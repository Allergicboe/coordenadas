import re
import streamlit as st
import gspread
from google.oauth2 import service_account

# --- 2. Funciones de Conexión y Carga de Datos ---
def init_connection():
    """Función para inicializar la conexión con Google Sheets."""
    try:
        credentials = service_account.Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=[
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
        )
        client = gspread.authorize(credentials)
        return client
    except Exception as e:
        st.error(f"Error en la conexión: {str(e)}")
        return None

def load_sheet(client):
    """Función para cargar la hoja de trabajo de Google Sheets."""
    try:
        return client.open_by_url(st.secrets["spreadsheet_url"]).sheet1
    except Exception as e:
        st.error(f"Error al cargar la planilla: {str(e)}")
        return None

# Función para convertir DMS a decimal
def dms_a_decimal(dms):
    # Expresión regular para manejar el formato de DMS con posibles espacios entre grados, minutos y segundos
    match = re.match(r"(\d{1,3})°\s*(\d{1,2})'\s*(\d+(\.\d+)?)\"\s*([NSWE])", str(dms))
    if not match:
        return None

    grados, minutos, segundos, _, direccion = match.groups()
    grados = int(grados)
    minutos = int(minutos)
    segundos = float(segundos)

    decimal = grados + minutos / 60 + segundos / 3600
    if direccion in ['S', 'W']:
        decimal = -decimal

    return round(decimal, 8)

# Función para convertir de decimal a DMS sin espacios y con redondeo de segundos
def decimal_a_dms(decimal, direccion):
    grados = int(abs(decimal))
    minutos = int((abs(decimal) - grados) * 60)
    segundos = (abs(decimal) - grados - minutos / 60) * 3600

    # Redondear segundos a un solo decimal
    segundos = round(segundos, 1)

    # Formatear a DMS sin espacios, asegurando dos dígitos para minutos y un decimal para segundos
    dms = f"{grados:02d}°{int(minutos):02d}'{segundos:04.1f}\"{direccion}"
    return dms

# Función para procesar la hoja y realizar la conversión
def procesar_hoja(sheet, conversion):
    datos = sheet.get_all_values()
    header = datos[0]
    data = datos[1:]

    # Obtener índices de las columnas necesarias
    col_m = header.index("Ubicación sonda google maps")
    col_n = header.index("Latitud sonda")
    col_o = header.index("Longitud Sonda")

    # Listas para almacenar los valores a actualizar
    updates = []

    # Procesar cada fila
    for i, fila in enumerate(data, start=2):  # Comienza en la fila 2 (índice 1 en listas)
        dms_sonda = fila[col_m].strip() if col_m < len(fila) else ""
        lat_decimal = fila[col_n].strip() if col_n < len(fila) else ""
        lon_decimal = fila[col_o].strip() if col_o < len(fila) else ""

        # Validación para evitar convertir valores no numéricos o vacíos
        try:
            if lat_decimal:
                lat_decimal = float(lat_decimal.replace(",", "."))
            else:
                lat_decimal = None

            if lon_decimal:
                lon_decimal = float(lon_decimal.replace(",", "."))
            else:
                lon_decimal = None
        except ValueError:
            lat_decimal = None
            lon_decimal = None

        # Si la conversión es de DMS a Decimal
        if conversion == "DMS a Decimal":
            if dms_sonda and re.search(r"\d+°\s*\d+'", dms_sonda):
                # Corregir y reemplazar DMS incorrecto
                corrected_dms = re.sub(r"\s+", "", dms_sonda)  # Eliminar espacios extra
                lat_decimal = dms_a_decimal(corrected_dms)
                lon_decimal = lat_decimal  # Ya que tenemos un solo valor decimal por DMS

                # Agregar las actualizaciones a la lista
                updates.append({"range": f"N{i}", "values": [[lat_decimal]]})  # Latitud decimal
                updates.append({"range": f"O{i}", "values": [[lon_decimal]]})  # Longitud decimal

                # También actualizamos la ubicación DMS corregida
                dms_sonda_corregido = decimal_a_dms(lat_decimal, "S" if lat_decimal < 0 else "N") + " " + decimal_a_dms(lon_decimal, "W" if lon_decimal < 0 else "E")
                updates.append({"range": f"M{i}", "values": [[dms_sonda_corregido]]})  # Ubicación corregida

        # Si la conversión es de Decimal a DMS
        elif conversion == "Decimal a DMS":
            if lat_decimal is not None and lon_decimal is not None:
                lat_dms = decimal_a_dms(lat_decimal, "S" if lat_decimal < 0 else "N")
                lon_dms = decimal_a_dms(lon_decimal, "W" if lon_decimal < 0 else "E")
                dms_sonda = f"{lat_dms} {lon_dms}"

                # Agregar las actualizaciones a la lista para reemplazar los valores en DMS
                updates.append({"range": f"M{i}", "values": [[dms_sonda]]})  # Ubicación formateada

    # Aplicar batch update
    if updates:
        sheet.batch_update(updates)

        # Aplicar formato a las columnas M, N y O
        format_updates = [
            {
                "range": f"M2:M{len(data)+1}",
                "format": {
                    "textFormat": {"fontFamily": "Arial", "fontSize": 11, "foregroundColor": {"red": 0, "green": 0, "blue": 0}},
                    "horizontalAlignment": "CENTER",
                    "backgroundColor": {"red": 1, "green": 1, "blue": 1}
                }
            },
            {
                "range": f"N2:N{len(data)+1}",
                "format": {
                    "textFormat": {"fontFamily": "Arial", "fontSize": 11, "foregroundColor": {"red": 0, "green": 0, "blue": 0}},
                    "horizontalAlignment": "CENTER",
                    "backgroundColor": {"red": 1, "green": 1, "blue": 1}
                }
            },
            {
                "range": f"O2:O{len(data)+1}",
                "format": {
                    "textFormat": {"fontFamily": "Arial", "fontSize": 11, "foregroundColor": {"red": 0, "green": 0, "blue": 0}},
                    "horizontalAlignment": "CENTER",
                    "backgroundColor": {"red": 1, "green": 1, "blue": 1}
                }
            }
        ]

        # Aplicar formato de celdas a las columnas M, N y O
        sheet.batch_update(format_updates)

        st.success("✅ Conversión completada y planilla actualizada.")
    else:
        st.warning("⚠️ No se encontraron datos válidos para actualizar.")

# --- Interfaz de Streamlit ---
st.title('Conversión de Coordenadas')
st.sidebar.header('Opciones')

# Selección del tipo de conversión
conversion = st.sidebar.radio("Seleccione el tipo de conversión", ('Decimal a DMS', 'DMS a Decimal'))

# Conectar a Google Sheets
client = init_connection()
if client:
    sheet = load_sheet(client)

    if sheet:
        if conversion == 'DMS a Decimal':
            if st.button("Convertir DMS a Decimal"):
                procesar_hoja(sheet, conversion)
        elif conversion == 'Decimal a DMS':
            if st.button("Convertir Decimal a DMS"):
                procesar_hoja(sheet, conversion)
