import streamlit as st
import gspread
from google.oauth2 import service_account
import re

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
    match = re.match(r"(\d{1,3})°\s*(\d{1,2})'(\d+(\.\d+)?)\"([NSWE])", str(dms))
    if not match:
        return None

    grados, minutos, segundos, _, direccion = match.groups()
    decimal = float(grados) + float(minutos) / 60 + round(float(segundos), 1) / 3600  # Redondeamos los segundos a 1 decimal
    if direccion in ['S', 'W']:
        decimal = -decimal

    return round(decimal, 8)

# Función para convertir de decimal a DMS sin espacios
def decimal_a_dms(decimal, direccion):
    grados = int(abs(decimal))
    minutos = int((abs(decimal) - grados) * 60)
    segundos = (abs(decimal) - grados - minutos / 60) * 3600

    # Redondear segundos a un solo decimal
    segundos = round(segundos, 1)

    # Formatear a DMS sin espacios, asegurando dos dígitos para minutos y un decimal para segundos
    dms = f"{grados:02d}°{int(minutos):02d}'{segundos:04.1f}\"{direccion}"
    return dms

# Función para formatear las celdas con el estilo deseado
def formatear_estilo(sheet, col_idx):
    """Aplica formato a la columna en el índice `col_idx`"""
    sheet.format(f'{col_idx}2:{col_idx}', {
        "horizontalAlignment": "CENTER",
        "textFormat": {"fontSize": 11, "fontFamily": "Arial", "bold": False},
        "backgroundColor": {"red": 1, "green": 1, "blue": 1},  # Sin relleno
        "textColor": {"red": 0, "green": 0, "blue": 0}  # Negro
    })

# Función para procesar la hoja y realizar la conversión
def procesar_hoja(sheet, conversion):
    datos = sheet.get_all_values()
    header = datos[0]
    data = datos[1:]

    # Obtener índices de las columnas necesarias (respetando mayúsculas)
    col_m = header.index("Ubicación sonda google maps")
    col_n = header.index("Latitud sonda")
    col_o = header.index("longitud Sonda")
    col_m_campo = header.index("Ubicación Campo")
    col_n_campo = header.index("Latitud campo")
    col_o_campo = header.index("Longitud Campo")

    # Listas para almacenar los valores a actualizar
    updates = []

    # Procesar cada fila
    for i, fila in enumerate(data, start=2):  # Comienza en la fila 2 (índice 1 en listas)
        dms_sonda = fila[col_m].strip() if col_m < len(fila) else ""
        lat_decimal_sonda = fila[col_n].strip() if col_n < len(fila) else ""
        lon_decimal_sonda = fila[col_o].strip() if col_o < len(fila) else ""

        dms_campo = fila[col_m_campo].strip() if col_m_campo < len(fila) else ""
        lat_decimal_campo = fila[col_n_campo].strip() if col_n_campo < len(fila) else ""
        lon_decimal_campo = fila[col_o_campo].strip() if col_o_campo < len(fila) else ""

        # Validación para evitar convertir valores no numéricos o vacíos
        try:
            if lat_decimal_sonda:
                lat_decimal_sonda = float(lat_decimal_sonda.replace(",", "."))
            else:
                lat_decimal_sonda = None

            if lon_decimal_sonda:
                lon_decimal_sonda = float(lon_decimal_sonda.replace(",", "."))
            else:
                lon_decimal_sonda = None

            if lat_decimal_campo:
                lat_decimal_campo = float(lat_decimal_campo.replace(",", "."))
            else:
                lat_decimal_campo = None

            if lon_decimal_campo:
                lon_decimal_campo = float(lon_decimal_campo.replace(",", "."))
            else:
                lon_decimal_campo = None

        except ValueError:
            lat_decimal_sonda = None
            lon_decimal_sonda = None
            lat_decimal_campo = None
            lon_decimal_campo = None

        # Si la conversión es de DMS a Decimal para Sonda
        if conversion == "DMS a Decimal":
            if dms_sonda and re.search(r"\d+°\s*\d+'", dms_sonda):
                lat_decimal_sonda = dms_a_decimal(dms_sonda)
                lon_decimal_sonda = lat_decimal_sonda  # Ya que tenemos un solo valor decimal por DMS

            # Agregar las actualizaciones a la lista
            updates.append({"range": f"N{i}", "values": [[lat_decimal_sonda]]})  # Latitud decimal
            updates.append({"range": f"O{i}", "values": [[lon_decimal_sonda]]})  # Longitud decimal

            # Si la conversión es de DMS a Decimal para Campo
            if dms_campo and re.search(r"\d+°\s*\d+'", dms_campo):
                lat_decimal_campo = dms_a_decimal(dms_campo)
                lon_decimal_campo = lat_decimal_campo  # Ya que tenemos un solo valor decimal por DMS

            # Agregar las actualizaciones a la lista
            updates.append({"range": f"Latitud campo{i}", "values": [[lat_decimal_campo]]})  # Latitud decimal
            updates.append({"range": f"Longitud Campo{i}", "values": [[lon_decimal_campo]]})  # Longitud decimal

        # Si la conversión es de Decimal a DMS para Sonda
        elif conversion == "Decimal a DMS":
            if lat_decimal_sonda is not None and lon_decimal_sonda is not None:
                lat_dms_sonda = decimal_a_dms(lat_decimal_sonda, "S" if lat_decimal_sonda < 0 else "N")
                lon_dms_sonda = decimal_a_dms(lon_decimal_sonda, "W" if lon_decimal_sonda < 0 else "E")
                dms_sonda = f"{lat_dms_sonda} {lon_dms_sonda}"

                # Agregar las actualizaciones a la lista para reemplazar los valores en DMS
                updates.append({"range": f"M{i}", "values": [[dms_sonda]]})  # Ubicación formateada

            # Si la conversión es de Decimal a DMS para Campo
            if lat_decimal_campo is not None and lon_decimal_campo is not None:
                lat_dms_campo = decimal_a_dms(lat_decimal_campo, "S" if lat_decimal_campo < 0 else "N")
                lon_dms_campo = decimal_a_dms(lon_decimal_campo, "W" if lon_decimal_campo < 0 else "E")
                dms_campo = f"{lat_dms_campo} {lon_dms_campo}"

                # Agregar las actualizaciones a la lista para reemplazar los valores en DMS
                updates.append({"range": f"M{i}", "values": [[dms_campo]]})  # Ubicación formateada

    # Aplicar batch update
    if updates:
        sheet.batch_update(updates)
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
