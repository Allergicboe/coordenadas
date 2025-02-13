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
    match = re.match(r"(\d{1,3})°(\d{1,2})'([\d.]+)\"([NSWE])", str(dms))
    if not match:
        return None

    grados, minutos, segundos, direccion = match.groups()
    decimal = float(grados) + float(minutos) / 60 + float(segundos) / 3600
    if direccion in ['S', 'W']:
        decimal = -decimal

    return round(decimal, 8)

# Función para convertir de decimal a DMS sin espacios
def decimal_a_dms(decimal, direccion):
    grados = int(abs(decimal))
    minutos = int((abs(decimal) - grados) * 60)
    segundos = (abs(decimal) - grados - minutos / 60) * 3600

    # Formatear a DMS sin espacios
    dms = f"{grados:02d}°{minutos:02d}'{segundos:0.1f}\"{direccion}"
    return dms

# Función para procesar la hoja y realizar la conversión
def procesar_hoja(sheet):
    datos = sheet.get_all_values()
    header = datos[0]
    data = datos[1:]

    # Obtener índices de las columnas necesarias
    col_m = header.index("Ubicación sonda google maps")
    col_n = header.index("Latitud sonda")
    col_o = header.index("longitud Sonda")

    # Listas para almacenar los valores a actualizar
    updates = []

    # Procesar cada fila
    for i, fila in enumerate(data, start=2):  # Comienza en la fila 2 (índice 1 en listas)
        dms_sonda = fila[col_m].strip() if col_m < len(fila) else ""
        lat_decimal = fila[col_n].strip() if col_n < len(fila) else ""
        lon_decimal = fila[col_o].strip() if col_o < len(fila) else ""

        # Si "Ubicación sonda google maps" tiene valor, convertir a decimal
        if dms_sonda and re.search(r"\d+°\s*\d+'", dms_sonda):
            coordenadas = dms_a_decimal(dms_sonda)
            if coordenadas:
                lat_decimal = coordenadas[0]
                lon_decimal = coordenadas[1]

        # Si "Latitud sonda" y "longitud Sonda" tienen valor, convertir a DMS
        elif lat_decimal and lon_decimal:
            lat_decimal = float(lat_decimal.replace(",", "."))
            lon_decimal = float(lon_decimal.replace(",", "."))
            lat_dms = decimal_a_dms(lat_decimal, "S" if lat_decimal < 0 else "N")
            lon_dms = decimal_a_dms(lon_decimal, "W" if lon_decimal < 0 else "E")
            dms_sonda = f"{lat_dms} {lon_dms}"

        # Agregar las actualizaciones a la lista
        updates.append({"range": f"M{i}", "values": [[dms_sonda]]})  # Ubicación formateada
        updates.append({"range": f"N{i}", "values": [[lat_decimal]]})  # Latitud decimal
        updates.append({"range": f"O{i}", "values": [[lon_decimal]]})  # Longitud decimal

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
                procesar_hoja(sheet)
        elif conversion == 'Decimal a DMS':
            if st.button("Convertir Decimal a DMS"):
                procesar_hoja(sheet)
