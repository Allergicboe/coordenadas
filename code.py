import streamlit as st
import gspread
from google.oauth2 import service_account
import re

# --- 1. Funciones de Conexión y Carga de Datos ---
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

def format_dms(value):
    pattern = r'(\d{2})°\s*(\d{1,2})\'\s*([\d\.]+)"\s*([NS])\s*(\d{2,3})°\s*(\d{1,2})\'\s*([\d\.]+)"\s*([EW])'
    match = re.match(pattern, value.strip())

    if match:
        lat_deg, lat_min, lat_sec, lat_dir, lon_deg, lon_min, lon_sec, lon_dir = match.groups()
        lat_sec = f"{float(lat_sec):04.1f}".replace(".", ".")
        lon_sec = f"{float(lon_sec):04.1f}".replace(".", ",")
        lon_deg = lon_deg.zfill(2)  # Asegurar dos dígitos en la longitud

        return f"{lat_deg}°{lat_min}'{lat_sec}\"{lat_dir} {lon_deg}°{lon_min}'{lon_sec}\"{lon_dir}"
    return None

# --- 2. Interfaz de Streamlit ---
st.title("Conexión y Conversión de Coordenadas DMS a Decimal")

# Iniciar conexión con Google Sheets
client = init_connection()
if client:
    sheet = load_sheet(client)
    if sheet:
        st.write("Datos cargados correctamente desde Google Sheets.")
        
        # Mostrar datos de las coordenadas DMS
        coordinates_column = "M"  # Columna que contiene las coordenadas DMS
        coordinates = sheet.col_values(ord(coordinates_column) - ord('A') + 1)
        
        st.subheader("Coordenadas DMS en la Hoja de Cálculo:")
        st.write(coordinates)

        # Botón para ejecutar la conversión
        if st.button("Convertir Coordenadas DMS a Decimal"):
            st.subheader("Conversión de DMS a Decimal:")
            for i, value in enumerate(coordinates[1:], 1):  # Ignorar encabezado
                if value:
                    decimal_value = format_dms(value)
                    if decimal_value:
                        st.write(f"Coordenadas {i}: {value} → {decimal_value}")
                    else:
                        st.write(f"Coordenadas {i}: Formato no válido.")
