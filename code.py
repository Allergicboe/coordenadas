import streamlit as st
import gspread
from google.oauth2 import service_account
import re

# --- 1. Funciones de Conexión y Carga de Datos ---
def init_connection():
    """Inicializa la conexión con Google Sheets."""
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
    """Carga la hoja de trabajo de Google Sheets."""
    try:
        return client.open_by_url(st.secrets["spreadsheet_url"]).sheet1
    except Exception as e:
        st.error(f"Error al cargar la planilla: {str(e)}")
        return None

def format_dms(value):
    """
    Convierte las coordenadas DMS al formato especificado.
    Ejemplo de entrada: "12° 34' 56.7" N 123° 45' 67.8" W"
    """
    pattern = r'(\d{2})°\s*(\d{1,2})\'\s*([\d\.]+)"\s*([NS])\s*(\d{2,3})°\s*(\d{1,2})\'\s*([\d\.]+)"\s*([EW])'
    match = re.match(pattern, value.strip())
    if match:
        lat_deg, lat_min, lat_sec, lat_dir, lon_deg, lon_min, lon_sec, lon_dir = match.groups()
        lat_sec = f"{float(lat_sec):04.1f}"
        lon_sec = f"{float(lon_sec):04.1f}".replace(".", ",")
        lat_deg = lat_deg.zfill(2)
        lon_deg = lon_deg.zfill(2)
        return f"{lat_deg}°{lat_min}'{lat_sec}\"{lat_dir} {lon_deg}°{lon_min}'{lon_sec}\"{lon_dir}"
    return None

def update_coordinates(sheet):
    """Actualiza en Google Sheets las coordenadas DMS con el nuevo formato."""
    try:
        # Se obtiene la columna M (columna 13) sin mostrarla en la interfaz
        coordinates = sheet.col_values(13)
        formatted_coordinates = []
        
        # Se ignora el encabezado
        for value in coordinates[1:]:
            if value:
                result = format_dms(value)
                formatted_coordinates.append(result if result else value)
            else:
                formatted_coordinates.append("")
        
        # Se prepara el rango a actualizar (M2:M{última fila})
        cell_list = sheet.range(f'M2:M{len(formatted_coordinates)+1}')
        for i, cell in enumerate(cell_list):
            cell.value = formatted_coordinates[i]
        
        # Se realiza la actualización en lote
        sheet.update_cells(cell_list)
        st.success("Coordenadas actualizadas correctamente en Google Sheets.")
    except Exception as e:
        st.error(f"Error al actualizar las coordenadas: {str(e)}")

# --- 2. Interfaz de Streamlit ---
st.title("Actualización de Coordenadas DMS")

client = init_connection()
if client:
    sheet = load_sheet(client)
    if sheet:
        # Únicamente se muestra el botón
        if st.button("Actualizar Coordenadas DMS"):
            update_coordinates(sheet)
