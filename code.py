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
    Convierte una cadena de coordenadas DMS al formato deseado.
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

# --- 2. Función para actualizar formato y coordenadas en Google Sheets ---
def update_coordinates(sheet):
    try:
        # Actualiza el formato de las columnas M, N y O
        cell_format = {
            "backgroundColor": {"red": 1, "green": 1, "blue": 1},  # Fondo blanco (sin relleno)
            "horizontalAlignment": "CENTER",
            "textFormat": {
                "foregroundColor": {"red": 0, "green": 0, "blue": 0},  # Texto en negro
                "fontFamily": "Arial",
                "fontSize": 11
            }
        }
        # Aplica el formato a todas las columnas M, N y O
        sheet.format("M:O", cell_format)

        # Lee la columna M (número 13) y actualiza sus valores con el formato DMS deseado
        coordinates = sheet.col_values(13)
        formatted_coordinates = []
        # Se omite el encabezado (fila 1)
        for value in coordinates[1:]:
            if value:
                result = format_dms(value)
                formatted_coordinates.append(result if result else value)
            else:
                formatted_coordinates.append("")
        
        # Prepara el rango a actualizar en la columna M (desde la fila 2 hasta la última)
        cell_range = f"M2:M{len(formatted_coordinates)+1}"
        cell_list = sheet.range(cell_range)
        for i, cell in enumerate(cell_list):
            cell.value = formatted_coordinates[i]
        
        # Ejecuta el batch update de la columna M
        sheet.update_cells(cell_list)

        st.success("Actualización completada: formato y coordenadas actualizadas.")
    except Exception as e:
        st.error(f"Error durante la actualización: {str(e)}")

# --- 3. Interfaz de Streamlit ---
st.title("Actualización de Formato y Coordenadas DMS")

client = init_connection()
if client:
    sheet = load_sheet(client)
    if sheet:
        # Únicamente se muestra un botón
        if st.button("Actualizar Formato y Coordenadas"):
            update_coordinates(sheet)
