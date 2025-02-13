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
    Ejemplo de entrada: "34° 4' 50.1" S 70° 39' 15.01" W"
    Ejemplo de salida: "34°4'50.1"S 70°39'15,01"W"
    """
    pattern = r'(\d{2})°\s*(\d{1,2})\'\s*([\d\.]+)"\s*([NS])\s*(\d{2,3})°\s*(\d{1,2})\'\s*([\d\.]+)"\s*([EW])'
    match = re.match(pattern, value.strip())
    if match:
        lat_deg, lat_min, lat_sec, lat_dir, lon_deg, lon_min, lon_sec, lon_dir = match.groups()

        # Asegurarse de que los segundos de latitud tengan un decimal
        lat_sec = f"{float(lat_sec):04.1f}"  # Un decimal para latitud
        
        # Asegurar que la longitud tenga dos decimales para los segundos
        lon_sec = f"{float(lon_sec):05.2f}"  # Dos decimales para longitud (con coma)
        
        # Cambiar el punto por coma en los segundos de longitud
        lon_sec = lon_sec.replace(".", ",")
        
        lat_deg = lat_deg.zfill(2)
        lon_deg = lon_deg.zfill(3)  # Asegurar que la longitud tenga 3 dígitos

        # Retornar el formato DMS con los segundos bien formateados
        return f"{lat_deg}°{lat_min}'{lat_sec}\"{lat_dir} {lon_deg}°{lon_min}'{lon_sec}\"{lon_dir}"
    return None

# --- 2. Función para actualizar formato y coordenadas en Google Sheets ---
def update_coordinates(sheet):
    try:
        # Actualiza el formato de las columnas M, N y O (excepto la fila 1)
        cell_format = {
            "backgroundColor": {"red": 1, "green": 1, "blue": 1},  # Fondo blanco (sin relleno)
            "horizontalAlignment": "CENTER",
            "textFormat": {
                "foregroundColor": {"red": 0, "green": 0, "blue": 0},  # Texto en negro
                "fontFamily": "Arial",
                "fontSize": 11
            }
        }
        # Aplica el formato a todas las filas de M, N, O, excepto la fila 1
        sheet.format('M2:O', cell_format)

        # Ahora actualizamos las coordenadas con el formato DMS adecuado
        coordinates_range = sheet.range('M2:M' + str(sheet.row_count))
        for cell in coordinates_range:
            if cell.value:
                formatted_value = format_dms(cell.value)
                if formatted_value:
                    sheet.update_cell(cell.row, 14, formatted_value)  # Actualiza la columna N

        st.success("Las coordenadas fueron actualizadas correctamente.")
    except Exception as e:
        st.error(f"Error al actualizar las coordenadas: {str(e)}")
