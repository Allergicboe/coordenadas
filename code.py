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
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"
            ],
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

# --- 2. Función para formatear las coordenadas ---
def format_dms(value):
    """
    Toma una cadena con coordenadas y la formatea al siguiente formato:
    34°04'50.1"S 70°39'15.1"W
    """
    pattern = r'(\d+)[°º]\s*(\d+)[\'’]\s*([\d\.]+)"\s*([NS])\s+(\d+)[°º]\s*(\d+)[\'’]\s*([\d\.]+)"\s*([EW])'
    m = re.match(pattern, value.strip())
    if m:
        lat_deg, lat_min, lat_sec, lat_dir, lon_deg, lon_min, lon_sec, lon_dir = m.groups()
        try:
            lat_deg = int(lat_deg)
            lat_min = int(lat_min)
            lat_sec = float(lat_sec)
            lon_deg = int(lon_deg)
            lon_min = int(lon_min)
            lon_sec = float(lon_sec)
        except ValueError:
            return None

        # Formateamos:
        # - Grados y minutos con dos dígitos.
        # - Segundos con un decimal.
        formatted_lat = f"{lat_deg:02d}°{lat_min:02d}'{lat_sec:04.1f}\"{lat_dir}"
        formatted_lon = f"{lon_deg:02d}°{lon_min:02d}'{lon_sec:04.1f}\"{lon_dir}"
        return f"{formatted_lat} {formatted_lon}"
    return None

# --- 3. Función para actualizar formato y coordenadas en Google Sheets ---
def update_coordinates(sheet):
    try:
        # 1. Actualizar el formato de las columnas M, N y O (a partir de la fila 2, sin modificar el encabezado)
        cell_format = {
            "backgroundColor": {"red": 1, "green": 1, "blue": 1},  # Sin relleno (fondo blanco)
            "horizontalAlignment": "CENTER",
            "textFormat": {
                "foregroundColor": {"red": 0, "green": 0, "blue": 0},  # Texto en negro
                "fontFamily": "Arial",
                "fontSize": 11,
            },
        }
        # Aplica el formato desde la fila 2 hasta el final de las columnas M, N y O
        sheet.format("M2:O", cell_format)

        # 2. Leer los valores actuales de la columna M (suponiendo que allí están las coordenadas originales)
        coords = sheet.col_values(13)  # La columna M es la número 13
        if len(coords) <= 1:
            st.warning("No se encontraron datos en la columna M (excepto el encabezado).")
            return

        # 3. Preparar el rango a actualizar (desde la fila 2 hasta donde existan datos)
        start_row = 2
        end_row = len(coords)
        cell_range = f"M{start_row}:M{end_row}"
        cells = sheet.range(cell_range)

        # 4. Para cada celda, formatear el contenido y actualizarlo
        for i, cell in enumerate(cells):
            original_value = coords[i + 1]  # coords[0] es el encabezado
            if original_value:
                new_val = format_dms(original_value)
                cell.value = new_val if new_val else original_value
            else:
                cell.value = ""
        # Actualizar en lote las celdas
        sheet.update_cells(cells)
        st.success("Formato y coordenadas actualizadas correctamente.")
    except Exception as e:
        st.error(f"Error durante la actualización: {str(e)}")

# --- 4. Interfaz de Streamlit ---
def main():
    st.title("Actualizar Coordenadas DMS en Google Sheets")
    st.write("Presiona el botón para actualizar el formato de las columnas M, N y O, y para formatear las coordenadas de la columna M.")
    
    client = init_connection()
    if not client:
        return
    sheet = load_sheet(client)
    if not sheet:
        return

    if st.button("Actualizar Formato y Coordenadas"):
        update_coordinates(sheet)

if __name__ == "__main__":
    main()
