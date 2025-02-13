import streamlit as st
import gspread
from google.oauth2 import service_account
import re

# --- 1. Conexión y carga de datos ---
def init_connection():
    """Inicializa la conexión con Google Sheets."""
    try:
        credentials = service_account.Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"],
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

# --- 2. Funciones para aplicar formato a las celdas ---
def apply_format(sheet, col_start, col_end):
    """Aplica formato a las celdas de la hoja de cálculo."""
    text_format = {
        "backgroundColor": {"red": 1, "green": 1, "blue": 1},
        "horizontalAlignment": "CENTER",
        "textFormat": {
            "foregroundColor": {"red": 0, "green": 0, "blue": 0},
            "fontFamily": "Arial",
            "fontSize": 11
        }
    }
    number_format = {
        "numberFormat": {
            "type": "NUMBER",
            "pattern": "#,##0.00000000"
        },
        "backgroundColor": {"red": 1, "green": 1, "blue": 1},
        "horizontalAlignment": "CENTER",
        "textFormat": {
            "foregroundColor": {"red": 0, "green": 0, "blue": 0},
            "fontFamily": "Arial",
            "fontSize": 11
        }
    }
    # Aplica formato a la columna de DMS (texto)
    sheet.format(f"{col_start}2:{col_end}", text_format)
    # Aplica formato a las columnas de Latitud y Longitud (números)
    sheet.format(f"{col_start+1}2:{col_end+1}", number_format)

# --- 3. Función para formatear la cadena DMS ---
def format_dms(value):
    """Formatea una cadena DMS al formato correcto."""
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
        formatted_lat = f"{lat_deg:02d}°{lat_min:02d}'{lat_sec:04.1f}\"{lat_dir.upper()}"
        formatted_lon = f"{lon_deg:02d}°{lon_min:02d}'{lon_sec:04.1f}\"{lon_dir.upper()}"
        return f"{formatted_lat} {formatted_lon}"
    return None

# --- 4. Actualizar el contenido de la columna DMS ---
def update_dms_format_column(sheet, col_start, col_end):
    """Actualiza la columna DMS en la hoja de cálculo."""
    dms_values = sheet.col_values(col_start)  # Columna M o E
    if len(dms_values) <= 1:
        return
    start_row = 2
    end_row = len(dms_values)
    cell_range = f"{col_start}{start_row}:{col_end}{end_row}"
    cells = sheet.range(cell_range)
    for i, cell in enumerate(cells):
        original_value = dms_values[i + 1]  # omite el encabezado
        if original_value:
            new_val = format_dms(original_value)
            cell.value = new_val if new_val is not None else original_value
    sheet.update_cells(cells)

# --- 5. Funciones de conversión ---
def dms_to_decimal(dms_str):
    """Convierte DMS a decimal."""
    pattern = r'(\d{2})[°º](\d{2})[\'’](\d{1,2}\.\d)"([NS])\s+(\d{2})[°º](\d{2})[\'’](\d{1,2}\.\d)"([EW])'
    m = re.match(pattern, dms_str.strip())
    if m:
        lat_deg, lat_min, lat_sec, lat_dir, lon_deg, lon_min, lon_sec, lon_dir = m.groups()
        lat = int(lat_deg) + int(lat_min) / 60 + float(lat_sec) / 3600
        lon = int(lon_deg) + int(lon_min) / 60 + float(lon_sec) / 3600
        if lat_dir.upper() == "S":
            lat = -lat
        if lon_dir.upper() == "W":
            lon = -lon
        return lat, lon
    return None

def decimal_to_dms(lat, lon):
    """Convierte decimal a DMS."""
    lat_dir = "N" if lat >= 0 else "S"
    abs_lat = abs(lat)
    lat_deg = int(abs_lat)
    lat_min = int((abs_lat - lat_deg) * 60)
    lat_sec = (abs_lat - lat_deg - lat_min / 60) * 3600
    lon_dir = "E" if lon >= 0 else "W"
    abs_lon = abs(lon)
    lon_deg = int(abs_lon)
    lon_min = int((abs_lon - lon_deg) * 60)
    lon_sec = (abs_lon - lon_deg - lon_min / 60) * 3600
    dms_lat = f"{lat_deg:02d}°{lat_min:02d}'{lat_sec:04.1f}\"{lat_dir}"
    dms_lon = f"{lon_deg:02d}°{lon_min:02d}'{lon_sec:04.1f}\"{lon_dir}"
    return f"{dms_lat} {dms_lon}"

# --- 6. Funciones que actualizan la hoja de cálculo ---
def update_decimal_from_dms(sheet, col_start, col_end):
    """Convierte DMS a decimal y actualiza las columnas correspondientes."""
    try:
        apply_format(sheet, col_start, col_end)
        update_dms_format_column(sheet, col_start, col_end)
        dms_values = sheet.col_values(col_start)  # Columna M o E
        if len(dms_values) <= 1:
            st.warning("No se encontraron datos en 'Ubicación Campo' o 'Ubicación Sonda'.")
            return
        num_rows = len(dms_values)
        lat_cells = sheet.range(f"{col_start+1}2:{col_start+1}{num_rows}")
        lon_cells = sheet.range(f"{col_start+2}2:{col_start+2}{num_rows}")
        for i, dms in enumerate(dms_values[1:]):  # omitir encabezado
            if dms:
                result = dms_to_decimal(dms)
                if result is not None:
                    lat, lon = result
                    lat_cells[i].value = round(lat, 8)
                    lon_cells[i].value = round(lon, 8)
        sheet.update_cells(lat_cells)
        sheet.update_cells(lon_cells)
        st.success("Conversión de DMS a decimal completada.")
    except Exception as e:
        st.error(f"Error en la conversión de DMS a decimal: {str(e)}")

def update_dms_from_decimal(sheet, col_start, col_end):
    """Convierte decimal a DMS y actualiza la columna correspondiente."""
    try:
        apply_format(sheet, col_start, col_end)
        update_dms_format_column(sheet, col_start, col_end)
        lat_values = sheet.col_values(col_start+1)  # Columna F
        lon_values = sheet.col_values(col_start+2)  # Columna G
        if len(lat_values) <= 1 or len(lon_values) <= 1:
            st.warning("No se encontraron datos en 'Latitud campo' o 'Longitud campo'.")
            return
        num_rows = min(len(lat_values), len(lon_values))
        dms_cells = sheet.range(f"{col_start}2:{col_start}{num_rows}")
        for i in range(1, num_rows):
            lat_str = lat_values[i]
            lon_str = lon_values[i]
            if lat_str and lon_str:
                try:
                    lat = float(lat_str.replace(",", "."))
                    lon = float(lon_str.replace(",", "."))
                    dms = decimal_to_dms(lat, lon)
                    dms_cells[i-1].value = dms
                except Exception:
                    pass
        sheet.update_cells(dms_cells)
        st.success("Conversión de decimal a DMS completada.")
    except Exception as e:
        st.error(f"Error en la conversión de decimal a DMS: {str(e)}")

# --- 7. Interfaz de usuario en Streamlit ---
def main():
    st.title("Conversión de Coordenadas: Sondas 📍 y Campos 🌱")
    st.write("Selecciona la conversión que deseas realizar:")

    client = init_connection()
    if not client:
        return
    sheet = load_sheet(client)
    if not sheet:
        return

    # Usar columnas para botones
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Sondas")
        if st.button("Convertir DMS a Decimal (Sondas)", help="Convierte las coordenadas DMS a formato decimal", key="dms_to_decimal_sonda", use_container_width=True):
            update_decimal_from_dms(sheet, 13, 15)
    with col2:
        st.subheader("Campos")
        if st.button("Convertir DMS a Decimal (Campos)", help="Convierte las coordenadas DMS a formato decimal", key="dms_to_decimal_campo", use_container_width=True):
            update_decimal_from_dms(sheet, 5, 7)

    # Separador entre botones
    st.markdown("---")

    # Usar columnas para las otras conversiones
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Convertir Decimal a DMS (Sondas)", help="Convierte las coordenadas decimales a formato DMS", key="decimal_to_dms_sonda", use_container_width=True):
            update_dms_from_decimal(sheet, 13, 15)
    with col2:
        if st.button("Convertir Decimal a DMS (Campos)", help="Convierte las coordenadas decimales a formato DMS", key="decimal_to_dms_campo", use_container_width=True):
            update_dms_from_decimal(sheet, 5, 7)

if __name__ == "__main__":
    main()
