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
def apply_format(sheet):
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
    # Aplica formato a la columna M (texto, DMS)
    sheet.format("M2:M", text_format)
    # Aplica formato a las columnas N y O (números)
    sheet.format("N2:O", number_format)
    # Aplica formato a las columnas E, F, G (Campo)
    sheet.format("E2:E", text_format)
    sheet.format("F2:G", number_format)

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

# --- 4. Funciones de conversión ---
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

# --- 5. Funciones que actualizan la hoja de cálculo ---
def update_decimal_from_dms(sheet):
    """Convierte DMS a decimal y actualiza las columnas correspondientes."""
    try:
        apply_format(sheet)
        update_dms_format_column(sheet)
        dms_values = sheet.col_values(13)  # Columna M
        if len(dms_values) <= 1:
            st.warning("No se encontraron datos en 'Ubicación sonda google maps'.")
            return
        num_rows = len(dms_values)
        lat_cells = sheet.range(f"N2:N{num_rows}")
        lon_cells = sheet.range(f"O2:O{num_rows}")
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

def update_dms_from_decimal(sheet):
    """Convierte decimal a DMS y actualiza la columna correspondiente."""
    try:
        apply_format(sheet)
        update_dms_format_column(sheet)
        lat_values = sheet.col_values(14)  # Columna N
        lon_values = sheet.col_values(15)  # Columna O
        if len(lat_values) <= 1 or len(lon_values) <= 1:
            st.warning("No se encontraron datos en 'Latitud sonda' o 'longitud Sonda'.")
            return
        num_rows = min(len(lat_values), len(lon_values))
        dms_cells = sheet.range(f"M2:M{num_rows}")
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

# --- 6. Funciones para Campo ---
def update_decimal_from_dms_campo(sheet):
    """Convierte DMS a decimal y actualiza las columnas E, F, G (Campo)."""
    try:
        apply_format(sheet)
        dms_values = sheet.col_values(5)  # Columna E (Ubicación Campo)
        if len(dms_values) <= 1:
            st.warning("No se encontraron datos en 'Ubicación Campo'.")
            return
        num_rows = len(dms_values)
        lat_cells = sheet.range(f"F2:F{num_rows}")
        lon_cells = sheet.range(f"G2:G{num_rows}")
        for i, dms in enumerate(dms_values[1:]):  # omitir encabezado
            if dms:
                result = dms_to_decimal(dms)
                if result is not None:
                    lat, lon = result
                    lat_cells[i].value = round(lat, 8)
                    lon_cells[i].value = round(lon, 8)
        sheet.update_cells(lat_cells)
        sheet.update_cells(lon_cells)
        st.success("Conversión de DMS a decimal para Campo completada.")
    except Exception as e:
        st.error(f"Error en la conversión de DMS a decimal para Campo: {str(e)}")

def update_dms_from_decimal_campo(sheet):
    """Convierte decimal a DMS y actualiza las columnas E, F, G (Campo)."""
    try:
        apply_format(sheet)
        lat_values = sheet.col_values(6)  # Columna F (Latitud Campo)
        lon_values = sheet.col_values(7)  # Columna G (Longitud Campo)
        if len(lat_values) <= 1 or len(lon_values) <= 1:
            st.warning("No se encontraron datos en 'Latitud Campo' o 'Longitud Campo'.")
            return
        num_rows = min(len(lat_values), len(lon_values))
        dms_cells = sheet.range(f"E2:E{num_rows}")
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
        st.success("Conversión de decimal a DMS para Campo completada.")
    except Exception as e:
        st.error(f"Error en la conversión de decimal a DMS para Campo: {str(e)}")

# --- 7. Interfaz de usuario en Streamlit ---
def main():
    st.title("Conversión de Coordenadas: Sondas y Campos📍")

    # Encabezado para Sonda
    st.header("Conversión de Coordenadas: Sonda")
    st.write("Convierte las coordenadas de la ubicación de la sonda entre los formatos DMS y Decimal.")
    if st.button("Actualizar decimal desde DMS (Sonda)"):
        client = init_connection()
        if client:
            sheet = load_sheet(client)
            if sheet:
                update_decimal_from_dms(sheet)
    if st.button("Actualizar DMS desde decimal (Sonda)"):
        client = init_connection()
        if client:
            sheet = load_sheet(client)
            if sheet:
                update_dms_from_decimal(sheet)

    # Encabezado para Campo
    st.header("Conversión de Coordenadas: Campo")
    st.write("Convierte las coordenadas de la ubicación del campo entre los formatos DMS y Decimal.")
    if st.button("Actualizar decimal desde DMS (Campo)"):
        client = init_connection()
        if client:
            sheet = load_sheet(client)
            if sheet:
                update_decimal_from_dms_campo(sheet)
    if st.button("Actualizar DMS desde decimal (Campo)"):
        client = init_connection()
        if client:
            sheet = load_sheet(client)
            if sheet:
                update_dms_from_decimal_campo(sheet)

if __name__ == "__main__":
    main()
