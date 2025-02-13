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

# --- 2. Función para aplicar formato a las celdas ---
def apply_format(sheet):
    # Formato para columna M (texto) y para columnas N y O (números)
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

# --- 3. Función para formatear la cadena DMS ---
def format_dms(value):
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

# --- 4. Funciones que actualizan la hoja de cálculo ---

def dms_to_decimal(dms):
    """Convierte coordenadas DMS a formato decimal."""
    pattern = r'(\d+)[°º]\s*(\d+)[\'’]\s*([\d\.]+)"\s*([NS])\s+(\d+)[°º]\s*(\d+)[\'’]\s*([\d\.]+)"\s*([EW])'
    m = re.match(pattern, dms.strip())
    if m:
        lat_deg, lat_min, lat_sec, lat_dir, lon_deg, lon_min, lon_sec, lon_dir = m.groups()
        lat = int(lat_deg) + int(lat_min) / 60 + float(lat_sec) / 3600
        lon = int(lon_deg) + int(lon_min) / 60 + float(lon_sec) / 3600
        if lat_dir.upper() == 'S':
            lat = -lat
        if lon_dir.upper() == 'W':
            lon = -lon
        return lat, lon
    return None

def decimal_to_dms(lat, lon):
    """Convierte coordenadas decimales a formato DMS."""
    lat_deg = int(lat)
    lat_min = int((lat - lat_deg) * 60)
    lat_sec = (lat - lat_deg - lat_min / 60) * 3600
    lon_deg = int(lon)
    lon_min = int((lon - lon_deg) * 60)
    lon_sec = (lon - lon_deg - lon_min / 60) * 3600
    lat_dir = 'N' if lat >= 0 else 'S'
    lon_dir = 'E' if lon >= 0 else 'W'
    return f"{abs(lat_deg)}°{abs(lat_min)}'{abs(lat_sec):.1f}\"{lat_dir} {abs(lon_deg)}°{abs(lon_min)}'{abs(lon_sec):.1f}\"{lon_dir}"

def update_decimal_from_dms(sheet):
    try:
        apply_format(sheet)
        dms_values = sheet.col_values(13)  # Columna M
        if len(dms_values) <= 1:
            st.warning("No se encontraron datos en 'Ubicación sonda google maps'.")
            return
        lat_cells = sheet.range(f"N2:N{len(dms_values)}")
        lon_cells = sheet.range(f"O2:O{len(dms_values)}")
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
    try:
        apply_format(sheet)
        lat_values = sheet.col_values(14)  # Columna N
        lon_values = sheet.col_values(15)  # Columna O
        if len(lat_values) <= 1 or len(lon_values) <= 1:
            st.warning("No se encontraron datos en 'Latitud sonda' o 'longitud Sonda'.")
            return
        dms_cells = sheet.range(f"M2:M{len(lat_values)}")
        for i in range(1, len(lat_values)):
            lat = lat_values[i]
            lon = lon_values[i]
            if lat and lon:
                dms = decimal_to_dms(float(lat), float(lon))
                dms_cells[i-1].value = dms
        sheet.update_cells(dms_cells)
        st.success("Conversión de decimal a DMS completada.")
    except Exception as e:
        st.error(f"Error en la conversión de decimal a DMS: {str(e)}")

# --- 7. Interfaz de usuario en Streamlit ---
def main():
    st.set_page_config(page_title="Conversión de Coordenadas", page_icon="🌍", layout="wide")
    st.markdown("<h1 style='text-align: center; color: #0b5394;'>Conversión de Coordenadas</h1>", unsafe_allow_html=True)
    st.markdown("""
        <p style='text-align: center; font-size: 18px;'>Utiliza los botones para convertir las coordenadas entre DMS y Decimal:</p>
    """, unsafe_allow_html=True)
    
    client = init_connection()
    if not client:
        return
    sheet = load_sheet(client)
    if not sheet:
        return
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Convertir DMS a Decimal", use_container_width=True, key="dms_to_decimal"):
            update_decimal_from_dms(sheet)
    with col2:
        if st.button("Convertir Decimal a DMS", use_container_width=True, key="decimal_to_dms"):
            update_dms_from_decimal(sheet)
    
    st.markdown("<hr>", unsafe_allow_html=True)
    st.write("Nota: Los cambios se aplicarán en las columnas correspondientes de la hoja de cálculo.")
    
if __name__ == "__main__":
    main()
