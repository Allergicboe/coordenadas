import streamlit as st
import gspread
from google.oauth2 import service_account
import re
import pandas as pd

# Configuración de la página
st.set_page_config(
    page_title="Conversor de Coordenadas",
    page_icon="📍",
    layout="wide"
)

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
    """Aplica formato a las celdas de la hoja de cálculo para Sondas."""
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
    # Columna M: Ubicación sonda google maps (texto, DMS)
    sheet.format("M2:M", text_format)
    # Columnas N y O: Latitud sonda y Longitud sonda (números)
    sheet.format("N2:O", number_format)

def apply_format_field(sheet):
    """Aplica formato a las celdas de la hoja de cálculo para Campo."""
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
    # Columna E: Ubicación campo (texto, DMS)
    sheet.format("E2:E", text_format)
    # Columnas F y G: Latitud campo y Longitud Campo (números)
    sheet.format("F2:G", number_format)

# --- 3. Función para formatear la cadena DMS ---
def format_dms(value):
    """Formatea una cadena DMS al formato correcto."""
    pattern = r'(\d+)[°º]\s*(\d+)[\'']\s*([\d\.]+)"\s*([NS])\s+(\d+)[°º]\s*(\d+)[\'']\s*([\d\.]+)"\s*([EW])'
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
def update_dms_format_column(sheet):
    """Actualiza la columna DMS en la hoja de cálculo para Sondas."""
    dms_values = sheet.col_values(13)  # Columna M
    if len(dms_values) <= 1:
        return
    start_row = 2
    end_row = len(dms_values)
    cell_range = f"M{start_row}:M{end_row}"
    cells = sheet.range(cell_range)
    for i, cell in enumerate(cells):
        original_value = dms_values[i + 1]  # omite el encabezado
        if original_value:
            new_val = format_dms(original_value)
            cell.value = new_val if new_val is not None else original_value
    sheet.update_cells(cells)

def update_dms_format_column_field(sheet):
    """Actualiza la columna DMS en la hoja de cálculo para Campo."""
    dms_values = sheet.col_values(5)  # Columna E
    if len(dms_values) <= 1:
        return
    start_row = 2
    end_row = len(dms_values)
    cell_range = f"E{start_row}:E{end_row}"
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
    pattern = r'(\d{2})[°º](\d{2})[\''](\d{1,2}\.\d)"([NS])\s+(\d{2})[°º](\d{2})[\''](\d{1,2}\.\d)"([EW])'
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
def update_decimal_from_dms(sheet):
    """Convierte DMS a decimal y actualiza las columnas correspondientes para Sondas."""
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
    """Convierte decimal a DMS y actualiza la columna correspondiente para Sondas."""
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

def update_decimal_from_dms_field(sheet):
    """Convierte DMS a decimal y actualiza las columnas correspondientes para Campo."""
    try:
        apply_format_field(sheet)
        update_dms_format_column_field(sheet)
        dms_values = sheet.col_values(5)  # Columna E
        if len(dms_values) <= 1:
            st.warning("No se encontraron datos en 'Ubicación campo'.")
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
        st.success("Conversión de DMS a decimal completada.")
    except Exception as e:
        st.error(f"Error en la conversión de DMS a decimal: {str(e)}")

def update_dms_from_decimal_field(sheet):
    """Convierte decimal a DMS y actualiza la columna correspondiente para Campo."""
    try:
        apply_format_field(sheet)
        update_dms_format_column_field(sheet)
        lat_values = sheet.col_values(6)  # Columna F
        lon_values = sheet.col_values(7)  # Columna G
        if len(lat_values) <= 1 or len(lon_values) <= 1:
            st.warning("No se encontraron datos en 'Latitud campo' o 'Longitud Campo'.")
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
        st.success("Conversión de decimal a DMS completada.")
    except Exception as e:
        st.error(f"Error en la conversión de decimal a DMS: {str(e)}")

def main():
    st.title("Conversión de Coordenadas: Sondas📍")
    st.write("Selecciona la conversión que deseas realizar:")

    client = init_connection()
    if not client:
        return
    sheet = load_sheet(client)
    if not sheet:
        return

# Botones para Sondas (Columnas M, N y O)
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Convertir DMS a Decimal (Sonda)", help="Convierte las coordenadas DMS a formato decimal", key="dms_to_decimal", use_container_width=True):
            with st.spinner("Procesando conversión DMS a Decimal para Sondas..."):
                update_decimal_from_dms(sheet)
    with col2:
        if st.button("Convertir Decimal a DMS (Sonda)", help="Convierte las coordenadas decimales a formato DMS", key="decimal_to_dms", use_container_width=True):
            with st.spinner("Procesando conversión Decimal a DMS para Sondas..."):
                update_dms_from_decimal(sheet)

    st.markdown("---")
    st.title("Conversión de Coordenadas: Campo 📍")
    st.write("Selecciona la conversión que deseas realizar:")

    # Botones para Campo (Columnas E, F y G)
    col3, col4 = st.columns(2)
    with col3:
        if st.button("Convertir DMS a Decimal (Campo)", help="Convierte las coordenadas DMS a formato decimal para Ubicación campo", key="dms_to_decimal_field", use_container_width=True):
            with st.spinner("Procesando conversión DMS a Decimal para Campo..."):
                update_decimal_from_dms_field(sheet)
    with col4:
        if st.button("Convertir Decimal a DMS (Campo)", help="Convierte las coordenadas decimales a formato DMS para Ubicación campo", key="decimal_to_dms_field", use_container_width=True):
            with st.spinner("Procesando conversión Decimal a DMS para Campo..."):
                update_dms_from_decimal_field(sheet)

    st.markdown("---")
    
    # Sección de previsualización
    st.subheader("📊 Previsualización de Datos")
    
    try:
        # Crear tabs para separar la visualización de Sondas y Campo
        tab1, tab2 = st.tabs(["Datos de Sondas", "Datos de Campo"])
        
        with tab1:
            # Obtener datos de Sondas
            sondas_data = {
                'Ubicación (DMS)': sheet.col_values(13)[1:6],  # Primeros 5 valores de la columna M
                'Latitud': sheet.col_values(14)[1:6],          # Primeros 5 valores de la columna N
                'Longitud': sheet.col_values(15)[1:6]          # Primeros 5 valores de la columna O
            }
            df_sondas = pd.DataFrame(sondas_data)
            if not df_sondas.empty:
                st.write("Mostrando las primeras 5 filas de datos de Sondas:")
                st.dataframe(df_sondas, use_container_width=True)
            else:
                st.info("No hay datos disponibles para mostrar en Sondas")

        with tab2:
            # Obtener datos de Campo
            campo_data = {
                'Ubicación (DMS)': sheet.col_values(5)[1:6],   # Primeros 5 valores de la columna E
                'Latitud': sheet.col_values(6)[1:6],           # Primeros 5 valores de la columna F
                'Longitud': sheet.col_values(7)[1:6]           # Primeros 5 valores de la columna G
            }
            df_campo = pd.DataFrame(campo_data)
            if not df_campo.empty:
                st.write("Mostrando las primeras 5 filas de datos de Campo:")
                st.dataframe(df_campo, use_container_width=True)
            else:
                st.info("No hay datos disponibles para mostrar en Campo")

    except Exception as e:
        st.error(f"Error al cargar la previsualización: {str(e)}")
    
    # Información adicional
    st.markdown("---")
    st.markdown("""
    ### ℹ️ Información importante
    - Los datos se actualizan automáticamente en la hoja de cálculo
    - La previsualización muestra las primeras 5 filas de cada sección
    - Formato DMS esperado: DD°MM'SS.S"N DD°MM'SS.S"E
    - Las conversiones se realizan en tiempo real
    - Asegúrese de que los datos estén en el formato correcto antes de convertir
    """)

    # Estado del sistema
    st.markdown("---")
    st.subheader("🔄 Estado del Sistema")
    
    # Verificar conexión y datos
    connection_status = "✅ Conectado" if client else "❌ Desconectado"
    sheet_status = "✅ Hoja de cálculo cargada" if sheet else "❌ Error al cargar la hoja"
    
    status_col1, status_col2 = st.columns(2)
    with status_col1:
        st.info(f"Estado de conexión: {connection_status}")
    with status_col2:
        st.info(f"Estado de la hoja: {sheet_status}")

if __name__ == "__main__":
    main()
