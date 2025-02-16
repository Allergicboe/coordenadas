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
    with st.spinner("Conectando con Google Sheets..."):
        try:
            credentials = service_account.Credentials.from_service_account_info(
                st.secrets["gcp_service_account"],
                scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"],
            )
            client = gspread.authorize(credentials)
            st.success("✅ Conexión establecida exitosamente")
            return client
        except Exception as e:
            st.error(f"❌ Error en la conexión: {str(e)}")
            return None

def load_sheet(client):
    """Carga la hoja de trabajo de Google Sheets."""
    with st.spinner("Cargando datos de la planilla..."):
        try:
            sheet = client.open_by_url(st.secrets["spreadsheet_url"]).sheet1
            data = pd.DataFrame(sheet.get_all_records())
            st.success(f"✅ Datos cargados: {len(data)} registros")
            return sheet
        except Exception as e:
            st.error(f"❌ Error al cargar la planilla: {str(e)}")
            return None

# [Mantener las funciones de formato y conversión existentes...]

# --- 6. Funciones que actualizan la hoja de cálculo con métricas ---
def update_decimal_from_dms(sheet):
    """Convierte DMS a decimal y actualiza las columnas correspondientes para Sondas."""
    try:
        with st.spinner("Aplicando formato a las celdas..."):
            apply_format(sheet)
            update_dms_format_column(sheet)
        
        dms_values = sheet.col_values(13)  # Columna M
        if len(dms_values) <= 1:
            st.warning("⚠️ No se encontraron datos en 'Ubicación sonda google maps'.")
            return
        
        # Inicializar contadores
        total_registros = len(dms_values) - 1  # Excluir encabezado
        conversiones_exitosas = 0
        conversiones_fallidas = 0
        
        with st.spinner("Procesando conversiones DMS a decimal..."):
            progress_bar = st.progress(0)
            num_rows = len(dms_values)
            lat_cells = sheet.range(f"N2:N{num_rows}")
            lon_cells = sheet.range(f"O2:O{num_rows}")
            
            for i, dms in enumerate(dms_values[1:]):
                if dms:
                    result = dms_to_decimal(dms)
                    if result is not None:
                        lat, lon = result
                        lat_cells[i].value = round(lat, 8)
                        lon_cells[i].value = round(lon, 8)
                        conversiones_exitosas += 1
                    else:
                        conversiones_fallidas += 1
                progress_bar.progress((i + 1) / total_registros)
            
            sheet.update_cells(lat_cells)
            sheet.update_cells(lon_cells)
        
        # Mostrar métricas
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total de Registros", total_registros)
        with col2:
            st.metric("Conversiones Exitosas", conversiones_exitosas)
        with col3:
            st.metric("Conversiones Fallidas", conversiones_fallidas)
        
        st.success("✅ Conversión de DMS a decimal completada")
        
    except Exception as e:
        st.error(f"❌ Error en la conversión de DMS a decimal: {str(e)}")

# [Implementar visualización similar para las otras funciones de actualización...]

# --- 7. Interfaz de usuario en Streamlit ---
def main():
    st.title("Conversión de Coordenadas: Sondas📍")
    
    # Inicialización
    with st.spinner("Iniciando aplicación..."):
        client = init_connection()
        if not client:
            return
        sheet = load_sheet(client)
        if not sheet:
            return
    
    # Estado del sistema
    system_status = st.sidebar.container()
    with system_status:
        st.subheader("📊 Estado del Sistema")
        st.write("Conexión: ✅ Activa")
        st.write("Última actualización: " + pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"))
    
    # Interfaz principal
    st.write("Selecciona la conversión que deseas realizar:")
    
    # Sondas
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Convertir DMS a Decimal (Sonda)", 
                     help="Convierte las coordenadas DMS a formato decimal",
                     key="dms_to_decimal",
                     use_container_width=True):
            update_decimal_from_dms(sheet)
    with col2:
        if st.button("Convertir Decimal a DMS (Sonda)",
                     help="Convierte las coordenadas decimales a formato DMS",
                     key="decimal_to_dms",
                     use_container_width=True):
            update_dms_from_decimal(sheet)
    
    st.markdown("---")
    
    # Campo
    st.title("Conversión de Coordenadas: Campo 📍")
    st.write("Selecciona la conversión que deseas realizar:")
    
    col3, col4 = st.columns(2)
    with col3:
        if st.button("Convertir DMS a Decimal (Campo)",
                     help="Convierte las coordenadas DMS a formato decimal para Ubicación campo",
                     key="dms_to_decimal_field",
                     use_container_width=True):
            update_decimal_from_dms_field(sheet)
    with col4:
        if st.button("Convertir Decimal a DMS (Campo)",
                     help="Convierte las coordenadas decimales a formato DMS para Ubicación campo",
                     key="decimal_to_dms_field",
                     use_container_width=True):
            update_dms_from_decimal_field(sheet)

if __name__ == "__main__":
    main()
