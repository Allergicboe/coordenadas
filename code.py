import streamlit as st
import gspread
from google.oauth2 import service_account
import re

# --- 1. Interfaz de Streamlit ---
def main():
    """Interfaz principal de Streamlit."""
    st.title("Conversión de Coordenadas DMS a Decimal")

    # Entrada de usuario para iniciar la conversión
    st.write("""
    Esta aplicación convierte coordenadas DMS (grados, minutos, segundos) a formato decimal.
    Puedes subir las coordenadas DMS a tu hoja de Google Sheets, luego presiona el botón para actualizar las celdas con las coordenadas convertidas.
    """)

    # Mostrar el estado de conexión
    if st.button('Conectar y Convertir Coordenadas'):
        client = init_connection()
        if client is None:
            return
        sheet = load_sheet(client)
        if sheet is None:
            return
        procesar_coordenadas(sheet)

# --- 2. Funciones de Conexión y Carga de Datos ---
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

# --- 3. Funciones para Formatear y Convertir las Coordenadas ---
def formatear_dms(dms):
    match = re.match(r"(\d{1,2})°\s*(\d{1,2})'\s*([\d.]+)\"\s*([NS])\s*(\d{1,3})°\s*(\d{1,2})'\s*([\d.]+)\"\s*([EW])", str(dms))
    if not match:
        return None

    lat_g, lat_m, lat_s, lat_dir, lon_g, lon_m, lon_s, lon_dir = match.groups()

    # Redondear segundos a un decimal
    lat_s = round(float(lat_s), 1)
    lon_s = round(float(lon_s), 1)

    # Si al redondear, los segundos se vuelven 60.0, ajustar
    if lat_s == 60.0:
        lat_s = 0.0
        lat_m = int(lat_m) + 1

    if lon_s == 60.0:
        lon_s = 0.0
        lon_m = int(lon_m) + 1

    # Asegurar formato de salida correcto
    lat = f"{int(lat_g):02d}°{int(lat_m):02d}'{lat_s:04.1f}\"{lat_dir}"

    # Corregir longitud para eliminar el cero adicional en los grados
    lon = f"{int(lon_g)}°{int(lon_m):02d}'{lon_s:04.1f}\"{lon_dir}"

    return lat, lon

def dms_a_decimal(dms):
    match = re.match(r"(\d{1,3})°(\d{1,2})'([\d.]+)\"([NSWE])", str(dms))
    if not match:
        return None

    grados, minutos, segundos, direccion = match.groups()
    decimal = float(grados) + float(minutos) / 60 + float(segundos) / 3600
    if direccion in ['S', 'W']:
        decimal = -decimal

    return round(decimal, 8)

# --- 4. Función para Procesar las Coordenadas ---
def procesar_coordenadas(sheet):
    """Función para procesar las coordenadas DMS y actualizar la planilla."""
    # Obtener todos los datos
    datos = sheet.get_all_values()
    header = datos[0]
    data = datos[1:]

    # Obtener índices de las columnas necesarias
    col_m = header.index("Ubicación sonda google maps")
    col_n = header.index("Latitud sonda")
    col_o = header.index("longitud Sonda")

    # Listas para almacenar los valores a actualizar
    updates = []

    # Procesar cada fila
    for i, fila in enumerate(data, start=2):  # Comienza en la fila 2 (índice 1 en listas)
        dms_sonda = fila[col_m].strip() if col_m < len(fila) else ""

        # Ignorar valores vacíos o inválidos
        if not dms_sonda or not re.search(r"\d+°\s*\d+'", dms_sonda):
            continue

        # Formatear coordenadas y convertirlas a decimal
        coordenadas = formatear_dms(dms_sonda)
        if coordenadas:
            lat_decimal = dms_a_decimal(coordenadas[0])
            lon_decimal = dms_a_decimal(coordenadas[1])
        else:
            lat_decimal = lon_decimal = ""

        # Agregar las actualizaciones a la lista
        updates.append({"range": f"M{i}", "values": [[f"{coordenadas[0]} {coordenadas[1]}"]]})  # Ubicación formateada
        updates.append({"range": f"N{i}", "values": [[lat_decimal]]})  # Latitud decimal
        updates.append({"range": f"O{i}", "values": [[lon_decimal]]})  # Longitud decimal

    # Aplicar batch update
    if updates:
        sheet.batch_update(updates)

    st.success("✅ Conversión completada y planilla actualizada.")

if __name__ == "__main__":
    main()
