import streamlit as st
import gspread
from google.oauth2 import service_account
import re

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

# --- 3. Función para Convertir Decimal a DMS ---
def decimal_a_dms(decimal, direccion):
    grados = int(decimal)
    minutos = int((decimal - grados) * 60)
    segundos = round((((decimal - grados) * 60) - minutos) * 60, 1)
    return f"{grados}° {minutos}' {segundos}\" {direccion}"

# --- 4. Función para Convertir DMS a Decimal ---
def dms_a_decimal(dms):
    match = re.match(r"(\d{1,3})°\s*(\d{1,2})'\s*([\d.]+)\"\s*([NSWE])", str(dms))
    if not match:
        return None

    grados, minutos, segundos, direccion = match.groups()
    decimal = float(grados) + float(minutos) / 60 + float(segundos) / 3600
    if direccion in ['S', 'W']:
        decimal = -decimal

    return round(decimal, 8)

# --- 5. Función para obtener datos de Google Sheets ---
def obtener_datos(sheet):
    datos = sheet.get_all_values()
    header = datos[0]
    data = datos[1:]

    col_m = header.index("Ubicación sonda google maps")
    col_n = header.index("Latitud sonda")
    col_o = header.index("Longitud Sonda")
    
    return header, data, col_m, col_n, col_o

# --- 6. Función para completar "Ubicación sonda google maps" (M) desde "Latitud sonda" (N) y "Longitud Sonda" (O) ---
def completar_dms(header, data, col_n, col_o):
    updates = []
    for i, fila in enumerate(data, start=2):
        lat_decimal = fila[col_n]
        lon_decimal = fila[col_o]
        
        if lat_decimal and lon_decimal:
            # Convertir de decimal a DMS
            lat_dms = decimal_a_dms(float(lat_decimal), "N" if float(lat_decimal) >= 0 else "S")
            lon_dms = decimal_a_dms(float(lon_decimal), "E" if float(lon_decimal) >= 0 else "W")
            
            # Preparar las actualizaciones
            updates.append({"range": f"M{i}", "values": [[f"{lat_dms}, {lon_dms}"]]})
    
    return updates

# --- 7. Función para completar "Latitud Sonda" (N) y "Longitud Sonda" (O) desde "Ubicación sonda google maps" (M) ---
def completar_decimal(header, data, col_m):
    updates = []
    for i, fila in enumerate(data, start=2):
        dms_sonda = fila[col_m]
        
        if dms_sonda:
            # Convertir de DMS a decimal
            coordenadas = dms_sonda.split(',')
            if len(coordenadas) == 2:
                lat_decimal = dms_a_decimal(coordenadas[0].strip())
                lon_decimal = dms_a_decimal(coordenadas[1].strip())
                
                # Preparar las actualizaciones
                updates.append({"range": f"N{i}", "values": [[lat_decimal]]})
                updates.append({"range": f"O{i}", "values": [[lon_decimal]]})

    return updates

# --- 8. Interfaz de Usuario ---
st.title("Conversión de Coordenadas: DMS a Decimal y Decimal a DMS")

# Conexión a Google Sheets
client = init_connection()
if client:
    sheet = load_sheet(client)
    if sheet:
        header, data, col_m, col_n, col_o = obtener_datos(sheet)

        # Sección para completar "Ubicación Sonda" (M) desde "Latitud" y "Longitud"
        st.subheader("Convertir de Latitud y Longitud a Ubicación Sonda (M)")
        if st.button("Convertir Latitud y Longitud a Ubicación Sonda"):
            updates = completar_dms(header, data, col_n, col_o)
            if updates:
                sheet.batch_update(updates)
                st.success("¡Ubicación Sonda completada!")

        # Sección para completar "Latitud" y "Longitud" desde "Ubicación Sonda"
        st.subheader("Convertir de Ubicación Sonda (M) a Latitud y Longitud")
        if st.button("Convertir Ubicación Sonda a Latitud y Longitud"):
            updates = completar_decimal(header, data, col_m)
            if updates:
                sheet.batch_update(updates)
                st.success("¡Latitud y Longitud completadas!")
