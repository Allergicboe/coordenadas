import streamlit as st
import gspread
from google.oauth2 import service_account
import re

# --- 1. Función para inicializar la conexión con Google Sheets ---
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

# --- 2. Función para cargar la hoja de trabajo ---
def load_sheet(client):
    """Función para cargar la hoja de trabajo de Google Sheets."""
    try:
        return client.open_by_url(st.secrets["spreadsheet_url"]).sheet1
    except Exception as e:
        st.error(f"Error al cargar la planilla: {str(e)}")
        return None

# --- 3. Funciones de conversión DMS a Decimal y viceversa ---
def dms_a_decimal(dms):
    """Convierte coordenadas DMS a formato decimal."""
    match = re.match(r"(\d{1,3})°\s*(\d{1,2})'\s*([\d.]+)\"\s*([NSWE])", str(dms))
    if not match:
        return None

    grados, minutos, segundos, direccion = match.groups()
    decimal = float(grados) + float(minutos) / 60 + float(segundos) / 3600
    if direccion in ['S', 'W']:
        decimal = -decimal

    return round(decimal, 8)

def decimal_a_dms(decimal, direccion):
    """Convierte coordenadas decimales a formato DMS."""
    grados = int(decimal)
    minutos = int((decimal - grados) * 60)
    segundos = round(((decimal - grados) * 60 - minutos) * 60, 4)
    
    return f"{grados}° {minutos}' {segundos}\" {direccion}"

# --- 4. Interfaz de usuario con Streamlit ---
st.title("Conversión de Coordenadas entre DMS y Decimal")

# Conexión a Google Sheets
client = init_connection()
sheet = load_sheet(client)

# Mostrar mensaje si no hay hoja
if not sheet:
    st.stop()

# Opciones de conversión
conversion_opcion = st.selectbox("Selecciona el tipo de conversión", ["De DMS a Decimal", "De Decimal a DMS"])

if conversion_opcion == "De DMS a Decimal":
    st.header("Conversión de DMS a Decimal")
    
    # Inputs para las coordenadas DMS
    dms_lat = st.text_input("Latitud (DMS):")
    dms_lon = st.text_input("Longitud (DMS):")
    
    # Botón para realizar la conversión
    if st.button("Convertir a Decimal"):
        if dms_lat and dms_lon:
            lat_decimal = dms_a_decimal(dms_lat)
            lon_decimal = dms_a_decimal(dms_lon)
            
            if lat_decimal and lon_decimal:
                st.success(f"Latitud: {lat_decimal}, Longitud: {lon_decimal}")
            else:
                st.error("Error en el formato de las coordenadas DMS")
        else:
            st.error("Por favor, ingrese ambas coordenadas DMS.")

elif conversion_opcion == "De Decimal a DMS":
    st.header("Conversión de Decimal a DMS")
    
    # Inputs para las coordenadas decimales
    decimal_lat = st.number_input("Latitud (Decimal):", format="%.8f")
    decimal_lon = st.number_input("Longitud (Decimal):", format="%.8f")
    
    # Botón para realizar la conversión
    if st.button("Convertir a DMS"):
        if decimal_lat and decimal_lon:
            lat_dms = decimal_a_dms(decimal_lat, 'S' if decimal_lat < 0 else 'N')
            lon_dms = decimal_a_dms(decimal_lon, 'W' if decimal_lon < 0 else 'E')
            
            st.success(f"Latitud: {lat_dms}, Longitud: {lon_dms}")
        else:
            st.error("Por favor, ingrese ambas coordenadas decimales.")

# --- 5. Función para actualizar los valores en Google Sheets ---
def actualizar_coordenadas():
    """Función para actualizar coordenadas en Google Sheets."""
    try:
        # Obtener los datos de la hoja
        datos = sheet.get_all_values()
        header = datos[0]
        data = datos[1:]

        # Obtener los índices de las columnas relevantes
        col_m = header.index("Ubicación sonda google maps")
        col_n = header.index("Latitud sonda")
        col_o = header.index("longitud Sonda")

        updates = []

        for i, fila in enumerate(data, start=2):  # Comienza en la fila 2
            dms_sonda = fila[col_m].strip() if col_m < len(fila) else ""
            lat_decimal = fila[col_n].strip() if col_n < len(fila) else ""
            lon_decimal = fila[col_o].strip() if col_o < len(fila) else ""

            # Actualizar coordenadas de DMS a Decimal
            if dms_sonda:
                lat_decimal = dms_a_decimal(dms_sonda)
                lon_decimal = dms_a_decimal(dms_sonda)
                updates.append({"range": f"N{i}", "values": [[lat_decimal]]})  # Latitud decimal
                updates.append({"range": f"O{i}", "values": [[lon_decimal]]})  # Longitud decimal
            elif lat_decimal and lon_decimal:
                dms_lat = decimal_a_dms(float(lat_decimal), 'S' if float(lat_decimal) < 0 else 'N')
                dms_lon = decimal_a_dms(float(lon_decimal), 'W' if float(lon_decimal) < 0 else 'E')
                updates.append({"range": f"M{i}", "values": [[f"{dms_lat} {dms_lon}"]]})  # Ubicación en DMS
        
        # Aplicar actualización batch
        if updates:
            sheet.batch_update(updates)
            st.success("¡Coordenadas actualizadas exitosamente!")
        else:
            st.warning("No se encontraron actualizaciones.")
    except Exception as e:
        st.error(f"Error al actualizar la hoja: {str(e)}")

# Botón para actualizar las coordenadas
if st.button("Actualizar Coordenadas en Google Sheets"):
    actualizar_coordenadas()
