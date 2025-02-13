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

# --- 5. Interfaz de Usuario ---
st.title("Conversión de Coordenadas: DMS a Decimal y Decimal a DMS")

# --- 5.1. Sección DMS a Decimal ---
st.subheader("De DMS a Decimal")

# Entrada de coordenadas DMS
dms_input = st.text_input("Ingresa las coordenadas DMS (Ejemplo: 45° 30' 15\" N, 67° 10' 30\" W)")

if st.button("Convertir DMS a Decimal"):
    if dms_input:
        # Separar las coordenadas
        lat_dms = dms_input.split(',')[0].strip()
        lon_dms = dms_input.split(',')[1].strip()
        
        lat_decimal = dms_a_decimal(lat_dms)
        lon_decimal = dms_a_decimal(lon_dms)

        if lat_decimal and lon_decimal:
            st.write(f"Latitud en Decimal: {lat_decimal}")
            st.write(f"Longitud en Decimal: {lon_decimal}")
        else:
            st.error("Formato DMS no válido, intenta de nuevo.")
    else:
        st.error("Por favor ingresa las coordenadas DMS.")

# --- 5.2. Sección Decimal a DMS ---
st.subheader("De Decimal a DMS")

# Entrada de coordenadas en formato decimal
lat_decimal_input = st.number_input("Ingresa la latitud en formato decimal", format="%.8f")
lon_decimal_input = st.number_input("Ingresa la longitud en formato decimal", format="%.8f")

if st.button("Convertir Decimal a DMS"):
    if lat_decimal_input and lon_decimal_input:
        lat_dms = decimal_a_dms(lat_decimal_input, "N" if lat_decimal_input >= 0 else "S")
        lon_dms = decimal_a_dms(lon_decimal_input, "E" if lon_decimal_input >= 0 else "W")
        
        st.write(f"Latitud en DMS: {lat_dms}")
        st.write(f"Longitud en DMS: {lon_dms}")
    else:
        st.error("Por favor ingresa las coordenadas en formato decimal.")
