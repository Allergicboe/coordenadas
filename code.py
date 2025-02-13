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

# Función para convertir DMS a decimal
def dms_a_decimal(dms):
    match = re.match(r"(\d{1,3})°(\d{1,2})'([\d.]+)\"([NSWE])", str(dms))
    if not match:
        return None

    grados, minutos, segundos, direccion = match.groups()
    decimal = float(grados) + float(minutos) / 60 + float(segundos) / 3600
    if direccion in ['S', 'W']:
        decimal = -decimal

    return round(decimal, 8)

# Función para convertir de decimal a DMS sin espacios
def decimal_a_dms(decimal, direccion):
    grados = int(abs(decimal))
    minutos = int((abs(decimal) - grados) * 60)
    segundos = (abs(decimal) - grados - minutos / 60) * 3600

    # Formatear a DMS sin espacios
    dms = f"{grados:02d}°{minutos:02d}'{segundos:0.1f}\"{direccion}"
    return dms

# Función para procesar la hoja y realizar la conversión
def procesar_hoja(sheet, conversion):
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
        lat_decimal = fila[col_n].strip() if col_n < len(fila) else ""
        lon_decimal = fila[col_o].strip() if col_o < len(fila) else ""

        # Validación para evitar convertir valores no numéricos o vacíos
        try:
            if lat_decimal:
                lat_decimal = float(lat_decimal.replace(",", "."))
            else:
                lat_decimal = None

            if lon_decimal:
                lon_decimal = float(lon_decimal.replace(",", "."))
            else:
                lon_decimal = None
        except ValueError:
            lat_decimal = None
            lon_decimal = None

        # Si la conversión es de DMS a Decimal
        if conversion == "DMS a Decimal":
            if dms_sonda and re.search(r"\d+°\s*\d+'", dms_sonda):
                lat_decimal = dms_a_decimal(dms_sonda)
                lon_decimal = lat_decimal  # Ya que tenemos un solo valor decimal por DMS

            # Agregar las actualizaciones a la lista
            updates.append({"range": f"N{i}", "values": [[lat_decimal]]})  # Latitud decimal
            updates.append({"range": f"O{i}", "values": [[lon_decimal]]})  # Longitud decimal

        # Si la conversión es de Decimal a DMS
        elif conversion == "Decimal a DMS":
            if lat_decimal is not None and lon_decimal is not None:
                lat_dms = decimal_a_dms(lat_decimal, "S" if lat_decimal < 0 else "N")
                lon_dms = decimal_a_dms(lon_decimal, "W" if lon_decimal < 0 else "E")
                dms_sonda = f"{lat_dms} {lon_dms}"

                # Agregar las actualizaciones a la lista para reemplazar los valores en DMS
                updates.append({"range": f"M{i}", "values": [[dms_sonda]]})  # Ubicación formateada

    # Aplicar batch update
    if updates:
        sheet.batch_update(updates)
        st.success("✅ Conversión completada y planilla actualizada.")
    else:
        st.warning("⚠️ No se encontraron datos válidos para actualizar.")

# --- Interfaz de Streamlit --- 
st.title('Conversión de Coordenadas')
st.markdown("""
**Selecciona el tipo de conversión y realiza la acción:**

- **DMS a Decimal:** Convierte coordenadas DMS a formato decimal.
- **Decimal a DMS:** Convierte coordenadas en formato decimal a DMS.
""")

# --- Opciones de conversión ---
conversion = st.radio(
    "¿Qué tipo de conversión deseas realizar?",
    ('DMS a Decimal', 'Decimal a DMS')
)

# Conectar a Google Sheets
client = init_connection()
if client:
    sheet = load_sheet(client)

    if sheet:
        st.markdown("### Proceso de conversión")
        st.markdown(f"Seleccionaste: **{conversion}**")
        
        if st.button("Ejecutar conversión"):
            procesar_hoja(sheet, conversion)
