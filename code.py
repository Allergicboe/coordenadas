import streamlit as st
import gspread
from google.oauth2 import service_account
import re

# --- Conexión con Google Sheets ---
def init_connection():
    """Inicializa la conexión con Google Sheets."""
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
    """Carga la hoja de trabajo de Google Sheets."""
    try:
        return client.open_by_url(st.secrets["spreadsheet_url"]).sheet1
    except Exception as e:
        st.error(f"Error al cargar la planilla: {str(e)}")
        return None

# --- Funciones de conversión ---
def dms_a_decimal(dms):
    """Convierte coordenadas en formato DMS a Decimal."""
    match = re.match(r"(\d{1,3})°\s*(\d{1,2})'(\d+(\.\d+)?)\"([NSWE])", str(dms))
    if not match:
        return None

    grados, minutos, segundos, _, direccion = match.groups()
    decimal = float(grados) + float(minutos) / 60 + round(float(segundos), 1) / 3600  # Redondeo correcto de segundos
    if direccion in ['S', 'W']:
        decimal = -decimal

    return round(decimal, 8)

def decimal_a_dms(decimal, direccion):
    """Convierte coordenadas en formato Decimal a DMS."""
    grados = int(abs(decimal))
    minutos = int((abs(decimal) - grados) * 60)
    segundos = (abs(decimal) - grados - minutos / 60) * 3600

    # Redondear segundos a un solo decimal
    segundos = round(segundos, 1)

    # Formato corregido sin espacios extra
    dms = f"{grados:02d}°{int(minutos):02d}'{segundos:04.1f}\"{direccion}"
    return dms

def formatear_estilo(sheet, col_letra):
    """Aplica formato a toda la columna especificada."""
    sheet.format(f'{col_letra}2:{col_letra}', {
        "horizontalAlignment": "CENTER",
        "textFormat": {
            "fontSize": 11,
            "fontFamily": "Arial",
            "bold": False
        },
        "backgroundColor": {
            "red": 1.0, "green": 1.0, "blue": 1.0  # Sin relleno (blanco)
        },
        "foregroundColorStyle": {  # Color de fuente negro
            "rgbColor": {"red": 0, "green": 0, "blue": 0}
        }
    })

def procesar_hoja(sheet, conversion):
    datos = sheet.get_all_values()
    header = datos[0]
    data = datos[1:]

    # Índices de columnas (respetando mayúsculas)
    col_m = header.index("Ubicación sonda google maps")
    col_n = header.index("Latitud sonda")
    col_o = header.index("longitud Sonda")

    updates = []

    for i, fila in enumerate(data, start=2):  # Comienza en la fila 2 (índice 1 en listas)
        dms_sonda = fila[col_m].strip() if col_m < len(fila) else ""
        lat_decimal_sonda = fila[col_n].strip() if col_n < len(fila) else ""
        lon_decimal_sonda = fila[col_o].strip() if col_o < len(fila) else ""

        try:
            lat_decimal_sonda = float(lat_decimal_sonda.replace(",", ".")) if lat_decimal_sonda else None
            lon_decimal_sonda = float(lon_decimal_sonda.replace(",", ".")) if lon_decimal_sonda else None
        except ValueError:
            lat_decimal_sonda, lon_decimal_sonda = None, None

        # Conversión de DMS a Decimal
        if conversion == "DMS a Decimal":
            if dms_sonda and re.search(r"\d+°\s*\d+'", dms_sonda):
                lat_decimal_sonda = dms_a_decimal(dms_sonda)
                lon_decimal_sonda = lat_decimal_sonda  # Un solo valor decimal por DMS

                updates.append({"range": f"N{i}", "values": [[lat_decimal_sonda]]})
                updates.append({"range": f"O{i}", "values": [[lon_decimal_sonda]]})

        # Conversión de Decimal a DMS
        elif conversion == "Decimal a DMS":
            if lat_decimal_sonda is not None and lon_decimal_sonda is not None:
                lat_dms_sonda = decimal_a_dms(lat_decimal_sonda, "S" if lat_decimal_sonda < 0 else "N")
                lon_dms_sonda = decimal_a_dms(lon_decimal_sonda, "W" if lon_decimal_sonda < 0 else "E")
                dms_sonda = f"{lat_dms_sonda} {lon_dms_sonda}"

                updates.append({"range": f"M{i}", "values": [[dms_sonda]]})

    if updates:
        sheet.batch_update(updates)
        st.success("✅ Conversión completada y planilla actualizada.")
    else:
        st.warning("⚠️ No se encontraron datos válidos para actualizar.")

    # Aplicar formato a las columnas
    formatear_estilo(sheet, "M")  # Ubicación Sonda
    formatear_estilo(sheet, "N")  # Latitud Sonda
    formatear_estilo(sheet, "O")  # Longitud Sonda

# --- Interfaz de Streamlit ---
st.title('Conversión de Coordenadas - Sonda')
st.sidebar.header('Opciones')

conversion = st.sidebar.radio("Seleccione el tipo de conversión", ('Decimal a DMS', 'DMS a Decimal'))

client = init_connection()
if client:
    sheet = load_sheet(client)

    if sheet:
        if conversion == 'DMS a Decimal':
            if st.button("Convertir DMS a Decimal"):
                procesar_hoja(sheet, conversion)
        elif conversion == 'Decimal a DMS':
            if st.button("Convertir Decimal a DMS"):
                procesar_hoja(sheet, conversion)
