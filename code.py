import gspread
from google.oauth2 import service_account
import re
import streamlit as st

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

# Función para convertir decimal a DMS con formato fijo
def decimal_a_dms(decimal, direccion):
    """Convierte coordenadas decimales a formato DMS."""
    if not decimal:
        return ""

    grados = abs(int(decimal))
    minutos = int((abs(decimal) - grados) * 60)
    segundos = round(((abs(decimal) - grados) * 60 - minutos) * 60, 1)

    # Asegurar que los minutos y segundos siempre tengan dos dígitos
    dms = f"{grados:02d}°{minutos:02d}'{segundos:04.1f}\"{direccion}"

    return dms

# Función para convertir DMS a decimal
def dms_a_decimal(dms):
    """Convierte coordenadas DMS a formato decimal."""
    match = re.match(r"(\d{1,3})°(\d{1,2})'([\d.]+)\"([NSWE])", str(dms))
    if not match:
        return None

    grados, minutos, segundos, direccion = match.groups()
    decimal = float(grados) + float(minutos) / 60 + float(segundos) / 3600
    if direccion in ['S', 'W']:
        decimal = -decimal

    return round(decimal, 8)

# Función para procesar y actualizar las celdas de Google Sheets
def procesar_hoja(sheet):
    """Procesa las celdas de Google Sheets y convierte coordenadas."""
    # Obtener los datos de la hoja
    datos = sheet.get_all_values()
    header = datos[0]
    data = datos[1:]

    # Obtener los índices de las columnas necesarias
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

        # Ignorar valores vacíos o inválidos
        if not dms_sonda and not lat_decimal and not lon_decimal:
            continue

        # Convertir de DMS a Decimal si la celda de DMS está llena
        if dms_sonda:
            lat_decimal = dms_a_decimal(dms_sonda.split(" ")[0])  # Latitud
            lon_decimal = dms_a_decimal(dms_sonda.split(" ")[1])  # Longitud

        # Convertir de Decimal a DMS si las celdas de latitud y longitud están llenas
        if lat_decimal and lon_decimal:
            lat_dms = decimal_a_dms(float(lat_decimal), "S" if float(lat_decimal) < 0 else "N")
            lon_dms = decimal_a_dms(float(lon_decimal), "W" if float(lon_decimal) < 0 else "E")

            # Actualizar la hoja con los nuevos valores de DMS
            updates.append({"range": f"M{i}", "values": [[f"{lat_dms} {lon_dms}"]]})  # Ubicación formateada
            updates.append({"range": f"N{i}", "values": [[lat_decimal]]})  # Latitud decimal
            updates.append({"range": f"O{i}", "values": [[lon_decimal]]})  # Longitud decimal

    # Aplicar batch update
    if updates:
        sheet.batch_update(updates)

# --- Interfaz de Streamlit ---
def main():
    # Conectar con Google Sheets
    client = init_connection()
    if not client:
        return
    sheet = load_sheet(client)
    if not sheet:
        return

    # Botón para convertir DMS a decimal
    if st.button("Convertir DMS a Decimal"):
        procesar_hoja(sheet)
        st.success("Coordenadas convertidas de DMS a decimal y actualizadas.")

    # Botón para convertir Decimal a DMS
    if st.button("Convertir Decimal a DMS"):
        procesar_hoja(sheet)
        st.success("Coordenadas convertidas de decimal a DMS y actualizadas.")

if __name__ == "__main__":
    main()
