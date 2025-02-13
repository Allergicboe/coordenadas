import streamlit as st
import gspread
from google.oauth2 import service_account
import re

# --- Funciones de Conexión y Carga de Datos ---
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
    match = re.match(r"(\d{1,3})°\s*(\d{1,2})'\s*([\d,]+)\"\s*([NSWE])", str(dms))
    if not match:
        return None

    grados, minutos, segundos, direccion = match.groups()

    # Convertir a decimal
    decimal = float(grados) + float(minutos) / 60 + float(segundos.replace(",", ".")) / 3600
    if direccion in ['S', 'W']:
        decimal = -decimal

    return round(decimal, 8)

# Función para convertir de decimal a DMS
def decimal_a_dms(decimal, direccion):
    grados = int(abs(decimal))
    minutos = int((abs(decimal) - grados) * 60)
    segundos = (abs(decimal) - grados - minutos / 60) * 3600

    # Formato DMS
    dms = f"{grados}° {minutos}' {segundos:0.4f}\" {direccion}"
    return dms

# Interfaz de Streamlit
st.title("Conversión de Coordenadas: DMS a Decimal y viceversa")

# Cargar hoja de cálculo
client = init_connection()
if client:
    sheet = load_sheet(client)

# Botones y entradas para conversión
if sheet:
    st.sidebar.title("Selección de Conversión")
    option = st.sidebar.selectbox("Selecciona la conversión:", ["De DMS a Decimal", "De Decimal a DMS"])

    # Cargar datos
    datos = sheet.get_all_values()
    header = datos[0]
    data = datos[1:]

    # Indices de las columnas necesarias
    col_m = header.index("Ubicación sonda google maps")
    col_n = header.index("Latitud sonda")
    col_o = header.index("longitud Sonda")

    # Listas para almacenar las actualizaciones
    updates = []

    if option == "De DMS a Decimal":
        st.subheader("Convertir de DMS a Decimal")

        # Procesar cada fila para convertir DMS a decimal
        for i, fila in enumerate(data, start=2):  # Comienza en la fila 2 (índice 1 en listas)
            dms_sonda = fila[col_m].strip() if col_m < len(fila) else ""

            if dms_sonda:
                # Verificar que la cadena tenga el formato esperado
                dms_parts = dms_sonda.split(" ")
                if len(dms_parts) == 2:
                    lat_decimal = dms_a_decimal(dms_parts[0])
                    lon_decimal = dms_a_decimal(dms_parts[1])

                    # Verificar si se obtuvo un valor válido para la latitud y longitud
                    if lat_decimal is not None and lon_decimal is not None:
                        updates.append({"range": f"N{i}", "values": [[lat_decimal]]})  # Latitud decimal
                        updates.append({"range": f"O{i}", "values": [[lon_decimal]]})  # Longitud decimal

        # Aplicar batch update
        if updates:
            try:
                sheet.batch_update(updates)
                st.success("✅ Conversión de DMS a Decimal completada y planilla actualizada.")
            except Exception as e:
                st.error(f"Error al actualizar la hoja: {str(e)}")

    elif option == "De Decimal a DMS":
        st.subheader("Convertir de Decimal a DMS")

        # Procesar cada fila para convertir de decimal a DMS
        for i, fila in enumerate(data, start=2):  # Comienza en la fila 2 (índice 1 en listas)
            lat_decimal = fila[col_n].strip() if col_n < len(fila) else ""
            lon_decimal = fila[col_o].strip() if col_o < len(fila) else ""

            if lat_decimal and lon_decimal:
                # Convertir decimal a DMS
                lat_decimal = float(lat_decimal.replace(",", "."))
                lon_decimal = float(lon_decimal.replace(",", "."))
                lat_dms = decimal_a_dms(lat_decimal, "S" if lat_decimal < 0 else "N")
                lon_dms = decimal_a_dms(lon_decimal, "W" if lon_decimal < 0 else "E")

                # Actualizar la ubicación sonda en formato DMS
                updates.append({"range": f"M{i}", "values": [[f"{lat_dms} {lon_dms}"]]})  # Ubicación formateada

        # Aplicar batch update
        if updates:
            try:
                sheet.batch_update(updates)
                st.success("✅ Conversión de Decimal a DMS completada y planilla actualizada.")
            except Exception as e:
                st.error(f"Error al actualizar la hoja: {str(e)}")
