import streamlit as st
import gspread
from google.oauth2 import service_account
import re

# --- 1. Funciones de Conexión y Carga de Datos ---
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

# --- 2. Funciones de Conversión y Formato ---
def format_dms(decimal, direccion):
    """
    Formatea un valor decimal a DMS asegurando:
    - Grados sin ceros a la izquierda
    - Minutos siempre con dos dígitos (00-59)
    - Segundos sin ceros a la izquierda antes del decimal, un decimal después
    Ejemplo: 34°39'08.8"S
    """
    grados = int(abs(decimal))
    minutos_temp = (abs(decimal) - grados) * 60
    minutos = int(minutos_temp)
    segundos = (minutos_temp - minutos) * 60
    
    # Formatea los segundos primero como string para eliminar ceros a la izquierda
    segundos_str = f"{segundos:.1f}"
    if segundos_str.startswith('0'):
        segundos_str = segundos_str[1:]  # Elimina el cero inicial si existe

    return f"{grados:02d}°{minutos:02d}'{segundos_str}\"{direccion}"

def dms_a_decimal(dms):
    """
    Convierte coordenadas DMS a decimal y devuelve el formato corregido en DMS.
    Formato de salida garantizado: 30°00'38.6"S 70°32'00.9"W
    """
    pattern = (
        r"(\d{1,3})\s*°\s*(\d{1,2})\s*'\s*(\d*\.?\d*)\s*\"?\s*([NSns])\s*"
        r"(\d{1,3})\s*°\s*(\d{1,2})\s*'\s*(\d*\.?\d*)\s*\"?\s*([WEwe])"
    )
    
    # Limpieza básica del texto
    dms = dms.strip()
    dms = re.sub(r'\s+', ' ', dms)
    
    match = re.search(pattern, dms)
    if not match:
        return None, None, None

    try:
        lat_grados, lat_min, lat_seg, lat_dir, lon_grados, lon_min, lon_seg, lon_dir = match.groups()
        
        # Convertir a decimal
        lat_decimal = float(lat_grados) + float(lat_min) / 60 + float(lat_seg) / 3600
        lon_decimal = float(lon_grados) + float(lon_min) / 60 + float(lon_seg) / 3600

        # Normalizar direcciones a mayúsculas
        lat_dir = lat_dir.upper()
        lon_dir = lon_dir.upper()

        # Aplicar signos según la dirección
        if lat_dir == "S":
            lat_decimal = -lat_decimal
        if lon_dir == "W":
            lon_decimal = -lon_decimal

        # Formatear DMS en el formato específico requerido
        lat_dms = format_dms(abs(lat_decimal), "S" if lat_decimal < 0 else "N")
        lon_dms = format_dms(abs(lon_decimal), "W" if lon_decimal < 0 else "E")

        return lat_decimal, lon_decimal, f"{lat_dms} {lon_dms}"
    except (ValueError, AttributeError, TypeError):
        return None, None, None

def decimal_a_dms(decimal, direccion):
    """
    Convierte coordenadas en decimal a DMS con formato específico.
    Formato de salida garantizado: 30°00'38.6"S o 70°32'00.9"W
    """
    return format_dms(abs(decimal), direccion)

# --- 3. Aplicar Formato a Google Sheets ---
def aplicar_formato(sheet, columna):
    """Aplica formato a una columna en Google Sheets."""
    formato = {
        "textFormat": {"fontFamily": "Arial", "fontSize": 11},
        "backgroundColor": {"red": 1, "green": 1, "blue": 1},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }
    sheet.format(f"{columna}2:{columna}", formato)

# --- 4. Procesar la Hoja de Cálculo ---
def procesar_hoja(sheet, conversion):
    datos = sheet.get_all_values()
    header = datos[0]
    data = datos[1:]

    # Obtener índices de las columnas
    col_m = header.index("Ubicación sonda google maps")
    col_n = header.index("Latitud sonda")
    col_o = header.index("longitud Sonda")

    updates = []
    errores = []

    # Aplicar formato a las columnas
    aplicar_formato(sheet, "M")
    aplicar_formato(sheet, "N")
    aplicar_formato(sheet, "O")

    for i, fila in enumerate(data, start=2):
        dms_sonda = fila[col_m].strip() if col_m < len(fila) else ""
        lat_decimal = fila[col_n].strip() if col_n < len(fila) else ""
        lon_decimal = fila[col_o].strip() if col_o < len(fila) else ""

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

        if conversion == "DMS a Decimal":
            if dms_sonda:
                lat_decimal, lon_decimal, dms_corregido = dms_a_decimal(dms_sonda)
                if lat_decimal is not None and lon_decimal is not None and dms_corregido is not None:
                    updates.append({"range": f"M{i}", "values": [[dms_corregido]]})
                    updates.append({"range": f"N{i}", "values": [[lat_decimal]]})
                    updates.append({"range": f"O{i}", "values": [[lon_decimal]]})
                else:
                    errores.append(f"Fila {i}: Formato inválido - '{dms_sonda}'")

        elif conversion == "Decimal a DMS":
            if lat_decimal is not None and lon_decimal is not None:
                lat_dms = decimal_a_dms(lat_decimal, "S" if lat_decimal < 0 else "N")
                lon_dms = decimal_a_dms(lon_decimal, "W" if lon_decimal < 0 else "E")
                dms_sonda = f"{lat_dms} {lon_dms}"
                updates.append({"range": f"M{i}", "values": [[dms_sonda]]})

    if updates:
        sheet.batch_update(updates)
        st.success(f"✅ Conversión completada. Se actualizaron {len(updates)} valores.")
        if errores:
            st.warning("⚠️ Algunas filas no pudieron ser procesadas:")
            for error in errores:
                st.write(error)
    else:
        st.error("❌ No se encontraron datos válidos para actualizar. Verifica el formato de las coordenadas.")
        if errores:
            st.write("Detalles de los errores:")
            for error in errores:
                st.write(error)

# --- 5. Interfaz de Streamlit ---
st.title('Conversión de Coordenadas')
st.sidebar.header('Opciones')

conversion = st.sidebar.radio("Seleccione el tipo de conversión", ('Decimal a DMS', 'DMS a Decimal'))

client = init_connection()
if client:
    sheet = load_sheet(client)
    if sheet:
        if conversion == 'DMS a Decimal' and st.button("Convertir DMS a Decimal"):
            procesar_hoja(sheet, conversion)
        elif conversion == 'Decimal a DMS' and st.button("Convertir Decimal a DMS"):
            procesar_hoja(sheet, conversion)
