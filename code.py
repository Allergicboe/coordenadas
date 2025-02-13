import streamlit as st
from google.colab import auth
from google.auth import default
import gspread
import re

# Autenticación con Google Sheets
auth.authenticate_user()
creds, _ = default()
gc = gspread.authorize(creds)

# Obtener la hoja de cálculo
SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1_74Vt8KL0bscmSME5Evm6hn4DWytLdGDGb98tHyNwtc/edit?usp=drive_link'
sheet = gc.open_by_url(SPREADSHEET_URL).sheet1

# Función para formatear coordenadas DMS
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

# Título de la app
st.title("Conversión de Coordenadas DMS a Decimal")

# Cargar datos
datos = sheet.get_all_values()
header = datos[0]
data = datos[1:]

# Obtener índices de las columnas necesarias
col_m = header.index("Ubicación sonda google maps")
col_n = header.index("Latitud sonda")
col_o = header.index("longitud Sonda")

# Listas para almacenar los valores a actualizar
updates = []

# Mostrar información
st.write(f"Se han encontrado {len(data)} registros para procesar.")

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

# Confirmación de éxito
st.success("✅ Conversión completada y planilla actualizada.")
