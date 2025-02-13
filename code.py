def dms_a_decimal(dms):
    """
    Convierte coordenadas DMS a decimal y devuelve el formato corregido en DMS.
    Formato de salida: 34°22'05.6"S 71°01'53.0"W
    """
    match = re.match(
        r"(\d{1,3})°\s*(\d{1,2})'(\d+(?:\.\d+)?)\"?\s*([NSWE])\s+"
        r"(\d{1,3})°\s*(\d{1,2})'(\d+(?:\.\d+)?)\"?\s*([NSWE])",
        dms
    )
    if not match:
        return None, None, None

    # Extraer los valores de latitud y longitud
    lat_grados, lat_min, lat_seg, lat_dir, lon_grados, lon_min, lon_seg, lon_dir = match.groups()
    
    # Convertir a decimal
    lat_decimal = float(lat_grados) + float(lat_min) / 60 + float(lat_seg) / 3600
    lon_decimal = float(lon_grados) + float(lon_min) / 60 + float(lon_seg) / 3600

    # Aplicar signos según la dirección
    if lat_dir == "S":
        lat_decimal = -lat_decimal
    if lon_dir == "W":
        lon_decimal = -lon_decimal

    # Formatear DMS en el formato específico requerido
    lat_dms = format_dms(abs(lat_decimal), "S" if lat_decimal < 0 else "N")
    lon_dms = format_dms(abs(lon_decimal), "W" if lon_decimal < 0 else "E")

    return lat_decimal, lon_decimal, f"{lat_dms} {lon_dms}"

def decimal_a_dms(decimal, direccion):
    """
    Convierte coordenadas en decimal a DMS con formato específico.
    Formato de salida: 34°22'05.6"S o 71°01'53.0"W
    """
    return format_dms(abs(decimal), direccion)

def format_dms(decimal, direccion):
    """
    Formatea un valor decimal a DMS en el formato específico:
    - Grados sin ceros a la izquierda
    - Minutos con dos dígitos (01-59)
    - Segundos con un decimal
    """
    grados = int(decimal)
    minutos_temp = (decimal - grados) * 60
    minutos = int(minutos_temp)
    segundos = (minutos_temp - minutos) * 60

    # Formatear con un solo decimal en segundos
    return f"{grados}°{minutos:02d}'{segundos:.1f}\"{direccion}"
