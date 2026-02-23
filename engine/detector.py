"""
Auto-detección de formato de archivo: headers, línea (DAMA/CABALLERO), tallas.
"""
from .styles import safe_int


def detect_headers(first_row):
    """Detecta si la primera fila son headers o datos.
    Returns True si tiene headers, False si son datos directos.
    """
    if not first_row:
        return False
    first_val = first_row[0]
    if first_val is None:
        return False
    # Si el primer valor es string no-numérico -> headers
    s = str(first_val).strip()
    if s.upper() in ('STR', 'STORE', 'TIENDA'):
        return True
    try:
        int(first_val)
        return False  # Es un número -> datos
    except (ValueError, TypeError):
        return True  # Es texto -> headers


def detect_line_and_sizes(records):
    """Detecta la línea (DAMA/CABALLERO) y rango de tallas a partir de los registros.
    Returns (linea, tallas) donde tallas es lista ordenada de enteros.
    """
    tallas_vistas = set()
    lineas_vistas = set()

    for r in records:
        size = r.get('SIZE', 0)
        if 20 <= size <= 35:
            tallas_vistas.add(size)

    tallas_sorted = sorted(tallas_vistas)

    # Determinar línea por rango de tallas
    if not tallas_sorted:
        return "DESCONOCIDO", []

    min_t = min(tallas_sorted)
    max_t = max(tallas_sorted)

    if min_t >= 26:
        linea = "CABALLERO"
    elif max_t <= 26:
        linea = "DAMA"
    else:
        # Rango mixto: checar dónde está la mayoría
        tallas_bajas = sum(1 for t in tallas_sorted if t < 26)
        tallas_altas = sum(1 for t in tallas_sorted if t >= 26)
        linea = "DAMA" if tallas_bajas > tallas_altas else "CABALLERO"

    # Tallas principales: las 4 más frecuentes en el rango
    if linea == "CABALLERO":
        tallas = [t for t in tallas_sorted if 26 <= t <= 30]
    else:
        tallas = [t for t in tallas_sorted if 22 <= t <= 27]

    # Tomar las 4 principales si hay más
    if len(tallas) > 4:
        # Contar frecuencia
        from collections import Counter
        talla_freq = Counter()
        for r in records:
            s = r.get('SIZE', 0)
            if s in tallas:
                talla_freq[s] += 1
        tallas = [t for t, _ in talla_freq.most_common(4)]
        tallas.sort()

    return linea, tallas
