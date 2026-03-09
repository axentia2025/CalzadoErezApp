"""
Algoritmo de distribución v3 - Calzado Erez.
Distribución pareja: reparte a TODAS las tiendas con cantidades equilibradas.
Solo salta tienda+talla si tiene inventario y cero ventas en 15 días.
"""
from collections import defaultdict


def separate_warehouse(records):
    """Separa registros de Almacén (T13) y Tiendas.

    Returns:
        almacen_por_producto: {(marca,modelo,color): {tallas:{}, sublinea, precio, total_inv}}
        tiendas_info: {tienda: {prod_key: {talla: {inven, vtas_15}}}}
        tiendas_ordenadas: lista ordenada de IDs de tienda
    """
    almacen_records = [r for r in records if r['STR'] == 13]
    tienda_records = [r for r in records if r['STR'] != 13]

    almacen_por_producto = {}
    for r in almacen_records:
        key = (r['MARCA'], r['DESC1'], r['ATTR'])
        if key not in almacen_por_producto:
            almacen_por_producto[key] = {
                'tallas': {}, 'sublinea': '', 'precio': 0, 'total_inv': 0
            }
        almacen_por_producto[key]['tallas'][r['SIZE']] = r['INVEN']
        almacen_por_producto[key]['sublinea'] = r['SUBLINEA']
        almacen_por_producto[key]['precio'] = r['PRECIO']
        almacen_por_producto[key]['total_inv'] += r['INVEN']

    tiendas_info = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: {'inven': 0, 'vtas_15': 0})))
    todas_tiendas = set()

    for r in tienda_records:
        t = r['STR']
        todas_tiendas.add(t)
        key = (r['MARCA'], r['DESC1'], r['ATTR'])
        tiendas_info[t][key][r['SIZE']]['inven'] = r['INVEN']
        tiendas_info[t][key][r['SIZE']]['vtas_15'] = r['VTAS_15']

    tiendas_ordenadas = sorted(todas_tiendas)

    return almacen_por_producto, tiendas_info, tiendas_ordenadas


def calculate_needs(almacen_por_producto, tiendas_info, tiendas_ordenadas, tallas):
    """Calcula necesidad por producto-tienda.

    Returns:
        necesidades: {prod_key: [lista de {tienda, score, prioridad, ...}]}
        tienda_vtas_totales: {tienda: ventas_totales}
    """
    tienda_vtas_totales = defaultdict(int)
    for tienda in tiendas_ordenadas:
        for prod_key in almacen_por_producto:
            tienda_data = tiendas_info[tienda].get(prod_key, {})
            for talla in tallas:
                td = tienda_data.get(talla, {'inven': 0, 'vtas_15': 0})
                tienda_vtas_totales[tienda] += td['vtas_15']

    total_vtas_todas = sum(tienda_vtas_totales.values())

    necesidades = {}
    for prod_key, prod_info in almacen_por_producto.items():
        if prod_info['total_inv'] == 0:
            continue

        prod_necesidades = []
        prod_vtas_total = 0

        for tienda in tiendas_ordenadas:
            tienda_data = tiendas_info[tienda].get(prod_key, {})
            vtas = sum(tienda_data.get(t, {'vtas_15': 0})['vtas_15'] for t in tallas)
            prod_vtas_total += vtas

        for tienda in tiendas_ordenadas:
            tienda_data = tiendas_info[tienda].get(prod_key, {})

            total_vtas_15 = 0
            total_inven = 0
            tallas_detalle = {}

            for talla in tallas:
                td = tienda_data.get(talla, {'inven': 0, 'vtas_15': 0})
                total_vtas_15 += td['vtas_15']
                total_inven += td['inven']
                tallas_detalle[talla] = td

            vtas_diaria = total_vtas_15 / 15.0

            # Score: 60% demanda directa + 30% peso tienda + 10% base × urgencia
            if prod_vtas_total > 0:
                score_demanda = total_vtas_15 / prod_vtas_total
            else:
                score_demanda = 0

            if total_vtas_todas > 0:
                score_tienda = tienda_vtas_totales[tienda] / total_vtas_todas
            else:
                score_tienda = 1.0 / len(tiendas_ordenadas)

            if vtas_diaria > 0:
                dias_inv = total_inven / vtas_diaria
                if dias_inv < 5:
                    factor_urgencia = 2.0
                elif dias_inv < 10:
                    factor_urgencia = 1.5
                elif dias_inv < 20:
                    factor_urgencia = 1.0
                else:
                    factor_urgencia = 0.5
            else:
                factor_urgencia = 0.3 if total_inven == 0 else 0.1

            score = (score_demanda * 0.6 + score_tienda * 0.3 + 0.01) * factor_urgencia

            if vtas_diaria > 0:
                dias_inv = total_inven / vtas_diaria
                if dias_inv < 5:
                    prioridad = "URGENTE"
                elif dias_inv < 10:
                    prioridad = "ALTA"
                elif dias_inv < 20:
                    prioridad = "NORMAL"
                else:
                    prioridad = "BAJA"
            else:
                prioridad = "OPORTUNIDAD" if total_inven == 0 else "BAJA"

            prod_necesidades.append({
                'tienda': tienda,
                'score': score,
                'prioridad': prioridad,
                'vtas_diaria': vtas_diaria,
                'inven_actual': total_inven,
                'dias_inv': total_inven / vtas_diaria if vtas_diaria > 0 else 999,
                'tallas_detalle': tallas_detalle,
            })

        prod_necesidades.sort(key=lambda x: -x['score'])
        necesidades[prod_key] = prod_necesidades

    return necesidades, tienda_vtas_totales


def distribute_stock(almacen_por_producto, necesidades, tallas, tiendas_info,
                     tiendas_ordenadas, rules=None, product_order=None):
    """Distribuye stock del almacén a tiendas con distribución pareja.

    Estrategia:
    - Para cada talla, reparte el stock de forma equitativa entre TODAS las tiendas elegibles.
    - Solo salta tienda+talla si la tienda tiene inventario > 0 Y cero ventas en 15 días.
    - Si repartir solo entre elegibles daría más de SPREAD_THRESHOLD por tienda,
      incluye también las tiendas saltadas para distribuir más ampliamente.
    - Las tiendas con mayor demanda reciben primero los pares de residuo (+1).
    - Objetivo: distribuir TODO el inventario.

    Args:
        rules: lista de reglas (opcional). Aplica max_per_size y max_total_store.
        product_order: lista ordenada de prod_keys (opcional). Controla orden de procesamiento.

    Returns:
        distribuciones: lista de dicts con cada asignación talla-tienda
        resumen_producto: {prod_key: {inv_original, distribuido, restante}}
    """
    SPREAD_THRESHOLD = 5  # Si por tienda superaría esto, expandir a más tiendas
    if rules is None:
        rules = []

    stock_disponible = {}
    for prod_key, prod_info in almacen_por_producto.items():
        stock_disponible[prod_key] = dict(prod_info['tallas'])

    distribuciones = []
    resumen_producto = {}
    store_running_totals = defaultdict(int)  # Para max_total_store

    # Usar orden personalizado si se proporcionó
    iteration_keys = product_order if product_order else list(necesidades.keys())

    for prod_key in iteration_keys:
        if prod_key not in necesidades:
            continue
        tiendas_necesidad = necesidades[prod_key]
        prod_info = almacen_por_producto[prod_key]
        marca, modelo, color = prod_key

        if prod_info['total_inv'] == 0:
            continue

        stock = stock_disponible[prod_key]
        if sum(stock.values()) == 0:
            continue

        distribuido_producto = 0

        for talla in tallas:
            stock_talla = stock.get(talla, 0)
            if stock_talla <= 0:
                continue

            # Clasificar tiendas en preferidas (elegibles) y saltables
            preferred = []   # Sin stock, o tiene stock con ventas
            skippable = []   # Tiene stock Y cero ventas en 15 días

            for nec in tiendas_necesidad:
                td = nec['tallas_detalle'].get(talla, {'inven': 0, 'vtas_15': 0})
                if td['inven'] > 0 and td['vtas_15'] == 0:
                    skippable.append(nec)
                else:
                    preferred.append(nec)

            # Ordenar por score (mayor demanda primero para el residuo)
            preferred.sort(key=lambda x: -x['score'])
            skippable.sort(key=lambda x: -x['score'])

            # Pool de distribución: empezar con preferidas
            pool = preferred

            # Si distribuir solo a preferidas daría demasiado por tienda,
            # expandir incluyendo las saltables
            if pool and stock_talla / len(pool) > SPREAD_THRESHOLD:
                pool = preferred + skippable
                pool.sort(key=lambda x: -x['score'])

            # Si no hay preferidas, usar todas las tiendas
            if not pool:
                pool = preferred + skippable
                pool.sort(key=lambda x: -x['score'])

            if not pool:
                continue

            # Distribución equitativa
            n = len(pool)
            base = stock_talla // n
            remainder = stock_talla % n

            # Si base es 0, solo 'remainder' tiendas reciben 1 par cada una
            # Si base > 0, todas reciben 'base' y las primeras 'remainder' reciben base+1
            assigned_total = 0
            for i, nec in enumerate(pool):
                qty = base + (1 if i < remainder else 0)
                if qty <= 0:
                    continue

                # No exceder el stock restante
                qty = min(qty, stock_talla - assigned_total)
                if qty <= 0:
                    break

                # Aplicar reglas de distribución (max_per_size, max_total_store)
                if rules:
                    from .rules import apply_distribution_rules
                    qty = apply_distribution_rules(
                        rules, nec['tienda'], talla, qty, prod_key,
                        almacen_por_producto, store_running_totals
                    )
                    if qty <= 0:
                        continue

                td = nec['tallas_detalle'].get(talla, {'inven': 0, 'vtas_15': 0})

                assigned_total += qty
                store_running_totals[nec['tienda']] += qty
                distribuciones.append({
                    'TIENDA': nec['tienda'],
                    'SUBLINEA': prod_info['sublinea'],
                    'MARCA': marca,
                    'MODELO': modelo,
                    'COLOR': color,
                    'TALLA': talla,
                    'QTY_ENVIAR': qty,
                    'PRECIO': prod_info['precio'],
                    'PRIORIDAD': nec['prioridad'],
                    'INV_ALMACEN': prod_info['tallas'].get(talla, 0),
                    'INV_TIENDA': td['inven'],
                    'VTAS_15D': td['vtas_15'],
                })

            stock[talla] -= assigned_total
            distribuido_producto += assigned_total

        resumen_producto[prod_key] = {
            'inv_original': prod_info['total_inv'],
            'distribuido': distribuido_producto,
            'restante': sum(stock.values()),
        }

    distribuciones.sort(key=lambda x: (
        x['TIENDA'], x['SUBLINEA'], x['MARCA'], x['MODELO'], x['COLOR'], x['TALLA']
    ))

    return distribuciones, resumen_producto


def build_summary(distribuciones, resumen_producto, almacen_por_producto):
    """Construye resumen por tienda, sublínea y prioridad.

    Returns:
        dict con: resumen_tienda, resumen_sublinea, resumen_prioridad,
                  total_pares, total_productos, total_valor, total_almacen,
                  pct_distribuido, total_restante, productos_completos
    """
    resumen_tienda = defaultdict(lambda: {
        'pares': 0, 'productos': set(), 'urgente': 0, 'alta': 0,
        'normal': 0, 'baja': 0, 'oportunidad': 0, 'valor': 0
    })
    resumen_sublinea = defaultdict(int)
    resumen_prioridad = defaultdict(int)

    for d in distribuciones:
        t = d['TIENDA']
        resumen_tienda[t]['pares'] += d['QTY_ENVIAR']
        resumen_tienda[t]['productos'].add((d['MARCA'], d['MODELO'], d['COLOR']))
        resumen_tienda[t]['valor'] += d['QTY_ENVIAR'] * d['PRECIO']
        resumen_tienda[t][d['PRIORIDAD'].lower()] += d['QTY_ENVIAR']
        resumen_sublinea[d['SUBLINEA']] += d['QTY_ENVIAR']
        resumen_prioridad[d['PRIORIDAD']] += d['QTY_ENVIAR']

    total_almacen = sum(p['total_inv'] for p in almacen_por_producto.values())
    total_pares = sum(r['pares'] for r in resumen_tienda.values())
    total_productos = len(set((d['MARCA'], d['MODELO'], d['COLOR']) for d in distribuciones))
    total_valor = sum(r['valor'] for r in resumen_tienda.values())
    pct_distribuido = total_pares / total_almacen * 100 if total_almacen > 0 else 0
    total_restante = sum(info['restante'] for info in resumen_producto.values())
    productos_completos = sum(1 for info in resumen_producto.values() if info['restante'] == 0)

    return {
        'resumen_tienda': dict(resumen_tienda),
        'resumen_sublinea': dict(resumen_sublinea),
        'resumen_prioridad': dict(resumen_prioridad),
        'total_pares': total_pares,
        'total_productos': total_productos,
        'total_valor': total_valor,
        'total_almacen': total_almacen,
        'productos_almacen': len(almacen_por_producto),
        'pct_distribuido': pct_distribuido,
        'total_restante': total_restante,
        'productos_completos': productos_completos,
    }


def run_distribution(records, tallas, progress_callback=None, rules=None):
    """Ejecuta el flujo completo de distribución.

    Args:
        rules: lista de reglas personalizadas (opcional).

    Returns:
        dict con: distribuciones, resumen_producto, summary, almacen_por_producto,
                  tiendas_info, tiendas_ordenadas, rules_applied
    """
    if rules is None:
        rules = []

    if progress_callback:
        progress_callback(0.80, "Separando almacén y tiendas...")

    almacen, tiendas_info, tiendas_ord = separate_warehouse(records)

    # Fase 1: Pre-filter rules (excluir tiendas/productos/tallas)
    if rules:
        from .rules import apply_pre_filter_rules
        tiendas_ord = apply_pre_filter_rules(rules, almacen, tiendas_info, tiendas_ord)

    if progress_callback:
        progress_callback(0.85, "Calculando necesidades...")

    necesidades, tienda_vtas = calculate_needs(almacen, tiendas_info, tiendas_ord, tallas)

    # Fase 2: Scoring rules (priorizar/depriorizar)
    if rules:
        from .rules import apply_scoring_rules
        apply_scoring_rules(rules, necesidades, almacen)

    # Fase 3: Ordering rules (reordenar productos)
    product_order = None
    if rules:
        from .rules import apply_ordering_rules
        product_order = apply_ordering_rules(rules, necesidades, almacen)

    if progress_callback:
        progress_callback(0.90, "Distribuyendo stock...")

    distribuciones, resumen_prod = distribute_stock(
        almacen, necesidades, tallas, tiendas_info, tiendas_ord,
        rules=rules, product_order=product_order
    )

    # Fase 5: Post-filter (safety net)
    if rules:
        from .rules import apply_post_filter_rules
        distribuciones = apply_post_filter_rules(rules, distribuciones)

    if progress_callback:
        progress_callback(0.95, "Generando resumen...")

    summary = build_summary(distribuciones, resumen_prod, almacen)

    return {
        'distribuciones': distribuciones,
        'resumen_producto': resumen_prod,
        'summary': summary,
        'almacen_por_producto': almacen,
        'tiendas_info': tiendas_info,
        'tiendas_ordenadas': tiendas_ord,
        'rules_applied': len(rules),
    }
