"""
Sistema de Distribución - Calzado Erez
Interfaz Streamlit para generar distribuciones diarias desde inventario de almacén.
"""
import streamlit as st
from datetime import date, datetime
from engine.loader import load_inventory
from engine.distributor import run_distribution
from engine.excel_writer import generate_excel
from engine.word_writer import generate_word

st.set_page_config(
    page_title="Distribución - Calzado Erez",
    page_icon="👞",
    layout="wide",
)

# CSS personalizado
st.markdown("""
<style>
    .main-header {
        font-size: 2rem;
        font-weight: 700;
        color: #1F4E79;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1.1rem;
        color: #666;
        margin-bottom: 2rem;
    }
    .metric-card {
        background: #f8f9fa;
        border-radius: 8px;
        padding: 1rem;
        border-left: 4px solid #1F4E79;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 8px 16px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">Sistema de Distribución - Calzado Erez</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Genera propuestas de distribución desde el inventario de almacén</div>', unsafe_allow_html=True)

# --- SIDEBAR: Upload + Config ---
with st.sidebar:
    st.header("Configuración")

    fecha_dist = st.date_input("Fecha de distribución", value=date.today())
    fecha_str = fecha_dist.strftime("%d de %B %Y").replace(
        "January", "Enero").replace("February", "Febrero").replace(
        "March", "Marzo").replace("April", "Abril").replace(
        "May", "Mayo").replace("June", "Junio").replace(
        "July", "Julio").replace("August", "Agosto").replace(
        "September", "Septiembre").replace("October", "Octubre").replace(
        "November", "Noviembre").replace("December", "Diciembre")

    st.divider()
    st.header("Archivos de Inventario")
    st.caption("Sube 1 o 2 archivos Excel (.xlsx) con el inventario diario del almacén.")

    uploaded_files = st.file_uploader(
        "Selecciona archivos de inventario",
        type=['xlsx'],
        accept_multiple_files=True,
        help="Formato esperado: archivo de inventario con columnas STR, MARCA, SIZE, INVEN, etc."
    )

    if uploaded_files and len(uploaded_files) > 2:
        st.error("Máximo 2 archivos (DAMA y CABALLERO)")
        uploaded_files = uploaded_files[:2]

    process_btn = st.button(
        "Procesar Distribución",
        type="primary",
        disabled=not uploaded_files,
        use_container_width=True,
    )

# --- PROCESAMIENTO ---
if not uploaded_files:
    st.info("Sube un archivo de inventario en el panel izquierdo para comenzar.")
    st.markdown("""
    ### Instrucciones
    1. Selecciona la **fecha de distribución** en el panel izquierdo
    2. Sube **1 o 2 archivos** de inventario (.xlsx)
       - El sistema detecta automáticamente si es **DAMA** o **CABALLERO**
       - Funciona con archivos con o sin headers
    3. Haz clic en **Procesar Distribución**
    4. Descarga los resultados en **Excel** (picking lists) y **Word** (observaciones)
    """)
    st.stop()

# Mostrar archivos subidos
if uploaded_files and not process_btn:
    st.subheader("Archivos cargados")
    for f in uploaded_files:
        st.write(f"- **{f.name}** ({f.size / 1024:.0f} KB)")
    st.info("Haz clic en **Procesar Distribución** para generar la propuesta.")
    st.stop()

if process_btn:
    results = []

    for file_idx, uploaded_file in enumerate(uploaded_files):
        file_label = f"Archivo {file_idx + 1}: {uploaded_file.name}"
        st.subheader(file_label)

        progress_bar = st.progress(0, text="Iniciando...")

        def update_progress(pct, msg):
            progress_bar.progress(min(pct, 1.0), text=msg)

        # Paso 1: Cargar
        try:
            data = load_inventory(uploaded_file, progress_callback=update_progress)
        except Exception as e:
            st.error(f"Error al cargar archivo: {e}")
            continue

        linea = data['linea']
        tallas = data['tallas']
        n_records = data['total_rows']

        # Mostrar detección
        col1, col2, col3 = st.columns(3)
        col1.metric("Línea detectada", linea)
        col2.metric("Registros", f"{n_records:,}")
        col3.metric("Tallas", " - ".join(str(t) for t in tallas))

        if not tallas:
            st.error("No se detectaron tallas válidas en el archivo.")
            progress_bar.empty()
            continue

        # Paso 2: Distribuir
        try:
            result = run_distribution(data['records'], tallas, progress_callback=update_progress)
        except Exception as e:
            st.error(f"Error en distribución: {e}")
            continue

        progress_bar.progress(1.0, text="Distribución completa")

        summary = result['summary']
        resumen_tienda = summary['resumen_tienda']

        # --- MÉTRICAS PRINCIPALES ---
        st.divider()
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Pares a distribuir", f"{summary['total_pares']:,}")
        m2.metric("Tiendas", len(resumen_tienda))
        m3.metric("% Almacén", f"{summary['pct_distribuido']:.0f}%")
        m4.metric("Valor", f"${summary['total_valor']:,.0f}")

        # --- TABS: Detalle ---
        tab_resumen, tab_prioridad, tab_tiendas = st.tabs(
            ["Resumen por Tienda", "Por Prioridad", "Detalle Tiendas"]
        )

        with tab_resumen:
            # Tabla resumen
            rows_data = []
            for t in sorted(resumen_tienda.keys()):
                r = resumen_tienda[t]
                rows_data.append({
                    'Tienda': f'T{t}',
                    'Pares': r['pares'],
                    'Productos': len(r['productos']),
                    'Urgente': r['urgente'],
                    'Alta': r['alta'],
                    'Normal': r['normal'],
                    'Baja': r['baja'],
                    'Valor $': f"${r['valor']:,.0f}",
                })
            st.dataframe(rows_data, use_container_width=True, hide_index=True)

        with tab_prioridad:
            prio_data = []
            for prio in ['URGENTE', 'ALTA', 'NORMAL', 'BAJA', 'OPORTUNIDAD']:
                qty = summary['resumen_prioridad'].get(prio, 0)
                if qty > 0:
                    pct_p = qty / summary['total_pares'] * 100 if summary['total_pares'] > 0 else 0
                    prio_data.append({
                        'Prioridad': prio,
                        'Pares': qty,
                        '% del Total': f"{pct_p:.1f}%",
                    })
            st.dataframe(prio_data, use_container_width=True, hide_index=True)

            st.divider()
            st.write("**Por Categoría:**")
            cat_data = []
            for sub, qty in sorted(summary['resumen_sublinea'].items(), key=lambda x: -x[1]):
                pct_c = qty / summary['total_pares'] * 100 if summary['total_pares'] > 0 else 0
                cat_data.append({
                    'Sublínea': sub,
                    'Pares': qty,
                    '% del Total': f"{pct_c:.1f}%",
                })
            st.dataframe(cat_data, use_container_width=True, hide_index=True)

        with tab_tiendas:
            tienda_sel = st.selectbox(
                "Selecciona tienda",
                sorted(resumen_tienda.keys()),
                format_func=lambda x: f"Tienda {x}",
                key=f"tienda_sel_{file_idx}",
            )
            if tienda_sel:
                items_tienda = [d for d in result['distribuciones'] if d['TIENDA'] == tienda_sel]
                if items_tienda:
                    # Agrupar por producto
                    from collections import defaultdict
                    prods = defaultdict(lambda: defaultdict(int))
                    prod_meta = {}
                    for item in items_tienda:
                        pk = f"{item['MARCA']} {item['MODELO']} {item['COLOR']}"
                        prods[pk][item['TALLA']] = item['QTY_ENVIAR']
                        prod_meta[pk] = item['PRIORIDAD']

                    detail_rows = []
                    for pk in sorted(prods.keys()):
                        row_d = {'Producto': pk}
                        for t in tallas:
                            row_d[str(t)] = prods[pk].get(t, '')
                        row_d['Total'] = sum(v for v in prods[pk].values() if isinstance(v, int))
                        row_d['Prioridad'] = prod_meta[pk]
                        detail_rows.append(row_d)

                    st.dataframe(detail_rows, use_container_width=True, hide_index=True)

        # --- DESCARGAS ---
        st.divider()
        st.subheader("Descargar Resultados")

        dl1, dl2 = st.columns(2)

        with dl1:
            excel_bytes = generate_excel(result, linea, fecha_str, tallas)
            nombre_excel = f"DISTRIBUCION_{linea}_{fecha_dist.strftime('%d%b%Y').upper()}.xlsx"
            st.download_button(
                label=f"Descargar Excel - {linea}",
                data=excel_bytes,
                file_name=nombre_excel,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"dl_excel_{file_idx}",
            )

        with dl2:
            word_bytes = generate_word(result, linea, fecha_str, tallas)
            nombre_word = f"OBSERVACIONES_{linea}_{fecha_dist.strftime('%d%b%Y').upper()}.docx"
            st.download_button(
                label=f"Descargar Word - {linea}",
                data=word_bytes,
                file_name=nombre_word,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                key=f"dl_word_{file_idx}",
            )

        results.append({'linea': linea, 'summary': summary})
        st.divider()

    # Resumen final si hay múltiples archivos
    if len(results) > 1:
        st.subheader("Resumen General")
        total_general = sum(r['summary']['total_pares'] for r in results)
        valor_general = sum(r['summary']['total_valor'] for r in results)
        g1, g2 = st.columns(2)
        g1.metric("Total Pares (ambas líneas)", f"{total_general:,}")
        g2.metric("Valor Total", f"${valor_general:,.0f}")
