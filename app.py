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
    .success-box {
        background: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 8px;
        padding: 1rem;
        margin: 1rem 0;
        color: #155724;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">Sistema de Distribución - Calzado Erez</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Genera propuestas de distribución desde el inventario de almacén</div>', unsafe_allow_html=True)

# --- Inicializar session_state ---
if 'processed_results' not in st.session_state:
    st.session_state.processed_results = None
if 'processed_files' not in st.session_state:
    st.session_state.processed_files = None

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

    # Detectar si los archivos cambiaron para limpiar resultados previos
    current_file_names = sorted([f.name for f in uploaded_files]) if uploaded_files else []
    if current_file_names != st.session_state.processed_files:
        st.session_state.processed_results = None

    process_btn = st.button(
        "Procesar Distribución",
        type="primary",
        disabled=not uploaded_files,
        use_container_width=True,
    )

# --- SIN ARCHIVOS ---
if not uploaded_files:
    st.session_state.processed_results = None
    st.session_state.processed_files = None
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

# --- PROCESAMIENTO (solo cuando se presiona el botón) ---
if process_btn:
    all_results = []

    for file_idx, uploaded_file in enumerate(uploaded_files):
        file_label = f"Archivo {file_idx + 1}: {uploaded_file.name}"

        with st.spinner(f"Procesando {uploaded_file.name}..."):
            st.subheader(file_label)
            progress_bar = st.progress(0, text="Iniciando carga del inventario...")

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

            progress_bar.progress(1.0, text="¡Distribución completa!")

            # Paso 3: Generar archivos
            progress_bar.progress(1.0, text="Generando archivos de descarga...")
            excel_bytes = generate_excel(result, linea, fecha_str, tallas)
            word_bytes = generate_word(result, linea, fecha_str, tallas)

            nombre_excel = f"DISTRIBUCION_{linea}_{fecha_dist.strftime('%d%b%Y').upper()}.xlsx"
            nombre_word = f"OBSERVACIONES_{linea}_{fecha_dist.strftime('%d%b%Y').upper()}.docx"

            progress_bar.progress(1.0, text="¡Proceso terminado!")

            all_results.append({
                'linea': linea,
                'tallas': tallas,
                'n_records': n_records,
                'result': result,
                'summary': result['summary'],
                'excel_bytes': excel_bytes,
                'word_bytes': word_bytes,
                'nombre_excel': nombre_excel,
                'nombre_word': nombre_word,
                'file_name': uploaded_file.name,
            })

    # Guardar en session_state para persistir entre reruns
    if all_results:
        st.session_state.processed_results = all_results
        st.session_state.processed_files = current_file_names
        st.rerun()

# --- MOSTRAR RESULTADOS (desde session_state) ---
if st.session_state.processed_results:
    all_results = st.session_state.processed_results

    st.markdown('<div class="success-box">&#9989; <strong>Distribución procesada exitosamente.</strong> Los archivos están listos para descarga.</div>', unsafe_allow_html=True)

    for file_idx, res in enumerate(all_results):
        linea = res['linea']
        tallas = res['tallas']
        result = res['result']
        summary = res['summary']
        resumen_tienda = summary['resumen_tienda']

        st.subheader(f"{linea} — {res['file_name']}")

        # Detección
        col1, col2, col3 = st.columns(3)
        col1.metric("Línea detectada", linea)
        col2.metric("Registros", f"{res['n_records']:,}")
        col3.metric("Tallas", " - ".join(str(t) for t in tallas))

        # --- MÉTRICAS PRINCIPALES ---
        st.divider()
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Pares a distribuir", f"{summary['total_pares']:,}")
        m2.metric("Tiendas", len(resumen_tienda))
        m3.metric("% Almacén", f"{summary['pct_distribuido']:.0f}%")
        m4.metric("Valor", f"${summary['total_valor']:,.0f}")

        # --- DESCARGAS (arriba, bien visibles) ---
        st.divider()
        st.subheader("Descargar Resultados")
        dl1, dl2 = st.columns(2)

        with dl1:
            st.download_button(
                label=f"📊 Descargar Excel - {linea}",
                data=res['excel_bytes'],
                file_name=res['nombre_excel'],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"dl_excel_{file_idx}",
            )

        with dl2:
            st.download_button(
                label=f"📝 Descargar Word - {linea}",
                data=res['word_bytes'],
                file_name=res['nombre_word'],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                key=f"dl_word_{file_idx}",
            )

        # --- TABS: Detalle ---
        tab_resumen, tab_prioridad, tab_tiendas = st.tabs(
            ["Resumen por Tienda", "Por Prioridad", "Detalle Tiendas"]
        )

        with tab_resumen:
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

        st.divider()

    # Resumen final si hay múltiples archivos
    if len(all_results) > 1:
        st.subheader("Resumen General")
        total_general = sum(r['summary']['total_pares'] for r in all_results)
        valor_general = sum(r['summary']['total_valor'] for r in all_results)
        g1, g2 = st.columns(2)
        g1.metric("Total Pares (ambas líneas)", f"{total_general:,}")
        g2.metric("Valor Total", f"${valor_general:,.0f}")

elif uploaded_files and not process_btn:
    # Archivos cargados pero aún no se ha procesado
    st.subheader("Archivos cargados")
    for f in uploaded_files:
        st.write(f"- **{f.name}** ({f.size / 1024:.0f} KB)")
    st.info("Haz clic en **Procesar Distribución** para generar la propuesta.")
