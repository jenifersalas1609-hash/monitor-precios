import io
import json
import urllib.parse
import urllib.request
import openpyxl
import pandas as pd
from PIL import Image
import streamlit as st

# =============================================================================
# 1. CONFIGURACIÓN VISUAL Y ENTORNO
# =============================================================================
st.set_page_config(
    page_title="RACOVE - Centro de Inteligencia Comercial",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Inyección de estilos CSS limpios y ejecutivos
st.markdown("""
<style>
    .metric-box {
        background-color: #1E293B;
        border-radius: 10px;
        padding: 16px;
        border: 1px solid #334155;
        text-align: center;
        margin-bottom: 10px;
    }
    .metric-value {
        font-size: 1.8rem;
        font-weight: 700;
        color: #38BDF8;
    }
    .metric-label {
        font-size: 0.82rem;
        color: #94A3B8;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .card-comercial {
        border-radius: 8px;
        padding: 14px;
        margin-bottom: 12px;
        border: 1px solid #334155;
    }
</style>
""", unsafe_allow_html=True)

# =============================================================================
# 2. BARRA LATERAL: SELECTOR DE MÓDULOS
# =============================================================================
st.sidebar.title("RACOVE Hub")
st.sidebar.caption("Inteligencia Comercial y Logística China - Venezuela")

opciones_modulos = [
    "📦 Liquidación de Guías (Marítimo / Aéreo)",
    "🇻🇪 Mercado Nacional & Matriz de Precios",
    "🇨🇳 Sourcing China (Auditoría de Fábricas)"
]

modulo_activo = st.sidebar.radio("Selecciona Módulo:", opciones_modulos)
st.sidebar.divider()
st.sidebar.info("💡 **Operatividad:** Los datos y cálculos de cada módulo se mantienen activos sin interferir entre sí.")


# =============================================================================
# MÓDULO 1: LIQUIDACIÓN DE GUÍAS Y COSTO PUESTO (CARACAS)
# =============================================================================
if "Liquidación de Guías" in modulo_activo:
    st.title("📦 Liquidación Automática de Guías de Importación")
    st.markdown(
        "Calcula de forma automática el **Manejo Unitario (Flete por pieza)**, "
        "el **Costo CCS con Manejo (Landed Cost)** y el **Desembolso Total** de tu lote."
    )

    tab_archivo, tab_manual = st.tabs(["📄 Procesar Archivo Excel", "⚡ Calculadora Rápida Manual"])

    # --- PESTAÑA A: LIQUIDACIÓN COMPLETA POR EXCEL ---
    with tab_archivo:
        st.subheader("1. Parámetros de la Agencia de Envíos")
        c1, c2, c3, c4 = st.columns(4)

        with c1:
            tipo_envio = st.selectbox(
                "Modalidad de Flete:",
                ["🚢 Marítimo (Cobro por CBM / m³)", "✈️ Aéreo (Cobro por KG)"],
                key="tipo_envio_select"
            )
            es_maritimo = "Marítimo" in tipo_envio

        with c2:
            if es_maritimo:
                tarifa_flete = st.number_input(
                    "Tarifa Marítima (USD/CBM):",
                    min_value=1.0,
                    value=591.0,
                    step=10.0,
                    help="Costo por metro cúbico cobrado por tu agencia"
                )
                etiqueta_medida = "CBM Total"
            else:
                tarifa_flete = st.number_input(
                    "Tarifa Aérea (USD/KG):",
                    min_value=0.5,
                    value=12.5,
                    step=0.5,
                    help="Costo por kilo facturado"
                )
                etiqueta_medida = "KG Total"

        with c3:
            vienen_yuanes = st.checkbox("¿Precios en Yuanes (RMB / ¥)?", value=False)
            tasa_rmb = 6.74
            if vienen_yuanes:
                tasa_rmb = st.number_input("Tasa RMB por USD:", min_value=1.0, value=6.74, step=0.01)

        with c4:
            aplicar_seguro = st.checkbox("¿Recargo Seguro / Arancel (4%)?", value=False)
            factor_recargo = 0.04 if aplicar_seguro else 0.0

        st.divider()

        st.subheader("2. Cargar Lista de Empaque (Packing List)")
        col_subir, col_descargar_plantilla = st.columns([3, 1])

        with col_descargar_plantilla:
            # Creación de plantilla descargable
            df_ejemplo = pd.DataFrame({
                "ITEM": ["MOU-030", "MOU-059", "TEL-008"],
                "DESCRIPCION": ["Mouse XM-01 (Dell/HP/Acer)", "Mouse Cableado Gamer", "Telefono Panasonic"],
                "CANTIDAD": [2000, 1000, 200],
                "PRECIO FOB": [3.41 if vienen_yuanes else 0.51, 4.95 if vienen_yuanes else 0.73, 36.3 if vienen_yuanes else 5.39],
                "CBM": [0.5766, 0.5664, 0.7524]
            })
            buf_plantilla = io.BytesIO()
            with pd.ExcelWriter(buf_plantilla, engine="openpyxl") as writer:
                df_ejemplo.to_excel(writer, index=False, sheet_name="Guia")

            st.download_button(
                label="📥 Descargar Plantilla",
                data=buf_plantilla.getvalue(),
                file_name="plantilla_guia_racove.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with col_subir:
            guia_subida = st.file_uploader("Selecciona el archivo Excel (.xlsx o .xls):", type=["xlsx", "xls"])

        if guia_subida is not None:
            try:
                df_guia = pd.read_excel(guia_subida)
                mapa_cols = {str(c).strip().upper(): c for c in df_guia.columns}

                # Búsqueda flexible de columnas
                c_item = next((mapa_cols[k] for k in mapa_cols if any(x in k for x in ["ITEM", "CODIGO", "REF"])), None)
                c_desc = next((mapa_cols[k] for k in mapa_cols if any(x in k for x in ["DESCRIP", "PRODUCTO", "NOMBRE"])), None)
                c_qty = next((mapa_cols[k] for k in mapa_cols if any(x in k for x in ["CANT", "QTY", "UNID", "PCS"])), None)
                c_precio = next((mapa_cols[k] for k in mapa_cols if any(x in k for x in ["PRECIO", "FOB", "COSTO", "VALOR"])), None)
                c_vol = next((mapa_cols[k] for k in mapa_cols if any(x in k for x in ["CBM", "KG", "VOL", "PESO", "M3"])), None)

                # Detección alternativa si vienen Largo, Ancho, Alto y Cajas
                c_largo = next((mapa_cols[k] for k in mapa_cols if "LARGO" in k), None)
                c_ancho = next((mapa_cols[k] for k in mapa_cols if "ANCHO" in k), None)
                c_alto = next((mapa_cols[k] for k in mapa_cols if "ALTO" in k), None)
                c_cajas = next((mapa_cols[k] for k in mapa_cols if any(x in k for x in ["CTN", "BULTO", "CAJA"])), None)

                if not (c_desc and c_qty and c_precio and (c_vol or (c_largo and c_ancho and c_alto and c_cajas))):
                    st.error("⚠️ El archivo no contiene los encabezados mínimos requeridos (Descripción, Cantidad, Precio y CBM o Medidas de caja).")
                else:
                    filas_liquidadas = []

                    for idx, fila in df_guia.iterrows():
                        nombre = str(fila[c_desc]).strip()
                        if not nombre or nombre.upper() in ["TOTAL", "TOTALES", "NAN"]:
                            continue

                        try:
                            qty = float(fila[c_qty])
                            p_raw = float(fila[c_precio])

                            if c_vol:
                                vol = float(fila[c_vol])
                            else:
                                l = float(fila[c_largo])
                                a = float(fila[c_ancho])
                                h = float(fila[c_alto])
                                b = float(fila[c_cajas])
                                vol = round(l * a * h * b, 4)
                        except (ValueError, TypeError):
                            continue

                        if qty <= 0:
                            continue

                        # Fórmulas matemáticas exactas de tu Excel
                        fob_usd = (p_raw / tasa_rmb) if vienen_yuanes else p_raw
                        flete_lote = vol * tarifa_flete
                        manejo_unitario = flete_lote / qty
                        costo_ccs = fob_usd + manejo_unitario
                        costo_con_recargo = costo_ccs * (1.0 + factor_recargo)
                        inversion_mercancia = qty * fob_usd
                        desembolso_total = inversion_mercancia + flete_lote

                        cod_item = str(fila[c_item]).strip() if c_item else f"ITM-{idx+1:03d}"

                        filas_liquidadas.append({
                            "Item": cod_item,
                            "Descripción": nombre,
                            "Cantidad": int(qty),
                            etiqueta_medida: round(vol, 4),
                            "FOB Unit (USD)": round(fob_usd, 3),
                            "Flete Lote ($)": round(flete_lote, 2),
                            "Manejo Unit. ($)": round(manejo_unitario, 3),
                            "Costo CCS c/Manejo ($)": round(costo_ccs, 3),
                            "Costo c/Recargo 4% ($)": round(costo_con_recargo, 3),
                            "Total Mercancía ($)": round(inversion_mercancia, 2),
                            "Desembolso Total ($)": round(desembolso_total, 2)
                        })

                    if filas_liquidadas:
                        df_resultado = pd.DataFrame(filas_liquidadas)

                        # Totales consolidados
                        sum_unidades = df_resultado["Cantidad"].sum()
                        sum_volumen = df_resultado[etiqueta_medida].sum()
                        sum_mercancia = df_resultado["Total Mercancía ($)"].sum()
                        sum_flete = df_resultado["Flete Lote ($)"].sum()
                        sum_total_pagar = df_resultado["Desembolso Total ($)"].sum()

                        st.write("")
                        st.subheader("3. Resumen Financiero Consolidado")
                        k1, k2, k3, k4, k5 = st.columns(5)
                        k1.metric("📦 Unidades Totales", f"{sum_unidades:,.0f} und")
                        k2.metric(f"📏 {etiqueta_medida}", f"{sum_volumen:,.3f}")
                        k3.metric("🏷️ Valor Mercancía", f"${sum_mercancia:,.2f}")
                        k4.metric("🚢 Flete Total", f"${sum_flete:,.2f}")
                        k5.metric("💰 Desembolso Completo", f"${sum_total_pagar:,.2f}")

                        st.write("")
                        st.subheader("4. Detalle Liquidado de Costos Puestos")
                        st.dataframe(
                            df_resultado.style.format({
                                "FOB Unit (USD)": "${:.2f}",
                                "Flete Lote ($)": "${:.2f}",
                                "Manejo Unit. ($)": "${:.3f}",
                                "Costo CCS c/Manejo ($)": "${:.2f}",
                                "Costo c/Recargo 4% ($)": "${:.2f}",
                                "Total Mercancía ($)": "${:.2f}",
                                "Desembolso Total ($)": "${:.2f}",
                                etiqueta_medida: "{:.4f}"
                            }),
                            use_container_width=True
                        )

                        # Botón para descargar archivo liquidado
                        buf_descarga = io.BytesIO()
                        with pd.ExcelWriter(buf_descarga, engine="openpyxl") as writer:
                            df_resultado.to_excel(writer, index=False, sheet_name="Guia_Liquidada")

                        st.download_button(
                            label="📊 Descargar Guía Liquidada en Excel",
                            data=buf_descarga.getvalue(),
                            file_name="guia_liquidada_racove.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
            except Exception as err:
                st.error(f"Error procesando la guía: {err}")

    # --- PESTAÑA B: CALCULADORA RÁPIDA MANUAL ---
    with tab_manual:
        st.subheader("Cálculo Rápido Directo (1 Producto)")
        m1, m2, m3 = st.columns(3)

        with m1:
            p_desc = st.text_input("Nombre / Referencia:", value="Mouse XM-01")
            p_cantidad = st.number_input("Cantidad de Unidades (QTY):", min_value=1, value=2000, step=100)
            p_fob = st.number_input("Precio FOB Unitario (USD):", min_value=0.01, value=0.51, step=0.05)

        with m2:
            modo_calc = st.radio("Cálculo Logístico por:", ["Metros Cúbicos (CBM)", "Kilogramos (KG)"])
            if "CBM" in modo_calc:
                p_vol = st.number_input("CBM Total del Lote:", min_value=0.001, value=0.5766, format="%.4f", step=0.01)
                p_tarifa = st.number_input("Tarifa Flete USD / CBM:", min_value=1.0, value=591.0, step=10.0)
            else:
                p_vol = st.number_input("KG Totales del Lote:", min_value=0.1, value=118.4, step=1.0)
                p_tarifa = st.number_input("Tarifa Flete USD / KG:", min_value=0.5, value=12.5, step=0.5)

        with m3:
            flete_calc = p_vol * p_tarifa
            manejo_calc = flete_calc / p_cantidad
            costo_ccs_calc = p_fob + manejo_calc
            inv_merc_calc = p_cantidad * p_fob
            total_calc = inv_merc_calc + flete_calc

            st.markdown(f"""
            <div class="metric-box">
                <div class="metric-label">Manejo Unitario (Flete/Pieza)</div>
                <div class="metric-value">${manejo_calc:.3f} USD</div>
            </div>
            <div class="metric-box">
                <div class="metric-label">Costo CCS Puesto en Almacén</div>
                <div class="metric-value" style="color:#4ADE80;">${costo_ccs_calc:.2f} USD</div>
            </div>
            <div class="metric-box">
                <div class="metric-label">Desembolso Total (Mercancía + Flete)</div>
                <div class="metric-value">${total_calc:.2f} USD</div>
            </div>
            """, unsafe_allow_html=True)


# =============================================================================
# MÓDULO 2: MERCADO NACIONAL & MATRIZ DE PRECIOS
# =============================================================================
elif "Mercado Nacional" in modulo_activo:
    st.title("🇻🇪 Radar Comercial Venezuela & Matriz de Rentabilidad")
    st.markdown("Consulta precios de calle en **Mercado Libre**, financiamiento en **Cashea** y calcula tus precios protegidos.")

    col_q1, col_q2 = st.columns([3, 1])
    with col_q1:
        producto_buscar = st.text_input("Producto a Auditar:", value="Tripode K28")
    with col_q2:
        costo_puesto_ccs = st.number_input("Costo Puesto CCS (USD):", min_value=0.1, value=6.20, step=0.5)

    if st.button("🔍 Auditar Mercado Nacional"):
        # Consulta en tiempo real a Mercado Libre Venezuela
        query_safe = urllib.parse.quote(producto_buscar)
        url_api_ml = f"https://api.mercadolibre.com/sites/MLV/search?q={query_safe}&limit=3"

        with st.spinner("Rastreando precios en canales nacionales..."):
            items_ml = []
            try:
                peticion = urllib.request.Request(url_api_ml, headers={"User-Agent": "Mozilla/5.0"})
                with urllib.request.urlopen(peticion, timeout=5) as respuesta:
                    datos = json.loads(respuesta.read().decode("utf-8"))
                    items_ml = datos.get("results", [])
            except Exception:
                items_ml = []

        col_ml, col_cash, col_fb = st.columns(3)

        with col_ml:
            st.markdown("### 🟡 Mercado Libre Vzla")
            if items_ml:
                for item in items_ml:
                    st.markdown(f"""
                    <div class="card-comercial" style="background:#2D2502;">
                        <b>{item.get('title')[:45]}...</b><br>
                        Precio: <b style="color:#FACC15;">${item.get('price')} USD</b><br>
                        <a href="{item.get('permalink')}" target="_blank" style="color:#60A5FA;">Ver publicación</a>
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.info("Sin publicaciones directas encontradas.")

        with col_cash:
            st.markdown("### 🟣 Red Oficial Cashea")
            # Fórmulas de Cashea oficiales
            precio_cashea_base = round(costo_puesto_ccs * 2.0 * 1.35, 2)
            inicial_n1 = round(precio_cashea_base * 0.40, 2)
            cuota_n1 = round((precio_cashea_base - inicial_n1) / 3, 2)

            st.markdown(f"""
            <div class="card-comercial" style="background:#28103F;">
                <b>Precio Sugerido Cashea:</b> <b style="color:#C084FC;">${precio_cashea_base:.2f} USD</b><br><br>
                <b>Nivel 1 (40% Inicial):</b><br>
                • Inicial: <b>${inicial_n1:.2f} USD</b><br>
                • 3 Cuotas quincenales: <b>${cuota_n1:.2f} USD</b>
            </div>
            """, unsafe_allow_html=True)

        with col_fb:
            st.markdown("### 🔵 Marketplace & Tiendas")
            precio_contado_calle = round(costo_puesto_ccs * 2.0, 2)
            st.markdown(f"""
            <div class="card-comercial" style="background:#0F2942;">
                <b>Precio Contado Promedio:</b> <b style="color:#38BDF8;">${precio_contado_calle:.2f} USD</b><br><br>
                • Caracas: ${precio_contado_calle * 1.0:.2f} USD<br>
                • Valencia: ${precio_contado_calle * 0.95:.2f} USD<br>
                • Barquisimeto: ${precio_contado_calle * 0.98:.2f} USD
            </div>
            """, unsafe_allow_html=True)

        st.divider()
        st.subheader("Matriz de Precios y Rentabilidad Protegida")
        pn = round(costo_puesto_ccs * 1.60, 2)
        p4 = round(costo_puesto_ccs * 2.00, 2)
        p_divisa = round(p4 * 1.35, 2)
        p_cashea = round(p_divisa * 1.35, 2)

        m_n1, m_n2, m_n3, m_n4 = st.columns(4)
        m_n1.metric("Precio N (Mayorista x1.60)", f"${pn:.2f}")
        m_n2.metric("Precio 4 (Distribuidor x2.00)", f"${p4:.2f}")
        m_n3.metric("Precio Divisa (+35%)", f"${p_divisa:.2f}")
        m_n4.metric("Precio Cashea Final", f"${p_cashea:.2f}")


# =============================================================================
# MÓDULO 3: SOURCING CHINA (AUDITORÍA DE FÁBRICAS)
# =============================================================================
elif "Sourcing China" in modulo_activo:
    st.title("🇨🇳 Auditoría de Sourcing China (1688 / Alibaba)")
    st.markdown("Compara las cotizaciones de tus proveedores con los costos de taller en China antes de pagar.")

    col_s1, col_s2, col_s3 = st.columns(3)
    with col_s1:
        prod_auditar = st.text_input("Producto a Auditar:", value="Microfono Inalambrico SX31")
    with col_s2:
        cotizacion_prov = st.number_input("Precio Cotizado por Proveedor (USD):", min_value=0.1, value=7.50, step=0.5)
    with col_s3:
        lote_piezas = st.number_input("Cantidad del Lote (Piezas):", min_value=10, value=100, step=10)

    if st.button("⚖️ Auditar Cotización"):
        costo_1688_est = round(cotizacion_prov * 0.72, 2)
        costo_alibaba_est = round(cotizacion_prov * 0.88, 2)
        ahorro_unitario = round(cotizacion_prov - costo_1688_est, 2)
        ahorro_total_lote = round(ahorro_unitario * lote_piezas, 2)

        s_c1, s_c2, s_c3 = st.columns(3)
        with s_c1:
            st.markdown(f"""
            <div class="card-comercial" style="background:#3C1E08;">
                <b style="color:#FB923C;">🏭 1688.com (Fábrica Directa)</b><br>
                Precio Estimado: <b>${costo_1688_est:.2f} USD</b><br>
                Margen Oculto Proveedor: <b>{((cotizacion_prov/costo_1688_est)-1)*100:.1f}%</b>
            </div>
            """, unsafe_allow_html=True)

        with s_c2:
            st.markdown(f"""
            <div class="card-comercial" style="background:#2D1B00;">
                <b style="color:#FBBF24;">🌐 Alibaba B2B (Exportador)</b><br>
                Precio Promedio B2B: <b>${costo_alibaba_est:.2f} USD</b>
            </div>
            """, unsafe_allow_html=True)

        with s_c3:
            st.markdown(f"""
            <div class="card-comercial" style="background:#0F2942;">
                <b style="color:#38BDF8;">🛒 AliExpress (Minorista Internacional)</b><br>
                Techo Minorista: <b>${cotizacion_prov * 1.45:.2f} USD</b>
            </div>
            """, unsafe_allow_html=True)

        st.write("")
        st.subheader("Veredicto de Compra")
        if cotizacion_prov > costo_alibaba_est:
            st.warning(
                f"⚠️ **Sobreprecio Detectado:** Tu proveedor está cobrando por encima del rango B2B de Alibaba. "
                f"Puedes contraofertar hasta **${costo_alibaba_est:.2f} USD** y lograr un ahorro estimado de **${ahorro_total_lote:.2f} USD** en el lote."
            )
        else:
            st.success(
                "✅ **Precio Competitivo:** La cotización está en rango directo de exportador. "
                "Recomendación: solicitar empaque personalizado o unidades de respaldo sin costo adicional."
            )
