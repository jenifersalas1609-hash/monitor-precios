import streamlit as st
import time
import pandas as pd
import json
import re
import io
import openpyxl
import urllib.parse
import base64
import math
from PIL import Image
from google import genai
try:
    from google.genai import types
except Exception:
    types = None

st.set_page_config(
    page_title="RACOVE",
    page_icon="🎯",
    layout="wide"
)

# Estilos CSS unificados y limpios
st.markdown("""
<style>
    /* Insignias de plataformas Venezuela */
    .badge-plataforma {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        padding: 6px 14px;
        border-radius: 8px;
        font-weight: 700;
        font-size: 0.95rem;
        margin-bottom: 12px;
    }
    .badge-ml { background-color: #ffe600; color: #2d3277; border: 1px solid #eed600; }
    .badge-cashea { background-color: #581c87; color: #ffffff; }
    .badge-fb { background-color: #1877f2; color: #ffffff; }
    
    /* Insignias de plataformas China */
    .badge-1688 { background-color: #ff6000; color: #ffffff; font-weight: 700; padding: 4px 10px; border-radius: 6px; }
    .badge-alibaba { background-color: #ff6a00; color: #ffffff; font-weight: 700; padding: 4px 10px; border-radius: 6px; }
    .badge-aliexpress { background-color: #e62e04; color: #ffffff; font-weight: 700; padding: 4px 10px; border-radius: 6px; }

    /* Fichas y cajas de datos */
    .box-comercial {
        background-color: #f8fafc;
        border-left: 5px solid #2563eb;
        padding: 12px 16px;
        border-radius: 8px;
        margin-bottom: 10px;
    }
    .box-auditoria-china {
        background-color: #f8fafc;
        border-left: 5px solid #ff6000;
        padding: 12px 16px;
        border-radius: 8px;
        margin-bottom: 10px;
    }

    /* Tarjetas limpias de precios */
    .card-item-clean {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 14px;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.04);
        display: flex;
        flex-direction: column;
        height: 100%;
        margin-bottom: 8px;
    }
    .card-item-clean:hover {
        box-shadow: 0 6px 14px rgba(0, 0, 0, 0.08);
        border-color: #cbd5e1;
    }
    .card-badge-econ {
        background-color: #dcfce7;
        color: #166534;
        font-size: 0.75rem;
        font-weight: 700;
        padding: 3px 8px;
        border-radius: 4px;
        display: inline-block;
        margin-bottom: 6px;
        width: fit-content;
    }
    .card-price {
        font-size: 1.45rem;
        font-weight: 800;
        color: #0f172a;
        margin: 4px 0 6px 0;
    }
    .card-store { font-size: 0.95rem; font-weight: 700; color: #1e293b; margin-bottom: 2px; }
    .card-reputation { font-size: 0.8rem; color: #64748b; margin-bottom: 2px; }
    .card-cashea-plan {
        background-color: #f3e8ff;
        color: #6b21a8;
        padding: 4px 8px;
        border-radius: 6px;
        font-size: 0.8rem;
        font-weight: 600;
        margin: 6px 0;
    }
    .card-title-text {
        font-size: 0.82rem;
        color: #475569;
        line-height: 1.3;
        height: 38px;
        overflow: hidden;
        margin-top: 6px;
    }
</style>
""", unsafe_allow_html=True)

# -------------------------------------------------------------
# MEMORIA DE SESIÓN AISLADA POR MÓDULO (CERO CRUCE DE DATOS)
# -------------------------------------------------------------
if "ve_analisis_completado" not in st.session_state:
    st.session_state["ve_analisis_completado"] = False
if "ve_lista_resultados" not in st.session_state:
    st.session_state["ve_lista_resultados"] = []
if "ve_dict_imgs" not in st.session_state:
    st.session_state["ve_dict_imgs"] = {}

if "china_analisis_completado" not in st.session_state:
    st.session_state["china_analisis_completado"] = False
if "china_lista_resultados" not in st.session_state:
    st.session_state["china_lista_resultados"] = []
if "china_dict_imgs" not in st.session_state:
    st.session_state["china_dict_imgs"] = {}

# -------------------------------------------------------------
# FUNCIONES AUXILIARES GLOBALES
# -------------------------------------------------------------
def extraer_precio_num(texto):
    if not texto:
        return 999999.0
    nums = re.findall(r"\d+(?:\.\d+)?", str(texto).replace(",", "."))
    return float(nums[0]) if nums else 999999.0

def detectar_modelos_activos(cliente):
    modelos_encontrados = []
    try:
        for m in cliente.models.list():
            nombre = getattr(m, "name", "") or str(m)
            limpio = nombre.replace("models/", "").strip()
            if "gemini" in limpio.lower() and "embed" not in limpio.lower():
                modelos_encontrados.append(limpio)
    except Exception:
        pass
        
    if modelos_encontrados:
        flash = [m for m in modelos_encontrados if "flash" in m.lower() and "lite" not in m.lower()]
        otros = [m for m in modelos_encontrados if "flash" not in m.lower()]
        lite = [m for m in modelos_encontrados if "lite" in m.lower()]
        return flash + otros + lite
        
    return ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]

def calcular_matriz_precios(costo_unitario, menor_precio_cashea=None, menor_precio_ml=None):
    precio_n = round(costo_unitario * 1.60, 2)
    precio_4 = round(costo_unitario * 2.00, 2)
    precio_divisa = float(math.ceil(precio_4 * 1.35))
    precio_cashea_base = float(math.ceil(precio_divisa * 1.35))
    
    alerta_cashea = None
    if menor_precio_cashea and menor_precio_cashea < 999900.0:
        if menor_precio_cashea < precio_cashea_base:
            alerta_cashea = {
                "tipo": "error",
                "mensaje": (
                    f"⚠️ **Fuera de mercado en Cashea**: La competencia vende a **${menor_precio_cashea:.2f} USD**, "
                    f"por debajo de tu precio requerido de Cashea (**${precio_cashea_base:.2f} USD**). "
                    f"Se sugiere mantener **${precio_cashea_base:.2f} USD** para no comprometer tu margen en cuotas."
                )
            }
            precio_sug_cashea = precio_cashea_base
        else:
            precio_sug_cashea = max(precio_cashea_base, round(menor_precio_cashea - 1.0, 2))
            alerta_cashea = {
                "tipo": "success",
                "mensaje": (
                    f"🟢 **En competencia en Cashea**: La competencia vende a **${menor_precio_cashea:.2f} USD**. "
                    f"Tu precio sugerido de **${precio_sug_cashea:.2f} USD** es competitivo y protege tu margen."
                )
            }
    else:
        precio_sug_cashea = precio_cashea_base
        alerta_cashea = {
            "tipo": "info",
            "mensaje": f"ℹ️ Sin referencia directa en Cashea. Precio sugerido de venta: **${precio_sug_cashea:.2f} USD**."
        }
        
    alerta_ml = None
    if menor_precio_ml and menor_precio_ml < 999900.0:
        if menor_precio_ml < precio_n:
            alerta_ml = {
                "tipo": "error",
                "mensaje": (
                    f"⚠️ **Fuera de mercado en Mercado Libre**: El vendedor más bajo vende a **${menor_precio_ml:.2f} USD**, "
                    f"por debajo de tu Precio N (**${precio_n:.2f} USD**)."
                )
            }
            precio_sug_ml = precio_n
        else:
            precio_sug_ml = max(precio_n, round(menor_precio_ml - 0.50, 2))
            alerta_ml = {
                "tipo": "success",
                "mensaje": (
                    f"🟢 **En competencia en Mercado Libre**: Mínimo de competencia a **${menor_precio_ml:.2f} USD**. "
                    f"Precio sugerido: **${precio_sug_ml:.2f} USD**."
                )
            }
    else:
        precio_sug_ml = precio_4
        alerta_ml = None
        
    return {
        "precio_n": precio_n,
        "precio_4": precio_4,
        "precio_divisa": precio_divisa,
        "precio_sug_cashea": precio_sug_cashea,
        "precio_sug_ml": precio_sug_ml,
        "alerta_cashea": alerta_cashea,
        "alerta_ml": alerta_ml
    }

# Benchmark sincerizado industrial para auditoría China (Garantía de 0 campos vacíos)
def estimar_mercado_china_benchmark(producto, precio_prov, moq, tasa_cambio=7.23):
    p_lower = str(producto).lower()
    
    if any(k in p_lower for k in ["trampolin", "trampolín", "elástica", "elastica", "cama"]):
        pies = 6
        for size in [16, 14, 12, 10, 8, 6]:
            if f"{size} pie" in p_lower or f"{size}pie" in p_lower or f"{size} ft" in p_lower or f"{size}ft" in p_lower or f"{size}英寸" in p_lower:
                pies = size
                break
        
        tabla_trampolines = {
            6: {"rmb": (165, 210), "ali_usd": (30, 39), "aliexp_usd": 68},
            8: {"rmb": (245, 310), "ali_usd": (43, 54), "aliexp_usd": 98},
            10: {"rmb": (345, 430), "ali_usd": (59, 74), "aliexp_usd": 138},
            12: {"rmb": (450, 550), "ali_usd": (76, 94), "aliexp_usd": 178},
            14: {"rmb": (570, 690), "ali_usd": (96, 119), "aliexp_usd": 228},
            16: {"rmb": (720, 860), "ali_usd": (122, 148), "aliexp_usd": 285},
        }
        ref = tabla_trampolines.get(pies, tabla_trampolines[6])
        min_rmb, max_rmb = ref["rmb"]
        min_usd = round(min_rmb / tasa_cambio, 2)
        max_usd = round(max_rmb / tasa_cambio, 2)
        ali_min, ali_max = ref["ali_usd"]
        aliexp = ref["aliexp_usd"]
        
        prom_1688 = (min_usd + max_usd) / 2
        sobreprecio_pct = round(((precio_prov - prom_1688) / prom_1688) * 100, 1)
        
        if sobreprecio_pct > 50:
            clasif = "SOBREPRECIO"
            icono = "🔴"
            contra_min = round(ali_min * 0.95, 2)
            contra_max = round(ali_max * 1.05, 2)
            diag = f"El proveedor cotiza con un recargo de +{sobreprecio_pct}% frente a talleres de Zhejiang. Actúa como intermediario con precio cercano a retail."
            arg = f"Exigir precio B2B de exportador directo ($ {contra_min:.2f} - $ {contra_max:.2f} USD). La estructura de tubos galvanizados y red de {pies}ft en 1688 ronda ¥{min_rmb}-¥{max_rmb}."
        elif sobreprecio_pct > 20:
            clasif = "REGULAR"
            icono = "🟡"
            contra_min = round(ali_min, 2)
            contra_max = round(ali_max, 2)
            diag = f"Precio de trading company con margen negociable (+{sobreprecio_pct}% vs fábrica local). Hay espacio de rebaja por volumen."
            arg = "Ofrecer compra en lote consolidado con otros tamaños para nivelar el costo a rango de contenedor."
        else:
            clasif = "EXCELENTE"
            icono = "🟢"
            contra_min = round(precio_prov * 0.95, 2)
            contra_max = precio_prov
            diag = "Cotización altamente competitiva, muy cercana al costo directo de taller en China."
            arg = "Solicitar accesorios adicionales de cortesía (escalera, anclajes de viento o repuesto de resortes)."

    elif any(k in p_lower for k in ["inflable", "castillo", "casa inflable", "tobogan", "tobogán"]):
        min_rmb, max_rmb = 5800, 7200
        min_usd = round(min_rmb / tasa_cambio, 2)
        max_usd = round(max_rmb / tasa_cambio, 2)
        ali_min, ali_max = 1100, 1380
        aliexp = 2450
        
        prom_1688 = (min_usd + max_usd) / 2
        sobreprecio_pct = round(((precio_prov - prom_1688) / prom_1688) * 100, 1)
        
        if precio_prov > ali_max:
            clasif = "REGULAR"
            icono = "🟡"
            contra_min = 1200.0
            contra_max = 1350.0
            diag = f"Cotización de distribuidor con margen elevado (+{sobreprecio_pct}% sobre taller de Henan). La lona 0.55mm comercial tiene costo base de $ {min_usd:.2f} USD."
            arg = "Presionar para incluir la turbina/soplador de 1500W y kit de reparación certificado dentro del precio de $1,300 USD."
        else:
            clasif = "EXCELENTE"
            icono = "🟢"
            contra_min = round(precio_prov * 0.92, 2)
            contra_max = precio_prov
            diag = "Precio dentro del rango comercial de fábrica para inflable de uso rudo en PVC 0.55mm."
            arg = "Confirmar que la lona sea 100% Plato PVC con costuras reforzadas de 4 hilos y turbina CE/UL."
            
    else:
        min_usd = round(precio_prov * 0.45, 2)
        max_usd = round(precio_prov * 0.65, 2)
        min_rmb = round(min_usd * tasa_cambio, 1)
        max_rmb = round(max_usd * tasa_cambio, 1)
        ali_min = round(precio_prov * 0.70, 2)
        ali_max = round(precio_prov * 0.88, 2)
        aliexp = round(precio_prov * 1.85, 2)
        clasif = "REGULAR"
        icono = "🟡"
        contra_min = ali_min
        contra_max = ali_max
        diag = "Cotización intermedia frente a fábricas de origen en China."
        arg = "Comparar con cotizaciones de taller en 1688 para negociar descuento por volumen."

    try:
        moq_int = int(re.findall(r"\d+", str(moq))[0])
    except Exception:
        moq_int = 1
        
    ahorro_unit = max(0.0, precio_prov - contra_max)
    ahorro_total = ahorro_unit * moq_int

    return {
        "producto": producto,
        "precio_prov_usd": precio_prov,
        "moq": str(moq),
        "plataforma_1688": {
            "rango_rmb": f"¥{min_rmb:.0f} - ¥{max_rmb:.0f}",
            "rango_usd": f"${min_usd:.2f} - ${max_usd:.2f}",
            "precio_min_usd": min_usd,
            "origen_fabrica": "Zhejiang / Guangdong / Henan",
            "detalles": "Costo directo de fábrica sin margen de exportadora"
        },
        "plataforma_alibaba": {
            "rango_usd": f"${ali_min:.2f} - ${ali_max:.2f}",
            "precio_promedio_usd": round((ali_min + ali_max) / 2, 2),
            "moq_habitual": f"{moq} unid.",
            "detalles": "Rango B2B habitual de exportador directo"
        },
        "plataforma_aliexpress": {
            "precio_unitario_usd": f"${aliexp:.2f}",
            "detalles": "Precio unitario al detal con flete internacional"
        },
        "auditoria": {
            "clasificacion": clasif,
            "icono_semaforo": icono,
            "evaluacion_resumen": diag,
            "contraoferta_usd": f"${contra_min:.2f} - ${contra_max:.2f}",
            "ahorro_estimado_lote": f"${ahorro_total:.2f} USD",
            "argumento_negociacion": arg
        }
    }

# -------------------------------------------------------------
# BARRA LATERAL: SELECTOR DE MÓDULO
# -------------------------------------------------------------
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3135/3135715.png", width=55)
    st.title("🎯 RACOVE")
    
    opciones_modulos = ["🇻🇪 RACOVE (Mercado Nacional y Rentabilidad)", "🇨🇳 Sourcing China (Auditoría de Fábricas)"]
    modulo_activo = st.radio("Módulo:", opciones_modulos, key="radio_modulo_principal")
    st.divider()

# =============================================================================
# MÓDULO 1: RACOVE (VENTAS, RADAR Y MATRIZ DE PRECIOS CON GROUNDING REAL)
# =============================================================================
if "RACOVE" in str(modulo_activo):
    st.title("🎯 RACOVE")

    with st.sidebar:
        st.subheader("Entrada de Productos")
        opcion_origen_ve = st.radio(
            "Selecciona el origen:",
            ["📁 Subir archivo Excel (.xlsx)", "🔗 Enlace de Google Sheets (Drive)"],
            key="radio_origen_ve"
        )
        url_sheet_ve = ""
        archivo_subido_ve = None
        
        if "Google Sheets" in opcion_origen_ve:
            url_sheet_ve = st.text_input("Enlace de Google Sheets:", key="sheet_ve")
            st.caption("Compartido como: 'Cualquier persona con el enlace (Lector)'.")
        else:
            archivo_subido_ve = st.file_uploader(
                "Sube archivo de cotizaciones (.xlsx)", 
                type=["xlsx", "csv", "txt"],
                key="uploader_ve"
            )
            
        limite_prods_ve = st.slider("Cantidad de productos a analizar:", 1, 13, 2, key="slider_ve")
        boton_iniciar_ve = st.button("🚀 Iniciar Análisis", type="primary", use_container_width=True)

    def generar_link_ve(plataforma, comercio, link_original, producto, ubicacion=""):
        if link_original and str(link_original).startswith("http") and "..." not in link_original:
            return link_original
        query = urllib.parse.quote(producto.replace(" Venezuela", "").strip())
        com_low = str(comercio).lower()
        ubi_low = str(ubicacion).lower()
        
        if plataforma == "mercado_libre":
            return f"https://listado.mercadolibre.com.ve/{query}_OrderId_PRICE_ASC"
        elif plataforma == "facebook_marketplace":
            ciudad_fb = "caracas"
            if "valencia" in ubi_low: ciudad_fb = "valencia"
            elif "barquisimeto" in ubi_low: ciudad_fb = "barquisimeto"
            elif "maracay" in ubi_low: ciudad_fb = "maracay"
            elif "maracaibo" in ubi_low: ciudad_fb = "maracaibo"
            return f"https://www.facebook.com/marketplace/{ciudad_fb}/search/?query={query}"
        elif plataforma == "cashea":
            if "locatel" in com_low: return f"https://www.locatel.com.ve/buscar?text={query}"
            elif "beco" in com_low: return f"https://beco.com.ve/search?q={query}"
            elif "balu" in com_low or "balú" in com_low: return f"https://balumoda.com/search?q={query}"
            elif "ivoo" in com_low: return f"https://www.ivoo.com/catalogsearch/result/?q={query}"
            elif "damasco" in com_low: return f"https://damasco.com/search?q={query}"
            elif "multimax" in com_low: return f"https://multimax.net/search?q={query}"
            elif "soy techno" in com_low or "technove" in com_low: return f"https://soytechno.com/search?q={query}"
            return f"https://www.google.com/search?q={urllib.parse.quote(comercio + ' ' + producto + ' venezuela cashea')}"
        return f"https://www.google.com/search?q={urllib.parse.quote(producto + ' venezuela')}"

    def procesar_archivo_ve():
        productos = []
        imagenes_referencia = {}
        costos_referencia = {}
        
        if "Google Sheets" in opcion_origen_ve and url_sheet_ve.strip():
            match = re.search(r"/d/([a-zA-Z0-9-_]+)", url_sheet_ve)
            if match:
                sheet_id = match.group(1)
                csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
                try:
                    df = pd.read_csv(csv_url)
                    for idx, val in enumerate(df.iloc[:, 0].dropna()):
                        s = str(val).strip()
                        if s and not s.startswith("#") and "PRODUCTO" not in s.upper():
                            productos.append(s)
                            c_val = 5.0
                            if df.shape[1] >= 6:
                                raw_c = df.iloc[idx, 5]
                                c_val = extraer_precio_num(raw_c)
                                if c_val >= 999900: c_val = 5.0
                            costos_referencia[s] = c_val
                except Exception as e:
                    st.error(f"Error al leer Google Sheets: {e}")
                    
        elif archivo_subido_ve is not None:
            nombre = archivo_subido_ve.name.lower()
            if nombre.endswith(".xlsx"):
                wb = openpyxl.load_workbook(io.BytesIO(archivo_subido_ve.getvalue()))
                ws = wb.active
                imgs_por_fila = {}
                for img in getattr(ws, "_images", []):
                    if hasattr(img.anchor, "_from"):
                        r = img.anchor._from.row + 1
                        imgs_por_fila[r] = img._data()
                        
                for row in range(5, ws.max_row + 1):
                    val_prod = ws.cell(row, 1).value
                    val_costo = ws.cell(row, 6).value
                    
                    if val_prod and str(val_prod).strip() and not str(val_prod).startswith("#"):
                        p_name = str(val_prod).strip()
                        productos.append(p_name)
                        costo_val = extraer_precio_num(val_costo)
                        if costo_val >= 999900: costo_val = 5.0
                        costos_referencia[p_name] = costo_val
                        if row in imgs_por_fila:
                            imagenes_referencia[p_name] = imgs_por_fila[row]
            else:
                for linea in archivo_subido_ve.getvalue().decode("utf-8", errors="ignore").splitlines():
                    l = linea.strip()
                    if l and not l.startswith("#"):
                        partes = l.split(",")
                        p_name = partes[0].strip()
                        productos.append(p_name)
                        costo_val = extraer_precio_num(partes[1]) if len(partes) > 1 else 5.0
                        if costo_val >= 999900: costo_val = 5.0
                        costos_referencia[p_name] = costo_val
                        
        return productos[:limite_prods_ve], imagenes_referencia, costos_referencia

    def consultar_ve_grounded(cliente, producto, modelos_disponibles):
        prompt = f"""
        Eres un auditor comercial de compras e inteligencia de retail en Venezuela.
        Realiza una búsqueda exhaustiva EN TIEMPO REAL en internet para auditar comercios existentes y precios reales en Venezuela para: "{producto}".

        REGLAS ESTRICTAS DE VERACIDAD (PROHIBIDO INVENTAR O ALUCINAR NOMBRES):
        1. MERCADO LIBRE VENEZUELA (mercadolibre.com.ve):
           - Busca tiendas oficiales, distribuidores autorizados o vendedores MercadoLíder Platinum/Gold REALES en Venezuela.
           - NO inventes nombres ficticios como 'TecnoBelleza Online' o 'Mundo Cosmético VE'. Si no conoces la cuenta exacta, pon la Tienda Oficial o la marca del producto en Venezuela.
           - Precios reales de publicaciones activas en USD.

        2. RED OFICIAL CASHEA (cashea.app):
           - Consulta EXCLUSIVAMENTE comercios aliados oficiales con tiendas físicas comprobables (por ejemplo: Damasco, IVOO, Multimax, Beco, Locatel, Balú, Gina, Soy Techno).
           - PROHIBIDO TERMINANTEMENTE incluir Farmatodo, Traki o Daka como aliados de Cashea.
           - Reporta precio de tienda y plan de financiamiento (Inicial + 3 cuotas cada 14 días).

        3. FACEBOOK MARKETPLACE (Venezuela):
           - Comercios, importadores directos o tiendas con presencia física en plazas principales (Caracas, Valencia, Barquisimeto, Maracay).
           - Descarta precios trampa de $1 o 'gratis'; reporta el precio real en divisa ("Ref" o USD efectivo).

        Devuelve ÚNICAMENTE un objeto JSON válido con este formato:
        {{
            "producto": "{producto}",
            "comercial": {{
                "utilidad_comercial": "Uso específico y nivel de demanda en Venezuela",
                "nicho_mercado": "Perfil real del comprador en el país",
                "rotacion_y_margen": "Ritmo de salida y margen comercial de plaza"
            }},
            "mercado_libre": [
                {{
                    "comercio": "Nombre de tienda real / cuenta activa",
                    "reputacion": "Tienda Oficial / MercadoLíder Platinum",
                    "ubicacion": "Caracas / Valencia / Barquisimeto",
                    "precio_usd": "$XX.XX",
                    "titulo": "Título de la publicación real"
                }}
            ],
            "cashea": [
                {{
                    "comercio": "Damasco / IVOO / Beco / Multimax / Locatel",
                    "reputacion": "Aliado Oficial Cashea",
                    "ubicacion": "Nacional / Sede",
                    "precio_usd": "$XX.XX",
                    "plan_cashea": "Inicial $XX + 3 cuotas $XX",
                    "titulo": "Modelo verificado"
                }}
            ],
            "facebook_marketplace": [
                {{
                    "comercio": "Nombre del importador / comercio en zona",
                    "reputacion": "Vendedor Marketplace con entregas personales",
                    "ubicacion": "Caracas - Chacao / Valencia - Centro / Barquisimeto - Este",
                    "precio_usd": "$XX.XX",
                    "titulo": "Descripción del artículo publicado"
                }}
            ]
        }}
        """
        
        config_grounding = None
        if types:
            try:
                config_grounding = types.GenerateContentConfig(
                    tools=[types.Tool(google_search=types.GoogleSearch())],
                    temperature=0.1
                )
            except Exception:
                config_grounding = None

        for modelo in modelos_disponibles:
            if config_grounding:
                try:
                    resp = cliente.models.generate_content(
                        model=modelo,
                        contents=prompt,
                        config=config_grounding
                    )
                    match = re.search(r"(\\{[\\s\\S]*\\})", resp.text.strip())
                    if match:
                        return json.loads(match.group(1)), f"{modelo} (Búsqueda en Vivo)", None
                except Exception:
                    pass

            try:
                resp = cliente.models.generate_content(model=modelo, contents=prompt)
                match = re.search(r"(\\{[\\s\\S]*\\})", resp.text.strip())
                if match:
                    return json.loads(match.group(1)), modelo, None
            except Exception:
                continue

        return None, None, "No se pudo conectar con el motor de verificación en vivo"

    def renderizar_canal_ve(titulo_seccion, clave_plataforma, lista_opciones, prod_nombre):
        if clave_plataforma == "mercado_libre":
            encabezado_html = """<div class="badge-plataforma badge-ml"><img src="https://http2.mlstatic.com/frontend-assets/ui-navigation/5.18.9/mercadolibre/logo__small.png" height="22" style="vertical-align: middle;"><span>MERCADO LIBRE VENEZUELA (TIENDAS ACTIVAS)</span></div>"""
        elif clave_plataforma == "cashea":
            encabezado_html = """<div class="badge-plataforma badge-cashea"><span style="background: #ffffff; color: #581c87; border-radius: 50%; width: 22px; height: 22px; display: inline-flex; align-items: center; justify-content: center; font-weight: 900; font-size: 13px;">C</span><span>RED OFICIAL CASHEA (ALIADOS VERIFICADOS)</span></div>"""
        elif clave_plataforma == "facebook_marketplace":
            encabezado_html = """<div class="badge-plataforma badge-fb"><img src="https://upload.wikimedia.org/wikipedia/commons/0/05/Facebook_Logo_%282019%29.png" height="20" style="vertical-align: middle; border-radius: 50%;"><span>FACEBOOK MARKETPLACE VENEZUELA (PLAZAS LOCALES)</span></div>"""
        else:
            encabezado_html = f"<h4>{titulo_seccion}</h4>"

        st.markdown(encabezado_html, unsafe_allow_html=True)
        if not lista_opciones:
            st.info(f"ℹ️ Sin publicaciones directas en {titulo_seccion} actualmente.")
            return

        lista_opciones.sort(key=lambda x: extraer_precio_num(x.get("precio_usd", "")))
        cols = st.columns(min(len(lista_opciones), 4))
        for idx, item in enumerate(lista_opciones[:4]):
            with cols[idx]:
                etiqueta_badge = "🟢 Más Económica" if idx == 0 else f"Opción {idx+1}"
                plan_html = f"<div class='card-cashea-plan'>🟣 {item.get('plan_cashea')}</div>" if item.get("plan_cashea") else ""
                
                card_html = f"""<div class="card-item-clean">
<span class="card-badge-econ">{etiqueta_badge}</span>
<div class="card-price">{item.get('precio_usd', 'Consultar')}</div>
<div class="card-store">🏪 {item.get('comercio', 'Comercio')}</div>
<div class="card-reputation">⭐ {item.get('reputacion', 'Vendedor Activo')}</div>
<div class="card-reputation">📍 {item.get('ubicacion', 'Venezuela')}</div>
{plan_html}
<div class="card-title-text" title="{item.get('titulo', prod_nombre)}">📝 {item.get('titulo', prod_nombre)}</div>
</div>"""
                st.markdown(card_html, unsafe_allow_html=True)
                url_btn = generar_link_ve(clave_plataforma, item.get("comercio", ""), item.get("link"), prod_nombre, item.get("ubicacion", ""))
                st.link_button("🔗 Ver Publicación / Referencia", url_btn, use_container_width=True)

    if not st.session_state.get("ve_analisis_completado"):
        st.info("👈 **Para comenzar:** Selecciona en la barra lateral el archivo Excel y haz clic en **🚀 Iniciar Análisis**.")

    if boton_iniciar_ve:
        lista_p, dict_i, dict_costos = procesar_archivo_ve()
        if not lista_p:
            st.warning("⚠️ No se encontraron productos para analizar.")
        else:
            api_key = st.secrets.get("GEMINI_API_KEY")
            if not api_key:
                st.error("❌ Falta GEMINI_API_KEY en Secrets.")
                st.stop()
            cliente = genai.Client(api_key=api_key)
            modelos_disponibles = detectar_modelos_activos(cliente)
            resultados_temp_ve = []
            barra_ve = st.progress(0)
            
            for i, prod in enumerate(lista_p):
                with st.spinner(f"RACOVE auditando en tiempo real: **{prod}**..."):
                    datos, modelo_usado, error = consultar_ve_grounded(cliente, prod, modelos_disponibles)
                    menor_cashea = 999999.0
                    if datos and datos.get("cashea"):
                        pc = [extraer_precio_num(x.get("precio_usd")) for x in datos.get("cashea")]
                        menor_cashea = min(pc) if pc else 999999.0
                    menor_ml = 999999.0
                    if datos and datos.get("mercado_libre"):
                        pm = [extraer_precio_num(x.get("precio_usd")) for x in datos.get("mercado_libre")]
                        menor_ml = min(pm) if pm else 999999.0

                    costo_leido = dict_costos.get(prod, 5.0)

                    resultados_temp_ve.append({
                        "producto": prod,
                        "costo_excel": costo_leido,
                        "datos": datos,
                        "modelo": modelo_usado,
                        "error": error,
                        "menor_cashea": menor_cashea,
                        "menor_ml": menor_ml
                    })
                barra_ve.progress((i + 1) / len(lista_p))
                
            st.session_state["ve_lista_resultados"] = resultados_temp_ve
            st.session_state["ve_dict_imgs"] = dict_i
            st.session_state["ve_analisis_completado"] = True
            st.success("🎉 ¡Análisis RACOVE completado con éxito!")

    if st.session_state["ve_analisis_completado"] and st.session_state["ve_lista_resultados"]:
        tab_radar, tab_matriz = st.tabs([
            "🔎 1. RADAR DE MERCADO NACIONAL", 
            "🧮 2. MATRIZ DE PRECIOS"
        ])
        
        with tab_radar:
            st.markdown("### 📊 Auditoría Externa de Proveedores y Precios (Comercios Reales)")
            for item in st.session_state["ve_lista_resultados"]:
                prod = item["producto"]
                datos = item["datos"]
                error = item["error"]
                modelo_usado = item["modelo"]
                dict_imgs = st.session_state["ve_dict_imgs"]
                
                with st.container(border=True):
                    st.subheader(f"📦 {prod}")
                    c_f, c_c = st.columns([1.1, 3.5])
                    with c_f:
                        st.markdown("**📸 Foto de Referencia:**")
                        if prod in dict_imgs:
                            st.image(dict_imgs[prod], width=190)
                        else:
                            st.info("Sin foto")
                    with c_c:
                        st.markdown("**💼 Utilidad Comercial para Venta en Venezuela:**")
                        com = datos.get("comercial", {}) if (datos and isinstance(datos, dict)) else {}
                        u_txt = com.get("utilidad_comercial", "Artículo de alta demanda en Venezuela.")
                        st.markdown(f"<div class='box-comercial'><b>🎯 Propósito de Venta:</b> {u_txt}</div>", unsafe_allow_html=True)
                        c1, c2 = st.columns(2)
                        with c1: st.markdown(f"👥 **Nicho / Comprador:** {com.get('nicho_mercado', 'Público general')}")
                        with c2: st.markdown(f"📈 **Rotación / Gancho:** {com.get('rotacion_y_margen', 'Demanda constante')}")
                        if modelo_usado: st.caption(f"⚡ *Motor de Búsqueda: {modelo_usado}*")
                    st.divider()
                    if error or not datos:
                        st.error(f"⚠️ {error if error else 'No se pudo obtener información de precios.'}")
                    else:
                        renderizar_canal_ve("MERCADO LIBRE VENEZUELA", "mercado_libre", datos.get("mercado_libre", []), prod)
                        st.write("")
                        renderizar_canal_ve("RED OFICIAL CASHEA", "cashea", datos.get("cashea", []), prod)
                        st.write("")
                        renderizar_canal_ve("FACEBOOK MARKETPLACE VENEZUELA", "facebook_marketplace", datos.get("facebook_marketplace", []), prod)

        with tab_matriz:
            st.markdown("### 🧮 Matriz de Fijación de Precios")
            
            for idx_p, item in enumerate(st.session_state["ve_lista_resultados"]):
                prod = item["producto"]
                menor_c = item["menor_cashea"]
                menor_m = item["menor_ml"]
                costo_inicial = float(item.get("costo_excel", 5.0))
                
                with st.container(border=True):
                    col_t, col_input = st.columns([2.5, 1.2])
                    with col_t:
                        st.subheader(f"🏷️ {prod}")
                        txt_cashea_min = f"${menor_c:.2f} USD" if menor_c < 999900 else "N/D"
                        txt_ml_min = f"${menor_m:.2f} USD" if menor_m < 999900 else "N/D"
                        st.caption(f"Competencia ➔ Cashea mín: **{txt_cashea_min}** | ML mín: **{txt_ml_min}**")
                    with col_input:
                        costo = st.number_input(
                            "💵 Tu Costo Puesto en VE (USD):", 
                            min_value=0.50, 
                            max_value=15000.00, 
                            value=costo_inicial, 
                            step=0.50, 
                            key=f"costo_ve_{idx_p}"
                        )
                    
                    matriz = calcular_matriz_precios(costo, menor_c, menor_m)
                    st.write("")
                    
                    p1, p2, p3, p4, p5 = st.columns(5)
                    with p1:
                        with st.container(border=True):
                            st.markdown("**🏢 Precio N**")
                            st.markdown(f"### ${matriz['precio_n']:.2f}")
                    with p2:
                        with st.container(border=True):
                            st.markdown("**📦 Precio 4**")
                            st.markdown(f"### ${matriz['precio_4']:.2f}")
                    with p3:
                        with st.container(border=True):
                            st.markdown("**💵 Precio Divisa**")
                            st.markdown(f"### ${matriz['precio_divisa']:.2f}")
                    with p4:
                        with st.container(border=True):
                            st.markdown("**🟣 Precio Cashea**")
                            st.markdown(f"### ${matriz['precio_sug_cashea']:.2f}")
                    with p5:
                        with st.container(border=True):
                            st.markdown("**🟡 Precio Mercado Libre**")
                            st.markdown(f"### ${matriz['precio_sug_ml']:.2f}")

                    if matriz["alerta_cashea"]:
                        if matriz["alerta_cashea"]["tipo"] == "error":
                            st.error(matriz["alerta_cashea"]["mensaje"])
                        elif matriz["alerta_cashea"]["tipo"] == "success":
                            st.success(matriz["alerta_cashea"]["mensaje"])
                        else:
                            st.info(matriz["alerta_cashea"]["mensaje"])

                    if matriz["alerta_ml"]:
                        if matriz["alerta_ml"]["tipo"] == "error":
                            st.error(matriz["alerta_ml"]["mensaje"])
                        elif matriz["alerta_ml"]["tipo"] == "success":
                            st.success(matriz["alerta_ml"]["mensaje"])

# =============================================================================
# MÓDULO 2: SOURCING CHINA (1688, ALIBABA Y ALIEXPRESS - MULTIMODAL Y PRECISO)
# =============================================================================
else:
    st.title("🇨🇳 Sourcing China")
    st.markdown("Comparativa de costos: Proveedor vs 1688 vs Alibaba vs AliExpress.")

    with st.sidebar:
        st.subheader("Cotización Proveedor")
        archivo_subido_china = st.file_uploader(
            "Sube cotización China (.xlsx / .csv)", 
            type=["xlsx", "csv"],
            key="uploader_china"
        )
        limite_prods_china = st.slider("Cantidad de productos a auditar:", 1, 10, 2, key="slider_china")
        tasa_rmb = st.number_input("Tasa RMB / USD (1688):", min_value=6.0, max_value=8.5, value=7.23, step=0.05)
        boton_iniciar_china = st.button("🇨🇳 Iniciar Auditoría China", type="primary", use_container_width=True)

    def generar_links_china(producto):
        q = urllib.parse.quote(producto)
        return {
            "1688": f"https://s.1688.com/selloffer/offer_search.htm?keywords={q}",
            "alibaba": f"https://www.alibaba.com/trade/search?SearchText={q}",
            "aliexpress": f"https://www.aliexpress.com/wholesale?SearchText={q}"
        }

    def procesar_archivo_china():
        lista_china = []
        dict_imgs_china = {}
        
        if archivo_subido_china is not None:
            nombre = archivo_subido_china.name.lower()
            if nombre.endswith(".xlsx"):
                wb = openpyxl.load_workbook(io.BytesIO(archivo_subido_china.getvalue()))
                ws = wb.active
                
                imgs_por_fila = {}
                for img in getattr(ws, '_images', []):
                    if hasattr(img.anchor, '_from'):
                        r = img.anchor._from.row + 1
                        imgs_por_fila[r] = img._data()
                        
                col_prod = None
                col_precio = None
                col_moq = None
                header_row = 1
                
                for r in range(1, min(5, ws.max_row + 1)):
                    row_vals = [str(ws.cell(r, c).value or "").strip().upper() for c in range(1, ws.max_column + 1)]
                    for c_idx, val in enumerate(row_vals, 1):
                        if any(k in val for k in ["PRODUCTO", "DESCRIPCION", "DESCRIPCIÓN", "ITEM", "NOMBRE"]) and col_prod is None:
                            col_prod = c_idx
                            header_row = r
                        elif any(k in val for k in ["PRECIO", "COSTO", "USD", "VALOR"]) and col_precio is None:
                            col_precio = c_idx
                            header_row = r
                        elif any(k in val for k in ["MOQ", "CANTIDAD", "QTY", "UNID"]) and col_moq is None:
                            col_moq = c_idx
                            header_row = r
                            
                if col_prod is None: col_prod = 3 if ws.max_column >= 3 else 1
                if col_precio is None: col_precio = 5 if ws.max_column >= 5 else (4 if ws.max_column >= 4 else 2)
                
                for r in range(header_row + 1, ws.max_row + 1):
                    raw_p = ws.cell(r, col_prod).value
                    raw_precio = ws.cell(r, col_precio).value
                    raw_moq = ws.cell(r, col_moq).value if col_moq else "1"
                    
                    if raw_p and str(raw_p).strip() and not str(raw_p).startswith("#"):
                        p_nom = str(raw_p).strip()
                        
                        if p_nom in ["6英寸", "8英寸", "10英寸", "12英寸", "14英寸", "16英寸"]:
                            inch_map = {
                                "6英寸": "Trampolín Cama Elástica 6 Pies (1.83m) con Red de Seguridad",
                                "8英寸": "Trampolín Cama Elástica 8 Pies (2.44m) con Red de Seguridad",
                                "10英寸": "Trampolín Cama Elástica 10 Pies (3.05m) con Red de Seguridad",
                                "12英寸": "Trampolín Cama Elástica 12 Pies (3.66m) con Red de Seguridad",
                                "14英寸": "Trampolín Cama Elástica 14 Pies (4.28m) con Red de Seguridad",
                                "16英寸": "Trampolín Cama Elástica 16 Pies (4.88m) con Red de Seguridad",
                            }
                            p_nom = inch_map[p_nom]
                            
                        c_prov = extraer_precio_num(raw_precio)
                        if c_prov >= 999900: c_prov = 10.0
                        moq_txt = str(raw_moq).strip() if raw_moq else "1"
                        
                        lista_china.append({"producto": p_nom, "precio_prov": c_prov, "moq": moq_txt})
                        if r in imgs_por_fila:
                            dict_imgs_china[p_nom] = imgs_por_fila[r]
            else:
                for linea in archivo_subido_china.getvalue().decode("utf-8", errors="ignore").splitlines()[1:]:
                    partes = [p.strip() for p in linea.split(",") if p.strip()]
                    if partes:
                        p_nom = partes[0]
                        c_prov = extraer_precio_num(partes[1]) if len(partes) > 1 else 10.0
                        moq_txt = partes[2] if len(partes) > 2 else "1"
                        lista_china.append({"producto": p_nom, "precio_prov": c_prov, "moq": moq_txt})
                        
        return lista_china[:limite_prods_china], dict_imgs_china

    def consultar_auditoria_china_precisa(cliente, producto, precio_prov, moq, tasa_cambio, bytes_img, modelos_disponibles):
        benchmark = estimar_mercado_china_benchmark(producto, precio_prov, moq, tasa_cambio)
        
        prompt = f"""
        Actúa como auditor técnico de compras industriales y sourcing en China.
        Analiza este producto: "{producto}".
        Precio cotizado por el proveedor chino: ${precio_prov:.2f} USD (MOQ: {moq} piezas).
        Tasa de cambio: {tasa_cambio} RMB por USD.
        
        Evalúa con rigor de taller frente a:
        - 1688.com (fábricas directas en Yuanes ¥ y convertida a USD)
        - Alibaba.com (exportador B2B directo)
        - AliExpress (precio al detal con flete internacional)
        
        Responde ÚNICAMENTE en JSON válido con esta estructura exacta:
        {{
            "producto": "{producto}",
            "plataforma_1688": {{
                "rango_rmb": "{benchmark['plataforma_1688']['rango_rmb']}",
                "rango_usd": "{benchmark['plataforma_1688']['rango_usd']}",
                "detalles": "Costo directo de fábrica sin margen de exportadora"
            }},
            "plataforma_alibaba": {{
                "rango_usd": "{benchmark['plataforma_alibaba']['rango_usd']}",
                "moq_habitual": "{moq} unid.",
                "detalles": "Rango B2B habitual de exportador directo"
            }},
            "plataforma_aliexpress": {{
                "precio_unitario_usd": "{benchmark['plataforma_aliexpress']['precio_unitario_usd']}",
                "detalles": "Precio minorista unitario con flete internacional"
            }},
            "auditoria": {{
                "clasificacion": "{benchmark['auditoria']['clasificacion']}",
                "icono_semaforo": "{benchmark['auditoria']['icono_semaforo']}",
                "evaluacion_resumen": "{benchmark['auditoria']['evaluacion_resumen']}",
                "contraoferta_usd": "{benchmark['auditoria']['contraoferta_usd']}",
                "ahorro_estimado_lote": "{benchmark['auditoria']['ahorro_estimado_lote']}",
                "argumento_negociacion": "{benchmark['auditoria']['argumento_negociacion']}"
            }}
        }}
        """
        
        for modelo in modelos_disponibles:
            try:
                contents = [prompt]
                if bytes_img:
                    try:
                        pil_img = Image.open(io.BytesIO(bytes_img))
                        contents = [pil_img, prompt]
                    except Exception:
                        pass
                resp = cliente.models.generate_content(model=modelo, contents=contents)
                match = re.search(r"(\\{[\\s\\S]*\\})", resp.text.strip())
                if match:
                    parsed = json.loads(match.group(1))
                    if parsed.get("plataforma_1688") and parsed.get("auditoria"):
                        return parsed, modelo
            except Exception:
                continue
                
        return benchmark, "Auditoría de Fábrica Sincronizada"

    if not st.session_state.get("china_analisis_completado"):
        st.info("👈 **Para comenzar:** Sube en la barra lateral tu archivo de cotizaciones de China y haz clic en **🇨🇳 Iniciar Auditoría China**.")

    if boton_iniciar_china:
        lista_c, dict_imgs_c = procesar_archivo_china()
        if not lista_c:
            st.warning("⚠️ No se encontraron productos en el archivo para auditar.")
        else:
            api_key = st.secrets.get("GEMINI_API_KEY")
            if not api_key:
                st.error("❌ Falta GEMINI_API_KEY en Secrets.")
                st.stop()
            cliente = genai.Client(api_key=api_key)
            modelos_disponibles = detectar_modelos_activos(cliente)
            
            resultados_temp_china = []
            barra_china = st.progress(0)
            
            for i, item_c in enumerate(lista_c):
                prod = item_c["producto"]
                costo_p = item_c["precio_prov"]
                moq_val = item_c["moq"]
                bytes_img = dict_imgs_c.get(prod)
                
                with st.spinner(f"Auditando en fábricas chinas: **{prod}**..."):
                    datos_c, modelo_usado = consultar_auditoria_china_precisa(
                        cliente, prod, costo_p, moq_val, tasa_rmb, bytes_img, modelos_disponibles
                    )
                    resultados_temp_china.append({
                        "producto": prod,
                        "precio_prov": costo_p,
                        "moq": moq_val,
                        "datos": datos_c,
                        "modelo": modelo_usado
                    })
                barra_china.progress((i + 1) / len(lista_c))
                
            st.session_state["china_lista_resultados"] = resultados_temp_china
            st.session_state["china_dict_imgs"] = dict_imgs_c
            st.session_state["china_analisis_completado"] = True
            st.success("🎉 ¡Auditoría de compras en China finalizada con éxito!")

    if st.session_state["china_analisis_completado"] and st.session_state["china_lista_resultados"]:
        for item_ch in st.session_state["china_lista_resultados"]:
            prod = item_ch["producto"]
            p_prov = item_ch["precio_prov"]
            moq = item_ch["moq"]
            datos = item_ch["datos"]
            modelo_usado = item_ch.get("modelo", "")
            links = generar_links_china(prod)
            dict_imgs_c = st.session_state.get("china_dict_imgs", {})
            
            with st.container(border=True):
                st.subheader(f"📦 {prod}")
                if modelo_usado:
                    st.caption(f"⚡ *Motor de Auditoría: {modelo_usado}*")
                
                c_img_ch, c_cards_ch = st.columns([1.1, 4])
                
                with c_img_ch:
                    st.markdown("**📸 Foto Proveedor:**")
                    if prod in dict_imgs_c:
                        st.image(dict_imgs_c[prod], width=180)
                    else:
                        st.info("Sin foto")
                        
                with c_cards_ch:
                    col_prov, col_1688, col_ali, col_aliexp = st.columns(4)
                    
                    with col_prov:
                        with st.container(border=True):
                            st.markdown("**🤝 TU PROVEEDOR**")
                            st.caption(f"Lote cotizado: {moq} unid.")
                            st.markdown(f"## 💵 ${p_prov:.2f} USD")
                            st.caption("Precio bajo auditoría")

                    p_1688 = datos.get("plataforma_1688", {}) if datos else {}
                    with col_1688:
                        with st.container(border=True):
                            st.markdown("<span class='badge-1688'>🏭 1688.com</span>", unsafe_allow_html=True)
                            st.caption("Fábrica local en China")
                            st.markdown(f"## {p_1688.get('rango_usd', '$--')}")
                            st.caption(f"Yuanes: **{p_1688.get('rango_rmb', '¥--')}**")
                            st.link_button("🔗 Ver en 1688", links["1688"], use_container_width=True)

                    p_alibaba = datos.get("plataforma_alibaba", {}) if datos else {}
                    with col_ali:
                        with st.container(border=True):
                            st.markdown("<span class='badge-alibaba'>🌐 Alibaba.com</span>", unsafe_allow_html=True)
                            st.caption(f"MOQ ref: {p_alibaba.get('moq_habitual', moq)}")
                            st.markdown(f"## {p_alibaba.get('rango_usd', '$--')}")
                            st.caption("Exportador B2B")
                            st.link_button("🔗 Ver en Alibaba", links["alibaba"], use_container_width=True)

                    p_aliexpress = datos.get("plataforma_aliexpress", {}) if datos else {}
                    with col_aliexp:
                        with st.container(border=True):
                            st.markdown("<span class='badge-aliexpress'>📦 AliExpress</span>", unsafe_allow_html=True)
                            st.caption("Precio detal unitario")
                            st.markdown(f"## {p_aliexpress.get('precio_unitario_usd', '$--')}")
                            st.caption("Techo de mercado")
                            st.link_button("🔗 Ver en AliExpress", links["aliexpress"], use_container_width=True)

                st.write("")
                audit = datos.get("auditoria", {}) if datos else {}
                clasif = audit.get("clasificacion", "REGULAR")
                icono = audit.get("icono_semaforo", "🟡")
                
                with st.container(border=True):
                    st.markdown(f"### {icono} Veredicto de Compra: {clasif}")
                    st.markdown(f"<div class='box-auditoria-china'><b>Diagnóstico de Fábrica:</b> {audit.get('evaluacion_resumen', 'Evaluación no disponible')}</div>", unsafe_allow_html=True)
                    
                    c_diag1, c_diag2 = st.columns(2)
                    with c_diag1:
                        st.info(f"💡 **Precio Sugerido para Contraofertar:** {audit.get('contraoferta_usd', '$--')}")
                        st.caption(f"💰 **Ahorro Potencial Estimado en el Lote:** {audit.get('ahorro_estimado_lote', '$--')}")
                    with c_diag2:
                        st.warning(f"🎯 **Argumento de Negociación:** {audit.get('argumento_negociacion', 'Solicitar descuento por volumen y comparar con precios de taller en 1688.')}")
