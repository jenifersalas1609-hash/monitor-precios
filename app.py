import streamlit as st
import time
import pandas as pd
import json
import re
import io
import openpyxl
import urllib.parse
from PIL import Image
from google import genai

st.set_page_config(
    page_title="Monitor de Precios y Utilidad Comercial - Venezuela",
    page_icon="🔎",
    layout="wide"
)

st.markdown("""
<style>
    .badge-ml { background-color: #ffe600; color: #2d3277; padding: 4px 10px; border-radius: 6px; font-weight: bold; font-size: 0.85rem; }
    .badge-cashea { background-color: #581c87; color: #ffffff; padding: 4px 10px; border-radius: 6px; font-weight: bold; font-size: 0.85rem; }
    .badge-fb { background-color: #1877f2; color: #ffffff; padding: 4px 10px; border-radius: 6px; font-weight: bold; font-size: 0.85rem; }
    .box-comercial {
        background-color: #f1f5f9;
        border-left: 5px solid #2563eb;
        padding: 12px 16px;
        border-radius: 8px;
        margin-bottom: 8px;
    }
</style>
""", unsafe_allow_html=True)

st.title("🔎 Monitor de Precios: Mercado Libre, Cashea y Marketplace")
st.markdown("Análisis de mercado con utilidad comercial de venta y referencias ordenadas de menor a mayor precio.")

with st.sidebar:
    st.header("⚙️ Entrada de Productos")
    opcion_origen = st.radio(
        "Selecciona el origen:",
        ["📁 Subir archivo Excel (.xlsx)", "🔗 Enlace de Google Sheets (Drive)"]
    )
    
    url_sheet = ""
    archivo_subido = None
    
    if "Google Sheets" in opcion_origen:
        url_sheet = st.text_input("Enlace de Google Sheets:")
        st.caption("Asegúrate de compartirlo como: 'Cualquier persona con el enlace (Lector)'.")
    else:
        archivo_subido = st.file_uploader(
            "Sube tu archivo de cotizaciones (.xlsx)", 
            type=["xlsx", "csv", "txt"]
        )
        
    limite_productos = st.slider("Cantidad de productos a analizar:", min_value=1, max_value=13, value=2)
    boton_iniciar = st.button("🔍 Iniciar Análisis Completo", type="primary", use_container_width=True)

def generar_link_directo(plataforma, comercio, link_original, producto):
    if link_original and str(link_original).startswith("http") and "..." not in link_original:
        return link_original
        
    query = urllib.parse.quote(producto.replace(" Venezuela", "").strip())
    com_low = str(comercio).lower()
    
    if plataforma == "mercado_libre":
        return f"https://listado.mercadolibre.com.ve/{query}_OrderId_PRICE_ASC"
    elif plataforma == "facebook_marketplace":
        return f"https://www.facebook.com/marketplace/caracas/search/?query={query}"
    elif plataforma == "cashea":
        if "farmatodo" in com_low:
            return f"https://www.farmatodo.com.ve/buscar?producto={query}"
        elif "ivoo" in com_low:
            return f"https://www.ivoo.com/catalogsearch/result/?q={query}"
        elif "damasco" in com_low:
            return f"https://damasco.com/search?q={query}"
        elif "multimax" in com_low:
            return f"https://multimax.net/search?q={query}"
        return f"https://www.google.com/search?q={urllib.parse.quote(comercio + ' ' + producto + ' venezuela')}"
    return f"https://www.google.com/search?q={urllib.parse.quote(producto + ' venezuela')}"

def procesar_archivo():
    productos = []
    imagenes_referencia = {}
    
    if "Google Sheets" in opcion_origen and url_sheet.strip():
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", url_sheet)
        if match:
            sheet_id = match.group(1)
            csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
            try:
                df = pd.read_csv(csv_url)
                for val in df.iloc[:, 0].dropna():
                    s = str(val).strip()
                    if s and not s.startswith("#") and "PRODUCTO" not in s.upper():
                        productos.append(s)
            except Exception as e:
                st.error(f"Error al leer Google Sheets: {e}")
                
    elif archivo_subido is not None:
        nombre = archivo_subido.name.lower()
        if nombre.endswith(".xlsx"):
            bytes_data = archivo_subido.getvalue()
            wb = openpyxl.load_workbook(io.BytesIO(bytes_data))
            ws = wb.active
            
            imgs_por_fila = {}
            for img in getattr(ws, "_images", []):
                if hasattr(img.anchor, "_from"):
                    r = img.anchor._from.row + 1
                    imgs_por_fila[r] = img._data()
                    
            for row in range(5, ws.max_row + 1):
                val_prod = ws.cell(row, 1).value
                if val_prod and str(val_prod).strip() and not str(val_prod).startswith("#"):
                    p_name = str(val_prod).strip()
                    productos.append(p_name)
                    if row in imgs_por_fila:
                        imagenes_referencia[p_name] = imgs_por_fila[row]
        else:
            contenido = archivo_subido.getvalue().decode("utf-8", errors="ignore")
            for linea in contenido.splitlines():
                l = linea.strip()
                if l and not l.startswith("#"):
                    productos.append(l.split(",")[0].strip())
                    
    return productos[:limite_productos], imagenes_referencia

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
        
    return ["gemini-2.5-flash", "gemini-2.0-flash"]

def consultar_analisis_comercial(cliente, producto, modelos_disponibles):
    prompt = f"""
    Actúa como un experto en importación y ventas mayoristas/detal en Venezuela.
    Realiza el levantamiento de mercado para el siguiente producto: "{producto}".
    
    1. UTILIDAD COMERCIAL PARA VENTA EN VENEZUELA (Sé breve, ultra preciso y con mentalidad de negocio):
       - utilidad_comercial: Explica en 2 líneas exactamente por qué se vende este producto en Venezuela y cuál es su función principal de uso.
       - nicho_mercado: Quién es el comprador final en Venezuela (ej: Maquilladoras a domicilio, salones de belleza, mujeres viajeras, jóvenes para TikTok/Instagram).
       - rotacion_y_margen: Nivel de rotación estimado en el mercado local (Alta rotación / Rotación media) y su gancho comercial.

    2. RELEVAMIENTO DE PRECIOS POR PLATAFORMA (Hasta 4 opciones por canal, de MENOR a MAYOR precio en USD):
       * REGLA DE REFERENCIA: Si no encuentras el modelo idéntico exacto publicado, entrega el producto EQUIVALENTE O REFERENCIA DIRECTA más similar en función y categoría disponible en Venezuela (colocando en detalles: "Referencia equivalente del mercado").
       
       A) MERCADO LIBRE VENEZUELA:
          - Publicaciones activas de vendedores calificados (MercadoLíder o Tiendas Oficiales).
          
       B) RED OFICIAL CASHEA:
          - Tiendas aliadas reales: Ivoo, Damasco, SoyTechno, Multimax, Farmatodo, Locatel, Balú, Beco, Mundo Total.
          * Traki NO está en Cashea (prohibido incluirlo).
          - Incluye precio total y plan estimado de cuotas.
          
       C) FACEBOOK MARKETPLACE VENEZUELA:
          - Precios de importadores o tiendas activas en Caracas, Valencia, Maracay o Maracaibo.
    
    Responde ÚNICAMENTE con este JSON:
    {{
        "producto": "{producto}",
        "comercial": {{
            "utilidad_comercial": "Breve explicación de para qué funciona y por qué tiene demanda comercial en Venezuela.",
            "nicho_mercado": "Público objetivo en Venezuela",
            "rotacion_y_margen": "Estimado de rotación y gancho de venta"
        }},
        "mercado_libre": [
            {{
                "comercio": "Vendedor o Tienda Oficial",
                "reputacion": "MercadoLíder Platinum / Gold / etc.",
                "ubicacion": "Caracas / Valencia / etc.",
                "precio_usd": "$XX",
                "titulo": "Título de la publicación o referencia equivalente",
                "detalles": "Condición o nota de referencia"
            }}
        ],
        "cashea": [
            {{
                "comercio": "Tienda aliada oficial",
                "reputacion": "Comercio Aliado Cashea",
                "ubicacion": "Nivel Nacional / Tiendas físicas",
                "precio_usd": "$XX",
                "plan_cashea": "Inicial $XX + 3 cuotas de $XX",
                "titulo": "Nombre del artículo o referencia en catálogo",
                "detalles": "Disponibilidad"
            }}
        ],
        "facebook_marketplace": [
            {{
                "comercio": "Importadora o tienda",
                "reputacion": "Tienda física / Importador directo",
                "ubicacion": "Sector y Ciudad",
                "precio_usd": "$XX",
                "titulo": "Título de la publicación",
                "detalles": "Modalidad de entrega"
            }}
        ]
    }}
    """
    
    ultimo_error = ""
    for modelo in modelos_disponibles:
        for intento in range(2):
            try:
                resp = cliente.models.generate_content(
                    model=modelo,
                    contents=prompt
                )
                txt = resp.text.strip()
                match = re.search(r"(\{[\s\S]*\})", txt)
                if match:
                    data = json.loads(match.group(1))
                    return data, modelo, None
            except Exception as e:
                err_str = str(e)
                ultimo_error = f"{modelo} -> {err_str}"
                if any(k in err_str for k in ["503", "429", "RESOURCE_EXHAUSTED", "UNAVAILABLE"]):
                    time.sleep(3)
                    continue
                break
                
    return None, None, f"Error: {ultimo_error}"

def renderizar_bloque_canal(titulo_seccion, clave_plataforma, badge_clase, badge_texto, lista_opciones, prod_nombre):
    st.markdown(f"#### <span class='{badge_clase}'>{badge_texto}</span> {titulo_seccion}", unsafe_allow_html=True)
    
    if not lista_opciones or len(lista_opciones) == 0:
        st.info(f"ℹ️ Sin publicaciones directas en {titulo_seccion} actualmente.")
        return

    lista_opciones.sort(key=lambda x: extraer_precio_num(x.get("precio_usd", "")))
    
    cols = st.columns(min(len(lista_opciones), 4))
    for idx, item in enumerate(lista_opciones[:4]):
        with cols[idx]:
            with st.container(border=True):
                etiqueta = "🟢 Opción Más Económica" if idx == 0 else f"Opción {idx+1}"
                st.markdown(f"**{etiqueta}**")
                st.markdown(f"## 💵 {item.get('precio_usd', 'Consultar')}")
                
                comercio = item.get('comercio', 'Comercio')
                st.markdown(f"🏪 **{comercio}**")
                st.caption(f"⭐ **Reputación:** {item.get('reputacion', 'Vendedor Activo')}")
                st.caption(f"📍 **Ubicación:** {item.get('ubicacion', 'Venezuela')}")
                
                if item.get("plan_cashea"):
                    st.info(f"🟣 **Cashea:** {item.get('plan_cashea')}")
                    
                st.caption(f"📝 *{item.get('titulo', prod_nombre)}*")
                if item.get("detalles"):
                    st.caption(f"ℹ️ {item.get('detalles')}")
                    
                url_btn = generar_link_directo(clave_plataforma, comercio, item.get("link"), prod_nombre)
                st.link_button("🔗 Ver Publicación / Referencia", url_btn, use_container_width=True)

if boton_iniciar:
    lista_prods, dict_imgs = procesar_archivo()
    
    if not lista_prods:
        st.warning("⚠️ No se encontraron productos para analizar.")
    else:
        st.info(f"📋 Analizando {len(lista_prods)} producto(s) con foco comercial y cotizaciones...")
        api_key = st.secrets.get("GEMINI_API_KEY")
        if not api_key:
            st.error("❌ Falta GEMINI_API_KEY en Secrets.")
            st.stop()
            
        cliente = genai.Client(api_key=api_key)
        
        with st.spinner("Conectando con Google y seleccionando modelo analítico..."):
            modelos_disponibles = detectar_modelos_activos(cliente)
            
        barra = st.progress(0)
        
        for i, prod in enumerate(lista_prods):
            with st.spinner(f"Analizando: **{prod}**..."):
                datos, modelo_usado, error = consultar_analisis_comercial(cliente, prod, modelos_disponibles)
                
                with st.container(border=True):
                    st.subheader(f"📦 {prod}")
                    
                    # Fila Superior: Tu foto del Excel + Ficha de Utilidad Comercial en Venezuela
                    col_foto, col_comercial = st.columns([1, 2.8])
                    
                    with col_foto:
                        st.markdown("**📸 Tu Foto de Referencia (Excel):**")
                        if prod in dict_imgs:
                            st.image(dict_imgs[prod], use_container_width=True)
                        else:
                            st.info("Sin foto en el archivo")
                            
                    with col_comercial:
                        st.markdown("**💼 Utilidad Comercial para Venta en Venezuela:**")
                        com = datos.get("comercial", {}) if (datos and isinstance(datos, dict)) else {}
                        
                        utilidad_txt = com.get("utilidad_comercial", "Artículo de alta demanda en el mercado de belleza y cuidado personal venezolano.")
                        st.markdown(f"<div class='box-comercial'><b>🎯 Propósito de Venta:</b> {utilidad_txt}</div>", unsafe_allow_html=True)
                        
                        c1, c2 = st.columns(2)
                        with c1:
                            st.markdown(f"👥 **Nicho / Comprador:** {com.get('nicho_mercado', 'Público general / Estilistas')}")
                        with c2:
                            st.markdown(f"📈 **Rotación / Gancho:** {com.get('rotacion_y_margen', 'Demanda constante')}")
                            
                        if modelo_usado:
                            st.caption(f"⚡ *Modelo analítico: {modelo_usado}*")
                            
                    st.divider()
                    
                    if error or not datos:
                        st.error(f"⚠️ {error if error else 'No se pudo obtener la información de precios.'}")
                    else:
                        # 1. MERCADO LIBRE
                        renderizar_bloque_canal(
                            "MERCADO LIBRE VENEZUELA", 
                            "mercado_libre", 
                            "badge-ml", 
                            "🟡 MERCADO LIBRE", 
                            datos.get("mercado_libre", []), 
                            prod
                        )
                        st.write("")
                        
                        # 2. RED OFICIAL CASHEA
                        renderizar_bloque_canal(
                            "RED OFICIAL CASHEA", 
                            "cashea", 
                            "badge-cashea", 
                            "🟣 CASHEA", 
                            datos.get("cashea", []), 
                            prod
                        )
                        st.write("")
                        
                        # 3. FACEBOOK MARKETPLACE
                        renderizar_bloque_canal(
                            "FACEBOOK MARKETPLACE VENEZUELA", 
                            "facebook_marketplace", 
                            "badge-fb", 
                            "🔵 MARKETPLACE", 
                            datos.get("facebook_marketplace", []), 
                            prod
                        )
                        
            barra.progress((i + 1) / len(lista_prods))
            time.sleep(3)
            
        st.balloons()
        st.success("🎉 ¡Análisis y cotizaciones completadas con éxito!")
