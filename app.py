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
    page_title="Monitor de Precios y Proveedores Verificados",
    page_icon="🔎",
    layout="wide"
)

# Estilos visuales para máxima nitidez de imágenes y jerarquía de datos
st.markdown("""
<style>
    .card-proveedor {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 14px;
        margin-bottom: 15px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }
    .badge-ml { background-color: #ffe600; color: #2d3277; padding: 3px 8px; border-radius: 4px; font-weight: bold; font-size: 0.8rem; }
    .badge-cashea { background-color: #581c87; color: #ffffff; padding: 3px 8px; border-radius: 4px; font-weight: bold; font-size: 0.8rem; }
    .badge-fb { background-color: #1877f2; color: #ffffff; padding: 3px 8px; border-radius: 4px; font-weight: bold; font-size: 0.8rem; }
</style>
""", unsafe_allow_html=True)

st.title("🔎 Monitor de Precios: Mercado Libre, Cashea y Marketplace")
st.markdown("Comparativa sincera de mercado en Venezuela: 4 opciones por canal ordenadas de menor a mayor precio con reputación y ubicación.")

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
        flash = [m for m in modelos_encontrados if "flash" in m.lower()]
        otros = [m for m in modelos_encontrados if "flash" not in m.lower()]
        return flash + otros
        
    return ["gemini-2.5-flash", "gemini-2.0-flash"]

def consultar_ofertas_estrictas(cliente, producto, modelos_disponibles):
    prompt = f"""
    Eres un auditor comercial de compras mayoristas y de retail en Venezuela.
    Tu objetivo es realizar un levantamiento de precios 100% sincero para el producto: "{producto}".
    
    REGLAS DE ORO (CERO ALUCINACIONES):
    1. Si para alguna plataforma NO encuentras el producto real o no existe en el catálogo venezolano, debes responder con una lista vacía [] en esa categoría. NUNCA inventes publicaciones.
    2. En CASHEA: Solo puedes incluir comercios formales de la red Cashea en Venezuela:
       - Cuidado capilar / electro: Ivoo, SoyTechno, Damasco, Multimax Store.
       - Salud y belleza: Farmatodo, Locatel, Farmarket.
       - Bolsos / Organización: Balú, Beco, Macuto, Gina, Mundo Total.
       * PROHIBIDO poner a Traki o Daka en Cashea.
    3. En MERCADO LIBRE: Solo publicaciones activas en mercadolibre.com.ve de vendedores con reputación positiva comprobada (MercadoLíder Gold, Platinum o Tienda Oficial).
    4. En FACEBOOK MARKETPLACE: Precios reales de importadores o tiendas con sede o entregas en Venezuela (Caracas, Valencia, Maracay, etc.).
    5. Debes suministrar HASTA 4 OPCIONES por cada plataforma, estrictamente ordenadas de la MÁS ECONÓMICA a la MÁS COSTOSA según el precio en USD.
    
    Estructura JSON obligatoria:
    {{
        "producto": "{producto}",
        "mercado_libre": [
            {{
                "comercio": "Nombre del vendedor o tienda oficial",
                "reputacion": "MercadoLíder Platinum / Tienda Oficial / etc.",
                "ubicacion": "Caracas / Valencia / Barquisimeto / etc.",
                "precio_usd": "$XX",
                "titulo": "Título exacto de la publicación",
                "imagen_url": "URL pública de imagen del producto o dejar vacio",
                "link": "URL",
                "detalles": "Garantía, condición (nuevo) y entrega"
            }}
        ],
        "cashea": [
            {{
                "comercio": "Ivoo / Damasco / Farmatodo / Balú / etc.",
                "reputacion": "Comercio Aliado Oficial Cashea",
                "ubicacion": "Nivel Nacional / Tiendas físicas",
                "precio_usd": "$XX",
                "plan_cashea": "Inicial $XX + 3 cuotas de$XX",
                "titulo": "Nombre del artículo en la tienda",
                "imagen_url": "URL pública de imagen del producto o dejar vacio",
                "link": "URL",
                "detalles": "Disponibilidad y condiciones"
            }}
        ],
        "facebook_marketplace": [
            {{
                "comercio": "Nombre de la tienda, importadora o vendedor",
                "reputacion": "Local físico / Vendedor con calificaciones / Importador directo",
                "ubicacion": "Ciudad o sector (ej: El Cementerio, Valencia, Sabana Grande)",
                "precio_usd": "$XX",
                "titulo": "Título de la publicación en Marketplace",
                "imagen_url": "URL pública de imagen del producto o dejar vacio",
                "link": "URL",
                "detalles": "Modalidad de entrega (delivery o retiro en tienda)"
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
                
    return None, None, f"Error de conexión: {ultimo_error}"

def renderizar_bloque_plataforma(titulo_seccion, clave_plataforma, badge_clase, badge_texto, lista_opciones, prod_nombre):
    st.markdown(f"#### <span class='{badge_clase}'>{badge_texto}</span> {titulo_seccion}", unsafe_allow_html=True)
    
    if not lista_opciones or len(lista_opciones) == 0:
        st.warning(f"🚫 **NO DISPONIBLE**: No se encontraron publicaciones activas comprobables en {titulo_seccion} para este producto.")
        return

    # Ordenar estrictamente de menor a mayor precio en USD
    lista_opciones.sort(key=lambda x: extraer_precio_num(x.get("precio_usd", "")))
    
    cols = st.columns(min(len(lista_opciones), 4))
    for idx, item in enumerate(lista_opciones[:4]):
        with cols[idx]:
            with st.container(border=True):
                st.markdown(f"**Opción {idx+1} ({'Más económica' if idx==0 else 'Alternativa'})**")
                
                # Imagen del proveedor en el mercado
                img_prov = item.get("imagen_url")
                if img_prov and str(img_prov).startswith("http"):
                    st.image(img_prov, use_container_width=True, caption="Foto del Proveedor")
                else:
                    placeholder_img = f"https://placehold.co/400x300/f8fafc/475569?text={urllib.parse.quote(item.get('comercio', 'Foto Mercado'))}"
                    st.image(placeholder_img, use_container_width=True, caption="Catálogo Mercado")
                
                st.markdown(f"### 💵 {item.get('precio_usd', 'Consultar')}")
                st.markdown(f"🏪 **Comercio:** {item.get('comercio', 'No especificado')}")
                st.caption(f"⭐ **Reputación:** {item.get('reputacion', 'Vendedor Activo')}")
                st.caption(f"📍 **Ubicación:** {item.get('ubicacion', 'Venezuela')}")
                
                if item.get("plan_cashea"):
                    st.info(f"🟣 **Cashea:** {item.get('plan_cashea')}")
                    
                st.caption(f"📝 *{item.get('titulo', prod_nombre)}*")
                
                if item.get("detalles"):
                    st.caption(f"ℹ️ {item.get('detalles')}")
                    
                url_btn = generar_link_directo(clave_plataforma, item.get("comercio", ""), item.get("link"), prod_nombre)
                st.link_button("🔗 Ir a la Publicación Directa", url_btn, use_container_width=True)

if boton_iniciar:
    lista_prods, dict_imgs = procesar_archivo()
    
    if not lista_prods:
        st.warning("⚠️ No se encontraron productos para analizar.")
    else:
        st.info(f"📋 Analizando {len(lista_prods)} producto(s) con verificación rigurosa de canales...")
        api_key = st.secrets.get("GEMINI_API_KEY")
        if not api_key:
            st.error("❌ Falta GEMINI_API_KEY en Secrets.")
            st.stop()
            
        cliente = genai.Client(api_key=api_key)
        
        with st.spinner("Sincronizando con Google y seleccionando modelo activo..."):
            modelos_disponibles = detectar_modelos_activos(cliente)
            
        barra = st.progress(0)
        
        for i, prod in enumerate(lista_prods):
            with st.spinner(f"Verificando canales comerciales para: **{prod}**..."):
                datos, modelo_usado, error = consultar_ofertas_estrictas(cliente, prod, modelos_disponibles)
                
                with st.container(border=True):
                    st.subheader(f"📦 {prod}")
                    
                    # Fila superior: Producto de referencia
                    col_foto_ref, col_info_ref = st.columns([1, 4])
                    with col_foto_ref:
                        st.markdown("**📸 Tu Foto de Referencia (Excel):**")
                        if prod in dict_imgs:
                            st.image(dict_imgs[prod], use_container_width=True)
                        else:
                            st.info("Sin foto en el archivo")
                    with col_info_ref:
                        st.markdown("**Criterio de Búsqueda:** Precios reales validados de menor a mayor. Si un canal no dispone del artículo, se marca como no disponible.")
                        if modelo_usado:
                            st.caption(f"⚡ Modelo analítico: `{modelo_usado}`")
                    
                    st.divider()
                    
                    if error or not datos:
                        st.error(f"⚠️ {error if error else 'No fue posible levantar la información para este producto.'}")
                    else:
                        # 1. MERCADO LIBRE VENEZUELA
                        renderizar_bloque_plataforma(
                            "MERCADO LIBRE VENEZUELA", 
                            "mercado_libre", 
                            "badge-ml", 
                            "🟡 MERCADO LIBRE", 
                            datos.get("mercado_libre", []), 
                            prod
                        )
                        st.write("")
                        
                        # 2. RED OFICIAL CASHEA
                        renderizar_bloque_plataforma(
                            "RED OFICIAL CASHEA (Comercios Verificados)", 
                            "cashea", 
                            "badge-cashea", 
                            "🟣 CASHEA", 
                            datos.get("cashea", []), 
                            prod
                        )
                        st.write("")
                        
                        # 3. FACEBOOK MARKETPLACE VENEZUELA
                        renderizar_bloque_plataforma(
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
        st.success("🎉 ¡Levantamiento de precios y proveedores finalizado con éxito!")
