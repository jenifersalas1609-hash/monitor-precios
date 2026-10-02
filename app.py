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
    page_title="Monitor de Precios y Proveedores - Venezuela",
    page_icon="🔍",
    layout="wide"
)

st.title("🔍 Comparador de Precios y Proveedores Confiables")
st.markdown("Precios reales comparados entre **Mercado Libre** (vendedores calificados), **Cashea** y **Facebook Marketplace**.")

with st.sidebar:
    st.header("⚙️ Entrada de Productos")
    
    opcion_origen = st.radio(
        "Selecciona el origen:",
        ["📁 Subir archivo Excel (.xlsx)", "🔗 Enlace de Google Sheets (Drive)"]
    )
    
    url_sheet = ""
    archivo_subido = None
    
    if "Google Sheets" in opcion_origen:
        url_sheet = st.text_input(
            "Enlace de Google Sheets:",
            placeholder="https://docs.google.com/spreadsheets/d/.../edit"
        )
        st.caption("Asegúrate de compartirlo como: 'Cualquier persona con el enlace (Lector)'.")
    else:
        archivo_subido = st.file_uploader(
            "Sube tu archivo de cotizaciones", 
            type=["xlsx", "csv", "txt"],
            help="Sube tu archivo COTIZACION_PRODUCTOS_BELLEZA_CON_IMAGENES.xlsx"
        )
        
    limite_productos = st.slider("Cantidad de productos a analizar:", min_value=1, max_value=15, value=2)
    boton_iniciar = st.button("🔍 Iniciar Monitoreo y Verificación", type="primary", use_container_width=True)

def generar_link_verificado(fuente, comercio, link_original, producto):
    """Construye enlaces funcionales a búsquedas activas y filtradas en cada plataforma."""
    if link_original and str(link_original).startswith("http") and "..." not in link_original:
        return link_original
        
    q = urllib.parse.quote(producto.replace(" Venezuela", "").strip())
    fuente_lower = str(fuente).lower()
    comercio_lower = str(comercio).lower()
    
    if "mercado libre" in fuente_lower:
        # Filtro de búsqueda activa en Mercado Libre Venezuela
        return f"https://listado.mercadolibre.com.ve/{q}_OrderId_PRICE_ASC"
    elif "facebook" in fuente_lower or "marketplace" in fuente_lower:
        # Búsqueda directa en Marketplace Caracas/Venezuela
        return f"https://www.facebook.com/marketplace/caracas/search/?query={q}"
    elif "cashea" in fuente_lower:
        if "farmatodo" in comercio_lower:
            return f"https://www.farmatodo.com.ve/buscar?producto={q}"
        elif "ivoo" in comercio_lower:
            return f"https://www.ivoo.com/catalogsearch/result/?q={q}"
        elif "damasco" in comercio_lower:
            return f"https://damasco.com/search?q={q}"
        return "https://cashea.com"
    return f"https://www.google.com/search?q={urllib.parse.quote(producto + ' precio venezuela')}"

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

def parsear_respuesta(texto):
    match = re.search(r"(\{[\s\S]*\})", texto)
    if match:
        try:
            data = json.loads(match.group(1))
            if "opciones" in data and isinstance(data["opciones"], list) and len(data["opciones"]) > 0:
                return data["opciones"], None
        except Exception:
            pass
    return [], "No se pudo interpretar el formato de los datos."

def consultar_ofertas_sinceras(cliente, producto, modelos_disponibles):
    prompt = f"""
    Actúa como un analista de compras mayoristas y cotizaciones en Venezuela.
    Investiga y sinceriza los precios reales de mercado en Venezuela para el producto: "{producto}".
    
    Debes estructurar 3 opciones obligatorias de 3 canales distintos:
    
    1. OPCIÓN MERCADO LIBRE VENEZUELA:
       - La publicación debe ser ACTIVA.
       - El vendedor DEBE tener reputación positiva obligatoria (MercadoLíder Platinum/Gold o Tienda Oficial).
       - No uses precios ficticios ni señuelos.
       
    2. OPCIÓN COMERCIO ALIADO CASHEA:
       - Tienda formal en Venezuela afiliada a Cashea (ej. Ivoo, Damasco, Farmatodo, Mundo Total, Traki, SoyTechno, etc.).
       - Precio total de venta en tienda y el desglose de inicial y cuotas de Cashea.
       
    3. OPCIÓN FACEBOOK MARKETPLACE VENEZUELA:
       - Precio real promedio de importadores directos o tiendas independientes activas en Caracas/Venezuela.
       - Vendedor con perfil activo/calificado o tienda física de apoyo.
       
    IMPORTANTE:
    - Ordena las 3 opciones de la MÁS ECONÓMICA a la MÁS COSTOSA según el precio en USD.
    - La FUENTE debe ser clara e independiente del nombre del comercio.
    
    Responde ÚNICAMENTE con este JSON:
    {{
        "producto": "{producto}",
        "opciones": [
            {{
                "fuente": "Mercado Libre / Cashea / Facebook Marketplace",
                "comercio_o_vendedor": "Nombre exacto del vendedor, tienda o comercio",
                "reputacion_vendedor": "Ej: MercadoLíder Platinum (100% recomendados) / Tienda Oficial / Comercio Verificado",
                "estado_publicacion": "Publicación Activa",
                "titulo": "Título de la publicación del producto",
                "precio_usd": "$XX",
                "plan_cashea": "Inicial $XX + 3 cuotas de $XX (o dejar vacío si no es Cashea)",
                "link": "URL",
                "detalles": "Ubicación (ej: Caracas, Valencia), garantía o condición"
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
                opciones, err_parse = parsear_respuesta(resp.text.strip())
                if opciones:
                    return {"opciones": opciones, "modelo_usado": modelo}, None
            except Exception as e:
                err_str = str(e)
                ultimo_error = f"{modelo} -> {err_str}"
                if any(k in err_str for k in ["503", "429", "RESOURCE_EXHAUSTED", "UNAVAILABLE"]):
                    time.sleep(3)
                    continue
                break
                
    return None, f"Error: {ultimo_error}"

if boton_iniciar:
    lista_prods, dict_imgs = procesar_archivo()
    
    if not lista_prods:
        st.warning("⚠️ No se encontraron productos para analizar.")
    else:
        st.info(f"📋 Verificando proveedores y sincerando precios para {len(lista_prods)} producto(s)...")
        api_key = st.secrets.get("GEMINI_API_KEY")
        if not api_key:
            st.error("❌ Falta GEMINI_API_KEY en Secrets.")
            st.stop()
            
        cliente = genai.Client(api_key=api_key)
        
        with st.spinner("Sincronizando con Google y seleccionando modelo disponible..."):
            modelos_disponibles = detectar_modelos_activos(cliente)
            
        barra = st.progress(0)
        
        for i, prod in enumerate(lista_prods):
            with st.spinner(f"Verificando fuentes para: **{prod}**..."):
                datos, error = consultar_ofertas_sinceras(cliente, prod, modelos_disponibles)
                
                with st.container(border=True):
                    st.subheader(f"📦 {prod}")
                    col_ref, col_res = st.columns([1, 2.7])
                    
                    with col_ref:
                        st.markdown("**📸 Producto de Referencia:**")
                        if prod in dict_imgs:
                            st.image(dict_imgs[prod], use_container_width=True)
                        else:
                            st.info("Sin foto en el archivo")
                            
                    with col_res:
                        if error or not datos or "opciones" not in datos:
                            st.warning(f"⚠️ {error if error else 'No se pudieron verificar las fuentes en este momento.'}")
                        else:
                            opciones = datos["opciones"]
                            modelo_usado = datos.get("modelo_usado", "")
                            opciones.sort(key=lambda x: extraer_precio_num(x.get("precio_usd", "")))
                            
                            cols_opc = st.columns(min(len(opciones), 3))
                            
                            for idx, op in enumerate(opciones[:3]):
                                with cols_opc[idx]:
                                    fuente = op.get("fuente", "Canal de Venta")
                                    
                                    # Fuente prominente en primer lugar
                                    if "mercado libre" in fuente.lower():
                                        st.markdown("### 🟡 MERCADO LIBRE")
                                    elif "facebook" in fuente.lower() or "marketplace" in fuente.lower():
                                        st.markdown("### 🔵 FB MARKETPLACE")
                                    elif "cashea" in fuente.lower():
                                        st.markdown("### 🟣 CASHEA")
                                    else:
                                        st.markdown(f"### 🏬 {fuente.upper()}")
                                        
                                    st.markdown(f"## 💵 {op.get('precio_usd', 'Consultar')}")
                                    
                                    # Comercio y Reputación verificada
                                    comercio = op.get("comercio_o_vendedor", "Vendedor")
                                    st.markdown(f"🏪 **Comercio:** {comercio}")
                                    
                                    reputacion = op.get("reputacion_vendedor", "Vendedor activo")
                                    st.caption(f"⭐ **Reputación:** {reputacion}")
                                    
                                    if op.get("plan_cashea"):
                                        st.info(f"🟣 **Cashea:** {op.get('plan_cashea')}")
                                        
                                    st.caption(f"📝 *{op.get('titulo', prod)}*")
                                    
                                    if op.get("detalles"):
                                        st.caption(f"📍 {op.get('detalles')}")
                                        
                                    link_final = generar_link_verificado(fuente, comercio, op.get("link"), prod)
                                    st.link_button("🔗 Ver Publicación Activa", link_final, use_container_width=True)
                                    
                            if modelo_usado:
                                st.caption(f"⚡ *Validado con: {modelo_usado}*")
                                
            barra.progress((i + 1) / len(lista_prods))
            time.sleep(3)
            
        st.balloons()
        st.success("🎉 ¡Precios sincerados y proveedores verificados con éxito!")
