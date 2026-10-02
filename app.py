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
from google.genai import types

st.set_page_config(
    page_title="Monitor de Precios - IA",
    page_icon="🛒",
    layout="wide"
)

st.title("🛒 Monitor de Precios: Mercado Libre vs. Cashea")
st.markdown("Analiza productos con su imagen de referencia y compara opciones de menor a mayor precio en comercios de Venezuela.")

with st.sidebar:
    st.header("⚙️️ Entrada de Productos")
    
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
        st.caption("Debe estar compartido como: 'Cualquier persona con el enlace (Lector)'.")
    else:
        archivo_subido = st.file_uploader(
            "Sube tu archivo de cotizaciones", 
            type=["xlsx", "csv", "txt"],
            help="Sube tu archivo COTIZACION_PRODUCTOS_BELLEZA_CON_IMAGENES.xlsx"
        )
        
    limite_productos = st.slider("Cantidad de productos a analizar:", min_value=1, max_value=15, value=3)
    boton_iniciar = st.button("🔍 Iniciar Monitoreo", type="primary", use_container_width=True)

def generar_link(tienda, link_original, producto):
    if link_original and str(link_original).startswith("http") and "..." not in link_original:
        return link_original
    q = urllib.parse.quote(producto.replace(" Venezuela", ""))
    t_lower = str(tienda).lower()
    if "mercado libre" in t_lower:
        return f"https://listado.mercadolibre.com.ve/{q}"
    elif "farmatodo" in t_lower:
        return f"https://www.farmatodo.com.ve/buscar?producto={q}"
    elif "ivoo" in t_lower:
        return f"https://www.ivoo.com/catalogsearch/result/?q={q}"
    elif "damasco" in t_lower:
        return f"https://damasco.com/search?q={q}"
    elif "cashea" in t_lower:
        return "https://cashea.com"
    return f"https://www.google.com/search?q={urllib.parse.quote(producto + ' ' + tienda)}"

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

def consultar_ofertas(cliente, producto):
    prompt = f"""
    Actúa como un cotizador de compras e investigación de mercado en Venezuela.
    Para el producto: "{producto}", genera 3 opciones comparativas representativas del comercio venezolano.
    
    Debes incluir opciones entre:
    1. Mercado Libre Venezuela (mercadolibre.com.ve)
    2. Comercios aliados a la red Cashea en Venezuela (por ejemplo Farmatodo, Traki, Mundo Total, Ivoo, Damasco, SoyTechno, Multimax, etc.).
    
    Requisitos:
    - Ordena las opciones de la MÁS ECONÓMICA a la MÁS COSTOSA según el precio en USD.
    - Indica el precio en USD (ejemplo: "$14", "$19", "$25").
    - Si la opción es de un aliado Cashea, marca 'es_cashea': true y detalla el pago en cuotas e inicial estimado.
    
    Responde estrictamente con un objeto JSON válido con esta estructura:
    {{
        "producto": "{producto}",
        "opciones": [
            {{
                "tienda": "Mercado Libre / Farmatodo / Ivoo / Traki / etc.",
                "titulo": "Descripción del producto o publicación",
                "precio_usd": "$XX",
                "es_cashea": true,
                "plan_cashea": "Inicial $XX + 3 cuotas de $XX (o dejar vacío si no es Cashea)",
                "link": "URL",
                "detalles": "Disponibilidad o garantía"
            }}
        ]
    }}
    """
    
    modelos = ["gemini-2.0-flash", "gemini-1.5-flash"]
    for modelo in modelos:
        for intento in range(2):
            try:
                resp = cliente.models.generate_content(
                    model=modelo,
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        temperature=0.2
                    )
                )
                txt = resp.text.strip()
                match = re.search(r'(\{[\s\S]*\})', txt)
                if match:
                    data = json.loads(match.group(1))
                    if "opciones" in data and len(data["opciones"]) > 0:
                        return data, None
            except Exception as e:
                err_str = str(e)
                if any(k in err_str for k in ["503", "429"]):
                    time.sleep(4)
                    continue
                break
    return None, "Servicio saturado. Intenta nuevamente."

if boton_iniciar:
    lista_prods, dict_imgs = procesar_archivo()
    
    if not lista_prods:
        st.warning("⚠️ No se encontraron productos para analizar.")
    else:
        st.info(f"📋 Analizando {len(lista_prods)} producto(s)...")
        api_key = st.secrets.get("GEMINI_API_KEY")
        if not api_key:
            st.error("❌ Falta GEMINI_API_KEY en Secrets.")
            st.stop()
            
        cliente = genai.Client(api_key=api_key)
        barra = st.progress(0)
        
        for i, prod in enumerate(lista_prods):
            with st.spinner(f"Buscando ofertas para: **{prod}**..."):
                datos, error = consultar_ofertas(cliente, prod)
                
                with st.container(border=True):
                    st.subheader(f"📦 {prod}")
                    col_ref, col_res = st.columns([1, 2.5])
                    
                    with col_ref:
                        st.markdown("**📸 Tu Producto de Referencia:**")
                        if prod in dict_imgs:
                            st.image(dict_imgs[prod], use_container_width=True)
                        else:
                            st.info("Sin foto en el archivo")
                            
                    with col_res:
                        if error or not datos or "opciones" not in datos:
                            st.warning(f"⚠️ {error if error else 'No se pudieron cargar opciones en este momento.'}")
                        else:
                            opciones = datos["opciones"]
                            opciones.sort(key=lambda x: extraer_precio_num(x.get("precio_usd", "")))
                            
                            cols_opc = st.columns(min(len(opciones), 3))
                            titulos = ["🟢 Más Económica", "🟡 Intermedia", "🟣 Opción Cashea / Alternativa"]
                            
                            for idx, op in enumerate(opciones[:3]):
                                with cols_opc[idx]:
                                    st.markdown(f"**{titulos[idx] if idx < len(titulos) else f'Opción {idx+1}'}**")
                                    st.markdown(f"### 💵 {op.get('precio_usd', 'Consultar')}")
                                    tienda_nombre = op.get('tienda', 'Tienda')
                                    st.markdown(f"🏪 **{tienda_nombre}**")
                                    if op.get("es_cashea") and op.get("plan_cashea"):
                                        st.caption(f"🟣 {op.get('plan_cashea')}")
                                    st.caption(f"📝 {op.get('titulo', prod)}")
                                    
                                    link_final = generar_link(tienda_nombre, op.get("link"), prod)
                                    st.link_button("🔗 Ver Producto", link_final, use_container_width=True)
                                    
            barra.progress((i + 1) / len(lista_prods))
            time.sleep(3)
            
        st.balloons()
        st.success("🎉 ¡Monitoreo completado!")
