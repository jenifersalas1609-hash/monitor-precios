import streamlit as st
import time
import pandas as pd
import json
import re
import io
import openpyxl
from PIL import Image
from google import genai

st.set_page_config(
    page_title="Monitor de Precios - IA",
    page_icon="🛒",
    layout="wide"
)

st.title("🛒 Monitor de Precios: Mercado Libre vs. Cashea")
st.markdown("Analiza productos con tu imagen de referencia y encuentra las mejores ofertas en Venezuela de menor a mayor precio.")

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
        
    limite_productos = st.slider("Cantidad de productos a analizar:", min_value=1, max_value=15, value=5)
    boton_iniciar = st.button("🔍 Iniciar Monitoreo", type="primary", use_container_width=True)

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
                # Tomar la primera columna con datos
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
            
            # Extraer imágenes incrustadas por fila
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
    Investiga en tiempo real precios en Venezuela para: "{producto}".
    Encuentra 2 o 3 opciones reales en:
    - Mercado Libre Venezuela (mercadolibre.com.ve)
    - Comercios aliados de Cashea (Ivoo, Damasco, SoyTechno, Multimax, etc.)
    
    Ordena las opciones de la MÁS BARATA a la MÁS CARA en USD.
    
    Responde estrictamente con un JSON válido:
    {{
        "producto": "{producto}",
        "opciones": [
            {{
                "tienda": "Mercado Libre / Ivoo / Damasco / etc.",
                "titulo": "Título de la publicación",
                "precio_usd": "$XX",
                "es_cashea": true,
                "plan_cashea": "Inicial $XX + cuotas (si aplica)",
                "link": "URL",
                "detalles": "Garantía o disponibilidad"
            }}
        ]
    }}
    """
    modelos = ["gemini-2.5-flash", "gemini-1.5-flash"]
    for modelo in modelos:
        for intento in range(2):
            try:
                resp = cliente.models.generate_content(model=modelo, contents=prompt)
                txt = resp.text.strip()
                if txt.startswith("```json"): txt = txt[7:]
                if txt.startswith("```"): txt = txt[3:]
                if txt.endswith("```"): txt = txt[:-3]
                return json.loads(txt.strip()), None
            except Exception as e:
                if any(k in str(e) for k in ["503", "429"]):
                    time.sleep(4)
                    continue
                break
    return None, "Servicio ocupado temporalmente."

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
                    
                    # Layout: Imagen de referencia a la izquierda, resultados a la derecha
                    col_ref, col_res = st.columns([1, 2.5])
                    
                    with col_ref:
                        st.markdown("**📸 Tu Producto de Referencia:**")
                        if prod in dict_imgs:
                            st.image(dict_imgs[prod], use_container_width=True)
                        else:
                            st.info("Sin foto en el archivo")
                            
                    with col_res:
                        if error or not datos or "opciones" not in datos:
                            st.warning("No se pudieron cargar opciones en este momento.")
                        else:
                            opciones = datos["opciones"]
                            opciones.sort(key=lambda x: extraer_precio_num(x.get("precio_usd", "")))
                            
                            cols_opc = st.columns(min(len(opciones), 3))
                            titulos = ["🟢 Más Económica", "🟡 Intermedia", "🟣 Aliado Cashea"]
                            
                            for idx, op in enumerate(opciones[:3]):
                                with cols_opc[idx]:
                                    st.markdown(f"**{titulos[idx] if idx < len(titulos) else f'Opción {idx+1}'}**")
                                    st.markdown(f"### 💵 {op.get('precio_usd', 'Consultar')}")
                                    st.markdown(f"🏪 **{op.get('tienda', 'Tienda')}**")
                                    if op.get("es_cashea") and op.get("plan_cashea"):
                                        st.caption(f"🟣 {op.get('plan_cashea')}")
                                    st.caption(f"📝 {op.get('titulo', prod)}")
                                    link = op.get("link") or "[https://www.mercadolibre.com.ve](https://www.mercadolibre.com.ve)"
                                    st.link_button("🔗 Ver Producto", link, use_container_width=True)
                                    
            barra.progress((i + 1) / len(lista_prods))
            time.sleep(3)
            
        st.balloons()
        st.success("🎉 ¡Monitoreo completado!")
