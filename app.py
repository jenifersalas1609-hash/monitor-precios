import streamlit as st
import time
import pandas as pd
import json
import re
from google import genai

# Configuración de la interfaz
st.set_page_config(
    page_title="Monitor de Precios - IA",
    page_icon="🛒",
    layout="wide"
)

st.title("🛒 Comparador de Precios: De Menor a Mayor")
st.markdown("Analiza productos desde Google Sheets, Excel (.xlsx), CSV o TXT en tiendas de Venezuela (Mercado Libre y comercios Cashea).")

# Barra lateral
with st.sidebar:
    st.header("⚙️ Configuración de Entrada")
    
    opcion_origen = st.radio(
        "Selecciona el origen de los productos:",
        ["📁 Subir archivo (.xlsx, .csv, .txt)", "🔗 Enlace de Google Sheets (Drive)"]
    )
    
    url_sheet = ""
    archivo_subido = None
    
    if "Google Sheets" in opcion_origen:
        url_sheet = st.text_input(
            "Pega el enlace de tu Google Sheet:",
            placeholder="https://docs.google.com/spreadsheets/d/.../edit"
        )
        st.caption("Asegúrate de compartirlo como: 'Cualquier persona con el enlace (Lector)'.")
    else:
        archivo_subido = st.file_uploader(
            "Sube tu archivo de productos", 
            type=["xlsx", "csv", "txt"],
            help="Soporta libros de Excel (.xlsx), tablas CSV o archivos de texto (.txt)."
        )
        
    limite_productos = st.slider("Cantidad de productos a analizar:", min_value=1, max_value=25, value=5)
    
    st.markdown("---")
    boton_iniciar = st.button("🔍 Iniciar Monitoreo Comparativo", type="primary", use_container_width=True)

def extraer_columna_productos(df):
    """Detecta de forma inteligente la columna con los nombres de productos."""
    col_candidata = None
    # Buscar columnas con nombres clave comunes
    for col in df.columns:
        col_str = str(col).strip().upper()
        if any(k in col_str for k in ["PRODUCTO", "DESCRIPCION", "DESCRIPCIÓN", "ARTICULO", "ARTÍCULO", "NOMBRE"]):
            col_candidata = col
            break
            
    if col_candidata is not None:
        serie = df[col_candidata].dropna()
    else:
        # Si no encuentra un encabezado específico, toma la primera columna con texto
        serie = df.iloc[:, 0].dropna()
        
    productos = [str(x).strip() for x in serie if str(x).strip() and not str(x).strip().startswith("#")]
    return productos

def obtener_productos():
    productos = []
    if "Google Sheets" in opcion_origen and url_sheet.strip():
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", url_sheet)
        if match:
            sheet_id = match.group(1)
            csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
            try:
                df = pd.read_csv(csv_url)
                productos = extraer_columna_productos(df)
            except Exception as e:
                st.error(f"Error al leer Google Sheets: {e}")
        else:
            st.error("El enlace de Google Sheets no es válido.")
            
    elif archivo_subido is not None:
        nombre = archivo_subido.name.lower()
        try:
            if nombre.endswith(".xlsx"):
                df = pd.read_excel(archivo_subido)
                productos = extraer_columna_productos(df)
            elif nombre.endswith(".csv"):
                df = pd.read_csv(archivo_subido)
                productos = extraer_columna_productos(df)
            else:  # Archivo de texto plano .txt
                contenido = archivo_subido.getvalue().decode("utf-8", errors="ignore")
                for linea in contenido.splitlines():
                    limpia = linea.strip()
                    if limpia and not limpia.startswith("#"):
                        productos.append(limpia.split(",")[0].strip())
        except Exception as e:
            st.error(f"Error al procesar el archivo: {e}")
                
    return productos[:limite_productos]

def extraer_numero_precio(texto_precio):
    """Extrae el valor numérico para ordenar de menor a mayor."""
    if not texto_precio:
        return 999999.0
    numeros = re.findall(r"\d+(?:\.\d+)?", str(texto_precio).replace(",", "."))
    if numeros:
        return float(numeros[0])
    return 999999.0

def consultar_ofertas_ordenadas(cliente, producto):
    prompt = f"""
    Actúa como un experto investigador de precios en Venezuela.
    Investiga en tiempo real precios y ofertas disponibles para el producto: "{producto}".
    
    Debes encontrar entre 2 y 3 opciones distintas en Venezuela:
    - Opciones en Mercado Libre Venezuela (mercadolibre.com.ve)
    - Opciones en tiendas oficiales o comercios aliados de Cashea (ej. Ivoo, SoyTechno, Damasco, Multimax).
    
    IMPORTANTE:
    - Debes ordenar las opciones estrictamente de la MÁS ECONÓMICA a la MÁS COSTOSA según su precio en USD.
    - Indica si la tienda acepta financiamiento con Cashea y el estimado de inicial y cuotas si aplica.
    
    Responde ÚNICAMENTE con un JSON válido con la siguiente estructura:
    {{
        "producto": "{producto}",
        "opciones": [
            {{
                "tienda_o_plataforma": "Mercado Libre / Ivoo / SoyTechno / etc.",
                "titulo_publicacion": "Título exacto de la publicación",
                "precio_usd": "$XXX",
                "imagen_url": "URL de la imagen del producto (si la consigues, sino vacío)",
                "link": "URL directa de la publicación o tienda",
                "es_cashea": true,
                "plan_cashea": "Inicial $XX + 3 cuotas de$XX (o vacío si no aplica)",
                "detalles": "Breve nota: garantía, estado, reputación del vendedor"
            }}
        ]
    }}
    """
    
    modelos = ["gemini-2.5-flash", "gemini-1.5-flash"]
    for modelo in modelos:
        for intento in range(2):
            try:
                resp = cliente.models.generate_content(
                    model=modelo,
                    contents=prompt,
                )
                texto = resp.text.strip()
                if texto.startswith("```json"):
                    texto = texto[7:]
                if texto.startswith("```"):
                    texto = texto[3:]
                if texto.endswith("```"):
                    texto = texto[:-3]
                data = json.loads(texto.strip())
                return data, None
            except Exception as e:
                err_str = str(e)
                if any(k in err_str for k in ["503", "429"]):
                    time.sleep(4)
                    continue
                break
    return None, "Servicio saturado temporalmente. Intenta nuevamente."

# Proceso al presionar el botón
if boton_iniciar:
    lista_productos = obtener_productos()
    
    if not lista_productos:
        st.warning("⚠️ No se encontraron productos. Revisa el archivo cargado o el enlace de Google Sheets.")
    else:
        st.info(f"📋 Procesando {len(lista_productos)} producto(s)...")
        
        api_key = st.secrets.get("GEMINI_API_KEY")
        if not api_key:
            st.error("❌ Falta configurar GEMINI_API_KEY en los Secrets.")
            st.stop()
            
        cliente = genai.Client(api_key=api_key)
        barra = st.progress(0)
        placeholder_img = "[https://placehold.co/400x300/f0f2f6/666666?text=Foto+del+Producto](https://placehold.co/400x300/f0f2f6/666666?text=Foto+del+Producto)"
        
        for i, prod in enumerate(lista_productos):
            with st.spinner(f"Buscando opciones de precio para: **{prod}**..."):
                datos, error = consultar_ofertas_ordenadas(cliente, prod)
                
                with st.container(border=True):
                    st.subheader(f"📦 {prod}")
                    
                    if error or not datos or "opciones" not in datos or not datos["opciones"]:
                        st.warning("No se pudieron cargar opciones para este producto en este intento.")
                    else:
                        opciones = datos["opciones"]
                        opciones.sort(key=lambda op: extraer_numero_precio(op.get("precio_usd", "")))
                        
                        num_cols = min(len(opciones), 3)
                        columnas = st.columns(num_cols)
                        
                        etiquetas = [
                            "🟢 Opción Más Económica",
                            "🟡 Opción Intermedia",
                            "🔵 Opción Alternativa / Más Alta"
                        ]
                        
                        for idx, op in enumerate(opciones[:3]):
                            col = columnas[idx]
                            with col:
                                etiqueta = etiquetas[idx] if idx < len(etiquetas) else f"Opción {idx + 1}"
                                st.markdown(f"**{etiqueta}**")
                                
                                img_url = op.get("imagen_url")
                                if not img_url or not str(img_url).startswith("http"):
                                    img_url = placeholder_img
                                    
                                st.image(img_url, use_container_width=True)
                                
                                precio = op.get("precio_usd", "Consultar")
                                st.markdown(f"### 💵 {precio}")
                                
                                tienda = op.get("tienda_o_plataforma", "Tienda online")
                                st.markdown(f"🏪 **{tienda}**")
                                
                                if op.get("es_cashea") and op.get("plan_cashea"):
                                    st.info(f"🟣 **Cashea:** {op.get('plan_cashea')}")
                                    
                                titulo = op.get("titulo_publicacion", prod)
                                st.caption(f"📝 {titulo}")
                                
                                if op.get("detalles"):
                                    st.caption(f"ℹ️ {op.get('detalles')}")
                                    
                                link_destino = op.get("link") or "[https://www.google.com](https://www.google.com)"
                                st.link_button("🔗 Ir a la publicación", link_destino, use_container_width=True)
                                
            barra.progress((i + 1) / len(lista_productos))
            time.sleep(3)
            
        st.balloons()
        st.success("🎉 ¡Comparativa de precios completada!")
