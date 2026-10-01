import streamlit as st
import time
import io
import csv
import re
from google import genai

# Configuración visual de la aplicación
st.set_page_config(
    page_title="Monitor de Precios - IA",
    page_icon="🛒",
    layout="wide"
)

st.title("🛒 Monitor de Precios en Vivo")
st.markdown("Sube tu archivo de productos para analizar tiendas, precios en USD y disponibilidad en tiempo real.")

# Barra lateral interactiva
with st.sidebar:
    st.header("⚙️ Configuración")
    archivo_subido = st.file_uploader(
        "Sube tu archivo de productos",
        type=["txt", "csv"],
        help="Sube un archivo .txt con un producto por línea o un archivo .csv"
    )
    
    st.markdown("---")
    boton_iniciar = st.button("🔍 Iniciar Monitoreo", type="primary", use_container_width=True)

def consultar_gemini(cliente, producto, reintentos=3):
    prompt = (
        f'Investiga el precio actual de "{producto}". '
        "Busca en al menos 2 o 3 tiendas/sitios distintos (por ejemplo Mercado Libre, "
        "tiendas online, distribuidores oficiales en Venezuela). "
        "Para cada tienda indica: nombre de la tienda, precio (en USD si es posible) "
        "y disponibilidad. "
        "Al final di claramente cuál es la mejor opción y por qué (mejor precio, "
        "mejor disponibilidad, más confiable, etc). "
        "Responde en español, de forma breve y estructurada."
    )
    
    for intento in range(reintentos):
        try:
            resp = cliente.models.generate_content(
                model="gemini-3.8-flash",
                contents=prompt,
            )
            return resp.text, None
        except Exception as e:
            error_str = str(e)
            if ("503" in error_str or "UNAVAILABLE" in error_str) and intento < reintentos - 1:
                time.sleep(5)
                continue
            return None, error_str

# Acción al presionar el botón
if boton_iniciar:
    if not archivo_subido:
        st.warning("⚠️ Por favor selecciona o arrastra primero un archivo (.txt o .csv) en la barra lateral.")
    else:
        contenido = archivo_subido.getvalue().decode("utf-8")
        lineas = contenido.splitlines()
        productos = []
        
        reader = csv.reader(lineas)
        for fila in reader:
            if fila and fila[0].strip() and not fila[0].strip().startswith("#"):
                productos.append(fila[0].strip())

        if not productos:
            st.error("❌ No se encontraron productos válidos en el archivo subido.")
        else:
            st.success(f"📋 Se cargaron {len(productos)} producto(s) exitosamente.")
            
            api_key = st.secrets.get("GEMINI_API_KEY")
            if not api_key:
                st.error("❌ Falta configurar la clave GEMINI_API_KEY en los Secrets de Streamlit.")
                st.stop()
                
            cliente = genai.Client(api_key=api_key)
            barra = st.progress(0)
            
            for i, prod in enumerate(productos):
                with st.spinner(f"Consultando precios para: **{prod}**..."):
                    texto, error = consultar_gemini(cliente, prod)
                    
                    with st.container(border=True):
                        st.subheader(f"📱 {prod}")
                        if error:
                            st.error(f"No se pudo consultar: {error}")
                        else:
                            st.markdown(texto)
                            
                barra.progress((i + 1) / len(productos))
                time.sleep(4)
                
            st.balloons()
            st.success("🎉 ¡Monitoreo completado con éxito!")
