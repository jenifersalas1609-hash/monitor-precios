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

st.set_page_config(
    page_title="RACOVE | Inteligencia Comercial y Sourcing",
    page_icon="🎯",
    layout="wide"
)

# Estilos CSS unificados para ambos módulos
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
    .badge-1688 { background-color: #ff6000; color: #ffffff; }
    .badge-alibaba { background-color: #ff6a00; color: #ffffff; }
    .badge-aliexpress { background-color: #e62e04; color: #ffffff; }

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

    /* Tarjetas estilo catálogo con foto panorámica superior */
    .card-item {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        display: flex;
        flex-direction: column;
        height: 100%;
        margin-bottom: 10px;
    }
    .card-item:hover { box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.09); }
    .card-img-top {
        width: 100%;
        height: 165px;
        object-fit: cover;
        border-radius: 12px 12px 0 0;
        background-color: #f1f5f9;
        display: block;
    }
    .card-content {
        padding: 12px;
        display: flex;
        flex-direction: column;
        flex-grow: 1;
    }
    .card-badge-econ {
        background-color: #dcfce7;
        color: #166534;
        font-size: 0.75rem;
        font-weight: 700;
        padding: 2px 8px;
        border-radius: 4px;
        display: inline-block;
        margin-bottom: 6px;
        width: fit-content;
    }
    .card-price {
        font-size: 1.45rem;
        font-weight: 800;
        color: #0f172a;
        margin: 4px 0;
    }
    .card-store { font-size: 0.95rem; font-weight: 700; color: #1e293b; }
    .card-reputation { font-size: 0.8rem; color: #64748b; margin-bottom: 4px; }
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
        font-size: 0.85rem;
        color: #475569;
        line-height: 1.25;
        height: 36px;
        overflow: hidden;
        margin-top: 4px;
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
# SELECTOR PRINCIPAL EN BARRA LATERAL
# -------------------------------------------------------------
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3135/3135715.png", width=60)
    st.title("🎯 RACOVE")
    st.caption("Radar Comercial Venezuela & Sourcing")
    modulo_activo = st.radio(
        "🌐 Selecciona el Módulo:",
        ["🇻🇪 RACOVE (Mercado Nacional y Rentabilidad)", "🇨🇳 Auditoría China (Fábricas y Compras)"]
    )
    st.divider()

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

def bytes_a_base64_img(bytes_img):
    try:
        encoded = base64.b64encode(bytes_img).decode("utf-8")
        return f"data:image/jpeg;base64,{encoded}"
    except Exception:
        return ""

def calcular_matriz_precios(costo_unitario, menor_precio_cashea=None, menor_precio_ml=None):
    # 1. Precio N: Costo + 60%
    precio_n = round(costo_unitario * 1.60, 2)
    # 2. Precio 4: Costo + 100%
    precio_4 = round(costo_unitario * 2.00, 2)
    # 3. Precio Divisa: Precio 4 + 35% SIEMPRE redondeado hacia arriba al entero
    precio_divisa_exacto = precio
