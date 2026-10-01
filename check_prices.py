import os
import sys
import html
import time
import urllib.request
import csv
import io
from datetime import datetime, timezone, timedelta
from google import genai
from google.genai import types

API_KEY = os.environ.get("GEMINI_API_KEY")
if not API_KEY:
    print("ERROR: falta la variable de entorno GEMINI_API_KEY", file=sys.stderr)
    sys.exit(1)

client = genai.Client(api_key=API_KEY)

def load_products(path="productos.txt"):
    # 1. Intentar cargar desde Google Sheets si existe la variable
    sheet_url = os.environ.get("SHEET_CSV_URL")
    if sheet_url:
        try:
            print("Detectada SHEET_CSV_URL. Descargando productos desde Google Sheets...")
            req = urllib.request.Request(sheet_url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=15) as resp:
                content = resp.read().decode("utf-8")
            reader = csv.reader(io.StringIO(content))
            productos = []
            for row in reader:
                if row and row[0].strip() and not row[0].strip().startswith("#"):
                    productos.append(row[0].strip())
            if productos:
                print(f"Se cargaron {len(productos)} productos desde Google Sheets.")
                return productos
        except Exception as e:
            print(f"Aviso: No se pudo leer Google Sheets ({e}). Buscando en archivo local...", file=sys.stderr)

    # 2. Si no hay Google Sheets o falló, usar productos.txt
    if not os.path.exists(path):
        return []
    productos = []
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            productos.append(line)
    return productos

def check_product(producto, reintentos=3):
    prompt = (
        f'Investiga el precio actual de "{producto}". '
        "Busca en al menos 2 o 3 tiendas/sitios distintos (por ejemplo Mercado Libre, "
        "tiendas online, distribuidores oficiales). "
        "Para cada tienda indica: nombre de la tienda, precio (en USD si es posible) "
        "y disponibilidad. "
        "Al final di claramente cual es la mejor opcion y por que (mejor precio, "
        "mejor disponibilidad, mas confiable, etc). "
        "Responde en español, de forma breve y clara, usando una lista corta."
    )
    
    # Intenta consultar a Gemini; si el servidor está saturado (503), espera y reintenta
    for intento in range(reintentos):
        try:
            response = client.models.generate_content(
                model="gemini-3.8-flash",
                contents=prompt,
            )
            texto = response.text or "(sin respuesta de Gemini)"
            fuentes = []
            return {"producto": producto, "texto": texto, "fuentes": fuentes, "error": None}
        except Exception as e:
            error_msg = str(e)
            if ("503" in error_msg or "UNAVAILABLE" in error_msg) and intento < reintentos - 1:
                print(f"Servidor ocupado para {producto}. Reintentando en 5 segundos... (intento {intento + 1}/{reintentos})")
                time.sleep(5)
                continue
            return {"producto": producto, "texto": "", "fuentes": [], "error": error_msg}

def format_texto(texto):
    texto = html.escape(texto)
    lineas = texto.split("\n")
    html_out = []
    en_lista = False
    for linea in lineas:
        l = linea.strip()
        if l.startswith("- ") or l.startswith("* "):
            if not en_lista:
                html_out.append("<ul>")
                en_lista = True
            html_out.append(f"<li>{l[2:]}</li>")
        else:
            if en_lista:
                html_out.append("</ul>")
                en_lista = False
            if l:
                html_out.append(f"<p>{l}</p>")
    if en_lista:
        html_out.append("</ul>")
    return "\n".join(html_out)

def build_html(resultados, fecha_str):
    tarjetas = []
    for r in resultados:
        nombre = html.escape(r["producto"])
        if r["error"]:
            cuerpo = f'<p class="error">No se pudo consultar este producto: {html.escape(r["error"])}</p>'
            fuentes_html = ""
        else:
            cuerpo = format_texto(r["texto"])
            if r["fuentes"]:
                items = "".join(
                    f'<li><a href="{html.escape(f["uri"])}" target="_blank" rel="noopener">{html.escape(f["title"] or f["uri"])}</a></li>'
                    for f in r["fuentes"]
                )
                fuentes_html = f'<details><summary>Fuentes</summary><ul>{items}</ul></details>'
            else:
                fuentes_html = ""
        tarjetas.append(
            f"""
            <div class="card">
              <h2>{nombre}</h2>
              {cuerpo}
              {fuentes_html}
            </div>
            """
        )
    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Monitor de Precios</title>
<style>
  body {{
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif;
    background: #0d1117;
    color: #e6edf3;
    margin: 0;
    padding: 24px;
  }}
  .container {{ max-width: 820px; margin: 0 auto; }}
  h1 {{ font-size: 1.6em; margin-bottom: 4px; }}
  .fecha {{ color: #8b949e; margin-bottom: 28px; }}
  .card {{
    background: #161b22;
    border: 1px solid #30363d;
    border-radius: 10px;
    padding: 18px 22px;
    margin-bottom: 18px;
  }}
  .card h2 {{ margin-top: 0; font-size: 1.15em; color: #58a6ff; }}
  .card p {{ margin: 6px 0; line-height: 1.5; }}
  .card ul {{ margin: 6px 0; padding-left: 20px; }}
  .card li {{ margin: 4px 0; line-height: 1.4; }}
  .error {{ color: #f85149; }}
  details {{ margin-top: 10px; font-size: 0.85em; color: #8b949e; }}
  details a {{ color: #8b949e; }}
  footer {{ color: #8b949e; font-size: 0.8em; margin-top: 30px; text-align: center; }}
</style>
</head>
<body>
  <div class="container">
    <h1>Monitor de Precios</h1>
    <div class="fecha">Actualizado: {fecha_str}</div>
    {"".join(tarjetas)}
    <footer>Generado automaticamente con Gemini. Actualiza tu Google Sheet o productos.txt para cambiar la lista.</footer>
  </div>
</body>
</html>
"""

def main():
    productos = load_products()
    if not productos:
        productos = ["Ejemplo: iPhone 13 128GB"]
    
    resultados = []
    for p in productos:
        resultados.append(check_product(p))
        time.sleep(3)  # Pausa de 3 segundos entre productos
        
    tz = timezone(timedelta(hours=-4))
    fecha_str = datetime.now(tz).strftime("%d/%m/%Y %H:%M") + " (Venezuela)"
    out = build_html(resultados, fecha_str)
    with open("index.html", "w", encoding="utf-8") as f:
        f.write(out)
    print(f"Reporte generado con {len(resultados)} producto(s).")

if __name__ == "__main__":
    main()
