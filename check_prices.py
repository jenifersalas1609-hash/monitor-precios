import os
import sys
import html
import csv
import requests
from io import StringIO
from datetime import datetime, timezone, timedelta
from google import genai
from google.genai import types

API_KEY = os.environ.get("GEMINI_API_KEY")
if not API_KEY:
    print("ERROR: falta la variable de entorno GEMINI_API_KEY", file=sys.stderr)
    sys.exit(1)

client = genai.Client(api_key=API_KEY)

def load_products(path="productos.txt"):
    # Intenta obtener la URL del CSV desde las variables de entorno
    csv_url = os.environ.get("SHEET_CSV_URL")
    nombre_columna = os.environ.get("SHEET_COLUMN_NAME", "Producto")
    productos = []

    if csv_url:
        print("Detectada SHEET_CSV_URL. Descargando productos desde Google Sheets...")
        try:
            response = requests.get(csv_url)
            response.raise_for_status()
            
            csv_data = StringIO(response.text)
            reader = csv.DictReader(csv_data)
            
            for fila in reader:
                if nombre_columna in fila and fila[nombre_columna].strip():
                    producto = fila[nombre_columna].strip()
                    if not producto.startswith("#"):
                        productos.append(producto)
            
            if productos:
                print(f"Éxito: Se cargaron {len(productos)} productos desde el CSV de Google Sheets.")
                return productos
            else:
                print(f"Advertencia: No se encontraron productos en la columna '{nombre_columna}'.")
        except Exception as e:
            print(f"Error al procesar el CSV: {e}. Cayendo de vuelta al archivo local...")

    # Plan de respaldo: Leer desde el archivo local (productos.txt)
    print(f"Leyendo productos desde {path}...")
    if not os.path.exists(path):
        return []

    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            productos.append(line)
            
    print(f"Éxito: Se cargaron {len(productos)} productos desde {path}.")
    return productos

def check_product(producto):
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
    try:
        grounding_tool = types.Tool(google_search=types.GoogleSearch())
        config = types.GenerateContentConfig(tools=[grounding_tool])
        response = client.models.generate_content(
            model="gemini-3.8-flash",
            contents=prompt,
            config=config,
        )
        texto = response.text or "(sin respuesta de Gemini)"
        fuentes = []
        try:
            chunks = response.candidates[0].grounding_metadata.grounding_chunks or []
            for c in chunks:
                if getattr(c, "web", None):
                    fuentes.append({"title": c.web.title, "uri": c.web.uri})
        except Exception:
            pass
        return {"producto": producto, "texto": texto, "fuentes": fuentes, "error": None}
    except Exception as e:
        return {"producto": producto, "texto": "", "fuentes": [], "error": str(e)}

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
    <footer>Generado automaticamente con Gemini. Edita productos.txt o conecta Google Sheets para cambiar la lista.</footer>
  </div>
</body>
</html>
"""

def main():
    productos = load_products()
    if not productos:
        productos = ["Ejemplo: iPhone 13 128GB"]
    resultados = [check_product(p) for p in productos]
    tz = timezone(timedelta(hours=-4))
    fecha_str = datetime.now(tz).strftime("%d/%m/%Y %H:%M") + " (Venezuela)"
    out = build_html(resultados, fecha_str)
    with open("index.html", "w", encoding="utf-8") as f:
        f.write(out)
    print(f"Reporte generado con {len(resultados)} producto(s).")

if __name__ == "__main__":
    main()
