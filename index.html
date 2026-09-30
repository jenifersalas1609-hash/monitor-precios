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
    csv_url = os.environ.get('SHEET_CSV_URL')
    nombre_columna = os.environ.get('SHEET_COLUMN_NAME', 'Producto')
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
                    if not producto.startswith('#'):
                        productos.append(producto)
            
            if productos:
                print(f"Éxito: Se cargaron {len(productos)} productos desde el CSV.")
                return productos
            else:
                print(f"Advertencia: No se encontraron productos en la columna '{nombre_columna}'.")
                
        except Exception as e:
            print(f"Error al procesar el CSV: {e}. Cayendo de vuelta a {path}...")

    # Plan de respaldo: Leer desde el archivo local (productos.txt)
    print(f"Leyendo productos desde {path}...")
    try:
        with open(path, 'r', encoding='utf-8') as f:
            for linea in f:
                linea = linea.strip()
                if linea and not linea.startswith('#'):
                    productos.append(linea)
        print(f"Éxito: Se cargaron {len(productos)} productos desde {path}.")
    except FileNotFoundError:
        print(f"Error crítico: No se encontró el archivo {path}")
        
    return productos

# --- IMPORTANTE: NO BORRES EL RESTO DEL CÓDIGO QUE ESTÁ DEBAJO DE ESTO ---
