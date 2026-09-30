# Contexto del proyecto: monitor-precios

Este archivo es para que cualquier IA (Gemini, Claude, etc.) que abra este repo entienda que es esto y que falta por hacer. Fecha de esta nota: 2026-09-30.

## Que es esto

Sistema automatizado que usa la API de Gemini (con Google Search grounding) para investigar precios de productos todos los dias y publicar un reporte HTML via GitHub Pages. Lo construyo Henry (webmaster541) junto con Claude (Anthropic).

## Arquitectura (archivos clave)

productos.txt: lista de productos a monitorear, un producto por linea. Lineas que empiezan con # son comentarios y se ignoran.

check_prices.py: script en Python que lee productos.txt, le pregunta a Gemini el precio de cada producto (usando grounding con Google Search), y genera index.html con el reporte final.

requirements.txt: dependencias de Python (google-genai).

.github/workflows/monitor-precios.yml: workflow de GitHub Actions. Corre check_prices.py todos los dias a las 12:00 UTC (8am hora Venezuela) via cron, y tambien se puede correr manualmente desde la pestana Actions (evento workflow_dispatch). Al terminar, hace commit y push de index.html de vuelta al repo si hubo cambios.

GitHub Pages esta activado (Settings, Pages, deploy desde rama main, carpeta raiz), asi que index.html se publica solo en: https://webmaster541.github.io/monitor-precios/

## Configuracion ya hecha (no repetir)

Secret GEMINI_API_KEY cargado en Settings, Secrets and variables, Actions, con la API key de Google AI Studio de Henry.

Workflow permissions puesto en "Read and write permissions" en Settings, Actions, General. Sin esto el workflow no puede hacer push del reporte (da error 403).

Modelo usado: gemini-3.8-flash. Ojo: gemini-2.5-flash fue descontinuado en septiembre 2026 y da error 404 si se usa.

## Limitacion conocida

La API key es de free tier. Con varios productos o varias corridas seguidas puede dar error 429 RESOURCE_EXHAUSTED. Para subir el limite hay que activar facturacion (billing) en el proyecto de Google Cloud/AI Studio asociado a esa key. Importante: la suscripcion paga de la app de Gemini (Gemini Advanced o Gemini Pro que ya tiene Henry) NO aumenta esta cuota, son sistemas de facturacion distintos.

## Pendiente, proximos pasos

Paso 1: Reemplazar los productos de ejemplo en productos.txt (iPhone 13, Redmi Note 13) por los productos reales que Henry quiere monitorear.

Paso 2 (opcional, todavia no implementado): conectar un Google Sheet como fuente de productos en vez de editar productos.txt a mano. Henry ya eligio el metodo: publicar el Sheet como CSV (Archivo, Compartir, Publicar en la web, elegir la hoja especifica, formato CSV, Publicar), guardar esa URL publicada como variable de repo llamada SHEET_CSV_URL (Settings, Secrets and variables, Actions, pestana Variables), y modificar check_prices.py para que si existe esa variable de entorno descargue el CSV (con la libreria requests, agregarla a requirements.txt) y use la columna de nombres de producto como lista, cayendo de vuelta a productos.txt si SHEET_CSV_URL no esta configurada. Falta que Henry mande el link publicado y el nombre de la columna para terminar esto.

Paso 3: si se agregan muchos productos, considerar activar facturacion paga en el proyecto de Google AI Studio para evitar los errores 429.

## Como probar cambios

Ir a la pestana Actions del repo, click en "Monitor de precios", boton "Run workflow" para correrlo manualmente sin esperar al cron diario. Revisar el resultado en https://webmaster541.github.io/monitor-precios/ despues de uno o dos minutos, porque el build de GitHub Pages tarda un poco despues del push.
