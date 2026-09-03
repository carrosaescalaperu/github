"""
================================================================================
 SCRAPER — Catálogo mayorista modelcarswholesale.com  →  Catálogo web (HTML)
================================================================================

CONTEXTO Y ADVERTENCIA IMPORTANTE (leer antes de usar):
--------------------------------------------------------------------------------
Al intentar inspeccionar la web para construir este script, la petición fue
bloqueada por un sistema anti-bot (probablemente Cloudflare u otro WAF). Esto
significa dos cosas:

  1. NO pude confirmar los selectores CSS reales (nombres de clases/IDs) de
     las tarjetas de producto, precio, SKU, etc. Los que ves abajo son los
     patrones MÁS COMUNES en catálogos tipo PrestaShop/OpenCart (el mismo
     motor que suelen usar sitios "hermanos" de distribución de scale models).
     Están marcados con "# AJUSTAR" — debes confirmarlos abriendo una página
     de producto real con F12 → pestaña "Elements" / "Inspeccionar".

  2. Un scraper con `requests` simple puede recibir error 403 (bloqueado).
     Por eso este archivo incluye DOS motores:
        - Motor A: requests + BeautifulSoup (rápido, pruébalo primero)
        - Motor B: Playwright (navegador real headless, esquiva bot-detection
          y es el único que te permite iniciar sesión si necesitas ver
          precios de distribuidor)

  3. Antes de scrapear en volumen, revisa https://www.modelcarswholesale.com/robots.txt
     y los Términos de Servicio del sitio. Como eres distribuidor con cuenta
     activa, lo más seguro (técnica y legalmente) es preguntarles si ofrecen
     un feed/export de catálogo (CSV, XML, API) — muchos mayoristas de este
     rubro sí lo tienen y te ahorra todo este trabajo.

--------------------------------------------------------------------------------
INSTALACIÓN (Windows, cmd o PowerShell):
--------------------------------------------------------------------------------
    python -m venv venv
    venv\\Scripts\\activate
    pip install requests beautifulsoup4 pandas openpyxl lxml
    pip install playwright
    playwright install chromium

--------------------------------------------------------------------------------
USO:
--------------------------------------------------------------------------------
    # Motor A (rápido, pruébalo primero)
    python scraper_modelcarswholesale.py --engine requests

    # Motor B (si el A da error 403 / bloqueo)
    python scraper_modelcarswholesale.py --engine playwright

    # Con login de distribuidor (variables de entorno, NUNCA hardcodeadas):
    #   set MCW_USER=tu_usuario
    #   set MCW_PASS=tu_clave
    python scraper_modelcarswholesale.py --engine playwright --login
================================================================================
"""

import os
import re
import sys
import time
import json
import random
import logging
import argparse
from html import escape as _html_escape
from datetime import datetime
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
import pandas as pd

# ==============================================================================
# 1. CONFIGURACIÓN GENERAL
# ==============================================================================

BASE_URL = "https://www.modelcarswholesale.com/search?keyword=&scale=all&type=available&country=all"

# AJUSTAR: rutas reales de las secciones que quieres scrapear.
# Entra al menú del sitio, haz click en "Available" / "All Available" /
# "Special Offers" y copia la URL exacta que te da el navegador.

CATALOG_SECTIONS = {
    "available": "/search?keyword=&scale=all&type=available&country=all&page=101",
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "es-ES,es;q=0.9,en-US;q=0.8,en;q=0.7",
    "Referer": BASE_URL,
    "Connection": "keep-alive",
}

# Pausa aleatoria entre requests para no saturar el servidor ni parecer un bot
REQUEST_DELAY_RANGE = (1.8, 4.0)  # segundos

# Límite de seguridad para no entrar en un loop infinito si la paginación falla.
# Con catálogos de 1000+ productos, sube este número si ves que corta antes
# de llegar a la última página real (revisa cuántas páginas tiene cada sección).
MAX_PAGES_PER_SECTION = 300

OUTPUT_FILE = f"catalogo_modelcarswholesale_{datetime.now():%Y%m%d_%H%M}.html"

# Cada cuántos productos se guarda un checkpoint (CSV) por si el proceso se
# corta a mitad de camino (caída de internet, bloqueo del sitio, etc.)
CHECKPOINT_CADA = 100
CHECKPOINT_FILE = "checkpoint_productos.csv"


def guardar_checkpoint(productos: list[dict]) -> None:
    """Guarda el progreso actual en CSV. Se sobreescribe cada vez (no acumula duplicados)."""
    if not productos:
        return
    pd.DataFrame(productos).to_csv(CHECKPOINT_FILE, index=False, encoding="utf-8-sig")
    logger.info(f"Checkpoint guardado: {len(productos)} productos en '{CHECKPOINT_FILE}'")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger("mcw_scraper")


# ==============================================================================
# 2. (OPCIONAL) LOGIN PARA TARIFAS DE DISTRIBUIDOR — Motor A (requests.Session)
# ==============================================================================
def login_con_requests(session: requests.Session) -> bool:
    """
    Bloque preparado para iniciar sesión vía requests.Session, por si el sitio
    requiere login para ver precios/stock de mayorista.

    PASOS PARA CONFIGURARLO (hazlo una sola vez):
      1. Abre el sitio en Chrome, F12 → pestaña "Network", marca "Preserve log".
      2. Haz login manualmente con tu usuario y clave.
      3. En la lista de requests busca el POST hacia la URL de login
         (normalmente algo como /login o /connexion).
      4. Haz clic en esa request → pestaña "Payload" / "Form Data" y copia
         los nombres EXACTOS de los campos (ej. "email", "password", "_token").
      5. Reemplaza los valores de LOGIN_URL y las llaves del diccionario
         `payload` más abajo con esos nombres reales.

    Las credenciales se leen de variables de entorno (nunca las escribas
    directamente en el código ni las subas a un repositorio público).
    """
    LOGIN_URL = f"{BASE_URL}/login"  # <-- AJUSTAR a la URL real de login

    usuario = os.getenv("MCW_USER")
    clave = os.getenv("MCW_PASS")

    if not usuario or not clave:
        logger.info(
            "No se encontraron credenciales en variables de entorno "
            "(MCW_USER / MCW_PASS). Se continúa en modo público."
        )
        return False

    try:
        # Muchos sitios PrestaShop/OpenCart incluyen un token CSRF oculto en
        # el formulario de login. Si es el caso, hay que leerlo primero:
        pagina_login = session.get(LOGIN_URL, headers=HEADERS, timeout=15)
        soup = BeautifulSoup(pagina_login.text, "lxml")

        token_input = soup.select_one('input[name="token"]')  # <-- AJUSTAR si aplica
        token = token_input["value"] if token_input else None

        payload = {
            "email": usuario,      # <-- AJUSTAR nombre real del campo
            "password": clave,     # <-- AJUSTAR nombre real del campo
        }
        if token:
            payload["token"] = token  # <-- AJUSTAR nombre real del campo CSRF

        resp = session.post(LOGIN_URL, data=payload, headers=HEADERS, timeout=15)
        resp.raise_for_status()

        # AJUSTAR: texto que solo aparece cuando el login fue exitoso
        # (ej. "Cerrar sesión", "Mi cuenta", el nombre del usuario, etc.)
        indicadores_login_ok = ["cerrar sesión", "mi cuenta", "logout"]
        texto_resp = resp.text.lower()

        if any(ind in texto_resp for ind in indicadores_login_ok):
            logger.info("Login exitoso (Motor A - requests).")
            return True

        logger.warning(
            "No se pudo confirmar el login con certeza. Revisa manualmente "
            "los selectores/indicadores en login_con_requests()."
        )
        return False

    except requests.RequestException as e:
        logger.error(f"Error durante el login (requests): {e}")
        return False


# ==============================================================================
# 3. MOTOR A — requests + BeautifulSoup
# ==============================================================================
def extraer_productos_de_pagina(html: str, url_pagina: str) -> list[dict]:
    """
    Recorre el HTML de una página de listado y extrae los datos de cada
    producto disponible. AJUSTAR los selectores según el HTML real del sitio.
    """
    productos = []
    soup = BeautifulSoup(html, "lxml")

    # AJUSTAR: selector del contenedor de cada tarjeta de producto.
    # Patrones típicos: "div.row product", "li.product", "div.product-miniature"
    tarjetas = soup.select("div.row.product")  # <-- AJUSTAR

    if not tarjetas:
        logger.warning(
            f"No se encontraron tarjetas de producto en {url_pagina}. "
            "Verifica el selector 'tarjetas' — puede que el sitio use "
            "JavaScript para renderizar el catálogo (usa --engine playwright)."
        )
        return productos

    for tarjeta in tarjetas:
        try:
            # --- Filtro de stock: saltar productos marcados como agotados ---
            # AJUSTAR: clase/texto que el sitio usa para marcar "sin stock"
            fuera_de_stock = tarjeta.select_one(".out-of-stock, .no-stock, .agotado")
            if fuera_de_stock:
                continue

            # --- Nombre / descripción ---
            nombre_tag = tarjeta.select_one(".row-two.hidden-xs")  # <-- AJUSTAR
            nombre = nombre_tag.get_text(strip=True) if nombre_tag else ""

            # --- URL del producto (por si luego quieres entrar al detalle) ---
            url_producto = nombre_tag["href"] if nombre_tag and nombre_tag.has_attr("href") else None
            if url_producto:
                url_producto = urljoin(BASE_URL, url_producto)

            # --- SKU / referencia ---
            sku_tag = tarjeta.select_one("row-one.hidden-xs")  # <-- AJUSTAR
            sku = sku_tag.get_text(strip=True) if sku_tag else ""

            # --- Precio ---
            precio_tag = tarjeta.select_one(".price, .product-price")  # <-- AJUSTAR
            precio_texto = precio_tag.get_text(strip=True) if precio_tag else ""
            precio = limpiar_precio(precio_texto)

            # --- Imagen en alta resolución ---
            img_tag = tarjeta.select_one(".col-sm-4.col-xs-12 img")
            imagen_url = None
            if img_tag:
                # Muchos catálogos usan lazy-loading: la URL real puede estar
                # en data-src / data-full-size-image en vez de src.
                imagen_url = (
                    img_tag.get("data-full-size-image")
                    or img_tag.get("data-src")
                    or img_tag.get("src")
                )
                if imagen_url:
                    imagen_url = urljoin(BASE_URL, imagen_url)

            # --- Marca / Fabricante y Escala ---
            # Muchos catálogos de este rubro no separan marca/escala en campos
            # propios, sino que van dentro del nombre del producto
            # (ej. "SOLIDO 1:18 Nissan Skyline R34 Z-Tune").
            # Por eso se intentan extraer con expresiones regulares del nombre.
           
            caja_info = tarjeta.select_one(".scale")
            texto_info = caja_info.get_text(strip=True) if caja_info else ""
            
            # Extraemos los números para la escala (ej: "1:43")
            escala = extraer_escala(texto_info)
            
            # Para la marca, borramos los números de la escala y dejamos solo el texto (ej: "LOOKSMART")
            marca = re.sub(r"1[:/\-]\s?\d{2,3}", "", texto_info).strip() 
            
            # --- Trademark ---
            # Reemplaza ".trademark-info" con el nombre real que te dé el inspector
            caja_trademark = tarjeta.select_one(".trademark") 
            trademark = caja_trademark.get_text(strip=True) if caja_trademark else ""

            # --- Código ---
            caja_codigo = tarjeta.select_one('.row-three div[style*="width: 70%"] b')
            sku = caja_codigo.get_text(strip=True) if caja_codigo else ""

            productos.append(
                {
                    "SKU": sku,
                    "Marca": marca,
                    "Escala": escala,
                    "Nombre": nombre,
                    "Trademark": trademark,
                    "Precio": precio,
                    "Imagen_URL": imagen_url,
                    "URL_Producto": url_producto,
                }
            )

        except Exception as e:
            # Nunca dejamos que un solo producto mal formado tumbe todo el script
            logger.error(f"Error procesando una tarjeta de producto en {url_pagina}: {e}")
            continue

    return productos


def limpiar_precio(texto: str) -> str:
    """Extrae solo el valor numérico de un string de precio (ej. '€ 45,90' -> '45.90')."""
    if not texto:
        return ""
    match = re.search(r"[\d.,]+", texto)
    if not match:
        return ""
    numero = match.group(0).replace(".", "").replace(",", ".")
    try:
        return f"{float(numero):.2f}"
    except ValueError:
        return texto


def extraer_escala(nombre: str) -> str:
    """Busca patrones de escala (ej. 1:18, 1/43, 1:8, 1:2) dentro del texto."""
    # Cambiamos a (\d{1,3}) para que acepte desde 1 hasta 3 dígitos
    match = re.search(r"1[:/\-]\s?(\d{1,3})", nombre)
    return f"1:{match.group(1)}" if match else ""


def extraer_marca(nombre: str) -> str:
    """
    Intenta detectar la marca/fabricante como la primera palabra del nombre.
    AJUSTAR: si el sitio tiene un campo de marca separado (ej. un link a
    "/brand/solido"), es mucho más confiable extraerlo de ahí en vez de
    adivinar desde el texto.
    """
    if not nombre:
        return ""
    return nombre.split()[0]


def encontrar_url_pagina_siguiente(html: str, url_actual: str) -> str | None:
    """
    Busca el link de la página siguiente basándose en el símbolo de una sola flecha ('>').
    """
    soup = BeautifulSoup(html, "lxml")
    
    # Buscamos cualquier etiqueta 'a' cuyo texto contenga '>' pero NO '>>'
    siguiente = soup.find("a", string=lambda t: t and ">" in t and ">>" not in t)
    
    if siguiente and siguiente.has_attr("href"):
        return urljoin(url_actual, siguiente["href"])
        
    return None

def scrapear_con_requests(usar_login: bool = False) -> pd.DataFrame:
    """Motor A: recorre todas las secciones y páginas usando requests + BeautifulSoup."""
    session = requests.Session()
    session.headers.update(HEADERS)

    if usar_login:
        login_con_requests(session)

    todos_los_productos = []

    for nombre_seccion, ruta in CATALOG_SECTIONS.items():
        url = urljoin(BASE_URL, ruta)
        pagina_num = 1

        while url and pagina_num <= MAX_PAGES_PER_SECTION:
            logger.info(f"[{nombre_seccion}] Descargando página {pagina_num}: {url}")

            try:
                resp = session.get(url, timeout=20)
                resp.raise_for_status()
            except requests.RequestException as e:
                logger.error(f"Error al descargar {url}: {e}")
                break

            if resp.status_code == 403:
                logger.error(
                    "Bloqueo 403 detectado (bot-detection). "
                    "Vuelve a intentar con --engine playwright."
                )
                break

            productos_pagina = extraer_productos_de_pagina(resp.text, url)
            todos_los_productos.extend(productos_pagina)
            logger.info(f"  -> {len(productos_pagina)} productos extraídos en esta página.")
            logger.info(f"  -> Total acumulado: {len(todos_los_productos)} productos.")

            if len(todos_los_productos) % CHECKPOINT_CADA < len(productos_pagina):
                guardar_checkpoint(todos_los_productos)

            url_siguiente = encontrar_url_pagina_siguiente(resp.text, url)
            if not url_siguiente or url_siguiente == url:
                break
            url = url_siguiente
            pagina_num += 1

            # Pausa para respetar el servidor
            time.sleep(random.uniform(*REQUEST_DELAY_RANGE))

    return pd.DataFrame(todos_los_productos)


# ==============================================================================
# 4. MOTOR B — Playwright (recomendado si hay bot-detection, y necesario
#    si vas a iniciar sesión para ver precios de distribuidor)
# ==============================================================================
def scrapear_con_playwright(usar_login: bool = False) -> pd.DataFrame:
    """
    Motor B: usa un navegador real (headless) para esquivar bot-detection y,
    opcionalmente, iniciar sesión antes de scrapear.
    Requiere: pip install playwright  &&  playwright install chromium
    """
    from playwright.sync_api import sync_playwright

    todos_los_productos = []

    
    with sync_playwright() as p:
        # Conectar al Chrome "títere" que ya abriste en el puerto 9222
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        
        # Tomar el control de esa ventana real
        context = browser.contexts[0]
        
        # Abrir una pestaña nueva en tu Chrome
        page = context.new_page()


        if usar_login:
            usuario = os.getenv("MCW_USER")
            clave = os.getenv("MCW_PASS")
            if usuario and clave:
                try:
                    page.goto(f"{BASE_URL}/login", timeout=30000)  # <-- AJUSTAR
                    # AJUSTAR selectores de los campos de login reales:
                    page.fill('input[name="email"]', usuario)       # <-- AJUSTAR
                    page.fill('input[name="password"]', clave)      # <-- AJUSTAR
                    page.click('button[type="submit"]')             # <-- AJUSTAR
                    page.wait_for_load_state("domcontentloaded")
                    logger.info("Login realizado (Motor B - Playwright).")
                except Exception as e:
                    logger.error(f"No se pudo iniciar sesión con Playwright: {e}")
            else:
                logger.info(
                    "No se encontraron credenciales en variables de entorno "
                    "(MCW_USER / MCW_PASS). Se continúa en modo público."
                )

        for nombre_seccion, ruta in CATALOG_SECTIONS.items():
            url = urljoin(BASE_URL, ruta)
            pagina_num = 1

            while url and pagina_num <= MAX_PAGES_PER_SECTION:
                logger.info(f"[{nombre_seccion}] Cargando página {pagina_num}: {url}")

                try:
                    page.goto(url, timeout=10000)
                    page.wait_for_load_state("domcontentloaded")
                    page.wait_for_timeout(10000)
                except Exception as e:
                    logger.error(f"Error al cargar {url}: {e}")
                    break

                html = page.content()
                productos_pagina = extraer_productos_de_pagina(html, url)
                todos_los_productos.extend(productos_pagina)
                logger.info(f"  -> {len(productos_pagina)} productos extraídos en esta página.")

                url_siguiente = encontrar_url_pagina_siguiente(html, url)
                if not url_siguiente or url_siguiente == url:
                    break
                url = url_siguiente
                pagina_num += 1

                time.sleep(random.uniform(*REQUEST_DELAY_RANGE))

        browser.close()

    return pd.DataFrame(todos_los_productos)


# ==============================================================================
# 5. GENERACIÓN DE CATÁLOGO WEB (HTML + CSS + JS EN UN SOLO ARCHIVO)
# ==============================================================================

# --- AJUSTAR AQUÍ el nombre de la tienda y la ruta/URL de tu logo ---
NOMBRE_TIENDA = "Carros Escala Perú - Catálogo a pedido"
RUTA_LOGO = "data:image/jpeg;base64,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"  # logo incrustado (Carros Escala Peru)

# Link de tu Google Sheet publicado como CSV (ver instrucciones más abajo).
# Si lo dejas vacío "", el catálogo usa los datos del Excel/scraping tal cual
# (modo estático, como hasta ahora). Si lo llenas, el catálogo carga los
# precios/datos en vivo desde el Sheet cada vez que alguien lo abre.
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1_xL3k8Kv4vJRbpfSqXE_XBYwk1rphiFE/export?format=csv&gid=1914991661"

# Plantilla base de la página. Se completa con .replace() (no f-string) para
# no tener que escapar cada { } del CSS/JS. Los tokens __TITULO__,
# __LOGO_SRC__ y __PRODUCTS_JSON__ se reemplazan más abajo.
_PLANTILLA_HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>__TITULO__</title>
<style>
  :root{
    --bg:#f5f4f1;
    --surface:#ffffff;
    --ink:#16181d;
    --ink-muted:#6f737a;
    --line:#e5e3dd;
    --accent:#ff4d2e;
    --accent-ink:#ffffff;
    --header-bg:#14161b;
    --header-ink:#f5f4f1;
    --radius:10px;
  }
  *{box-sizing:border-box;}
  html,body{margin:0;padding:0;}
  body{
    background:var(--bg);
    color:var(--ink);
    font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
    -webkit-font-smoothing:antialiased;
  }
  a{color:inherit;text-decoration:none;}

  /* ---------- Header ---------- */
  .site-header{ background:var(--header-bg); color:var(--header-ink); }
  .header-inner{
    max-width:1280px; margin:0 auto; padding:18px 24px;
    display:flex; align-items:center; gap:16px;
  }
  .logo-wrap{
    height:48px; width:auto; max-width:190px; flex:0 0 auto; border-radius:8px;
    overflow:hidden; background:var(--surface); padding:4px 10px;
    display:flex; align-items:center; justify-content:center;
  }
  .logo-wrap img{ width:100%; height:100%; object-fit:contain; }
  .store-title{ margin:0; font-size:19px; font-weight:700; letter-spacing:0.2px; }

  /* ---------- Toolbar ---------- */
  .toolbar{
    position:sticky; top:0; z-index:5;
    background:var(--bg); border-bottom:1px solid var(--line);
  }
  .toolbar-inner{
    max-width:1280px; margin:0 auto; padding:14px 24px;
    display:flex; flex-wrap:wrap; align-items:center; gap:10px;
  }
  #searchInput{
    flex:1 1 240px; min-width:0; padding:10px 14px;
    border:1px solid var(--line); border-radius:8px;
    background:var(--surface); font-size:14px; color:var(--ink);
  }
  #scaleSelect{
    flex:0 0 auto; padding:10px 12px;
    border:1px solid var(--line); border-radius:8px;
    background:var(--surface); font-size:14px; color:var(--ink);
  }
  #searchInput:focus, #scaleSelect:focus{
    outline:2px solid var(--accent); outline-offset:1px;
  }
  .counter{
    flex:0 0 auto; margin-left:auto; font-size:13px;
    color:var(--ink-muted); font-variant-numeric:tabular-nums; white-space:nowrap;
  }

  /* ---------- Grid ---------- */
  main{ max-width:1280px; margin:0 auto; padding:22px 24px 60px; }
  .grid{
    display:grid;
    grid-template-columns:repeat(auto-fill, minmax(210px, 1fr));
    gap:16px;
  }
  .card{
    display:flex; flex-direction:column;
    background:var(--surface); border:1px solid var(--line);
    border-radius:var(--radius); overflow:hidden;
    transition:transform .15s ease, box-shadow .15s ease;
  }
  .card:hover{ transform:translateY(-3px); box-shadow:0 8px 20px rgba(20,22,27,0.08); }
  .card-img{
    aspect-ratio:4/3; background:#f0efeb;
    display:flex; align-items:center; justify-content:center; overflow:hidden;
  }
  .card-img img{ width:100%; height:100%; object-fit:contain; padding:10px; }
  .card-img .no-img{ font-size:12px; color:var(--ink-muted); }
  .card-body{ padding:12px 14px 14px; display:flex; flex-direction:column; gap:6px; flex:1; }
  .badges{ display:flex; gap:6px; flex-wrap:wrap; }
  .badge{
    font-size:10.5px; font-weight:700; letter-spacing:0.5px; text-transform:uppercase;
    padding:3px 7px; border-radius:4px;
    font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;
  }
  .badge-escala{ background:var(--accent); color:var(--accent-ink); }
  .badge-marca{ background:var(--header-bg); color:var(--header-ink); }
  .card-title{
    margin:2px 0 0; font-size:14px; font-weight:600; line-height:1.35;
    display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden;
  }
  .card-sku{ font-size:11px; color:var(--ink-muted); }
  .card-footer{ margin-top:auto; display:flex; align-items:center; justify-content:space-between; padding-top:8px; }
  .card-price{ font-size:15px; font-weight:700; }

  /* ---------- Estado vacío ---------- */
  .empty-state{ text-align:center; color:var(--ink-muted); padding:60px 20px; font-size:14px; }

  /* ---------- Paginación ---------- */
  .pagination{
    display:flex; align-items:center; justify-content:center; gap:16px;
    margin-top:28px; padding-top:20px; border-top:1px solid var(--line);
  }
  .page-btn{
    padding:9px 16px; border:1px solid var(--line); border-radius:8px;
    background:var(--surface); color:var(--ink); font-size:13.5px; cursor:pointer;
  }
  .page-btn:hover:not(:disabled){ border-color:var(--accent); color:var(--accent); }
  .page-btn:disabled{ opacity:0.4; cursor:not-allowed; }
  .page-info{ font-size:13.5px; color:var(--ink-muted); white-space:nowrap; }

  @media (prefers-reduced-motion: reduce){ .card{ transition:none; } }

  @media (max-width: 520px){
    .header-inner{ padding:14px 16px; }
    .toolbar-inner{ padding:12px 16px; }
    main{ padding:16px 16px 40px; }
    .counter{ width:100%; margin-left:0; order:3; }
  }
</style>
</head>
<body>

  <header class="site-header">
    <div class="header-inner">
      <div class="logo-wrap">
        <img src="__LOGO_SRC__" alt="Logo" onerror="this.style.display='none'">
      </div>
      <h1 class="store-title">__TITULO__</h1>
    </div>
  </header>

  <div class="toolbar">
    <div class="toolbar-inner">
      <input type="search" id="searchInput" placeholder="Buscar por nombre, marca o modelo...">
      <select id="scaleSelect">
        <option value="">Todas las escalas</option>
      </select>
      <select id="sortSelect">
        <option value="">Sin ordenar</option>
        <option value="precio-asc">Precio: menor a mayor</option>
        <option value="precio-desc">Precio: mayor a menor</option>
        <option value="nombre-asc">Nombre: A-Z</option>
      </select>
      <span id="counter" class="counter">0 productos</span>
    </div>
  </div>

  <main>
    <div id="grid" class="grid"></div>
    <p id="emptyState" class="empty-state" hidden>No se encontraron productos con esos filtros.</p>
    <div id="pagination" class="pagination" hidden>
      <button id="prevPage" class="page-btn" type="button">&larr; Anterior</button>
      <span id="pageInfo" class="page-info"></span>
      <button id="nextPage" class="page-btn" type="button">Siguiente &rarr;</button>
    </div>
  </main>

<script>
  let PRODUCTOS = __PRODUCTS_JSON__;
  const PRODUCTOS_INICIALES = PRODUCTOS.slice();
  const SHEET_CSV_URL = "__SHEET_CSV_URL__";

  const grid = document.getElementById('grid');
  const searchInput = document.getElementById('searchInput');
  const scaleSelect = document.getElementById('scaleSelect');
  const sortSelect = document.getElementById('sortSelect');
  const counter = document.getElementById('counter');
  const emptyState = document.getElementById('emptyState');
  const pagination = document.getElementById('pagination');
  const pageInfo = document.getElementById('pageInfo');
  const prevPageBtn = document.getElementById('prevPage');
  const nextPageBtn = document.getElementById('nextPage');

  const PRODUCTOS_POR_PAGINA = 30;
  let paginaActual = 1;

  function normalizar(texto){
    return String(texto || '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLowerCase();
  }

  function limpiarPrecio(valor){
    if(!valor) return '';
    // Quita cualquier símbolo de moneda o texto (S/, $, €, espacios...) que ya
    // venga en el dato de origen, dejando solo el número.
    return String(valor).replace(/[^\d.,]/g, '').trim();
  }

  function numeroEscala(escala){
    const m = String(escala || '').match(/(\d+)/);
    return m ? parseInt(m[1], 10) : Number.MAX_SAFE_INTEGER;
  }

  function poblarEscalas(){
    scaleSelect.querySelectorAll('option:not([value=""])').forEach(o => o.remove());

    const escalas = [...new Set(
      PRODUCTOS.map(p => (p.Escala || '').trim()).filter(Boolean)
    )].sort((a, b) => numeroEscala(a) - numeroEscala(b));

    for(const escala of escalas){
      const opt = document.createElement('option');
      opt.value = escala;
      opt.textContent = escala;
      scaleSelect.appendChild(opt);
    }
  }

  function crearTarjeta(p){
    const el = document.createElement(p.URL_Producto ? 'a' : 'div');
    el.className = 'card';
    if(p.URL_Producto){
      el.href = p.URL_Producto;
      el.target = '_blank';
      el.rel = 'noopener noreferrer';
    }

    const imgWrap = document.createElement('div');
    imgWrap.className = 'card-img';
    if(p.Imagen_URL){
      const img = document.createElement('img');
      img.src = p.Imagen_URL;
      img.alt = p.Nombre || '';
      img.loading = 'lazy';
      img.onerror = function(){ imgWrap.innerHTML = '<span class="no-img">Sin imagen</span>'; };
      imgWrap.appendChild(img);
    } else {
      imgWrap.innerHTML = '<span class="no-img">Sin imagen</span>';
    }

    const body = document.createElement('div');
    body.className = 'card-body';

    const badges = document.createElement('div');
    badges.className = 'badges';
    if(p.Escala){
      const b = document.createElement('span');
      b.className = 'badge badge-escala';
      b.textContent = p.Escala;
      badges.appendChild(b);
    }
    const marcaTexto = p.Trademark || p.Marca;
    if(marcaTexto){
      const b = document.createElement('span');
      b.className = 'badge badge-marca';
      b.textContent = marcaTexto;
      badges.appendChild(b);
    }

    const titulo = document.createElement('p');
    titulo.className = 'card-title';
    titulo.textContent = p.Nombre || 'Sin nombre';

    const sku = document.createElement('p');
    sku.className = 'card-sku';
    sku.textContent = p.SKU ? `Ref. ${p.SKU}` : '';

    const footer = document.createElement('div');
    footer.className = 'card-footer';
    const precio = document.createElement('span');
    precio.className = 'card-price';
    precio.textContent = p.Precio ? `S/ ${limpiarPrecio(p.Precio)}` : '';
    footer.appendChild(precio);

    body.appendChild(badges);
    body.appendChild(titulo);
    if(sku.textContent) body.appendChild(sku);
    body.appendChild(footer);

    el.appendChild(imgWrap);
    el.appendChild(body);
    return el;
  }

  function coincideBusqueda(p, termino){
    if(!termino) return true;
    const campos = [p.Nombre, p.Marca, p.Trademark, p.SKU];
    return campos.some(campo => normalizar(campo).includes(termino));
  }

  function numeroPrecio(precio){
    const n = parseFloat(precio);
    return isNaN(n) ? -Infinity : n;
  }

  function ordenar(lista, criterio){
    const copia = lista.slice();
    if(criterio === 'precio-asc'){
      copia.sort((a, b) => numeroPrecio(a.Precio) - numeroPrecio(b.Precio));
    } else if(criterio === 'precio-desc'){
      copia.sort((a, b) => numeroPrecio(b.Precio) - numeroPrecio(a.Precio));
    } else if(criterio === 'nombre-asc'){
      copia.sort((a, b) => normalizar(a.Nombre).localeCompare(normalizar(b.Nombre)));
    }
    return copia;
  }

  function render(){
    const termino = normalizar(searchInput.value.trim());
    const escala = scaleSelect.value;

    let filtrados = PRODUCTOS.filter(p =>
      coincideBusqueda(p, termino) &&
      (!escala || (p.Escala || '') === escala)
    );
    filtrados = ordenar(filtrados, sortSelect.value);

    const totalProductos = filtrados.length;
    const totalPaginas = Math.max(1, Math.ceil(totalProductos / PRODUCTOS_POR_PAGINA));
    if(paginaActual > totalPaginas) paginaActual = totalPaginas;
    if(paginaActual < 1) paginaActual = 1;

    const inicio = (paginaActual - 1) * PRODUCTOS_POR_PAGINA;
    const pagina = filtrados.slice(inicio, inicio + PRODUCTOS_POR_PAGINA);

    grid.innerHTML = '';
    const frag = document.createDocumentFragment();
    pagina.forEach(p => frag.appendChild(crearTarjeta(p)));
    grid.appendChild(frag);

    emptyState.hidden = totalProductos !== 0;
    counter.textContent = totalProductos === 1 ? '1 producto' : `${totalProductos} productos`;

    pagination.hidden = totalProductos === 0;
    if(totalProductos > 0){
      pageInfo.textContent = `Página ${paginaActual} de ${totalPaginas}`;
      prevPageBtn.disabled = paginaActual <= 1;
      nextPageBtn.disabled = paginaActual >= totalPaginas;
    }
  }

  function irAPagina(delta){
    paginaActual += delta;
    render();
    if(typeof grid.scrollIntoView === 'function'){
      grid.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  }

  function parseCSV(texto){
    const filas = [];
    let fila = [], campo = '', dentroComillas = false;
    for(let i = 0; i < texto.length; i++){
      const c = texto[i];
      if(dentroComillas){
        if(c === '"'){
          if(texto[i+1] === '"'){ campo += '"'; i++; }
          else dentroComillas = false;
        } else {
          campo += c;
        }
      } else if(c === '"'){
        dentroComillas = true;
      } else if(c === ','){
        fila.push(campo); campo = '';
      } else if(c === '\n' || c === '\r'){
        if(c === '\r' && texto[i+1] === '\n') i++;
        fila.push(campo); campo = '';
        filas.push(fila);
        fila = [];
      } else {
        campo += c;
      }
    }
    if(campo !== '' || fila.length){ fila.push(campo); filas.push(fila); }
    return filas.filter(f => f.some(v => v !== ''));
  }

  function csvAProductos(texto){
    const filas = parseCSV(texto);
    if(filas.length < 2) return [];
    const encabezados = filas[0].map(h => h.trim().toLowerCase());
    const indice = {};
    ['sku', 'marca', 'escala', 'nombre', 'trademark', 'precio', 'imagen_url', 'foto', 'url_producto'].forEach(campo => {
      indice[campo] = encabezados.indexOf(campo);
    });
    // Alias: si la hoja usa "Foto" en vez de "Imagen_URL" para la imagen, igual funciona
    if(indice.imagen_url < 0 && indice.foto >= 0){
      indice.imagen_url = indice.foto;
    }
    const obtener = (fila, campo) => (indice[campo] >= 0 ? (fila[indice[campo]] || '').trim() : '');

    return filas.slice(1)
      .map(fila => ({
        SKU: obtener(fila, 'sku'),
        Marca: obtener(fila, 'marca'),
        Escala: obtener(fila, 'escala'),
        Nombre: obtener(fila, 'nombre'),
        Trademark: obtener(fila, 'trademark'),
        Precio: obtener(fila, 'precio'),
        Imagen_URL: obtener(fila, 'imagen_url'),
        URL_Producto: obtener(fila, 'url_producto'),
      }))
      .filter(p => p.Nombre);
  }

  async function cargarDesdeSheet(){
    if(!SHEET_CSV_URL) return;
    try{
      const resp = await fetch(SHEET_CSV_URL, { cache: 'no-store' });
      if(!resp.ok) throw new Error('HTTP ' + resp.status);
      const texto = await resp.text();
      const nuevos = csvAProductos(texto);
      if(nuevos.length){
        // Si el Sheet no trae imagen para una fila (ej. la columna Foto sigue
        // con la fórmula =IMAGE() y no exporta como texto), se mantiene la
        // imagen que ya venía cargada en el catálogo inicial, en vez de dejarla
        // en blanco. Se asume mismo orden de filas entre ambas fuentes.
        const fusionados = nuevos.map((p, i) => {
          if(!p.Imagen_URL && PRODUCTOS_INICIALES[i] && PRODUCTOS_INICIALES[i].Imagen_URL){
            return { ...p, Imagen_URL: PRODUCTOS_INICIALES[i].Imagen_URL };
          }
          return p;
        });
        PRODUCTOS = fusionados;
        poblarEscalas();
        paginaActual = 1;
        render();
      }
    } catch(e){
      console.error('No se pudo actualizar desde Google Sheets:', e);
    }
  }

  function resetPaginaYRender(){
    paginaActual = 1;
    render();
  }

  poblarEscalas();
  searchInput.addEventListener('input', resetPaginaYRender);
  scaleSelect.addEventListener('change', resetPaginaYRender);
  sortSelect.addEventListener('change', resetPaginaYRender);
  prevPageBtn.addEventListener('click', () => irAPagina(-1));
  nextPageBtn.addEventListener('click', () => irAPagina(1));
  render();
  cargarDesdeSheet();
</script>
</body>
</html>
"""


# ==============================================================================
# 5.5 SINCRONIZACIÓN CON MATRIZ A PEDIDO (Google Sheets, por SKU)
# ==============================================================================
# Compara lo que acaba de traer el scraper contra tu Sheet privado
# "MATRIZ A PEDIDO" (hoja "copia de productos"), emparejando por SKU:
#   - Producto que YA EXISTE (mismo SKU)  -> actualiza Foto/Marca/Escala/Nombre,
#     NUNCA toca costos ni el precio público (columna F en adelante).
#   - Producto NUEVO (SKU que no estaba)  -> se agrega como fila nueva, con el
#     costo en blanco para que tú lo completes.
#   - Producto que YA NO aparece en el proveedor -> se marca "DESCONTINUADO"
#     en la columna de Estado, sin borrar la fila.
#
# Requiere: pip install gspread google-auth
# Y una cuenta de servicio de Google con acceso de Editor al Sheet privado
# (ver instrucciones que te di aparte para crearla y descargar el JSON).

SHEET_ID_MATRIZ = "PON_AQUI_EL_ID_DE_MATRIZ_A_PEDIDO"
HOJA_MATRIZ = "copia de productos"
ARCHIVO_CREDENCIALES_GOOGLE = "credenciales_google.json"

# Mapa de columnas real de MATRIZ A PEDIDO (confirmado con Dylan):
COL_FOTO = "A"
COL_MARCA = "B"
COL_ESCALA = "C"
COL_NOMBRE = "D"
COL_SKU = "E"
# COL F = FOB EUR, columna G-J = costos/margen/precio público -> nunca se tocan
COL_ESTADO = "K"  # AJUSTAR si ya usas la columna K para otra cosa

FILA_INICIO_MATRIZ = 2  # primera fila de datos, después del encabezado


def _col_a_indice(col_letra: str) -> int:
    """Convierte 'A' -> 1, 'B' -> 2, 'K' -> 11, etc."""
    indice = 0
    for c in col_letra:
        indice = indice * 26 + (ord(c.upper()) - ord("A") + 1)
    return indice


def sincronizar_con_matriz(df_scrapeado: pd.DataFrame) -> None:
    """
    Sincroniza df_scrapeado (resultado de scrapear_con_requests()/
    scrapear_con_playwright()) contra MATRIZ A PEDIDO, emparejando por SKU.
    No requiere ni toca generar_catalogo_web() — son pasos independientes.
    """
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError:
        logger.error(
            "Faltan librerías: corre 'pip install gspread google-auth' "
            "antes de usar --sincronizar."
        )
        return

    if df_scrapeado is None or df_scrapeado.empty:
        logger.warning("No hay productos scrapeados — no se sincroniza nada.")
        return

    creds = Credentials.from_service_account_file(
        ARCHIVO_CREDENCIALES_GOOGLE,
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    cliente = gspread.authorize(creds)
    hoja = cliente.open_by_key(SHEET_ID_MATRIZ).worksheet(HOJA_MATRIZ)

    todas_las_filas = hoja.get_all_values()
    filas_datos = todas_las_filas[FILA_INICIO_MATRIZ - 1:]

    idx_sku = _col_a_indice(COL_SKU) - 1

    # SKU -> número de fila real en el Sheet (1-indexed)
    fila_por_sku = {}
    for i, fila in enumerate(filas_datos):
        sku = fila[idx_sku].strip() if len(fila) > idx_sku else ""
        if sku:
            fila_por_sku[sku] = FILA_INICIO_MATRIZ + i

    skus_scrapeados = set()
    actualizaciones = []  # [(fila, col_letra, valor), ...]
    filas_nuevas = []

    for _, p in df_scrapeado.iterrows():
        sku = str(p.get("SKU", "")).strip()
        if not sku:
            continue
        skus_scrapeados.add(sku)

        if sku in fila_por_sku:
            fila = fila_por_sku[sku]
            actualizaciones.append((fila, COL_FOTO, p.get("Imagen_URL", "") or ""))
            actualizaciones.append((fila, COL_MARCA, p.get("Marca", "") or ""))
            actualizaciones.append((fila, COL_ESCALA, p.get("Escala", "") or ""))
            actualizaciones.append((fila, COL_NOMBRE, p.get("Nombre", "") or ""))
            actualizaciones.append((fila, COL_ESTADO, "ACTIVO"))
        else:
            filas_nuevas.append(p)

    # Productos que ya no aparecen en el proveedor -> marcar, no borrar
    for sku, fila in fila_por_sku.items():
        if sku not in skus_scrapeados:
            actualizaciones.append((fila, COL_ESTADO, "DESCONTINUADO"))

    if actualizaciones:
        celdas = []
        for fila, col, valor in actualizaciones:
            celda = hoja.cell(fila, _col_a_indice(col))
            celda.value = valor
            celdas.append(celda)
        hoja.update_cells(celdas, value_input_option="USER_ENTERED")

    if filas_nuevas:
        filas_a_insertar = []
        ancho_fila = _col_a_indice(COL_ESTADO)
        for p in filas_nuevas:
            fila = [""] * ancho_fila
            fila[_col_a_indice(COL_FOTO) - 1] = p.get("Imagen_URL", "") or ""
            fila[_col_a_indice(COL_MARCA) - 1] = p.get("Marca", "") or ""
            fila[_col_a_indice(COL_ESCALA) - 1] = p.get("Escala", "") or ""
            fila[_col_a_indice(COL_NOMBRE) - 1] = p.get("Nombre", "") or ""
            fila[_col_a_indice(COL_SKU) - 1] = p.get("SKU", "") or ""
            fila[_col_a_indice(COL_ESTADO) - 1] = "ACTIVO (nuevo, falta costo)"
            filas_a_insertar.append(fila)
        hoja.append_rows(filas_a_insertar, value_input_option="USER_ENTERED")

    descontinuados = sum(1 for s in fila_por_sku if s not in skus_scrapeados)
    logger.info(
        f"Sincronización con MATRIZ A PEDIDO completa: "
        f"{len(fila_por_sku)} productos existentes revisados, "
        f"{len(filas_nuevas)} nuevos agregados, "
        f"{descontinuados} marcados como DESCONTINUADO."
    )


def generar_catalogo_web(df: pd.DataFrame = None, ruta_salida: str = "catalogo.html") -> None:
    """
    Genera un catálogo web en un único archivo HTML (con CSS y JS incrustados)
    a partir del DataFrame de productos scrapeados y/o de un Google Sheet
    publicado como CSV (ver SHEET_CSV_URL arriba del archivo).

    Incluye:
      - Header con espacio para logo (RUTA_LOGO, editable arriba) y nombre
        de tienda (NOMBRE_TIENDA, editable arriba).
      - Buscador en tiempo real (filtra por Nombre, Marca, Trademark y SKU).
      - Selector de Escala que se llena dinámicamente con las escalas
        presentes en los datos, ordenadas numéricamente (1:18, 1:24, 1:43...).
      - Selector de orden (precio asc/desc, nombre A-Z).
      - Contador de productos mostrados.
      - Grid de tarjetas responsivo (imagen, escala, marca, nombre, SKU, precio).
      - Si SHEET_CSV_URL está configurado, la página además intenta cargar
        los datos en vivo desde ese Sheet cada vez que alguien la abre,
        reemplazando los datos iniciales cuando la carga tiene éxito.

    Parámetros
    ----------
    df : pd.DataFrame o None
        Datos iniciales (se muestran de inmediato al abrir la página, y
        sirven de respaldo si el Sheet no carga). Puede omitirse si vas a
        depender 100% de SHEET_CSV_URL.
    ruta_salida : str
        Nombre/ruta del archivo .html a generar. Por defecto "catalogo.html".
    """
    columnas_esperadas = [
        "SKU", "Marca", "Escala", "Nombre", "Trademark", "Precio",
        "Imagen_URL", "URL_Producto",
    ]

    if df is not None and not df.empty:
        df = df.copy()
        for col in columnas_esperadas:
            if col not in df.columns:
                df[col] = ""
        df = df[columnas_esperadas].fillna("")
        productos = df.to_dict(orient="records")
        # ensure_ascii=False conserva tildes/ñ legibles; el .replace("</", "<\\/")
        # evita que un valor con "</script>" rompa la etiqueta.
        productos_json = json.dumps(productos, ensure_ascii=False).replace("</", "<\\/")
    else:
        productos_json = "[]"
        if not SHEET_CSV_URL.strip():
            logger.warning(
                "No hay DataFrame ni SHEET_CSV_URL configurado — no se generará el catálogo."
            )
            return

    html = (
        _PLANTILLA_HTML
        .replace("__TITULO__", _html_escape(NOMBRE_TIENDA))
        .replace("__LOGO_SRC__", _html_escape(RUTA_LOGO))
        .replace("__PRODUCTS_JSON__", productos_json)
        .replace("__SHEET_CSV_URL__", _html_escape(SHEET_CSV_URL))
    )

    with open(ruta_salida, "w", encoding="utf-8") as f:
        f.write(html)

    modo = "en vivo desde Google Sheets" if SHEET_CSV_URL.strip() else "estático"
    n = len(df) if df is not None else 0
    logger.info(f"Catálogo web generado ({modo}): {ruta_salida}  ({n} productos iniciales)")


# ==============================================================================
# 6. MAIN
# ==============================================================================
def main():
    parser = argparse.ArgumentParser(description="Scraper de catálogo modelcarswholesale.com")
    parser.add_argument(
        "--engine",
        choices=["requests", "playwright"],
        default="playwright",
        help="Motor de scraping a usar. 'playwright' recomendado si hay bot-detection.",
    )
    parser.add_argument(
        "--login",
        action="store_true",
        help="Intenta iniciar sesión con las credenciales MCW_USER / MCW_PASS (variables de entorno).",
    )
    parser.add_argument(
        "--sincronizar",
        action="store_true",
        help="Después de scrapear, sincroniza productos nuevos/descontinuados con MATRIZ A PEDIDO por SKU.",
    )
    args = parser.parse_args()

    logger.info(f"Iniciando scraping con motor: {args.engine}")

    if args.engine == "requests":
        df = scrapear_con_requests(usar_login=args.login)
    else:
        df = scrapear_con_playwright(usar_login=args.login)

    logger.info(f"Total de productos extraídos (todas las secciones): {len(df)}")

    if args.sincronizar:
        sincronizar_con_matriz(df)
    else:
        generar_catalogo_web(df, OUTPUT_FILE)


if __name__ == "__main__":
    main()