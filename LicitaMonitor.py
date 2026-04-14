"""
LICITAMONITOR — Soldesp
========================
Monitor de licitaciones industriales — Portal Minero.
Fuentes: /wp/oportunidades-de-negocios/ · /display/acce/Muro+de+Actividades

REQUISITOS:
    pip install selenium openpyxl webdriver-manager requests

CONFIGURACIÓN:
    Editar config.json con credenciales de Portal Minero.
"""

import json
import logging
import os
import re
import smtplib
import sys
import time
from abc import ABC, abstractmethod
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path

from selenium import webdriver
from selenium.common.exceptions import (
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ═══════════════════════════════════════════════════════════════
# PALETA OSCURA CORPORATIVA
# ═══════════════════════════════════════════════════════════════

C_HEADER_BG      = "0A1628"   # Casi negro azulado — cabeceras principales
C_SUBHEADER_BG   = "0F2236"   # Azul muy oscuro — subtítulos
C_ROW_EVEN       = "112033"   # Filas pares
C_ROW_ODD        = "0D1A2B"   # Filas impares
C_RESUMEN_BG     = "0F1E30"   # Fondo hoja resumen
C_RESUMEN_ROW    = "162840"   # Filas hoja resumen

C_BADGE_MINERO   = "6B1515"   # Rojo oscuro — badge Portal Minero

C_TEXT_HEADER    = "FFFFFF"   # Blanco puro — texto en cabeceras
C_TEXT_DATA      = "C8DCF0"   # Azul muy claro — texto en celdas
C_TEXT_MUTED     = "5A7A99"   # Gris azulado — texto secundario
C_LINK           = "4FC3F7"   # Azul eléctrico — hipervínculos
C_BORDER         = "1A3550"   # Borde oscuro


# ═══════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ═══════════════════════════════════════════════════════════════

def cargar_config() -> dict:
    """Carga config.json desde el directorio del ejecutable o script."""
    base = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
    config_path = base / "config.json"

    if not config_path.exists():
        _crear_config_default(config_path)
        print(f"[INFO] config.json creado en {config_path}. Edítalo antes de continuar.")
        sys.exit(0)

    with open(config_path, encoding="utf-8") as f:
        return json.load(f)


def _crear_config_default(path: Path) -> None:
    config = {
        "_comentario": "Edita este archivo con tus credenciales. NO subas a GitHub.",
        "_version": "3.1",
        "portales": {
            "portal_minero": {
                "activo": True,
                "usuario": "",
                "password": ""
            }
        },
        "rutas": {
            "output_dir": str(Path.home() / "Desktop" / "LicitaMonitor"),
            "log_dir":    str(Path.home() / "Desktop" / "LicitaMonitor" / "logs"),
        },
        "filtros": {
            "horas_atras": 24,
            "keywords": [
                "estructura", "caldereria", "caldería",
                "piping", "fabricacion", "fabricación",
                "montaje", "spool", "pipeline", "manifold",
                "tk", "estanques", "soldadura"
            ],
        },
        "selenium": {
            "headless": True,
            "timeout_segundos": 30,
            "max_reintentos": 3,
            "pausa_entre_paginas": 2,
        },
        "notificaciones": {
            "activar_email": False,
            "smtp_servidor": "smtp.gmail.com",
            "smtp_puerto": 587,
            "smtp_usuario": "",
            "smtp_password": "",
            "destinatarios": [],
        },
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


# ═══════════════════════════════════════════════════════════════
# LOGGING
# ═══════════════════════════════════════════════════════════════

def configurar_logging(log_dir: str) -> logging.Logger:
    Path(log_dir).mkdir(parents=True, exist_ok=True)
    fecha    = datetime.now().strftime("%Y-%m-%d")
    log_file = Path(log_dir) / f"licitamonitor_{fecha}.log"

    logger = logging.getLogger("LicitaMonitor")
    if logger.handlers:
        return logger  # ya configurado

    logger.setLevel(logging.DEBUG)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S")

    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)

    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)

    logger.addHandler(fh)
    logger.addHandler(ch)
    return logger


# ═══════════════════════════════════════════════════════════════
# WEBDRIVER COMPARTIDO
# ═══════════════════════════════════════════════════════════════

def iniciar_driver(cfg_selenium: dict) -> webdriver.Chrome:
    options = webdriver.ChromeOptions()
    if cfg_selenium.get("headless", True):
        options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--lang=es-CL")
    # User-Agent real para evitar bloqueos por bot-detection (Portal Minero, etc.)
    options.add_argument(
        "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    # Logging de red para interceptar XHR (útil en SPAs como Wherex)
    options.set_capability("goog:loggingPrefs", {"performance": "ALL"})

    service = Service(ChromeDriverManager().install())
    driver  = webdriver.Chrome(service=service, options=options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver


def _guardar_screenshot(driver, nombre: str, logger) -> None:
    try:
        driver.save_screenshot(str(Path(nombre)))
        logger.debug("Screenshot: %s", nombre)
    except Exception:
        pass


def _es_relevante(texto: str, keywords: list) -> bool:
    """Verifica relevancia. Keywords de ≤2 chars usan word boundary para evitar falsos positivos."""
    return bool(_keywords_encontrados(texto, keywords))


def _keywords_encontrados(texto: str, keywords: list) -> str:
    """Devuelve string con los keywords encontrados en el texto, separados por ' · '."""
    txt = texto.lower()
    encontrados = []
    for kw in keywords:
        kw_l = kw.lower()
        if len(kw_l) <= 2:
            if re.search(r'\b' + re.escape(kw_l) + r'\b', txt):
                encontrados.append(kw)
        else:
            if kw_l in txt:
                encontrados.append(kw)
    return " · ".join(encontrados)


def _dentro_de_ventana(fecha_str: str, horas_atras: int, logger) -> bool:
    """True si la fecha está dentro de las últimas N horas. Retorna True si no parseable."""
    if not fecha_str or fecha_str.strip() in ("—", "-", ""):
        return True  # Sin fecha → incluir (portales sin columna de fecha)
    limite   = datetime.now() - timedelta(hours=horas_atras)
    formatos = [
        "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y",
        "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d",
        "%d-%m-%Y", "%d.%m.%Y",
    ]
    for fmt in formatos:
        try:
            return datetime.strptime(fecha_str.strip(), fmt) >= limite
        except ValueError:
            continue
    logger.debug("Fecha no parseable '%s' — incluida por defecto.", fecha_str)
    return True


# ═══════════════════════════════════════════════════════════════
# BASE PORTAL
# ═══════════════════════════════════════════════════════════════

class PortalBase(ABC):
    """Interfaz común para todos los scrapers de portal."""

    nombre_portal: str  = "Portal"
    color_badge:   str  = "1255A3"
    tiene_fechas:  bool = True   # False en portales sin columna de fecha

    def __init__(self, driver: webdriver.Chrome, cfg: dict, logger: logging.Logger):
        self.driver  = driver
        self.cfg     = cfg
        self.logger  = logger
        self.timeout = int(cfg.get("selenium", {}).get("timeout_segundos", 30))
        self.pausa   = float(cfg.get("selenium", {}).get("pausa_entre_paginas", 2))

    @abstractmethod
    def login(self) -> None: ...

    @abstractmethod
    def extraer_licitaciones(self, keywords: list, horas_atras: int) -> list: ...

    # Campos estándar que debe devolver cada licitación:
    # {
    #   "portal":           str,
    #   "id":               str,
    #   "comprador":        str,
    #   "descripcion":      str,
    #   "region":           str,
    #   "fecha_pub":        str,
    #   "fecha_cierre":     str,
    #   "link":             str,
    # }


# ═══════════════════════════════════════════════════════════════
# PORTAL: PORTAL MINERO
#
# Fuente 1 — PÚBLICA (no requiere login):
#   https://www.portalminero.com/wp/oportunidades-de-negocios/
#   · 50 artículos WordPress (selector: article, skip primero)
#   · Cada article: texto "Publicado en Portal Minero el: DD/MM/YYYY\n[Título]"
#   · Link: <a href="https://www.portalminero.com/pages/viewpage.action?pageId=...">
#
# Fuente 2 — REQUIERE LOGIN (Confluence Seraph):
#   https://www.portalminero.com/display/acce/Muro+de+Actividades
#   · ~15 elementos .update-item
#   · Cada item: Título\nactualizando hace X
#   · Link: <a href="https://www.portalminero.com/pages/viewpage.action?pageId=...">
#
# Login: POST a /dologin.action, campos os_username / os_password,
#        botón input[type='submit'][name='login'] (NO existe #loginButton)
# ═══════════════════════════════════════════════════════════════

class PortalMinero(PortalBase):
    nombre_portal = "Portal Minero"
    color_badge   = C_BADGE_MINERO
    tiene_fechas  = True  # Fuente 1 tiene "Publicado en Portal Minero el: DD/MM/YYYY"

    LOGIN_URL   = "https://www.portalminero.com/login.action"
    BASE_URL    = "https://www.portalminero.com"
    FUENTE_WP   = "https://www.portalminero.com/wp/oportunidades-de-negocios/"
    FUENTE_MURO = "https://www.portalminero.com/display/acce/Muro+de+Actividades"

    # ──────────────────────────────────────────────────────────────
    # LOGIN
    # ──────────────────────────────────────────────────────────────

    def login(self) -> None:
        self.logger.info("[Portal Minero] Iniciando sesión...")
        cfg    = self.cfg["portales"]["portal_minero"]
        wait   = WebDriverWait(self.driver, self.timeout)

        self.driver.get(self.LOGIN_URL)

        # Esperar campo os_username
        try:
            wait.until(EC.presence_of_element_located((By.NAME, "os_username")))
        except TimeoutException:
            _guardar_screenshot(self.driver, "pm_login_error.png", self.logger)
            raise RuntimeError("[Portal Minero] Formulario de login no cargó.")

        self.driver.find_element(By.NAME, "os_username").send_keys(cfg["usuario"])
        self.driver.find_element(By.NAME, "os_password").send_keys(cfg["password"])

        # Botón submit: input[type='submit'] con name='login'
        # (no existe #loginButton en este sitio)
        self.driver.find_element(By.CSS_SELECTOR, "input[type='submit']").click()

        # Seraph redirige fuera de /login.action al autenticar correctamente
        try:
            wait.until(lambda d: "login.action" not in d.current_url)
            self.logger.info("[Portal Minero] Login exitoso → %s", self.driver.current_url)
        except TimeoutException:
            msg = "—"
            for sel in [".error", "#loginErrorDiv", ".aui-message-error"]:
                try:
                    msg = self.driver.find_element(By.CSS_SELECTOR, sel).text.strip()
                    break
                except Exception:
                    pass
            _guardar_screenshot(self.driver, "pm_login_failed.png", self.logger)
            raise RuntimeError(f"[Portal Minero] Login fallido: {msg}")

    # ──────────────────────────────────────────────────────────────
    # EXTRACCIÓN PRINCIPAL
    # ──────────────────────────────────────────────────────────────

    def extraer_licitaciones(self, keywords: list, horas_atras: int) -> list:
        resultados: list = []
        vistos: set      = set()

        # Fuente 1: página pública WordPress
        try:
            r = self._scrapear_oportunidades(keywords, horas_atras, vistos)
            self.logger.info("[Portal Minero] Oportunidades WP: %d relevantes", len(r))
            resultados.extend(r)
        except Exception as exc:
            self.logger.warning("[Portal Minero] Error Oportunidades WP: %s", exc)
            _guardar_screenshot(self.driver, "pm_error_wp.png", self.logger)

        # Fuente 2: Muro de Actividades Confluence (requiere sesión)
        try:
            r = self._scrapear_muro(keywords, vistos)
            self.logger.info("[Portal Minero] Muro Actividades: %d relevantes", len(r))
            resultados.extend(r)
        except Exception as exc:
            self.logger.warning("[Portal Minero] Error Muro Actividades: %s", exc)
            _guardar_screenshot(self.driver, "pm_error_muro.png", self.logger)

        self.logger.info("[Portal Minero] Total: %d licitaciones.", len(resultados))
        return resultados

    # ──────────────────────────────────────────────────────────────
    # FUENTE 1 — /wp/oportunidades-de-negocios/
    # ──────────────────────────────────────────────────────────────

    def _scrapear_oportunidades(self, keywords: list, horas_atras: int,
                                vistos: set) -> list:
        """
        Estructura confirmada:
          - article[0]: contenedor con los 50 links (se omite)
          - article[1..N]: cada oportunidad individual
              TEXT: "Publicado en Portal Minero el: DD/MM/YYYY\n[Título]"
              LINK: a[href*='viewpage.action']
        """
        self.driver.get(self.FUENTE_WP)
        wait = WebDriverWait(self.driver, self.timeout)
        try:
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "article")))
            time.sleep(1)
        except TimeoutException:
            _guardar_screenshot(self.driver, "pm_error_wp_load.png", self.logger)
            return []

        articles = self.driver.find_elements(By.CSS_SELECTOR, "article")
        # Omitir article[0] (es el contenedor general)
        items = articles[1:] if len(articles) > 1 else articles
        self.logger.debug("[Portal Minero] WP articles: %d", len(items))

        limite = datetime.now() - timedelta(hours=horas_atras)
        resultados = []

        for art in items:
            try:
                texto = art.text.strip()
                if not texto:
                    continue

                # Extraer fecha: "Publicado en Portal Minero el: DD/MM/YYYY"
                fecha_pub = "—"
                m = re.search(r'Publicado en Portal Minero el:\s*(\d{2}/\d{2}/\d{4})', texto)
                if m:
                    fecha_pub = m.group(1)
                    # Filtrar por ventana de tiempo
                    try:
                        fdt = datetime.strptime(fecha_pub, "%d/%m/%Y")
                        if fdt < limite:
                            continue
                    except ValueError:
                        pass

                # Extraer título y link del primer <a>
                try:
                    a    = art.find_element(By.CSS_SELECTOR, "a[href]")
                    titulo = a.text.strip()
                    link   = a.get_attribute("href") or ""
                except NoSuchElementException:
                    titulo = texto.split("\n")[-1][:200]
                    link   = ""

                if not titulo:
                    continue

                # Filtrar por keywords
                kws = _keywords_encontrados(titulo, keywords)
                if not kws:
                    continue

                clave = titulo[:60]
                if clave in vistos:
                    continue
                vistos.add(clave)

                resultados.append({
                    "portal":         self.nombre_portal,
                    "fuente":         "Oportunidades de Negocio",
                    "id":             "—",
                    "comprador":      "—",
                    "descripcion":    titulo,
                    "region":         "—",
                    "fecha_pub":      fecha_pub,
                    "fecha_cierre":   "—",
                    "link":           link,
                    "keywords_match": kws,
                })

            except StaleElementReferenceException:
                continue
            except Exception as exc:
                self.logger.debug("[Portal Minero] Error article: %s", exc)

        return resultados

    # ──────────────────────────────────────────────────────────────
    # FUENTE 2 — /display/acce/Muro+de+Actividades
    # ──────────────────────────────────────────────────────────────

    def _scrapear_muro(self, keywords: list, vistos: set) -> list:
        """
        Estructura confirmada:
          - .update-item: cada actividad reciente
              TEXT: "[Título]\nactualizado hace X / actualizado ayer a las HH:MM"
              LINK[0]: a[href*='viewpage.action'] → página de detalle
              LINK[1]: a[href*='diffpages'] → diff de versiones (ignorar)
        """
        self.driver.get(self.FUENTE_MURO)
        wait = WebDriverWait(self.driver, self.timeout)
        try:
            wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, ".update-item")))
            time.sleep(1)
        except TimeoutException:
            _guardar_screenshot(self.driver, "pm_error_muro_load.png", self.logger)
            return []

        items = self.driver.find_elements(By.CSS_SELECTOR, ".update-item")
        self.logger.debug("[Portal Minero] Muro items: %d", len(items))

        resultados = []

        for item in items:
            try:
                texto = item.text.strip()
                if not texto:
                    continue

                lineas = [l.strip() for l in texto.split("\n") if l.strip()]
                titulo = lineas[0] if lineas else texto[:200]

                # Fecha relativa (ej. "actualizado hace 2 horas")
                fecha_rel = lineas[1] if len(lineas) > 1 else "—"

                # Primer link que NO sea diffpages
                link = ""
                for a in item.find_elements(By.CSS_SELECTOR, "a[href]"):
                    href = a.get_attribute("href") or ""
                    if "diffpages" not in href and href:
                        link = href
                        break

                # Filtrar por keywords
                kws = _keywords_encontrados(titulo, keywords)
                if not kws:
                    continue

                clave = titulo[:60]
                if clave in vistos:
                    continue
                vistos.add(clave)

                resultados.append({
                    "portal":         self.nombre_portal,
                    "fuente":         "Muro de Actividades",
                    "id":             "—",
                    "comprador":      "—",
                    "descripcion":    titulo,
                    "region":         "—",
                    "fecha_pub":      fecha_rel,
                    "fecha_cierre":   "—",
                    "link":           link,
                    "keywords_match": kws,
                })

            except StaleElementReferenceException:
                continue
            except Exception as exc:
                self.logger.debug("[Portal Minero] Error muro item: %s", exc)

        return resultados

    def _encontrar_elemento(self, selectors: list):
        for by, sel in selectors:
            try:
                el = self.driver.find_element(by, sel)
                if el.is_displayed():
                    return el
            except Exception:
                continue
        return None


# ═══════════════════════════════════════════════════════════════
# GENERACIÓN DE EXCEL — PALETA OSCURA
# ═══════════════════════════════════════════════════════════════

def _borde_oscuro():
    lado = Side(style="thin", color=C_BORDER)
    return Border(left=lado, right=lado, top=lado, bottom=lado)

def _fill(color: str) -> PatternFill:
    return PatternFill("solid", fgColor=color)

def _font(bold=False, size=10, color=C_TEXT_DATA, italic=False, underline=None):
    return Font(bold=bold, size=size, color=color, italic=italic,
                underline=underline or "none")


def generar_excel(
    licitaciones: list,
    output_dir: str,
    keywords: list,
    horas_atras: int,
    logger: logging.Logger,
) -> str:
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    ts       = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filepath = Path(output_dir) / f"LicitaMonitor_Soldesp_{ts}.xlsx"

    wb = openpyxl.Workbook()

    _hoja_licitaciones(wb.active, licitaciones, keywords, horas_atras)

    # Hojas por portal
    portales_presentes = sorted({l["portal"] for l in licitaciones})
    for portal in portales_presentes:
        subset = [l for l in licitaciones if l["portal"] == portal]
        hoja_nombre = portal[:31]
        ws_p = wb.create_sheet(hoja_nombre)
        _hoja_licitaciones(ws_p, subset, keywords, horas_atras, subtitulo=f"Portal: {portal}")

    _hoja_resumen(wb.create_sheet("Resumen"), licitaciones, keywords, horas_atras, filepath.name)

    try:
        wb.save(str(filepath))
        logger.info("Excel guardado: %s", filepath)
    except PermissionError:
        alt = filepath.with_stem(filepath.stem + "_v2")
        wb.save(str(alt))
        logger.warning("Archivo en uso — guardado como: %s", alt)
        return str(alt)

    return str(filepath)


def _hoja_licitaciones(ws, licitaciones: list, keywords: list, horas_atras: int,
                       subtitulo: str | None = None) -> None:
    ws.sheet_view.showGridLines = False

    borde = _borde_oscuro()
    aln_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
    aln_i = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    NCOLS = 7  # A..G

    # ── Fila 1: Título principal ─────────────────────────────────────────
    ws.merge_cells(f"A1:{get_column_letter(NCOLS)}1")
    c = ws["A1"]
    c.value     = (f"LICITAMONITOR — SOLDESP  ·  "
                   f"{datetime.now().strftime('%d/%m/%Y %H:%M')}  ·  "
                   f"Período: últimas {horas_atras}h")
    c.font      = _font(bold=True, size=13, color=C_TEXT_HEADER)
    c.fill      = _fill(C_HEADER_BG)
    c.alignment = aln_c
    ws.row_dimensions[1].height = 34

    # ── Fila 2: Filtros / subtítulo ──────────────────────────────────────
    ws.merge_cells(f"A2:{get_column_letter(NCOLS)}2")
    c2 = ws["A2"]
    c2.value     = subtitulo or ("Filtros: " + " · ".join(kw.capitalize() for kw in keywords))
    c2.font      = _font(italic=True, size=9, color=C_TEXT_MUTED)
    c2.fill      = _fill(C_SUBHEADER_BG)
    c2.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 18

    # ── Fila 3: Cabeceras ────────────────────────────────────────────────
    # Columnas orientadas a evaluar si vale la pena participar
    headers = [
        "Fuente",           # A — Oportunidades / Muro Actividades
        "Título / Servicio solicitado",  # B
        "Keywords detectados",  # C
        "Fecha Publicación",    # D
        "Fecha Cierre",         # E
        "Empresa / Región",     # F
        "Link Directo",         # G
    ]
    anchos = [22, 60, 30, 18, 16, 22, 22]

    for col_i, (hdr, ancho) in enumerate(zip(headers, anchos), start=1):
        c = ws.cell(row=3, column=col_i, value=hdr)
        c.font      = _font(bold=True, size=10, color=C_TEXT_HEADER)
        c.fill      = _fill(C_HEADER_BG)
        c.alignment = aln_c
        c.border    = borde
        ws.column_dimensions[get_column_letter(col_i)].width = ancho
    ws.row_dimensions[3].height = 26

    # ── Filas de datos ───────────────────────────────────────────────────
    if not licitaciones:
        ws.merge_cells(f"A4:{get_column_letter(NCOLS)}4")
        c = ws["A4"]
        c.value     = "No se encontraron licitaciones para los keywords configurados."
        c.font      = _font(italic=True, color=C_TEXT_MUTED)
        c.fill      = _fill(C_ROW_ODD)
        c.alignment = aln_c
        ws.row_dimensions[4].height = 28
        return

    for row_i, lic in enumerate(licitaciones, start=4):
        bg        = C_ROW_EVEN if row_i % 2 == 0 else C_ROW_ODD
        fill_data = _fill(bg)

        # A — Fuente (badge con color de portal)
        fuente = lic.get("fuente", lic["portal"])
        c_f = ws.cell(row=row_i, column=1, value=fuente)
        c_f.font      = _font(bold=True, size=9, color=C_TEXT_HEADER)
        c_f.fill      = _fill(_color_portal(lic["portal"]))
        c_f.alignment = aln_c
        c_f.border    = borde

        # B — Título (texto largo, wrap)
        c_t = ws.cell(row=row_i, column=2, value=lic["descripcion"])
        c_t.font      = _font(size=9, color=C_TEXT_DATA, bold=True)
        c_t.fill      = fill_data
        c_t.alignment = aln_i
        c_t.border    = borde

        # C — Keywords detectados
        kws = lic.get("keywords_match") or _keywords_encontrados(
            f"{lic['descripcion']} {lic.get('comprador','')} {lic.get('region','')}",
            keywords
        )
        c_k = ws.cell(row=row_i, column=3, value=kws)
        c_k.font      = _font(size=9, color="FFD54F")  # amarillo cálido — destaca
        c_k.fill      = fill_data
        c_k.alignment = aln_c
        c_k.border    = borde

        # D — Fecha publicación
        c_fp = ws.cell(row=row_i, column=4, value=lic["fecha_pub"])
        c_fp.font      = _font(size=9, color=C_TEXT_DATA)
        c_fp.fill      = fill_data
        c_fp.alignment = aln_c
        c_fp.border    = borde

        # E — Fecha cierre
        c_fc = ws.cell(row=row_i, column=5, value=lic["fecha_cierre"])
        c_fc.font      = _font(size=9, color=C_TEXT_DATA)
        c_fc.fill      = fill_data
        c_fc.alignment = aln_c
        c_fc.border    = borde

        # F — Empresa / Región
        empresa_region = " / ".join(
            v for v in [lic.get("comprador", ""), lic.get("region", "")]
            if v and v != "—"
        ) or "—"
        c_e = ws.cell(row=row_i, column=6, value=empresa_region)
        c_e.font      = _font(size=9, color=C_TEXT_MUTED)
        c_e.fill      = fill_data
        c_e.alignment = aln_i
        c_e.border    = borde

        # G — Link directo (hipervínculo)
        c_l = ws.cell(row=row_i, column=7)
        if lic.get("link"):
            c_l.value     = "Ver licitación →"
            c_l.hyperlink = lic["link"]
            c_l.font      = _font(size=9, color=C_LINK, underline="single")
        else:
            c_l.value = "Sin enlace"
            c_l.font  = _font(size=9, color=C_TEXT_MUTED)
        c_l.fill      = fill_data
        c_l.alignment = aln_c
        c_l.border    = borde

        # Altura dinámica según largo del título
        ws.row_dimensions[row_i].height = min(60, max(26, len(lic["descripcion"]) // 4))

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:{get_column_letter(NCOLS)}{3 + len(licitaciones)}"


def _hoja_resumen(ws, licitaciones: list, keywords: list,
                  horas_atras: int, nombre_archivo: str) -> None:
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 22
    borde = _borde_oscuro()

    # Título
    ws.merge_cells("A1:B1")
    c = ws["A1"]
    c.value     = "RESUMEN DE EJECUCIÓN — LicitaMonitor Soldesp"
    c.font      = _font(bold=True, size=12, color=C_TEXT_HEADER)
    c.fill      = _fill(C_HEADER_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    # Estadísticas globales
    portales = sorted({l["portal"] for l in licitaciones})
    stats = [
        ("Fecha de ejecución",            datetime.now().strftime("%d/%m/%Y %H:%M:%S")),
        ("Período analizado",             f"Últimas {horas_atras} horas"),
        ("─── Resultados globales ───",   ""),
        ("Total licitaciones encontradas", len(licitaciones)),
        ("Compradores distintos",          len({l["comprador"] for l in licitaciones})),
        ("Regiones distintas",             len({l["region"]    for l in licitaciones if l["region"] != "—"})),
        ("─── Por portal ───",            ""),
    ]
    for p in portales:
        cnt = sum(1 for l in licitaciones if l["portal"] == p)
        stats.append((f"  · {p}", cnt))

    stats += [
        ("─── Filtros aplicados ───",     ""),
        ("Keywords",                      " · ".join(kw.capitalize() for kw in keywords)),
        ("Archivo generado",              nombre_archivo),
    ]

    for row_i, (label, valor) in enumerate(stats, start=2):
        c_label = ws.cell(row=row_i, column=1, value=label)
        c_valor = ws.cell(row=row_i, column=2, value=valor)

        es_seccion = label.startswith("───")
        bg = C_SUBHEADER_BG if es_seccion else (C_ROW_EVEN if row_i % 2 == 0 else C_ROW_ODD)

        c_label.font = _font(bold=True, size=10,
                              color=(C_TEXT_MUTED if es_seccion else C_TEXT_DATA))
        c_valor.font = _font(size=10, color=C_TEXT_DATA)

        for c in (c_label, c_valor):
            c.fill      = _fill(bg)
            c.border    = borde
            c.alignment = Alignment(vertical="center", wrap_text=True)
        ws.row_dimensions[row_i].height = 22


def _color_portal(nombre: str) -> str:
    tabla = {
        "Portal Minero": C_BADGE_MINERO,
    }
    return tabla.get(nombre, "1B3A6B")


# ═══════════════════════════════════════════════════════════════
# NOTIFICACIÓN EMAIL
# ═══════════════════════════════════════════════════════════════

def enviar_email(cfg_notif: dict, filepath: str, n_lics: int, logger: logging.Logger) -> None:
    if not cfg_notif.get("activar_email", False) or not cfg_notif.get("destinatarios"):
        return
    try:
        msg            = MIMEMultipart()
        msg["From"]    = cfg_notif["smtp_usuario"]
        msg["To"]      = ", ".join(cfg_notif["destinatarios"])
        msg["Subject"] = (f"LicitaMonitor Soldesp — {n_lics} licitaciones "
                          f"({datetime.now().strftime('%d/%m/%Y')})")

        cuerpo = (
            f"Se encontraron <b>{n_lics}</b> licitaciones relevantes.<br>"
            f"Archivo adjunto: <code>{Path(filepath).name}</code><br><br>"
            f"<i>Generado automáticamente por LicitaMonitor Soldesp</i>"
        )
        msg.attach(MIMEText(cuerpo, "html", "utf-8"))

        with open(filepath, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition",
                        f'attachment; filename="{Path(filepath).name}"')
        msg.attach(part)

        with smtplib.SMTP(cfg_notif["smtp_servidor"], int(cfg_notif["smtp_puerto"])) as s:
            s.ehlo(); s.starttls()
            s.login(cfg_notif["smtp_usuario"], cfg_notif["smtp_password"])
            s.sendmail(cfg_notif["smtp_usuario"],
                       cfg_notif["destinatarios"], msg.as_string())

        logger.info("Email enviado a: %s", ", ".join(cfg_notif["destinatarios"]))
    except Exception as exc:
        logger.warning("Email no enviado: %s", exc)


# ═══════════════════════════════════════════════════════════════
# ORQUESTADOR PRINCIPAL
# ═══════════════════════════════════════════════════════════════

PORTAL_MAP = {
    "portal_minero": PortalMinero,
}

def ejecutar_proceso(
    logger: logging.Logger | None = None,
    portales_activos: list | None = None,
) -> tuple[str | None, int]:
    """
    Ejecuta scraping en todos los portales activos y genera un Excel unificado.
    Retorna (ruta_excel, total_licitaciones).
    """
    cfg     = cargar_config()
    log_dir = cfg["rutas"].get("log_dir", str(Path.home() / "Desktop" / "LicitaMonitor" / "logs"))
    if logger is None:
        logger = configurar_logging(log_dir)

    keywords    = cfg["filtros"]["keywords"]
    horas_atras = int(cfg["filtros"].get("horas_atras", 24))
    output_dir  = cfg["rutas"]["output_dir"]
    cfg_sel     = cfg.get("selenium", {})
    reintentos  = int(cfg_sel.get("max_reintentos", 3))

    # Determinar portales a ejecutar
    portales_cfg = cfg.get("portales", {})
    if portales_activos is None:
        portales_activos = [k for k, v in portales_cfg.items() if v.get("activo", True)]

    logger.info("=" * 60)
    logger.info("  LICITAMONITOR — %s", datetime.now().strftime("%d/%m/%Y %H:%M"))
    logger.info("  Portales activos: %s", ", ".join(portales_activos))
    logger.info("  Keywords: %s", ", ".join(keywords))
    logger.info("=" * 60)

    todas_licitaciones: list = []
    driver = None

    for nombre_portal in portales_activos:
        cls = PORTAL_MAP.get(nombre_portal)
        if not cls:
            logger.warning("Portal desconocido: %s — omitido.", nombre_portal)
            continue

        for intento in range(1, reintentos + 1):
            try:
                if intento > 1:
                    logger.info("Reintento %d/%d para %s...", intento, reintentos, nombre_portal)
                    time.sleep(5 * intento)

                driver  = iniciar_driver(cfg_sel)
                portal  = cls(driver, cfg, logger)
                portal.login()
                lics    = portal.extraer_licitaciones(keywords, horas_atras)
                todas_licitaciones.extend(lics)
                logger.info("[%s] %d licitaciones agregadas.", nombre_portal, len(lics))
                break

            except RuntimeError as exc:
                logger.error("Error en %s: %s", nombre_portal, exc)
            except WebDriverException as exc:
                logger.error("WebDriver error en %s: %s", nombre_portal, exc)
            except Exception as exc:
                logger.exception("Error inesperado en %s: %s", nombre_portal, exc)
            finally:
                if driver:
                    try:
                        driver.quit()
                    except Exception:
                        pass
                    driver = None

    archivo = None
    if todas_licitaciones or True:  # Generar Excel siempre (aunque esté vacío)
        archivo = generar_excel(todas_licitaciones, output_dir, keywords, horas_atras, logger)

    if archivo:
        logger.info("")
        logger.info("╔══════════════════════════════════════════════════════════╗")
        logger.info("║  PROCESO COMPLETADO                                      ║")
        logger.info("╠══════════════════════════════════════════════════════════╣")
        logger.info("║  Total licitaciones: %-36d║", len(todas_licitaciones))
        logger.info("║  Archivo: %-48s║", Path(archivo).name[:48])
        logger.info("╚══════════════════════════════════════════════════════════╝")
        enviar_email(cfg.get("notificaciones", {}), archivo, len(todas_licitaciones), logger)

    return archivo, len(todas_licitaciones)


def main() -> None:
    cfg     = cargar_config()
    log_dir = cfg["rutas"].get("log_dir", str(Path.home() / "Desktop" / "LicitaMonitor" / "logs"))
    logger  = configurar_logging(log_dir)
    ejecutar_proceso(logger)


if __name__ == "__main__":
    main()
