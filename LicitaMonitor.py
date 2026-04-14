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
    txt = texto.lower()
    for kw in keywords:
        kw_l = kw.lower()
        if len(kw_l) <= 2:
            if re.search(r'\b' + re.escape(kw_l) + r'\b', txt):
                return True
        else:
            if kw_l in txt:
                return True
    return False


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
# Login:   https://www.portalminero.com/login.action  (Confluence Seraph)
# Fuente 1: /wp/oportunidades-de-negocios/            (WordPress)
# Fuente 2: /display/acce/Muro+de+Actividades         (Confluence)
# ═══════════════════════════════════════════════════════════════

class PortalMinero(PortalBase):
    nombre_portal = "Portal Minero"
    color_badge   = C_BADGE_MINERO
    tiene_fechas  = False   # ninguna de las dos fuentes tiene columna de fecha fiable

    LOGIN_URL  = "https://www.portalminero.com/login.action"
    BASE_URL   = "https://www.portalminero.com"

    # Las dos fuentes de licitaciones a scrapear
    FUENTES = [
        "https://www.portalminero.com/wp/oportunidades-de-negocios/",
        "https://www.portalminero.com/display/acce/Muro+de+Actividades",
    ]

    # ──────────────────────────────────────────────────────────────
    # LOGIN
    # ──────────────────────────────────────────────────────────────

    def login(self) -> None:
        self.logger.info("[Portal Minero] Iniciando sesión...")
        cfg_portal = self.cfg["portales"]["portal_minero"]
        usuario    = cfg_portal["usuario"]
        password   = cfg_portal["password"]

        self.driver.get(self.LOGIN_URL)
        wait = WebDriverWait(self.driver, self.timeout)

        # Esperar formulario Seraph (sin comas en selector — C1)
        found = False
        for sel in ["#os_username", "input[name='os_username']", "#username"]:
            try:
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, sel)))
                found = True
                break
            except TimeoutException:
                continue
        if not found:
            _guardar_screenshot(self.driver, "pm_error_login_page.png", self.logger)
            raise RuntimeError("[Portal Minero] Página de login no cargó en %ds." % self.timeout)

        CAMPOS_USER = [
            (By.ID,   "os_username"),
            (By.NAME, "os_username"),
            (By.ID,   "username"),
            (By.NAME, "username"),
            (By.CSS_SELECTOR, "input[type='text']:not([disabled])"),
        ]
        CAMPOS_PASS = [
            (By.ID,   "os_password"),
            (By.NAME, "os_password"),
            (By.ID,   "password"),
            (By.NAME, "password"),
            (By.CSS_SELECTOR, "input[type='password']:not([disabled])"),
        ]
        CAMPOS_BTN = [
            (By.ID,   "loginButton"),
            (By.CSS_SELECTOR, "input#loginButton"),
            (By.CSS_SELECTOR, "button#loginButton"),
            (By.CSS_SELECTOR, "input[type='submit']"),
            (By.CSS_SELECTOR, "button[type='submit']"),
            (By.XPATH, "//input[@value='Iniciar sesión']"),
            (By.XPATH, "//input[@value='Log In']"),
        ]

        campo_user = self._encontrar_elemento(CAMPOS_USER)
        campo_pass = self._encontrar_elemento(CAMPOS_PASS)

        if not campo_user or not campo_pass:
            _guardar_screenshot(self.driver, "pm_error_campos.png", self.logger)
            raise RuntimeError("[Portal Minero] No se encontraron los campos usuario/contraseña.")

        self.driver.execute_script("arguments[0].value = '';", campo_user)
        campo_user.send_keys(usuario)
        self.driver.execute_script("arguments[0].value = '';", campo_pass)
        campo_pass.send_keys(password)

        btn = self._encontrar_elemento(CAMPOS_BTN)
        if btn:
            self.driver.execute_script("arguments[0].click();", btn)
        else:
            campo_pass.submit()

        # Detectar login exitoso — Seraph redirige fuera de /login.action
        try:
            wait.until(lambda d: (
                "login.action" not in d.current_url
                and "login" not in d.current_url.lower().split("?")[0]
            ))
            self.logger.info("[Portal Minero] Login exitoso → %s", self.driver.current_url)
        except TimeoutException:
            try:
                err = self.driver.find_element(
                    By.CSS_SELECTOR,
                    ".aui-message-error, #loginErrorDiv, .error, [class*='error-message']"
                )
                msg = err.text.strip()
            except Exception:
                msg = "sin mensaje"
            _guardar_screenshot(self.driver, "pm_error_redirect.png", self.logger)
            raise RuntimeError(
                f"[Portal Minero] Login fallido — {msg}. "
                "Verifica usuario y contraseña en config.json."
            )

    # ──────────────────────────────────────────────────────────────
    # EXTRACCIÓN — itera sobre las dos fuentes
    # ──────────────────────────────────────────────────────────────

    def extraer_licitaciones(self, keywords: list, horas_atras: int) -> list:
        resultados: list = []
        vistos: set      = set()

        for url in self.FUENTES:
            self.logger.info("[Portal Minero] → %s", url)
            try:
                parciales = self._scrapear_fuente(url, keywords, vistos)
                resultados.extend(parciales)
                self.logger.info("[Portal Minero] %d resultados en %s", len(parciales), url)
            except Exception as exc:
                self.logger.warning("[Portal Minero] Error en %s: %s", url, exc)
                _guardar_screenshot(self.driver, f"pm_error_{url.split('/')[-1][:20]}.png",
                                    self.logger)

        self.logger.info("[Portal Minero] Total: %d licitaciones relevantes.", len(resultados))
        return resultados

    def _scrapear_fuente(self, url: str, keywords: list, vistos: set) -> list:
        """Scrapea una URL concreta, paginando hasta agotar resultados."""
        self.driver.get(url)
        wait    = WebDriverWait(self.driver, self.timeout)
        pagina  = 1
        parcial = []

        # Esperar que cargue algo útil
        time.sleep(2)
        for sel in ["#main-content", "#content", ".entry-content",
                    "table", ".wiki-content", "article"]:
            try:
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, sel)))
                break
            except TimeoutException:
                continue

        while True:
            self.logger.debug("[Portal Minero] Fuente %s — página %d", url, pagina)
            items = self._obtener_items_pagina()

            if not items:
                self.logger.info("[Portal Minero] Sin ítems en página %d de %s", pagina, url)
                break

            self.logger.debug("[Portal Minero] %d ítems detectados.", len(items))

            for item in items:
                datos = self._parsear_item(item)
                if not datos:
                    continue
                texto_busqueda = (
                    f"{datos['descripcion']} {datos['comprador']} {datos['region']}"
                )
                if not _es_relevante(texto_busqueda, keywords):
                    continue

                clave = datos["descripcion"][:60]
                if clave in vistos:
                    continue
                vistos.add(clave)
                parcial.append(datos)

            siguiente = self._siguiente_pagina(wait)
            if not siguiente:
                break
            try:
                self.driver.execute_script("arguments[0].click();", siguiente)
                time.sleep(self.pausa)
                pagina += 1
            except Exception as exc:
                self.logger.warning("[Portal Minero] Error paginando: %s", exc)
                break

        return parcial

    # ──────────────────────────────────────────────────────────────
    # OBTENER ÍTEMS DE LA PÁGINA ACTUAL
    # ──────────────────────────────────────────────────────────────

    def _obtener_items_pagina(self) -> list:
        """
        Intenta múltiples selectores adaptados a las dos fuentes:
        - WordPress (/wp/...): divs con clase post/entry/card o filas de tabla
        - Confluence (/display/...): tablas .confluenceTable o #main-content
        """
        SELECTORES = [
            # ── WordPress ──────────────────────────────────────
            ".entry-content table tbody tr",
            "article",
            ".post",
            ".oportunidad",
            ".proyecto",
            "ul.oportunidades li",
            "ul.proyectos li",
            # ── Confluence ─────────────────────────────────────
            "table.confluenceTable tbody tr",
            "#main-content table tbody tr",
            ".wiki-content table tbody tr",
            "#content table tbody tr",
            # ── Genérico ───────────────────────────────────────
            "table tbody tr",
        ]
        for sel in SELECTORES:
            items = self.driver.find_elements(By.CSS_SELECTOR, sel)
            items = [i for i in items if i.text.strip()]
            if items:
                self.logger.debug("[Portal Minero] Selector: %s → %d ítems", sel, len(items))
                return items
        return []

    # ──────────────────────────────────────────────────────────────
    # PARSEAR ÍTEM INDIVIDUAL
    # ──────────────────────────────────────────────────────────────

    def _parsear_item(self, item) -> dict | None:
        try:
            texto = item.text.strip()
            if not texto or len(texto) < 4:
                return None

            tag = item.tag_name.lower()

            # ── Si es un div/article (WordPress) ─────────────────────────
            if tag in ("article", "div", "li", "section"):
                desc = texto[:200]
                link = ""
                try:
                    a    = item.find_element(By.CSS_SELECTOR, "a[href]")
                    href = a.get_attribute("href") or ""
                    if href and not href.startswith("javascript"):
                        link = href if href.startswith("http") else self.BASE_URL + href
                    # Preferir el texto del enlace como descripción si es más corto y útil
                    titulo = a.text.strip()
                    if titulo and len(titulo) > 5:
                        desc = titulo
                except NoSuchElementException:
                    pass

                # Intentar extraer región/país de sub-elementos
                region = ""
                for cls in [".region", ".pais", ".ubicacion", ".location"]:
                    try:
                        region = item.find_element(By.CSS_SELECTOR, cls).text.strip()
                        break
                    except NoSuchElementException:
                        pass

                return {
                    "portal":       self.nombre_portal,
                    "id":           "—",
                    "comprador":    "—",
                    "descripcion":  desc,
                    "region":       region or "—",
                    "fecha_pub":    "—",
                    "fecha_cierre": "—",
                    "link":         link,
                }

            # ── Si es una fila de tabla (Confluence / tabla WP) ───────────
            celdas = item.find_elements(By.TAG_NAME, "td")
            n = len(celdas)

            if n == 0:
                return None  # <tr> con <th> solamente — cabecera

            # Estructura real de Portal Minero: Nombre | País | Región | Ver Detalles
            desc   = celdas[0].text.strip() if n > 0 else texto[:120]
            pais   = celdas[1].text.strip() if n > 1 else "—"
            region = celdas[2].text.strip() if n > 2 else "—"

            # Link en col 0 (nombre) o col 3 (Ver Detalles)
            link = ""
            for sel in ["a[href*='/display/']", "a[href*='portalminero']",
                        "a[href*='/wp/']", "a[href]"]:
                try:
                    el   = item.find_element(By.CSS_SELECTOR, sel)
                    href = el.get_attribute("href") or ""
                    if href and not href.startswith("javascript"):
                        link = href if href.startswith("http") else self.BASE_URL + href
                        break
                except NoSuchElementException:
                    continue

            return {
                "portal":       self.nombre_portal,
                "id":           "—",
                "comprador":    pais,
                "descripcion":  desc,
                "region":       region,
                "fecha_pub":    "—",
                "fecha_cierre": "—",
                "link":         link,
            }

        except StaleElementReferenceException:
            return None
        except Exception as exc:
            self.logger.debug("[Portal Minero] Error parseando item: %s", exc)
            return None

    # ──────────────────────────────────────────────────────────────
    # PAGINACIÓN
    # ──────────────────────────────────────────────────────────────

    def _siguiente_pagina(self, wait) -> object | None:
        """
        Soporta:
        1. javascript:pagina(N,1) — estilo Confluence Portal Minero
        2. Selectores estándar de paginación (Next, siguiente, etc.)
        3. WordPress next page links
        """
        # 1. Patrón javascript:pagina(N,1)
        try:
            enlaces = self.driver.find_elements(By.CSS_SELECTOR, "a")
            pagina_actual  = None
            proximo        = None

            for a in enlaces:
                for fuente in [a.get_attribute("href") or "", a.get_attribute("onclick") or ""]:
                    m = re.search(r'pagina\((\d+)\s*,\s*1\)', fuente)
                    if m:
                        clases = (a.get_attribute("class") or "").lower()
                        num = int(m.group(1))
                        if "active" in clases or "current" in clases or "selected" in clases:
                            pagina_actual = num
                        else:
                            if proximo is None or num > proximo[0]:
                                proximo = (num, a)

            if proximo:
                num, el = proximo
                if pagina_actual is not None and num != pagina_actual + 1:
                    # Buscar exactamente pagina_actual + 1
                    for a in enlaces:
                        for fuente in [a.get_attribute("href") or "",
                                       a.get_attribute("onclick") or ""]:
                            m = re.search(r'pagina\((\d+)\s*,\s*1\)', fuente)
                            if m and int(m.group(1)) == pagina_actual + 1:
                                return a
                    return None
                return el
        except Exception as exc:
            self.logger.debug("[Portal Minero] paginación JS: %s", exc)

        # 2. Selectores estándar
        for selector in [
            "a.next",
            "a[aria-label='Next']",
            "a[rel='next']",
            ".nav-previous a",
            ".pagination li:not(.disabled) a.next",
            "a.siguiente",
            "a[title*='siguiente' i]",
            "a[title*='next' i]",
            "li.next:not(.disabled) a",
        ]:
            try:
                el = self.driver.find_element(By.CSS_SELECTOR, selector)
                if el.is_displayed() and el.is_enabled():
                    return el
            except NoSuchElementException:
                continue
        return None

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

    borde   = _borde_oscuro()
    aln_c   = Alignment(horizontal="center", vertical="center", wrap_text=True)
    aln_i   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    # ── Fila 1: Título principal ────────────────────────────────────────
    ws.merge_cells("A1:I1")
    c = ws["A1"]
    c.value     = (f"LICITAMONITOR — SOLDESP  ·  "
                   f"{datetime.now().strftime('%d/%m/%Y %H:%M')}  ·  "
                   f"Período: últimas {horas_atras}h")
    c.font      = _font(bold=True, size=13, color=C_TEXT_HEADER)
    c.fill      = _fill(C_HEADER_BG)
    c.alignment = aln_c
    ws.row_dimensions[1].height = 34

    # ── Fila 2: Subtítulo / filtros ─────────────────────────────────────
    ws.merge_cells("A2:I2")
    c2 = ws["A2"]
    if subtitulo:
        c2.value = subtitulo
    else:
        c2.value = "Filtros: " + " · ".join(kw.capitalize() for kw in keywords)
    c2.font      = _font(italic=True, size=9, color=C_TEXT_MUTED)
    c2.fill      = _fill(C_SUBHEADER_BG)
    c2.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 18

    # ── Fila 3: Cabeceras ────────────────────────────────────────────────
    headers = ["Portal", "ID / N°", "Comprador / Mandante",
               "Descripción / Servicio", "Región / Zona",
               "Fecha Publicación", "Fecha Cierre", "Link Directo", "Estado"]
    anchos  = [18,       12,          28,
               50,                      16,
               20,                 20,              24,              12]

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
        ws.merge_cells("A4:I4")
        c = ws["A4"]
        c.value     = f"No se encontraron licitaciones en las últimas {horas_atras} horas."
        c.font      = _font(italic=True, color=C_TEXT_MUTED)
        c.fill      = _fill(C_ROW_ODD)
        c.alignment = aln_c
        ws.row_dimensions[4].height = 28
    else:
        portales_colores = {}
        for row_i, lic in enumerate(licitaciones, start=4):
            bg = C_ROW_EVEN if row_i % 2 == 0 else C_ROW_ODD
            fill_data = _fill(bg)

            # Celda portal con color de badge
            portal_color = _color_portal(lic["portal"])
            c_portal = ws.cell(row=row_i, column=1, value=lic["portal"])
            c_portal.font      = _font(bold=True, size=9, color=C_TEXT_HEADER)
            c_portal.fill      = _fill(portal_color)
            c_portal.alignment = aln_c
            c_portal.border    = borde

            # Resto de columnas
            valores = [
                lic["id"],
                lic["comprador"],
                lic["descripcion"],
                lic["region"],
                lic["fecha_pub"],
                lic["fecha_cierre"],
                None,            # placeholder Link
                "Vigente",
            ]

            for col_i, valor in enumerate(valores, start=2):
                c = ws.cell(row=row_i, column=col_i, value=valor)
                c.font      = _font(size=9, color=C_TEXT_DATA)
                c.fill      = fill_data
                c.border    = borde
                c.alignment = aln_c if col_i in (2, 5, 6, 7, 9) else aln_i

            # Hipervínculo en columna 8
            lc = ws.cell(row=row_i, column=8)
            if lic["link"]:
                lc.value     = "Ver licitación →"
                lc.hyperlink = lic["link"]
                lc.font      = _font(size=9, color=C_LINK, underline="single")
            else:
                lc.value = "Sin enlace"
                lc.font  = _font(size=9, color=C_TEXT_MUTED)
            lc.fill      = fill_data
            lc.border    = borde
            lc.alignment = aln_c

            ws.row_dimensions[row_i].height = 26

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:I{3 + len(licitaciones)}"


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
