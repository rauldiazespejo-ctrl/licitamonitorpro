"""
LICITAMONITOR — Interfaz Gráfica
==================================
Monitor de licitaciones industriales — Portal Minero.

Uso:
    python LicitaMonitor_GUI.py           → abre la interfaz gráfica
    python LicitaMonitor_GUI.py --auto    → ejecución silenciosa (Programador de Tareas)
"""

import json
import logging
import math
import os
import queue
import subprocess
import sys
import threading
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, font as tkfont, messagebox, scrolledtext, ttk

sys.path.insert(0, str(Path(__file__).parent))

try:
    import LicitaMonitor as core
except ImportError as e:
    print(f"[ERROR] No se pudo importar LicitaMonitor.py: {e}")
    sys.exit(1)


# ═══════════════════════════════════════════════════════════════
# PALETA OSCURA — GUI
# ═══════════════════════════════════════════════════════════════

C_BG_PRINCIPAL  = "#0D1B2A"   # Fondo principal
C_BG_PANEL      = "#101E30"   # Paneles laterales
C_BG_CARD       = "#132233"   # Tarjetas / cards
C_BG_INPUT      = "#0A1624"   # Inputs
C_BORDE         = "#1A3550"   # Bordes

C_AZUL_VIV      = "#1B6FB5"   # Azul vivo — acento
C_AZUL_CLARO    = "#4FC3F7"   # Azul claro — links / resaltes
C_VERDE         = "#1A7A4A"   # Verde — éxito
C_ROJO          = "#9B2020"   # Rojo — error
C_AMARILLO      = "#B8860B"   # Amarillo oscuro — advertencia
C_BLANCO        = "#E8F0F8"   # Texto principal
C_GRIS          = "#5A7A99"   # Texto secundario

C_MINERO_BADGE  = "#6B1515"

FONT_TITULO    = ("Segoe UI", 14, "bold")
FONT_SUBTITULO = ("Segoe UI", 10, "bold")
FONT_NORMAL    = ("Segoe UI", 9)
FONT_LOG       = ("Consolas", 8)
FONT_BTN       = ("Segoe UI", 10, "bold")
FONT_SMALL     = ("Segoe UI", 8)


# ═══════════════════════════════════════════════════════════════
# SPLASH SCREEN
# ═══════════════════════════════════════════════════════════════

class SplashScreen:
    W, H        = 500, 360
    DURACION_MS = 2800
    FPS         = 30

    def __init__(self):
        self.root = tk.Tk()
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.update_idletasks()
        sx = (self.root.winfo_screenwidth()  - self.W) // 2
        sy = (self.root.winfo_screenheight() - self.H) // 2
        self.root.geometry(f"{self.W}x{self.H}+{sx}+{sy}")

        self.canvas = tk.Canvas(self.root, width=self.W, height=self.H,
                                bg="#0D1B2A", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        self._logo_img  = None
        self._frame     = 0
        self._total     = int(self.DURACION_MS / 1000 * self.FPS)

        self._dibujar_fondo()
        self._cargar_logo()
        self._animar()

    def _dibujar_fondo(self):
        c = self.canvas
        W, H = self.W, self.H
        pasos = 40
        for i in range(pasos):
            r = int(13  + (18  - 13)  * i / pasos)
            g = int(27  + (45  - 27)  * i / pasos)
            b = int(42  + (80  - 42)  * i / pasos)
            color = f"#{r:02x}{g:02x}{b:02x}"
            y0 = int(H * i / pasos)
            y1 = int(H * (i + 1) / pasos)
            c.create_rectangle(0, y0, W, y1, fill=color, outline="")

        # Borde
        c.create_rectangle(1, 1, W-2, H-2, outline="#1B6FB5", width=1)
        c.create_rectangle(3, 3, W-4, H-4, outline="#0F2A45", width=1)

        # Líneas decorativas
        col = "#0F2A45"
        for cx, cy in [(20, 20), (W-20, 20), (20, H-20), (W-20, H-20)]:
            dx = 1 if cx < W//2 else -1
            dy = 1 if cy < H//2 else -1
            c.create_line(cx, cy, cx + 40*dx, cy, fill=col, width=1)
            c.create_line(cx, cy, cx, cy + 40*dy, fill=col, width=1)
            c.create_oval(cx-3, cy-3, cx+3, cy+3, fill="#1B6FB5", outline="")

        # Textos fijos
        c.create_text(W//2, H-60, text="AGENCIA DE INTELIGENCIA ARTIFICIAL",
                      fill="#2A4A6A", font=("Segoe UI", 8, "bold"), anchor="center")
        c.create_text(W//2, H-42, text="Iniciando sistema...",
                      fill=C_GRIS, font=("Segoe UI", 8), anchor="center", tags="txt_estado")

    def _cargar_logo(self):
        c = self.canvas
        W = self.W
        base = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
        logo_path = base / "logo.png"

        if logo_path.exists():
            try:
                from PIL import Image, ImageTk
                img = Image.open(logo_path).resize((240, 170), Image.LANCZOS)
                self._logo_img = ImageTk.PhotoImage(img)
                c.create_image(W//2, 145, image=self._logo_img, anchor="center", tags="logo")
                return
            except Exception:
                pass

        # Fallback texto
        c.create_text(W//2 - 50, 105, text="LICITA", fill=C_BLANCO,
                      font=("Segoe UI", 40, "bold"), anchor="center", tags="logo")
        c.create_text(W//2 + 52, 105, text="MONITOR", fill="#4FC3F7",
                      font=("Segoe UI", 24, "bold"), anchor="center", tags="logo")
        c.create_text(W//2, 148, text="⬡  Portal Minero",
                      fill="#2A4A6A", font=("Segoe UI", 10), anchor="center", tags="logo")

    def _animar(self):
        c = self.canvas
        W, H = self.W, self.H
        f  = self._frame
        tf = self._total
        p  = f / tf  # 0→1

        BAR_X1, BAR_Y1 = 60, H - 28
        BAR_X2, BAR_Y2 = W - 60, H - 16
        c.delete("barra")
        c.create_rectangle(BAR_X1, BAR_Y1, BAR_X2, BAR_Y2,
                           fill="#091524", outline="#1A3550", tags="barra")
        fx = BAR_X1 + (BAR_X2 - BAR_X1) * p
        if fx > BAR_X1 + 2:
            c.create_rectangle(BAR_X1+1, BAR_Y1+1, fx, BAR_Y2-1,
                               fill="#1B6FB5", outline="", tags="barra")
            c.create_rectangle(max(fx-5, BAR_X1+1), BAR_Y1+1, fx, BAR_Y2-1,
                               fill="#4FC3F7", outline="", tags="barra")

        # Texto estado
        msgs = [(0.0,"Iniciando sistema..."), (0.3,"Cargando portales..."),
                (0.6,"Preparando interfaz..."), (0.85,"Listo.")]
        txt = msgs[0][1]
        for u, m in msgs:
            if p >= u:
                txt = m
        c.delete("txt_estado")
        c.create_text(W//2, H-42, text=txt, fill="#2A5A80",
                      font=("Segoe UI", 8), anchor="center", tags="txt_estado")

        # Fade
        alpha = min(p/0.12, 1.0) if p < 0.12 else (max((1-p)/0.12, 0.0) if p > 0.88 else 1.0)
        try:
            self.root.attributes("-alpha", round(alpha, 2))
        except Exception:
            pass

        self._frame += 1
        if self._frame <= tf:
            self.root.after(int(1000/self.FPS), self._animar)
        else:
            self.root.destroy()

    def mostrar(self):
        self.root.mainloop()


# ═══════════════════════════════════════════════════════════════
# LOG HANDLER → GUI QUEUE
# ═══════════════════════════════════════════════════════════════

class QueueHandler(logging.Handler):
    def __init__(self, q: queue.Queue):
        super().__init__()
        self.q = q
    def emit(self, record):
        self.q.put(self.format(record))


# ═══════════════════════════════════════════════════════════════
# APLICACIÓN PRINCIPAL
# ═══════════════════════════════════════════════════════════════

class LicitaMonitorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.config_data  = core.cargar_config()
        self.log_queue    = queue.Queue()
        self._ejecutando  = False
        self._ultimo_xlsx = None

        self._setup_window()
        self._build_ui()
        self._setup_logging()
        self._actualizar_estado_tarea()
        self._poll_logs()

    # ── Ventana ──────────────────────────────────────────────────────────────

    def _setup_window(self):
        self.title("LicitaMonitor — Soldesp")
        self.resizable(True, True)
        self.minsize(900, 640)
        self.configure(bg=C_BG_PRINCIPAL)
        self.update_idletasks()
        w, h = 1020, 720
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")
        ico = Path(__file__).parent / "icon.ico"
        if ico.exists():
            try:
                self.iconbitmap(str(ico))
            except Exception:
                pass

    # ── UI ───────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Header ──────────────────────────────────────────────────────────
        header = tk.Frame(self, bg="#091524", height=66)
        header.pack(fill="x")
        header.pack_propagate(False)

        tk.Label(header, text="  LICITA", bg="#091524",
                 fg=C_BLANCO, font=("Segoe UI", 18, "bold")).pack(side="left", pady=14, padx=(18, 0))
        tk.Label(header, text="MONITOR", bg="#091524",
                 fg=C_AZUL_CLARO, font=("Segoe UI", 18, "bold")).pack(side="left", pady=14)
        tk.Label(header, text="  ·  Soldesp  ·  Monitor de licitaciones industriales",
                 bg="#091524", fg="#2A4A6A", font=("Segoe UI", 9)).pack(side="left", pady=20)

        # Badge portal
        badge_frame = tk.Frame(header, bg="#091524")
        badge_frame.pack(side="right", padx=18, pady=14)
        self._badge(badge_frame, "Portal Minero", C_MINERO_BADGE)

        # ── Body ─────────────────────────────────────────────────────────────
        body = tk.Frame(self, bg=C_BG_PRINCIPAL)
        body.pack(fill="both", expand=True, padx=14, pady=10)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(0, weight=1)

        # Panel izquierdo
        left = tk.Frame(body, bg=C_BG_PRINCIPAL)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        self._card_portales(left)
        self._card_config(left)
        self._card_acciones(left)
        self._card_tarea(left)

        # Panel derecho — log
        right = tk.Frame(body, bg=C_BG_PRINCIPAL)
        right.grid(row=0, column=1, sticky="nsew")
        right.rowconfigure(1, weight=1)

        tk.Label(right, text="Registro de actividad", bg=C_BG_PRINCIPAL,
                 fg=C_BLANCO, font=FONT_SUBTITULO).grid(row=0, column=0, sticky="w", pady=(0,4))
        right.columnconfigure(0, weight=1)

        log_outer = tk.Frame(right, bg=C_BORDE, bd=1)
        log_outer.grid(row=1, column=0, sticky="nsew")

        self.txt_log = scrolledtext.ScrolledText(
            log_outer, wrap="word", state="disabled",
            bg="#071018", fg="#4FC3F7",
            insertbackground=C_BLANCO, font=FONT_LOG,
            relief="flat", bd=0, padx=8, pady=6,
            selectbackground=C_AZUL_VIV,
        )
        self.txt_log.pack(fill="both", expand=True)
        self.txt_log.tag_config("INFO",     foreground="#4FC3F7")
        self.txt_log.tag_config("WARNING",  foreground="#B8860B")
        self.txt_log.tag_config("ERROR",    foreground="#C0392B")
        self.txt_log.tag_config("CRITICAL", foreground="#C0392B")
        self.txt_log.tag_config("DEBUG",    foreground="#2A5A80")

        tk.Button(right, text="Limpiar log", bg=C_BG_PRINCIPAL, fg=C_GRIS,
                  font=FONT_SMALL, relief="flat", cursor="hand2",
                  command=self._limpiar_log).grid(row=2, column=0, sticky="e", pady=(4,0))

        # Status bar
        self.lbl_status = tk.Label(self, text="Listo", bg="#071018",
                                   fg=C_GRIS, font=FONT_SMALL, anchor="w", padx=12)
        self.lbl_status.pack(fill="x", side="bottom")

    def _badge(self, parent, texto, color):
        tk.Label(parent, text=f"  {texto}  ", bg=color,
                 fg=C_BLANCO, font=("Segoe UI", 8, "bold"),
                 relief="flat", padx=6, pady=2).pack(side="left", padx=3)

    def _card(self, parent, titulo, color_titulo=C_AZUL_VIV):
        frame = tk.LabelFrame(
            parent, text=f"  {titulo}  ", bg=C_BG_CARD,
            fg=color_titulo, font=FONT_SUBTITULO,
            relief="flat", bd=1,
            highlightthickness=1, highlightbackground=C_BORDE,
        )
        frame.pack(fill="x", pady=(0, 8))
        return frame

    def _card_portales(self, parent):
        card = self._card(parent, "Portales activos")

        portales_cfg = self.config_data.get("portales", {})
        self.var_minero = tk.BooleanVar(
            value=portales_cfg.get("portal_minero", {}).get("activo", True))

        row = tk.Frame(card, bg=C_BG_CARD)
        row.pack(fill="x", padx=10, pady=8)

        cb_style = {"bg": C_BG_CARD, "fg": C_BLANCO, "font": FONT_NORMAL,
                    "activebackground": C_BG_CARD, "activeforeground": C_AZUL_CLARO,
                    "selectcolor": C_BG_INPUT, "relief": "flat", "cursor": "hand2"}

        tk.Checkbutton(row, text="  Portal Minero", variable=self.var_minero,
                       **cb_style).pack(side="left")

    def _card_config(self, parent):
        card = self._card(parent, "Configuración")

        # Carpeta salida
        tk.Label(card, text="Carpeta de salida:", bg=C_BG_CARD,
                 fg=C_BLANCO, font=FONT_NORMAL).pack(anchor="w", padx=10, pady=(8,2))
        row_dir = tk.Frame(card, bg=C_BG_CARD)
        row_dir.pack(fill="x", padx=10, pady=(0,6))

        self.var_output = tk.StringVar(
            value=self.config_data["rutas"].get("output_dir",
                  str(Path.home() / "Desktop" / "LicitaMonitor")))
        entry = tk.Entry(row_dir, textvariable=self.var_output, font=FONT_NORMAL,
                         bg=C_BG_INPUT, fg=C_BLANCO, relief="flat", bd=1,
                         insertbackground=C_AZUL_CLARO,
                         highlightthickness=1, highlightbackground=C_BORDE)
        entry.pack(side="left", fill="x", expand=True, ipady=4)
        tk.Button(row_dir, text="📁", bg=C_AZUL_VIV, fg=C_BLANCO,
                  font=("Segoe UI", 10), relief="flat", cursor="hand2",
                  width=3, command=self._elegir_carpeta).pack(side="left", padx=(4,0))

        # Período de búsqueda
        tk.Label(card, text="Buscar licitaciones del:",
                 bg=C_BG_CARD, fg=C_BLANCO, font=FONT_NORMAL).pack(anchor="w", padx=10, pady=(0,2))

        # Mapa: etiqueta → horas
        self.PERIODOS = {
            "Último día   (24 h)":      24,
            "Últimos 2 días  (48 h)":   48,
            "Última semana  (7 días)":  168,
            "Última quincena (15 días)": 360,
            "Último mes  (30 días)":    720,
            "Últimos 2 meses (60 días)": 1440,
            "Últimos 3 meses (90 días)": 2160,
        }
        horas_actuales = self.config_data["filtros"].get("horas_atras", 24)
        # Seleccionar la etiqueta que coincida con las horas guardadas
        etiqueta_default = next(
            (k for k, v in self.PERIODOS.items() if v == horas_actuales),
            "Último día   (24 h)"
        )
        self.var_periodo = tk.StringVar(value=etiqueta_default)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Dark.TCombobox",
                        fieldbackground=C_BG_INPUT,
                        background=C_BG_INPUT,
                        foreground=C_BLANCO,
                        selectbackground=C_AZUL_VIV,
                        selectforeground=C_BLANCO,
                        arrowcolor=C_AZUL_CLARO,
                        bordercolor=C_BORDE,
                        darkcolor=C_BG_INPUT,
                        lightcolor=C_BG_INPUT)
        style.map("Dark.TCombobox",
                  fieldbackground=[("readonly", C_BG_INPUT)],
                  foreground=[("readonly", C_BLANCO)],
                  background=[("readonly", C_BG_INPUT)])

        combo = ttk.Combobox(
            card,
            textvariable=self.var_periodo,
            values=list(self.PERIODOS.keys()),
            state="readonly",
            font=FONT_NORMAL,
            style="Dark.TCombobox",
        )
        combo.pack(fill="x", padx=10, pady=(0,6), ipady=4)

        tk.Button(card, text="Guardar configuración",
                  bg=C_AZUL_VIV, fg=C_BLANCO, font=FONT_SMALL,
                  relief="flat", cursor="hand2", padx=12, pady=4,
                  command=self._guardar_config).pack(anchor="e", padx=10, pady=(0,8))

    def _card_acciones(self, parent):
        card = self._card(parent, "Ejecución")

        self.btn_ejecutar = tk.Button(
            card, text="▶  Ejecutar Ahora",
            bg=C_VERDE, fg=C_BLANCO, font=FONT_BTN,
            relief="flat", cursor="hand2", padx=16, pady=10,
            activebackground="#115c33", activeforeground=C_BLANCO,
            command=self._ejecutar,
        )
        self.btn_ejecutar.pack(fill="x", padx=10, pady=(10,4))

        self.btn_abrir = tk.Button(
            card, text="📂  Abrir último Excel generado",
            bg=C_BG_INPUT, fg=C_GRIS, font=FONT_NORMAL,
            relief="flat", cursor="hand2", padx=12, pady=6,
            state="disabled", command=self._abrir_excel,
        )
        self.btn_abrir.pack(fill="x", padx=10, pady=(0,6))

        self.lbl_ultimo = tk.Label(card, text="Sin ejecuciones en esta sesión.",
                                   bg=C_BG_CARD, fg=C_GRIS, font=FONT_SMALL,
                                   wraplength=280, justify="left")
        self.lbl_ultimo.pack(anchor="w", padx=10, pady=(0,8))

    def _card_tarea(self, parent):
        card = self._card(parent, "Programador de Tareas (Windows)")

        self.lbl_tarea = tk.Label(card, text="Verificando...",
                                  bg=C_BG_CARD, fg=C_GRIS, font=FONT_NORMAL)
        self.lbl_tarea.pack(anchor="w", padx=10, pady=(8,2))

        row = tk.Frame(card, bg=C_BG_CARD)
        row.pack(fill="x", padx=10, pady=(0,6))

        tk.Button(row, text="Instalar tarea (15:00)",
                  bg="#0F2A45", fg=C_BLANCO, font=FONT_SMALL,
                  relief="flat", cursor="hand2", padx=10, pady=5,
                  command=self._instalar_tarea).pack(side="left", padx=(0,6))
        tk.Button(row, text="Eliminar",
                  bg=C_BG_INPUT, fg=C_GRIS, font=FONT_SMALL,
                  relief="flat", cursor="hand2", padx=10, pady=5,
                  command=self._eliminar_tarea).pack(side="left")

        tk.Label(card, text="Ejecuta silenciosamente cada día a las 15:00.",
                 bg=C_BG_CARD, fg="#2A4A6A", font=("Segoe UI", 7),
                 wraplength=280, justify="left").pack(anchor="w", padx=10, pady=(0,8))

    # ── Logging ──────────────────────────────────────────────────────────────

    def _setup_logging(self):
        log_dir = self.config_data["rutas"].get(
            "log_dir", str(Path(self.var_output.get()) / "logs"))
        self.logger = core.configurar_logging(log_dir)
        qh = QueueHandler(self.log_queue)
        qh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S"))
        self.logger.addHandler(qh)

    def _poll_logs(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self._append_log(msg)
        except queue.Empty:
            pass
        self.after(150, self._poll_logs)

    def _append_log(self, texto: str):
        self.txt_log.config(state="normal")
        tag = "INFO"
        for n in ("ERROR", "CRITICAL", "WARNING", "DEBUG"):
            if f"[{n}]" in texto:
                tag = n
                break
        self.txt_log.insert("end", texto + "\n", tag)
        self.txt_log.see("end")
        self.txt_log.config(state="disabled")

    def _limpiar_log(self):
        self.txt_log.config(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.config(state="disabled")

    # ── Acciones ─────────────────────────────────────────────────────────────

    def _elegir_carpeta(self):
        carpeta = filedialog.askdirectory(
            title="Seleccionar carpeta de salida",
            initialdir=self.var_output.get() if Path(self.var_output.get()).exists()
                       else str(Path.home()),
        )
        if carpeta:
            self.var_output.set(carpeta)

    def _guardar_config(self):
        self.config_data["rutas"]["output_dir"] = self.var_output.get()
        self.config_data["rutas"]["log_dir"]    = str(Path(self.var_output.get()) / "logs")
        self.config_data["filtros"]["horas_atras"] = self.PERIODOS.get(
            self.var_periodo.get(), 24
        )

        # Sincronizar portales activos
        if "portales" not in self.config_data:
            self.config_data["portales"] = {}
        if "portal_minero" not in self.config_data["portales"]:
            self.config_data["portales"]["portal_minero"] = {}
        self.config_data["portales"]["portal_minero"]["activo"] = self.var_minero.get()

        base = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
        config_path = base / "config.json"
        try:
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(self.config_data, f, ensure_ascii=False, indent=2)
            self.logger.info("Configuración guardada en %s", config_path)
            self._set_status("Configuración guardada.", C_VERDE)
        except Exception as exc:
            self.logger.error("Error guardando config: %s", exc)
            messagebox.showerror("Error", f"No se pudo guardar la configuración:\n{exc}")

    def _ejecutar(self):
        if self._ejecutando:
            return

        portales = []
        if self.var_minero.get():
            portales.append("portal_minero")

        if not portales:
            messagebox.showwarning("Aviso", "Selecciona al menos un portal antes de ejecutar.")
            return

        self._guardar_config()
        self._ejecutando = True
        self.btn_ejecutar.config(state="disabled", text="⏳  Ejecutando...", bg="#333333")
        self._set_status("Ejecutando — por favor espera...", C_AMARILLO)
        self.logger.info("─" * 55)
        self.logger.info("Ejecución manual — portales: %s | período: %s",
                         ", ".join(portales), self.var_periodo.get().strip())
        threading.Thread(target=self._hilo_ejecucion,
                         args=(portales,), daemon=True).start()

    def _hilo_ejecucion(self, portales: list):
        archivo = None
        n_lics  = 0
        try:
            archivo, n_lics = core.ejecutar_proceso(
                logger=self.logger,
                portales_activos=portales,
            )
        except Exception as exc:
            self.logger.exception("Error inesperado: %s", exc)
            self.after(0, lambda: self._fin_ejecucion(None, 0, str(exc)))
            return
        self.after(0, lambda: self._fin_ejecucion(archivo, n_lics))

    def _fin_ejecucion(self, archivo, n_lics, error=None):
        self._ejecutando  = False
        self._ultimo_xlsx = archivo
        self.btn_ejecutar.config(state="normal", text="▶  Ejecutar Ahora", bg=C_VERDE)

        if error:
            self._set_status(f"Error: {error[:80]}", C_ROJO)
            self.lbl_ultimo.config(text=f"Error: {error[:100]}", fg=C_ROJO)
        else:
            ts = datetime.now().strftime("%H:%M:%S")
            self.lbl_ultimo.config(
                text=(f"Última ejecución: {ts}\n"
                      f"Licitaciones encontradas: {n_lics}\n"
                      f"Archivo: {Path(archivo).name if archivo else '—'}"),
                fg=C_AZUL_CLARO if n_lics > 0 else C_GRIS,
            )
            self._set_status(f"Completado — {n_lics} licitaciones encontradas.", C_VERDE)
            if archivo:
                self.btn_abrir.config(state="normal", bg=C_AZUL_VIV, fg=C_BLANCO)

    def _abrir_excel(self):
        if self._ultimo_xlsx and Path(self._ultimo_xlsx).exists():
            try:
                if sys.platform == "win32":
                    os.startfile(self._ultimo_xlsx)
                elif sys.platform == "darwin":
                    subprocess.run(["open", self._ultimo_xlsx], check=False)
                else:
                    subprocess.run(["xdg-open", self._ultimo_xlsx], check=False)
            except Exception as exc:
                messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{exc}")
        else:
            messagebox.showwarning("Aviso", "El archivo no existe en la ruta indicada.")

    # ── Tarea Programada ─────────────────────────────────────────────────────

    def _actualizar_estado_tarea(self):
        if sys.platform != "win32":
            self.lbl_tarea.config(text="⚠ Solo disponible en Windows", fg=C_AMARILLO)
            return
        try:
            r = subprocess.run(["schtasks", "/Query", "/TN", "LicitaMonitor Soldesp"],
                               capture_output=True, text=True, timeout=8)
            if r.returncode == 0:
                self.lbl_tarea.config(text="✔ Tarea activa (diario 15:00)", fg=C_VERDE)
            else:
                self.lbl_tarea.config(text="✘ Tarea NO configurada", fg=C_ROJO)
        except Exception:
            self.lbl_tarea.config(text="⚠ No se pudo verificar", fg=C_AMARILLO)

    def _instalar_tarea(self):
        if sys.platform != "win32":
            messagebox.showinfo("Solo Windows",
                                "El Programador de Tareas solo está disponible en Windows.")
            return
        exe_path = sys.executable if getattr(sys, "frozen", False) \
                   else f'"{sys.executable}" "{Path(__file__).resolve()}" --auto'
        cmd = ["schtasks", "/Create",
               "/TN", "LicitaMonitor Soldesp",
               "/SC", "DAILY", "/ST", "15:00",
               "/TR", exe_path if getattr(sys, "frozen", False)
                     else f'{sys.executable} "{Path(__file__).resolve()}" --auto',
               "/RL", "HIGHEST", "/F"]
        try:
            r = subprocess.run(cmd, capture_output=True, text=True, timeout=15)
            if r.returncode == 0:
                self.lbl_tarea.config(text="✔ Tarea activa (diario 15:00)", fg=C_VERDE)
                messagebox.showinfo("Tarea instalada",
                                    "LicitaMonitor se ejecutará automáticamente cada día a las 15:00.")
            else:
                raise RuntimeError(r.stderr or r.stdout)
        except Exception as exc:
            messagebox.showerror("Error",
                                 f"No se pudo instalar la tarea:\n{exc}\n\n"
                                 "Intenta ejecutar como Administrador.")

    def _eliminar_tarea(self):
        if sys.platform != "win32":
            return
        if not messagebox.askyesno("Confirmar", "¿Eliminar la tarea programada?"):
            return
        try:
            r = subprocess.run(
                ["schtasks", "/Delete", "/TN", "LicitaMonitor Soldesp", "/F"],
                capture_output=True, text=True, timeout=10)
            if r.returncode == 0:
                self.lbl_tarea.config(text="✘ Tarea NO configurada", fg=C_ROJO)
                messagebox.showinfo("Eliminada", "Tarea programada eliminada.")
            else:
                raise RuntimeError(r.stderr)
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo eliminar la tarea:\n{exc}")

    # ── Status bar ───────────────────────────────────────────────────────────

    def _set_status(self, texto: str, color: str = "#2A4A6A"):
        self.lbl_status.config(text=f"  {texto}", fg=color)


# ═══════════════════════════════════════════════════════════════
# PUNTO DE ENTRADA
# ═══════════════════════════════════════════════════════════════

def main():
    if "--auto" in sys.argv:
        cfg     = core.cargar_config()
        log_dir = cfg["rutas"].get("log_dir", str(Path.home() / "Desktop" / "LicitaMonitor" / "logs"))
        logger  = core.configurar_logging(log_dir)
        logger.info("=" * 60)
        logger.info("  LICITAMONITOR (modo automático) — %s",
                    datetime.now().strftime("%d/%m/%Y %H:%M"))
        logger.info("=" * 60)
        core.ejecutar_proceso(logger=logger)
        return

    SplashScreen().mostrar()

    app = LicitaMonitorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
