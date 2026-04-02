#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
RestoreGuard Pro - Crea puntos de restauración del sistema Windows
y respalda configuraciones reales a C: y D:
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
import subprocess, json, os, sys, threading, time, shutil, ctypes, winreg
from datetime import datetime
from pathlib import Path
import queue

try:
    import schedule; SCHEDULE_OK = True
except: SCHEDULE_OK = False

try:
    from PIL import Image, ImageDraw; PIL_OK = True
except: PIL_OK = False

try:
    import pystray; TRAY_OK = True
except: TRAY_OK = False

# ─────────────────────────────────────────────────────────────────
# RUTAS Y CONSTANTES ─ DETECCIÓN AUTOMÁTICA (UNIVERSAL)
# ─────────────────────────────────────────────────────────────────
APP_NAME     = "RestoreGuard Pro"
PREFIX       = "RestoreGuard"

# Carpeta principal en C: → funciona para cualquier usuario de Windows
BASE_DIR     = Path.home() / APP_NAME
CONFIG_FILE  = BASE_DIR / "config.json"
HISTORY_FILE = BASE_DIR / "historial.json"
LOG_FILE     = BASE_DIR / "restore_log.txt"

# Disco D: → se detecta automáticamente. Si no existe se usa carpeta en C:
_DRIVE_D = Path("D:/")
HAS_DRIVE_D  = _DRIVE_D.exists()
BACKUP_D     = _DRIVE_D / APP_NAME if HAS_DRIVE_D else BASE_DIR / "respaldo_secundario"

DEFAULT_CFG = {
    "auto_time": "08:00",
    "auto_enabled": True,
    "startup_enabled": False,
    "max_points": 5,
    "backup_path_d": str(BACKUP_D),
    "create_on_startup": False,
}

# ─────────────────────────────────────────────────────────────────
# TEMAS DE COLOR (DARK / LIGHT)
# ─────────────────────────────────────────────────────────────────
THEMES = {
    "dark": {
        "BG":      "#0D1117",
        "CARD":    "#161B22",
        "CARD2":   "#1C2128",
        "ACCENT":  "#3D7EFF",
        "ACCH":    "#5A96FF",
        "SUCCESS": "#238636",
        "SUCCH":   "#2EA043",
        "WARN":    "#D29922",
        "DANGER":  "#DA3633",
        "TPRI":    "#E6EDF3",
        "TSEC":    "#8B949E",
        "BORDER":  "#30363D",
    },
    "light": {
        "BG":      "#F5F7FA",
        "CARD":    "#FFFFFF",
        "CARD2":   "#EEF1F6",
        "ACCENT":  "#1A3EBB",
        "ACCH":    "#2550D4",
        "SUCCESS": "#1A7A3C",
        "SUCCH":   "#22A050",
        "WARN":    "#FF6B00",
        "DANGER":  "#CC2200",
        "TPRI":    "#0D0D0D",
        "TSEC":    "#555E6E",
        "BORDER":  "#D0D7E2",
    },
}

_current_theme = "dark"

def _T(key):
    return THEMES[_current_theme][key]

# Aliases dinámicos (retrocompatibilidad)
class _ThemeProxy:
    def __getattr__(self, name):
        return THEMES[_current_theme].get(name, "")

_theme = _ThemeProxy()

BG      = "#0D1117"; CARD    = "#161B22"; CARD2   = "#1C2128"
ACCENT  = "#3D7EFF"; ACCH    = "#5A96FF"
SUCCESS = "#238636"; SUCCH   = "#2EA043"
WARN    = "#D29922"; DANGER  = "#DA3633"
TPRI    = "#E6EDF3"; TSEC    = "#8B949E"; BORDER  = "#30363D"

def apply_theme(name):
    """Actualiza las variables globales de color al tema dado."""
    global _current_theme, BG, CARD, CARD2, ACCENT, ACCH, SUCCESS, SUCCH, WARN, DANGER, TPRI, TSEC, BORDER
    _current_theme = name
    t = THEMES[name]
    BG=t["BG"]; CARD=t["CARD"]; CARD2=t["CARD2"]
    ACCENT=t["ACCENT"]; ACCH=t["ACCH"]
    SUCCESS=t["SUCCESS"]; SUCCH=t["SUCCH"]
    WARN=t["WARN"]; DANGER=t["DANGER"]
    TPRI=t["TPRI"]; TSEC=t["TSEC"]; BORDER=t["BORDER"]

# ─────────────────────────────────────────────────────────────────
# ADMIN
# ─────────────────────────────────────────────────────────────────
def is_admin():
    try: return ctypes.windll.shell32.IsUserAnAdmin()
    except: return False

def request_admin():
    if not is_admin():
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable,
            " ".join(f'"{a}"' for a in sys.argv), None, 1)
        sys.exit(0)

# ─────────────────────────────────────────────────────────────────
# UTILIDADES
# ─────────────────────────────────────────────────────────────────
def log(msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    BASE_DIR.mkdir(parents=True, exist_ok=True)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {msg}\n")
    except: pass

def load_config():
    try:
        if CONFIG_FILE.exists():
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            for k, v in DEFAULT_CFG.items():
                cfg.setdefault(k, v)
            return cfg
    except: pass
    return dict(DEFAULT_CFG)

def save_config(cfg):
    BASE_DIR.mkdir(parents=True, exist_ok=True)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)

def load_history():
    try:
        if HISTORY_FILE.exists():
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except: pass
    return []

def save_history(history):
    BASE_DIR.mkdir(parents=True, exist_ok=True)
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history, f, indent=2, ensure_ascii=False)
    try:
        cfg = load_config()
        d = Path(cfg["backup_path_d"])
        d.mkdir(parents=True, exist_ok=True)
        shutil.copy2(HISTORY_FILE, d / "historial.json")
    except: pass

def run_ps(cmd, timeout=90):
    return subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", cmd],
        capture_output=True, text=True, timeout=timeout, encoding="utf-8", errors="ignore"
    )

def run_cmd(args, timeout=30):
    return subprocess.run(
        args, capture_output=True, timeout=timeout,
        encoding="cp850", errors="ignore"
    )

# ─────────────────────────────────────────────────────────────────
# VENTANA DE PROGRESO
# ─────────────────────────────────────────────────────────────────
class ProgressWindow(ctk.CTkToplevel):
    def __init__(self, parent, title="Creando Punto de Restauración"):
        super().__init__(parent)
        self.title(title)
        self.geometry("640x480")
        self.resizable(False, False)
        self.configure(fg_color=CARD)
        self.attributes("-topmost", True)
        self.protocol("WM_DELETE_WINDOW", lambda: None)  # no cerrar durante proceso

        # Centrar sobre la ventana padre
        self.update_idletasks()
        px = parent.winfo_x() + parent.winfo_width()//2 - 320
        py = parent.winfo_y() + parent.winfo_height()//2 - 240
        self.geometry(f"640x480+{px}+{py}")

        self._build()
        self.q = queue.Queue()
        self._poll()

    def _build(self):
        self.grid_columnconfigure(0, weight=1)

        # Título
        ctk.CTkLabel(self, text="🛡️  RestoreGuard Pro",
            font=ctk.CTkFont("Segoe UI", 18, "bold"),
            text_color=TPRI).grid(row=0, column=0, pady=(24, 4))

        self._status_lbl = ctk.CTkLabel(self,
            text="Iniciando proceso...",
            font=ctk.CTkFont("Segoe UI", 12),
            text_color=TSEC)
        self._status_lbl.grid(row=1, column=0, pady=(0, 12))

        # Barra de progreso
        self._pbar = ctk.CTkProgressBar(self, height=10, corner_radius=5,
            fg_color=CARD2, progress_color=ACCENT, width=560)
        self._pbar.set(0)
        self._pbar.grid(row=2, column=0, padx=40, pady=(0, 4))

        self._pct_lbl = ctk.CTkLabel(self, text="0%",
            font=ctk.CTkFont("Segoe UI", 11),
            text_color=ACCENT)
        self._pct_lbl.grid(row=3, column=0, pady=(0, 14))

        # Log de archivos
        ctk.CTkLabel(self, text="📋 Archivos y procesos:",
            font=ctk.CTkFont("Segoe UI", 11, "bold"),
            text_color=TSEC, anchor="w"
        ).grid(row=4, column=0, padx=42, sticky="w")

        self._log_box = ctk.CTkTextbox(self,
            fg_color=BG, corner_radius=8,
            font=ctk.CTkFont("Consolas", 10),
            text_color=TPRI, wrap="word",
            width=560, height=240)
        self._log_box.grid(row=5, column=0, padx=40, pady=(4, 20))
        self._log_box.configure(state="disabled")

        # Botón cerrar (deshabilitado hasta terminar)
        self._close_btn = ctk.CTkButton(self,
            text="Procesando...", state="disabled",
            fg_color=CARD2, text_color=TSEC,
            width=180, height=38, corner_radius=10,
            command=self.destroy)
        self._close_btn.grid(row=6, column=0, pady=(0, 24))

    def _poll(self):
        """Lee la cola de mensajes y actualiza la UI."""
        try:
            while True:
                item = self.q.get_nowait()
                if item[0] == "progress":
                    val = item[1] / 100
                    self._pbar.set(val)
                    self._pct_lbl.configure(text=f"{item[1]}%")
                elif item[0] == "status":
                    self._status_lbl.configure(text=item[1])
                elif item[0] == "log":
                    self._log_box.configure(state="normal")
                    self._log_box.insert("end", item[1] + "\n")
                    self._log_box.see("end")
                    self._log_box.configure(state="disabled")
                elif item[0] == "done":
                    self._pbar.set(1)
                    self._pct_lbl.configure(text="100%")
                    color = SUCCH if item[1] else DANGER
                    self._status_lbl.configure(text=item[2], text_color=color)
                    self._close_btn.configure(
                        state="normal",
                        text="✅ Cerrar" if item[1] else "❌ Cerrar",
                        fg_color=SUCCESS if item[1] else DANGER,
                        hover_color=SUCCH if item[1] else "#c0392b",
                        text_color=TPRI)
                    return
        except queue.Empty:
            pass
        self.after(80, self._poll)

    def send(self, *args):
        self.q.put(args)

# ─────────────────────────────────────────────────────────────────
# CREADOR DE PUNTO DE RESTAURACIÓN (con progreso real)
# ─────────────────────────────────────────────────────────────────
class RestoreCreator:
    """
    Crea el punto de restauración del sistema Windows y además
    respalda archivos de configuración reales a C: y D:.
    Reporta progreso mediante una función callback.
    """

    def __init__(self, tipo="MANUAL", pw: ProgressWindow = None):
        self.tipo = tipo
        self.pw = pw
        cfg = load_config()
        self.max_pts = cfg.get("max_points", 5)
        self.ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.desc = f"{PREFIX}_{tipo}_{self.ts}"
        self.backup_sub = f"respaldo_{self.ts}"
        self.dirs = {
            "c": BASE_DIR / self.backup_sub,
            "d": Path(cfg.get("backup_path_d", str(BACKUP_D))) / self.backup_sub,
        }

    def _p(self, pct, status, log_line=None):
        if self.pw:
            self.pw.send("progress", pct)
            self.pw.send("status", status)
            if log_line:
                self.pw.send("log", log_line)
        log(log_line or status)

    def _log(self, line):
        if self.pw:
            self.pw.send("log", line)
        log(line)

    def run(self):
        ok = False
        msg = ""
        try:
            ok, msg = self._do_create()
        except Exception as e:
            msg = str(e)
            ok = False
        finally:
            if self.pw:
                self.pw.send("done", ok, msg)
        return ok, msg

    def _do_create(self):
        # ── Paso 1: Preparar carpetas
        self._p(2, "Preparando carpetas de respaldo...", "")
        for label, d in self.dirs.items():
            d.mkdir(parents=True, exist_ok=True)
            self._log(f"  📁 Carpeta creada: {d}")

        # ── Paso 2: Habilitar System Restore en C: y bypass limit
        self._p(5, "Habilitando System Restore en C: ...",
                "⚙️  Enable-ComputerRestore -Drive 'C:\\'")
        run_ps("Enable-ComputerRestore -Drive 'C:\\'", timeout=20)
        run_cmd(["reg", "add", r"HKLM\Software\Microsoft\Windows NT\CurrentVersion\SystemRestore", "/v", "SystemRestorePointCreationFrequency", "/t", "REG_DWORD", "/d", "0", "/f"])

        # ── Paso 3: Eliminar puntos viejos si hay más del límite
        self._p(10, "Verificando puntos existentes...",
                f"🔍 Máximo configurado: {self.max_pts} puntos")
        self._cleanup_old_points()

        # ── Paso 4: Crear point de restauración Windows
        self._p(18, "⏳ Creando Punto de Restauración del Sistema...",
                f"🛡️  Checkpoint-Computer: '{self.desc}'")
        self._p(20, "⏳ Esto puede tardar 30-60 segundos, por favor espera...")

        r = run_ps(
            f"Checkpoint-Computer -Description '{self.desc}' "
            f"-RestorePointType 'MODIFY_SETTINGS' -ErrorAction Stop",
            timeout=120
        )

        if r.returncode != 0:
            err = (r.stderr or r.stdout or "Error desconocido").strip()
            # Intento alternativo con WMI
            self._log("  ⚠️  Reintentando con WMI...")
            r2 = run_ps(
                "$rp = [wmiclass]'root/default:SystemRestore'; "
                f"$rp.CreateRestorePoint('{self.desc}', 0, 100)",
                timeout=90
            )
            if r2.returncode != 0:
                return False, f"No se pudo crear el punto de restauración:\n{err[:300]}"

        self._log("  ✅  Punto de restauración del sistema creado correctamente")
        self._p(40, "✅ Punto del sistema creado. Guardando respaldos...", "")

        # ── Paso 5: Respaldos de configuración
        steps = [
            (45, self._backup_ipconfig),
            (52, self._backup_netsh),
            (58, self._backup_wifi),
            (64, self._backup_registry_network),
            (70, self._backup_registry_services),
            (76, self._backup_hosts),
            (82, self._backup_drivers),
            (87, self._backup_sysinfo),
            (92, self._backup_programs),
            (96, self._mirror_to_d),
        ]

        for pct, fn in steps:
            try:
                fn()
            except Exception as e:
                self._log(f"  ⚠️  {fn.__name__}: {e}")
            self._p(pct, f"Guardando respaldos... ({pct}%)")

        # ── Paso 6: Guardar historial
        self._p(98, "Guardando historial...", "📝  Actualizando historial.json")
        h = load_history()
        h.insert(0, {
            "id": self.ts,
            "description": self.desc,
            "tipo": self.tipo,
            "fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "estado": "✅ Éxito",
            "carpeta": str(self.dirs["c"]),
        })
        save_history(h[:50])

        self._p(100, "✅ ¡Completado con éxito!", "")
        self._log(f"\n✅  PROCESO COMPLETADO")
        self._log(f"   C: {self.dirs['c']}")
        self._log(f"   D: {self.dirs['d']}")
        return True, f"Punto creado y respaldo guardado:\n{self.desc}"

    # ── Backups individuales ──────────────────────────────────────
    def _backup_ipconfig(self):
        self._log("  💾  ipconfig /all → ipconfig.txt")
        r = run_cmd(["ipconfig", "/all"])
        for d in self.dirs.values():
            (d / "ipconfig.txt").write_text(r.stdout, encoding="utf-8", errors="ignore")

    def _backup_netsh(self):
        self._log("  💾  netsh → configuracion_red.txt")
        r1 = run_cmd(["netsh", "interface", "ip", "show", "config"])
        r2 = run_cmd(["netsh", "interface", "ipv4", "show", "route"])
        r3 = run_cmd(["netsh", "winsock", "show", "catalog"])
        txt = "=== IP CONFIG ===\n" + r1.stdout
        txt += "\n=== RUTAS ===\n" + r2.stdout
        txt += "\n=== WINSOCK ===\n" + r3.stdout
        for d in self.dirs.values():
            (d / "configuracion_red.txt").write_text(txt, encoding="utf-8", errors="ignore")

    def _backup_wifi(self):
        self._log("  📡  Perfiles WiFi → carpeta wifi_profiles/")
        for d in self.dirs.values():
            wdir = d / "wifi_profiles"
            wdir.mkdir(exist_ok=True)
            subprocess.run(
                ["netsh", "wlan", "export", "profile", f"folder={wdir}", "key=clear"],
                capture_output=True, timeout=20
            )
            r = run_cmd(["netsh", "wlan", "show", "profiles"])
            (wdir / "lista_redes.txt").write_text(r.stdout, encoding="utf-8", errors="ignore")
            self._log(f"     → {wdir}")

    def _backup_registry_network(self):
        self._log("  🗂️  Registro → HKLM\\SYSTEM\\CurrentControlSet\\Services\\Tcpip")
        key = "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Tcpip"
        for d in self.dirs.values():
            out = d / "registro_tcpip.reg"
            subprocess.run(
                ["reg", "export", key, str(out), "/y"],
                capture_output=True, timeout=30
            )
            self._log(f"     → {out.name}")

    def _backup_registry_services(self):
        self._log("  🗂️  Registro → HKLM\\SYSTEM\\CurrentControlSet\\Services\\Dnscache")
        key = "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Dnscache"
        for d in self.dirs.values():
            out = d / "registro_dns.reg"
            subprocess.run(
                ["reg", "export", key, str(out), "/y"],
                capture_output=True, timeout=30
            )
            self._log(f"     → {out.name}")

    def _backup_hosts(self):
        hosts = Path(r"C:\Windows\System32\drivers\etc\hosts")
        self._log(f"  📄  Copiando hosts → hosts.txt")
        if hosts.exists():
            for d in self.dirs.values():
                shutil.copy2(hosts, d / "hosts.txt")
                self._log(f"     → {d / 'hosts.txt'}")

    def _backup_drivers(self):
        self._log("  📋  Lista de drivers → drivers.txt")
        r = run_ps("Get-WmiObject Win32_PnPSignedDriver | "
                   "Select-Object DeviceName, DriverVersion, Manufacturer | "
                   "Sort-Object DeviceName | Format-Table -AutoSize | Out-String -Width 200",
                   timeout=40)
        for d in self.dirs.values():
            (d / "drivers.txt").write_text(r.stdout, encoding="utf-8", errors="ignore")
            self._log(f"     → {d / 'drivers.txt'}")

    def _backup_sysinfo(self):
        self._log("  🖥️  Información del sistema → sysinfo.txt")
        r = run_cmd(["systeminfo"])
        for d in self.dirs.values():
            (d / "sysinfo.txt").write_text(r.stdout, encoding="utf-8", errors="ignore")
            self._log(f"     → {d / 'sysinfo.txt'}")

    def _backup_programs(self):
        self._log("  📦  Programas instalados → programas_instalados.txt")
        r = run_ps(
            "Get-ItemProperty HKLM:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\* "
            "| Select-Object DisplayName, DisplayVersion, Publisher, InstallDate "
            "| Sort-Object DisplayName | Format-Table -AutoSize | Out-String -Width 200",
            timeout=30
        )
        for d in self.dirs.values():
            (d / "programas_instalados.txt").write_text(r.stdout, encoding="utf-8", errors="ignore")
            self._log(f"     → {d / 'programas_instalados.txt'}")

    def _mirror_to_d(self):
        self._log("  🔄  Verificando mirror en D: ...")
        dc = self.dirs["c"]
        dd = self.dirs["d"]
        dd.mkdir(parents=True, exist_ok=True)
        count = 0
        for f in dc.iterdir():
            if f.is_file():
                dst = dd / f.name
                shutil.copy2(f, dst)
                self._log(f"     → D: {dst.name}")
                count += 1
        self._log(f"  ✅  {count} archivo(s) copiados a D:")

    # ── Limpiar puntos viejos ─────────────────────────────────────
    def _cleanup_old_points(self):
        r = run_ps(
            f"Get-ComputerRestorePoint | "
            f"Where-Object {{$_.Description -like '{PREFIX}*'}} | "
            f"Sort-Object CreationTime | "
            f"Select-Object SequenceNumber, Description | ConvertTo-Json",
            timeout=30
        )
        try:
            data = json.loads(r.stdout.strip()) if r.stdout.strip() else []
            if isinstance(data, dict): data = [data]
        except: data = []

        self._log(f"  🔍  Puntos anteriores encontrados: {len(data)}")

        while len(data) >= self.max_pts:
            oldest = data[0]
            seq = oldest.get("SequenceNumber", 0)
            self._log(f"  🗑️  Eliminando punto antiguo #{seq}: {oldest.get('Description','')}")
            run_ps(
                f'$p = Get-WmiObject -Namespace root\\default -Class SystemRestore '
                f'-Filter "SequenceNumber={seq}"; if ($p) {{ $p.Delete() }}',
                timeout=30
            )
            data = data[1:]

        # También limpiar respaldo antiguo en C: y D:
        self._cleanup_old_backups()

    def _cleanup_old_backups(self):
        for label, base in [("C:", BASE_DIR), ("D:", Path(load_config().get("backup_path_d", str(BACKUP_D))))]:
            if not base.exists(): continue
            subs = sorted([d for d in base.iterdir()
                           if d.is_dir() and d.name.startswith("respaldo_")],
                          key=lambda d: d.stat().st_mtime)
            while len(subs) >= self.max_pts:
                old = subs.pop(0)
                shutil.rmtree(old, ignore_errors=True)
                self._log(f"  🗑️  Respaldo antiguo eliminado en {label}: {old.name}")

# ─────────────────────────────────────────────────────────────────
# SCHEDULER THREAD
# ─────────────────────────────────────────────────────────────────
class SchedulerThread(threading.Thread):
    def __init__(self, callback):
        super().__init__(daemon=True)
        self.callback = callback
        self._stop = threading.Event()
        self._setup()

    def _setup(self):
        if not SCHEDULE_OK: return
        schedule.clear()
        cfg = load_config()
        if cfg.get("auto_enabled", True):
            t = cfg.get("auto_time", "08:00")
            schedule.every().day.at(t).do(self.callback)
            log(f"Programador: auto a las {t}")

    def run(self):
        if not SCHEDULE_OK: return
        while not self._stop.is_set():
            schedule.run_pending()
            time.sleep(20)

    def update(self, new_time):
        self._setup()

# ─────────────────────────────────────────────────────────────────
# WIDGETS HELPERS
# ─────────────────────────────────────────────────────────────────
def lbl(parent, text, size=12, weight="normal", color=None, **kw):
    if color is None: color = TPRI
    return ctk.CTkLabel(parent, text=text,
        font=ctk.CTkFont("Segoe UI", size, weight),
        text_color=color, **kw)

def card(parent, **kw):
    return ctk.CTkFrame(parent, fg_color=CARD, corner_radius=14, **kw)

# ─────────────────────────────────────────────────────────────────
# RESTAURADOR (Restore Wizard Backend)
# ─────────────────────────────────────────────────────────────────
class SystemRestorer:
    def __init__(self, history_entry, pw: ProgressWindow = None):
        self.entry = history_entry
        self.pw = pw
        self.backup_dir = Path(self.entry.get("carpeta", ""))
        self.desc = self.entry.get("description", "")
        
    def _p(self, pct, status, log_line=None):
        if self.pw:
            self.pw.send("progress", pct)
            self.pw.send("status", status)
            if log_line: self.pw.send("log", log_line)
        log(log_line or status)

    def _log(self, line):
        if self.pw: self.pw.send("log", line)
        log(line)

    def run_windows_restore(self):
        ok, msg = False, ""
        try: ok, msg = self._do_windows_restore()
        except Exception as e: ok, msg = False, str(e)
        finally:
            if self.pw: self.pw.send("done", ok, msg)
        return ok, msg

    def run_manual_restore(self):
        ok, msg = False, ""
        try: ok, msg = self._do_manual_restore()
        except Exception as e: ok, msg = False, str(e)
        finally:
            if self.pw: self.pw.send("done", ok, msg)
        return ok, msg

    def _do_windows_restore(self):
        self._p(10, "Buscando punto en Windows...", f"🔍 Buscando secuencia para: '{self.desc}'")
        r = run_ps(
            f"Get-ComputerRestorePoint | "
            f"Where-Object {{$_.Description -eq '{self.desc}'}} | "
            f"Sort-Object CreationTime -Descending | "
            f"Select-Object -First 1 SequenceNumber | ConvertTo-Json",
            timeout=30
        )
        try:
            data = json.loads(r.stdout.strip())
            seq = data.get("SequenceNumber", None) if isinstance(data, dict) else None
        except:
            seq = None
            
        if not seq:
            return False, f"No se encontró el punto '{self.desc}' nativo en Windows. Prueba rescate manual."
            
        self._p(30, "Iniciando restauración...", f"🛡️ Ejecutando Restore-Computer -RestorePoint {seq}")
        self._p(50, "⏳ El sistema comenzará a restaurarse y se reiniciará inmediatamente...")
        r2 = run_ps(f"Restore-Computer -RestorePoint {seq} -Confirm:$False", timeout=60)
        
        if r2.returncode != 0:
            return False, f"Fallo al invocar Restore-Computer: {r2.stderr or r2.stdout}"
            
        self._p(100, "✅ Restauración iniciada. Reiniciando...", "")
        return True, "Windows se está reiniciando para restaurar el sistema."

    def _do_manual_restore(self):
        if not self.backup_dir.exists():
            return False, f"La carpeta de respaldo no existe: {self.backup_dir}"
        
        self._p(10, "Restaurando servicios de red...", "")
        tcp_reg = self.backup_dir / "registro_tcpip.reg"
        dns_reg = self.backup_dir / "registro_dns.reg"
        
        if tcp_reg.exists():
            self._log(f"  🗂️ Importando {tcp_reg.name}")
            run_cmd(["reg", "import", str(tcp_reg)])
        if dns_reg.exists():
            self._log(f"  🗂️ Importando {dns_reg.name}")
            run_cmd(["reg", "import", str(dns_reg)])
            
        self._p(40, "Restaurando archivo hosts...", "")
        hosts_file = self.backup_dir / "hosts.txt"
        if hosts_file.exists():
            sys_hosts = Path(r"C:\Windows\System32\drivers\etc\hosts")
            try:
                shutil.copy2(hosts_file, sys_hosts)
                self._log("  📄 Archivo hosts restaurado")
            except Exception as e:
                self._log(f"  ⚠️ Error copiando hosts: {e}")
                
        self._p(60, "Restaurando perfiles de WiFi...", "")
        wifi_dir = self.backup_dir / "wifi_profiles"
        if wifi_dir.exists():
            count = 0
            for xml in wifi_dir.glob("*.xml"):
                subprocess.run(["netsh", "wlan", "add", "profile", f"filename={xml}"], capture_output=True)
                count += 1
            self._log(f"  📡 Se restauraron {count} perfiles WiFi")
            
        self._p(90, "Actualizando estado...", "  ✅ Registro, Hosts y Red inyectados")
        self._p(100, "✅ Restauración manual completada", "")
        return True, "Configuraciones inyectadas con éxito."

# ─────────────────────────────────────────────────────────────────
# ASISTENTE WIZARD
# ─────────────────────────────────────────────────────────────────
class RestoreWizard(ctk.CTkToplevel):
    def __init__(self, parent, entry):
        super().__init__(parent)
        self.entry = entry
        self.app = parent
        self.title("Asistente de Restauración")
        self.geometry("640x530")
        self.resizable(False, False)
        self.configure(fg_color=CARD)
        self.attributes("-topmost", True)
        self.grab_set()

        self.update_idletasks()
        try:
            px = parent.winfo_x() + parent.winfo_width()//2 - 320
            py = parent.winfo_y() + parent.winfo_height()//2 - 265
            self.geometry(f"640x530+{px}+{py}")
        except: pass

        self._build()

    def _build(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        lbl(self, "🚑 Asistente de Restauración", 18, "bold").grid(row=0, column=0, pady=(24, 10))
        
        desc = self.entry.get("description", "Desconocido")
        fecha = self.entry.get("fecha", "Desconocida")
        
        info = card(self)
        info.grid(row=1, column=0, sticky="ew", padx=30, pady=10)
        info.grid_columnconfigure(0, weight=1)
        lbl(info, "Punto Seleccionado:", 12, "bold", TSEC).grid(row=0, column=0, sticky="w", padx=16, pady=(12,0))
        lbl(info, f"📅 {fecha}   |   🏷️ {desc}", 12, "bold", ACCENT).grid(row=1, column=0, sticky="w", padx=16, pady=(4,12))

        self._opts = ctk.CTkFrame(self, fg_color="transparent")
        self._opts.grid(row=2, column=0, sticky="nsew", padx=30, pady=10)
        self._opts.grid_columnconfigure(0, weight=1)
        
        o1 = card(self._opts)
        o1.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        o1.grid_columnconfigure(0, weight=1)
        lbl(o1, "1️⃣ Restauración Automática (Windows)", 13, "bold").grid(row=0, column=0, sticky="w", padx=16, pady=(16,4))
        lbl(o1, "Devuelve toda la PC a este punto (Registro, Windows, Drivers).\n⚠️ ¡La computadora se REINICIARÁ inmediatamente!", 11, color=WARN, justify="left").grid(row=1, column=0, sticky="w", padx=16, pady=(0, 10))
        ctk.CTkButton(o1, text="Ejecutar Auto-Restauración", fg_color=WARN, hover_color=DANGER, font=ctk.CTkFont(weight="bold"), command=self._run_windows).grid(row=2, column=0, sticky="e", padx=16, pady=(0,16))

        o2 = card(self._opts)
        o2.grid(row=1, column=0, sticky="ew")
        o2.grid_columnconfigure(0, weight=1)
        lbl(o2, "2️⃣ Rescate Manual Temerario (Solo Archivos y Red)", 13, "bold").grid(row=0, column=0, sticky="w", padx=16, pady=(16,4))
        lbl(o2, "Repara tu red inyectando en silencio los registros, IPs y Wi-Fi de esta fecha.\n✔️ Rápido, no toca a Windows superficial, y no reinicia la PC.", 11, color=TSEC, justify="left").grid(row=1, column=0, sticky="w", padx=16, pady=(0, 10))
        ctk.CTkButton(o2, text="Ejecutar Rescate Manual", fg_color=SUCCESS, hover_color=SUCCH, font=ctk.CTkFont(weight="bold"), command=self._run_manual).grid(row=2, column=0, sticky="e", padx=16, pady=(0,16))

    def _run_windows(self):
        self._opts.destroy()
        self._launch_restore("WINDOWS")
        
    def _run_manual(self):
        self._opts.destroy()
        self._launch_restore("MANUAL")
        
    def _launch_restore(self, mode):
        pw = ProgressWindow(self.master, title="Restaurando...")
        try:
            pw.geometry(f"640x480+{self.winfo_x()}+{self.winfo_y()}")
        except: pass
        pw.grab_set()
        self.destroy()

        def _run():
            restorer = SystemRestorer(self.entry, pw)
            if mode == "WINDOWS": ok, msg = restorer.run_windows_restore()
            else:                 ok, msg = restorer.run_manual_restore()
        threading.Thread(target=_run, daemon=True).start()

# ─────────────────────────────────────────────────────────────────
# APP PRINCIPAL
# ─────────────────────────────────────────────────────────────────
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        BASE_DIR.mkdir(parents=True, exist_ok=True)
        apply_theme("light")
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        self.title("RestoreGuard Pro v1.0.0")
        self.geometry("860x700")
        self.minsize(720, 580)
        self.configure(fg_color=BG)
        self._creating = False
        self._theme_name = "light"

        self._build()
        self.refresh()

        self.sched = SchedulerThread(lambda: self._bg_create("AUTO"))
        self.sched.start()

        if TRAY_OK and PIL_OK:
            threading.Thread(target=self._start_tray, daemon=True).start()

        self.protocol("WM_DELETE_WINDOW", self.withdraw)

        cfg = load_config()
        if cfg.get("create_on_startup", False):
            threading.Timer(4.0, lambda: self._bg_create("AUTO")).start()

    # ── Construcción UI ───────────────────────────────────────────
    def _build(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        self._build_header()
        self._build_tabs()
        self._build_pages()
        self._switch("inicio")

    def _build_header(self):
        self._header = ctk.CTkFrame(self, fg_color=CARD, corner_radius=0, height=72)
        self._header.grid(row=0, column=0, sticky="ew")
        self._header.grid_propagate(False)
        self._header.grid_columnconfigure(1, weight=1)
        lbl(self._header, "🛡️", 34).grid(row=0, column=0, padx=(18,6), pady=16)
        tf = ctk.CTkFrame(self._header, fg_color="transparent")
        tf.grid(row=0, column=1, sticky="w")
        lbl(tf,"RestoreGuard Pro v1.0.0",20,"bold").pack(anchor="w")
        lbl(tf,"Protección automática real del sistema Windows",11,color=TSEC).pack(anchor="w")
        self._badge = lbl(self._header,"● Activo",12,"bold",SUCCH)
        self._badge.grid(row=0, column=2, padx=20)
        # Botón toggle de tema
        self._theme_btn = ctk.CTkButton(
            self._header, text="🌙  Modo Oscuro",
            font=ctk.CTkFont("Segoe UI", 11, "bold"),
            fg_color="#1A3EBB", hover_color="#2550D4",
            width=120, height=34, corner_radius=10,
            command=self._toggle_theme)
        self._theme_btn.grid(row=0, column=3, padx=(8, 8))
        
        # Botón Salir
        self._exit_btn = ctk.CTkButton(
            self._header, text="❌ Salir",
            font=ctk.CTkFont("Segoe UI", 11, "bold"),
            fg_color=DANGER, hover_color="#c0392b", text_color="white",
            width=80, height=34, corner_radius=10,
            command=lambda: os._exit(0))
        self._exit_btn.grid(row=0, column=4, padx=(0, 8))
        
        # Botón Acerca de
        self._about_btn = ctk.CTkButton(
            self._header, text="💡 Acerca de",
            font=ctk.CTkFont("Segoe UI", 11, "bold"),
            fg_color="transparent", hover_color=BG, text_color=TSEC, border_width=1, border_color=BORDER,
            width=100, height=34, corner_radius=10,
            command=self._show_about)
        self._about_btn.grid(row=0, column=5, padx=(0, 20))

    def _build_tabs(self):
        tb = ctk.CTkFrame(self, fg_color=CARD2, corner_radius=0, height=46)
        tb.grid(row=1, column=0, sticky="ew")
        tb.grid_propagate(False)
        self._tab_btns = {}
        for i,(k,label) in enumerate([
            ("inicio","🏠  Inicio"),("historial","📋  Historial"),("config","⚙️  Configuración"),("ayuda","❓  Ayuda")]):
            b = ctk.CTkButton(tb, text=label,
                font=ctk.CTkFont(size=12), fg_color="transparent",
                hover_color=CARD, text_color=TSEC,
                corner_radius=0, height=46, width=180,
                command=lambda key=k: self._switch(key))
            b.grid(row=0, column=i)
            self._tab_btns[k] = b

    def _build_pages(self):
        self._content = ctk.CTkFrame(self, fg_color=BG, corner_radius=0)
        self._content.grid(row=2, column=0, sticky="nsew")
        self._content.grid_columnconfigure(0, weight=1)
        self._content.grid_rowconfigure(0, weight=1)
        self._pages = {
            "inicio":    self._page_inicio(),
            "historial": self._page_historial(),
            "config":    self._page_config(),
            "ayuda":     self._page_ayuda(),
        }

    def _switch(self, key):
        for p in self._pages.values(): p.grid_remove()
        self._pages[key].grid(row=0, column=0, sticky="nsew")
        self._active_tab = key
        for k, b in self._tab_btns.items():
            b.configure(text_color=ACCENT if k==key else TSEC,
                        fg_color=CARD if k==key else "transparent")

    def _toggle_theme(self):
        """Alterna entre modo oscuro y modo claro y reconstruye la UI."""
        new_theme = "light" if self._theme_name == "dark" else "dark"
        self._theme_name = new_theme
        apply_theme(new_theme)
        ctk.set_appearance_mode("light" if new_theme == "light" else "dark")
        self.configure(fg_color=BG)
        # Reconstruir toda la UI con los nuevos colores
        active = getattr(self, "_active_tab", "inicio")
        for widget in self.winfo_children():
            widget.destroy()
        self._creating = False
        self._build()
        self.refresh()
        self._switch(active)
        # Ajustar el botón de toggle
        if new_theme == "light":
            self._theme_btn.configure(
                text="🌙  Modo Oscuro",
                fg_color="#1A3EBB", hover_color="#2550D4")
        else:
            self._theme_btn.configure(
                text="☀️  Modo Claro",
                fg_color=ACCENT, hover_color=ACCH)

    def _show_about(self):
        t = ctk.CTkToplevel(self)
        t.title("Acerca de RestoreGuard Pro")
        t.geometry("450x280")
        t.resizable(False, False)
        t.configure(fg_color=BG)
        t.attributes("-topmost", True)
        
        # Centrar ventana
        t.update_idletasks()
        px = self.winfo_x() + self.winfo_width()//2 - 225
        py = self.winfo_y() + self.winfo_height()//2 - 140
        t.geometry(f"+{px}+{py}")
        
        c = card(t)
        c.pack(fill="both", expand=True, padx=20, pady=20)
        lbl(c, "🛡️ RestoreGuard Pro v1.0.0", 16, "bold", ACCENT).pack(pady=(20, 10))
        
        info = ("Autor y Creador Original: Ing. Lucidio Fuenmayor\n\n"
              "Licencia: Uso personal protegido.\n\n"
              "Aplicación diseñada para la protección, mantenimiento\ny rescate seguro de sistemas Windows.")
        lbl(c, info, 12, color=TPRI, justify="center").pack(pady=10)
        
        ctk.CTkButton(c, text="Cerrar", width=120, fg_color=CARD2, hover_color=BORDER, text_color=TPRI, command=t.destroy).pack(pady=(10, 20))

    # ── INICIO ────────────────────────────────────────────────────
    def _page_inicio(self):
        page = ctk.CTkFrame(self._content, fg_color=BG, corner_radius=0)
        page.grid_columnconfigure(0, weight=1)
        sc = ctk.CTkScrollableFrame(page, fg_color=BG, corner_radius=0)
        sc.grid(row=0, column=0, sticky="nsew")
        page.grid_rowconfigure(0, weight=1)
        sc.grid_columnconfigure(0, weight=1)

        # Stat cards
        row0 = ctk.CTkFrame(sc, fg_color="transparent")
        row0.grid(row=0, column=0, sticky="ew", padx=20, pady=(20,10))
        row0.grid_columnconfigure((0,1,2), weight=1)
        self._s_last  = self._stat_card(row0,"🕐 Último Punto","—",TSEC,0)
        self._s_next  = self._stat_card(row0,"⏰ Próximo Auto","—",WARN,1)
        self._s_total = self._stat_card(row0,"🛡️ Total Guardados","—",ACCENT,2)

        # Botón crear
        bc = card(sc); bc.grid(row=1,column=0,sticky="ew",padx=20,pady=10)
        bc.grid_columnconfigure(0, weight=1)
        lbl(bc,"Protección Manual",14,"bold").grid(row=0,column=0,pady=(20,4))
        lbl(bc,"Crea ahora un punto de restauración + respaldo completo",11,color=TSEC
            ).grid(row=1,column=0,pady=(0,14))
        self._create_btn = ctk.CTkButton(bc,
            text="🛡️   Crear Punto de Restauración Ahora",
            font=ctk.CTkFont("Segoe UI",15,"bold"),
            fg_color=ACCENT, hover_color=ACCH,
            height=58, corner_radius=12,
            command=self._on_manual)
        self._create_btn.grid(row=2,column=0,padx=40,pady=(0,20),sticky="ew")

        # Qué se guarda
        ic = card(sc); ic.grid(row=2,column=0,sticky="ew",padx=20,pady=10)
        ic.grid_columnconfigure(0, weight=1)
        lbl(ic,"📦 Qué se guarda en cada respaldo",13,"bold"
            ).grid(row=0,column=0,sticky="w",padx=20,pady=(16,8))
        items = [
            ("🛡️","Punto de Restauración del Sistema Windows","(registro, drivers, configuración del sistema)"),
            ("🌐","Configuración de red completa","ipconfig, netsh, rutas, DNS, Winsock"),
            ("📡","Perfiles WiFi","Todas las redes guardadas con contraseñas"),
            ("🗂️","Claves de registro de red","Tcpip, DNS cache (archivos .reg)"),
            ("📄","Archivo hosts","C:\\Windows\\System32\\drivers\\etc\\hosts"),
            ("🖥️","Información del sistema","CPU, RAM, SO, componentes"),
            ("🚗","Lista de drivers","Todos los controladores instalados"),
            ("📦","Programas instalados","Lista completa de software"),
        ]
        for i,(ico,title,desc) in enumerate(items):
            rf = ctk.CTkFrame(ic, fg_color=CARD2 if i%2==0 else CARD, corner_radius=8)
            rf.grid(row=i+1,column=0,sticky="ew",padx=16,pady=2)
            rf.grid_columnconfigure(1, weight=1)
            lbl(rf,ico,16).grid(row=0,column=0,padx=12,pady=8)
            lbl(rf,title,11,"bold").grid(row=0,column=1,sticky="w",padx=4,pady=8)
            lbl(rf,desc,10,color=TSEC).grid(row=0,column=2,sticky="e",padx=16,pady=8)
        ctk.CTkLabel(ic,text="").grid(row=len(items)+1,column=0,pady=6)

        # Ubicaciones
        lc = card(sc); lc.grid(row=3,column=0,sticky="ew",padx=20,pady=10)
        lc.grid_columnconfigure(0, weight=1)
        lbl(lc,"📁 Ubicaciones de Respaldo",13,"bold"
            ).grid(row=0,column=0,sticky="w",padx=20,pady=(16,8))
        self._loc_c = self._loc_row(lc,"💾 Disco C:",str(BASE_DIR),1)
        cfg = load_config()
        d_label = "💿 Disco D:" if HAS_DRIVE_D else "💿 Disco D: (no disponible → usando C:)"
        self._loc_d = self._loc_row(lc, d_label, cfg.get("backup_path_d",str(BACKUP_D)), 2, is_d=True)
        ctk.CTkLabel(lc,text="").grid(row=3,column=0,pady=6)

        # Actividad reciente
        ac = card(sc); ac.grid(row=4,column=0,sticky="ew",padx=20,pady=(10,20))
        ac.grid_columnconfigure(0, weight=1)
        lbl(ac,"⚡ Actividad Reciente",13,"bold"
            ).grid(row=0,column=0,sticky="w",padx=20,pady=(16,8))
        self._recent = ctk.CTkFrame(ac, fg_color="transparent")
        self._recent.grid(row=1,column=0,sticky="ew",padx=20,pady=(0,16))
        self._recent.grid_columnconfigure(0, weight=1)

        return page

    def _stat_card(self, parent, title, value, color, col):
        f = ctk.CTkFrame(parent, fg_color=CARD, corner_radius=12)
        f.grid(row=0, column=col, padx=5, sticky="ew")
        f.grid_columnconfigure(0, weight=1)
        lbl(f,title,11,color=TSEC).grid(row=0,column=0,padx=16,pady=(14,2))
        v = lbl(f,value,14,"bold",color)
        v.grid(row=1,column=0,padx=16,pady=(0,14))
        return v

    def _loc_row(self, parent, label, path, row, is_d=False):
        f = ctk.CTkFrame(parent, fg_color=CARD2, corner_radius=8)
        f.grid(row=row,column=0,sticky="ew",padx=16,pady=3)
        f.grid_columnconfigure(1, weight=1)
        lbl(f,label,12,"bold").grid(row=0,column=0,padx=12,pady=10,sticky="w")
        lbl(f,path,10,color=TSEC).grid(row=0,column=1,padx=8,pady=10,sticky="w")
        if is_d and not HAS_DRIVE_D:
            status_txt = "ℹ️ No disponible (usando C:)"
            status_col = ACCENT
        else:
            ex = Path(path).exists()
            status_txt = "✅ OK" if ex else "⚠️ No existe"
            status_col = SUCCH if ex else WARN
        s = lbl(f, status_txt, 11, "bold", status_col)
        s.grid(row=0,column=2,padx=12,pady=10)
        return s

    # ── HISTORIAL ─────────────────────────────────────────────────
    def _page_historial(self):
        page = ctk.CTkFrame(self._content, fg_color=BG, corner_radius=0)
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(1, weight=1)

        hdr = card(page); hdr.grid(row=0,column=0,sticky="ew",padx=20,pady=(20,10))
        hdr.grid_columnconfigure(0, weight=1)
        lbl(hdr,"📋 Historial de Puntos de Restauración",14,"bold"
            ).grid(row=0,column=0,sticky="w",padx=20,pady=14)
        ctk.CTkButton(hdr, text="🔄 Actualizar",
            font=ctk.CTkFont(size=11), fg_color=CARD2,
            hover_color=BORDER, width=110, height=30, corner_radius=8,
            command=self.refresh).grid(row=0,column=1,padx=14,pady=12)

        tbl = card(page); tbl.grid(row=1,column=0,sticky="nsew",padx=20,pady=(0,20))
        tbl.grid_columnconfigure(0, weight=1)
        tbl.grid_rowconfigure(1, weight=1)

        hrow = ctk.CTkFrame(tbl, fg_color=CARD2, corner_radius=0)
        hrow.grid(row=0,column=0,sticky="ew")
        hrow.grid_columnconfigure(2, weight=1)
        for i,t in enumerate(["📅 Fecha","🏷️ Tipo","📝 Descripción","✅ Estado", ""]):
            lbl(hrow,t,11,"bold",TSEC).grid(row=0,column=i,padx=14,pady=10,sticky="w")

        self._hist_sc = ctk.CTkScrollableFrame(tbl, fg_color="transparent", corner_radius=0)
        self._hist_sc.grid(row=1,column=0,sticky="nsew")
        self._hist_sc.grid_columnconfigure((0,1,2,3,4), weight=1)
        return page

    # ── CONFIG ────────────────────────────────────────────────────
    def _page_config(self):
        page = ctk.CTkFrame(self._content, fg_color=BG, corner_radius=0)
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(0, weight=1)

        sc = ctk.CTkScrollableFrame(page, fg_color=BG, corner_radius=0)
        sc.grid(row=0,column=0,sticky="nsew")
        sc.grid_columnconfigure(0, weight=1)

        cfg = load_config()

        s1 = card(sc); s1.grid(row=0,column=0,sticky="ew",padx=20,pady=(20,10))
        s1.grid_columnconfigure(1, weight=1)
        lbl(s1,"⏰ Programación Automática",13,"bold"
            ).grid(row=0,column=0,columnspan=2,sticky="w",padx=20,pady=(16,8))

        lbl(s1,"Activar puntos automáticos diarios:").grid(row=1,column=0,sticky="w",padx=20,pady=8)
        self._sw_auto = ctk.CTkSwitch(s1, text="", progress_color=ACCENT, width=50)
        self._sw_auto.grid(row=1,column=1,sticky="e",padx=20,pady=8)
        if cfg.get("auto_enabled",True): self._sw_auto.select()

        lbl(s1,"Hora del punto automático (HH:MM):").grid(row=2,column=0,sticky="w",padx=20,pady=8)
        self._auto_time = tk.StringVar(value=cfg.get("auto_time","08:00"))
        tf = ctk.CTkFrame(s1, fg_color="transparent")
        tf.grid(row=2,column=1,sticky="e",padx=20,pady=8)
        ctk.CTkEntry(tf,textvariable=self._auto_time,width=80,
            font=ctk.CTkFont(size=13),justify="center").pack(side="left")
        lbl(tf,"(HH:MM)",10,color=TSEC).pack(side="left",padx=6)

        lbl(s1,"Crear punto al iniciar Windows:").grid(row=3,column=0,sticky="w",padx=20,pady=8)
        self._sw_startup = ctk.CTkSwitch(s1, text="", progress_color=ACCENT, width=50)
        self._sw_startup.grid(row=3,column=1,sticky="e",padx=20,pady=8)
        if cfg.get("create_on_startup",False): self._sw_startup.select()
        ctk.CTkLabel(s1,text="").grid(row=4,column=0,pady=4)

        s2 = card(sc); s2.grid(row=1,column=0,sticky="ew",padx=20,pady=10)
        s2.grid_columnconfigure(1, weight=1)
        lbl(s2,"💾 Almacenamiento",13,"bold"
            ).grid(row=0,column=0,columnspan=2,sticky="w",padx=20,pady=(16,8))

        lbl(s2,"Máx. puntos a conservar (2–10):").grid(row=1,column=0,sticky="w",padx=20,pady=8)
        mf = ctk.CTkFrame(s2, fg_color="transparent")
        mf.grid(row=1,column=1,sticky="e",padx=20,pady=8)
        self._max_var = tk.IntVar(value=cfg.get("max_points",5))
        self._max_lbl = lbl(mf,str(self._max_var.get()),14,"bold",ACCENT)
        self._max_lbl.pack(side="right",padx=(8,0))
        sl = ctk.CTkSlider(mf,from_=2,to=10,number_of_steps=8,
            variable=self._max_var,progress_color=ACCENT,width=140)
        sl.pack(side="right")
        sl.configure(command=lambda v: self._max_lbl.configure(text=str(int(v))))

        lbl(s2,"Carpeta respaldo en D:").grid(row=2,column=0,sticky="w",padx=20,pady=8)
        self._d_path = tk.StringVar(value=cfg.get("backup_path_d",str(BACKUP_D)))
        ctk.CTkEntry(s2,textvariable=self._d_path,
            font=ctk.CTkFont(size=11),height=32
            ).grid(row=2,column=1,sticky="ew",padx=20,pady=8)
        ctk.CTkLabel(s2,text="").grid(row=3,column=0,pady=4)

        s3 = card(sc); s3.grid(row=2,column=0,sticky="ew",padx=20,pady=10)
        s3.grid_columnconfigure(1, weight=1)
        lbl(s3,"🖥️ Sistema",13,"bold"
            ).grid(row=0,column=0,columnspan=2,sticky="w",padx=20,pady=(16,8))
        lbl(s3,"Iniciar RestoreGuard con Windows:").grid(row=1,column=0,sticky="w",padx=20,pady=8)
        self._sw_win = ctk.CTkSwitch(s3, text="", progress_color=ACCENT, width=50)
        self._sw_win.grid(row=1,column=1,sticky="e",padx=20,pady=8)
        if cfg.get("startup_enabled",False): self._sw_win.select()
        ctk.CTkLabel(s3,text="").grid(row=2,column=0,pady=4)

        ctk.CTkButton(sc,text="💾  Guardar Configuración",
            font=ctk.CTkFont("Segoe UI",13,"bold"),
            fg_color=SUCCESS,hover_color=SUCCH,
            height=48,corner_radius=10,
            command=self._save_cfg
            ).grid(row=3,column=0,padx=20,pady=(10,30),sticky="ew")

        return page

    # ── AYUDA ─────────────────────────────────────────────────────
    def _page_ayuda(self):
        page = ctk.CTkFrame(self._content, fg_color=BG, corner_radius=0)
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(0, weight=1)

        sc = ctk.CTkScrollableFrame(page, fg_color=BG, corner_radius=0)
        sc.grid(row=0,column=0,sticky="nsew")
        sc.grid_columnconfigure(0, weight=1)

        h1 = card(sc); h1.grid(row=0,column=0,sticky="ew",padx=20,pady=(20,10))
        h1.grid_columnconfigure(0, weight=1)
        lbl(h1,"🚑 ¿Cómo restauro mi sistema ante un fallo?",14,"bold").grid(row=0,column=0,sticky="w",padx=20,pady=(16,8))

        t1 = ("1️⃣ Restauración Automática (Recomendada)\n"
              "Si tienes problemas graves o te quedaste sin red, usa el sistema de Windows:\n\n"
              "  • Presiona la tecla Windows en tu teclado y busca 'Crear un punto de restauración'.\n"
              "  • Abre esa opción y haz clic en el botón 'Restaurar sistema...'.\n"
              "  • Dale a Siguiente (marca 'Mostrar más puntos' si aparece).\n"
              "  • Busca en la lista el punto que dice 'RestoreGuard_...' con la fecha correcta.\n"
              "  • Seleccionalo y haz clic en Siguiente y luego Finalizar. Tu PC se reiniciará reparada.")
        lbl(h1,t1,12,color=TPRI,justify="left").grid(row=1,column=0,sticky="w",padx=20,pady=(0,16))

        h2 = card(sc); h2.grid(row=1,column=0,sticky="ew",padx=20,pady=10)
        h2.grid_columnconfigure(0, weight=1)
        
        t2 = ("2️⃣ Recuperación Manual (Desde tu Disco D:)\n"
              "Si la restauración automática no devolvió tus contraseñas o Ips específicas:\n\n"
              "  • Ve a tu disco D:\\ y entra a la carpeta 'punto de restauración'.\n"
              "  • Busca la carpeta 'respaldo_' con la fecha deseada.\n"
              "  • Doble clic en 'registro_tcpip.reg' y 'registro_dns.reg' para reparar los servicios.\n"
              "  • Abre la subcarpeta 'wifi_profiles\\' para leer tus antiguas claves de Wi-Fi.\n"
              "  • Abre 'configuracion_red.txt' o 'ipconfig.txt' para ver tus IP y DNS anteriores.")
        lbl(h2,t2,12,color=TPRI,justify="left").grid(row=0,column=0,sticky="w",padx=20,pady=16)

        return page

    # ── ACCIONES ──────────────────────────────────────────────────
    def _on_manual(self):
        if self._creating: return
        self._creating = True
        self._create_btn.configure(state="disabled", text="⏳  Procesando...")

        pw = ProgressWindow(self)
        pw.grab_set()

        def _run():
            creator = RestoreCreator("MANUAL", pw)
            ok, msg = creator.run()
            self.after(0, lambda: self._done(ok))

        threading.Thread(target=_run, daemon=True).start()

    def _done(self, ok):
        self._creating = False
        self._create_btn.configure(state="normal",
            text="🛡️   Crear Punto de Restauración Ahora")
        self.refresh()

    def _bg_create(self, tipo="AUTO"):
        if self._creating: return
        self._creating = True
        pw = ProgressWindow(self)

        def _run():
            creator = RestoreCreator(tipo, pw)
            ok, msg = creator.run()
            self.after(0, lambda: self._done(ok))

        threading.Thread(target=_run, daemon=True).start()

    def _save_cfg(self):
        cfg = {
            "auto_time":         self._auto_time.get(),
            "auto_enabled":      bool(self._sw_auto.get()),
            "startup_enabled":   bool(self._sw_win.get()),
            "max_points":        int(self._max_var.get()),
            "backup_path_d":     self._d_path.get(),
            "create_on_startup": bool(self._sw_startup.get()),
        }
        save_config(cfg)
        self.sched.update(cfg["auto_time"])
        self._apply_startup(cfg["startup_enabled"])
        self.refresh()
        self._toast("✅ Configuración guardada")

    def _apply_startup(self, enable):
        try:
            key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
            name = "RestoreGuardPro"
            cmd = f'"{sys.executable}" "{Path(__file__).resolve()}"'
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as k:
                if enable: winreg.SetValueEx(k, name, 0, winreg.REG_SZ, cmd)
                else:
                    try: winreg.DeleteValue(k, name)
                    except FileNotFoundError: pass
        except Exception as e:
            log(f"Error startup: {e}")

    # ── REFRESH ───────────────────────────────────────────────────
    def refresh(self):
        history = load_history()
        cfg = load_config()

        if history:
            self._s_last.configure(text=history[0]["fecha"][:16], text_color=SUCCH)
            self._s_total.configure(text=f"{len(history)} puntos", text_color=ACCENT)
        else:
            self._s_last.configure(text="Sin puntos aún", text_color=TSEC)
            self._s_total.configure(text="0 puntos", text_color=TSEC)

        auto_en = cfg.get("auto_enabled", True)
        self._s_next.configure(
            text=f"Hoy {cfg.get('auto_time','08:00')}" if auto_en else "Desactivado",
            text_color=WARN if auto_en else TSEC)

        for w in self._recent.winfo_children(): w.destroy()
        if not history:
            lbl(self._recent,"Sin actividad aún",11,color=TSEC).grid(row=0,column=0,sticky="w")
        for i,e in enumerate(history[:4]):
            ico = "🤖" if e.get("tipo")=="AUTO" else "👆"
            t = e.get("fecha","")[:16]
            estado = e.get("estado","")
            lbl(self._recent, f"{ico}  {t}   {e.get('tipo','')}   {estado}",
                11).grid(row=i,column=0,sticky="w",pady=1)

        d_path = Path(cfg.get("backup_path_d", str(BACKUP_D)))
        self._loc_c.configure(text="✅ OK" if BASE_DIR.exists() else "⚠️ No existe",
            text_color=SUCCH if BASE_DIR.exists() else WARN)
        self._loc_d.configure(text="✅ OK" if d_path.exists() else "⚠️ No existe",
            text_color=SUCCH if d_path.exists() else WARN)

        for w in self._hist_sc.winfo_children(): w.destroy()
        if not history:
            lbl(self._hist_sc,"Sin puntos registrados aún",12,color=TSEC
                ).grid(row=0,column=0,columnspan=4,pady=30)
        for i,e in enumerate(history):
            bg = CARD if i%2==0 else CARD2
            tipo = e.get("tipo","?")
            rf = ctk.CTkFrame(self._hist_sc, fg_color=bg, corner_radius=0, height=42)
            rf.grid(row=i,column=0,columnspan=4,sticky="ew")
            rf.grid_propagate(False)
            rf.grid_columnconfigure((0,1,2), weight=1)
            lbl(rf,e.get("fecha","")[:16],11).grid(row=0,column=0,padx=12,sticky="w")
            lbl(rf,tipo,11,"bold",ACCENT if tipo=="AUTO" else SUCCH
                ).grid(row=0,column=1,padx=12,sticky="w")
            desc = e.get("description","")
            desc = ("…"+desc[-24:]) if len(desc)>24 else desc
            lbl(rf,desc,10,color=TSEC).grid(row=0,column=2,padx=12,sticky="w")
            lbl(rf,e.get("estado",""),11).grid(row=0,column=3,padx=12,sticky="w")
            
            btn = ctk.CTkButton(rf, text="Restaurar", width=70, height=26, fg_color=WARN, hover_color=DANGER, font=ctk.CTkFont(size=11, weight="bold"), command=lambda e_=e: RestoreWizard(self, e_))
            btn.grid(row=0, column=4, padx=12)

    # ── TOAST ─────────────────────────────────────────────────────
    def _toast(self, msg):
        t = ctk.CTkToplevel(self)
        t.overrideredirect(True)
        t.attributes("-topmost", True)
        t.configure(fg_color=CARD)
        lbl(t, msg, 12, "bold").pack(padx=24, pady=16)
        t.update_idletasks()
        x = self.winfo_x() + self.winfo_width()//2 - t.winfo_width()//2
        y = self.winfo_y() + self.winfo_height() - 100
        t.geometry(f"+{x}+{y}")
        t.after(2800, t.destroy)

    # ── TRAY ─────────────────────────────────────────────────────
    def _start_tray(self):
        img = Image.new("RGBA",(64,64),(0,0,0,0))
        d = ImageDraw.Draw(img)
        d.ellipse([0,0,63,63],fill=(61,126,255,255))
        d.polygon([(32,8),(54,20),(54,40),(32,57),(10,40),(10,20)],fill=(255,255,255,220))
        d.polygon([(32,15),(47,24),(47,38),(32,50),(17,38),(17,24)],fill=(61,126,255,255))
        d.line([(23,33),(30,41)],fill=(255,255,255),width=4)
        d.line([(30,41),(43,24)],fill=(255,255,255),width=4)

        menu = pystray.Menu(
            pystray.MenuItem("📋 Abrir",
                lambda i,it: self.after(0,lambda:(self.deiconify(),self.lift())), default=True),
            pystray.MenuItem("🛡️ Crear Punto",
                lambda i,it: self.after(0, self._on_manual)),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("❌ Salir",
                lambda i,it: (i.stop(), self.after(0, self.destroy))),
        )
        pystray.Icon("RG", img, "RestoreGuard Pro", menu).run()

# ─────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    try:
        request_admin()
        App().mainloop()
    except Exception as e:
        print(f"Error: {e}")
        try:
            import traceback
            with open(BASE_DIR / "crash.log", "w", encoding="utf-8") as f:
                f.write(traceback.format_exc())
        except:
            pass


