#!/usr/bin/env python
# -*- coding: utf-8 -*-

# ROBOT - JUSTIFICACIONES (Local, visible en Chrome, Flask + SocketIO + Playwright ASYNC)
# UI en http://localhost:8767
# Motor ÚNICO: AsyncRobot (Playwright async_api) para evitar errores de hilos del sync_api.

import os
import re
import json
import atexit
import subprocess
from dataclasses import dataclass
from typing import List, Optional, Any

from flask import Flask, request, jsonify, render_template
from flask_socketio import SocketIO
from playwright.async_api import TimeoutError as PlaywrightTimeoutError

from robot_async import AsyncRobot

# =========================
# Configuración general
# =========================

PORT = 8771
URL_PORTAL = "https://portal.gestion.sedepkd.red.gob.es/portal/espacioAD"
HEADLESS = False            # Navegador visible
BROWSER_CHANNEL = "chrome"  # Lanzar Google Chrome real
DEFAULT_SPEED = "medio"     # rapido | medio | lento

DELAY_PRESETS = {
    "rapido": 0.25,
    "medio": 0.6,
    "lento": 1.2,
}

ENGINE_TAG = "async-only-1"

# Selectores de la plataforma (según HTML facilitado)
SEL = {
    "menu_tramitacion": "#navbarTramitAcuerdos",
    "kd_justificaciones": "#kd-jus",
    "kc_justificaciones": "#kc-jus",
    "btn_clave_cert": "button:has-text(\"Acceso DNIe / Certificado electrónico\")",
    "btn_aceptar_clave": "button:has-text(\"Aceptar\")",
    "btn_advanced_search": "#advancedSearch",
    "table_justificaciones": "#tableJustificaciones",
    "btn_presentar": "#submitForm",
    "btn_firma_clave": "#submitfirmaNotCrypto",
    "paginate_next": "#tableJustificaciones_next",
}

# =========================
# App Flask + SocketIO
# =========================

app = Flask(__name__, template_folder="templates")
app.config["TEMPLATES_AUTO_RELOAD"] = True
app.jinja_env.auto_reload = True
socketio = SocketIO(app, cors_allowed_origins="*", async_mode="threading")

@app.after_request
def add_no_cache_headers(response):
    if response.mimetype == "text/html":
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    return response

STATE = {
    "server": "Conectado",
    "robot": "Detenido",
    "speed": DEFAULT_SPEED,
}

def emit_status():
    socketio.emit("status", STATE)

def log(msg: str):
    safe = (msg or "").replace("\r", "").replace("\x00", "")
    socketio.emit("log", {"message": safe})

# =========================
# Certificados (Windows)
# =========================

CN_REGEX = re.compile(r"CN=([^,]+)")

@dataclass
class CertInfo:
    cn: str
    issuer: str
    serial: str   # Número de serie (hex)
    thumbprint: str
    not_after: str

def _powershell_json(cmd: str) -> Any:
    completed = subprocess.run(
        ["powershell", "-NoProfile", "-Command", cmd],
        capture_output=True, text=True, encoding="utf-8", errors="ignore"
    )
    if completed.returncode != 0:
        raise RuntimeError(completed.stderr.strip() or "Error ejecutando PowerShell")
    out = completed.stdout.strip()
    if not out:
        return []
    try:
        return json.loads(out)
    except json.JSONDecodeError:
        return []

def _parse_cn(subject: str) -> str:
    m = CN_REGEX.search(subject or "")
    if m:
        return m.group(1).strip()
    return (subject or "").strip()

def list_windows_certs() -> List[CertInfo]:
    # Personal del usuario actual: Cert:\CurrentUser\My, no expirados
    ps = r"""
$now = Get-Date
Get-ChildItem -Path Cert:\CurrentUser\My `
| Where-Object { $_.NotAfter -gt $now } `
| Select-Object Subject, Issuer, SerialNumber, Thumbprint, NotAfter `
| ConvertTo-Json -Compress
"""
    data = _powershell_json(ps)
    certs: List[CertInfo] = []
    if isinstance(data, dict):
        data = [data]
    for item in data or []:
        cn = _parse_cn(item.get("Subject", ""))
        issuer = item.get("Issuer", "") or ""
        serial = (item.get("SerialNumber", "") or "").replace(" ", "").upper()
        thumb = (item.get("Thumbprint", "") or "").upper()
        not_after = str(item.get("NotAfter", ""))
        certs.append(CertInfo(cn=cn, issuer=issuer, serial=serial, thumbprint=thumb, not_after=not_after))
    certs.sort(key=lambda c: c.not_after, reverse=True)
    return certs

def get_chrome_path() -> Optional[str]:
    """Obtiene la ruta real de chrome.exe desde el Registro (HKLM/HKCU)."""
    try:
        ps = r"""
$paths = @(
  'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe',
  'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe'
)
foreach ($k in $paths) {
  try {
    $v = (Get-ItemProperty -Path $k -ErrorAction Stop).'(default)'
    if ($v) { Write-Output $v; break }
  } catch {}
}
"""
        completed = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps],
            capture_output=True, text=True, encoding="utf-8", errors="ignore"
        )
        p = (completed.stdout or "").strip()
        return p if p else None
    except Exception:
        return None

# =========================
# Instancia del Motor ASYNC ÚNICO
# =========================

async_robot = AsyncRobot(
    logger=log,
    status_cb=emit_status,
    url_portal=URL_PORTAL,
    selectors=SEL,
    delay_presets=DELAY_PRESETS,
    headless=HEADLESS,
    browser_channel=BROWSER_CHANNEL,
)

@atexit.register
def _cleanup():
    try:
        async_robot.stop()
    except Exception:
        pass

# =========================
# Rutas HTTP / API
# =========================

@app.route("/")
def index():
    return render_template("index.html", port=PORT, engine=ENGINE_TAG)

@app.route("/api/status", methods=["GET"])
def api_status():
    return jsonify(STATE)

@app.route("/api/ping", methods=["GET"])
def api_ping():
    # Comprobación rápida desde el navegador para verificar versión del servidor
    return jsonify({"ok": True, "port": PORT, "engine": ENGINE_TAG})

@app.route("/api/certificates", methods=["GET"])
def api_certificates():
    try:
        certs = list_windows_certs()
        payload = [
            {
                "cn": c.cn,
                "issuer": c.issuer,
                "serial": c.serial,
                "thumbprint": c.thumbprint,
                "not_after": c.not_after,
            }
            for c in certs
        ]
        return jsonify({"ok": True, "certificates": payload})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route("/api/open", methods=["POST"])
def api_open():
    """
    Abre una VISTA PREVIA del portal en el navegador del sistema (nueva ventana si es posible).
    La sesión Playwright del robot se abre al pulsar 'Iniciar Proceso'.
    """
    try:
        data = request.get_json(force=True) or {}
        speed = (data.get("speed", DEFAULT_SPEED) or DEFAULT_SPEED)
        async_robot.set_speed(speed)

        opened = False
        chrome_path = get_chrome_path()
        if chrome_path and os.path.isfile(chrome_path):
            try:
                ps = f'Start-Process -FilePath "{chrome_path}" -ArgumentList \'--new-window\', \'{URL_PORTAL}\''
                subprocess.run(["powershell", "-NoProfile", "-Command", ps], check=False)
                opened = True
            except Exception:
                opened = False

        if not opened:
            try:
                subprocess.Popen(["cmd", "/c", "start", "", "chrome", "--new-window", URL_PORTAL], shell=True)
                opened = True
            except Exception:
                opened = False

        if not opened:
            try:
                import webbrowser
                opened = webbrowser.open(URL_PORTAL, new=2)
            except Exception:
                opened = False

        if not opened:
            try:
                os.startfile(URL_PORTAL)  # type: ignore[attr-defined]
                opened = True
            except Exception:
                opened = False

        STATE["robot"] = "Detenido"
        emit_status()

        if opened:
            log("[Robot] Vista previa abierta en nueva ventana (o pestaña si no fue posible). El robot abrirá su sesión propia al pulsar 'Iniciar Proceso'.")
            return jsonify({"ok": True})
        else:
            log("[Robot] No se pudo abrir la vista previa automáticamente. Abre manualmente: " + URL_PORTAL)
            return jsonify({"ok": False, "error": "No se pudo abrir el navegador automáticamente"}), 500
    except Exception as e:
        log(f"[Robot] Error en /api/open: {e}")
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route("/api/start", methods=["POST"])
def api_start():
    try:
        data = request.get_json(force=True) or {}
        categoria = data.get("categoria", "KD")
        thumb = (data.get("thumbprint", "") or "").upper()
        serial = ""
        cn = ""
        issuer_cn = ""
        for c in list_windows_certs():
            if c.thumbprint.upper() == thumb:
                serial = c.serial
                cn = c.cn
                issuer_cn = _parse_cn(c.issuer)
                break
        if not serial:
            return jsonify({"ok": False, "error": "No se pudo resolver el certificado seleccionado"}), 400

        speed = data.get("speed", DEFAULT_SPEED)
        async_robot.set_speed(speed)
        log("[Async] start: categoría=%s" % categoria)

        STATE["robot"] = "En ejecución"
        emit_status()
        async_robot.start(categoria, serial, speed, cn, issuer_cn)
        return jsonify({"ok": True})
    except Exception as e:
        STATE["robot"] = "Detenido"
        emit_status()
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route("/api/stop", methods=["POST"])
def api_stop():
    try:
        async_robot.stop()
        STATE["robot"] = "Detenido"
        emit_status()
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

# Socket events
@socketio.on("connect")
def on_connect():
    emit_status()

@socketio.on("disconnect")
def on_disconnect():
    pass

# =========================
# Arranque
# =========================

if __name__ == "__main__":
    print(f"UI disponible en http://localhost:{PORT} | ENGINE={ENGINE_TAG}")
    socketio.run(app, host="127.0.0.1", port=PORT, allow_unsafe_werkzeug=True)
