#!/usr/bin/env python
# -*- coding: utf-8 -*-

import asyncio
import re
import json
import time
import threading
from typing import Dict, Any, Callable, Optional
from typing import Optional as Opt
from pywinauto import Desktop, timings
import subprocess
import sys
import os

from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError


class AsyncRobot:
    """
    Motor de automatización Playwright en modo asíncrono, ejecutándose SIEMPRE
    en el mismo hilo con su propio event loop. Los endpoints Flask (sync) envían
    órdenes con run_coroutine_threadsafe para evitar problemas de hilos.
    """

    def __init__(
        self,
        logger: Callable[[str], None],
        status_cb: Callable[[], None],
        url_portal: str,
        selectors: Dict[str, str],
        delay_presets: Dict[str, float],
        headless: bool = False,
        browser_channel: str = "chrome",
    ):
        self.log = logger
        self.emit_status = status_cb
        self.URL_PORTAL = url_portal
        self.SEL = selectors
        self.DELAY_PRESETS = delay_presets
        self.headless = headless
        self.browser_channel = browser_channel

        # Estado
        self.speed = "medio"
        self.delay = self.DELAY_PRESETS.get(self.speed, 0.6)
        self.stop_flag = asyncio.Event()

        # Objetos Playwright
        self._pw = None
        self._browser = None
        self._context = None
        self.page = None

        # Event loop dedicado
        self.loop = asyncio.new_event_loop()
        self._t = threading.Thread(target=self._run_loop, daemon=True)
        self._t.start()
        # Watcher de certificado (nativo) en hilo aparte
        self._cert_watch_stop = threading.Event()
        self._cert_watch_thread: Optional[threading.Thread] = None
        self._cert_watch_active = False

    # ------------- Infra -------------
    def _run_loop(self):
        asyncio.set_event_loop(self.loop)
        self.loop.run_forever()

    def set_speed(self, speed: str):
        if speed not in self.DELAY_PRESETS:
            speed = "medio"
        self.speed = speed
        self.delay = self.DELAY_PRESETS[speed]
        self.log(f"[Robot] Velocidad: {self.speed} (delay {self.delay}s)")

    def start(self, categoria: str, serial: str, speed: Optional[str] = None, cn: Optional[str] = None, issuer_cn: Optional[str] = None):
        if speed:
            self.set_speed(speed)
        # Construir política dinámica de auto-selección (si tenemos datos)
        try:
            filt: Dict[str, Dict[str, str]] = {}
            if cn:
                filt["SUBJECT"] = {"CN": cn}
            if issuer_cn:
                filt["ISSUER"] = {"CN": issuer_cn}
            if filt:
                entries = [
                    {"pattern": "https://pasarela.clave.gob.es", "filter": filt},
                    {"pattern": "https://pasarela.clave.gob.es:443", "filter": filt},
                    {"pattern": "https://*.clave.gob.es", "filter": filt},
                    {"pattern": "https://*.ident.clave.gob.es", "filter": filt},
                ]
                self._auto_select_arg = json.dumps(entries, ensure_ascii=False)
                self.log(f"[Robot] AutoSelectCertificateForUrls (SUBJECT='{cn or ''}' ISSUER='{issuer_cn or ''}')")
            else:
                self._auto_select_arg = None
        except Exception:
            self._auto_select_arg = None
        # Lanzar corrida
        fut = asyncio.run_coroutine_threadsafe(self._run_full(categoria, serial, cn), self.loop)
        return fut

    def stop(self):
        # Señal de parada y cerrar
        asyncio.run_coroutine_threadsafe(self._stop_and_close(), self.loop)

    # ------------- Helpers async -------------
    async def _sleep(self, mult: float = 1.0):
        await asyncio.sleep(self.delay * mult)

    async def _ensure_browser(self):
        if self._browser:
            return
        self.log("[Robot] Iniciando Playwright/Chrome...")
        self._pw = await async_playwright().start()
        args = ["--start-maximized"]
        try:
            if getattr(self, "_auto_select_arg", None):
                args.append(f"--auto-select-certificate-for-urls={self._auto_select_arg}")
        except Exception:
            pass
        self._browser = await self._pw.chromium.launch(
            headless=self.headless,
            channel=self.browser_channel,
            args=args,
        )
        # viewport=None para permitir maximizado real
        self._context = await self._browser.new_context(ignore_https_errors=True, viewport={"width": 1400, "height": 900})
        self.page = await self._context.new_page()

    async def _close(self):
        try:
            if self._context:
                await self._context.close()
            if self._browser:
                await self._browser.close()
            if self._pw:
                await self._pw.stop()
        except Exception:
            pass
        finally:
            self._context = None
            self._browser = None
            self._pw = None
            self.page = None

    async def _stop_and_close(self):
        try:
            self.stop_flag.set()
        except Exception:
            pass
        await self._close()

    # ------------- Pasos de navegación -------------
    async def open_portal(self):
        await self._ensure_browser()
        self.log(f"[Robot] Abriendo portal: {self.URL_PORTAL}")
        await self.page.goto(self.URL_PORTAL, wait_until="networkidle")
        await self._sleep(1.0)

    async def navigate_to_justificaciones(self, categoria: str):
        await self._ensure_browser()
        self.log(f"[Robot] Navegando a Justificaciones ({'Kit Digital' if categoria=='KD' else 'Kit Consulting'})")
        if not self.page or ("portal/espacioAD" not in (self.page.url or "")):
            await self.open_portal()

        await self.page.click(self.SEL["menu_tramitacion"], timeout=15000)
        await self._sleep()

        if categoria == "KD":
            await self.page.click(self.SEL["kd_justificaciones"], timeout=15000)
        else:
            await self.page.click(self.SEL["kc_justificaciones"], timeout=15000)

        await self._sleep(1.2)

    # ---------- UIA (Windows) para diálogo nativo de certificados ----------
    async def _uia_pick_cert(self, serial: str, cn: Optional[str] = None, timeout: float = 10.0) -> bool:
        """
        Intenta seleccionar por UI Automation (pywinauto) el certificado en el diálogo nativo
        y pulsar 'Aceptar'. Coincide por prefijo de número de serie y/o por CN.
        Devuelve True si pudo aceptar el diálogo.
        """
        def norm(s: Optional[str]) -> str:
            return re.sub(r"[^A-Z0-9]", "", (s or "").upper())

        target = norm(serial)
        subj = norm(cn or "")
        timings.Timings.after_clickinput_wait = 0.3

        t0 = time.time()
        dlg = None
        while time.time() - t0 < timeout:
            try:
                wins = Desktop(backend="uia").windows()
                for w in wins:
                    try:
                        title = (w.window_text() or "")
                        cls = (getattr(w, "element_info", None).class_name or "")
                        # Priorizar ventana de Chrome que aloja la pasarela/diálogo
                        if re.search(r"(Cl@ve|clave|pasarela)", title, re.I) or cls == "Chrome_WidgetWin_1":
                            # Buscar descendiente que sea el propio diálogo
                            try:
                                cand = w.child_window(title_re="Seleccionar.*certificado|Select.*certificate", control_type="Window")
                                dlg = cand.wrapper_object()
                            except Exception:
                                dlg = w
                            break
                    except Exception:
                        continue
                if dlg:
                    break
            except Exception:
                pass
            time.sleep(0.3)

        if not dlg:
            return False

        try:
            # Buscar tabla/filas
            try:
                table = dlg.child_window(control_type="Table").wrapper_object()
            except Exception:
                table = None

            rows = []
            if table:
                try:
                    rows = table.children()
                except Exception:
                    rows = []
            if not rows:
                # Fallback: cualquier fila tipo DataItem bajo el diálogo
                try:
                    rows = dlg.descendants(control_type="DataItem")
                except Exception:
                    rows = []

            picked = False
            def row_texts(r):
                try:
                    parts = []
                    for el in r.descendants(control_type="Text"):
                        try:
                            parts.append(el.window_text())
                        except Exception:
                            continue
                    return " ".join(parts)
                except Exception:
                    return ""

            for r in rows:
                txt = row_texts(r)
                n = norm(txt)
                if (target and (target in n or n in target)) or (subj and subj in n):
                    try:
                        r.click_input()
                        picked = True
                        break
                    except Exception:
                        continue

            if not picked:
                return False

            # Botón Aceptar
            try:
                btn = dlg.child_window(title_re="Aceptar|Accept", control_type="Button").wrapper_object()
                btn.click_input()
                return True
            except Exception:
                # Intentar con tecla Enter
                try:
                    dlg.type_keys("{ENTER}")
                    return True
                except Exception:
                    return False
        except Exception:
            return False

    def _run_cert_clicker(self, serial: str, cn: Optional[str], secs: float = 1.0) -> bool:
        """Ejecuta helper local que selecciona y acepta el diálogo nativo con pywinauto (muy rápido)."""
        try:
            # Resolver ruta del script de forma robusta relativa a este archivo
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # .../just-signer
            script = os.path.join(base_dir, "tools", "cert_clicker.py")
            if not os.path.isfile(script):
                return False
            cn_arg = cn or ""
            t_int = max(1, int(secs))
            proc = subprocess.run(
                [sys.executable, script, "--serial", serial, "--cn", cn_arg, "--timeout", str(t_int)],
                capture_output=True, text=True, timeout=t_int + 1
            )
            out = (proc.stdout or "") + (proc.stderr or "")
            self.log(f"[Watcher] cert_clicker rc={proc.returncode} | {out.strip()[:160]}")
            return proc.returncode == 0
        except subprocess.TimeoutExpired:
            self.log("[Watcher] cert_clicker timeout")
            return False
        except Exception as e:
            self.log(f"[Watcher] cert_clicker error: {e}")
            return False

    # -------- Watcher residente: cierra el diálogo en cuanto aparece --------
    def _cert_watcher(self, serial: str, cn: Optional[str]):
        try:
            self._cert_watch_active = True
            t0 = time.time()
            attempt = 0
            # Bucle de vigilancia muy rápido (~ 12 s)
            while not self._cert_watch_stop.is_set() and not self.stop_flag.is_set():
                if time.time() - t0 > 12:
                    break
                attempt += 1
                try:
                    ok = self._run_cert_clicker(serial, cn, secs=1.0)
                    if ok:
                        self.log("[Watcher] Diálogo aceptado (helper).")
                        break
                except Exception:
                    pass
                # Pequeña espera para no saturar CPU
                time.sleep(0.15)
        finally:
            self._cert_watch_active = False

    def _start_cert_watcher(self, serial: str, cn: Optional[str]):
        try:
            self._cert_watch_stop.clear()
            if self._cert_watch_thread and self._cert_watch_thread.is_alive():
                return
            th = threading.Thread(target=self._cert_watcher, args=(serial, cn), daemon=True)
            self._cert_watch_thread = th
            th.start()
            self.log("[Watcher] Activado.")
        except Exception as e:
            self.log(f"[Watcher] Error al activar: {e}")

    def _stop_cert_watcher(self):
        try:
            self._cert_watch_stop.set()
            if self._cert_watch_thread and self._cert_watch_thread.is_alive():
                self._cert_watch_thread.join(timeout=2.0)
            self._cert_watch_thread = None
            self.log("[Watcher] Detenido.")
        except Exception as e:
            self.log(f"[Watcher] Error al detener: {e}")

    async def _detect_clave(self) -> bool:
        url = (self.page.url or "").lower()
        if "pasarela.clave.gob.es" in url:
            return True
        try:
            loc = self.page.locator("text=Plataforma de identificación").first
            await loc.wait_for(state="visible", timeout=1000)
            return True
        except Exception:
            return False

    async def _select_cert_in_clave(self, serial: str, cn: Optional[str] = None):
        """
        Selecciona certificado en Cl@ve:
        - Maneja números de serie truncados en la tabla (solo prefijo visible).
        - Fallback por CN si por serie no se localiza.
        - Busca también dentro de iframes del diálogo si los hubiese.
        """
        def norm(s: Opt[str]) -> str:
            return re.sub(r"[^A-Z0-9]", "", (s or "").upper())

        target = norm(serial)
        if not target:
            raise RuntimeError("Número de serie no proporcionado")

        self.log("[Robot] Seleccionando certificado en Cl@ve por Número de serie...")
        # Intento 0: carrera rápida (helper nativo + UIA) en bucle corto para evitar timeouts de la pasarela
        try:
            for _ in range(8):  # ~ 8 * (0.3s + llamadas rápidas) ≈ 3-4s
                ok_helper = await asyncio.to_thread(self._run_cert_clicker, serial, cn, 2.0)
                if ok_helper:
                    self.log("[Robot] Certificado seleccionado por helper local y aceptado (rápido).")
                    await self._sleep(0.6)
                    return
                ok_native = await self._uia_pick_cert(serial, cn, timeout=2.0)
                if ok_native:
                    self.log("[Robot] Certificado seleccionado por diálogo nativo (UIA) y aceptado (rápido).")
                    await self._sleep(0.6)
                    return
                await asyncio.sleep(0.3)
        except Exception:
            pass

        # Intentar pulsar el botón de acceso por certificado (si existe en DOM)
        try:
            await self.page.click(self.SEL["btn_clave_cert"], timeout=6000)
            # Carrera inmediata tras abrir la opción
            try:
                for _ in range(10):  # ~ 10 * (0.25s + llamadas rápidas) ≈ 4-5s
                    ok_helper2 = await asyncio.to_thread(self._run_cert_clicker, serial, cn, 2.0)
                    if ok_helper2:
                        self.log("[Robot] Certificado seleccionado por helper local tras abrir opción de certificado.")
                        await self._sleep(0.6)
                        return
                    ok_native2 = await self._uia_pick_cert(serial, cn, timeout=2.0)
                    if ok_native2:
                        self.log("[Robot] Certificado seleccionado por diálogo nativo (UIA) tras abrir opción de certificado.")
                        await self._sleep(0.6)
                        return
                    await asyncio.sleep(0.25)
            except Exception:
                pass
        except Exception:
            pass

        # Buscar un frame que contenga el título del diálogo
        scope = self.page
        try:
            for fr in self.page.frames:
                try:
                    if await fr.get_by_text("Seleccionar un certificado", exact=False).count() > 0:
                        scope = fr
                        break
                except Exception:
                    continue
        except Exception:
            pass

        # Localizar filas dentro del diálogo de selección (en scope deducido)
        rows = scope.locator("xpath=//*[contains(@role,'dialog') or contains(@class,'modal')]//table//tbody/tr")
        # Fallback si no detecta diálogo/clase: usar cualquier tabla en la página
        try:
            if await rows.count() == 0:
                rows = scope.locator("table tbody tr")
        except Exception:
            rows = scope.locator("table tbody tr")
        try:
            await rows.first.wait_for(timeout=8000)
        except Exception:
            # Fallback: búsqueda textual por prefijo de serie en todo el scope
            pref16 = serial[:16]
            pref12 = serial[:12]
            try:
                await scope.get_by_text(pref16, exact=False).first.click(timeout=5000)
                await scope.get_by_role("button", name=re.compile("Aceptar", re.I)).first.click(timeout=8000)
                await self._sleep(2.0)
                return
            except Exception:
                try:
                    await scope.get_by_text(pref12, exact=False).first.click(timeout=5000)
                    await scope.get_by_role("button", name=re.compile("Aceptar", re.I)).first.click(timeout=8000)
                    await self._sleep(2.0)
                    return
                except Exception:
                    raise RuntimeError(f"No se encontró el certificado con Nº de serie {serial} en Cl@ve")

        found = False
        count = await rows.count()
        serial_hits = []
        # Primero intentar por número de serie (columna 3)
        for i in range(count):
            row = rows.nth(i)
            try:
                cell_serial = row.locator("td").nth(2)
                row_serial_raw = await cell_serial.inner_text(timeout=3000)
                row_serial = norm(row_serial_raw)
                serial_hits.append(row_serial)
                # Coincidencia si una es prefijo de la otra o contienen el mismo prefijo
                if row_serial and (target.startswith(row_serial) or row_serial.startswith(target) or (row_serial in target) or (target in row_serial)):
                    await row.click()
                    found = True
                    break
                # Coincidencia por prefijo visible (primeros 12–16 chars)
                if target[:16] and target[:16] in row_serial:
                    await row.click()
                    found = True
                    break
            except Exception:
                continue

        # Si no coincidió por serie, probar por CN (columna 1) y por texto directo
        if not found and cn:
            cn_norm = norm(cn)
            for i in range(count):
                row = rows.nth(i)
                try:
                    cell_cn = row.locator("td").nth(0)
                    row_cn_raw = await cell_cn.inner_text(timeout=3000)
                    row_cn = norm(row_cn_raw)
                    if cn_norm and (cn_norm in row_cn or row_cn in cn_norm):
                        await row.click()
                        found = True
                        break
                except Exception:
                    continue
            if not found:
                try:
                    await scope.get_by_text(cn, exact=False).first.click(timeout=4000)
                    found = True
                except Exception:
                    pass

        if not found:
            self.log(f"[Robot] Series en Cl@ve (normalizadas): {', '.join([s[:12]+'...' for s in serial_hits if s])}")
            # Último intento por prefijo de serie en todo el scope
            for pref in [serial[:20], serial[:16], serial[:12]]:
                try:
                    if pref:
                        await scope.get_by_text(pref, exact=False).first.click(timeout=3000)
                        found = True
                        break
                except Exception:
                    continue

        if not found:
            # Fallback adicional: buscar cualquier celda que contenga un prefijo del número de serie
            picked = False
            for pref in [serial[:20], serial[:18], serial[:16], serial[:14], serial[:12], serial[:10]]:
                pref = (pref or "").strip()
                if not pref:
                    continue
                try:
                    td = scope.locator(f"xpath=//td[contains(normalize-space(), '{pref}')]").first
                    await td.click(timeout=3000)
                    picked = True
                    break
                except Exception:
                    continue

            # Fallback por CN si aún no
            if not picked and cn:
                try:
                    await scope.get_by_text(cn, exact=False).first.click(timeout=4000)
                    picked = True
                except Exception:
                    pass

            # Último recurso: seleccionar la primera fila si existe
            if not picked:
                try:
                    await rows.first.click(timeout=2000)
                    picked = True
                except Exception:
                    picked = False

            # Verificación estricta: si hay una fila seleccionada y coincide con nuestro certificado, aceptamos
            if not picked:
                try:
                    selected = scope.locator("xpath=//tr[contains(@class,'selected') or @aria-selected='true']").first
                    await selected.wait_for(timeout=1000)
                    sel_cn_raw = await selected.locator("td").nth(0).inner_text(timeout=2000)
                    sel_cn = norm(sel_cn_raw)
                    sel_serial_raw = await selected.locator("td").nth(2).inner_text(timeout=2000)
                    sel_serial = norm(sel_serial_raw)
                    self.log(f"[Robot] Fila seleccionada: CN='{sel_cn_raw.strip()}' Serie='{sel_serial_raw.strip()}'")
                    if (sel_serial and (target.startswith(sel_serial) or sel_serial.startswith(target) or (sel_serial in target) or (target in sel_serial))) or (cn and sel_cn and (norm(cn) in sel_cn or sel_cn in norm(cn))):
                        picked = True
                except Exception:
                    pass

            if not picked:
                raise RuntimeError(f"No se encontró el certificado con Nº de serie {serial} en Cl@ve")

        # Aceptar
        try:
            await scope.get_by_role("button", name=re.compile("Aceptar", re.I)).first.click(timeout=10000)
        except Exception:
            try:
                await self.page.click(self.SEL["btn_aceptar_clave"], timeout=10000)
            except Exception:
                # Enter como última opción
                try:
                    await self.page.keyboard.press("Enter")
                except Exception:
                    pass
        await self._sleep(2.0)

    async def authenticate_if_needed(self, serial: str, cn: Optional[str] = None):
        if await self._detect_clave():
            self.log("[Robot] Autenticación requerida en Cl@ve")
            # Forzar velocidad rápida durante autenticación
            prev_speed = self.speed
            try:
                self.set_speed("rapido")
            except Exception:
                prev_speed = "medio"
            # Arrancar watcher residente (cierre inmediato del diálogo si aparece)
            self._start_cert_watcher(serial, cn)
            try:
                await self._select_cert_in_clave(serial, cn)
                await self.page.wait_for_load_state("networkidle", timeout=30000)
                if await self._detect_clave():
                    raise RuntimeError("No fue posible completar la autenticación en Cl@ve")
                self.log("[Robot] Autenticado en Cl@ve")
            finally:
                # Detener watcher y restaurar velocidad
                self._stop_cert_watcher()
                try:
                    self.set_speed(prev_speed)
                except Exception:
                    pass

    async def _scan_rows_and_open_first_pending(self) -> bool:
        table = self.page.locator(self.SEL["table_justificaciones"])
        await table.wait_for(timeout=20000)

        rows = table.locator("tbody tr")
        count = await rows.count()
        for i in range(count):
            if self.stop_flag.is_set():
                return False
            row = rows.nth(i)
            try:
                estado_text = (await row.locator("td").nth(3).inner_text(timeout=5000)).strip()
            except Exception:
                continue
            if "Pdte. presentar" in estado_text:
                link = row.locator("td").nth(0).locator("a")
                if await link.count() > 0:
                    cod = (await link.inner_text()).strip()
                    self.log(f"[Robot] Abriendo expediente {cod} (Estado: {estado_text})")
                    await link.click(timeout=10000)
                    await self.page.wait_for_load_state("domcontentloaded", timeout=20000)
                    await self._sleep(0.8)
                    return True
        return False

    async def _go_back_to_list(self):
        try:
            await self.page.go_back(timeout=15000)
            await self.page.wait_for_load_state("networkidle", timeout=15000)
        except Exception:
            await self.page.reload(wait_until="networkidle")

    async def _try_presentar_autofirma(self) -> bool:
        try:
            if await self.page.locator(self.SEL["btn_presentar"]).count() == 0:
                return False
            await self.page.click(self.SEL["btn_presentar"], timeout=8000)
            self.log("[Robot] Pulsado 'Presentar' (AutoFirma). Esperando resultado...")
            for _ in range(40):
                await asyncio.sleep(0.5)
                if await self.page.locator(self.SEL["btn_presentar"]).count() == 0:
                    self.log("[Robot] 'Presentar' ya no está visible (posible éxito).")
                    return True
                try:
                    if await self.page.locator(self.SEL["btn_presentar"]).first.is_disabled():
                        self.log("[Robot] 'Presentar' deshabilitado (posible éxito).")
                        return True
                except Exception:
                    pass
            self.log("[Robot] No se detectó progreso suficiente con AutoFirma en tiempo razonable.")
            return False
        except Exception as e:
            self.log(f"[Robot] Error al intentar 'Presentar': {e}")
            return False

    async def _try_firma_clave(self, serial: str) -> bool:
        try:
            if await self.page.locator(self.SEL["btn_firma_clave"]).count() == 0:
                return False
            await self.page.click(self.SEL["btn_firma_clave"], timeout=8000)
            self.log("[Robot] Pulsado 'Firma con Cl@ve y presentar'.")
            await self._sleep(1.0)
            if await self._detect_clave():
                await self._select_cert_in_clave(serial)
                await self.page.wait_for_load_state("networkidle", timeout=30000)
            for _ in range(30):
                await asyncio.sleep(0.5)
                if (
                    await self.page.locator(self.SEL["btn_firma_clave"]).count() == 0
                    and await self.page.locator(self.SEL["btn_presentar"]).count() == 0
                ):
                    self.log("[Robot] Botones de firma ya no visibles (posible éxito).")
                    return True
                try:
                    p_disabled = await self.page.locator(self.SEL["btn_presentar"]).first.is_disabled()
                except Exception:
                    p_disabled = True
                try:
                    c_disabled = await self.page.locator(self.SEL["btn_firma_clave"]).first.is_disabled()
                except Exception:
                    c_disabled = True
                if p_disabled and c_disabled:
                    self.log("[Robot] Botones deshabilitados (posible éxito).")
                    return True
            self.log("[Robot] No se detectó confirmación clara tras firmar con Cl@ve.")
            return False
        except Exception as e:
            self.log(f"[Robot] Error al intentar 'Firma con Cl@ve y presentar': {e}")
            return False

    async def _sign_current_record(self, serial: str) -> bool:
        if await self._try_presentar_autofirma():
            return True
        if await self._try_firma_clave(serial):
            return True
        return False

    async def _run_full(self, categoria: str, serial: str, cn: Optional[str] = None):
        try:
            self.stop_flag.clear()
            await self._close()  # garantizar estado limpio
            await self.navigate_to_justificaciones(categoria)
            await self.authenticate_if_needed(serial, cn)

            total_firmados = 0
            while not self.stop_flag.is_set():
                self.log("[Robot] Buscando expedientes 'Pdte. presentar' en la página actual...")
                opened = await self._scan_rows_and_open_first_pending()
                if not opened:
                    try:
                        next_btn = self.page.locator(self.SEL["paginate_next"]).first
                        classes = (await next_btn.get_attribute("class")) or ""
                        if "disabled" in classes:
                            self.log("[Robot] No hay más páginas ni expedientes 'Pdte. presentar'. Proceso completado.")
                            break
                        await next_btn.click(timeout=8000)
                        await self._sleep(0.8)
                        continue
                    except Exception:
                        self.log("[Robot] No hay más 'Pdte. presentar' en esta página. Proceso completado.")
                        break

                self.log("[Robot] Intentando firmar expediente (Preferencia: Presentar/AutoFirma)...")
                ok = await self._sign_current_record(serial)
                if ok:
                    total_firmados += 1
                    self.log("[Robot] Expediente firmado. Volviendo al listado...")
                else:
                    self.log("[Robot] No fue posible completar la firma. Continuando con el siguiente.")
                await self._go_back_to_list()
                await self.page.wait_for_load_state("networkidle", timeout=20000)
                await self._sleep(0.8)

            self.log(f"[Robot] Fin del proceso. Expedientes firmados: {total_firmados}")
        except PlaywrightTimeoutError as te:
            self.log(f"[Robot] Timeout: {te}")
        except Exception as e:
            self.log(f"[Robot] Error durante la ejecución: {e}")
        finally:
            # Devolver estado al servidor
            try:
                self.emit_status()
            except Exception:
                pass
            await self._close()
