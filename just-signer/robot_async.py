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
        headless: bool = False,
        browser_channel: str = "chrome",
    ):
        self.log = logger
        self.emit_status = status_cb
        self.URL_PORTAL = url_portal
        self.SEL = selectors
        self.headless = headless
        self.browser_channel = browser_channel

        # Estado
        self.delay = 0.15  # Delay mínimo inteligente para dar tiempo a JavaScript/animaciones
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

    def start(self, categoria: str, serial: str, cn: Optional[str] = None, issuer_cn: Optional[str] = None):
        # Construir política dinámica de auto-selección (si tenemos datos)
        try:
            filt: Dict[str, Dict[str, str]] = {}
            
            # CAMBIO CRÍTICO: NO usar SERIALNUMBER en el filtro porque Chrome no lo respeta bien
            # En su lugar, usar solo CN que es más confiable, y reforzar la selección manual
            # con el número de serie en los helpers nativos (cert_clicker y UIA)
            if cn:
                filt["SUBJECT"] = {"CN": cn}
                self.log(f"[Robot] Configurando certificado específico: CN='{cn}', Serial='{serial[:16]}...'")
            elif serial:
                # Si no hay CN, intentar solo con serial (aunque es menos confiable)
                serial_formatted = ':'.join([serial[i:i+2] for i in range(0, len(serial), 2)])
                filt["SUBJECT"] = {"SERIALNUMBER": serial_formatted}
                self.log(f"[Robot] Configurando certificado por Serial: '{serial[:16]}...'")
                
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
                self.log(f"[Robot] AutoSelectCertificateForUrls configurado con CN del certificado")
            else:
                self._auto_select_arg = None
        except Exception as e:
            self.log(f"[Robot] Error configurando auto-selección: {e}")
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
                auto_select_arg = self._auto_select_arg
                args.append(f"--auto-select-certificate-for-urls={auto_select_arg}")
                self.log(f"[Robot] Flag de auto-selección aplicado: {auto_select_arg[:200]}...")
        except Exception as e:
            self.log(f"[Robot] No se pudo aplicar flag de auto-selección: {e}")
        
        self.log(f"[Robot] Argumentos de Chrome: {args}")
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
        
        # Detectar y corregir error "Not Found" si aparece
        await self._detect_and_fix_not_found()

    # ---------- UIA (Windows) para diálogo nativo de certificados ----------
    async def _uia_pick_cert(self, serial: str, cn: Optional[str] = None, timeout: float = 10.0) -> bool:
        """
        Intenta seleccionar por UI Automation (pywinauto) el certificado en el diálogo nativo
        y pulsar 'Aceptar'. DEBE buscar y seleccionar activamente el certificado correcto.
        Devuelve True si pudo seleccionar el certificado correcto y aceptar el diálogo.
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
            self.log("[UIA] No se encontró el diálogo de certificado")
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

            if not rows:
                self.log("[UIA] No se encontraron filas en el diálogo")
                return False

            # CRÍTICO: Buscar activamente el certificado correcto por serial Y CN
            picked = False
            matched_row = None
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

            self.log(f"[UIA] Buscando certificado: Serial={serial[:16]}... CN={cn or 'N/A'}")
            
            for i, r in enumerate(rows):
                txt = row_texts(r)
                n = norm(txt)
                
                # Prioridad 1: Coincidir por número de serie (más específico)
                if target and (target in n or n in target or target[:16] in n):
                    matched_row = r
                    self.log(f"[UIA] Fila {i} coincide por serial: {txt[:100]}")
                    break
                
                # Prioridad 2: Coincidir por CN si no hay match por serial
                if subj and subj in n:
                    if matched_row is None:
                        matched_row = r
                        self.log(f"[UIA] Fila {i} coincide por CN: {txt[:100]}")

            if not matched_row:
                self.log(f"[UIA] NO se encontró el certificado correcto entre {len(rows)} filas")
                return False

            # Seleccionar activamente la fila correcta
            try:
                matched_row.click_input()
                picked = True
                time.sleep(0.3)  # Pequeña espera para asegurar selección
                self.log("[UIA] Certificado correcto seleccionado")
            except Exception as e:
                self.log(f"[UIA] Error al hacer clic en la fila: {e}")
                return False

            if not picked:
                self.log("[UIA] No se pudo seleccionar el certificado")
                return False

            # Botón Aceptar
            try:
                btn = dlg.child_window(title_re="Aceptar|Accept", control_type="Button").wrapper_object()
                btn.click_input()
                self.log("[UIA] Botón Aceptar pulsado")
                return True
            except Exception:
                # Intentar con tecla Enter
                try:
                    dlg.type_keys("{ENTER}")
                    self.log("[UIA] Enter pulsado")
                    return True
                except Exception as e:
                    self.log(f"[UIA] Error al aceptar: {e}")
                    return False
        except Exception as e:
            self.log(f"[UIA] Error general: {e}")
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
        
        # PASO 1: Hacer clic en el botón "Acceso DNIe / Certificado electrónico" PRIMERO
        try:
            self.log("[Robot] Haciendo clic en 'Acceso DNIe / Certificado electrónico'...")
            await self.page.click(self.SEL["btn_clave_cert"], timeout=6000)
            await self._sleep(0.5)  # Pequeña espera para que aparezca el diálogo
        except Exception as e:
            self.log(f"[Robot] No se pudo hacer clic en el botón de certificado: {e}")
            # Si no existe el botón, tal vez ya está en la pantalla de selección
            pass
        
        # PASO 2: Ahora sí, intentar métodos rápidos (helper nativo + UIA) 
        # porque el diálogo ya debería estar abierto
        try:
            for _ in range(10):  # ~ 10 * (0.25s + llamadas rápidas) ≈ 4-5s
                ok_helper = await asyncio.to_thread(self._run_cert_clicker, serial, cn, 2.0)
                if ok_helper:
                    self.log("[Robot] Certificado seleccionado por helper local y aceptado.")
                    await self._sleep(0.6)
                    return
                ok_native = await self._uia_pick_cert(serial, cn, timeout=2.0)
                if ok_native:
                    self.log("[Robot] Certificado seleccionado por diálogo nativo (UIA) y aceptado.")
                    await self._sleep(0.6)
                    return
                await asyncio.sleep(0.25)
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
                self.log(f"[Robot] ✗ NO se encontró el certificado correcto entre {len(serial_hits)} certificados disponibles")
                return False  # Retornar False en lugar de lanzar error

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
        return True  # Éxito al seleccionar y aceptar el certificado

    async def authenticate_if_needed(self, serial: str, cn: Optional[str] = None, categoria: Optional[str] = None, max_retries: int = 2):
        if await self._detect_clave():
            self.log("[Robot] Autenticación requerida en Cl@ve")
            
            for attempt in range(1, max_retries + 1):
                # Arrancar watcher residente (cierre inmediato del diálogo si aparece)
                self._start_cert_watcher(serial, cn)
                try:
                    self.log(f"[Robot] Intento de autenticación {attempt}/{max_retries}")
                    cert_selected = await self._select_cert_in_clave(serial, cn)
                    
                    if cert_selected is False:
                        # Certificado no encontrado en la lista
                        self.log(f"[Robot] ⚠ Certificado no encontrado en intento {attempt}/{max_retries}")
                        if attempt < max_retries and categoria:
                            self.log("[Robot] 🔄 Navegando a página de inicio para reintentar...")
                            # Cerrar diálogo de certificados si está abierto
                            try:
                                await self.page.keyboard.press("Escape")
                                await self._sleep(0.5)
                            except Exception:
                                pass
                            # Navegar a justificaciones desde inicio
                            await self.navigate_to_justificaciones(categoria)
                            await self._sleep(1.0)
                            continue  # Reintentar
                        else:
                            # Agotados los reintentos o no hay categoría
                            raise RuntimeError(f"No se encontró el certificado correcto después de {max_retries} intentos")
                    
                    # Certificado seleccionado correctamente
                    await self.page.wait_for_load_state("networkidle", timeout=30000)
                    if await self._detect_clave():
                        raise RuntimeError("No fue posible completar la autenticación en Cl@ve")
                    self.log(f"[Robot] ✓ Autenticado en Cl@ve (intento {attempt}/{max_retries})")
                    break  # Éxito, salir del bucle
                    
                finally:
                    # Detener watcher
                    self._stop_cert_watcher()

    async def _scan_rows_and_open_first_pending(self) -> bool:
        try:
            table = self.page.locator(self.SEL["table_justificaciones"])
            await table.wait_for(timeout=20000)
            
            # CRÍTICO: Esperar activamente a que la tabla tenga contenido
            # La tabla puede estar visible pero sin filas después de una recarga/re-renderizado
            rows = table.locator("tbody tr")
            max_wait_attempts = 30  # 30 intentos × 0.5s = 15 segundos máximo
            count = 0
            
            self.log("[Robot] Esperando a que la tabla tenga contenido...")
            for attempt in range(max_wait_attempts):
                count = await rows.count()
                if count > 0:
                    self.log(f"[Robot] ✓ Tabla lista con {count} expedientes (intento {attempt + 1}/{max_wait_attempts})")
                    break
                if attempt < max_wait_attempts - 1:
                    await asyncio.sleep(0.5)
                else:
                    self.log(f"[Robot] ⚠ Advertencia: Tabla sin filas después de {max_wait_attempts * 0.5}s")
            
            if count == 0:
                self.log("[Robot] ⚠ La tabla no tiene filas después de esperar. No hay expedientes para procesar.")
                return False
            
            self.log(f"[Robot] Procesando {count} expedientes en la tabla")
            
            for i in range(count):
                if self.stop_flag.is_set():
                    return False
                row = rows.nth(i)
                try:
                    # Leer todas las celdas de la fila primero
                    cells = row.locator("td")
                    cell_count = await cells.count()
                    self.log(f"[Robot] Fila {i} tiene {cell_count} columnas")
                    
                    # Leer el código del expediente (columna 0)
                    link = cells.nth(0).locator("a")
                    cod = ""
                    if await link.count() > 0:
                        cod = (await link.inner_text(timeout=10000)).strip()
                    else:
                        cod = (await cells.nth(0).inner_text(timeout=10000)).strip()
                    
                    # Leer el estado - puede estar en columna 3, 4, o la última
                    estado_text = ""
                    # Intentar leer todas las columnas para encontrar el estado
                    for col_idx in range(cell_count):
                        try:
                            texto = (await cells.nth(col_idx).inner_text(timeout=10000)).strip()
                            if "Pdte" in texto or "presentar" in texto or "Presentado" in texto or "Aprobado" in texto:
                                estado_text = texto
                                self.log(f"[Robot] Estado encontrado en columna {col_idx}: '{estado_text}'")
                                break
                        except Exception:
                            continue
                    
                    if not estado_text:
                        self.log(f"[Robot] Expediente {cod}: No se pudo leer el estado, saltando...")
                        continue
                    else:
                        self.log(f"[Robot] Expediente {cod}: Estado='{estado_text}'")
                        
                        # CRÍTICO: Filtro estricto - SOLO "Pdte. presentar" exacto
                        if estado_text == "Pdte. presentar":
                            if await link.count() > 0:
                                self.log(f"[Robot] Abriendo expediente {cod} (Estado: {estado_text})")
                                await link.click(timeout=10000)
                                # CRÍTICO: Esperar a networkidle, no domcontentloaded
                                await self.page.wait_for_load_state("networkidle", timeout=30000)
                                await self._sleep(0.8)
                                return True
                        else:
                            self.log(f"[Robot] Expediente {cod} saltado: Estado no es 'Pdte. presentar' (es '{estado_text}')")
                            continue
                except Exception as e:
                    self.log(f"[Robot] Error al procesar fila {i}: {e}")
                    continue
            return False
        except Exception as e:
            self.log(f"[Robot] Error al buscar expedientes: {e}")
            raise

    async def _go_back_to_list(self, categoria: str, serial: str, cn: Optional[str] = None):
        """
        Vuelve al listado de justificaciones. Si detecta que está atascado en la pasarela
        de Cl@ve, navega directamente a la URL del listado y re-autentica si es necesario.
        CRÍTICO: Siempre re-aplica la búsqueda avanzada para mantener el filtro de "Pdte. presentar"
        """
        current_url = (self.page.url or "").lower()
        
        # Si estamos atascados en la pasarela de Cl@ve, navegar directamente
        if "pasarela.clave.gob.es" in current_url or "clave.gob.es" in current_url:
            self.log("[Robot] Detectada redirección a Cl@ve, navegando directamente al listado...")
            try:
                # Navegar a la página principal de justificaciones
                await self.navigate_to_justificaciones(categoria)
                
                # Verificar si necesita re-autenticación
                if await self._detect_clave():
                    self.log("[Robot] Re-autenticación requerida después de navegación...")
                    await self.authenticate_if_needed(serial, cn, categoria)
                
                # RE-APLICAR búsqueda avanzada para mantener el filtro
                await self._use_advanced_search()
                return
            except Exception as nav_err:
                self.log(f"[Robot] Error al navegar directamente: {nav_err}")
        
            # Intento normal con go_back
        try:
            await self.page.go_back(timeout=15000)
            await self.page.wait_for_load_state("networkidle", timeout=15000)
            
            # Detectar y corregir error "Not Found" después de go_back
            await self._detect_and_fix_not_found()
            
            # Verificar si seguimos en pasarela después de go_back
            current_url = (self.page.url or "").lower()
            if "pasarela.clave.gob.es" in current_url or "clave.gob.es" in current_url:
                self.log("[Robot] Aún en pasarela después de go_back, navegando...")
                await self.navigate_to_justificaciones(categoria)
                
                # Verificar si necesita re-autenticación
                if await self._detect_clave():
                    self.log("[Robot] Re-autenticación requerida...")
                    await self.authenticate_if_needed(serial, cn, categoria)
                
                # RE-APLICAR búsqueda avanzada
                await self._use_advanced_search()
            else:
                # Verificar si perdimos el filtro de búsqueda avanzada
                # Si la URL no contiene los parámetros de búsqueda, re-aplicar
                if "search" not in current_url and "advanced" not in current_url:
                    self.log("[Robot] Filtro de búsqueda perdido, re-aplicando búsqueda avanzada...")
                    await self._use_advanced_search()
        except Exception:
            # Fallback: navegar y re-aplicar búsqueda
            try:
                self.log("[Robot] Error en go_back, navegando directamente y re-aplicando búsqueda...")
                await self.navigate_to_justificaciones(categoria)
                
                if await self._detect_clave():
                    await self.authenticate_if_needed(serial, cn, categoria)
                
                await self._use_advanced_search()
            except Exception as e:
                self.log(f"[Robot] Error crítico en navegación: {e}")
                raise


    async def _detect_and_fix_not_found(self) -> bool:
        """
        Detecta si aparece el error 'Not Found' en la página y lo corrige refrescando.
        Retorna True si no hay error o se solucionó, False si el error persiste.
        """
        try:
            not_found = self.page.locator("text=Not Found")
            if await not_found.count() > 0:
                self.log("[Robot] ⚠ Error 'Not Found' detectado. Refrescando página...")
                await self.page.reload(wait_until="networkidle", timeout=15000)
                await self._sleep(1.0)
                # Verificar que se solucionó
                if await not_found.count() == 0:
                    self.log("[Robot] ✓ Página refrescada correctamente")
                    return True
                else:
                    self.log("[Robot] ✗ Error 'Not Found' persiste después de refrescar")
                    return False
            return True  # No hay error
        except Exception as e:
            self.log(f"[Robot] Advertencia al verificar 'Not Found': {e}")
            return True  # Continuar si hay error en la verificación

    async def _try_firma_clave(self, serial: str, cn: Optional[str] = None) -> bool:
        try:
            self.log("[Robot] Buscando botón 'Firma con Cl@ve y presentar'...")
            btn_firma = self.page.locator(self.SEL["btn_firma_clave"])
            
            # Esperar a que desaparezca el spinner si existe
            try:
                spinner = self.page.locator("#spinner-div")
                if await spinner.count() > 0:
                    self.log("[Robot] Esperando que desaparezca el spinner...")
                    await spinner.wait_for(state="hidden", timeout=10000)
                    await self._sleep(0.5)
            except Exception:
                pass  # Si no hay spinner o ya desapareció, continuar
            
            if await btn_firma.count() == 0:
                self.log("[Robot] ✗ Botón 'Firma con Cl@ve' no encontrado")
                return False
            
            self.log("[Robot] ✓ Botón 'Firma con Cl@ve' encontrado")
            # TIMEOUT AUMENTADO: 8s → 30s para dar tiempo a la firma digital
            await btn_firma.click(timeout=30000)
            self.log("[Robot] Pulsado 'Firma con Cl@ve y presentar'. Esperando firma digital...")
            # ESPERA AUMENTADA: 1.0s → 3.0s después del clic
            await self._sleep(3.0)
            
            if await self._detect_clave():
                self.log("[Robot] Pasarela Cl@ve detectada durante firma, seleccionando certificado SIN watcher...")
                # CRÍTICO: NO usar watcher durante la firma de expedientes
                # El watcher acepta demasiado rápido sin verificar el certificado correcto
                # Usar solo los métodos que verifican activamente (UIA y cert_clicker con timeout largo)
                await self._select_cert_in_clave(serial, cn)
                await self.page.wait_for_load_state("networkidle", timeout=30000)
            
            # CICLO DE VERIFICACIÓN AUMENTADO: 30 intentos → 60 intentos (30 segundos totales)
            for _ in range(60):
                await asyncio.sleep(0.5)
                # Solo verificar el botón de Cl@ve, sin referencias a AutoFirma
                if await btn_firma.count() == 0:
                    self.log("[Robot] Botón de firma ya no visible (posible éxito).")
                    return True
                try:
                    if await btn_firma.first.is_disabled():
                        self.log("[Robot] Botón de firma deshabilitado (posible éxito).")
                        return True
                except Exception:
                    pass
            self.log("[Robot] No se detectó confirmación clara tras firmar con Cl@ve (timeout).")
            return False
        except Exception as e:
            self.log(f"[Robot] Error al intentar 'Firma con Cl@ve y presentar': {e}")
            return False

    async def _sign_current_record(self, serial: str, cn: Optional[str] = None) -> bool:
        # Solo firma con Cl@ve, sin fallback
        return await self._try_firma_clave(serial, cn)

    async def _use_advanced_search(self):
        """Usa la búsqueda avanzada para filtrar solo expedientes 'Pdte. presentar'"""
        try:
            # Paso 1: Hacer clic en el botón que navega a la página de búsqueda avanzada
            self.log("[Robot] Navegando a la página de búsqueda avanzada...")
            await self.page.click(self.SEL["btn_advanced_search"], timeout=10000)
            await self._sleep(1.0)
            
            # Esperar a que se cargue la página de búsqueda
            await self.page.wait_for_load_state("domcontentloaded", timeout=15000)
            await self._sleep(0.5)
            
            # Paso 2: Seleccionar "Pdte. presentar" (valor "3") en el desplegable de estado
            self.log("[Robot] Seleccionando estado 'Pdte. presentar' (valor: 3)...")
            await self.page.select_option(self.SEL["select_estado"], value="3", timeout=10000)
            self.log("[Robot] Estado seleccionado correctamente.")
            await self._sleep(0.5)
            
            # Paso 3: Hacer clic en el botón buscar con reintentos y verificación
            search_button = self.page.locator("#advancedSearch")
            max_attempts = 2
            
            for attempt in range(1, max_attempts + 1):
                try:
                    self.log(f"[Robot] Intento {attempt}/{max_attempts}: Esperando que el botón de búsqueda esté listo...")
                    
                    # Esperar a que el botón esté visible y habilitado
                    await search_button.wait_for(state="visible", timeout=5000)
                    is_enabled = await search_button.is_enabled()
                    self.log(f"[Robot] Botón de búsqueda - Visible: ✓, Habilitado: {'✓' if is_enabled else '✗'}")
                    
                    if not is_enabled:
                        self.log("[Robot] Botón deshabilitado, esperando...")
                        await self._sleep(0.5)
                    
                    # Intentar el clic
                    self.log(f"[Robot] Haciendo clic en buscar (intento {attempt}/{max_attempts})...")
                    
                    if attempt == 1:
                        # Primer intento: clic normal
                        await search_button.click(timeout=5000)
                    else:
                        # Segundo intento: clic forzado o Enter
                        self.log("[Robot] Usando clic forzado...")
                        try:
                            await search_button.click(force=True, timeout=5000)
                        except Exception:
                            self.log("[Robot] Clic forzado falló, intentando Enter en el formulario...")
                            await self.page.keyboard.press("Enter")
                    
                    self.log("[Robot] Clic ejecutado, esperando resultados...")
                    await self._sleep(1.0)
                    
                    # Verificar que apareció la tabla de resultados
                    self.log("[Robot] Verificando que apareció la tabla de resultados...")
                    table = self.page.locator(self.SEL["table_justificaciones"])
                    
                    try:
                        await table.wait_for(state="visible", timeout=8000)
                        self.log("[Robot] Tabla visible, esperando a que termine de cargar...")
                        
                        # CRÍTICO: Esperar networkidle ANTES de contar filas
                        # Esto asegura que la página termine de cargar y la tabla se renderice con datos
                        try:
                            await self.page.wait_for_load_state("networkidle", timeout=15000)
                            self.log("[Robot] ✓ Página en networkidle después de buscar")
                        except Exception as net_err:
                            self.log(f"[Robot] Advertencia: No se alcanzó networkidle: {net_err}")
                        
                        # Ahora sí, esperar activamente a que la tabla tenga contenido
                        rows = table.locator("tbody tr")
                        max_wait_attempts = 30  # 30 intentos × 0.5s = 15 segundos máximo
                        row_count = 0
                        
                        self.log("[Robot] Verificando contenido de la tabla...")
                        for attempt in range(max_wait_attempts):
                            row_count = await rows.count()
                            if row_count > 0:
                                self.log(f"[Robot] ✓ Tabla cargada con {row_count} filas (intento {attempt + 1}/{max_wait_attempts})")
                                break
                            if attempt < max_wait_attempts - 1:
                                await asyncio.sleep(0.5)
                            else:
                                self.log(f"[Robot] ⚠ Advertencia: Tabla visible pero sin datos después de {max_wait_attempts * 0.5}s")
                        
                        # Si aún no hay filas después de esperar, podría ser un problema
                        if row_count == 0:
                            self.log("[Robot] ⚠ La tabla no tiene filas. Puede que no haya expedientes o hubo un error.")
                        
                        self.log("[Robot] ✓ Búsqueda avanzada completada exitosamente.")
                        
                        # Detectar y corregir error "Not Found" después de la búsqueda
                        await self._detect_and_fix_not_found()
                        
                        return True
                        
                    except Exception as table_err:
                        self.log(f"[Robot] ✗ Tabla no apareció después del clic: {table_err}")
                        if attempt < max_attempts:
                            self.log(f"[Robot] Reintentando... ({attempt + 1}/{max_attempts})")
                            await self._sleep(1.0)
                            continue
                        else:
                            raise RuntimeError("La tabla de resultados no apareció después de hacer clic en buscar")
                
                except Exception as click_err:
                    self.log(f"[Robot] Error en intento {attempt}: {click_err}")
                    if attempt < max_attempts:
                        await self._sleep(1.0)
                        continue
                    else:
                        raise
            
            # Si llegamos aquí, ningún intento funcionó
            self.log("[Robot] ✗ No se pudo completar la búsqueda después de todos los intentos")
            return False
            
        except Exception as e:
            self.log(f"[Robot] ✗ Error crítico al usar búsqueda avanzada: {e}")
            return False

    async def _run_full(self, categoria: str, serial: str, cn: Optional[str] = None):
        try:
            self.stop_flag.clear()
            await self._close()  # garantizar estado limpio
            await self.navigate_to_justificaciones(categoria)
            await self.authenticate_if_needed(serial, cn, categoria)

            # Usar búsqueda avanzada para filtrar solo "Pdte. presentar"
            search_ok = await self._use_advanced_search()
            if not search_ok:
                self.log("[Robot] No se pudo usar la búsqueda avanzada. Abortando...")
                return

            total_firmados = 0
            failed_expedientes = set()  # Expedientes que han fallado para evitar bucles
            error_count = 0  # Contador de errores consecutivos
            max_errors = 3  # Máximo de errores antes de reiniciar
            
            while not self.stop_flag.is_set():
                self.log("[Robot] Buscando expedientes en la página actual...")
                
                # Detectar y corregir error "Not Found" antes de escanear
                await self._detect_and_fix_not_found()
                
                try:
                    opened = await self._scan_rows_and_open_first_pending()
                    error_count = 0  # Resetear contador si funciona
                except Exception as e:
                    error_count += 1
                    self.log(f"[Robot] Error al buscar expedientes (intento {error_count}/{max_errors}): {e}")
                    if error_count >= max_errors:
                        self.log("[Robot] Demasiados errores consecutivos. Navegando a página inicial...")
                        # Navegar a la página inicial sin cerrar Chrome
                        await self.navigate_to_justificaciones(categoria)
                        await self.authenticate_if_needed(serial, cn, categoria)
                        # Volver a aplicar búsqueda avanzada
                        search_ok = await self._use_advanced_search()
                        if not search_ok:
                            self.log("[Robot] No se pudo reestablecer la búsqueda avanzada. Abortando...")
                            break
                        error_count = 0
                        failed_expedientes.clear()
                    continue
                
                if not opened:
                    try:
                        next_btn = self.page.locator(self.SEL["paginate_next"]).first
                        classes = (await next_btn.get_attribute("class")) or ""
                        if "disabled" in classes:
                            self.log("[Robot] No hay más páginas. Proceso completado.")
                            break
                        await next_btn.click(timeout=8000)
                        await self._sleep(0.8)
                        continue
                    except Exception:
                        self.log("[Robot] No hay más expedientes. Proceso completado.")
                        break

                # Obtener código del expediente actual para tracking
                try:
                    current_url = self.page.url or ""
                    expediente_match = re.search(r'(KD|KC)/\d+-\d+', current_url)
                    expediente_id = expediente_match.group(0) if expediente_match else ""
                    
                    # Verificar si ya falló este expediente
                    if expediente_id and expediente_id in failed_expedientes:
                        self.log(f"[Robot] Expediente {expediente_id} ya ha fallado antes. Saltando...")
                        await self._go_back_to_list(categoria, serial, cn)
                        continue
                except Exception:
                    expediente_id = ""

                # CRÍTICO: Esperar a que la página del expediente cargue completamente
                self.log("[Robot] Esperando a que cargue la página del expediente...")
                
                # 1. Esperar networkidle para asegurar que la página termine de cargar
                try:
                    await self.page.wait_for_load_state("networkidle", timeout=20000)
                    self.log("[Robot] ✓ Página en estado 'networkidle'")
                except Exception as e:
                    self.log(f"[Robot] Advertencia: No se alcanzó 'networkidle': {e}")
                
                # 2. Espera activa de hasta 10 segundos buscando que aparezca el botón de firma
                self.log("[Robot] Esperando activamente a que aparezca el botón de firma...")
                max_wait_button = 20  # 20 intentos × 0.5s = 10 segundos
                button_appeared = False
                
                for attempt in range(max_wait_button):
                    try:
                        btn_clave_count = await self.page.locator(self.SEL["btn_firma_clave"]).count()
                        
                        if btn_clave_count > 0:
                            button_appeared = True
                            self.log(f"[Robot] ✓ Botón de firma detectado (intento {attempt + 1}/{max_wait_button})")
                            break
                    except Exception:
                        pass
                    
                    if attempt < max_wait_button - 1:
                        await asyncio.sleep(0.5)
                
                if not button_appeared:
                    self.log("[Robot] ⚠ No apareció el botón de firma después de 10 segundos")
                
                # 3. Si aparece spinner, esperarlo
                try:
                    self.log("[Robot] Verificando si hay spinner de carga...")
                    spinner = self.page.locator("#spinner-div")
                    if await spinner.count() > 0:
                        self.log("[Robot] Esperando a que desaparezca el spinner...")
                        await spinner.wait_for(state="hidden", timeout=15000)
                        self.log("[Robot] ✓ Spinner desaparecido")
                        await self._sleep(1.5)  # Espera adicional aumentada de 0.5s a 1.5s
                    else:
                        await self._sleep(1.0)  # Pequeña espera si no hay spinner
                except Exception as spinner_err:
                    self.log(f"[Robot] Advertencia: Error con spinner: {spinner_err}")
                    await self._sleep(1.0)
                
                # 4. Verificación final: esperar explícitamente a que el botón esté visible y listo
                self.log("[Robot] Verificación final: esperando botón de firma con Cl@ve...")
                button_found = False
                try:
                    # Esperar al botón de Cl@ve - timeout aumentado a 15s
                    await self.page.wait_for_selector(self.SEL["btn_firma_clave"], timeout=15000, state="visible")
                    self.log("[Robot] ✓ Botón 'Firma con Cl@ve' detectado y visible")
                    button_found = True
                except Exception as btn_err:
                    self.log(f"[Robot] ✗ No se detectó el botón de firma: {btn_err}")
                
                if not button_found:
                    self.log("[Robot] ✗ No se encontró botón de firma después de esperar. Saltando expediente...")
                    # CRÍTICO: Marcar como fallido solo si no hay botón (problema estructural)
                    if expediente_id:
                        failed_expedientes.add(expediente_id)
                    await self._go_back_to_list(categoria, serial, cn)
                    continue
                
                self.log("[Robot] Intentando firmar expediente con Cl@ve...")
                ok = await self._sign_current_record(serial, cn)
                if ok:
                    total_firmados += 1
                    self.log("[Robot] ✓ Expediente firmado exitosamente. Volviendo al listado...")
                else:
                    # CRÍTICO: NO marcar como fallido si falla por timeout
                    # El expediente puede reintentarse en la siguiente pasada
                    self.log(f"[Robot] ✗ No fue posible completar la firma (posible timeout). Continuando con siguiente expediente...")
                
                try:
                    await self._go_back_to_list(categoria, serial, cn)
                    await self.page.wait_for_load_state("networkidle", timeout=20000)
                    await self._sleep(0.8)
                except Exception as nav_err:
                    self.log(f"[Robot] Error al volver al listado: {nav_err}")
                    error_count += 1
                    if error_count >= max_errors:
                        self.log("[Robot] Demasiados errores de navegación. Navegando a página inicial...")
                        # Navegar a la página inicial sin cerrar Chrome
                        await self.navigate_to_justificaciones(categoria)
                        await self.authenticate_if_needed(serial, cn, categoria)
                        # Volver a aplicar búsqueda avanzada
                        search_ok = await self._use_advanced_search()
                        if not search_ok:
                            self.log("[Robot] No se pudo reestablecer la búsqueda avanzada. Abortando...")
                            break
                        error_count = 0
                        failed_expedientes.clear()
                    else:
                        # Último intento: navegar directamente
                        try:
                            self.log("[Robot] Intento de recuperación: navegando directamente al listado...")
                            if categoria == "KD":
                                list_url = "https://portal.gestion.sedepkd.red.gob.es/justificaciones/tic/business-intelligence-2"
                            else:
                                list_url = "https://portal.gestion.sedepkd.red.gob.es/justificaciones/tic/kit-consulting-2"
                            await self.page.goto(list_url, wait_until="domcontentloaded", timeout=20000)
                            await self._sleep(1.0)
                            
                            # Verificar si necesita re-autenticación
                            if await self._detect_clave():
                                self.log("[Robot] Re-autenticación requerida...")
                                await self.authenticate_if_needed(serial, cn, categoria)
                            error_count = 0  # Resetear si funciona la recuperación
                        except Exception:
                            self.log("[Robot] No se pudo recuperar la navegación.")

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
