#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
cert_clicker.py
Helper local para cerrar el diálogo nativo de selección de certificado (Chrome/Windows)
- Selecciona la fila por prefijo de número de serie y/o por CN (SUBJECT CN)
- Pulsa el botón "Aceptar" (o ENTER) y sale con:
    0 -> éxito
    1 -> no se encontró diálogo / fila no coincide
    2 -> timeout u otro error

Uso:
  python cert_clicker.py --serial 2DBB3D0D060A312066E40192B2676785 --cn "YUSTA PLIEGO PABLO - 11086279A" --timeout 25
"""

import sys
import re
import time
import argparse
from typing import Optional
from pywinauto import mouse

try:
    from pywinauto import Desktop, timings
except Exception as e:
    print(f"[cert_clicker] pywinauto no disponible: {e}", file=sys.stderr)
    sys.exit(2)


def norm(s: Optional[str]) -> str:
    return re.sub(r"[^A-Z0-9]", "", (s or "").upper())


def find_dialog(timeout: float = 20.0):
    """Localiza el diálogo de 'Seleccionar un certificado' como descendiente de una ventana de Chrome."""
    t0 = time.time()
    dlg = None
    timings.Timings.after_clickinput_wait = 0.3

    while time.time() - t0 < timeout:
        try:
            wins = Desktop(backend="uia").windows()
            for w in wins:
                try:
                    title = (w.window_text() or "")
                    cls = (getattr(w, "element_info", None).class_name or "")
                    if re.search(r"(Cl@ve|clave|pasarela)", title, re.I) or cls == "Chrome_WidgetWin_1":
                        # Buscar descendiente que sea el propio diálogo
                        try:
                            cand = w.child_window(title_re="Seleccionar.*certificado|Select.*certificate", control_type="Window")
                            dlg = cand.wrapper_object()
                            return dlg
                        except Exception:
                            # Si no expone un control Window, devolvemos la propia w
                            dlg = w
                            # Validar que dentro hay textos relacionados
                            texts = [c.window_text() for c in dlg.descendants(control_type="Text")]
                            if any("certificado" in (t or "").lower() for t in texts):
                                return dlg
                except Exception:
                    continue
        except Exception:
            pass
        time.sleep(0.3)
    return None


def list_rows(dlg):
    """Recoge posibles filas del grid (Table/DataItem)."""
    rows = []
    try:
        table = dlg.child_window(control_type="Table").wrapper_object()
    except Exception:
        table = None
    if table:
        try:
            rows = table.children()
        except Exception:
            rows = []
    if not rows:
        try:
            rows = dlg.descendants(control_type="DataItem")
        except Exception:
            rows = []
    return rows


def row_text(r):
    """Concatena textos de una fila."""
    parts = []
    try:
        for el in r.descendants(control_type="Text"):
            try:
                parts.append(el.window_text())
            except Exception:
                continue
    except Exception:
        pass
    return " ".join([p for p in parts if p])


def click_accept(dlg):
    """Pulsa Aceptar o Enter en el diálogo con varios respaldos."""
    # Intento 1: botón como control Button
    try:
        btn = dlg.child_window(title_re="Aceptar|Accept", control_type="Button").wrapper_object()
        dlg.set_focus()
        btn.click_input()
        return True
    except Exception:
        pass

    # Intento 2: cualquier control cuyo name sea Aceptar/Accept
    try:
        for el in dlg.descendants():
            try:
                name = (el.window_text() or "")
                if re.search(r"^(Aceptar|Accept)$", name, re.I):
                    w = el.wrapper_object()
                    dlg.set_focus()
                    w.click_input()
                    return True
            except Exception:
                continue
    except Exception:
        pass

    # Intento 3: tecla Enter con foco en el diálogo
    try:
        dlg.set_focus()
        dlg.type_keys("{ENTER}")
        return True
    except Exception:
        pass

    # Intento 4: clic heurístico en la zona inferior central-derecha del diálogo
    try:
        r = dlg.rectangle()
        # Heurística: botón de aceptar suele estar ~ a 1/3 desde la derecha y ~40 px por encima del borde inferior
        x = r.right - int((r.right - r.left) * 0.35)
        y = r.bottom - 50
        mouse.click(button="left", coords=(x, y))
        return True
    except Exception:
        return False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--serial", required=True, help="Número de serie completo (se usa prefijo para coincidir)")
    ap.add_argument("--cn", default="", help="CN (SUBJECT) a coincidir como respaldo")
    ap.add_argument("--timeout", type=float, default=20.0, help="Timeout total en segundos")
    args = ap.parse_args()

    target = norm(args.serial)
    subj = norm(args.cn)

    dlg = find_dialog(timeout=args.timeout)
    # Traer al frente
    try:
        if dlg:
            dlg.set_focus()
    except Exception:
        pass
    if not dlg:
        print("[cert_clicker] No se encontró el diálogo de certificado.")
        sys.exit(1)

    rows = list_rows(dlg)
    if not rows:
        print("[cert_clicker] No se encontraron filas en el diálogo.")
        sys.exit(1)

    # CRÍTICO: Buscar activamente el certificado correcto por serial (más específico) o CN
    picked = False
    matched_row = None
    
    print(f"[cert_clicker] Buscando certificado: Serial={args.serial[:16]}... CN={args.cn or 'N/A'}")
    print(f"[cert_clicker] Total de filas encontradas: {len(rows)}")
    
    # Prioridad 1: Coincidir por número de serie (más específico)
    for i, r in enumerate(rows):
        txt = row_text(r)
        n = norm(txt)
        if target and (target in n or n in target or target[:16] in n):
            matched_row = r
            print(f"[cert_clicker] Fila {i} coincide por serial: {txt.strip()[:100]}")
            break
    
    # Prioridad 2: Coincidir por CN si no hay match por serial
    if not matched_row and subj:
        for i, r in enumerate(rows):
            txt = row_text(r)
            n = norm(txt)
            if subj and subj in n:
                matched_row = r
                print(f"[cert_clicker] Fila {i} coincide por CN: {txt.strip()[:100]}")
                break

    if not matched_row:
        print(f"[cert_clicker] NO se encontró el certificado correcto entre {len(rows)} filas")
        # Imprimir todas las filas disponibles para diagnóstico
        for i, r in enumerate(rows):
            txt = row_text(r)
            print(f"[cert_clicker] Fila {i}: {txt.strip()[:100]}")
        sys.exit(1)

    # Seleccionar activamente la fila correcta
    try:
        matched_row.click_input()
        picked = True
        time.sleep(0.3)  # Pequeña espera para asegurar selección
        print(f"[cert_clicker] Certificado correcto seleccionado")
    except Exception as e:
        print(f"[cert_clicker] Error al hacer clic en la fila: {e}")
        sys.exit(2)

    if not picked:
        print("[cert_clicker] No se pudo seleccionar el certificado")
        sys.exit(1)

    ok = click_accept(dlg)
    if not ok:
        print("[cert_clicker] No se pudo pulsar Aceptar.")
        sys.exit(2)

    print("[cert_clicker] Aceptado con éxito.")
    sys.exit(0)


if __name__ == "__main__":
    try:
        main()
    except SystemExit as se:
        raise
    except Exception as e:
        print(f"[cert_clicker] Error: {e}", file=sys.stderr)
        sys.exit(2)
