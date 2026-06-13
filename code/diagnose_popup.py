"""
Script de diagnóstico v3: prueba la nueva lógica de popup dinámico para VL06F.
    python diagnose_popup.py
"""
import win32com.client
import time


def wait_ready(session, timeout=10.0):
    end = time.time() + timeout
    while time.time() < end:
        try:
            if not session.Busy:
                return
        except Exception:
            pass
        time.sleep(0.3)


def wnd_exists(session, wnd_id):
    try:
        session.findById(wnd_id)
        return True
    except Exception:
        return False


def wnd_title(session, wnd_id):
    try:
        return session.findById(wnd_id).Text
    except Exception:
        return "(no disponible)"


def show_tree(session, target, max_depth=3):
    def _show(el, depth=0):
        if depth > max_depth:
            return
        indent = "  " * depth
        try:
            print(f"{indent}{el.Id}  [{el.Type}]")
            for i in range(min(el.Children.Count, 20)):
                _show(el.Children(i), depth + 1)
        except Exception:
            pass
    try:
        _show(session.findById(target))
    except Exception as e:
        print(f"  Error mostrando árbol de {target}: {e}")


def find_popup_wnd(session):
    for n in (2, 1):
        try:
            session.findById(f"wnd[{n}]/usr/tabsTAB_STRIP")
            return f"wnd[{n}]"
        except Exception:
            continue
    return None


def main():
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    app = sap_gui_auto.GetScriptingEngine

    session = None
    for i in range(app.Children.Count):
        conn = app.Children(i)
        if conn.Children.Count > 0:
            candidate = conn.Children(0)
            try:
                candidate.findById("wnd[0]/usr/txtRSYST-BNAME")
                continue
            except Exception:
                session = candidate
                break

    if not session:
        print("No se encontró sesión SAP activa.")
        return

    print("Sesión encontrada.")

    # ---- 1. Navegar a VL06F ----
    print("\n[1] Navegando a VL06F...")
    session.findById("wnd[0]/tbar[0]/okcd").Text = "VL06F"
    session.findById("wnd[0]").sendVKey(0)
    wait_ready(session)
    print(f"    Título wnd[0]: {wnd_title(session, 'wnd[0]')}")

    # ---- 2. Limpiar fechas ----
    print("\n[2] Limpiando fechas...")
    for field in ("ctxtIT_WADAT-LOW", "ctxtIT_WADAT-HIGH"):
        try:
            session.findById(f"wnd[0]/usr/{field}").text = ""
            print(f"    {field} -> limpiado")
        except Exception:
            print(f"    {field} -> no encontrado (OK si no existe)")

    # ---- 3. Presionar btn[14] ----
    print("\n[3] Presionando btn[14] (Dynamic Selections)...")
    try:
        session.findById("wnd[0]/tbar[1]/btn[14]").press()
        wait_ready(session)
        print("    btn[14] presionado.")
    except Exception as e:
        print(f"    btn[14] FALLÓ: {e}")

    wnd1_after_btn14 = wnd_exists(session, "wnd[1]")
    print(f"    wnd[1] existe después de btn[14]: {wnd1_after_btn14}")
    if wnd1_after_btn14:
        print(f"    wnd[1] título: {wnd_title(session, 'wnd[1]')}")
        print(f"\n    Árbol de wnd[1] (para identificar si es diálogo de selecciones):")
        show_tree(session, "wnd[1]", max_depth=2)

    # ---- 4. Si btn[14] abrió wnd[1], cerrarlo ----
    if wnd1_after_btn14:
        print("\n[3b] Cerrando wnd[1] con btn[0] (Enter/OK)...")
        try:
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            wait_ready(session)
            still_open = wnd_exists(session, "wnd[1]")
            print(f"    wnd[1] después de btn[0]: {'SIGUE ABIERTO' if still_open else 'cerrado OK'}")
        except Exception as e:
            print(f"    btn[0] FALLÓ: {e}")
            # Intentar con Escape
            try:
                session.findById("wnd[0]").sendVKey(12)  # F12 = Escape
                wait_ready(session)
                print("    Intentado F12 para cerrar.")
            except Exception:
                pass

    # ---- 5. Verificar botón IT_VBELN ----
    print("\n[4] Verificando botón IT_VBELN...")
    btn_wnd0 = "wnd[0]/usr/btn%_IT_VBELN_%_APP_%-VALU_PUSH"
    btn_wnd1 = "wnd[1]/usr/btn%_IT_VBELN_%_APP_%-VALU_PUSH"
    print(f"    btn en wnd[0]: {wnd_exists(session, btn_wnd0)}")
    print(f"    btn en wnd[1]: {wnd_exists(session, btn_wnd1)}")

    # ---- 6. Presionar el botón (donde exista) ----
    print("\n[5] Presionando botón IT_VBELN...")
    btn_pressed = False
    for btn_path in (btn_wnd0, btn_wnd1):
        if wnd_exists(session, btn_path):
            try:
                session.findById(btn_path).press()
                wait_ready(session)
                print(f"    Presionado: {btn_path}")
                btn_pressed = True
                break
            except Exception as e:
                print(f"    Error al presionar {btn_path}: {e}")

    if not btn_pressed:
        print("    Botón IT_VBELN no encontrado en ninguna ventana.")
        return

    # ---- 7. Estado después del botón ----
    print("\n[6] Ventanas después de presionar IT_VBELN:")
    for w in ("wnd[0]", "wnd[1]", "wnd[2]"):
        exists = wnd_exists(session, w)
        print(f"    {w}: {'existe' if exists else 'no existe'}", end="")
        if exists:
            print(f"  título='{wnd_title(session, w)}'")
        else:
            print()

    # ---- 8. Detectar popup ----
    popup = find_popup_wnd(session)
    if popup:
        print(f"\n[7] Popup detectado en: {popup}")
        print(f"    Árbol completo de {popup}:")
        show_tree(session, popup, max_depth=4)

        # Verificar campos de la tabla
        table = (
            f"{popup}/usr/tabsTAB_STRIP/tabpSIVA"
            "/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE"
        )
        print(f"\n[8] Verificando tabla en: {table}")
        for prefix in ("ctxtRSCSEL_255-SLOW_I[1,0]", "txtRSCSEL_255-SLOW_I[1,0]"):
            full = f"{table}/{prefix}"
            found = wnd_exists(session, full)
            print(f"    {prefix}: {'ENCONTRADO OK' if found else 'no encontrado'}")

        print(f"\n[9] Cerrando popup con Escape (F12)...")
        try:
            session.findById("wnd[0]").sendVKey(12)
            wait_ready(session)
            print("    Popup cerrado.")
        except Exception as e:
            print(f"    Error cerrando popup: {e}")
    else:
        print("\n[7] No se encontró popup (wnd[1] ni wnd[2] tienen tabsTAB_STRIP).")

    print("\nDiagnóstico completado.")


if __name__ == "__main__":
    main()
