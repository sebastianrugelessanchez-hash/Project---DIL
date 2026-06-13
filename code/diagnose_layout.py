"""
Diagnóstico: flujo correcto de VL06F
  1. Navegar a VL06F
  2. Limpiar fechas de 'Pland Gds Mvmnt Date'
  3. Entrar tickets en campo Delivery (IT_VBELN)
  4. F8
  5. Ctrl+F9 -> mostrar estructura del diálogo Choose Layout
  6. Seleccionar /02C y leer grid

Corre con UNA sola sesión SAP abierta.
    python diagnose_layout.py
"""
import win32com.client
import time

TICKETS = ["336330550", "336330552", "336330553"]

GRID_PATHS = [
    "wnd[0]/usr/cntlGRID1/shellcont/shell",
    "wnd[0]/usr/cntlGRID/shellcont/shell",
    "wnd[0]/usr/cntlALV_GRID/shellcont/shell",
    "wnd[0]/usr/shellcont/shell",
]


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


def show_tree(session, target, max_depth=3, max_children=100):
    def _show(el, depth=0):
        if depth > max_depth:
            return
        try:
            print("  " * depth + f"{el.Id}  [{el.Type}]")
            for i in range(min(el.Children.Count, max_children)):
                _show(el.Children(i), depth + 1)
        except Exception:
            pass
    try:
        _show(session.findById(target))
    except Exception as e:
        print(f"  Error: {e}")


def get_sbar(session):
    try:
        return session.findById("wnd[0]/sbar").Text
    except Exception:
        return "(sbar no disponible)"


def get_session():
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    app = sap_gui_auto.GetScriptingEngine
    for i in range(app.Children.Count):
        conn = app.Children(i)
        for j in range(conn.Children.Count):
            candidate = conn.Children(j)
            try:
                candidate.findById("wnd[0]/usr/txtRSYST-BNAME")
                continue
            except Exception:
                return candidate
    return None


def main():
    session = get_session()
    if not session:
        print("No se encontró sesión SAP.")
        return

    try:
        print(f"Sesión: '{session.findById('wnd[0]').Text}'")
    except Exception:
        print("Sesión encontrada.")

    # Cerrar diálogos abiertos
    for w in ("wnd[2]", "wnd[1]"):
        if wnd_exists(session, w):
            try:
                session.findById(w).sendVKey(12)
                wait_ready(session)
            except Exception:
                pass

    # ---- 1. Navegar a VL06F (con /n para reset) ----
    print("\n[1] Navegando a VL06F con /nVL06F...")
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nVL06F"
    session.findById("wnd[0]").sendVKey(0)
    wait_ready(session)
    try:
        print(f"    Pantalla: '{session.findById('wnd[0]').Text}'")
    except Exception:
        pass
    print(f"    Status bar: '{get_sbar(session)}'")

    # ---- 2. Limpiar fechas (mostrar antes y después) ----
    print("\n[2] Limpiando fechas de 'Pland Gds Mvmnt Date'...")
    for field in ("ctxtIT_WADAT-LOW", "ctxtIT_WADAT-HIGH"):
        try:
            before = session.findById(f"wnd[0]/usr/{field}").text
            session.findById(f"wnd[0]/usr/{field}").text = ""
            after = session.findById(f"wnd[0]/usr/{field}").text
            print(f"    {field}: '{before}' -> '{after}'")
        except Exception as e:
            print(f"    {field} NO encontrado: {e}")

    # ---- 3. Campo Delivery (IT_VBELN) ----
    print("\n[3] Abriendo popup IT_VBELN (Delivery)...")
    btn_path = "wnd[0]/usr/btn%_IT_VBELN_%_APP_%-VALU_PUSH"
    if not wnd_exists(session, btn_path):
        print("    ADVERTENCIA: botón IT_VBELN no visible — árbol completo:")
        show_tree(session, "wnd[0]/usr", max_depth=2, max_children=100)
        return
    else:
        print("    Botón IT_VBELN encontrado OK")

    session.findById(btn_path).press()
    wait_ready(session)

    popup = None
    for n in (2, 1):
        try:
            session.findById(f"wnd[{n}]/usr/tabsTAB_STRIP")
            popup = f"wnd[{n}]"
            break
        except Exception:
            continue

    if not popup:
        print("    ERROR: popup no encontrado — árbol wnd[1]:")
        show_tree(session, "wnd[1]", max_depth=3)
        return

    print(f"    Popup en {popup}")
    tbl = (f"{popup}/usr/tabsTAB_STRIP/tabpSIVA"
           "/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE")

    try:
        session.findById(f"{popup}/tbar[0]/btn[24]").press()
    except Exception:
        pass

    for i, t in enumerate(TICKETS):
        row = i % 8
        if i > 0 and row == 0:
            session.findById(tbl).verticalScrollbar.Position = i
        try:
            session.findById(f"{tbl}/ctxtRSCSEL_255-SLOW_I[1,{row}]").Text = t
        except Exception:
            pass

    session.findById(f"{popup}/tbar[0]/btn[8]").press()
    wait_ready(session)
    print(f"    Tickets ingresados: {TICKETS}")

    # ---- 4. F8 ----
    print("\n[4] Ejecutando F8...")
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    wait_ready(session)

    try:
        print(f"    Pantalla: '{session.findById('wnd[0]').Text}'")
    except Exception:
        pass
    print(f"    Status bar: '{get_sbar(session)}'")

    # Verificar popups después de F8
    for w in ("wnd[1]", "wnd[2]"):
        if wnd_exists(session, w):
            try:
                print(f"    Popup {w}: '{session.findById(w).Text}'")
            except Exception:
                print(f"    Popup {w} abierto (sin título)")
            try:
                session.findById(w).sendVKey(12)
                wait_ready(session)
            except Exception:
                pass

    # Buscar grid en todos los paths conocidos
    grid_found = False
    grid = None
    for path in GRID_PATHS:
        try:
            g = session.findById(path)
            print(f"    Grid en: {path}  RowCount={g.RowCount}")
            grid_found = True
            grid = g
            break
        except Exception:
            pass

    if not grid_found:
        print("    Grid NO encontrado — árbol COMPLETO wnd[0]/usr:")
        show_tree(session, "wnd[0]/usr", max_depth=3, max_children=100)
        return

    # ---- 5. Ctrl+F9 -> Choose Layout ----
    print("\n[5] Presionando Ctrl+F9 (VKey 33) para Choose Layout...")
    session.findById("wnd[0]").sendVKey(33)
    wait_ready(session)

    if wnd_exists(session, "wnd[1]"):
        try:
            title = session.findById("wnd[1]").Text
            print(f"    Diálogo abierto: '{title}'")
        except Exception:
            print("    Diálogo wnd[1] abierto")

        SHELL_PATH = (
            "wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501"
            "/cntlG51_CONTAINER/shellcont/shell"
        )
        try:
            shell = session.findById(SHELL_PATH)
            print(f"    Shell encontrado. RowCount={shell.RowCount}")
            for col in ("VARIANT", "LAYOUT", "DISVARIANT", "TEXT"):
                vals = []
                for i in range(min(shell.RowCount, 10)):
                    try:
                        v = shell.GetCellValue(i, col)
                        if v:
                            vals.append(f"[{i}]={v!r}")
                    except Exception:
                        break
                if vals:
                    print(f"    Col {col!r}: {', '.join(vals)}")

            print("\n    Buscando /02C...")
            found = False
            for i in range(shell.RowCount):
                for col in ("VARIANT", "LAYOUT", "DISVARIANT"):
                    try:
                        val = (shell.GetCellValue(i, col) or "").strip()
                        if val == "/02C":
                            print(f"    /02C en fila {i}, col {col!r}")
                            shell.setCurrentCell(i, col)
                            shell.selectedRows = str(i)
                            shell.doubleClickCurrentCell()
                            wait_ready(session)
                            found = True
                            break
                    except Exception:
                        continue
                if found:
                    break
            if not found:
                print("    /02C NO encontrado — cerrando diálogo")
                session.findById("wnd[1]").sendVKey(12)
                wait_ready(session)
        except Exception as e:
            print(f"    Shell NO encontrado: {e}")
            print("    Árbol del diálogo:")
            show_tree(session, "wnd[1]", max_depth=3)
            session.findById("wnd[1]").sendVKey(12)
            wait_ready(session)
    else:
        print("    wnd[1] no existe después de Ctrl+F9")

    # ---- 6. Leer grid ----
    print("\n[6] Leyendo grid (con o sin layout /02C)...")
    for path in GRID_PATHS:
        try:
            grid = session.findById(path)
            print(f"    Grid: {path}  RowCount={grid.RowCount}")
            for col in ["VBELN", "ZZVBELN", "ZZFKNUM", "ZZTKNUM", "WBSTK", "ZZVBELN_IL"]:
                try:
                    val = grid.GetCellValue(0, col)
                    print(f"    Col={col!r}: {val!r}")
                except Exception:
                    print(f"    Col={col!r}: NO DISPONIBLE")
            break
        except Exception:
            pass

    print("\nDiagnóstico completado.")


if __name__ == "__main__":
    main()
