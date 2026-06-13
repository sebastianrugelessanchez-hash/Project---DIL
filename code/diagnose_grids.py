"""
Diagnóstico de LAYOUT / GRIDS de SAP.

Objetivo: descubrir los NOMBRES TÉCNICOS reales de las columnas de cada grid
(ZCMR, ZSD_DEL_TICKETS, VL06F) y las variantes/layouts disponibles, para poder
definir los pasos del pipeline con certeza en vez de adivinar nombres de columna.

Motivación (ver logs de fallos Batch 6-8):
  - Batch 6 (ZCMR) leía 0 órdenes  -> nombres de columna del sub-grid no coinciden.
  - Batch 7 (ZSD) no mapeaba tickets -> la columna del ticket se llama 'TICKET_CODE',
    no 'TICKET' como asumía el código.
  - VL06F fallaba al seleccionar layout ('virtual key is not enabled').

USO (requiere una sesión SAP YA abierta y logueada):
    python diagnose_grids.py zcmr 96358938 96358939
    python diagnose_grids.py zsd 96358938
    python diagnose_grids.py vl06f 96358938
    python diagnose_grids.py all              # usa tickets de muestra del Excel
    python diagnose_grids.py layouts zcmr     # solo lista layouts disponibles

Si no se pasan tickets, intenta leer los primeros del Excel vía read_zcmr().
Toda la salida se imprime en consola Y se guarda en Data-bases/Logs/.
"""
import sys
import time

import win32com.client

from sap_utils import _navigate_to, _wait_ready, _enter_multi_values, _POPUP_TABLE
from log_util import setup_logging


# Paths conocidos donde SAP suele exponer el grid ALV principal
_GRID_PATHS = [
    "wnd[0]/usr/cntlGRID1/shellcont/shell",
    "wnd[0]/usr/cntlGRID/shellcont/shell",
    "wnd[0]/usr/cntlALV_GRID/shellcont/shell",
    "wnd[0]/usr/shellcont/shell",
]

# Sub-grid de ZCMR (al hacer doble clic en una fila del grid principal)
_ZCMR_SUB_GRID = "wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell"


# --------------------------------------------------------------------------- #
# Conexión a la sesión SAP activa                                              #
# --------------------------------------------------------------------------- #
def get_session():
    """Retorna la primera sesión SAP que NO esté en la pantalla de login."""
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    app = sap_gui_auto.GetScriptingEngine
    for i in range(app.Children.Count):
        conn = app.Children(i)
        for j in range(conn.Children.Count):
            candidate = conn.Children(j)
            try:
                candidate.findById("wnd[0]/usr/txtRSYST-BNAME")
                continue  # está en login -> saltar
            except Exception:
                return candidate
    return None


def _sbar(session):
    try:
        bar = session.findById("wnd[0]/sbar")
        return f"[{bar.MessageType}] {bar.Text}"
    except Exception:
        return "(sbar no disponible)"


def _close_open_popups(session):
    for w in ("wnd[2]", "wnd[1]"):
        try:
            session.findById(w)
            session.findById(w).sendVKey(12)  # F12 = cancelar
            _wait_ready(session)
        except Exception:
            pass


# --------------------------------------------------------------------------- #
# Núcleo: volcar columnas + samples de cualquier grid ALV                      #
# --------------------------------------------------------------------------- #
def dump_grid(session, grid_id: str, label: str, max_rows: int = 3):
    """
    Imprime RowCount, tipo y TODAS las columnas (nombre técnico + valores de
    muestra de las primeras filas) de un grid ALV. Retorna el grid o None.
    """
    print(f"\n--- GRID [{label}] @ {grid_id} ---")
    try:
        grid = session.findById(grid_id)
    except Exception as e:
        print(f"    NO ENCONTRADO: {e}")
        return None

    try:
        rc = int(grid.RowCount)
    except Exception:
        rc = 0
    try:
        gtype = grid.Type
    except Exception:
        gtype = "?"
    print(f"    Type={gtype}  RowCount={rc}")

    try:
        cols = list(grid.ColumnOrder)
    except Exception as e:
        print(f"    No se pudo leer ColumnOrder: {e}")
        return grid

    nrows = min(rc, max_rows)
    print(f"    {len(cols)} columnas (nombre técnico -> samples de las primeras {nrows} filas):")
    for c in cols:
        samples = []
        for r in range(nrows):
            try:
                samples.append(repr(grid.GetCellValue(r, c)))
            except Exception:
                samples.append("ERR")
        print(f"      {c!r:38s} -> {samples}")
    return grid


def dump_layouts(session):
    """
    Abre el diálogo 'Choose Layout' (Ctrl+F9 = VKey 33) y lista las variantes
    disponibles con su nombre técnico. Cierra el diálogo al terminar.
    """
    print("\n--- LAYOUTS DISPONIBLES (Ctrl+F9) ---")
    try:
        session.findById("wnd[0]").sendVKey(33)
        _wait_ready(session)
    except Exception as e:
        print(f"    Ctrl+F9 falló: {e}")
        return

    if not _wnd_exists(session, "wnd[1]"):
        print("    No se abrió diálogo de layout (wnd[1] inexistente).")
        return

    shell_path = (
        "wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501"
        "/cntlG51_CONTAINER/shellcont/shell"
    )
    try:
        shell = session.findById(shell_path)
        rc = int(shell.RowCount)
        print(f"    {rc} layouts encontrados:")
        cols = list(shell.ColumnOrder)
        for i in range(min(rc, 40)):
            vals = {}
            for col in cols:
                try:
                    v = shell.GetCellValue(i, col)
                    if v:
                        vals[col] = v
                except Exception:
                    continue
            print(f"      [{i}] {vals}")
    except Exception as e:
        print(f"    No se pudo leer el shell de layouts: {e}")
    finally:
        try:
            session.findById("wnd[1]").sendVKey(12)
            _wait_ready(session)
        except Exception:
            pass


def _wnd_exists(session, wnd_id):
    try:
        session.findById(wnd_id)
        return True
    except Exception:
        return False


def _find_main_grid(session):
    for path in _GRID_PATHS:
        try:
            session.findById(path)
            return path
        except Exception:
            continue
    return None


def _enter_tickets_via_popup(session, tickets, plant_star=False):
    """
    Patrón común ZCMR/ZSD: abre el popup multi-valor de P_TICKET, limpia
    valores previos, sube los tickets y ejecuta (btn[8]).
    """
    if plant_star:
        try:
            session.findById("wnd[0]/usr/ctxtP_PLANT-LOW").Text = "*"
            session.findById("wnd[0]/usr/ctxtP_PLANT-LOW").SetFocus()
        except Exception as e:
            print(f"    (P_PLANT-LOW no seteado: {e})")

    session.findById("wnd[0]/usr/btn%_P_TICKET_%_APP_%-VALU_PUSH").press()
    _wait_ready(session)
    try:
        session.findById("wnd[1]/tbar[0]/btn[16]").press()  # limpiar selección
        _wait_ready(session)
    except Exception:
        pass
    _enter_multi_values(session, _POPUP_TABLE, tickets)
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    _wait_ready(session)


# --------------------------------------------------------------------------- #
# Diagnósticos por transacción                                                 #
# --------------------------------------------------------------------------- #
def diagnose_zcmr(session, tickets):
    print("\n========== ZCMR ==========")
    _navigate_to(session, "ZCMR")
    _wait_ready(session)
    print(f"  Pantalla: {session.findById('wnd[0]').Text!r}")

    _enter_tickets_via_popup(session, tickets, plant_star=True)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()  # F8
    _wait_ready(session)
    print(f"  Status bar tras F8: {_sbar(session)}")

    grid_path = _find_main_grid(session)
    if not grid_path:
        print("  Grid principal NO encontrado en ningún path conocido.")
        return
    grid = dump_grid(session, grid_path, "ZCMR-MAIN")
    dump_layouts(session)

    # Drill-down al sub-grid: doble clic en la primera fila con datos
    if grid is None:
        return
    print("\n  >> Drill-down al SUB-GRID (doble clic en fila 0)...")
    try:
        grid.currentCellRow = 0
        grid.selectedRows = "0"
        grid.doubleClickCurrentCell()
        _wait_ready(session)
        dump_grid(session, _ZCMR_SUB_GRID, "ZCMR-SUBGRID")
    except Exception as e:
        print(f"     Drill-down falló: {e}")


def diagnose_zsd(session, tickets):
    print("\n========== ZSD_DEL_TICKETS ==========")
    _navigate_to(session, "ZSD_DEL_TICKETS")
    _wait_ready(session)
    print(f"  Pantalla: {session.findById('wnd[0]').Text!r}")

    _enter_tickets_via_popup(session, tickets, plant_star=False)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()  # F8
    _wait_ready(session)
    print(f"  Status bar tras F8: {_sbar(session)}")

    grid_path = _find_main_grid(session)
    if not grid_path:
        print("  Grid NO encontrado.")
        return
    dump_grid(session, grid_path, "ZSD")
    dump_layouts(session)


def diagnose_vl06f(session, tickets):
    print("\n========== VL06F ==========")
    _navigate_to(session, "VL06F")
    _wait_ready(session)
    print(f"  Pantalla: {session.findById('wnd[0]').Text!r}")

    # Limpiar fechas de movimiento de mercancía
    for field in ("ctxtIT_WADAT-LOW", "ctxtIT_WADAT-HIGH"):
        try:
            session.findById(f"wnd[0]/usr/{field}").Text = ""
        except Exception:
            pass

    # Popup de Delivery (IT_VBELN)
    try:
        session.findById("wnd[0]/usr/btn%_IT_VBELN_%_APP_%-VALU_PUSH").press()
        _wait_ready(session)
        try:
            session.findById("wnd[1]/tbar[0]/btn[16]").press()
            _wait_ready(session)
        except Exception:
            pass
        _enter_multi_values(session, _POPUP_TABLE, tickets)
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        _wait_ready(session)
    except Exception as e:
        print(f"  Error ingresando deliveries: {e}")

    session.findById("wnd[0]/tbar[1]/btn[8]").press()  # F8
    _wait_ready(session)
    print(f"  Status bar tras F8: {_sbar(session)}")

    grid_path = _find_main_grid(session)
    if not grid_path:
        print("  Grid NO encontrado.")
        return
    dump_grid(session, grid_path, "VL06F")
    dump_layouts(session)


# --------------------------------------------------------------------------- #
# Entry point                                                                  #
# --------------------------------------------------------------------------- #
_DISPATCH = {
    "zcmr": diagnose_zcmr,
    "zsd": diagnose_zsd,
    "vl06f": diagnose_vl06f,
}


def _resolve_tickets(args_tickets):
    if args_tickets:
        return args_tickets
    print("  (sin tickets en CLI — leyendo los primeros 3 del Excel vía read_zcmr)")
    try:
        from excel_reader import read_zcmr
        tickets = read_zcmr()[:3]
        if tickets:
            print(f"  Tickets de muestra: {tickets}")
            return tickets
    except Exception as e:
        print(f"  No se pudo leer del Excel: {e}")
    raise SystemExit("Pasa al menos un ticket: python diagnose_grids.py zcmr <ticket> ...")


def main():
    argv = sys.argv[1:]
    if not argv:
        print(__doc__)
        return

    target = argv[0].lower()

    # Modo 'layouts <tcode>': solo listar layouts sin volcar grid completo
    if target == "layouts":
        tcode = argv[1].lower() if len(argv) > 1 else "zcmr"
        tickets = _resolve_tickets(argv[2:])
        setup_logging(prefix=f"diagnose_layouts_{tcode}")
        session = get_session()
        if not session:
            print("No se encontró sesión SAP abierta.")
            return
        _close_open_popups(session)
        _DISPATCH.get(tcode, diagnose_zcmr)(session, tickets)
        return

    setup_logging(prefix=f"diagnose_{target}")
    session = get_session()
    if not session:
        print("No se encontró sesión SAP abierta. Abre SAP GUI y loguéate primero.")
        return
    _close_open_popups(session)

    if target == "all":
        tickets = _resolve_tickets(argv[1:])
        for fn in (diagnose_zcmr, diagnose_zsd, diagnose_vl06f):
            try:
                fn(session, tickets)
            except Exception as e:
                print(f"  ERROR en {fn.__name__}: {e}", file=sys.stderr)
            _close_open_popups(session)
    elif target in _DISPATCH:
        tickets = _resolve_tickets(argv[1:])
        _DISPATCH[target](session, tickets)
    else:
        print(f"Target desconocido: {target!r}. Usa: zcmr | zsd | vl06f | all | layouts")
        return

    print("\nDiagnóstico completado.")


if __name__ == "__main__":
    main()
