"""
diagnose_batch_6_8.py — Diagnóstico aislado de Batches 6-8 sobre tickets específicos.

Uso (un plant por ticket, formato 'plant:ticket'):
    python diagnose_batch_6_8.py 8710:301661838 8713:114252799

Uso (todos los tickets del mismo plant, usar --plant):
    python diagnose_batch_6_8.py --plant 8710 301661838 301661839

Ejecuta:
    Batch 6: Lee orders desde ZCMR (un F8 por plant) + intenta borrar.
    Batch 7: Verifica en ZSD_DEL_TICKETS (un solo F8 con todos los tickets).
    Batch 8: Si hay pendientes, intenta cancel vía VA02.

NO toca state.json ni reporte.xlsx. Output a stdout con timestamps + detalle SAP.
"""
import sys
import argparse
import traceback
from datetime import datetime

# Forzar UTF-8 en stdout/stderr — Windows usa cp1252 por default y rompe con
# caracteres unicode (acentos, flechas, etc.) al redirigir a archivo con `>`.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from sap_login import SapAutomation, CREDENTIALS_FILE
from sap_orders import (
    cancel_order_by_ticket,
    _read_zcmr_orders,
    _delete_intercompany_order_va02,
    _delete_intracompany_order_me22n,
)
from sap_utils import _navigate_to, _wait_ready, _POPUP_TABLE, _enter_multi_values
from verifications import verify_zsd_del_tickets


def _ts() -> str:
    return datetime.now().strftime("%H:%M:%S.%f")[:-3]


def _log(label: str, msg: str = "") -> None:
    print(f"[{_ts()}] {label:14s} {msg}", flush=True)


def _dump_va02_state(session, order: str, label: str) -> None:
    """Imprime estado actual de SAP (wnd[0] título, sbar, popups)."""
    try:
        title = session.findById("wnd[0]/titl").Text or ""
        _log(label, f"  wnd[0] título: {title!r}")
    except Exception:
        pass
    try:
        sbar = session.findById("wnd[0]/sbar")
        text = (sbar.Text or "").strip()
        mtype = (sbar.MessageType or "").strip()
        if text:
            _log(label, f"  sbar [{mtype}]: {text!r}")
    except Exception:
        pass
    for wnd in ("wnd[1]", "wnd[2]"):
        try:
            wtitle = session.findById(wnd).Text or ""
            _log(label, f"  {wnd} título: {wtitle!r}")
        except Exception:
            continue
        for field in ("txtMESSTXT1", "txtMESSTXT2", "txtMESSTXT3"):
            try:
                v = session.findById(f"{wnd}/usr/{field}").Text or ""
                if v:
                    _log(label, f"  {wnd}/usr/{field}: {v!r}")
            except Exception:
                pass


def _inspect_zcmr_raw(session, plant: str, tickets: list) -> None:
    """
    Inspección manual de ZCMR para diagnóstico. Setea plant + tickets y vuelca
    el grid principal en crudo.
    """
    _log("ZCMR INSPECT", f"Inspeccionando ZCMR con plant={plant!r} tickets={tickets}")
    try:
        _navigate_to(session, "ZCMR")
        _wait_ready(session)
    except Exception as e:
        _log("ZCMR INSPECT", f"  EXCEPCIÓN navegando: {e!r}")
        return

    # Setear plant (campo obligatorio)
    try:
        session.findById("wnd[0]/usr/ctxtP_PLANT-LOW").Text = plant
        session.findById("wnd[0]/usr/ctxtP_PLANT-LOW").SetFocus()
        _log("ZCMR INSPECT", f"  Seteado P_PLANT-LOW = {plant!r}")
    except Exception as e:
        _log("ZCMR INSPECT", f"  No se pudo setear P_PLANT-LOW: {e!r}")

    # Tickets vía popup multi-valor
    try:
        session.findById("wnd[0]/usr/btn%_P_TICKET_%_APP_%-VALU_PUSH").press()
        _wait_ready(session)
        try:
            session.findById("wnd[1]/tbar[0]/btn[16]").press()
            _wait_ready(session)
        except Exception:
            pass
        _enter_multi_values(session, _POPUP_TABLE, tickets)
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        _wait_ready(session)
        _log("ZCMR INSPECT", f"  Tickets cargados: {tickets}")
    except Exception as e:
        _log("ZCMR INSPECT", f"  EXCEPCIÓN aplicando tickets: {e!r}")
        return

    # F8
    try:
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        _wait_ready(session)
        _log("ZCMR INSPECT", "  F8 presionado.")
    except Exception as e:
        _log("ZCMR INSPECT", f"  EXCEPCIÓN F8: {e!r}")
        return

    # Status bar post-F8
    try:
        sbar = session.findById("wnd[0]/sbar")
        text = (sbar.Text or "").strip()
        mtype = (sbar.MessageType or "").strip()
        if text:
            _log("ZCMR INSPECT", f"  sbar post-F8 [{mtype}]: {text!r}")
        else:
            _log("ZCMR INSPECT", "  sbar post-F8: (vacío)")
    except Exception:
        pass

    # Grid principal — leer crudo
    try:
        grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        rc = grid.RowCount
        _log("ZCMR INSPECT", f"  Grid principal RowCount = {rc}")
        for i in range(min(rc, 20)):
            row_dump = {}
            for col in ("PLANT", "TICKET_DATE", "TICKET_LOADED", "TICKET_QTY",
                        "ORDER_TYPE", "TICKET_IN_SAP"):
                try:
                    v = grid.GetCellValue(i, col) or ""
                    if v:
                        row_dump[col] = v
                except Exception:
                    pass
            _log("ZCMR INSPECT", f"  Fila {i}: {row_dump}")
    except Exception as e:
        _log("ZCMR INSPECT", f"  EXCEPCIÓN leyendo grid: {e!r}")


def diagnose_grouped(by_plant: dict) -> None:
    """
    Ejecuta diagnóstico Batches 6-8 sobre tickets agrupados por plant.

    Args:
        by_plant: dict {plant: [ticket1, ticket2, ...]}
    """
    all_tickets = [t for tlist in by_plant.values() for t in tlist]
    _log("DIAGNÓSTICO", f"Plants a diagnosticar: {list(by_plant.keys())}")
    _log("DIAGNÓSTICO", f"Total tickets: {len(all_tickets)} -> {all_tickets}")
    _log("DIAGNÓSTICO", "Login SAP iniciando...")

    sap = SapAutomation(CREDENTIALS_FILE)
    sap.run()
    s1, s2, s3 = sap.session, sap.session2, sap.session3
    if not s1:
        _log("ERROR", "No se pudo abrir sesión SAP. ABORTAR.")
        return

    distintas = (s2 is not s1) and (s3 is not s1) and (s3 is not s2)
    _log("DIAGNÓSTICO",
         f"Sesiones distintas: {distintas} "
         f"(s1={hex(id(s1))}, s2={hex(id(s2))}, s3={hex(id(s3))})")
    if not distintas:
        _log("DIAGNÓSTICO",
             "ADVERTENCIA: las 3 sesiones no son distintas — diagnóstico continúa.")

    # ============================================================ #
    # PRE-BATCH 6 — Inspección ZCMR por cada plant                  #
    # ============================================================ #
    for plant, tickets in by_plant.items():
        _inspect_zcmr_raw(s3, plant, tickets)

    # ============================================================ #
    # BATCH 6 — ZCMR read + delete (por plant)                      #
    # ============================================================ #
    all_zcmr_orders = []
    for plant, tickets in by_plant.items():
        _log("BATCH 6", f"--- Plant {plant!r} con {len(tickets)} ticket(s) ---")
        try:
            orders = _read_zcmr_orders(s3, tickets, plant=plant)
            _log("BATCH 6", f"  ZCMR (plant={plant}) encontró {len(orders)} items:")
            for item in orders:
                _log("BATCH 6", f"    {item}")
                all_zcmr_orders.append(item)
        except Exception as e:
            _log("BATCH 6", f"  EXCEPCIÓN en _read_zcmr_orders: {e!r}")
            traceback.print_exc()

    if not all_zcmr_orders:
        _log("BATCH 6", "ZCMR no devolvió ningún item en ningún plant. ABORTAR.")
        return

    for item in all_zcmr_orders:
        ticket = item.get("ticket", "?")
        order = item.get("order", "?")
        delivery = item.get("delivery", "?")
        intra = item.get("is_intracompany", False)
        _log("BATCH 6",
             f"Procesando ticket={ticket} order={order} delivery={delivery} "
             f"intracompany={intra}")
        try:
            if intra:
                _delete_intracompany_order_me22n(s3, order)
                _log("BATCH 6", "  -> ME22N: sin excepción.")
                _dump_va02_state(s3, order, "BATCH 6 post-ME22N")
            else:
                _dump_va02_state(s3, order, "BATCH 6 pre-VA02")
                _delete_intercompany_order_va02(s3, order, delivery=ticket)
                _log("BATCH 6", "  -> VA02: sin excepción.")
                _dump_va02_state(s3, order, "BATCH 6 post-VA02")
        except Exception as e:
            _log("BATCH 6", f"  -> EXCEPCIÓN: {e!r}")
            _dump_va02_state(s3, order, "BATCH 6 ERR")
            traceback.print_exc()

    # ============================================================ #
    # BATCH 7 — ZSD_DEL_TICKETS (no requiere plant)                 #
    # ============================================================ #
    _log("BATCH 7", "Verificando en ZSD_DEL_TICKETS...")
    try:
        exitosos, fallidos, ticket_to_order = verify_zsd_del_tickets(s1, all_tickets)
        _log("BATCH 7", f"Exitosos (no pendientes): {exitosos}")
        _log("BATCH 7", f"Fallidos (pendientes):   {fallidos}")
        _log("BATCH 7", f"ticket_to_order: {ticket_to_order}")
    except Exception as e:
        _log("BATCH 7", f"EXCEPCIÓN: {e!r}")
        traceback.print_exc()
        ticket_to_order = {}

    # ============================================================ #
    # BATCH 8                                                       #
    # ============================================================ #
    if not ticket_to_order:
        _log("BATCH 8", "No hay órdenes pendientes para cancelar. FIN.")
        _log("DIAGNÓSTICO", "Diagnóstico completo.")
        return

    _log("BATCH 8", f"Cancelando {len(ticket_to_order)} órdenes pendientes...")
    for ticket, order in ticket_to_order.items():
        if not order:
            _log("BATCH 8",
                 f"  ticket {ticket}: order VACÍA — no se puede cancelar.")
            continue
        _log("BATCH 8", f"  Cancelando ticket={ticket} order={order}")
        try:
            cancel_order_by_ticket(s3, order, ticket)
            _log("BATCH 8", "  -> sin excepción.")
            _dump_va02_state(s3, order, "BATCH 8 post")
        except Exception as e:
            _log("BATCH 8", f"  -> EXCEPCIÓN: {e!r}")
            _dump_va02_state(s3, order, "BATCH 8 ERR")
            traceback.print_exc()

    _log("DIAGNÓSTICO", "Diagnóstico completo.")


def _parse_ticket_arg(arg: str, default_plant: str = "") -> tuple:
    """
    Parse 'plant:ticket' o solo 'ticket' (usa default_plant).
    Retorna (plant, ticket).
    """
    if ":" in arg:
        plant, ticket = arg.split(":", 1)
        return plant.strip(), ticket.strip()
    return (default_plant, arg.strip())


def main():
    parser = argparse.ArgumentParser(
        description="Diagnóstico aislado de Batches 6-8 sobre tickets específicos.",
        epilog="Ejemplos:\n"
               "  python diagnose_batch_6_8.py 8710:301661838 8713:114252799\n"
               "  python diagnose_batch_6_8.py --plant 8710 301661838 301661839",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--plant", default="",
        help="Plant a usar para todos los tickets que no especifiquen su propio "
             "plant (formato 'plant:ticket' tiene prioridad).",
    )
    parser.add_argument(
        "tickets", nargs="+",
        help="Tickets en formato 'plant:ticket' (ej: '8710:301661838') o solo "
             "'ticket' (usa --plant).",
    )
    args = parser.parse_args()

    # Agrupar por plant
    by_plant: dict = {}
    sin_plant = []
    for arg in args.tickets:
        plant, ticket = _parse_ticket_arg(arg, args.plant)
        if not plant:
            sin_plant.append(ticket)
            continue
        by_plant.setdefault(plant, []).append(ticket)

    if sin_plant:
        print(f"Error: tickets sin plant detectados: {sin_plant}", file=sys.stderr)
        print("Usa formato 'plant:ticket' o pasa --plant XXXX.", file=sys.stderr)
        sys.exit(1)

    if not by_plant:
        print("Error: no se proporcionaron tickets válidos.", file=sys.stderr)
        sys.exit(1)

    diagnose_grouped(by_plant)


if __name__ == "__main__":
    main()
