import argparse
import sys
import time
from datetime import date
from sap_login import SapAutomation, CREDENTIALS_FILE
from sap_vl06f import read_vl06f_data, delete_bol, delete_bol_bulk
from sap_batches import (
    delete_billing_documents_bulk,  # Batch 1 — bulk (una llamada VF11 para todos)
    delete_shipment_costs_all,      # Batch 2 — permanece en VI05
    delete_shipment_numbers_all,    # Batch 3 — permanece en VT02N
    reverse_pgi_bulk,               # Batch 4 — bulk (una llamada VL09 para todos)
)
from sap_orders import (
    delete_orders_from_zcmr, delete_orders_from_zsd, cancel_order_by_ticket,
    _delete_order_lines_for_tickets,
)
from verifications import (
    verify_billing_documents_bulk,
    verify_shipment_costs_bulk,
    verify_shipment_numbers_bulk,
    verify_pgi_reversed_bulk,
    verify_bol_deleted_bulk,
    verify_zsd_del_tickets,
)
from report_writer import write_report_xlsx
from checkpoint import (
    load_state, init_state, save_chunk_result, is_chunk_completed,
    consolidate_state, clear_state, get_failed_tickets_per_batch,
    has_pending_work,
)
from path import get_billing_file
from log_util import setup_logging


VALID_BATCHES = {1, 2, 3, 4, 5, 6, 7, 8}
DEFAULT_CHUNK_SIZE = 100


def parse_batch_spec(spec: str) -> set:
    """
    Parsea spec de batches en formato '1,2,3' o '6-8' o 'all' o '1,4,6-8'.
    Retorna un set de números de batch.
    """
    spec = spec.strip().lower()
    if spec == "all":
        return VALID_BATCHES.copy()

    batches = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            start, end = part.split("-", 1)
            batches.update(range(int(start.strip()), int(end.strip()) + 1))
        else:
            batches.add(int(part))

    invalid = batches - VALID_BATCHES
    if invalid:
        raise ValueError(
            f"Batches inválidos: {sorted(invalid)}. Válidos: {sorted(VALID_BATCHES)}"
        )
    return batches


def parse_args():
    parser = argparse.ArgumentParser(description="DIL Pipeline — automatización SAP")
    parser.add_argument(
        "--batches",
        default="all",
        help="Batches a ejecutar. Ej: 'all' (default), '1,2,3', '6-8', '1,4,6-8'",
    )
    parser.add_argument(
        "--chunk-size",
        type=int,
        default=DEFAULT_CHUNK_SIZE,
        help=f"Tamaño de chunk para procesar tickets (default {DEFAULT_CHUNK_SIZE}). "
             f"SAP VL09 tiene límite práctico ~60-100 por timeout de work process.",
    )
    parser.add_argument(
        "--fresh",
        action="store_true",
        help="Ignora cualquier state previo y arranca desde chunk 0.",
    )
    parser.add_argument(
        "--report-only",
        action="store_true",
        help="No toca SAP. Solo regenera el reporte xlsx desde state existente.",
    )
    parser.add_argument(
        "--manual-only",
        action="store_true",
        help="Modo recovery: procesa SOLO los pares (ticket, order) de la hoja "
             "'Manual Orders' del Excel. Corre Batches 6->7->8 únicamente "
             "(skip lectura VL06F y batches 0-5). Útil para orders huérfanas.",
    )
    parser.add_argument(
        "--retry-failed",
        action="store_true",
        help="Reintenta solo los tickets que quedaron como fallidos en el state "
             "previo. Reusa vl06f cacheado del state; NO re-lee VL06F. Encadena "
             "batches posteriores cuando el retry de un batch tiene éxito "
             "(cascada completa). Combinable con --batches para limitar a "
             "batches específicos.",
    )
    return parser.parse_args()


def print_report(resultados: dict, total_tickets: int,
                 tickets_no_encontrados: list,
                 tickets_con_factura: list = None,
                 tickets_critical_error: list = None) -> None:
    hoy = date.today().strftime("%Y-%m-%d")
    linea = "=" * 45

    tickets_con_factura = tickets_con_factura or []
    tickets_critical_error = tickets_critical_error or []
    no_encontrados = len(tickets_no_encontrados)
    con_factura = len(tickets_con_factura)
    critical_err = len(tickets_critical_error)
    # Tickets que entraron al pipeline = total - excluidos - no detectados - crasheados
    procesados = total_tickets - no_encontrados - con_factura - critical_err

    print(f"\n{linea}")
    print(f"  REPORTE DIL — {hoy}")
    print(f"{linea}")
    print(f"  Tickets en Excel:           {total_tickets}")
    print(f"  Procesados por el pipeline: {procesados}")
    if no_encontrados:
        print(f"  NO encontrados en VL06F:    {no_encontrados} — REQUIEREN INVESTIGACIÓN (deberían estar en VL06F)")
    if con_factura:
        print(f"  Con factura intercompany:   {con_factura} — EXCLUIDOS por R001 (cancelar factura primero)")
    if critical_err:
        print(f"  Error crítico (chunk):      {critical_err} — NO procesados (re-correr pipeline)")
    print()

    tickets_con_fallo = set()

    for batch_name, (exitosos, fallidos) in resultados.items():
        print(f"  {batch_name}")
        print(f"    Exitosos ({len(exitosos)}): {', '.join(exitosos) if exitosos else '—'}")
        print(f"    Fallidos ({len(fallidos)}): {', '.join(fallidos) if fallidos else '—'}")
        print()
        tickets_con_fallo.update(fallidos)

    completados = procesados - len(tickets_con_fallo)
    print(linea)
    print(f"  Exitosos (todos los batches): {completados}/{procesados} procesados ({completados}/{total_tickets} del Excel)")
    if tickets_con_fallo:
        print(f"  Con fallos parciales:         {len(tickets_con_fallo)}/{procesados}")
        print(f"  Revisar manualmente:          {', '.join(sorted(tickets_con_fallo))}")
    if no_encontrados:
        print(f"  NO encontrados en VL06F:      {no_encontrados}/{total_tickets} — pipeline no los detectó")
    if con_factura:
        print(f"  Con factura intercompany:     {con_factura}/{total_tickets} — excluidos por R001")
        print(f"    Tickets:                    {', '.join(tickets_con_factura)}")
    if critical_err:
        print(f"  Error crítico (no procesados):{critical_err}/{total_tickets} — chunk crasheó, RE-CORRER PIPELINE")
        print(f"    Tickets:                    {', '.join(tickets_critical_error)}")
    print(f"{linea}\n")


def _finalize_order_tracking(order_tracking: dict) -> None:
    """
    Calcula el `final_status` de cada entry del tracking según los estados de
    los 3 batches. Reglas:

      - Si batch_8_cancel == 'ok'                      -> deleted (recovery por B8)
      - Elif batch_7_verify == 'done' AND
             batch_6 == 'ok'                           -> deleted (camino feliz)
      - Elif batch_7_verify == 'pending' AND
             batch_8_cancel != 'ok'                    -> pending (recovery falló o no aplica)
      - Elif batch_6 == 'failed' AND
             batch_7_verify not in ('', 'done')        -> failed
      - Else                                           -> unknown
    """
    for ticket, tr in order_tracking.items():
        b6 = tr.get("batch_6", "")
        b7 = tr.get("batch_7_verify", "")
        b8 = tr.get("batch_8_cancel", "")
        if b8 == "ok":
            tr["final_status"] = "deleted"
        elif b7 == "done" and b6 == "ok":
            tr["final_status"] = "deleted"
        elif b7 == "pending":
            tr["final_status"] = "pending"
        elif b6 == "failed" and b7 != "done":
            tr["final_status"] = "failed"
        else:
            tr["final_status"] = tr.get("final_status", "unknown") or "unknown"


def _cancel_orders_deduped(session, ticket_to_order: dict, label: str,
                           cancel_failures: dict) -> tuple:
    """
    Batch 8: cancela órdenes vía VA02 DEDUPLICANDO. Muchos tickets pueden apuntar
    a la misma orden; sin dedup, la misma orden se intentaba borrar N veces (ej.
    50x la orden 1150084576). Aquí cada orden distinta se intenta UNA vez y el
    resultado (éxito/fallo) se propaga a todos sus tickets.

    Retorna (exitosos, fallidos) como listas de tickets.
    """
    orden_a_tickets: dict = {}
    exitosos, fallidos = [], []
    for ticket, order in ticket_to_order.items():
        if not order:
            fallidos.append(ticket)
            continue
        orden_a_tickets.setdefault(order, []).append(ticket)

    print(f"  {label}: {len(orden_a_tickets)} órdenes distintas "
          f"({len(ticket_to_order)} tickets)...")

    for order, tickets in orden_a_tickets.items():
        try:
            # SEGURO: borra solo las líneas de los tickets en alcance de esta
            # orden (o la orden completa si TODAS sus líneas lo están).
            _delete_order_lines_for_tickets(session, order, set(tickets))
            exitosos.extend(tickets)
            for t in tickets:
                cancel_failures.pop(t, None)
        except Exception as e:
            print(f"    [{label}] order {order} ({len(tickets)} ticket(s)): {e}",
                  file=sys.stderr)
            fallidos.extend(tickets)
            for t in tickets:
                cancel_failures[t] = order
    return exitosos, fallidos


def process_manual_chunk(
    pairs: dict,
    session1, session3,
) -> tuple:
    """
    Procesa pares (ticket, order) manuales — modo recovery --manual-only.

    NO toca VL06F. Skip Batches 0-5. Solo corre:
      - Batch 6: cancel_order_by_ticket por cada (ticket, order) — dispatch
        a VA02 (intercompany) o ME22N (intracompany) según prefijo de la orden
      - Batch 7: verificación en ZSD_DEL_TICKETS
      - Batch 8: retry line-level delete para los que ZSD sigue mostrando

    Args:
        pairs: dict {ticket: order} cargado desde hoja 'Manual Orders'.
        session1: para Batch 7 (ZSD_DEL_TICKETS).
        session3: para Batches 6 y 8 (VA02 cancel).

    Returns la misma tupla que process_chunk() para compatibilidad con
    el state y el reporter.
    """
    resultados: dict = {}
    zcmr_failures: dict = {}
    ticket_to_order: dict = {}
    cancel_failures: dict = {}

    tickets = list(pairs.keys())

    # --- Batch 6 (manual): line-level delete por cada (ticket, order) ---
    # Usa cancel_order_by_ticket: borra solo la línea de la order que matchea
    # el ticket (por PO Number = ticket padded), NO la order completa. Si la
    # order queda vacía SAP la borra; si quedan otras líneas, persiste.
    print(f"\n  Batch 6 (manual): Eliminando líneas de {len(pairs)} pares (ticket, order) vía VA02...")
    exitosos_6, fallidos_6 = [], []
    for ticket, order in pairs.items():
        try:
            cancel_order_by_ticket(session3, order, ticket)
            exitosos_6.append(ticket)
        except Exception as e:
            print(f"    [VA02] order {order} (ticket {ticket}): {e}", file=sys.stderr)
            fallidos_6.append(ticket)
            zcmr_failures[ticket] = [order]
    resultados["BATCH 6 — ZCMR Orders"] = (exitosos_6, fallidos_6)

    # --- Batch 7: verificar en ZSD_DEL_TICKETS ---
    print("  Batch 7 (manual): Verificación final con ZSD_DEL_TICKETS...")
    exitosos_7, fallidos_7, zsd_ticket_to_order = verify_zsd_del_tickets(session1, tickets)
    resultados["BATCH 7 — Verificación Final ZSD"] = (exitosos_7, fallidos_7)

    # Override: si ZSD no pudo mapear ticket->order (layout oculto, etc.),
    # usar la mapping manual para que Batch 8 sepa qué cancelar.
    ticket_to_order = dict(zsd_ticket_to_order)
    for ticket in fallidos_7:
        if ticket not in ticket_to_order and ticket in pairs:
            ticket_to_order[ticket] = pairs[ticket]

    # Ground truth: si ZSD muestra pendiente, Batch 6 NO fue realmente exitoso
    if ticket_to_order:
        tickets_pendientes = set(ticket_to_order.keys())
        new_exit6 = [t for t in exitosos_6 if t not in tickets_pendientes]
        new_fail6 = sorted(set(fallidos_6) | tickets_pendientes)
        if len(new_exit6) != len(exitosos_6):
            print(f"    [Batch 6 corregido por ZSD] {len(exitosos_6) - len(new_exit6)} "
                  f"ticket(s) movido(s) de exitoso a fallido")
        resultados["BATCH 6 — ZCMR Orders"] = (new_exit6, new_fail6)

    # --- Batch 8: retry line-level delete para los que siguen pendientes ---
    # Mismo enfoque que Batch 6: line-level matching por ticket, NO order-level.
    if ticket_to_order:
        print(f"  Batch 8 (manual): Retry line-delete de {len(ticket_to_order)} líneas vía VA02...")
        cancel_exitosos, cancel_fallidos = [], []
        for ticket, order in ticket_to_order.items():
            try:
                cancel_order_by_ticket(session3, order, ticket)
                cancel_exitosos.append(ticket)
            except Exception as e:
                print(f"    [Batch 8] order {order} (ticket {ticket}): {e}", file=sys.stderr)
                cancel_fallidos.append(ticket)
                cancel_failures[ticket] = order
        resultados["BATCH 8 — Order Cancellation"] = (cancel_exitosos, cancel_fallidos)
    else:
        print("  Batch 8 (manual): Sin órdenes pendientes — omitido.")

    # vl06f vacío: en modo manual no leemos VL06F. tickets_con_factura tampoco
    # aplica en modo manual (R001 solo se evalúa sobre lecturas de VL06F).
    # order_tracking en modo manual queda vacío (este modo no usa ZSD para
    # discover, usa pares (ticket, order) ya conocidos del sheet).
    return (resultados, {}, zcmr_failures, ticket_to_order, cancel_failures,
            [], [], {})


def process_chunk(
    tickets: list,
    session1, session2, session3,
    batches_to_run: set,
    needs_vl06f: bool,
) -> tuple:
    """
    Procesa UN chunk de tickets a través de todos los batches seleccionados.
    Retorna: (resultados, vl06f, zcmr_failures, ticket_to_order, cancel_failures,
              tickets_no_encontrados, tickets_con_factura)
    """
    # --- 3.5. Leer datos de VL06F para este chunk ---
    if needs_vl06f:
        print("  Leyendo datos de VL06F...")
        vl06f = read_vl06f_data(session1, tickets)
        print(f"    {len(vl06f)} deliveries encontrados en VL06F.")

        tickets_no_encontrados = [t for t in tickets if t not in vl06f]

        # Auto-recovery: si la primera lectura no encontró todos los tickets,
        # reintentar UNA vez solo con los faltantes. Ataca fallos transitorios
        # (filtro de fecha residual, layout no aplicado, popup) sin tocar los
        # ya leídos.
        if tickets_no_encontrados:
            print(f"    [Auto-recovery VL06F] Re-leyendo VL06F para "
                  f"{len(tickets_no_encontrados)} tickets no encontrados...")
            try:
                retry_vl06f = read_vl06f_data(session1, tickets_no_encontrados)
            except Exception as e:
                print(f"    [Auto-recovery VL06F] Error en re-lectura: {e}",
                      file=sys.stderr)
                retry_vl06f = {}
            if retry_vl06f:
                vl06f.update(retry_vl06f)
                recuperados = list(retry_vl06f.keys())
                tickets_no_encontrados = [t for t in tickets if t not in vl06f]
                print(f"    [Auto-recovery VL06F] Recuperados: {len(recuperados)}. "
                      f"Siguen sin encontrarse: {len(tickets_no_encontrados)}")

        coincidencias = sum(1 for t in tickets if t in vl06f)
        print(f"    [DEBUG] Coincidencias ticket<->VBELN: {coincidencias} de {len(tickets)}")

        if tickets_no_encontrados:
            print(
                f"    Advertencia: {len(tickets_no_encontrados)} tickets sin datos en VL06F: "
                f"{', '.join(tickets_no_encontrados)}",
                file=sys.stderr,
            )

        # R001: tickets con invoice_il que empieza con '7' tienen factura
        # intercompany — se EXCLUYEN del procesamiento completo (todos los
        # batches). Requieren cancelación manual de la factura primero.
        tickets_con_factura = [
            t for t in tickets
            if t in vl06f and vl06f[t].get("invoice_il", "").startswith("7")
        ]
        if tickets_con_factura:
            print(f"    [R001] {len(tickets_con_factura)} tickets con factura "
                  f"intercompany — EXCLUIDOS del pipeline (todos los batches).")

        tickets_activos = [
            t for t in tickets
            if t in vl06f and not vl06f[t].get("invoice_il", "").startswith("7")
        ]
    else:
        print("  Omitiendo lectura de VL06F (batches 1-5 no seleccionados).")
        tickets_con_factura = []
        vl06f = {
            t: {"billing_doc": "", "shpt_cst": "", "shipment": "",
                "wbstk": "", "delivery": t, "invoice_il": "", "block": ""}
            for t in tickets
        }
        tickets_no_encontrados = []
        tickets_activos = tickets[:]

    resultados = {}
    zcmr_failures: dict = {}
    ticket_to_order: dict = {}
    cancel_failures: dict = {}
    order_tracking: dict = {}  # tracking por ticket de su trayectoria en B6→B8

    # --- Batch 1: Billing Documents — BULK (VF11) ---
    if 1 in batches_to_run:
        print("\n  Batch 1: Billing Documents (bulk)...")
        billing_docs = [vl06f[t]["billing_doc"] for t in tickets_activos if vl06f[t]["billing_doc"]]
        print(f"    {len(billing_docs)} billing documents a eliminar.")
        try:
            if billing_docs:
                delete_billing_documents_bulk(session2, billing_docs)
        except Exception as e:
            print(f"    [Batch 1] Error en operación bulk: {e}", file=sys.stderr)
        exitosos, fallidos = verify_billing_documents_bulk(session1, tickets_activos)
        resultados["BATCH 1 — Billing Documents"] = (exitosos, fallidos)
        tickets_activos = exitosos

    # --- Batch 2: Shipment Cost — VI05 ---
    if 2 in batches_to_run:
        print("  Batch 2: Shipment Cost (stay-in-VI05)...")
        shpt_csts = [vl06f[t]["shpt_cst"] for t in tickets_activos if vl06f[t]["shpt_cst"]]
        print(f"    {len(shpt_csts)} shipment costs a eliminar.")
        try:
            delete_shipment_costs_all(session2, shpt_csts)
        except Exception as e:
            print(f"    [Batch 2] Error en operación: {e}", file=sys.stderr)
        exitosos, fallidos = verify_shipment_costs_bulk(session1, tickets_activos)
        resultados["BATCH 2 — Shipment Cost"] = (exitosos, fallidos)
        tickets_activos = exitosos

    # --- Batch 3: Shipment Number — VT02N ---
    if 3 in batches_to_run:
        print("  Batch 3: Shipment Number (stay-in-VT02N)...")
        shipments = [vl06f[t]["shipment"] for t in tickets_activos if vl06f[t]["shipment"]]
        print(f"    {len(shipments)} shipment numbers a eliminar.")
        try:
            delete_shipment_numbers_all(session2, shipments)
        except Exception as e:
            print(f"    [Batch 3] Error en operación: {e}", file=sys.stderr)
        exitosos, fallidos = verify_shipment_numbers_bulk(session1, tickets_activos)
        resultados["BATCH 3 — Shipment Number"] = (exitosos, fallidos)
        tickets_activos = exitosos

    # --- Batch 4: Reverse PGI — VL09 ---
    if 4 in batches_to_run:
        print("  Batch 4: Reverse PGI (bulk)...")
        # Solo reversar deliveries con PGI activo (wbstk distinto de "A" Y "").
        # wbstk=="" = nunca hubo PGI -> nada que reversar; wbstk=="A" = ya reversed.
        deliveries_pgi = [
            vl06f[t]["delivery"] for t in tickets_activos
            if vl06f[t]["wbstk"] not in ("A", "") and vl06f[t]["delivery"]
        ]
        print(f"    {len(deliveries_pgi)} deliveries a reversar.")
        try:
            if deliveries_pgi:
                reverse_pgi_bulk(session2, deliveries_pgi)
        except Exception as e:
            print(f"    [Batch 4] Error en operación bulk: {e}", file=sys.stderr)
        exitosos, fallidos = verify_pgi_reversed_bulk(session1, tickets_activos)
        resultados["BATCH 4 — Reverse PGI"] = (exitosos, fallidos)
        tickets_activos = exitosos

    # --- Batch 5: Eliminar BOL numbers — VL06F (per-ticket) ---
    # FLUJO:
    #   1. ACTION: delete_bol por ticket (con step-by-step diagnostics y status
    #      bar capture en sap_vl06f.delete_bol).
    #   2. VERIFY: re-leer VL06F con verify_bol_deleted_bulk para confirmar
    #      cuáles realmente se borraron. Sin este verify, fallos silenciosos
    #      (SAP rechaza sin popup) se reportarían como exitosos por error.
    if 5 in batches_to_run:
        print("  Batch 5: Eliminar BOL numbers (per-ticket)...")
        action_errors: dict = {}  # ticket -> mensaje de error de delete_bol
        for ticket in tickets_activos:
            try:
                delete_bol(session1, vl06f[ticket]["delivery"])
            except Exception as e:
                print(f"    [BOL] Error en {ticket}: {e}", file=sys.stderr)
                action_errors[ticket] = str(e)

        # VERIFY post-batch: re-leer VL06F. Source of truth para clasificar.
        print(f"  Batch 5: Verificando con VL06F que las deliveries se borraron...")
        exitosos, fallidos = verify_bol_deleted_bulk(session1, tickets_activos)
        if action_errors:
            print(f"  Batch 5: {len(action_errors)} tickets lanzaron error en delete_bol; "
                  f"verify confirma {len([t for t in action_errors if t in fallidos])} "
                  f"siguen en VL06F.")
        resultados["BATCH 5 — BOL Deletion"] = (exitosos, fallidos)
        tickets_activos = exitosos

    # --- Batch 6: Eliminar Orders vía ZSD_DEL_TICKETS — session1 (lee) + session3 (borra) ---
    # Usamos ZSD en vez de ZCMR porque ZCMR filtra los tickets cuyo BOL ya
    # fue eliminado por Batch 5 (creando un race condition arquitectónico).
    # ZSD_DEL_TICKETS muestra el mapping ticket->order independientemente del
    # estado del BOL.
    if 6 in batches_to_run:
        print("  Batch 6: Eliminando orders vía ZSD_DEL_TICKETS...")
        try:
            zcmr_failures, b6_tracking = delete_orders_from_zsd(
                session1, session3, tickets_activos)
            order_tracking.update(b6_tracking)
            fallidos_6 = list(zcmr_failures.keys())
            exitosos_6 = [t for t in tickets_activos if t not in zcmr_failures]
            resultados["BATCH 6 — ZCMR Orders"] = (exitosos_6, fallidos_6)
        except Exception as e:
            print(f"    [ZSD/Batch 6] Error general: {e}", file=sys.stderr)
            resultados["BATCH 6 — ZCMR Orders"] = ([], tickets_activos[:])

    # --- Batch 7: Verificación final con ZSD_DEL_TICKETS — session1 ---
    if 7 in batches_to_run:
        print("  Batch 7: Verificación final con ZSD_DEL_TICKETS...")
        exitosos, fallidos, ticket_to_order = verify_zsd_del_tickets(session1, tickets_activos)
        resultados["BATCH 7 — Verificación Final ZSD"] = (exitosos, fallidos)
        if ticket_to_order:
            print(f"    Orders pendientes en ZSD_DEL_TICKETS: {ticket_to_order}")
        else:
            print("    Sin orders pendientes en ZSD_DEL_TICKETS.")

        # Actualizar tracking con resultado de B7:
        #   - ticket que NO aparece como pendiente en ZSD → batch_7_verify = "done"
        #   - ticket que sí aparece pendiente → batch_7_verify = "pending"
        pendientes_b7 = set(ticket_to_order.keys())
        for t in tickets_activos:
            if t in order_tracking:
                if t in pendientes_b7:
                    order_tracking[t]["batch_7_verify"] = "pending"
                else:
                    order_tracking[t]["batch_7_verify"] = "done"

        # Ground truth: si ZSD muestra que la orden sigue pendiente, Batch 6 NO
        # fue realmente exitoso aunque no haya lanzado excepción.
        if 6 in batches_to_run and ticket_to_order:
            old_exit6, old_fail6 = resultados.get("BATCH 6 — ZCMR Orders", ([], []))
            tickets_con_pendientes = set(ticket_to_order.keys())
            new_exit6 = [t for t in old_exit6 if t not in tickets_con_pendientes]
            new_fail6 = sorted(set(old_fail6) | tickets_con_pendientes)
            if len(new_exit6) != len(old_exit6):
                print(f"    [Batch 6 corregido por ZSD] {len(old_exit6) - len(new_exit6)} "
                      f"ticket(s) movido(s) de exitoso a fallido")
            resultados["BATCH 6 — ZCMR Orders"] = (new_exit6, new_fail6)

    # --- Batch 8: Cancelar orders pendientes vía VA02 — session3 ---
    if 8 in batches_to_run:
        if ticket_to_order:
            cancel_exitosos, cancel_fallidos = _cancel_orders_deduped(
                session3, ticket_to_order, "Batch 8", cancel_failures)
            resultados["BATCH 8 — Order Cancellation"] = (cancel_exitosos, cancel_fallidos)
            # Actualizar tracking con resultado de B8
            for t in cancel_exitosos:
                if t in order_tracking:
                    order_tracking[t]["batch_8_cancel"] = "ok"
            for t in cancel_fallidos:
                if t in order_tracking:
                    order_tracking[t]["batch_8_cancel"] = "failed"
                    if not order_tracking[t]["error_msg"]:
                        order_tracking[t]["error_msg"] = "Batch 8 cancel falló"
        else:
            print("  Batch 8: Sin órdenes pendientes — omitido.")

    # --- Consolidar final_status del tracking ---
    _finalize_order_tracking(order_tracking)

    return (resultados, vl06f, zcmr_failures, ticket_to_order, cancel_failures,
            tickets_no_encontrados, tickets_con_factura, order_tracking)


def _result_to_chunk_dict(result: tuple, chunk_tickets: list) -> dict:
    """
    Convierte la tupla retornada por process_chunk/process_manual_chunk en un
    dict con la misma estructura que un chunk persistido en el state. Útil para
    pasar el resultado de un chunk recién procesado a process_chunk_retry sin
    tener que ir al disco.
    """
    # Unpack tolerante a 7 u 8 elementos (retro-compat con tuplas viejas)
    if len(result) == 8:
        (resultados, vl06f, zcmr_failures, ticket_to_order,
         cancel_failures, tickets_no_encontrados, tickets_con_factura,
         order_tracking) = result
    else:
        (resultados, vl06f, zcmr_failures, ticket_to_order,
         cancel_failures, tickets_no_encontrados, tickets_con_factura) = result
        order_tracking = {}
    return {
        "tickets": list(chunk_tickets),
        "vl06f": vl06f,
        "resultados": {
            name: [list(exitosos), list(fallidos)]
            for name, (exitosos, fallidos) in resultados.items()
        },
        "zcmr_failures": zcmr_failures,
        "ticket_to_order": ticket_to_order,
        "cancel_failures": cancel_failures,
        "tickets_no_encontrados": list(tickets_no_encontrados),
        "tickets_con_factura": list(tickets_con_factura),
        "order_tracking": order_tracking,
    }


def _merge_retry_into_prev(prev_resultados: dict, nuevos: dict) -> dict:
    """
    Combina resultados de retry con los previos del state.

    Para cada batch reintentado, los nuevos exitosos se SUMAN a los previos
    exitosos y los nuevos fallidos REEMPLAZAN a los previos fallidos solo para
    los tickets reintentados. Los tickets que no se reintentaron mantienen su
    clasificación previa.
    """
    merged = dict(prev_resultados)
    for name, (new_exit, new_fail) in nuevos.items():
        prev_exit, prev_fail = prev_resultados.get(name, ([], []))
        retried = set(new_exit) | set(new_fail)
        merged_exit = sorted(set(prev_exit) | set(new_exit))
        merged_fail = sorted((set(prev_fail) - retried) | set(new_fail))
        merged[name] = (merged_exit, merged_fail)
    return merged


def process_chunk_retry(
    prev_chunk: dict,
    session1, session2, session3,
    batches_to_run: set,
) -> tuple:
    """
    Reintenta los batches con fallidos de un chunk previamente completado.

    Reusa vl06f cacheado del state (no re-lee VL06F para tickets que ya estaban),
    pero SÍ re-lee VL06F para los tickets_no_encontrados previos: si su ausencia
    fue transitoria, ahora aparecerán y se inyectan al pool desde Batch 1.
    Cascada completa: si el Batch N retry pasa, encadena Batches N+1, N+2, ...
    Mergea resultados con los previos antes de retornar.
    """
    vl06f = dict(prev_chunk["vl06f"])  # copia para no mutar el state directamente
    prev_resultados = {
        name: (list(par[0]), list(par[1]))
        for name, par in prev_chunk["resultados"].items()
    }

    # Re-leer VL06F para los tickets que no se encontraron en la corrida previa.
    # Si la causa fue transitoria, ahora aparecerán y entrarán al flujo Batch 1+.
    no_encontrados_prev = list(prev_chunk.get("tickets_no_encontrados", []))
    tickets_recuperados: list = []
    if no_encontrados_prev:
        print(f"    [Retry VL06F] Re-leyendo VL06F para "
              f"{len(no_encontrados_prev)} tickets no encontrados previamente...")
        try:
            nuevo_vl06f = read_vl06f_data(session1, no_encontrados_prev)
        except Exception as e:
            print(f"    [Retry VL06F] Error: {e}", file=sys.stderr)
            nuevo_vl06f = {}
        if nuevo_vl06f:
            vl06f.update(nuevo_vl06f)
            tickets_recuperados = list(nuevo_vl06f.keys())
            print(f"    [Retry VL06F] Recuperados: {len(tickets_recuperados)}")
        no_encontrados_actuales = [t for t in no_encontrados_prev if t not in vl06f]
        if no_encontrados_actuales:
            print(f"    [Retry VL06F] Siguen sin encontrarse: "
                  f"{len(no_encontrados_actuales)}")
    else:
        no_encontrados_actuales = []

    failed_per_batch = get_failed_tickets_per_batch(prev_chunk, batches_to_run)

    if not failed_per_batch and not tickets_recuperados:
        print("    Sin fallidos ni tickets recuperados — skip.")
        return (
            prev_resultados, vl06f,
            dict(prev_chunk.get("zcmr_failures", {})),
            dict(prev_chunk.get("ticket_to_order", {})),
            dict(prev_chunk.get("cancel_failures", {})),
            no_encontrados_actuales,
            list(prev_chunk.get("tickets_con_factura", [])),
            dict(prev_chunk.get("order_tracking", {})),
        )

    # Tickets recuperados arrancan desde Batch 1 (nunca pasaron por ninguno).
    # Se inyectan como "fallidos virtuales de Batch 1" para que entren al flujo.
    if tickets_recuperados:
        existing_b1 = set(failed_per_batch.get(1, []))
        failed_per_batch[1] = sorted(existing_b1 | set(tickets_recuperados))

    earliest = min(failed_per_batch.keys())
    tickets_activos = list(failed_per_batch[earliest])
    print(f"    Reintentando desde Batch {earliest} con {len(tickets_activos)} tickets.")

    nuevos: dict = {}
    zcmr_failures: dict = dict(prev_chunk.get("zcmr_failures", {}))
    ticket_to_order: dict = dict(prev_chunk.get("ticket_to_order", {}))
    cancel_failures: dict = dict(prev_chunk.get("cancel_failures", {}))

    # --- Batch 1: Billing Documents — BULK (VF11) ---
    if 1 in batches_to_run and 1 >= earliest and tickets_activos:
        print("\n  Batch 1 (retry): Billing Documents (bulk)...")
        billing_docs = [vl06f[t]["billing_doc"] for t in tickets_activos if vl06f[t]["billing_doc"]]
        try:
            if billing_docs:
                delete_billing_documents_bulk(session2, billing_docs)
        except Exception as e:
            print(f"    [Batch 1 retry] {e}", file=sys.stderr)
        exitosos, fallidos = verify_billing_documents_bulk(session1, tickets_activos)
        nuevos["BATCH 1 — Billing Documents"] = (exitosos, fallidos)
        tickets_activos = exitosos

    # --- Batch 2: Shipment Cost — VI05 ---
    if 2 in batches_to_run and 2 >= earliest and tickets_activos:
        print("  Batch 2 (retry): Shipment Cost...")
        shpt_csts = [vl06f[t]["shpt_cst"] for t in tickets_activos if vl06f[t]["shpt_cst"]]
        try:
            if shpt_csts:
                delete_shipment_costs_all(session2, shpt_csts)
        except Exception as e:
            print(f"    [Batch 2 retry] {e}", file=sys.stderr)
        exitosos, fallidos = verify_shipment_costs_bulk(session1, tickets_activos)
        nuevos["BATCH 2 — Shipment Cost"] = (exitosos, fallidos)
        tickets_activos = exitosos

    # --- Batch 3: Shipment Number — VT02N ---
    if 3 in batches_to_run and 3 >= earliest and tickets_activos:
        print("  Batch 3 (retry): Shipment Number...")
        shipments = [vl06f[t]["shipment"] for t in tickets_activos if vl06f[t]["shipment"]]
        try:
            if shipments:
                delete_shipment_numbers_all(session2, shipments)
        except Exception as e:
            print(f"    [Batch 3 retry] {e}", file=sys.stderr)
        exitosos, fallidos = verify_shipment_numbers_bulk(session1, tickets_activos)
        nuevos["BATCH 3 — Shipment Number"] = (exitosos, fallidos)
        tickets_activos = exitosos

    # --- Batch 4: Reverse PGI — VL09 ---
    if 4 in batches_to_run and 4 >= earliest and tickets_activos:
        print("  Batch 4 (retry): Reverse PGI (bulk)...")
        # Solo reversar deliveries con PGI activo (wbstk distinto de "A" Y "").
        deliveries_pgi = [
            vl06f[t]["delivery"] for t in tickets_activos
            if vl06f[t].get("wbstk") not in ("A", "") and vl06f[t].get("delivery")
        ]
        try:
            if deliveries_pgi:
                reverse_pgi_bulk(session2, deliveries_pgi)
        except Exception as e:
            print(f"    [Batch 4 retry] {e}", file=sys.stderr)
        exitosos, fallidos = verify_pgi_reversed_bulk(session1, tickets_activos)
        nuevos["BATCH 4 — Reverse PGI"] = (exitosos, fallidos)
        tickets_activos = exitosos

    # --- Batch 5: BOL Deletion — per-ticket ---
    if 5 in batches_to_run and 5 >= earliest and tickets_activos:
        print("  Batch 5 (retry): BOL Deletion (per-ticket)...")
        for ticket in tickets_activos:
            try:
                delete_bol(session1, vl06f[ticket]["delivery"])
            except Exception as e:
                print(f"    [BOL retry] {ticket}: {e}", file=sys.stderr)
        exitosos, fallidos = verify_bol_deleted_bulk(session1, tickets_activos)
        nuevos["BATCH 5 — BOL Deletion"] = (exitosos, fallidos)
        tickets_activos = exitosos

    # --- Batch 6: ZSD Orders — session1 (lee) + session3 (borra) ---
    if 6 in batches_to_run and 6 >= earliest and tickets_activos:
        print("  Batch 6 (retry): Eliminando orders vía ZSD_DEL_TICKETS...")
        try:
            new_zcmr = delete_orders_from_zsd(session1, session3, tickets_activos)
            fallidos_6 = list(new_zcmr.keys())
            exitosos_6 = [t for t in tickets_activos if t not in new_zcmr]
            nuevos["BATCH 6 — ZCMR Orders"] = (exitosos_6, fallidos_6)
            # Mergear failures de orders al dict acumulado
            for t, orders in new_zcmr.items():
                zcmr_failures[t] = orders
            # Limpiar entradas previas de zcmr_failures para tickets que ahora pasaron
            for t in exitosos_6:
                zcmr_failures.pop(t, None)
        except Exception as e:
            print(f"    [ZCMR retry] Error general: {e}", file=sys.stderr)
            nuevos["BATCH 6 — ZCMR Orders"] = ([], list(tickets_activos))

    # --- Batch 7: Verificación Final ZSD ---
    # Si Batch 6 corrió en retry, tickets_activos para B7 son los exitosos de B6.
    # Si solo se reintenta B7 directamente, usar los fallidos previos de B7.
    if 7 in batches_to_run and 7 >= earliest:
        if 6 in batches_to_run and "BATCH 6 — ZCMR Orders" in nuevos:
            tickets_para_7 = nuevos["BATCH 6 — ZCMR Orders"][0]
        else:
            tickets_para_7 = list(failed_per_batch.get(7, []))
        if tickets_para_7:
            print("  Batch 7 (retry): Verificación final con ZSD_DEL_TICKETS...")
            exitosos, fallidos, new_tto = verify_zsd_del_tickets(session1, tickets_para_7)
            nuevos["BATCH 7 — Verificación Final ZSD"] = (exitosos, fallidos)
            # Refrescar ticket_to_order con el resultado del retry
            for t in tickets_para_7:
                ticket_to_order.pop(t, None)
            ticket_to_order.update(new_tto)

            # Ground truth: si ZSD dice pendiente, Batch 6 NO fue exitoso
            if 6 in batches_to_run and new_tto and "BATCH 6 — ZCMR Orders" in nuevos:
                old_exit6, old_fail6 = nuevos["BATCH 6 — ZCMR Orders"]
                pendientes = set(new_tto.keys())
                new_exit6 = [t for t in old_exit6 if t not in pendientes]
                new_fail6 = sorted(set(old_fail6) | pendientes)
                if len(new_exit6) != len(old_exit6):
                    print(f"    [Batch 6 retry corregido por ZSD] "
                          f"{len(old_exit6) - len(new_exit6)} ticket(s) movido(s) a fallido")
                nuevos["BATCH 6 — ZCMR Orders"] = (new_exit6, new_fail6)

    # --- Batch 8: Order Cancellation retry ---
    if 8 in batches_to_run and 8 >= earliest:
        # Tickets candidatos: los detectados por B7 retry, o si solo se reintenta
        # B8, los fallidos previos de B8 (que tienen mapping en ticket_to_order).
        if "BATCH 7 — Verificación Final ZSD" in nuevos:
            tickets_para_8 = list(ticket_to_order.keys())
        else:
            prev_b8_fail = failed_per_batch.get(8, [])
            tickets_para_8 = [t for t in prev_b8_fail if t in ticket_to_order]
        if tickets_para_8:
            # Dedup: restringir ticket_to_order a los tickets de este retry.
            tto_retry = {t: ticket_to_order.get(t) for t in tickets_para_8}
            cancel_exitosos, cancel_fallidos = _cancel_orders_deduped(
                session3, tto_retry, "Batch 8 retry", cancel_failures)
            nuevos["BATCH 8 — Order Cancellation"] = (cancel_exitosos, cancel_fallidos)

    merged = _merge_retry_into_prev(prev_resultados, nuevos)
    return (merged, vl06f, zcmr_failures, ticket_to_order, cancel_failures,
            no_encontrados_actuales,
            list(prev_chunk.get("tickets_con_factura", [])),
            dict(prev_chunk.get("order_tracking", {})))


def process_manual_chunk_retry(prev_chunk: dict, session1, session3) -> tuple:
    """
    Reintenta los pares (ticket, order) con fallidos de un chunk manual previo.

    Reusa el mapping ticket_to_order del state y aplica cascada en
    Batch 6 -> Batch 7 -> Batch 8 (que es el flujo completo del modo manual).
    """
    prev_resultados = {
        name: (list(par[0]), list(par[1]))
        for name, par in prev_chunk["resultados"].items()
    }
    prev_tto = dict(prev_chunk.get("ticket_to_order", {}))
    zcmr_failures: dict = dict(prev_chunk.get("zcmr_failures", {}))
    cancel_failures: dict = dict(prev_chunk.get("cancel_failures", {}))
    ticket_to_order: dict = dict(prev_tto)

    failed_per_batch = get_failed_tickets_per_batch(prev_chunk)
    if not failed_per_batch:
        print("    Sin fallidos para reintentar en este chunk manual — skip.")
        return (prev_resultados, {}, zcmr_failures, ticket_to_order,
                cancel_failures, [], [],
                dict(prev_chunk.get("order_tracking", {})))

    earliest = min(failed_per_batch.keys())
    tickets_activos = list(failed_per_batch[earliest])
    print(f"    Reintentando (manual) desde Batch {earliest} con "
          f"{len(tickets_activos)} tickets.")

    nuevos: dict = {}

    # Batch 6 (manual): line-level delete por (ticket, order)
    if 6 >= earliest and tickets_activos:
        print(f"\n  Batch 6 retry (manual): {len(tickets_activos)} pares (ticket, order)...")
        exitosos_6, fallidos_6 = [], []
        for ticket in tickets_activos:
            order = prev_tto.get(ticket) or (zcmr_failures.get(ticket) or [None])[0]
            if not order:
                print(f"    [Manual B6] {ticket}: sin order mapeada — skip.", file=sys.stderr)
                fallidos_6.append(ticket)
                continue
            try:
                cancel_order_by_ticket(session3, order, ticket)
                exitosos_6.append(ticket)
                zcmr_failures.pop(ticket, None)
            except Exception as e:
                print(f"    [Manual B6] order {order} (ticket {ticket}): {e}", file=sys.stderr)
                fallidos_6.append(ticket)
                zcmr_failures[ticket] = [order]
        nuevos["BATCH 6 — ZCMR Orders"] = (exitosos_6, fallidos_6)
        tickets_activos = exitosos_6

    # Batch 7: verify en ZSD
    if 7 >= earliest:
        if "BATCH 6 — ZCMR Orders" in nuevos:
            tickets_para_7 = nuevos["BATCH 6 — ZCMR Orders"][0]
        else:
            tickets_para_7 = list(failed_per_batch.get(7, []))
        if tickets_para_7:
            print("  Batch 7 retry (manual): ZSD_DEL_TICKETS...")
            exitosos_7, fallidos_7, zsd_tto = verify_zsd_del_tickets(
                session1, tickets_para_7)
            nuevos["BATCH 7 — Verificación Final ZSD"] = (exitosos_7, fallidos_7)
            # Mapping override: si ZSD no devolvió order, usar la previa
            for t in tickets_para_7:
                ticket_to_order.pop(t, None)
            for t, o in zsd_tto.items():
                ticket_to_order[t] = o
            for t in fallidos_7:
                if t not in ticket_to_order and t in prev_tto:
                    ticket_to_order[t] = prev_tto[t]

            # Ground truth: si ZSD muestra pendiente, B6 retry NO fue exitoso
            if "BATCH 6 — ZCMR Orders" in nuevos and ticket_to_order:
                pendientes = set(ticket_to_order.keys()) & set(tickets_para_7)
                if pendientes:
                    old_exit6, old_fail6 = nuevos["BATCH 6 — ZCMR Orders"]
                    new_exit6 = [t for t in old_exit6 if t not in pendientes]
                    new_fail6 = sorted(set(old_fail6) | pendientes)
                    if len(new_exit6) != len(old_exit6):
                        print(f"    [Manual B6 corregido por ZSD] "
                              f"{len(old_exit6) - len(new_exit6)} ticket(s) -> fallido")
                    nuevos["BATCH 6 — ZCMR Orders"] = (new_exit6, new_fail6)

    # Batch 8: retry line-level delete sobre los que ZSD sigue mostrando
    if 8 >= earliest:
        if "BATCH 7 — Verificación Final ZSD" in nuevos:
            tickets_para_8 = list(ticket_to_order.keys())
        else:
            prev_b8 = failed_per_batch.get(8, [])
            tickets_para_8 = [t for t in prev_b8 if t in ticket_to_order]
        if tickets_para_8:
            print(f"  Batch 8 retry (manual): {len(tickets_para_8)} líneas vía VA02...")
            cancel_exitosos, cancel_fallidos = [], []
            for ticket in tickets_para_8:
                order = ticket_to_order.get(ticket)
                if not order:
                    cancel_fallidos.append(ticket)
                    continue
                try:
                    cancel_order_by_ticket(session3, order, ticket)
                    cancel_exitosos.append(ticket)
                    cancel_failures.pop(ticket, None)
                except Exception as e:
                    print(f"    [Manual B8] order {order} (ticket {ticket}): {e}",
                          file=sys.stderr)
                    cancel_fallidos.append(ticket)
                    cancel_failures[ticket] = order
            nuevos["BATCH 8 — Order Cancellation"] = (cancel_exitosos, cancel_fallidos)

    merged = _merge_retry_into_prev(prev_resultados, nuevos)
    return (merged, {}, zcmr_failures, ticket_to_order, cancel_failures, [], [],
            dict(prev_chunk.get("order_tracking", {})))


def _format_duration(seconds: float) -> str:
    """Formatea segundos como 'Xh Ym Zs' omitiendo componentes en 0."""
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    s = seconds - h * 3600 - m * 60
    parts = []
    if h:
        parts.append(f"{h}h")
    if m or h:
        parts.append(f"{m}m")
    parts.append(f"{s:.1f}s")
    return " ".join(parts)


def _print_pipeline_duration(seconds: float, label: str = "") -> None:
    """Imprime el tiempo total del pipeline en formato legible."""
    suffix = f" ({label})" if label else ""
    print(f"\n  Tiempo total del pipeline{suffix}: "
          f"{_format_duration(seconds)}  [{seconds:.0f}s]")


def _assert_multi_session(session1, session2, session3) -> bool:
    """
    Verifica que las 3 sesiones SAP sean objetos COM DISTINTOS.

    Si session2 o session3 cayeron al fallback de sap_login (mismo objeto que
    session1), el pipeline pierde el paralelismo entre sesiones — todas las
    operaciones serializan en una única sesión visual con riesgo de race
    conditions entre lecturas y operaciones destructivas.

    Retorna True si las 3 son distintas; False y muestra mensaje destacado
    en stderr si alguna apunta al mismo objeto que otra.
    """
    s2_distinta = session2 is not None and session2 is not session1
    s3_distinta = (session3 is not None
                   and session3 is not session1
                   and session3 is not session2)

    if s2_distinta and s3_distinta:
        print(f"  [Pipeline] Multi-sesión OK — paralelismo activo:")
        print(f"    session1 (lectura/verifies): {hex(id(session1))}")
        print(f"    session2 (operaciones):       {hex(id(session2))}")
        print(f"    session3 (órdenes):           {hex(id(session3))}")
        return True

    bar = "=" * 60
    print(f"\n{bar}", file=sys.stderr)
    print("  ERROR: El pipeline requiere 3 sesiones SAP DISTINTAS.", file=sys.stderr)
    print(f"{bar}", file=sys.stderr)
    print("", file=sys.stderr)
    print("  Estado detectado:", file=sys.stderr)
    print(f"    session1 != session2: {s2_distinta}", file=sys.stderr)
    print(f"    session1 != session3 (y != session2): {s3_distinta}", file=sys.stderr)
    print("", file=sys.stderr)
    print("  Causa: SAP rechazó abrir 2 sesiones adicionales, probablemente", file=sys.stderr)
    print("  porque el usuario ya tiene el límite máximo de sesiones abiertas.", file=sys.stderr)
    print("", file=sys.stderr)
    print("  Acción requerida:", file=sys.stderr)
    print("    1. Cierra TODAS las ventanas SAP innecesarias (deja solo SAP Logon).", file=sys.stderr)
    print("    2. Vuelve a correr el comando.", file=sys.stderr)
    print("", file=sys.stderr)
    print("  El pipeline se aborta para evitar correr en modo degradado lento", file=sys.stderr)
    print("  con riesgo de race conditions entre sesiones compartidas.", file=sys.stderr)
    print(f"{bar}\n", file=sys.stderr)
    return False


def _emit_consolidated_report(state: dict,
                              tickets_critical_error: list = None) -> None:
    """Consolida el state acumulado y emite reporte consola + xlsx.

    tickets_critical_error: tickets de chunks que fallaron con excepción no
    controlada y NO se guardaron en state. Aparecen como "no procesados" en
    el reporte para evitar falsos positivos del tipo "1/1 exitosos" cuando
    en realidad el chunk crasheó.
    """
    tickets_critical_error = tickets_critical_error or []
    print(f"\n{'=' * 60}")
    print("  REPORTE CONSOLIDADO — TODOS LOS CHUNKS")
    print(f"{'=' * 60}")
    (resultados, vl06f, zcmr_failures, ticket_to_order,
     cancel_failures, tickets_no_encontrados,
     tickets_con_factura, order_tracking) = consolidate_state(state)
    print_report(resultados, state["total_tickets"], tickets_no_encontrados,
                 tickets_con_factura, tickets_critical_error)
    write_report_xlsx(
        resultados, state["total_tickets"], vl06f,
        zcmr_failures, ticket_to_order, cancel_failures,
        tickets_no_encontrados, tickets_con_factura,
        tickets_critical_error, order_tracking,
    )


def _run_retry(batches_to_run: set) -> None:
    """
    Flujo --retry-failed: NO lee Excel ni inicializa state nuevo. Carga el state
    previo, identifica chunks con fallidos y reintenta solo esos batches sobre
    esos tickets reusando vl06f cacheado. Emite reporte consolidado al final.
    """
    pipeline_start = time.time()

    try:
        input_file = get_billing_file()
    except Exception as e:
        print(f"Error localizando archivo Excel para checkpoint: {e}", file=sys.stderr)
        return

    # Probar ambos modes para detectar el del state existente
    state = load_state(input_file, expected_mode="normal")
    run_mode = "normal"
    if state is None:
        state = load_state(input_file, expected_mode="manual")
        run_mode = "manual"
    if state is None:
        print("Error: --retry-failed requiere un state previo válido para este "
              "Excel. No se encontró.", file=sys.stderr)
        return

    print(f"  [Retry] State detectado en modo {run_mode!r} con "
          f"{len(state['chunks'])}/{state['total_chunks']} chunks completados.")

    chunks_con_fallidos = [
        int(k) for k in state["chunks"]
        if has_pending_work(state["chunks"][k], batches_to_run)
    ]
    if not chunks_con_fallidos:
        print("No hay tickets fallidos ni tickets no encontrados en el state "
              "para los batches seleccionados. Nada que reintentar.")
        return
    print(f"  [Retry] {len(chunks_con_fallidos)} chunk(s) con trabajo pendiente: "
          f"{sorted(chunks_con_fallidos)}")

    # --- Login SAP (una sola vez para todos los chunks del retry) ---
    sap = SapAutomation(CREDENTIALS_FILE)
    sap.run()
    session1 = sap.session
    session2 = sap.session2
    session3 = sap.session3
    if not session1:
        print("No se pudo establecer sesión con SAP.", file=sys.stderr)
        return

    # Fail-fast si SAP no permitió abrir 3 sesiones distintas. Sin esto el
    # pipeline correría serializando todo en session1 sin que el usuario lo note.
    if not _assert_multi_session(session1, session2, session3):
        return

    n_chunks = state["total_chunks"]
    linea = "=" * 60
    for chunk_idx in sorted(chunks_con_fallidos):
        prev_chunk = state["chunks"][str(chunk_idx)]
        chunk_tickets = prev_chunk["tickets"]
        print(f"\n{linea}")
        print(f"  RETRY CHUNK {chunk_idx + 1}/{n_chunks} — "
              f"{len(chunk_tickets)} tickets originales")
        print(f"{linea}")

        t0 = time.time()
        try:
            if run_mode == "manual":
                result = process_manual_chunk_retry(prev_chunk, session1, session3)
            else:
                result = process_chunk_retry(
                    prev_chunk, session1, session2, session3, batches_to_run,
                )
        except Exception as e:
            print(f"  [RETRY {chunk_idx + 1}] Error crítico: {e}", file=sys.stderr)
            print(f"  [RETRY {chunk_idx + 1}] NO se actualiza state — "
                  f"se podrá reintentar en otra corrida.", file=sys.stderr)
            continue

        duration = time.time() - t0

        try:
            save_chunk_result(state, chunk_idx, chunk_tickets, result, duration,
                              input_file)
            print(f"\n  [Checkpoint] Chunk {chunk_idx + 1}/{n_chunks} actualizado "
                  f"({duration:.0f}s).")
        except Exception as e:
            print(f"  [Checkpoint] ERROR guardando state: {e}", file=sys.stderr)

    _emit_consolidated_report(state)

    _print_pipeline_duration(time.time() - pipeline_start, label="--retry-failed")


def main():
    setup_logging()
    args = parse_args()

    if args.retry_failed and args.fresh:
        print("Error: --retry-failed y --fresh son mutuamente excluyentes.",
              file=sys.stderr)
        return
    if args.retry_failed and args.report_only:
        print("Error: --retry-failed y --report-only son mutuamente excluyentes.",
              file=sys.stderr)
        return

    try:
        batches_to_run = parse_batch_spec(args.batches)
    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        return
    chunk_size = max(1, args.chunk_size)
    print(f"Batches a ejecutar: {sorted(batches_to_run)}")
    print(f"Tamaño de chunk:    {chunk_size}")

    if args.retry_failed:
        _run_retry(batches_to_run)
        return

    pipeline_start = time.time()

    # --- Lectura del Excel ---
    # Modo --manual-only: lee SOLO la hoja 'Manual Orders' con pares (ticket, order)
    # Modo normal: lee las columnas ZCMR (C) y Ticket Number (J)
    from excel_reader import read_zcmr, read_manual_orders

    if args.manual_only:
        try:
            manual_pairs = read_manual_orders()
        except Exception as e:
            print(f"Error leyendo la hoja Manual Orders: {e}", file=sys.stderr)
            return

        if not manual_pairs:
            print(
                "No se encontraron pares (ticket, order) en la hoja "
                "'Manual Orders'. Verifica que la hoja exista y tenga "
                "columnas A=Ticket y B=Order desde fila 2.",
                file=sys.stderr,
            )
            return

        # En modo manual, los "tickets" son las keys de manual_pairs
        tickets = list(manual_pairs.keys())
        print(f"Pares manuales:     {len(manual_pairs)} (modo recovery, Batches 6-8)")
    else:
        try:
            tickets = read_zcmr()
        except Exception as e:
            print(f"Error leyendo el archivo Excel: {e}", file=sys.stderr)
            return

        if not tickets:
            print("No se encontraron tickets en las columnas ZCMR / Ticket Number.",
                  file=sys.stderr)
            return

        manual_pairs = {}
        print(f"Tickets cargados:   {len(tickets)}")

    # --- Checkpointing: cargar o inicializar state ---
    try:
        input_file = get_billing_file()
    except Exception as e:
        print(f"Error localizando archivo Excel para checkpoint: {e}", file=sys.stderr)
        return

    if args.fresh:
        clear_state(input_file)

    # Mode separation: el state file distingue 'normal' vs 'manual' para que
    # cambiar entre modos invalide el state automáticamente (los tickets son
    # distintos: ZCMR/Ticket Number vs pares de Manual Orders).
    run_mode = "manual" if args.manual_only else "normal"

    state = load_state(input_file, expected_mode=run_mode)
    if state is None:
        state = init_state(input_file, tickets, chunk_size, batches_to_run, mode=run_mode)
        print(f"  [Checkpoint] Nuevo state inicializado (mode={run_mode!r}) en Data-bases/Estado/")
    else:
        n_done = len(state["chunks"])
        print(f"  [Checkpoint] State previo encontrado: {n_done}/{state['total_chunks']} "
              f"chunks ya completados. Reanudando.")
        # Si chunk_size cambió respecto al state, advertir
        if state.get("chunk_size") != chunk_size:
            print(f"  [Checkpoint] WARNING: chunk_size cambió "
                  f"({state.get('chunk_size')} -> {chunk_size}). Re-usando state previo "
                  f"con chunk_size original.", file=sys.stderr)
            chunk_size = state["chunk_size"]

    # --- --report-only: regenera reporte sin tocar SAP ---
    if args.report_only:
        if not state["chunks"]:
            print("No hay chunks completados en el state — nada que reportar.")
            return
        _emit_consolidated_report(state)
        _print_pipeline_duration(time.time() - pipeline_start, label="--report-only")
        return

    # --- Login a SAP (una sola vez, sesiones se reusan entre chunks) ---
    sap = SapAutomation(CREDENTIALS_FILE)
    sap.run()
    session1 = sap.session   # VL06F — lectura, verificaciones, BOL
    session2 = sap.session2  # Operaciones — VF11, VI05, VT02N, VL09
    session3 = sap.session3  # Órdenes — ZCMR, VA02, ME22N

    if not session1:
        print("No se pudo establecer sesión con SAP.", file=sys.stderr)
        return

    # Fail-fast si SAP no permitió abrir 3 sesiones distintas. Sin esto el
    # pipeline correría serializando todo en session1 sin que el usuario lo note.
    if not _assert_multi_session(session1, session2, session3):
        return

    # En modo --manual-only se ignoran los flags --batches y needs_vl06f.
    # El flujo manual SIEMPRE corre Batches 6-7-8 (definido en process_manual_chunk).
    needs_vl06f = bool(batches_to_run & {1, 2, 3, 4, 5}) and not args.manual_only

    # --- Partir tickets en chunks y procesar cada uno por separado ---
    n_chunks = state["total_chunks"]
    mode_label = "modo MANUAL (Batches 6-8)" if args.manual_only else "modo normal"
    print(f"Total chunks:       {n_chunks} de hasta {chunk_size} tickets c/u — {mode_label}\n")

    # Tickets de chunks que fallaron con error crítico (excepción no controlada).
    # No se guardan en state, pero deben aparecer en el reporte como "no procesados
    # — requieren re-corrida del pipeline". Sin esto, el reporte miente diciendo
    # "1/1 exitosos" cuando el chunk realmente falló.
    tickets_critical_error: list = []

    for chunk_idx in range(n_chunks):
        # Skip si ya está completado (resume)
        if is_chunk_completed(state, chunk_idx):
            chunk_done = state["chunks"][str(chunk_idx)]
            print(f"\n  CHUNK {chunk_idx + 1}/{n_chunks} ya completado el "
                  f"{chunk_done['completed_at']} — saltando.")
            continue

        start_idx = chunk_idx * chunk_size
        end_idx = min(start_idx + chunk_size, len(tickets))
        chunk = tickets[start_idx:end_idx]

        linea = "=" * 60
        print(f"\n{linea}")
        print(f"  CHUNK {chunk_idx + 1}/{n_chunks}  —  {len(chunk)} tickets "
              f"({chunk[0]} ... {chunk[-1]})")
        print(f"{linea}")

        t0 = time.time()
        try:
            if args.manual_only:
                chunk_pairs = {t: manual_pairs[t] for t in chunk}
                result = process_manual_chunk(chunk_pairs, session1, session3)
            else:
                result = process_chunk(
                    chunk, session1, session2, session3, batches_to_run, needs_vl06f,
                )
        except Exception as e:
            print(f"  [CHUNK {chunk_idx + 1}] Error crítico: {e}", file=sys.stderr)
            print(f"  [CHUNK {chunk_idx + 1}] NO se guarda en state — se reintentará "
                  f"al re-correr el pipeline.", file=sys.stderr)
            tickets_critical_error.extend(chunk)
            continue

        # Auto-retry inline: si quedaron fallidos en este chunk, reintentar
        # UNA vez antes de persistir. Ataca fallos transitorios (locks SAP,
        # timeouts, layouts mal aplicados) sin requerir --retry-failed manual.
        fake_prev = _result_to_chunk_dict(result, chunk)
        if get_failed_tickets_per_batch(fake_prev, batches_to_run):
            print(f"\n  [Auto-retry] Fallidos detectados en chunk {chunk_idx + 1}. "
                  f"Reintentando una vez antes de persistir...")
            try:
                if args.manual_only:
                    result = process_manual_chunk_retry(fake_prev, session1, session3)
                else:
                    result = process_chunk_retry(
                        fake_prev, session1, session2, session3, batches_to_run,
                    )
            except Exception as e:
                print(f"  [Auto-retry] Error: {e} — se conserva el resultado original.",
                      file=sys.stderr)

        duration = time.time() - t0

        # Persistir progreso del chunk en disco (atómico)
        try:
            save_chunk_result(state, chunk_idx, chunk, result, duration, input_file)
            print(f"\n  [Checkpoint] Chunk {chunk_idx + 1}/{n_chunks} guardado "
                  f"({duration:.0f}s).")
        except Exception as e:
            print(f"  [Checkpoint] ERROR guardando state: {e}", file=sys.stderr)

    # --- Reporte consolidado final desde el state acumulado ---
    _emit_consolidated_report(state, tickets_critical_error)

    _print_pipeline_duration(time.time() - pipeline_start)


if __name__ == "__main__":
    main()
