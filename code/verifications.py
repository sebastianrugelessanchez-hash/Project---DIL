import sys
import time

from sap_utils import _navigate_to, _wait_ready, _POPUP_TABLE, _enter_multi_values
from sap_vl06f import read_vl06f_data, _vl06f_delivery_filter


def _get_ticket_data(session, ticket: str) -> dict:
    """
    Navega a VL06F filtrando por un único ticket y retorna los valores clave del grid.
    Retorna {} si el ticket no se encuentra o hay un error.
    """
    _navigate_to(session, "VL06F")
    _wait_ready(session)

    for field in ("ctxtIT_WADAT-LOW", "ctxtIT_WADAT-HIGH"):
        try:
            session.findById(f"wnd[0]/usr/{field}").text = ""
        except Exception:
            pass

    try:
        _vl06f_delivery_filter(session, [ticket])
    except Exception as e:
        print(f"  [Verif] Error abriendo popup VL06F para {ticket}: {e}", file=sys.stderr)
        return {}

    session.findById("wnd[0]/tbar[1]/btn[8]").press()  # F8
    _wait_ready(session)

    try:
        grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        row_count = grid.RowCount
        for row in range(max(0, row_count - 1)):
            vbeln = grid.GetCellValue(row, "VBELN")
            if vbeln == ticket:
                return {
                    "ZZVBELN": grid.GetCellValue(row, "ZZVBELN") or "",
                    "ZZFKNUM": grid.GetCellValue(row, "ZZFKNUM") or "",
                    "ZZTKNUM": grid.GetCellValue(row, "ZZTKNUM") or "",
                    "WBSTK":   grid.GetCellValue(row, "WBSTK")   or "",
                }
    except Exception as e:
        print(f"  [Verif] Error leyendo VL06F para {ticket}: {e}", file=sys.stderr)

    return {}


def verify_billing_document(session, ticket: str) -> bool:
    """True si el Billing Document del ticket está vacío en VL06F."""
    return _get_ticket_data(session, ticket).get("ZZVBELN", "") == ""


def verify_shipment_cost(session, ticket: str) -> bool:
    """True si el Shipment Cost del ticket está vacío en VL06F."""
    return _get_ticket_data(session, ticket).get("ZZFKNUM", "") == ""


def verify_shipment_number(session, ticket: str) -> bool:
    """True si el Shipment Number del ticket está vacío en VL06F."""
    return _get_ticket_data(session, ticket).get("ZZTKNUM", "") == ""


def verify_pgi_reversed(session, ticket: str) -> bool:
    """True si el PGI fue revertido (WBSTK = 'A') en VL06F."""
    return _get_ticket_data(session, ticket).get("WBSTK", "") == "A"


# ---------------------------------------------------------------------------
# Verificaciones bulk — 1 sola consulta VL06F para todos los tickets
# ---------------------------------------------------------------------------

def _classify(tickets: list, data: dict, field: str, success_value="") -> tuple[list, list]:
    """
    Clasifica tickets en exitosos/fallidos basándose en VL06F re-lectura.

    success_value puede ser un str (un único valor de éxito) o una tupla/set
    de strings (varios valores de éxito aceptables). Útil cuando un mismo
    estado lógico puede corresponder a varios códigos SAP — p. ej. WBSTK 'A'
    (PGI reversed) y '' (nunca hubo PGI) son ambos "no requiere reverse".

    Reglas:
    - Si el ticket NO está en VL06F (data) -> EXITOSO (asumimos que la operación
      lo procesó completo y SAP lo sacó del monitor, o ya estaba limpio).
    - Si el ticket está en VL06F y data[t][field] está en success_value -> EXITOSO
    - En cualquier otro caso -> FALLIDO

    Tratar "no encontrado" como éxito evita cascadas falsas cuando VL06F
    cambia su vista después de una operación. La verificación final en
    Batch 7 (ZSD_DEL_TICKETS) catchea cualquier falso positivo.
    """
    if isinstance(success_value, str):
        success_values = {success_value}
    else:
        success_values = set(success_value)

    exitosos = []
    fallidos = []
    for t in tickets:
        if t not in data:
            exitosos.append(t)  # no en VL06F = procesado/inexistente = OK
        elif data[t].get(field, "") in success_values:
            exitosos.append(t)
        else:
            fallidos.append(t)
    return exitosos, fallidos


def verify_billing_documents_bulk(session, tickets: list) -> tuple[list, list]:
    """Una consulta VL06F. Ticket es exitoso si billing_doc=="" o no está en VL06F."""
    if not tickets:
        return [], []
    data = read_vl06f_data(session, tickets)
    return _classify(tickets, data, "billing_doc", "")


def verify_shipment_costs_bulk(session, tickets: list) -> tuple[list, list]:
    """Una consulta VL06F. Ticket es exitoso si shpt_cst=="" o no está en VL06F."""
    if not tickets:
        return [], []
    data = read_vl06f_data(session, tickets)
    return _classify(tickets, data, "shpt_cst", "")


def verify_shipment_numbers_bulk(session, tickets: list) -> tuple[list, list]:
    """Una consulta VL06F. Ticket es exitoso si shipment=="" o no está en VL06F."""
    if not tickets:
        return [], []
    data = read_vl06f_data(session, tickets)
    return _classify(tickets, data, "shipment", "")


def verify_pgi_reversed_bulk(session, tickets: list) -> tuple[list, list]:
    """
    Una consulta VL06F. Ticket es exitoso si:
      - wbstk == 'A' (PGI reversed) — lo que el Batch 4 acaba de hacer, O
      - wbstk == ''  (nunca hubo PGI — nada que reversar, semánticamente OK), O
      - no está en VL06F.
    Cualquier otro valor (ej: 'C' = PGI completo sin reverse) -> fallido.
    """
    if not tickets:
        return [], []
    data = read_vl06f_data(session, tickets)
    return _classify(tickets, data, "wbstk", ("A", ""))


def verify_bol_deleted_bulk(session, tickets: list) -> tuple[list, list]:
    """
    Verifica que las deliveries (BOLs) fueron eliminadas de VL06F.

    Semántica:
      - Ticket NO aparece en VL06F (out-of-queue) -> EXITOSO (delivery borrada)
      - Ticket aparece en VL06F                    -> FALLIDO (delivery sigue viva)

    Es la confirmación crítica de Batch 5. `delete_bol` puede ejecutarse sin
    lanzar excepción y aún así no haber borrado nada (rechazo silencioso de
    SAP, click sin efecto, popup auto-cerrado). Solo este re-read confirma
    que la operación tuvo efecto real.

    Sin este verify, fallos silenciosos se propagan a Batch 6 con tickets
    que en realidad nunca se borraron — exactamente el bug que el code
    review identificó.
    """
    if not tickets:
        return [], []
    data = read_vl06f_data(session, tickets)
    exitosos = [t for t in tickets if t not in data]
    fallidos = [t for t in tickets if t in data]
    return exitosos, fallidos


# ---------------------------------------------------------------------------
# Verificación final — ZSD_DEL_TICKETS
# ---------------------------------------------------------------------------

def verify_zsd_del_tickets(session, tickets: list) -> tuple[list, list, dict]:
    """
    ZSD_DEL_TICKETS: verificación final del pipeline.
    Retorna (exitosos, fallidos, ticket_to_order).

    - ticket_to_order: dict[ticket, order] solo para tickets con órdenes pendientes
    - Captura el par (ticket, order) de la MISMA fila del grid
    - Si tickets es vacío -> return inmediato sin tocar SAP
    - Si ZSD falla -> todos exitosos, mapping vacío
    """
    if not tickets:
        return [], [], {}

    try:
        _navigate_to(session, "ZSD_DEL_TICKETS")
        _wait_ready(session)

        session.findById("wnd[0]/usr/btn%_P_TICKET_%_APP_%-VALU_PUSH").press()
        _wait_ready(session)
        # btn[16] = Delete Entire Selection (Shift+F4) — limpiar valores previos
        try:
            session.findById("wnd[1]/tbar[0]/btn[16]").press()
            _wait_ready(session)
        except Exception:
            pass
        _enter_multi_values(session, _POPUP_TABLE, tickets)
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        _wait_ready(session)

        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        _wait_ready(session)

    except Exception as e:
        print(f"  [ZSD] Error navegando a ZSD_DEL_TICKETS: {e}", file=sys.stderr)
        print("  [ZSD] Se asumen todos como exitosos — verificar manualmente.", file=sys.stderr)
        return tickets[:], [], {}

    ticket_to_order: dict = {}
    all_orders_found: list = []  # dedup en orden de inserción
    try:
        grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        row_count = grid.RowCount

        if row_count == 0:
            return tickets[:], [], {}

        # Diagnóstico: listar TODAS las columnas disponibles + sample de fila 0
        try:
            cols_avail = list(grid.ColumnOrder)
            print(f"  [ZSD] Grid RowCount={row_count}, columnas ({len(cols_avail)}):")
            for c in cols_avail:
                try:
                    sample = grid.GetCellValue(0, c)
                    print(f"    {c!r:35s} sample row0={sample!r}")
                except Exception:
                    print(f"    {c!r:35s} (no se pudo leer)")
        except Exception as e:
            print(f"  [ZSD] no se pudieron listar columnas: {e}")

        # Sales Order de SAP: típicamente 10 dígitos empezando con "1" (ej. 1150403358).
        # Excluir valores cortos (códigos como "1067167" de 7 dígitos NO son sales orders)
        # o "0" o vacío.
        def _is_real_sales_order(val: str) -> bool:
            v = (val or "").strip()
            return v.isdigit() and len(v) >= 10 and v != "0" * len(v)

        # Candidatos para la columna del TICKET y del SALES ORDER real.
        # NOTA: en ZSD_DEL_TICKETS el ticket vive en 'TICKET_CODE' (confirmado en
        # logs de diagnóstico); 'ORDER_CODE' es un código distinto y suele venir
        # vacío, el sales order real está en 'SD_ORDER'.
        _TICKET_COLS = ("TICKET_CODE", "TICKET", "VBELN", "DELIVERY")
        _ORDER_COLS  = ("SD_ORDER", "VBELN_VA", "ORDER_NUM", "AUFNR", "ORDER", "VGBEL")

        for row in range(row_count):
            row_ticket = ""
            for col in _TICKET_COLS:
                try:
                    val = (grid.GetCellValue(row, col) or "").strip()
                    if val in tickets:
                        row_ticket = val
                        break
                except Exception:
                    continue

            row_order = ""
            for col in _ORDER_COLS:
                try:
                    val = (grid.GetCellValue(row, col) or "").strip()
                    # Validar que parezca un sales order real (>=10 dígitos)
                    if _is_real_sales_order(val):
                        row_order = val
                        break
                except Exception:
                    continue

            if row_order and row_order not in all_orders_found:
                all_orders_found.append(row_order)

            if row_ticket and row_order:
                ticket_to_order[row_ticket] = row_order

    except Exception:
        return tickets[:], [], {}

    # FALLBACK conservador nivel 1: si hay orders en el grid pero no se mapeó
    # ningún ticket (ej: layout oculta TICKET pero ORDER sí se leyó), marcar
    # TODOS los tickets como fallidos y distribuir las orders cíclicamente.
    if all_orders_found and not ticket_to_order:
        print(f"  [ZSD] WARNING: {len(all_orders_found)} orders detectadas en el grid "
              f"pero no se mapearon a tickets (¿columna TICKET oculta en el layout?).")
        print(f"  [ZSD] Orders detectadas: {all_orders_found}")
        print(f"  [ZSD] Marcando los {len(tickets)} tickets como fallidos por seguridad.")
        for i, ticket in enumerate(tickets):
            ticket_to_order[ticket] = all_orders_found[i % len(all_orders_found)]

    # FALLBACK conservador nivel 2: si el grid tiene filas (row_count > 0) pero
    # NI ticket NI order se pudieron leer (ej: nombres técnicos de columna no
    # están en _TICKET_COLS/_ORDER_COLS), asumimos que TODOS los tickets están
    # pendientes. El grid es ground truth: si tiene filas, hay trabajo por hacer.
    # Esto previene el falso positivo "12/12 exitosos pero orders siguen en SAP".
    if row_count > 0 and not ticket_to_order:
        print(f"  [ZSD] WARNING: grid tiene {row_count} filas pero no se pudo leer "
              f"ni TICKET ni ORDER. Probable mismatch de nombres técnicos de columna.")
        print(f"  [ZSD] _TICKET_COLS={_TICKET_COLS}, _ORDER_COLS={_ORDER_COLS}")
        print(f"  [ZSD] Revisar el diagnóstico de columnas (líneas 'sample row0=...') "
              f"y ajustar las tuplas si es necesario.")
        print(f"  [ZSD] Marcando los {len(tickets)} tickets como fallidos por seguridad.")
        for ticket in tickets:
            ticket_to_order[ticket] = ""  # sin order conocida, Batch 8 no podrá borrar

    if ticket_to_order:
        print(f"  [ZSD] {len(ticket_to_order)} tickets con orders pendientes.")

    fallidos = list(ticket_to_order.keys())
    exitosos = [t for t in tickets if t not in ticket_to_order]
    return exitosos, fallidos, ticket_to_order
