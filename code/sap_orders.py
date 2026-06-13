import sys
import time

from sap_utils import (
    _navigate_to, _go_back, _wait_ready, _enter_multi_values, _POPUP_TABLE,
    _normalize_ticket,
)


_ZCMR_SUB_GRID = "wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell"

# Candidatos de nombres técnicos de columna del SUB-GRID de ZCMR (el detalle
# ticket<->orden<->delivery). El grid principal es solo resumen (PLANT, TICKET_QTY,
# etc.) y NO tiene estas columnas. El layout puede variar, así que se resuelve
# el primer nombre que exista vía _resolve_col en lugar de asumir uno fijo.
_ZCMR_ORDER_COLS    = ("SD_ORDER", "ORDER", "VBELN_VA", "AUFNR", "VGBEL")
_ZCMR_DELIVERY_COLS = ("DELIVERY", "VBELN", "VBELN_VL", "DEL_NUM")
_ZCMR_TICKET_COLS   = ("TICKET_CODE", "TICKET", "TICKET_NUM", "P_TICKET")


def _resolve_col(grid, candidates):
    """Retorna el primer nombre de `candidates` presente en grid.ColumnOrder, o None."""
    try:
        available = set(grid.ColumnOrder)
    except Exception:
        return candidates[0] if candidates else None
    for c in candidates:
        if c in available:
            return c
    return None


def _dump_subgrid_columns(sub) -> None:
    """
    Volcado one-shot de las columnas reales del sub-grid de ZCMR (diagnóstico).
    Queda en el log para conocer los nombres técnicos sin correr un script aparte.
    """
    try:
        cols = list(sub.ColumnOrder)
        print(f"  [ZCMR] Sub-grid columnas reales ({len(cols)}): {cols}")
        for c in cols:
            try:
                print(f"    {c!r:30s} sample row0={sub.GetCellValue(0, c)!r}")
            except Exception:
                pass
    except Exception as e:
        print(f"  [ZCMR] no se pudo volcar columnas del sub-grid: {e}", file=sys.stderr)

_VA02_TABLE = (
    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01"
    "/ssubSUBSCREEN_BODY:SAPMV45A:4400"
    "/subSUBSCREEN_TC:SAPMV45A:4900"
    "/tblSAPMV45ATCTRL_U_ERF_AUFTRAG"
)

_ME22N_TABLE = (
    "wnd[0]/usr/subSUB0:SAPLMEGUI:0019"
    "/subSUB2:SAPLMEVIEWS:1100"
    "/subSUB2:SAPLMEVIEWS:1200"
    "/subSUB1:SAPLMEGUI:1211"
    "/tblSAPLMEGUITC_1211"
)

_ME22N_DELETE_BTN = (
    "wnd[0]/usr/subSUB0:SAPLMEGUI:0019"
    "/subSUB2:SAPLMEVIEWS:1100"
    "/subSUB2:SAPLMEVIEWS:1200"
    "/subSUB1:SAPLMEGUI:1211/btnDELETE"
)

# Flag one-shot para volcar las columnas de la tabla ME22N una sola vez por run.
_ME22N_COLS_DUMPED = False


def delete_orders_from_zcmr(session, tickets: list) -> dict:
    """
    ZCMR -> VA02 / ME22N: Lee los orders de todos los tickets en ZCMR y los elimina.
    Intercompany usa VA02; Intracompany (order[:2] == '47') usa ME22N.

    Retorna dict[ticket, list[orders_que_fallaron]] — solo tickets con fallos.

    NOTA: Esta función está deprecada en favor de delete_orders_from_zsd, que
    es más robusta porque ZSD_DEL_TICKETS sigue mostrando el mapping
    ticket->order incluso después de que Batch 5 (BOL Deletion) haya quitado
    el ticket de ZCMR. Mantener por backward compatibility — el pipeline
    principal ahora usa delete_orders_from_zsd.
    """
    orders = _read_zcmr_orders(session, tickets)
    print(f"  [ZCMR] {len(orders)} orders encontrados para eliminar.")

    failures: dict = {}
    for item in orders:
        order = item["order"]
        delivery = item["delivery"]
        ticket = item.get("ticket", "")
        try:
            if item["is_intracompany"]:
                _delete_intracompany_order_me22n(session, order)
            else:
                _delete_intercompany_order_va02(session, order, delivery)
        except Exception as e:
            print(f"  [ZCMR] Error en order {order}: {e}", file=sys.stderr)
            if ticket:
                failures.setdefault(ticket, []).append(order)
            else:
                failures.setdefault("(sin ticket)", []).append(order)

    return failures


def delete_orders_from_zsd(session_zsd, session_orders, tickets: list) -> tuple:
    """
    ZSD_DEL_TICKETS -> VA02 / ME22N: variante robusta de Batch 6.

    Usa ZSD_DEL_TICKETS para obtener el mapping ticket->order (en vez de ZCMR,
    que filtra tickets cuyo BOL ya fue eliminado por Batch 5). ZSD_DEL_TICKETS
    sigue mostrando todos los tickets independientemente del estado del BOL,
    indicando con order vacía si ya fue eliminada o con número si sigue pendiente.

    Para cada ticket con order pendiente:
        - Si order[:2] == '47'  -> ME22N (intracompany purchase order)
        - Else                  -> VA02 line-delete (intercompany sales order)

    Args:
        session_zsd: sesión para leer ZSD_DEL_TICKETS (típicamente session1).
        session_orders: sesión para borrar en VA02/ME22N (típicamente session3).
        tickets: lista de ticket numbers a procesar.

    Retorna tupla (failures, order_tracking):
        failures: dict[ticket, list[orders_que_fallaron]] — solo tickets con fallos.
        order_tracking: dict[ticket, dict] con la trayectoria de cada ticket
            que tenía orden pendiente en ZSD. Campos:
                order, intracompany, transaction, detected_in_zsd,
                batch_6, batch_7_verify, batch_8_cancel, final_status, error_msg.
            Otros batches (7, 8) actualizan este dict después.
    """
    if not tickets:
        return {}, {}

    # Import local para evitar circularidad con verifications.py
    from verifications import verify_zsd_del_tickets

    _, _, ticket_to_order = verify_zsd_del_tickets(session_zsd, tickets)

    order_tracking: dict = {}

    if not ticket_to_order:
        print(f"  [ZSD] Ningún ticket tiene order pendiente en ZSD_DEL_TICKETS.")
        return {}, order_tracking

    print(f"  [ZSD] {len(ticket_to_order)} ticket(s) con order pendiente. "
          f"Borrando vía VA02/ME22N...")

    failures: dict = {}
    for ticket, order in ticket_to_order.items():
        intra = bool(order and str(order).startswith("47"))
        order_tracking[ticket] = {
            "order": order or "",
            "intracompany": intra,
            "transaction": "ME22N" if intra else "VA02",
            "detected_in_zsd": True,
            "batch_6": "skipped",
            "batch_7_verify": "",
            "batch_8_cancel": "",
            "final_status": "unknown",
            "error_msg": "",
        }

        if not order:
            print(f"  [ZSD] ticket {ticket}: pending sin order mapeada (layout?). "
                  f"NO se puede borrar.", file=sys.stderr)
            failures.setdefault(ticket, []).append("(sin order)")
            order_tracking[ticket]["batch_6"] = "failed"
            order_tracking[ticket]["error_msg"] = "(sin order mapeada en ZSD)"
            order_tracking[ticket]["final_status"] = "failed"
            continue
        try:
            # cancel_order_by_ticket dispatch según order[:2] == '47' a ME22N
            # o VA02 line-delete (usa el ticket como delivery para line-match).
            cancel_order_by_ticket(session_orders, order, ticket)
            order_tracking[ticket]["batch_6"] = "ok"
        except Exception as e:
            print(f"  [ZSD] Error en order {order} (ticket {ticket}): {e}",
                  file=sys.stderr)
            failures.setdefault(ticket, []).append(order)
            order_tracking[ticket]["batch_6"] = "failed"
            order_tracking[ticket]["error_msg"] = str(e)[:200]

    return failures, order_tracking


def _read_zcmr_orders(session, tickets: list, plant: str = "*") -> list:
    """
    Navega a ZCMR, ejecuta con todos los tickets y lee los orders de cada sub-grid.
    Retorna lista de {order, delivery, is_intracompany}.
    Lee todo en memoria antes de navegar a VA02/ME22N (single-session).

    Args:
        session: sesión SAP (típicamente session3).
        tickets: lista de ticket numbers a buscar en ZCMR.
        plant: valor del campo Plant (P_PLANT-LOW). SAP requiere este campo
            obligatoriamente — '*' no funciona como wildcard. Default '*' por
            compatibilidad pero el caller DEBE pasar un plant válido (ej:
            '8710') para que ZCMR devuelva resultados.

    Si tickets es vacío -> return inmediato sin navegar (F8 con filtro vacío
    devolvería todo ZCMR y SAP se congelaría).
    """
    if not tickets:
        return []

    _navigate_to(session, "ZCMR")
    _wait_ready(session)

    session.findById("wnd[0]/usr/ctxtP_PLANT-LOW").Text = plant
    session.findById("wnd[0]/usr/ctxtP_PLANT-LOW").SetFocus()

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

    session.findById("wnd[0]/tbar[1]/btn[8]").press()  # F8
    _wait_ready(session)

    orders = []
    subgrid_dumped = False

    try:
        grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        row_count = max(0, grid.RowCount - 2)

        for i in range(row_count):
            try:
                grid.currentCellRow = i
                grid.selectedRows = i

                fecha = grid.GetCellValue(i, "TICKET_DATE") or ""
                if not fecha:
                    continue

                # Capturar ticket de la fila principal (columnas candidatas)
                ticket_row = ""
                for col in ("TICKET", "P_TICKET", "ZTICKET", "TICKET_NUM"):
                    try:
                        v = (grid.GetCellValue(i, col) or "").strip()
                        if v:
                            ticket_row = v
                            break
                    except Exception:
                        continue

                current_deletion_raw = grid.GetCellValue(i, "TICKET_LOADED") or "0"
                current_deletion = int(current_deletion_raw.replace(".", "")) if current_deletion_raw.strip() else 0
                if current_deletion == 0:
                    continue

                grid.currentCellRow = i
                grid.selectedRows = i
                grid.doubleClickCurrentCell()
                _wait_ready(session)

                # Usar el RowCount real del sub-grid como límite,
                # NO confiar ciegamente en current_deletion (puede venir mal parseado
                # del locale o ser un total acumulado en vez de per-ticket).
                try:
                    sub = session.findById(_ZCMR_SUB_GRID)
                    sub_rows = sub.RowCount
                except Exception as e:
                    print(f"  [ZCMR] Error accediendo sub-grid en fila {i}: {e}", file=sys.stderr)
                    continue

                # Diagnóstico one-shot: volcar las columnas reales del sub-grid
                # la primera vez que lo abrimos. Así el log revela los nombres
                # técnicos sin tener que correr diagnose_grids.py por separado.
                if not subgrid_dumped:
                    _dump_subgrid_columns(sub)
                    subgrid_dumped = True

                # Resolver nombres de columna reales (tolerante a variaciones de
                # layout): probamos varios candidatos en vez de uno fijo.
                order_col = _resolve_col(sub, _ZCMR_ORDER_COLS)
                delivery_col = _resolve_col(sub, _ZCMR_DELIVERY_COLS)
                ticket_col = _resolve_col(sub, _ZCMR_TICKET_COLS)
                if order_col is None:
                    print(f"  [ZCMR] WARNING fila {i}: no se encontró columna de ORDER en "
                          f"el sub-grid. Candidatos={_ZCMR_ORDER_COLS}, "
                          f"reales={list(sub.ColumnOrder)}", file=sys.stderr)
                    try:
                        session.findById("wnd[0]/tbar[0]/btn[3]").press()  # colapsar
                        _wait_ready(session)
                    except Exception:
                        pass
                    continue

                iter_count = min(current_deletion, sub_rows)
                consecutive_errors = 0
                # Dedup por (order, delivery) — antes era solo por `order`
                # consecutivo y descartaba tickets distintos con misma order.
                seen_pairs = set()

                for j in range(iter_count):
                    try:
                        try:
                            sub.setCurrentCell(j, order_col)
                        except Exception:
                            pass
                        sub.selectedRows = j

                        delivery = (sub.GetCellValue(j, delivery_col) or "") if delivery_col else ""
                        order = sub.GetCellValue(j, order_col) or ""
                        ticket_cell = (sub.GetCellValue(j, ticket_col) or "").strip() if ticket_col else ""
                        # Preferir el ticket leído del propio sub-grid; si no hay
                        # columna, caer al capturado del grid principal.
                        ticket_final = ticket_cell or ticket_row

                        if not delivery and not order:
                            consecutive_errors = 0
                            continue

                        pair = (order, delivery)
                        if order and pair not in seen_pairs:
                            seen_pairs.add(pair)
                            orders.append({
                                "ticket": ticket_final,
                                "order": order,
                                "delivery": delivery,
                                "is_intracompany": order[:2] == "47",
                            })
                        consecutive_errors = 0

                    except Exception as e:
                        consecutive_errors += 1
                        if consecutive_errors >= 3:
                            print(
                                f"  [ZCMR] Bail out: 3 errores consecutivos en sub-grid "
                                f"de fila {i} (j={j}). Probablemente sub-grid truncado.",
                                file=sys.stderr,
                            )
                            break
                        print(f"  [ZCMR] Error sub-grid fila {j}: {e}", file=sys.stderr)
                        continue

                session.findById("wnd[0]/tbar[0]/btn[3]").press()  # Colapsar sub-grid
                _wait_ready(session)

            except Exception as e:
                print(f"  [ZCMR] Error fila {i}: {e}", file=sys.stderr)
                continue

    except Exception as e:
        print(f"  [ZCMR] Error leyendo grid principal: {e}", file=sys.stderr)

    return orders


_ERROR_KEYWORDS = (
    "cannot be deleted",
    "no se puede borrar",
    "no se puede eliminar",
    "subsequent document",
    "documento subsiguiente",
    "documento posterior",
    "error",
)


# Mensajes de SAP que indican que la orden YA NO existe / ya fue borrada. En ese
# caso el objetivo (que la orden no exista) ya está cumplido -> se cuenta éxito,
# no error. Da idempotencia: re-correr el pipeline no falla en orders ya borradas.
_ORDER_GONE_KEYWORDS = (
    "does not exist",
    "no existe",
    "is not in the database",
    "no está en la base de datos",
    "has been archived",
    "ha sido archivad",          # archivado/archivada
    "no longer exists",
    "ya no existe",
)


def _read_sbar(session) -> tuple:
    """Retorna (texto, tipo) del status bar de wnd[0]; ('','') si no se puede."""
    try:
        bar = session.findById("wnd[0]/sbar")
        return (bar.Text or "").strip(), (bar.MessageType or "").upper()
    except Exception:
        return "", ""


def _order_already_gone(session, order: str, popup_msg: str = "") -> bool:
    """
    True si SAP indica que la orden ya no existe (popup inicial o status bar).
    Imprime un mensaje informativo cuando lo detecta.
    """
    sbar_text, _ = _read_sbar(session)
    combined = f"{popup_msg} {sbar_text}".strip().lower()
    if any(kw in combined for kw in _ORDER_GONE_KEYWORDS):
        print(f"  [VA02] orden {order}: ya no existe en SAP "
              f"({(sbar_text or popup_msg)!r}). Se considera borrada (OK).")
        return True
    return False


def _dump_va02_menu(session, order: str) -> None:
    """
    Diagnóstico one-shot: vuelca al log los submenús de wnd[0]/mbar/menu[0] (id
    y texto) cuando no se pudo seleccionar el menú Delete. Sirve para confirmar
    dónde está realmente "Delete" si el popup no era la causa.
    """
    try:
        menu0 = session.findById("wnd[0]/mbar/menu[0]")
        print(f"  [VA02] orden {order}: volcado de menú '{menu0.Text}' "
              f"({menu0.Children.Count} items):", file=sys.stderr)
        for i in range(menu0.Children.Count):
            try:
                item = menu0.Children(i)
                print(f"    menu[0]/menu[{i}] = {item.Text!r}", file=sys.stderr)
            except Exception:
                continue
    except Exception as e:
        print(f"  [VA02] orden {order}: no se pudo volcar el menú: {e}",
              file=sys.stderr)


_SAVE_POPUP_KEYWORDS = (
    "incomplete", "incompleto",
    "save the document", "guardar el documento",
    "save anyway", "guardar de todas",
)

# Popups informativos que SAP muestra al borrar líneas en VA02. Todos requieren
# presionar Enter/Continue para que el save proceda. Si los cerramos con la X,
# SAP cancela el delete silenciosamente.
_INFO_POPUP_KEYWORDS = (
    "delivery group",                # "Delivery group N consists of only one item"
    "consists of only",
    "subsequent function",           # popups de funciones subsecuentes
    "information",                   # popups con título "Information"
    "será desfijado", "will be unfixed",
    "será removido", "will be removed",
)


def _press_continue_on_popup(session, wnd_id: str) -> bool:
    """
    Presiona Enter / Continue en un popup informativo de SAP. Prueba varios
    IDs comunes hasta encontrar uno que funcione.
    """
    # Estrategia: sendVKey(0) en el popup envía Enter — funciona en la mayoría
    # de popups informativos sin tener que encontrar el botón exacto.
    try:
        session.findById(wnd_id).sendVKey(0)
        _wait_ready(session)
        return True
    except Exception:
        pass
    # Fallback: probar botones explícitos de Continue
    for btn_path in ("tbar[0]/btn[0]", "usr/btnSPOP-OPTION1"):
        try:
            session.findById(f"{wnd_id}/{btn_path}").press()
            _wait_ready(session)
            return True
        except Exception:
            continue
    return False


def _handle_va02_post_save_popups(session, max_wait: float = 4.0,
                                  max_popups: int = 8) -> int:
    """
    Maneja los popups que VA02 muestra después de borrar líneas y presionar Save.
    En orden de aparición:
      1. Popup informativo "Delivery group N consists of only one item" — Enter
      2. Posibles otros popups informativos — Enter
      3. Popup "Save Incomplete Document" — Save (btnSPOP-OPTION1/VAROPTION1)

    El popup puede aparecer DESPUÉS de que session.Busy vuelva a False, por eso
    hace polling con timeout. También puede haber MÚLTIPLES popups en cascada.

    Si se cerrara cualquiera de estos popups con btn[0] (X = Cancel), SAP
    descartaría el delete silenciosamente — exactamente el bug que aparecía.

    COTAS ANTI-BUCLE-INFINITO (el deadline se resetea con cada popup, así que sin
    estas cotas un popup que NO se cierra al presionarlo colgaba el pipeline):
      - max_popups: tope de popups manejados.
      - hard_deadline: tope absoluto que NO se resetea.
      - detección de "mismo popup repetido": si tras manejarlo reaparece idéntico
        varias veces, se asume que no se cierra y se aborta.

    Retorna el número de popups manejados (0 si no había ninguno).
    """
    deadline = time.time() + max_wait
    hard_deadline = time.time() + max_wait * (max_popups + 2)
    handled = 0
    other_popup_seen = ""
    last_key = None       # (wnd_id, popup_text) del último popup visto
    repeat = 0            # cuántas veces seguidas reapareció el mismo popup

    while time.time() < deadline and time.time() < hard_deadline:
        if handled >= max_popups:
            print(f"  [VA02] WARNING: tope de {max_popups} popups alcanzado; "
                  f"probable popup que no se cierra. Abortando handler.",
                  file=sys.stderr)
            break

        popup_processed_this_iter = False
        for wnd_id in ("wnd[1]", "wnd[2]"):
            try:
                session.findById(wnd_id)
            except Exception:
                continue

            popup_text = ""
            try:
                popup_text = session.findById(wnd_id).Text or ""
            except Exception:
                pass
            for field in ("txtMESSTXT1", "txtMESSTXT2", "txtMESSTXT3"):
                try:
                    val = session.findById(f"{wnd_id}/usr/{field}").Text or ""
                    if val:
                        popup_text += " " + val
                except Exception:
                    continue

            lower = popup_text.lower()
            is_save = any(kw in lower for kw in _SAVE_POPUP_KEYWORDS)
            is_info = any(kw in lower for kw in _INFO_POPUP_KEYWORDS)

            if not (is_save or is_info):
                # No reconocido: guardar para diagnóstico y seguir.
                if popup_text and not other_popup_seen:
                    other_popup_seen = popup_text[:200]
                continue

            # Detección de popup que no se cierra: mismo (wnd, texto) seguido.
            key = (wnd_id, popup_text)
            if key == last_key:
                repeat += 1
            else:
                repeat = 0
                last_key = key
            if repeat >= 3:
                print(f"  [VA02] WARNING: el mismo popup reaparece tras manejarlo "
                      f"({repeat}x) y no se cierra: {popup_text[:80]!r}. Abortando.",
                      file=sys.stderr)
                return handled

            # 1) Save Incomplete Document -> Save (primer botón)
            if is_save:
                pressed = False
                for btn_id in ("btnSPOP-OPTION1", "btnSPOP-VAROPTION1"):
                    try:
                        session.findById(f"{wnd_id}/usr/{btn_id}").press()
                        _wait_ready(session)
                        print(f"  [VA02] Popup 'Save Incomplete Document' -> Save "
                              f"({btn_id}).")
                        pressed = True
                        break
                    except Exception:
                        continue
                if pressed:
                    handled += 1
                    popup_processed_this_iter = True
                    # Resetear el deadline para esperar el próximo popup en cascada
                    deadline = time.time() + max_wait
                    break
                else:
                    print(f"  [VA02] WARNING: popup incomplete pero no se pudo "
                          f"presionar Save en {wnd_id}.", file=sys.stderr)
                    return handled

            # 2) Popup informativo -> Enter / Continue
            if is_info:
                if _press_continue_on_popup(session, wnd_id):
                    print(f"  [VA02] Popup informativo -> Continue: "
                          f"{popup_text[:80]!r}")
                    handled += 1
                    popup_processed_this_iter = True
                    deadline = time.time() + max_wait
                    break
                else:
                    print(f"  [VA02] WARNING: popup informativo pero no se pudo "
                          f"presionar Continue en {wnd_id}.", file=sys.stderr)
                    return handled

        if not popup_processed_this_iter:
            time.sleep(0.3)

    if other_popup_seen and handled == 0:
        print(f"  [VA02] DEBUG: popup detectado pero NO matcheó ningún keyword: "
              f"{other_popup_seen!r}", file=sys.stderr)
    return handled


# Alias retrocompatible: la lógica vieja llamaba _handle_save_incomplete_popup,
# ahora delegamos al handler completo que maneja también popups informativos.
def _handle_save_incomplete_popup(session, max_wait: float = 4.0) -> bool:
    return _handle_va02_post_save_popups(session, max_wait) > 0


def _check_va02_error_after_action(session, order: str, action_label: str) -> None:
    """
    Verifica si VA02 mostró un popup de error o un mensaje en el sbar
    después de una acción de eliminación. Lanza excepción si detecta error.

    Maneja el caso típico:
      'Item 000010 cannot be deleted because of subsequent document XXX'
    """
    # Detectar popup informativo de error
    for wnd_id in ("wnd[2]", "wnd[1]"):
        try:
            txt_msg = ""
            # Probar varios IDs comunes de mensaje en popups SAP
            for field in ("txtMESSTXT1", "txtMESSTXT2", "txtMESSTXT3"):
                try:
                    val = session.findById(f"{wnd_id}/usr/{field}").Text or ""
                    if val:
                        txt_msg += val + " "
                except Exception:
                    continue
            txt_msg = txt_msg.strip()

            if txt_msg:
                lower = txt_msg.lower()
                if any(kw in lower for kw in _ERROR_KEYWORDS):
                    # Cerrar popup y reportar
                    try:
                        session.findById(f"{wnd_id}/tbar[0]/btn[0]").press()
                        _wait_ready(session)
                    except Exception:
                        pass
                    raise RuntimeError(
                        f"VA02 [{action_label}] orden {order}: {txt_msg}"
                    )
                # Popup informativo pero no de error -> cerrar
                try:
                    session.findById(f"{wnd_id}/tbar[0]/btn[0]").press()
                    _wait_ready(session)
                except Exception:
                    pass
        except RuntimeError:
            raise
        except Exception:
            continue

    # También chequear el status bar
    try:
        sbar = session.findById("wnd[0]/sbar")
        sbar_text = (sbar.Text or "").strip()
        sbar_type = (sbar.MessageType or "").upper()
        if sbar_text and sbar_type in ("E", "A"):
            raise RuntimeError(
                f"VA02 [{action_label}] orden {order} (sbar {sbar_type}): {sbar_text}"
            )
    except RuntimeError:
        raise
    except Exception:
        pass


def _pad_ticket(ticket: str) -> str:
    """
    Padding de ticket al formato del PO Number de VA02 (BSTKD_E): un cero
    adelante; si queda en 9 dígitos, otro cero -> 10. Mismo criterio que la
    lógica original line-level.
    """
    padded = "0" + str(ticket)
    if len(padded) == 9:
        padded = "0" + padded
    return padded


def _va02_delete_selected_lines(session, order: str) -> None:
    """Asume filas ya seleccionadas en _VA02_TABLE: borra esas líneas y guarda."""
    delete_btn = (
        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01"
        "/ssubSUBSCREEN_BODY:SAPMV45A:4400"
        "/subSUBSCREEN_TC:SAPMV45A:4900"
        "/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POLO"
    )
    session.findById(delete_btn).press()
    _wait_ready(session)
    try:
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        _wait_ready(session)
    except Exception:
        pass
    session.findById("wnd[0]/tbar[0]/btn[11]").press()  # Save
    _wait_ready(session)
    _handle_save_incomplete_popup(session)
    _check_va02_error_after_action(session, order, "delete-line")


def _delete_intercompany_lines_va02(session, order: str, tickets: set) -> None:
    """
    VA02 "verificar y decidir" (SEGURO con órdenes mixtas):
      - Lee TODAS las líneas de la orden (PO Number BSTKD_E = ticket padded).
      - Si TODAS las líneas están en `tickets` -> borra la ORDEN COMPLETA.
      - Si está MEZCLADA -> borra SOLO las líneas en alcance; el resto persiste.
      - Si NINGUNA línea está en alcance -> no toca nada (idempotente).

    Lanza RuntimeError si SAP reporta error (subsequent document, etc.).
    """
    _go_back(session, 3)
    _navigate_to(session, "VA02")
    _wait_ready(session)

    session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = order
    session.findById("wnd[0]").sendVKey(0)
    _wait_ready(session)

    # Cerrar popup informativo inicial (si lo hay).
    mensaje = ""
    try:
        mensaje = session.findById("wnd[1]/usr/txtMESSTXT1").Text or ""
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        _wait_ready(session)
    except Exception:
        pass

    # Idempotencia: orden ya borrada -> éxito.
    if _order_already_gone(session, order, mensaje):
        return

    padded = {_pad_ticket(t) for t in tickets}

    # Leer Y seleccionar en UNA sola pasada: la selección de una fila requiere
    # que esté VISIBLE, así que se marca en el momento de leerla (mecánica
    # probada del borrado line-level original). Borrar solo las líneas en alcance
    # vacía la orden si TODAS lo están (SAP la elimina sola) o conserva las demás
    # si está mezclada. Si la tabla no está, la orden no abrió en modo edición.
    try:
        table = session.findById(_VA02_TABLE)
    except Exception:
        sbar_text, sbar_type = _read_sbar(session)
        raise RuntimeError(
            f"VA02 orden {order}: no abrió la tabla de items en modo edición. "
            f"SAP sbar: {sbar_text!r} (type={sbar_type})."
        )
    try:
        row_count = int(table.RowCount)
    except Exception:
        row_count = 0
    try:
        visible_count = int(table.VisibleRowCount) or 8
    except Exception:
        visible_count = 8

    in_scope = 0
    otras = 0
    first_visible = 0
    for absolute_row in range(row_count):
        if absolute_row >= first_visible + visible_count:
            try:
                session.findById(_VA02_TABLE).verticalScrollbar.Position = absolute_row
                _wait_ready(session)
                first_visible = absolute_row
            except Exception as e:
                print(f"  [VA02] orden {order}: error en scroll a fila "
                      f"{absolute_row}: {e}", file=sys.stderr)
                continue
        visible_row = absolute_row - first_visible
        try:
            po = session.findById(
                f"{_VA02_TABLE}/txtVBKD-BSTKD_E[6,{visible_row}]"
            ).Text or ""
        except Exception:
            continue
        if po and po in padded:
            # La fila está visible AHORA -> seleccionarla en el momento.
            try:
                session.findById(_VA02_TABLE).getAbsoluteRow(absolute_row).Selected = True
                in_scope += 1
            except Exception as e:
                print(f"  [VA02] orden {order}: error marcando fila "
                      f"{absolute_row}: {e}", file=sys.stderr)
        elif po:
            otras += 1

    print(f"  [VA02] orden {order}: {row_count} línea(s); "
          f"{in_scope} en alcance (seleccionadas), {otras} a conservar.")

    if in_scope == 0:
        # Las líneas de nuestros tickets ya no están -> idempotente.
        print(f"  [VA02] orden {order}: sin líneas en alcance (ya borradas). OK.")
        return

    if otras == 0:
        print(f"  [VA02] orden {order}: todas las líneas en alcance -> se vacía la orden.")
    else:
        print(f"  [VA02] orden {order}: MEZCLADA -> borrar {in_scope}, conservar {otras}.")

    _va02_delete_selected_lines(session, order)


def _delete_intercompany_order_va02(session, order: str, delivery: str) -> None:
    """
    VA02: Elimina un Intercompany order.
    Si existe línea con el delivery -> elimina esa línea.
    Si no -> elimina el order completo vía menú.

    Lanza excepción si SAP muestra popup de error (ej: "cannot be deleted
    because of subsequent document") — esto permite que Batch 6 lo marque
    como fallido correctamente en vez de silenciosamente reportar éxito.
    """
    _go_back(session, 3)
    _navigate_to(session, "VA02")
    _wait_ready(session)

    session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = order
    session.findById("wnd[0]").sendVKey(0)
    _wait_ready(session)

    mensaje = ""
    try:
        mensaje = session.findById("wnd[1]/usr/txtMESSTXT1").Text or ""
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        _wait_ready(session)
    except Exception:
        pass

    # Idempotencia: si la orden ya no existe (ya fue borrada en una corrida
    # previa), el objetivo está cumplido -> salir sin error.
    if _order_already_gone(session, order, mensaje):
        return

    if mensaje:
        delivery_padded = "0" + delivery
        if len(delivery_padded) == 9:
            delivery_padded = "0" + delivery_padded

        deleted = False
        po_numbers_seen = []  # diagnóstico: todos los PO Numbers que vimos
        lineas_matched = 0

        # Usar RowCount del table control como límite duro. El loop anterior
        # terminaba prematuramente con `break` cuando una celda no estaba
        # accesible (porque SAP aún no había cargado la siguiente página tras
        # el scroll) o cuando una celda intermedia tenía BSTKD_E vacío. Ahora
        # iteramos exactamente row_count veces y manejamos errores con
        # `continue` para no perder filas por timing.
        # Si la tabla no está, la orden no abrió en modo edición: error claro
        # (con el mensaje real de SAP) en vez del críptico 'control not found'.
        try:
            table = session.findById(_VA02_TABLE)
        except Exception:
            sbar_text, sbar_type = _read_sbar(session)
            raise RuntimeError(
                f"VA02 orden {order}: no abrió la tabla de items en modo edición. "
                f"SAP sbar: {sbar_text!r} (type={sbar_type}). "
                f"popup inicial: {mensaje[:80]!r}"
            )
        try:
            row_count = int(table.RowCount)
        except Exception:
            row_count = 0
        try:
            visible_count = int(table.VisibleRowCount) or 8
        except Exception:
            visible_count = 8

        print(f"  [VA02] orden {order}: table RowCount={row_count}, "
              f"VisibleRowCount={visible_count}")

        first_visible = 0
        for absolute_row in range(row_count):
            # Scroll cuando la fila absoluta cae fuera del rango visible actual
            if absolute_row >= first_visible + visible_count:
                try:
                    session.findById(_VA02_TABLE).verticalScrollbar.Position = absolute_row
                    _wait_ready(session)
                    first_visible = absolute_row
                    # Re-obtener referencia tras scroll (los controles se recargan)
                    table = session.findById(_VA02_TABLE)
                except Exception as e:
                    print(f"  [VA02] orden {order}: error en scroll a fila "
                          f"{absolute_row}: {e}", file=sys.stderr)
                    continue

            visible_row = absolute_row - first_visible
            try:
                po_number = session.findById(
                    f"{_VA02_TABLE}/txtVBKD-BSTKD_E[6,{visible_row}]"
                ).Text or ""
            except Exception:
                # Línea no accesible momentáneamente — continuar, NO break
                # (era el bug que abortaba el loop prematuramente).
                continue

            if po_number:
                po_numbers_seen.append(po_number)

            if po_number == delivery_padded:
                try:
                    session.findById(_VA02_TABLE).getAbsoluteRow(absolute_row).Selected = True
                    deleted = True
                    lineas_matched += 1
                except Exception as e:
                    print(f"  [VA02] orden {order}: error marcando fila "
                          f"{absolute_row}: {e}", file=sys.stderr)

        if deleted:
            print(f"  [VA02] orden {order}: {lineas_matched} línea(s) matching "
                  f"delivery {delivery_padded} seleccionadas. Eliminando...")
            delete_btn = (
                "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01"
                "/ssubSUBSCREEN_BODY:SAPMV45A:4400"
                "/subSUBSCREEN_TC:SAPMV45A:4900"
                "/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POLO"
            )
            session.findById(delete_btn).press()
            _wait_ready(session)
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
            _wait_ready(session)
            session.findById("wnd[0]/tbar[0]/btn[11]").press()  # Save
            _wait_ready(session)

            # Manejar popup "Save Incomplete Document" ANTES del check de errores
            # (sino _check_va02_error_after_action lo cierra con Cancel y pierde
            # los cambios silenciosamente).
            _handle_save_incomplete_popup(session)

            # Detectar popup de error "subsequent document" o sbar error
            _check_va02_error_after_action(session, order, "delete-line")
        else:
            # DIAGNÓSTICO: imprimir todos los PO Numbers que SÍ vimos en VA02
            unicos = sorted(set(po_numbers_seen))
            print(f"  [VA02] orden {order}: NO se encontró línea con PO={delivery_padded}")
            print(f"  [VA02] PO Numbers visibles en la orden ({len(unicos)} únicos): {unicos[:10]}")
            raise RuntimeError(
                f"VA02 orden {order}: no se encontró línea matching delivery {delivery} "
                f"(padded={delivery_padded}). PO Numbers visibles: {unicos[:5]}"
            )
    else:
        # Si el menú Delete no está, la orden no abrió en modo edición: error
        # claro con el mensaje real de SAP en vez del críptico 'control not found'.
        try:
            session.findById("wnd[0]/mbar/menu[0]/menu[11]").Select()
            _wait_ready(session)
        except Exception:
            sbar_text, sbar_type = _read_sbar(session)
            raise RuntimeError(
                f"VA02 orden {order}: no se pudo abrir el menú Delete "
                f"(¿la orden abrió en modo edición?). SAP sbar: {sbar_text!r} "
                f"(type={sbar_type})"
            )
        try:
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
            _wait_ready(session)
        except Exception:
            pass

        # Save explícito: el menu Delete puede dejar el documento marcado pero
        # sin persistir hasta el save. Sin esto, el siguiente /n descarta todo.
        try:
            session.findById("wnd[0]/tbar[0]/btn[11]").press()
            _wait_ready(session)
        except Exception:
            pass

        # Manejar popup "Save Incomplete Document" ANTES del check de errores
        _handle_save_incomplete_popup(session)

        # Detectar popup de error después del menú Delete
        _check_va02_error_after_action(session, order, "delete-order")
        _go_back(session)


def _read_me22n_rows(session) -> tuple:
    """
    Lee la tabla de items de ME22N (GuiTableControl) con scroll. Retorna
    (data, col_names) con data = {absolute_row: {col_name: value}}.
    Best-effort: ante cualquier incertidumbre retorna lo que pudo leer.
    """
    table = session.findById(_ME22N_TABLE)
    try:
        total = int(table.RowCount)
    except Exception:
        total = 0
    try:
        visible = int(table.VisibleRowCount) or 1
    except Exception:
        visible = 1

    col_names = []
    try:
        ccount = int(table.Columns.Count)
        for i in range(ccount):
            try:
                col_names.append(table.Columns.ElementAt(i).Name)
            except Exception:
                try:
                    col_names.append(table.Columns.Item(i).Name)
                except Exception:
                    col_names.append(str(i))
    except Exception:
        col_names = []

    data = {}
    top = 0
    guard = 0
    while top < total and guard <= total + visible:
        guard += visible
        try:
            table = session.findById(_ME22N_TABLE)
            table.verticalScrollbar.Position = top
            _wait_ready(session)
            table = session.findById(_ME22N_TABLE)
        except Exception:
            pass
        for vis in range(min(visible, total - top)):
            abs_row = top + vis
            row_vals = {}
            for ci in range(len(col_names)):
                try:
                    row_vals[col_names[ci]] = table.GetCell(vis, ci).Text or ""
                except Exception:
                    continue
            data[abs_row] = row_vals
        top += visible
    return data, col_names


def _delete_intracompany_lines_me22n(session, order: str, tickets: set) -> None:
    """
    ME22N "verificar y decidir" (SEGURO con POs mixtos). A diferencia del código
    viejo (que borraba la fila 0 a ciegas), aquí:
      - Lee las líneas del PO e identifica la columna del ticket por VALOR
        (auto-detección: la columna cuyos valores matchean los tickets en alcance).
      - Selecciona y borra SOLO las líneas en alcance (si todas lo están, eso es
        el PO completo). Las líneas de otros tickets se conservan.
      - Si NO se puede identificar la columna del ticket -> error seguro (NO borra
        nada), para no afectar líneas ajenas. El volcado de columnas queda en el
        log para fijar el campo.
    """
    global _ME22N_COLS_DUMPED

    _go_back(session, 3)
    _navigate_to(session, "ME22N")
    _wait_ready(session)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()  # Other purchase order
    _wait_ready(session)
    session.findById(
        "wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN"
    ).Text = order
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    _wait_ready(session)
    time.sleep(2)

    # Idempotencia: si el PO ya no existe.
    if _order_already_gone(session, order):
        return

    data, col_names = _read_me22n_rows(session)

    # Volcado one-shot de columnas (diagnóstico para identificar el campo del ticket).
    if not _ME22N_COLS_DUMPED:
        _ME22N_COLS_DUMPED = True
        print(f"  [ME22N] columnas de la tabla de items ({len(col_names)}): "
              f"{col_names}", file=sys.stderr)
        for abs_row, vals in list(data.items())[:3]:
            print(f"    fila {abs_row}: {vals}", file=sys.stderr)

    # Auto-detección del campo del ticket: columna cuyos valores (normalizados)
    # matchean los tickets en alcance.
    in_scope_norm = {_normalize_ticket(t) for t in tickets}
    best_col, best_hits = None, 0
    for cname in col_names:
        hits = sum(1 for vals in data.values()
                   if _normalize_ticket(vals.get(cname, "")) in in_scope_norm)
        if hits > best_hits:
            best_hits, best_col = hits, cname

    if not best_col or best_hits == 0:
        # No se identificó el campo del ticket -> NO borrar a ciegas (seguridad).
        raise RuntimeError(
            f"ME22N PO {order}: no se identificó la columna del ticket "
            f"(columnas={col_names}). NO se borra para no afectar líneas ajenas. "
            f"Revisar el volcado de columnas en el log y fijar el campo."
        )

    in_scope_rows = [r for r, vals in data.items()
                     if _normalize_ticket(vals.get(best_col, "")) in in_scope_norm]
    # "otras" = todas las que NO están confirmadas en alcance (incluye ticket
    # vacío/no clasificable). Solo se borran las in_scope, nunca estas.
    in_scope_set = set(in_scope_rows)
    otras_rows = [r for r in data if r not in in_scope_set]

    print(f"  [ME22N] PO {order}: campo ticket='{best_col}'; {len(data)} línea(s); "
          f"{len(in_scope_rows)} en alcance, {len(otras_rows)} a conservar.")

    if not in_scope_rows:
        print(f"  [ME22N] PO {order}: sin líneas en alcance (ya borradas). OK.")
        return

    # Seleccionar SOLO las líneas en alcance (si todas lo están, es el PO completo).
    # La selección requiere que la fila esté visible -> scroll a cada una antes.
    marcadas = 0
    for r in sorted(in_scope_rows):
        try:
            table = session.findById(_ME22N_TABLE)
            table.verticalScrollbar.Position = r
            _wait_ready(session)
            table = session.findById(_ME22N_TABLE)
            table.getAbsoluteRow(r).Selected = True
            marcadas += 1
        except Exception as e:
            print(f"  [ME22N] PO {order}: error marcando fila {r}: {e}",
                  file=sys.stderr)
    if marcadas == 0:
        raise RuntimeError(
            f"ME22N PO {order}: no se pudo marcar ninguna línea en alcance.")

    if otras_rows:
        print(f"  [ME22N] PO {order}: MEZCLADO -> borrar {marcadas} línea(s), "
              f"conservar {len(otras_rows)}.")
    else:
        print(f"  [ME22N] PO {order}: todas las líneas en alcance -> PO completo.")

    session.findById(_ME22N_DELETE_BTN).press()
    _wait_ready(session)
    try:
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        _wait_ready(session)
    except Exception:
        pass
    session.findById("wnd[0]/tbar[0]/btn[11]").press()  # Save
    _wait_ready(session)
    try:
        session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()
        _wait_ready(session)
    except Exception:
        pass
    _check_va02_error_after_action(session, order, "delete-PO-ME22N")


def _delete_intracompany_order_me22n(session, order: str) -> None:
    """
    ME22N: Elimina un Intracompany purchase order.
    Lanza excepción si SAP muestra popup de error o sbar error después del save.
    """
    _go_back(session, 3)
    _navigate_to(session, "ME22N")
    _wait_ready(session)

    session.findById("wnd[0]/tbar[1]/btn[17]").press()  # Other purchase order
    _wait_ready(session)
    session.findById(
        "wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN"
    ).Text = order
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    _wait_ready(session)
    time.sleep(2)

    session.findById(_ME22N_TABLE).getAbsoluteRow(0).Selected = True
    session.findById(f"{_ME22N_TABLE}/btnMEPO1211-STATUSICON[0,0]").SetFocus()

    session.findById(
        "wnd[0]/usr/subSUB0:SAPLMEGUI:0019"
        "/subSUB2:SAPLMEVIEWS:1100"
        "/subSUB2:SAPLMEVIEWS:1200"
        "/subSUB1:SAPLMEGUI:1211/btnDELETE"
    ).press()
    _wait_ready(session)
    try:
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        _wait_ready(session)
    except Exception:
        pass
    session.findById("wnd[0]/tbar[0]/btn[11]").press()  # Save
    _wait_ready(session)

    try:
        session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()
        _wait_ready(session)
    except Exception:
        pass

    # Detectar popup de error o sbar error (mismo patrón que VA02)
    _check_va02_error_after_action(session, order, "delete-PO-ME22N")


def _delete_order_lines_for_tickets(session, order: str, tickets: set) -> None:
    """
    Entrypoint SEGURO de borrado de órdenes para Batch 8: "verificar y decidir".
    Despacha según el tipo de orden:
      - order empieza con '47' -> Intracompany (ME22N).
      - si no                  -> Intercompany (VA02).

    En ambos casos borra SOLO las líneas de los tickets en alcance (o la orden
    completa si TODAS sus líneas lo están), nunca líneas de tickets ajenos.

    `tickets`: conjunto de tickets en alcance que mapean a esta orden.
    """
    tickets = set(tickets)
    if str(order).startswith("47"):
        _delete_intracompany_lines_me22n(session, order, tickets)
    else:
        _delete_intercompany_lines_va02(session, order, tickets)


def cancel_order_by_ticket(session, order: str, ticket: str) -> None:
    """
    Modo manual / line-level: borra solo las líneas de la order que matchean
    el ticket dado (por PO Number = ticket padded). Si la order queda vacía,
    SAP la borra automáticamente. Si quedan otras líneas (de tickets que NO
    están en la hoja Manual Orders), la order persiste con esas líneas.

    Dispatches según prefijo del order:
      - order[:2] == "47" -> ME22N (intracompany purchase order)
      - else              -> VA02 line-delete (intercompany sales order)

    Reusa la lógica probada de `_delete_intercompany_order_va02` y
    `_delete_intracompany_order_me22n`, pasando el ticket como "delivery"
    porque en este SAP el campo BSTKD_E (PO Number en VA02) guarda el ticket.

    Lanza RuntimeError si SAP muestra popup de error (subsequent document,
    cannot be deleted, etc.) — el batch que la invoca debe capturarlo.
    """
    if str(order).startswith("47"):
        _delete_intracompany_order_me22n(session, order)
    else:
        _delete_intercompany_order_va02(session, order, delivery=ticket)
