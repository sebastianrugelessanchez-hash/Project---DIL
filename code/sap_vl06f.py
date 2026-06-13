import sys

from sap_utils import (
    _navigate_to, _go_back, _wait_ready,
    _enter_multi_values, _POPUP_TABLE,
    _wnd_exists, _find_popup_wnd, _popup_table,
    _normalize_ticket,
)


def _vl06f_delivery_filter(session, tickets: list) -> None:
    """
    Abre el popup multi-valor de IT_VBELN (Delivery) y entra los tickets.
    El campo Delivery es accesible directamente en la pantalla de selección de VL06F.
    """
    session.findById("wnd[0]/usr/btn%_IT_VBELN_%_APP_%-VALU_PUSH").press()
    _wait_ready(session)

    popup = _find_popup_wnd(session)
    tbl   = _popup_table(popup)

    # btn[16] = "Delete Entire Selection" (Shift+F4) — limpia valores previos
    # NOTA: antes había btn[24] aquí que en realidad es "Upload from Clipboard" —
    # eso causaba un upload con basura del clipboard antes del upload real.
    try:
        session.findById(f"{popup}/tbar[0]/btn[16]").press()
        _wait_ready(session)
    except Exception:
        pass

    _enter_multi_values(session, tbl, tickets)
    session.findById(f"{popup}/tbar[0]/btn[8]").press()  # btn[8] = Copy (F8) — cerrar popup
    _wait_ready(session)


def _select_bol_layout(session) -> None:
    """
    Selecciona el layout /02C (BOL STATUS) en la pantalla de resultados de VL06F.
    Usa Ctrl+F9 (VKey 33) para abrir el diálogo Choose Layout.

    Este layout estabiliza los nombres técnicos de columnas (VBELN, ZZVBELN, etc.)
    de forma que GetCellValue funcione consistentemente. SIN este switch, el grid
    usa la layout default del usuario que puede tener columnas con nombres distintos.
    """
    try:
        session.findById("wnd[0]").sendVKey(33)  # Ctrl+F9 = Choose Layout
        _wait_ready(session)

        if not _wnd_exists(session, "wnd[1]"):
            print("  [VL06F] Choose Layout no abrió diálogo — sigo con layout actual", file=sys.stderr)
            return

        shell = session.findById(
            "wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501"
            "/cntlG51_CONTAINER/shellcont/shell"
        )
        layouts_disponibles = []
        for i in range(shell.RowCount):
            for col in ("VARIANT", "LAYOUT", "DISVARIANT"):
                try:
                    val = (shell.GetCellValue(i, col) or "").strip()
                    if val:
                        layouts_disponibles.append(val)
                        if val == "/02C":
                            shell.setCurrentCell(i, col)
                            shell.selectedRows = str(i)
                            shell.doubleClickCurrentCell()
                            _wait_ready(session)
                            print("  [VL06F] Layout /02C BOL STATUS aplicado")
                            return
                except Exception:
                    continue

        print(f"  [VL06F] Layout /02C NO encontrado. Disponibles: {layouts_disponibles[:20]}", file=sys.stderr)
        try:
            session.findById("wnd[1]").sendVKey(12)
        except Exception:
            pass
    except Exception as e:
        print(f"  [VL06F] Error seleccionando layout: {e}", file=sys.stderr)


_VL06F_COL_CANDIDATES = {
    "vbeln":       ("VBELN", "VBELN_VL", "DELIVERY", "LIKP-VBELN"),
    "billing_doc": ("ZZVBELN", "VBELN_VF", "VRBELN", "FKART", "VFAKT"),
    "shpt_cst":    ("ZZFKNUM", "FKNUM", "VFKP_FKNUM", "SHPTCOST"),
    "shipment":    ("ZZTKNUM", "TKNUM", "VTTK_TKNUM", "SHIPMENT"),
    "wbstk":       ("WBSTK", "GBSTK", "GM_STATUS", "GM"),
    "invoice_il":  ("ZZVBELN_IL", "VBELN_IL", "INVOICE_LIST"),
    "block":       ("LIFSK", "FAKSP", "BLOCK"),
}


def _build_vl06f_column_map(grid) -> dict:
    """
    Detecta el nombre técnico real de cada columna usada por el pipeline.
    SIEMPRE imprime las columnas disponibles para diagnóstico.
    """
    try:
        cols = list(grid.ColumnOrder)
    except Exception:
        cols = []

    print(f"  [VL06F] Columnas disponibles ({len(cols)}): {cols}")

    mapping = {}
    not_found = []
    for field, options in _VL06F_COL_CANDIDATES.items():
        match = next((opt for opt in options if opt in cols), None)
        if match is None:
            not_found.append(field)
            mapping[field] = options[0]  # fallback al primero
        else:
            mapping[field] = match

    if not_found:
        print(f"  [VL06F] Columnas NO encontradas: {not_found}", file=sys.stderr)

    print(f"  [VL06F] Column map final: {mapping}")
    return mapping


def read_vl06f_data(session, tickets: list) -> dict:
    """
    Navega a VL06F, filtra por los tickets dados y lee el grid de entregas.

    Retorna: {vbeln: {billing_doc, shpt_cst, shipment, wbstk, delivery, invoice_il}}

    NOTA: La regla R001 (skip tickets con factura intercompany — invoice_il
    empieza con '7') NO se aplica aquí. Esta función solo LEE VL06F. La
    exclusión del ticket completo se hace en main.process_chunk a partir del
    campo `invoice_il` retornado.
    """
    # CRÍTICO: si tickets está vacío, NO navegar a VL06F. Si lo hacemos,
    # el popup se queda vacío, F8 filtra con NADA y SAP devuelve TODA
    # la base de deliveries -> congelamiento.
    if not tickets:
        print("  [VL06F] Lista de tickets vacía — skip (evita F8 sin filtro)")
        return {}

    _navigate_to(session, "VL06F")
    _wait_ready(session)

    # Expandir todos los criterios dinámicos (VBA: btn[19] "All selections").
    # Asegura que los campos IT_WADAT-LOW/HIGH estén accesibles para limpiar.
    try:
        session.findById("wnd[0]/tbar[1]/btn[19]").press()
        _wait_ready(session)
    except Exception:
        pass

    for field in ("ctxtIT_WADAT-LOW", "ctxtIT_WADAT-HIGH"):
        try:
            session.findById(f"wnd[0]/usr/{field}").text = ""
        except Exception:
            pass

    _vl06f_delivery_filter(session, tickets)

    session.findById("wnd[0]/tbar[1]/btn[8]").press()  # F8
    _wait_ready(session)

    # Aplicar layout /02C — fija los nombres técnicos de columnas (VBELN, etc.)
    # para que GetCellValue funcione. Sin esto, la layout default del usuario
    # puede tener columnas con nombres distintos -> todas las filas devuelven "".
    _select_bol_layout(session)

    data = {}
    try:
        grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        row_count = grid.RowCount
        print(f"  [VL06F] Grid RowCount={row_count}")

        # Detectar nombres reales de columnas (puede variar según versión SAP / layout)
        col = _build_vl06f_column_map(grid)

        # Materializar TODAS las filas. Los grids ALV virtualizados devuelven ""
        # en GetCellValue para filas fuera del viewport. Combinamos selectAll
        # con scroll forzado (firstVisibleRow al final -> SAP carga todo).
        try:
            grid.selectAll()
        except Exception:
            try:
                grid.SelectAll()
            except Exception:
                pass

        try:
            grid.firstVisibleRow = max(0, row_count - 1)
            _wait_ready(session)
            grid.firstVisibleRow = 0
            _wait_ready(session)
        except Exception:
            pass

        for row in range(row_count):
            try:
                vbeln_raw = grid.GetCellValue(row, col["vbeln"])
                vbeln = _normalize_ticket(vbeln_raw)
                if not vbeln:
                    continue  # skip filas de subtotal/agrupamiento, NO terminar

                invoice_il = grid.GetCellValue(row, col["invoice_il"]) or ""
                billing_doc = grid.GetCellValue(row, col["billing_doc"]) or ""

                data[vbeln] = {
                    "billing_doc": billing_doc,
                    "shpt_cst":    grid.GetCellValue(row, col["shpt_cst"]) or "",
                    "shipment":    grid.GetCellValue(row, col["shipment"]) or "",
                    "wbstk":       grid.GetCellValue(row, col["wbstk"])   or "",
                    "block":       grid.GetCellValue(row, col["block"])   or "",
                    "delivery":    vbeln,
                    "invoice_il":  invoice_il,
                }
            except Exception:
                continue

        print(f"  [VL06F] Filas leídas con VBELN válido: {len(data)}")

        # Diagnóstico: ¿cuántos tickets tienen cada campo poblado?
        n_billing = sum(1 for d in data.values() if d["billing_doc"])
        n_shpt    = sum(1 for d in data.values() if d["shpt_cst"])
        n_ship    = sum(1 for d in data.values() if d["shipment"])
        n_wbstk   = sum(1 for d in data.values() if d["wbstk"])
        n_block   = sum(1 for d in data.values() if d["block"])
        print(f"  [VL06F] Tickets con datos poblados:")
        print(f"           billing_doc={n_billing}, shpt_cst={n_shpt}, "
              f"shipment={n_ship}, wbstk={n_wbstk}, BLOCKED={n_block} (de {len(data)})")
        if n_block:
            # Mostrar distribución de códigos de bloqueo
            from collections import Counter
            block_codes = Counter(d["block"] for d in data.values() if d["block"])
            print(f"  [VL06F] Códigos de bloqueo encontrados: {dict(block_codes)}")

        # Muestras de 3 filas (primera, media, última) para verificar lectura
        if data:
            keys = list(data.keys())
            sample_indices = sorted(set([0, len(keys) // 2, len(keys) - 1]))
            for idx in sample_indices:
                k = keys[idx]
                print(f"  [VL06F] Muestra row {idx} ({k}): {data[k]}")

    except Exception as e:
        print(f"  [VL06F] Error leyendo grid: {e}", file=sys.stderr)

    return data


def delete_bol(session, delivery: str) -> None:
    """
    VL06F: Elimina el delivery/BOL del monitor de entregas (per-ticket).
    Corresponde a reversarElDelivery en workFlow.bas.

    Implementación con step-by-step diagnostics + status bar capture.
    Cualquier excepción contiene el prefix [step=N] para identificar
    dónde falló. La verificación REAL (si SAP efectivamente borró la
    delivery) se hace en `verify_bol_deleted_bulk` post-batch en main.py.
    """
    if not delivery:
        return

    # PASO 1: navegar a VL06F
    try:
        _navigate_to(session, "VL06F")
        _wait_ready(session)
    except Exception as e:
        raise RuntimeError(f"[BOL {delivery}][step=1 navegar VL06F] {e}")

    # Limpiar filtros de fecha (best effort, no fatal si falla)
    for field in ("ctxtIT_WADAT-LOW", "ctxtIT_WADAT-HIGH"):
        try:
            session.findById(f"wnd[0]/usr/{field}").Text = ""
        except Exception:
            pass

    # PASO 2: filtrar por el delivery
    try:
        _vl06f_delivery_filter(session, [delivery])
    except Exception as e:
        raise RuntimeError(f"[BOL {delivery}][step=2 filtrar delivery] {e}")

    # PASO 3: F8 ejecutar
    try:
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        _wait_ready(session)
    except Exception as e:
        raise RuntimeError(f"[BOL {delivery}][step=3 F8] {e}")

    # PASO 4: validar que el grid tenga filas (defensive contract)
    try:
        grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        row_count = grid.RowCount
    except Exception as e:
        raise RuntimeError(f"[BOL {delivery}][step=4 acceder grid] {e}")

    if row_count == 0:
        # Grid vacío post-F8 = delivery no encontrado. Capturar mensaje SAP.
        sbar_msg = ""
        try:
            sbar_msg = (session.findById("wnd[0]/sbar").Text or "").strip()
        except Exception:
            pass
        raise RuntimeError(
            f"[BOL {delivery}][step=4 grid vacío post-F8] "
            f"delivery no encontrado en VL06F. sbar: {sbar_msg!r}"
        )

    # PASO 5: seleccionar fila 0
    try:
        grid.setCurrentCell(-1, "")
        grid.selectedRows = "0"
    except Exception as e:
        raise RuntimeError(f"[BOL {delivery}][step=5 select row 0] {e}")

    # PASO 6a: btn[13] (preparación del delete)
    try:
        session.findById("wnd[0]/tbar[1]/btn[13]").press()
        _wait_ready(session)
    except Exception as e:
        raise RuntimeError(f"[BOL {delivery}][step=6a btn[13]] {e}")

    # PASO 6b: btn[14] (ejecuta el delete)
    try:
        session.findById("wnd[0]/tbar[1]/btn[14]").press()
        _wait_ready(session)
    except Exception as e:
        raise RuntimeError(f"[BOL {delivery}][step=6b btn[14]] {e}")

    # PASO 7: confirmar popup principal (best effort — algunos flujos no lo muestran)
    try:
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        _wait_ready(session)
    except Exception:
        pass
    # Algunas versiones de SAP muestran una 2da confirmación
    try:
        session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()
        _wait_ready(session)
    except Exception:
        pass

    # PASO 8: capturar status bar (observability)
    sap_msg = ""
    sap_mtype = ""
    try:
        sbar = session.findById("wnd[0]/sbar")
        sap_msg = (sbar.Text or "").strip()
        sap_mtype = (sbar.MessageType or "").strip()
        if sap_msg:
            print(f"  [BOL {delivery}] SAP ({sap_mtype}): {sap_msg}")
    except Exception:
        pass

    # PASO 9: refresh — SIEMPRE corre (fuera de paso fallable), para dejar
    # VL06F en estado limpio para la próxima iteración. Si el refresh falla,
    # tampoco bloqueamos al ticket (la falla queda implícita en el sbar).
    try:
        session.findById("wnd[0]/tbar[1]/btn[20]").press()
        _wait_ready(session)
    except Exception:
        pass

    # PASO 10: si SAP reportó error (E o A) en sbar, propagar como fallo.
    # Esto cubre el caso de "rechazo silencioso" — SAP responde con error
    # en sbar pero sin lanzar popup; sin esto, marcaríamos exitoso por error.
    if sap_mtype in ("E", "A"):
        raise RuntimeError(
            f"[BOL {delivery}][step=8 SAP rechazó delete] ({sap_mtype}): {sap_msg}"
        )


def delete_bol_bulk(session, deliveries: list) -> None:
    """
    VL06F BULK: Elimina TODOS los BOLs en UNA sola pasada.

    Patrón:
    1. Navega a VL06F
    2. Filtra por TODOS los deliveries con _vl06f_delivery_filter (upload from clipboard)
    3. F8 -> resultados
    4. selectAll -> selecciona todas las filas del grid
    5. btn[13] + btn[14] -> aplica el delete a la selección completa
    6. Confirma popup
    7. Refresh

    Mismo patrón usado por reverse_pgi_bulk en VL09 (Batch 4).
    """
    if not deliveries:
        return

    _navigate_to(session, "VL06F")
    _wait_ready(session)

    # Expandir todos los criterios (btn[19] = All selections)
    try:
        session.findById("wnd[0]/tbar[1]/btn[19]").press()
        _wait_ready(session)
    except Exception:
        pass

    # Limpiar fechas
    for field in ("ctxtIT_WADAT-LOW", "ctxtIT_WADAT-HIGH"):
        try:
            session.findById(f"wnd[0]/usr/{field}").text = ""
        except Exception:
            pass

    # Filtrar por TODOS los deliveries de una vez (clipboard upload)
    _vl06f_delivery_filter(session, deliveries)

    # F8
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    _wait_ready(session)

    # Seleccionar TODAS las filas del grid (no solo la primera)
    grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    grid.currentCellColumn = ""
    try:
        grid.selectAll()
    except Exception:
        try:
            grid.SelectAll()
        except Exception:
            # Fallback: seleccionar manualmente todas las filas conocidas
            try:
                grid.selectedRows = ",".join(str(i) for i in range(grid.RowCount))
            except Exception:
                grid.selectedRows = "0"

    print(f"  [BOL bulk] Aplicando delete a {grid.RowCount} filas seleccionadas...")

    # Mismo flujo que delete_bol pero con todas las filas seleccionadas
    session.findById("wnd[0]/tbar[1]/btn[13]").press()
    _wait_ready(session)
    session.findById("wnd[0]/tbar[1]/btn[14]").press()
    _wait_ready(session)

    # Confirmar popup principal (puede aparecer una o varias veces según versión SAP)
    try:
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        _wait_ready(session)
    except Exception:
        pass

    try:
        session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()
        _wait_ready(session)
    except Exception:
        pass

    # Refresh
    try:
        session.findById("wnd[0]/tbar[1]/btn[20]").press()
        _wait_ready(session)
    except Exception:
        pass
    _wait_ready(session)
