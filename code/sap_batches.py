import sys
import time
from datetime import date

from sap_utils import _navigate_to, _go_back, _wait_ready, _enter_multi_values, _POPUP_TABLE


_VT02N_HEADER = (
    "wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE"
    "/ssubG_HEADER_SUBSCREEN2:SAPMV56A:1025"
)
_VT02N_PLANNING = (
    "wnd[0]/usr/subPLANNING:SAPLV56I_PLAN_SCREEN:0110"
    "/cntlV56I_PLAN_SCREEN_CONTAINER/shellcont/shell"
    "/shellcont[0]/shell"
)


# ---------------------------------------------------------------------------
# Batch 1 — VF11: Eliminar Billing Document
# ---------------------------------------------------------------------------

def delete_billing_document(session, billing_doc: str) -> None:
    """VF11: Reversa el billing document dado (uso per-ticket)."""
    if not billing_doc:
        return

    today = date.today().strftime("%m/%d/%Y")
    _navigate_to(session, "VF11")
    _wait_ready(session)

    session.findById(
        "wnd[0]/usr/tblSAPMV60ATCTRL_ERF_FAKT/ctxtKOMFK-VBELN[0,0]"
    ).Text = billing_doc
    session.findById("wnd[0]/usr/ctxtRV60A-FKDAT").Text = today
    session.findById("wnd[0]/tbar[0]/btn[11]").press()
    _wait_ready(session)
    _go_back(session)


def delete_billing_documents_bulk(session, billing_docs: list) -> None:
    """
    VF11: Reversa todos los billing documents en una sola transacción.
    La tabla de VF11 muestra 14 filas a la vez; btn[7] agrega una nueva página de 14.
    """
    if not billing_docs:
        return

    today = date.today().strftime("%m/%d/%Y")
    print(f"  [VF11] Cancelando {len(billing_docs)} billing doc(s). Fecha={today}")

    _navigate_to(session, "VF11")
    _wait_ready(session)

    j = 0
    for doc in billing_docs:
        session.findById(
            f"wnd[0]/usr/tblSAPMV60ATCTRL_ERF_FAKT/ctxtKOMFK-VBELN[0,{j}]"
        ).Text = doc
        if j == 13:
            session.findById("wnd[0]/tbar[1]/btn[7]").press()  # Nueva página de 14 filas
            _wait_ready(session)
            j = 0
        else:
            j += 1

    # Set fecha de cancelación y presionar Enter para validar
    session.findById("wnd[0]/usr/ctxtRV60A-FKDAT").Text = today
    session.findById("wnd[0]").sendVKey(0)  # Enter para validar la fecha
    _wait_ready(session)

    # Save (btn[11] = Ctrl+S)
    session.findById("wnd[0]/tbar[0]/btn[11]").press()
    _wait_ready(session)

    # Manejar popups de confirmación que VF11 muestra después del save
    for _ in range(3):
        try:
            popup_title = session.findById("wnd[1]").Text
            print(f"  [VF11] Popup detectado: '{popup_title}'")
            # Intentar confirmar con Yes/OK
            for btn_path in (
                "wnd[1]/usr/btnSPOP-OPTION1",       # Yes en POPUP estándar
                "wnd[1]/tbar[0]/btn[0]",            # OK
                "wnd[1]/tbar[0]/btn[11]",           # Save
            ):
                try:
                    session.findById(btn_path).press()
                    _wait_ready(session)
                    print(f"  [VF11] Popup cerrado con {btn_path}")
                    break
                except Exception:
                    continue
        except Exception:
            break  # No hay más popups

    # Leer status bar para confirmar
    try:
        sbar_text = session.findById("wnd[0]/sbar").Text or ""
        sbar_type = session.findById("wnd[0]/sbar").MessageType or ""
        if sbar_text:
            print(f"  [VF11] Status bar ({sbar_type}): {sbar_text}")
    except Exception:
        pass

    _go_back(session)


# ---------------------------------------------------------------------------
# Batch 2 — VI05: Eliminar Shipment Cost
# ---------------------------------------------------------------------------

def _enumerate_screen_controls(session, root_id: str = "wnd[0]/usr", max_depth: int = 3) -> list:
    """
    Diagnóstico: recorre los hijos del nodo root y devuelve los IDs encontrados.
    Útil para descubrir el ID real de un control cuando el esperado no existe.
    """
    found = []

    def _recurse(node_id, depth):
        if depth > max_depth:
            return
        try:
            node = session.findById(node_id)
        except Exception:
            return
        try:
            n_children = node.Children.Count
        except Exception:
            n_children = 0
        for i in range(n_children):
            try:
                child = node.Children.Item(i)
                cid = child.Id
                ctype = getattr(child, "Type", "?")
                cname = getattr(child, "Name", "")
                found.append(f"{ctype:15s} {cid}  name={cname!r}")
                _recurse(cid, depth + 1)
            except Exception:
                continue

    _recurse(root_id, 0)
    return found


def _cambiar_estado_transferencia(session, fecha: str) -> None:
    """
    Cambia el estado de transferencia del ShptCst de C a abierto antes de eliminar.
    Diagnóstico paso a paso para identificar exactamente qué control falla en
    tu versión/layout de VI05.
    """
    # SUB-PASO 7.1: seleccionar checkbox de la fila (chk[1,4])
    try:
        session.findById("wnd[0]/usr/chk[1,4]").Selected = True
    except Exception as e:
        # Volcar TODOS los controles visibles en el área wnd[0]/usr para
        # encontrar el ID real del checkbox de selección.
        print(f"  [_cambiar_estado] FALLO chk[1,4]: {e}", file=sys.stderr)
        print(f"  [_cambiar_estado] Controles visibles en wnd[0]/usr (top 30):", file=sys.stderr)
        try:
            ctrls = _enumerate_screen_controls(session, "wnd[0]/usr", max_depth=4)
            # Filtrar a los más relevantes: checkboxes, botones, labels
            relevant = [c for c in ctrls if any(t in c[:15] for t in ("GuiCheckBox", "GuiButton", "GuiLabel", "GuiGridView", "GuiShell"))]
            for c in relevant[:30]:
                print(f"    {c}", file=sys.stderr)
        except Exception as ee:
            print(f"    (no se pudo enumerar: {ee})", file=sys.stderr)
        raise RuntimeError(f"sub-step 7.1 chk[1,4].Selected=True: {e}")

    # SUB-PASO 7.2: entrar a modo edición (btn[17] = F8 Display->Change)
    try:
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        _wait_ready(session)
    except Exception as e:
        raise RuntimeError(f"sub-step 7.2 btn[17] edit mode: {e}")

    # SUB-PASO 7.3: ASEGURAR SLSTOR=True (= "solicitar cancelación de storage/transfer FI").
    # Esto es lo que habilita el campo STDAT (cancelación necesita fecha).
    # Manejo idempotente:
    #   - Si SLSTOR ya estaba en True: no tocar (alguien lo marcó antes manualmente).
    #   - Si SLSTOR estaba en False (caso típico): setear True y disparar refresh.
    try:
        chk_slstor = session.findById(
            "wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/chkVFKPD-SLSTOR[9,0]"
        )
        try:
            cur_sel = chk_slstor.Selected
            cur_ch  = getattr(chk_slstor, "Changeable", "?")
            print(f"  [_cambiar_estado] SLSTOR antes: Selected={cur_sel} Changeable={cur_ch}")
        except Exception:
            cur_sel = False

        if not cur_sel:
            chk_slstor.Selected = True
            chk_slstor.SetFocus()
            session.findById("wnd[0]").sendVKey(0)
            _wait_ready(session)
            print("  [_cambiar_estado] SLSTOR seteado a True (cancelación solicitada)")
        else:
            print("  [_cambiar_estado] SLSTOR ya era True — skip set")
    except Exception as e:
        print(f"  [_cambiar_estado] FALLO en chkVFKPD-SLSTOR[9,0]: {e}", file=sys.stderr)
        try:
            ctrls = _enumerate_screen_controls(session, "wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP", max_depth=3)
            print(f"  [_cambiar_estado] Controles del tblSAPMV54ACRTL_ITEMS_VFKP (top 20):", file=sys.stderr)
            for c in ctrls[:20]:
                print(f"    {c}", file=sys.stderr)
        except Exception:
            pass
        raise RuntimeError(f"sub-step 7.3 chkVFKPD-SLSTOR[9,0]: {e}")

    # SUB-PASO 7.3b: ¿apareció algún popup después del Enter?
    try:
        popup = session.findById("wnd[1]")
        popup_txt = (getattr(popup, "Text", "") or "")[:80]
        print(f"  [_cambiar_estado] POPUP detectado tras SLSTOR: '{popup_txt}'")
        # Intentar cerrarlo con OK
        try:
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            _wait_ready(session)
            print(f"  [_cambiar_estado] popup cerrado con tbar[0]/btn[0]")
        except Exception:
            pass
    except Exception:
        pass  # no hay popup, OK

    # SUB-PASO 7.4: setear fecha STDAT (probar múltiples formatos)
    try:
        fld_stdat = session.findById(
            "wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/ctxtVFKPD-STDAT[12,0]"
        )
        # Diagnóstico exhaustivo del estado del campo ANTES de escribir
        try:
            current  = fld_stdat.Text or ""
            max_len  = getattr(fld_stdat, "MaxLength", "?")
            ch       = getattr(fld_stdat, "Changeable", "?")
            ftype    = getattr(fld_stdat, "Type", "?")
            fname    = getattr(fld_stdat, "Name", "?")
            tooltip  = getattr(fld_stdat, "Tooltip", "?")
            print(f"  [_cambiar_estado] STDAT antes de set:")
            print(f"           Text={current!r} MaxLength={max_len} Changeable={ch}")
            print(f"           Type={ftype} Name={fname} Tooltip={tooltip!r}")
        except Exception as e:
            print(f"  [_cambiar_estado] no se pudieron leer props del STDAT: {e}")

        # SetFocus ANTES de escribir (en algunos casos SAP requiere foco activo)
        try:
            fld_stdat.SetFocus()
            _wait_ready(session)
        except Exception:
            pass

        today_obj = date.today()
        candidatos = [
            today_obj.strftime("%m/%d/%Y"),  # 05/20/2026 (US)
            today_obj.strftime("%d.%m.%Y"),  # 20.05.2026 (DE)
            today_obj.strftime("%d/%m/%Y"),  # 20/05/2026 (ES/CO)
            today_obj.strftime("%Y-%m-%d"),  # 2026-05-20 (ISO)
            today_obj.strftime("%m-%d-%Y"),  # 05-20-2026
            today_obj.strftime("%d-%m-%Y"),  # 20-05-2026
        ]
        formato_ok = None
        last_err = None
        for cand in candidatos:
            try:
                fld_stdat.Text = cand
                formato_ok = cand
                break
            except Exception as e:
                last_err = e
                continue
        if formato_ok is None:
            raise RuntimeError(f"ningún formato de fecha aceptado. Último error: {last_err}")
        print(f"  [_cambiar_estado] STDAT seteado con formato={formato_ok!r}")
    except Exception as e:
        raise RuntimeError(f"sub-step 7.4 ctxtVFKPD-STDAT[12,0]: {e}")

    # SUB-PASO 7.5: save (btn[11] = Ctrl+S)
    try:
        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        _wait_ready(session)
    except Exception as e:
        raise RuntimeError(f"sub-step 7.5 btn[11] save: {e}")

    # SUB-PASO 7.6: refresh (btn[8] = F8)
    try:
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        _wait_ready(session)
    except Exception as e:
        raise RuntimeError(f"sub-step 7.6 btn[8] refresh: {e}")


def _eliminar_shpt_cst(session, shpt_cst: str = "") -> None:
    """Elimina el ShptCst seleccionado y captura el mensaje de SAP."""
    session.findById("wnd[0]/usr/chk[1,4]").Selected = True
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    _wait_ready(session)
    session.findById("wnd[0]/tbar[1]/btn[14]").press()
    _wait_ready(session)
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    _wait_ready(session)

    # Capturar mensaje de SAP para diagnóstico (bloqueos, docs subsiguientes, etc.)
    try:
        sbar = session.findById("wnd[0]/sbar")
        msg = (sbar.Text or "").strip()
        mtype = (sbar.MessageType or "").strip()
        if msg:
            tag = f" {shpt_cst}" if shpt_cst else ""
            print(f"  [VI05{tag}] SAP ({mtype}): {msg}")
    except Exception:
        pass


def delete_shipment_cost(session, shpt_cst: str) -> None:
    """VI05: Elimina el shipment cost dado (uso per-ticket)."""
    if not shpt_cst:
        return

    today = date.today().strftime("%m/%d/%Y")
    _go_back(session, 3)
    _navigate_to(session, "VI05")
    _wait_ready(session)

    session.findById("wnd[0]/usr/btn%_S_FKNUM_%_APP_%-VALU_PUSH").press()
    _wait_ready(session)
    session.findById(f"{_POPUP_TABLE}/ctxtRSCSEL_255-SLOW_I[1,0]").Text = shpt_cst
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    _wait_ready(session)

    session.findById("wnd[0]/usr/btn%_S_STBER_%_APP_%-VALU_PUSH").press()
    _wait_ready(session)
    session.findById("wnd[1]/tbar[0]/btn[16]").press()  # Select All
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    _wait_ready(session)

    session.findById("wnd[0]/tbar[1]/btn[8]").press()  # F8
    _wait_ready(session)

    try:
        estado = session.findById("wnd[0]/usr/lbl[33,5]").Text.strip()
    except Exception:
        estado = ""

    if estado == "C":
        _cambiar_estado_transferencia(session, today)

    _eliminar_shpt_cst(session)

    session.findById("wnd[0]/tbar[1]/btn[8]").press()  # Refresh
    _wait_ready(session)


def _vi05_setup_selection_screen(session) -> None:
    """Navega a VI05 limpio y configura el rango de fechas via btn S_STBER."""
    _navigate_to(session, "VI05")
    _wait_ready(session)

    try:
        session.findById("wnd[0]/usr/btn%_S_STBER_%_APP_%-VALU_PUSH").press()
        _wait_ready(session)
        session.findById("wnd[1]/tbar[0]/btn[16]").press()  # Select All
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        _wait_ready(session)
    except Exception as e:
        print(f"  [VI05 setup] WARNING al configurar fechas: {e}", file=sys.stderr)


def _process_single_shpt_cst(session, shpt_cst: str, today: str) -> None:
    """
    Procesa UN ShptCst con diagnóstico por paso y recovery a VI05 al inicio.
    Lanza RuntimeError con prefijo [VI05 step=X] si falla, para identificar
    exactamente dónde se rompe.
    """
    # PASO 0: navegación limpia a VI05 (rompe cualquier cascade de pantalla)
    try:
        _vi05_setup_selection_screen(session)
    except Exception as e:
        raise RuntimeError(f"[step=0 setup VI05] {e}")

    # PASO 1: abrir popup multi-valor de S_FKNUM
    try:
        session.findById("wnd[0]/usr/btn%_S_FKNUM_%_APP_%-VALU_PUSH").press()
        _wait_ready(session)
    except Exception as e:
        raise RuntimeError(f"[step=1 abrir popup S_FKNUM] {e}")

    # PASO 2: limpiar selección previa (btn[16] = Delete Entire Selection)
    try:
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        _wait_ready(session)
    except Exception:
        pass  # no fatal si no hay nada que limpiar

    # PASO 3: ingresar el shpt_cst
    try:
        session.findById(f"{_POPUP_TABLE}/ctxtRSCSEL_255-SLOW_I[1,0]").Text = shpt_cst
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        _wait_ready(session)
    except Exception as e:
        raise RuntimeError(f"[step=3 entrar shpt_cst en popup] {e}")

    # PASO 4: F8 ejecutar
    try:
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        _wait_ready(session)
    except Exception as e:
        raise RuntimeError(f"[step=4 F8 ejecutar] {e}")

    # PASO 5: capturar status bar post-F8 (info, no fatal)
    try:
        sbar_msg = (session.findById("wnd[0]/sbar").Text or "").strip()
        if sbar_msg:
            print(f"  [VI05 {shpt_cst}] sbar post-F8: {sbar_msg}")
            if "no" in sbar_msg.lower() and ("encontr" in sbar_msg.lower() or "found" in sbar_msg.lower()):
                print(f"  [VI05 {shpt_cst}] No encontrado en VI05 — skip")
                return
    except Exception:
        pass

    # PASO 6: leer estado de transferencia (lbl[33,5])
    estado = ""
    try:
        estado = session.findById("wnd[0]/usr/lbl[33,5]").Text.strip()
        print(f"  [VI05 {shpt_cst}] estado={estado!r}")
    except Exception as e:
        print(f"  [VI05 {shpt_cst}] WARN: no se pudo leer lbl[33,5]: {e}")

    # PASO 7: si estado=C, cambiar transferencia antes de eliminar
    if estado == "C":
        try:
            _cambiar_estado_transferencia(session, today)
        except Exception as e:
            raise RuntimeError(f"[step=7 cambiar_estado_transferencia] {e}")

    # PASO 8: eliminar
    try:
        _eliminar_shpt_cst(session, shpt_cst)
    except Exception as e:
        raise RuntimeError(f"[step=8 eliminar_shpt_cst] {e}")


def delete_shipment_costs_all(session, shpt_csts: list) -> None:
    """
    VI05: Elimina shipment costs uno por uno con navegación limpia al inicio
    de cada iteración. Más lento que stay-in-transaction pero robusto frente a
    errores: el fallo de un ticket NO contamina los siguientes.
    """
    if not shpt_csts:
        return

    today = date.today().strftime("%m/%d/%Y")
    _go_back(session, 3)

    for shpt_cst in shpt_csts:
        try:
            _process_single_shpt_cst(session, shpt_cst, today)
        except Exception as e:
            print(f"  [VI05] Error en {shpt_cst}: {e}", file=sys.stderr)


# ---------------------------------------------------------------------------
# Batch 3 — VT02N: Eliminar Shipment Number
# ---------------------------------------------------------------------------

def delete_shipment_number(session, shipment: str) -> None:
    """VT02N: Desasigna y elimina el shipment dado (uso per-ticket)."""
    if not shipment:
        return

    _go_back(session, 3)
    _navigate_to(session, "VT02N")
    _wait_ready(session)

    session.findById("wnd[0]/usr/ctxtVTTK-TKNUM").Text = shipment
    session.findById("wnd[0]").sendVKey(0)
    _wait_ready(session)

    session.findById(f"{_VT02N_HEADER}/btn*RV56A-ICON_STABF").press()
    _wait_ready(session)
    session.findById(f"{_VT02N_HEADER}/btn*RV56A-ICON_STDIS").press()
    _wait_ready(session)

    session.findById("wnd[0]/tbar[1]/btn[7]").press()  # Ver planning screen
    _wait_ready(session)

    session.findById(
        f"{_VT02N_PLANNING}/shellcont[1]/shell[1]"
    ).selectItem("          2", "&Hierarchy")

    session.findById(_VT02N_PLANNING).pressButton("MM_UNAS        10001")
    _wait_ready(session)

    session.findById("wnd[0]/tbar[0]/btn[11]").press()
    _wait_ready(session)
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    _wait_ready(session)


def _process_single_shipment(session, shipment: str) -> None:
    """
    Procesa UN shipment en VT02N con navegación limpia y diagnóstico por paso.
    Lanza RuntimeError con prefijo [VT02N step=X] si falla.
    """
    # PASO 0: navegación limpia a VT02N (rompe cascade)
    try:
        _navigate_to(session, "VT02N")
        _wait_ready(session)
    except Exception as e:
        raise RuntimeError(f"[step=0 navegar VT02N] {e}")

    # PASO 1: ingresar shipment number
    try:
        session.findById("wnd[0]/usr/ctxtVTTK-TKNUM").Text = shipment
        session.findById("wnd[0]").sendVKey(0)
        _wait_ready(session)
    except Exception as e:
        raise RuntimeError(f"[step=1 ingresar TKNUM] {e}")

    # PASO 2: presionar STABF (cancel transp.) y STDIS (cancel plng compl.)
    try:
        session.findById(f"{_VT02N_HEADER}/btn*RV56A-ICON_STABF").press()
        _wait_ready(session)
    except Exception as e:
        raise RuntimeError(f"[step=2a STABF] {e}")
    try:
        session.findById(f"{_VT02N_HEADER}/btn*RV56A-ICON_STDIS").press()
        _wait_ready(session)
    except Exception as e:
        raise RuntimeError(f"[step=2b STDIS] {e}")

    # PASO 3: abrir planning screen (btn[7] = F7)
    try:
        session.findById("wnd[0]/tbar[1]/btn[7]").press()
        _wait_ready(session)
    except Exception as e:
        raise RuntimeError(f"[step=3 btn[7] planning] {e}")

    # PASO 4a: seleccionar nodo "2" en la jerarquía del planning tree
    try:
        tree = session.findById(f"{_VT02N_PLANNING}/shellcont[1]/shell[1]")
        try:
            top_node = tree.GetNodeKeyByPath("          2")
            print(f"  [VT02N {shipment}] tree nodo '2' encontrado, key={top_node!r}")
        except Exception:
            pass
        tree.selectItem("          2", "&Hierarchy")
        _wait_ready(session)
    except Exception as e:
        print(f"  [VT02N {shipment}] step=4a selectItem falló: {e}", file=sys.stderr)
        # Diagnóstico: listar nodos del tree
        try:
            tree = session.findById(f"{_VT02N_PLANNING}/shellcont[1]/shell[1]")
            top_nodes = list(tree.GetAllNodeKeys())
            print(f"  [VT02N {shipment}] nodos disponibles en tree: {top_nodes[:20]}", file=sys.stderr)
        except Exception as ee:
            print(f"  [VT02N {shipment}] no se pudo enumerar tree: {ee}", file=sys.stderr)
        raise RuntimeError(f"[step=4a selectItem '2' &Hierarchy] {e}")

    # PASO 4b: presionar botón Unassign (MM_UNAS        10001)
    # _VT02N_PLANNING es un GuiSplitterShell con 2 GuiContainerShell hijos.
    # Cada container tiene shell[0] (toolbar) y shell[1] (contenido).
    # El tree vive en shellcont[1]/shell[1], así que su toolbar pareado es shellcont[1]/shell[0].
    candidates_press = [
        ("shellcont[1]/shell[0] (toolbar del tree)", f"{_VT02N_PLANNING}/shellcont[1]/shell[0]"),
        ("shellcont[0]/shell[0] (toolbar left)",     f"{_VT02N_PLANNING}/shellcont[0]/shell[0]"),
        ("shellcont[1]/shell[1] (tree)",             f"{_VT02N_PLANNING}/shellcont[1]/shell[1]"),
        ("shellcont[0]/shell[1]",                    f"{_VT02N_PLANNING}/shellcont[0]/shell[1]"),
    ]
    pressed = False
    last_err = None
    for label, ctrl_id in candidates_press:
        try:
            ctrl = session.findById(ctrl_id)
            ctype = getattr(ctrl, "Type", "?")
            print(f"  [VT02N {shipment}] probando pressButton en {label} (Type={ctype})")
            ctrl.pressButton("MM_UNAS        10001")
            _wait_ready(session)
            print(f"  [VT02N {shipment}] pressButton OK en {label}")
            pressed = True
            break
        except Exception as e:
            last_err = e
            continue

    if not pressed:
        print(f"  [VT02N {shipment}] step=4b pressButton falló en TODOS los candidatos", file=sys.stderr)
        # Diagnóstico recursivo: muestra hijos del planning Y de cada container
        def _dump(node_id, prefix=""):
            try:
                node = session.findById(node_id)
            except Exception as e:
                print(f"  {prefix}NO se pudo acceder {node_id}: {e}", file=sys.stderr)
                return
            try:
                ntype = getattr(node, "Type", "?")
                n_child = node.Children.Count
                print(f"  {prefix}{node_id} Type={ntype} Children={n_child}", file=sys.stderr)
                for i in range(min(n_child, 10)):
                    try:
                        c = node.Children.Item(i)
                        cid = getattr(c, "Id", "?")
                        ctype = getattr(c, "Type", "?")
                        print(f"  {prefix}  [{i}] {ctype}  {cid}", file=sys.stderr)
                    except Exception:
                        continue
            except Exception as e:
                print(f"  {prefix}error: {e}", file=sys.stderr)

        _dump(_VT02N_PLANNING, "")
        _dump(f"{_VT02N_PLANNING}/shellcont[0]", "  shellcont[0]: ")
        _dump(f"{_VT02N_PLANNING}/shellcont[1]", "  shellcont[1]: ")
        raise RuntimeError(f"[step=4b pressButton 'MM_UNAS 10001'] último error: {last_err}")

    # PASO 5: save (btn[11])
    try:
        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        _wait_ready(session)
    except Exception as e:
        raise RuntimeError(f"[step=5 save] {e}")

    # PASO 6: confirmar popup (puede o no aparecer)
    try:
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        _wait_ready(session)
    except Exception:
        pass  # no fatal si no hay popup

    # PASO 7: capturar status bar
    try:
        sbar = session.findById("wnd[0]/sbar")
        msg = (sbar.Text or "").strip()
        mtype = (sbar.MessageType or "").strip()
        if msg:
            print(f"  [VT02N {shipment}] SAP ({mtype}): {msg}")
    except Exception:
        pass


def delete_shipment_numbers_all(session, shipments: list) -> None:
    """
    VT02N: Elimina shipment numbers uno por uno con navegación limpia al inicio
    de cada iteración. Más lento pero robusto: errores no contaminan al siguiente.
    """
    if not shipments:
        return

    _go_back(session, 3)

    for shipment in shipments:
        try:
            _process_single_shipment(session, shipment)
        except Exception as e:
            print(f"  [VT02N] Error en {shipment}: {e}", file=sys.stderr)


# ---------------------------------------------------------------------------
# Batch 4 — VL09: Reversar PGI (Goods Issue)
# ---------------------------------------------------------------------------

def reverse_pgi(session, delivery: str) -> None:
    """VL09: Reversa el Post Goods Issue del delivery dado (uso per-ticket)."""
    if not delivery:
        return

    _go_back(session, 3)
    _navigate_to(session, "VL09")
    _wait_ready(session)

    session.findById("wnd[0]/usr/btn%_I_VBELN_%_APP_%-VALU_PUSH").press()
    _wait_ready(session)
    session.findById(f"{_POPUP_TABLE}/ctxtRSCSEL_255-SLOW_I[1,0]").Text = delivery
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    _wait_ready(session)

    session.findById("wnd[0]/tbar[1]/btn[8]").press()  # F8
    _wait_ready(session)

    grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    grid.setCurrentCell(-1, "")
    grid.SelectAll()

    session.findById("wnd[0]/tbar[1]/btn[5]").press()
    _wait_ready(session)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    _wait_ready(session)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    _wait_ready(session)

    session.findById("wnd[0]/tbar[1]/btn[16]").press()  # Refresh lista
    _wait_ready(session)
    _go_back(session, 2)


def reverse_pgi_bulk(session, deliveries: list) -> None:
    """
    VL09: Reversa el PGI de todos los deliveries en una sola transacción.
    Usa el multi-valor del popup y selecciona todo para reversar.
    """
    if not deliveries:
        return

    print(f"  [VL09] Reversando PGI de {len(deliveries)} delivery(s)...")

    _go_back(session, 3)
    _navigate_to(session, "VL09")
    _wait_ready(session)

    session.findById("wnd[0]/usr/btn%_I_VBELN_%_APP_%-VALU_PUSH").press()
    _wait_ready(session)

    _enter_multi_values(session, _POPUP_TABLE, deliveries)

    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    _wait_ready(session)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()  # F8
    _wait_ready(session)

    # Verificar que el grid tiene resultados antes de procesar
    try:
        grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        row_count = grid.RowCount
        print(f"  [VL09] Grid después de F8: {row_count} filas")
        if row_count == 0:
            print("  [VL09] WARNING: grid vacío — VL09 no encontró deliveries para reversar")
            _go_back(session, 2)
            return
    except Exception as e:
        print(f"  [VL09] ERROR accediendo al grid: {e}")
        try:
            sbar = session.findById("wnd[0]/sbar").Text or ""
            if sbar:
                print(f"  [VL09] Status bar: {sbar}")
        except Exception:
            pass
        return

    grid.setCurrentCell(-1, "")
    grid.SelectAll()

    session.findById("wnd[0]/tbar[1]/btn[5]").press()
    _wait_ready(session)
    try:
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        _wait_ready(session)
    except Exception:
        pass
    try:
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        _wait_ready(session)
    except Exception:
        pass

    # Leer status bar para confirmar el resultado
    try:
        sbar_text = session.findById("wnd[0]/sbar").Text or ""
        sbar_type = session.findById("wnd[0]/sbar").MessageType or ""
        if sbar_text:
            print(f"  [VL09] Status bar ({sbar_type}): {sbar_text}")
    except Exception:
        pass

    try:
        session.findById("wnd[0]/tbar[1]/btn[16]").press()  # Refresh lista
        _wait_ready(session)
    except Exception:
        pass
    _go_back(session, 2)
