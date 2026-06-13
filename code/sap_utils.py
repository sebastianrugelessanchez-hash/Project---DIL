import time

import win32clipboard


_POPUP_TABLE = (
    "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA"
    "/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE"
)


def _normalize_ticket(value) -> str:
    """
    Canonicaliza un ticket / VBELN para comparación consistente:
    - strip de whitespace
    - lstrip de ceros a la izquierda si es numérico
    - asegura string

    Permite que '44688706' (Excel) y '0044688706' (VL06F) coincidan.
    """
    if value is None:
        return ""
    s = str(value).strip()
    if s.isdigit():
        s = s.lstrip("0") or "0"
    return s


def _navigate_to(session, t_code: str) -> None:
    prefix = "" if t_code.startswith("/") else "/n"
    session.findById("wnd[0]/tbar[0]/okcd").Text = prefix + t_code
    session.findById("wnd[0]").sendVKey(0)


def _go_back(session, times: int = 1) -> None:
    for _ in range(times):
        try:
            session.findById("wnd[0]/tbar[0]/btn[3]").press()
            time.sleep(0.3)
        except Exception:
            break


def _wait_ready(session, timeout: float = 10.0) -> None:
    end = time.time() + timeout
    while time.time() < end:
        try:
            if not session.Busy:
                return
        except Exception:
            pass
        time.sleep(0.3)


def _wnd_exists(session, wnd_id: str) -> bool:
    try:
        session.findById(wnd_id)
        return True
    except Exception:
        return False


def _find_popup_wnd(session) -> str:
    """
    Return the window path of the multi-value selection popup after a %_VALU_PUSH press.
    Checks wnd[2] first so it works when wnd[1] is already occupied (e.g. btn[14] dialog).
    """
    for n in (2, 1):
        try:
            session.findById(f"wnd[{n}]/usr/tabsTAB_STRIP")
            return f"wnd[{n}]"
        except Exception:
            continue
    raise RuntimeError("No se abrió el popup de selección múltiple (tabsTAB_STRIP no encontrado)")


def _popup_table(wnd: str) -> str:
    """Build the full table path for a multi-value popup in the given window."""
    return (
        f"{wnd}/usr/tabsTAB_STRIP/tabpSIVA"
        "/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE"
    )


def _set_clipboard_text(text: str) -> None:
    """Copia texto al clipboard de Windows (formato CF_UNICODETEXT)."""
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
    finally:
        win32clipboard.CloseClipboard()


def _popup_table_row_count(session, table_id: str) -> int:
    """Cuenta cuántas filas con valor tiene el popup multi-valor."""
    ctxt_prefix = table_id + "/ctxtRSCSEL_255-SLOW_I[1,"
    txt_prefix  = table_id + "/txtRSCSEL_255-SLOW_I[1,"
    count = 0
    for row in range(50):
        for prefix in (ctxt_prefix, txt_prefix):
            try:
                val = session.findById(prefix + f"{row}]").Text or ""
                if val.strip():
                    count += 1
                break
            except Exception:
                continue
    return count


def _enter_multi_values(session, table_id: str, values: list) -> None:
    """
    Llena un popup multi-valor de SAP con N valores usando clipboard + botón
    'Upload from Clipboard' (btn[24], confirmado vía diagnose_popup_buttons.py).

    Imprime diagnóstico: cuántos valores intentamos subir vs cuántos quedan
    realmente en el popup después del upload.
    """
    if not values:
        return

    n_values = len(values)
    print(f"  [MultiValues] Subiendo {n_values} valores al popup vía clipboard (btn[24])...")

    # 1. Copiar valores al clipboard (uno por línea, formato Windows \r\n).
    #    Setea AMBOS formatos (CF_UNICODETEXT y CF_TEXT) para máxima compatibilidad SAP.
    text = "\r\n".join(str(v) for v in values)
    try:
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
            try:
                win32clipboard.SetClipboardText(text.encode("latin-1", errors="ignore"), win32clipboard.CF_TEXT)
            except Exception:
                pass
        finally:
            win32clipboard.CloseClipboard()
    except Exception as e:
        print(f"  [MultiValues] ERROR setting clipboard: {e}")
        _enter_multi_values_legacy(session, table_id, values)
        return

    # 2. Identificar la ventana del popup (wnd[1] o wnd[2])
    popup_wnd = table_id.split("/usr/")[0]

    # 3. btn[24] = "Upload from Clipboard (Shift+F12)" en este SAP.
    uploaded = False
    try:
        session.findById(f"{popup_wnd}/tbar[0]/btn[24]").press()
        _wait_ready(session)
        uploaded = True
    except Exception as e:
        print(f"  [MultiValues] btn[24] falló: {e}")

    # 4. FALLBACK: atajo de teclado Shift+F12 si btn[24] falló
    if not uploaded:
        try:
            session.findById(popup_wnd).sendVKey(24)
            _wait_ready(session)
            uploaded = True
        except Exception:
            pass

    # 5. NO contamos filas del popup: es virtualizado (solo ~8 visibles), el conteo
    # daría falsos positivos. La verificación real del upload se hace por el
    # RowCount del grid después de F8 (ver read_vl06f_data).
    if uploaded:
        return

    # 6. ÚLTIMO RECURSO: método legacy con scroll (solo confiable para <=16 valores)
    print(f"  [MultiValues] Upload falló — usando método legacy con scroll")
    _enter_multi_values_legacy(session, table_id, values)


def _enter_multi_values_legacy(session, table_id: str, values: list) -> None:
    """
    Método anterior: llena la tabla manualmente scrolleando cada 8 filas.
    Solo es confiable para <= ~16 valores. Se mantiene como fallback.
    """
    VISIBLE = 8

    # Detectar tipo de campo: ctxt (char fields) o txt (numeric fields)
    ctxt_prefix = table_id + "/ctxtRSCSEL_255-SLOW_I[1,"
    txt_prefix  = table_id + "/txtRSCSEL_255-SLOW_I[1,"

    try:
        session.findById(ctxt_prefix + "0]")
        field_prefix = ctxt_prefix
    except Exception:
        field_prefix = txt_prefix

    for i, value in enumerate(values):
        row = i % VISIBLE
        if i > 0 and row == 0:
            session.findById(table_id).verticalScrollbar.Position = i
        session.findById(field_prefix + str(row) + "]").Text = value
