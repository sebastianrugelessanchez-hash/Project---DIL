"""
Diagnóstico: identificar el botón "Upload from Clipboard" en el popup multi-valor de VL06F.

Pasos:
  1. Se conecta a la sesión SAP que ya esté abierta (debes tener SAP abierto y logueado)
  2. Navega a VL06F
  3. Abre el popup multi-valor del campo Delivery (IT_VBELN)
  4. Imprime TODOS los botones del toolbar del popup con su tooltip
  5. Imprime opciones del menú (mbar) si hay
  6. Cierra el popup sin hacer cambios

Correr con UNA sola sesión SAP abierta y SAP en pantalla inicial:
    python diagnose_popup_buttons.py
"""
import sys
import time

import win32com.client


def wait_ready(session, timeout: float = 10.0) -> None:
    end = time.time() + timeout
    while time.time() < end:
        try:
            if not session.Busy:
                return
        except Exception:
            pass
        time.sleep(0.2)


def wnd_exists(session, wnd_id: str) -> bool:
    try:
        session.findById(wnd_id)
        return True
    except Exception:
        return False


def get_session():
    """Encuentra la primera sesión SAP que no esté en login screen."""
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    app = sap_gui_auto.GetScriptingEngine
    for i in range(app.Children.Count):
        conn = app.Children(i)
        for j in range(conn.Children.Count):
            candidate = conn.Children(j)
            try:
                candidate.findById("wnd[0]/usr/txtRSYST-BNAME")
                continue  # esta sesión está en login, saltar
            except Exception:
                return candidate
    return None


def enumerate_toolbar(session, base: str, label: str) -> None:
    """Imprime los botones disponibles en un toolbar (tbar[0], tbar[1])."""
    print(f"\n  -- {label} --")
    encontrados = 0
    for btn_id in range(50):
        try:
            btn = session.findById(f"{base}/btn[{btn_id}]")
            tip  = (getattr(btn, "Tooltip", "") or "").strip()
            text = (getattr(btn, "Text", "") or "").strip()
            print(f"    btn[{btn_id:2d}]  tooltip={tip!r:60s}  text={text!r}")
            encontrados += 1
        except Exception:
            continue
    if encontrados == 0:
        print(f"    (sin botones)")


def enumerate_menubar(session, wnd_id: str) -> None:
    """Imprime los items del menubar del popup (si existe)."""
    print(f"\n  -- {wnd_id}/mbar --")
    try:
        mbar = session.findById(f"{wnd_id}/mbar")
    except Exception:
        print("    (no hay mbar)")
        return

    try:
        for i in range(mbar.Children.Count):
            menu = mbar.Children(i)
            try:
                name = menu.Text or menu.Name
            except Exception:
                name = "?"
            print(f"    menu[{i}]: {name!r}")
            # listar sub-items del menu (1 nivel)
            try:
                for j in range(menu.Children.Count):
                    sub = menu.Children(j)
                    try:
                        sname = sub.Text or sub.Name
                    except Exception:
                        sname = "?"
                    print(f"      menu[{i}]/menu[{j}]: {sname!r}")
            except Exception:
                pass
    except Exception as e:
        print(f"    Error iterando mbar: {e}")


def main() -> None:
    session = get_session()
    if not session:
        print("ERROR: no se encontró sesión SAP activa.")
        print("Asegúrate de tener SAP abierto y logueado antes de correr este script.")
        return

    try:
        print(f"Sesión SAP encontrada: '{session.findById('wnd[0]').Text}'")
    except Exception:
        print("Sesión SAP encontrada.")

    # Cerrar diálogos abiertos por si acaso
    for w in ("wnd[2]", "wnd[1]"):
        if wnd_exists(session, w):
            try:
                session.findById(w).sendVKey(12)  # F12 = Cancel
                wait_ready(session)
            except Exception:
                pass

    # ---- 1. Navegar a VL06F ----
    print("\n[1] Navegando a /nVL06F...")
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nVL06F"
    session.findById("wnd[0]").sendVKey(0)
    wait_ready(session)
    try:
        print(f"    Pantalla: '{session.findById('wnd[0]').Text}'")
    except Exception:
        pass

    # ---- 2. Abrir popup multi-valor de Delivery (IT_VBELN) ----
    print("\n[2] Abriendo popup multi-valor de IT_VBELN (Delivery)...")
    btn_path = "wnd[0]/usr/btn%_IT_VBELN_%_APP_%-VALU_PUSH"
    if not wnd_exists(session, btn_path):
        print(f"    ERROR: no se encontró el botón {btn_path}")
        print("    ¿Estás seguro de que VL06F cargó correctamente?")
        return

    session.findById(btn_path).press()
    wait_ready(session)

    # Identificar la ventana del popup
    popup_wnd = None
    for n in (2, 1):
        if wnd_exists(session, f"wnd[{n}]/usr/tabsTAB_STRIP"):
            popup_wnd = f"wnd[{n}]"
            break

    if not popup_wnd:
        print("    ERROR: el popup no se abrió o no tiene tabsTAB_STRIP")
        return

    try:
        print(f"    Popup abierto en: {popup_wnd}")
        print(f"    Título: '{session.findById(popup_wnd).Text}'")
    except Exception:
        pass

    # ---- 3. Enumerar TODOS los botones del popup ----
    print("\n[3] BOTONES del popup:")
    enumerate_toolbar(session, f"{popup_wnd}/tbar[0]", f"{popup_wnd}/tbar[0]")
    enumerate_toolbar(session, f"{popup_wnd}/tbar[1]", f"{popup_wnd}/tbar[1]")

    # ---- 4. Enumerar menúes ----
    print("\n[4] MENÚES del popup:")
    enumerate_menubar(session, popup_wnd)

    # ---- 5. Listar tabs del TAB_STRIP ----
    print("\n[5] TABS del TAB_STRIP:")
    try:
        tab_strip = session.findById(f"{popup_wnd}/usr/tabsTAB_STRIP")
        for i in range(tab_strip.Children.Count):
            child = tab_strip.Children(i)
            try:
                name = child.Name
                text = (getattr(child, "Text", "") or "").strip()
                print(f"    tab[{i}]: name={name!r}  text={text!r}")
            except Exception:
                pass
    except Exception as e:
        print(f"    Error iterando tabs: {e}")

    # ---- 6. Cerrar popup sin hacer cambios ----
    print("\n[6] Cerrando popup (F12 = Cancel)...")
    try:
        session.findById(popup_wnd).sendVKey(12)
        wait_ready(session)
    except Exception:
        pass

    print("\nDiagnóstico completado. Busca en la lista de arriba el botón con tooltip")
    print("'Upload from clipboard', 'Importar desde portapapeles' o similar.")
    print("Comparte el btn[N] correspondiente conmigo para actualizar el código.")


if __name__ == "__main__":
    main()
