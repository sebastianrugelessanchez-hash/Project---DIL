# Entendimiento de Módulos — Proyecto DIL

## Arquitectura general

El proyecto automatiza un proceso que antes era manual en SAP: tomar una lista de tickets de transporte y eliminar/revertir todos sus documentos asociados en el orden correcto. Piénsalo como una **cadena de desmontaje** — cada pieza debe quitarse antes de poder quitar la siguiente.

```
credentials.json
      │
      ▼
 sap_login.py ──────────────────────────────────────────────────────┐
      │                                                              │
      │ session (objeto COM de SAP GUI)                             │
      ▼                                                              │
  main.py  ◄── orquestador central                                  │
      │                                                              │
      ├── sap_vl06f.py    (leer datos + eliminar BOL)               │
      ├── sap_batches.py  (VF11 · VI05 · VT02N · VL09)             │
      ├── sap_orders.py   (ZCMR · VA02 · ME22N)                    │
      └── verifications.py (verificar resultados)                   │
                │                                                    │
                └── todos importan sap_utils.py ◄───────────────────┘
```

---

## Módulo 1: `sap_utils.py` — El lenguaje base

Es el módulo más pequeño pero el más fundamental. Todos los demás dependen de él.

```python
_POPUP_TABLE = (
    "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA"
    "/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE"
)
```

Esta constante es la dirección interna del campo de texto dentro del popup de selección múltiple de SAP. En SAP GUI Scripting, cada elemento de pantalla tiene un ID único en forma de ruta, igual que una URL. `wnd[1]` es la ventana emergente (popup), y el resto es la ruta dentro de esa ventana hasta la tabla donde se escriben los valores. Se repite en casi todas las transacciones que aceptan múltiples valores.

```python
def _navigate_to(session, t_code: str) -> None:
    session.findById("wnd[0]/tbar[0]/okcd").Text = t_code
    session.findById("wnd[0]").sendVKey(0)
```

`okcd` es el campo de código de transacción (el cuadro de texto arriba a la izquierda en SAP donde escribes `VL06F`, `VF11`, etc.). `sendVKey(0)` equivale a presionar Enter. Así se navega a cualquier transacción sin tocar el mouse.

```python
def _go_back(session, times: int = 1) -> None:
    for _ in range(times):
        try:
            session.findById("wnd[0]/tbar[0]/btn[3]").press
            time.sleep(0.3)
        except Exception:
            break
```

`btn[3]` en la barra de herramientas principal es siempre el botón F3 (Back/Atrás) en SAP. El `try/except` existe porque si ya estamos en la pantalla inicial, presionar Back puede lanzar una excepción — en ese caso simplemente paramos.

```python
def _wait_ready(session, timeout: float = 10.0) -> None:
    end = time.time() + timeout
    while time.time() < end:
        try:
            if not session.Busy:
                return
        except Exception:
            pass
        time.sleep(0.3)
```

SAP GUI es asíncrono: cuando presionas un botón, SAP procesa en el backend y la pantalla queda "ocupada" (`session.Busy = True`) hasta que termina. Sin esta función, el script lanzaría comandos antes de que SAP terminara de responder, causando errores. Es el equivalente a `await` en programación asíncrona.

```python
def _enter_multi_values(session, table_id: str, values: list) -> None:
    VISIBLE = 8
    field_prefix = f"{table_id}/ctxtRSCSEL_255-SLOW_I[1,"
    for i, value in enumerate(values):
        row = i % VISIBLE
        if i > 0 and row == 0:
            session.findById(table_id).verticalScrollbar.Position = i
        session.findById(f"{field_prefix}{row}]").Text = value
```

El popup de selección múltiple de SAP solo muestra 8 filas a la vez. La lógica es: escribe en fila 0..7, cuando llegas a la fila 8 haces scroll a la posición `i` y empiezas de nuevo desde la fila 0. Así puedes ingresar 100 tickets aunque el popup solo muestre 8 visibles.

---

## Módulo 2: `sap_login.py` — Conexión con SAP

```python
SAP_PATH = r'C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe'
SAP_SYSTEM = "PRD - ECC Production"
CLIENT = "900"
```

Define dónde está instalado SAP y a qué sistema conectar.

```python
def launch_sap(self) -> bool:
    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        ...
        self.app = engine
        return True  # SAP ya estaba abierto
    except Exception:
        pass
    subprocess.Popen(SAP_PATH)  # Abrir SAP desde cero
    return self._wait_for_sap()
```

Primero intenta conectarse a una instancia de SAP GUI ya abierta (para no abrir duplicados). Si no existe, lanza el ejecutable y espera hasta 10 reintentos a que responda.

```python
def login(self) -> bool:
    ...
    subprocess.run(["clip"], input=self.password.encode("utf-16-le"), check=True)
    pwd_field = self.session.findById("wnd[0]/usr/pwdRSYST-BCODE")
    pwd_field.SetFocus()
    win32com.client.Dispatch("WScript.Shell").SendKeys("^v")
    subprocess.run(["clip"], input="".encode("utf-16-le"), check=True)
```

Este bloque es la parte más interesante: SAP bloquea el acceso directo al campo de contraseña por seguridad (no se puede escribir `.Text = password` en campos `pwd`). La solución: copiar la contraseña al portapapeles de Windows (`clip`), hacer foco en el campo, pegar con `Ctrl+V`, y luego limpiar el portapapeles. Al final limpia el portapapeles para no dejar la contraseña en memoria.

```python
def run(self):
    if not self.launch_sap(): return
    if not self.open_session(): return
    self.login()
```

`run()` es el método principal: lanza SAP → abre conexión → hace login. Al final, `self.session` contiene el objeto COM de la sesión activa — ese objeto es lo que se pasa a todos los demás módulos.

---

## Módulo 3: `sap_vl06f.py` — La fuente de datos

Este módulo es crítico porque `read_vl06f_data` es **la primera y más importante operación** de todo el pipeline. Lee de una sola vez todos los datos que los demás batches necesitarán.

```python
def read_vl06f_data(session, tickets: list) -> dict:
```

**Flujo completo:**

**Paso 1 — Navegar a VL06F y limpiar filtros previos:**
```python
_navigate_to(session, "VL06F")
session.findById("wnd[0]/usr/ctxtIT_WADAT-LOW").SetFocus
session.findById("wnd[0]/tbar[1]/btn[14]").press  # Limpiar variante/layout
```
VL06F es el "Monitor de Entregas" de SAP. `btn[14]` en `tbar[1]` es el botón de limpiar filtros para que no queden valores de ejecuciones anteriores.

**Paso 2 — Ingresar todos los tickets en el filtro de selección múltiple:**
```python
session.findById("wnd[0]/usr/btn%_IT_VBELN_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press  # Clear
_enter_multi_values(session, _POPUP_TABLE, tickets)
session.findById("wnd[1]/tbar[0]/btn[8]").press  # Cerrar popup
```
El botón `btn%_IT_VBELN_%_APP_%-VALU_PUSH` es el botón de múltiple selección para el campo "Delivery" (VBELN). Se abren todos los tickets de una vez en lugar de buscar uno por uno.

**Paso 3 — Ejecutar la búsqueda y leer el grid:**
```python
session.findById("wnd[0]/tbar[1]/btn[8]").press  # F8 = Ejecutar
grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
for row in range(row_count - 1):
    vbeln = grid.GetCellValue(row, "VBELN")
    invoice_il = grid.GetCellValue(row, "ZZVBELN_IL") or ""
    billing_doc = grid.GetCellValue(row, "ZZVBELN") or ""
    if invoice_il.startswith("7"):
        billing_doc = ""  # Filtro: no eliminar si invoice empieza con "7"
    data[vbeln] = {
        "billing_doc": billing_doc,
        "shpt_cst":    grid.GetCellValue(row, "ZZFKNUM") or "",
        "shipment":    grid.GetCellValue(row, "ZZTKNUM") or "",
        "wbstk":       grid.GetCellValue(row, "WBSTK")   or "",
        "delivery":    vbeln,
    }
```

El resultado es un diccionario como este:
```python
{
  "T0012345": {
      "billing_doc": "9001234",   # VF11 lo eliminará
      "shpt_cst":    "5006789",   # VI05 lo eliminará
      "shipment":    "0000123",   # VT02N lo eliminará
      "wbstk":       "C",         # VL09 lo revertirá (si != "A")
      "delivery":    "T0012345",
  },
  "T0012346": { ... },
  ...
}
```

Este dict se pasa a `main.py` y **todos los batches leen de él** sin volver a consultar SAP por datos. Es la piedra angular de la optimización.

---

## Módulo 4: `sap_batches.py` — Las 4 operaciones principales

Contiene las funciones que realizan las eliminaciones/reversiones en el orden correcto del proceso DIL.

### Batch 1 — VF11: Eliminar Billing Documents

```python
def delete_billing_documents_bulk(session, billing_docs: list) -> None:
    j = 0
    for doc in billing_docs:
        session.findById(
            f"wnd[0]/usr/tblSAPMV60ATCTRL_ERF_FAKT/ctxtKOMFK-VBELN[0,{j}]"
        ).Text = doc
        if j == 13:
            session.findById("wnd[0]/tbar[1]/btn[7]").press  # Nueva página
            j = 0
        else:
            j += 1
    session.findById("wnd[0]/usr/ctxtRV60A-FKDAT").Text = today
    session.findById("wnd[0]/tbar[0]/btn[11]").press  # Guardar
```

VF11 ("Reversar Facturación") tiene una tabla en pantalla que acepta 14 documentos por página. El índice `j` controla la fila actual (0 a 13). Cuando llega a 13, presiona `btn[7]` para agregar una nueva página de 14 filas y reinicia `j=0`. Así se pueden enviar 100 billing docs en una sola transacción.

### Batch 2 — VI05: Eliminar Shipment Costs

La lógica más compleja del sistema. Cada shipment cost puede estar en estado `"C"` (Completely Transferred, bloqueado) o abierto:

```python
def _process_single_shpt_cst(session, shpt_cst: str, today: str) -> None:
    # Abrir popup, limpiar valor anterior, ingresar nuevo
    session.findById("wnd[1]/tbar[0]/btn[24]").press  # Clear
    session.findById(f"{_POPUP_TABLE}/...").Text = shpt_cst
    session.findById("wnd[0]/tbar[1]/btn[8]").press  # F8

    estado = session.findById("wnd[0]/usr/lbl[33,5]").Text.strip()
    if estado == "C":
        _cambiar_estado_transferencia(session, today)  # Desbloquear primero
    _eliminar_shpt_cst(session)

    session.findById("wnd[0]/tbar[1]/btn[8]").press  # Refresh
    _go_back(session, 1)  # Volver a pantalla de selección (sin salir de VI05)
```

La función `_cambiar_estado_transferencia` desbloquea el documento (desmarca `SLSTOR`, pone fecha actual, guarda). Luego `_eliminar_shpt_cst` lo selecciona y elimina. El `_go_back(session, 1)` al final regresa a la pantalla de selección de VI05 — **sin salir de la transacción** — para procesar el siguiente.

El contenedor `delete_shipment_costs_all` navega a VI05 una sola vez, configura las fechas una sola vez, y luego itera:
```python
def delete_shipment_costs_all(session, shpt_csts: list) -> None:
    _navigate_to(session, "VI05")          # UNA sola vez
    # Configurar fechas Select All (también UNA sola vez)
    for shpt_cst in shpt_csts:
        _process_single_shpt_cst(...)      # vuelve a selección de VI05 al final
```

### Batch 3 — VT02N: Eliminar Shipment Numbers

VT02N es el editor de envíos. Para eliminar un shipment hay que ir a la pestaña de Planning y desasignar las entregas:

```python
def _process_single_shipment(session, shipment: str) -> None:
    session.findById("wnd[0]/usr/ctxtVTTK-TKNUM").Text = shipment
    session.findById("wnd[0]").sendVKey(0)  # Cargar el shipment

    # Presionar los botones de status (requeridos por SAP antes de editar)
    session.findById(f"{_VT02N_HEADER}/btn*RV56A-ICON_STABF").press
    session.findById(f"{_VT02N_HEADER}/btn*RV56A-ICON_STDIS").press

    session.findById("wnd[0]/tbar[1]/btn[7]").press  # Ver planning screen

    # Seleccionar el nivel "2" de la jerarquía y desasignar
    session.findById(f"{_VT02N_PLANNING}/shellcont[1]/shell[1]"
    ).selectItem("          2", "&Hierarchy")
    session.findById(_VT02N_PLANNING).pressButton("MM_UNAS        10001")

    session.findById("wnd[0]/tbar[0]/btn[11]").press  # Guardar
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press  # Confirmar
    _go_back(session, 1)  # Volver a entrada de VT02N
```

El `selectItem("          2", "&Hierarchy")` selecciona el nodo de nivel 2 en el árbol jerárquico del planning (los deliveries asignados). `pressButton("MM_UNAS        10001")` es el código del botón "Unassign" del toolbar de la pantalla de planning.

### Batch 4 — VL09: Reversar PGI

```python
def reverse_pgi_bulk(session, deliveries: list) -> None:
    _navigate_to(session, "VL09")
    session.findById("wnd[0]/usr/btn%_I_VBELN_%_APP_%-VALU_PUSH").press
    _enter_multi_values(session, _POPUP_TABLE, deliveries)  # Todos a la vez
    session.findById("wnd[0]/tbar[1]/btn[8]").press  # F8

    grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    grid.SelectAll()
    session.findById("wnd[0]/tbar[1]/btn[5]").press   # Reversar
    session.findById("wnd[1]/tbar[0]/btn[0]").press   # Confirmar (1)
    session.findById("wnd[1]/tbar[0]/btn[0]").press   # Confirmar (2)
```

VL09 acepta múltiples deliveries en el filtro, muestra todos en un grid, y con `SelectAll()` + btn[5] los reversa todos de una vez. SAP pide dos confirmaciones porque revertir movimiento de mercancías es una operación irreversible.

---

## Módulo 5: `sap_orders.py` — Eliminar los pedidos de compra/venta

El más complejo porque involucra 3 transacciones y 2 tipos de órdenes.

```python
def _read_zcmr_orders(session, tickets: list) -> list:
```

ZCMR es una transacción personalizada (Z = custom) que muestra tickets de la empresa. Tiene un grid principal con una fila por ticket, y dentro de cada fila un **sub-grid colapsado** con las órdenes asociadas.

```python
for i in range(row_count):
    fecha = grid.GetCellValue(i, "TICKET_DATE")
    if not fecha:
        continue
    grid.doubleClickCurrentCell()  # Expandir sub-grid

    sub = session.findById(_ZCMR_SUB_GRID)
    delivery = sub.GetCellValue(j, "DELIVERY")
    order    = sub.GetCellValue(j, "SD_ORDER")
    orders.append({
        "order": order,
        "delivery": delivery,
        "is_intracompany": order[:2] == "47",  # Detectar tipo
    })

    session.findById("wnd[0]/tbar[0]/btn[3]").press  # Colapsar sub-grid
```

Se lee TODO en memoria primero, antes de navegar a VA02 o ME22N. Si no hiciéramos esto, al navegar a VA02 perderíamos el estado del grid de ZCMR (sesión única).

La detección del tipo de orden: `order[:2] == "47"` — si los primeros 2 caracteres del número de orden son "47", es una orden de compra intracompany (ME22N). Si no, es una orden de venta intercompany (VA02).

```python
def _delete_intercompany_order_va02(session, order, delivery):
    ...
    mensaje = session.findById("wnd[1]/usr/txtMESSTXT1").Text
    if mensaje:  # El order tiene líneas
        # Buscar la línea que tiene el PO number == delivery y eliminar solo esa línea
    else:        # El order está vacío
        session.findById("wnd[0]/mbar/menu[0]/menu[11]").Select  # Menú → Eliminar
```

VA02 abre el order y muestra un popup con un mensaje informativo si tiene ítems. Si tiene ítems, buscamos la línea específica que corresponde a este delivery (comparando el campo `BSTKD_E` con el número de delivery) y la eliminamos. Si el order está vacío, lo eliminamos completo desde el menú.

---

## Módulo 6: `verifications.py` — Confirmar que todo salió bien

```python
def verify_billing_documents_bulk(session, tickets: list) -> tuple[list, list]:
    data = read_vl06f_data(session, tickets)
    exitosos = [t for t in tickets if data.get(t, {}).get("billing_doc", "") == ""]
    fallidos  = [t for t in tickets if t not in exitosos]
    return exitosos, fallidos
```

Después de cada batch, vuelve a leer VL06F para todos los tickets de una sola consulta. Si el campo que debía vaciarse está vacío (`== ""`), el ticket fue exitoso. Si sigue con valor, falló. Las 4 funciones bulk siguen exactamente el mismo patrón, cambiando solo el campo que verifican:

| Función | Campo verificado | Condición de éxito |
|---------|-----------------|-------------------|
| `verify_billing_documents_bulk` | `billing_doc` | `== ""` |
| `verify_shipment_costs_bulk` | `shpt_cst` | `== ""` |
| `verify_shipment_numbers_bulk` | `shipment` | `== ""` |
| `verify_pgi_reversed_bulk` | `wbstk` | `== "A"` |

---

## Módulo 7: `main.py` — El director de orquesta

Aquí se define el **orden del proceso** y el **flujo de control**.

```python
vl06f = read_vl06f_data(session, tickets)
tickets_activos = [t for t in tickets if t in vl06f]
```

Primero lee todo de VL06F. `tickets_activos` empieza con todos los tickets encontrados.

Luego para cada batch, el patrón es siempre el mismo:
```python
# 1. Filtrar: solo los que tienen el documento a eliminar
docs = [vl06f[t]["campo"] for t in tickets_activos if vl06f[t]["campo"]]

# 2. Operar: ejecutar la eliminación
delete_X(session, docs)

# 3. Verificar: re-leer VL06F y clasificar
exitosos, fallidos = verify_X_bulk(session, tickets_activos)
resultados["BATCH X"] = (exitosos, fallidos)

# 4. Filtrar: solo los exitosos continúan al siguiente batch
tickets_activos = exitosos
```

El punto 4 es clave: **si un ticket falla en el Batch 1, no entra al Batch 2**. Esto evita que un batch posterior intente operar sobre un ticket cuyo estado es inconsistente.

```python
def print_report(resultados: dict, total_tickets: int) -> None:
    tickets_con_fallo = set()
    for batch_name, (exitosos, fallidos) in resultados.items():
        tickets_con_fallo.update(fallidos)
    completados = total_tickets - len(tickets_con_fallo)
```

El reporte al final agrega todos los fallidos de todos los batches en un set (sin duplicados) y muestra cuántos tickets completaron el proceso 100%.

---

## Flujo completo de punta a punta (100 tickets)

```
INICIO
  │
  ├─ sap_login.py → conectar SAP, obtener session
  │
  ├─ sap_vl06f.read_vl06f_data()
  │     └─ 1 consulta VL06F → dict con datos de los 100 tickets
  │
  ├─ BATCH 1 (VF11)
  │     ├─ sap_batches.delete_billing_documents_bulk()  ← 1 navegación
  │     └─ verifications.verify_billing_documents_bulk() ← 1 VL06F
  │
  ├─ BATCH 2 (VI05)
  │     ├─ sap_batches.delete_shipment_costs_all()  ← 1 navegación a VI05
  │     │     └─ _process_single_shpt_cst() × N  ← sin salir de VI05
  │     └─ verifications.verify_shipment_costs_bulk() ← 1 VL06F
  │
  ├─ BATCH 3 (VT02N)
  │     ├─ sap_batches.delete_shipment_numbers_all()  ← 1 navegación a VT02N
  │     │     └─ _process_single_shipment() × N  ← sin salir de VT02N
  │     └─ verifications.verify_shipment_numbers_bulk() ← 1 VL06F
  │
  ├─ BATCH 4 (VL09)
  │     ├─ sap_batches.reverse_pgi_bulk()  ← 1 navegación
  │     └─ verifications.verify_pgi_reversed_bulk() ← 1 VL06F
  │
  ├─ PASO 6 (VL06F delete)
  │     └─ sap_vl06f.delete_bol() × N  ← per-ticket (no existe bulk en SAP)
  │
  ├─ PASO 7 (ZCMR → VA02 / ME22N)
  │     ├─ sap_orders._read_zcmr_orders()  ← leer todo en memoria
  │     └─ VA02 o ME22N × N  ← una por order
  │
  └─ print_report() → resumen en consola
```

El diseño garantiza que si SAP tiene un error en cualquier punto, el ticket afectado queda en `fallidos` y el proceso continúa con los demás — nunca falla todo el lote por un ticket problemático.
