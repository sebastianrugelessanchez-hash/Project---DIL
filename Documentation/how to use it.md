# How to Use It

Pasos para ejecutar el pipeline de automatización SAP DIL.

---

## Prerequisitos

- Anaconda instalado con el entorno `mi_entorno` (Python 3.10.18)
- SAP GUI instalado en `C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe`
- Archivo `credentials.json` dentro de la carpeta `code/` con tu usuario y contraseña SAP
- Archivo Excel de billing ubicado en `Data-bases/Entradas/<subcarpeta-del-mes>/`
- Carpeta de salida creada en `Data-bases/Salidas/<subcarpeta-del-mes>/`

---

## Estructura de carpetas esperada

```
Proyecto DIL/
├── code/
│   ├── main.py
│   ├── credentials.json       ← credenciales SAP (nunca compartir)
│   └── rules.json             ← reglas de negocio del pipeline
├── Data-bases/
│   ├── Entradas/
│   │   └── May-2026/
│   │       └── _Billing USA - May 13 2026.xlsx   ← archivo de tickets
│   ├── Salidas/
│   │   └── May-2026/          ← el reporte xlsx se genera aquí
│   └── Estado/                ← state JSON para resume (auto-creada)
└── Documentation/
```

---

## Paso 1 — Abrir SAP GUI manualmente

Abre SAP GUI e inicia sesión con tu usuario. El pipeline necesita que SAP esté abierto y con al menos una sesión activa antes de correr.

---

## Paso 2 — Abrir Anaconda Prompt

Busca **Anaconda Prompt** en el menú de inicio de Windows y ábrelo.

---

## Paso 3 — Activar el entorno

```bash
conda activate mi_entorno
```

---

## Paso 4 — Navegar a la carpeta del proyecto

```bash
cd "C:\Users\srugeles\Desktop\Proyecto DIL\code"
```

---

## Paso 5 — Ejecutar el pipeline

```bash
python main.py
```

---

## Qué ocurre al ejecutarlo

1. **Lectura del Excel** — el pipeline busca automáticamente el archivo `.xlsx` más reciente en `Data-bases/Entradas/<subcarpeta-más-reciente>/` y extrae los tickets de la columna ZCMR
2. **Login SAP** — se conecta a la sesión SAP abierta e intenta abrir 3 sesiones paralelas
3. **Checkpoint** — carga el state previo (si existe) o crea uno nuevo. Si el Excel cambió, el state previo se descarta automáticamente
4. **Chunking** — los tickets se parten en chunks de **100** (configurable). Cada chunk se procesa COMPLETO por todos los batches antes del siguiente:

   Para cada chunk de 100 tickets:

   - **Lectura VL06F** — carga datos del chunk (billing doc, shipment cost, shipment number, delivery, WBSTK)
   - **Batch 1 — Billing Documents** (VF11): elimina billing docs y verifica en VL06F
   - **Batch 2 — Shipment Cost** (VI05): elimina shipment costs y verifica
   - **Batch 3 — Shipment Number** (VT02N): elimina shipment numbers y verifica
   - **Batch 4 — Reverse PGI** (VL09): reversa PGI y verifica
   - **Batch 5 — BOL Deletion** (VL06F): elimina el BOL por ticket
   - **Batch 6 — ZCMR Orders** (ZCMR + VA02/ME22N): elimina las órdenes
   - **Batch 7 — Verificación final** (ZSD_DEL_TICKETS): confirma que no quedan órdenes
   - **Batch 8 — Order Cancellation retry** (VA02): cancela las que quedaron pendientes
   - **Checkpoint** — guarda el resultado del chunk en `Data-bases/Estado/{nombre_excel}_state.json`

5. **Reporte consolidado** — al terminar TODOS los chunks, imprime resumen en consola y guarda un `.xlsx` en `Data-bases/Salidas/<subcarpeta-más-reciente>/` con los datos acumulados de todos los chunks

> **Regla importante:** si un ticket falla en cualquier batch, no avanza a los siguientes batches DENTRO de su chunk, pero el pipeline continúa procesando los demás tickets sin detenerse.

> **Si la corrida se interrumpe** (timeout, kill, crash), al volver a correr `python main.py` se reanuda desde el último chunk no completado. Los chunks anteriores no se re-procesan.

---

## Reporte de salida

Al finalizar verás en consola:

```
=============================================
  REPORTE DIL — 2026-05-13
=============================================
  Tickets procesados: 45

  BATCH 1 — Billing Documents
    Exitosos (45): 70811039, 70811040, ...
    Fallidos  (0): —
  ...
=============================================
  Completados al 100%:  43/45
  Con fallos parciales:  2/45
  Revisar manualmente:  70811041, 70811055
=============================================

  Reporte xlsx guardado en: Data-bases\Salidas\May-2026\Reporte DIL - 2026-05-13.xlsx
```

El archivo Excel de salida tiene dos hojas:
- **Resumen** — conteos de exitosos/fallidos por batch
- **Fallos** — lista de tickets fallidos con el batch donde ocurrió el error

---

## Argumentos del pipeline

`python main.py` acepta los siguientes flags. Todos son opcionales.

| Flag | Default | Descripción |
|------|---------|-------------|
| `--batches` | `all` | Subset de batches a correr. Formato: `1,2,3` o `6-8` o `1,4,6-8` |
| `--chunk-size` | `100` | Tickets por chunk. Límite práctico VL09 es 60-100 |
| `--fresh` | `false` | Elimina state previo antes de arrancar (forzar fresh run) |
| `--report-only` | `false` | NO toca SAP. Solo regenera el reporte xlsx desde state existente |
| `--retry-failed` | `false` | Reintenta SOLO los tickets que quedaron como fallidos en el state previo. Reusa vl06f cacheado (no relee VL06F). Cascada completa: si un batch pasa en retry, encadena los siguientes |

### Sintaxis de `--batches`

```bash
python main.py --batches <especificación>
```

| Formato | Ejemplo | Qué hace |
|---------|---------|----------|
| `all` | `--batches all` | Todos los batches (igual al default) |
| Lista | `--batches 1,2,3` | Solo los batches indicados |
| Rango | `--batches 6-8` | Desde el primero hasta el último (inclusivo) |
| Mixto | `--batches 1,4,6-8` | Combinación de listas y rangos |
| Único | `--batches 5` | Un solo batch |

> **Batches válidos:** 1, 2, 3, 4, 5, 6, 7, 8.

### Casos de uso comunes

**Corrida normal (todos los tickets, todos los batches, chunks de 100):**

```bash
python main.py
```

**Re-procesar órdenes pendientes** (cuando los tickets ya se borraron de VL06F pero quedan órdenes en SAP):

```bash
python main.py --batches 6-8
```

Como ningún batch 1-5 está seleccionado, el pipeline **omite la lectura de VL06F** y pasa los tickets del Excel directamente a Batch 6.

**Solo eliminar Shipment Cost y Shipment Number:**

```bash
python main.py --batches 2,3
```

**Solo verificación final con cancelación de órdenes:**

```bash
python main.py --batches 7,8
```

**Chunks más pequeños** (si SAP está lento o aparecen timeouts):

```bash
python main.py --chunk-size 50
```

**Ignorar progreso previo y arrancar desde cero:**

```bash
python main.py --fresh
```

**Solo regenerar el reporte xlsx desde el state actual** (NO toca SAP, útil después de inspeccionar manualmente o reanudar reporte sin re-procesar):

```bash
python main.py --report-only
```

**Reintentar solo los tickets fallidos del state previo** (sin reprocesar los ya exitosos):

```bash
python main.py --retry-failed
```

Recorre los chunks ya completados del state, identifica los tickets que quedaron como **fallidos** en cualquier batch, y los reintenta. Reusa el `vl06f` cacheado en el state (no vuelve a leer VL06F en SAP). Si un batch pasa en el retry, los siguientes batches también se intentan sobre esos tickets (cascada completa).

**Reintentar solo batches específicos:**

```bash
python main.py --retry-failed --batches 5,6     # solo BOL + ZCMR Orders
python main.py --retry-failed --batches 7-8     # solo verificación y cancelación
```

**Combinaciones:**

```bash
python main.py --batches 6-8 --fresh             # subset + fresh
python main.py --chunk-size 50 --batches 1-5     # solo eliminación con chunks pequeños
python main.py --retry-failed --batches 5,6      # retry quirúrgico de B5 y B6
```

> **Conflictos:** `--retry-failed` NO se puede combinar con `--fresh` (se contradicen: uno reusa state, el otro lo borra) ni con `--report-only` (uno re-ejecuta SAP, el otro no). El pipeline aborta con error claro si se pasan juntos.

### Comportamiento importante

| Caso | Qué hace el pipeline |
|------|----------------------|
| No seleccionas ningún batch 1-5 | **Omite la lectura de VL06F** y usa todos los tickets del Excel directamente |
| Seleccionas algún batch 1-5 | Lee VL06F normalmente (necesario para los números de documento) |
| Seleccionas Batch 8 sin Batch 7 | Se omite con aviso (necesita las órdenes detectadas por Batch 7) |
| Batch inválido (ej: `--batches 9`) | Error claro mostrando los batches válidos (1-8) |

### Reporte cuando se corre parcialmente

El xlsx generado **solo incluye sheets de los batches que realmente se ejecutaron**. Si haces `--batches 6-8`, el reporte solo mostrará sheets de Batch 6, 7 y 8 (más el "Resumen").

---

## Resume y checkpointing

El pipeline guarda el progreso por chunk en `Data-bases/Estado/{nombre_excel}_state.json`. Esto permite:

### Si la corrida se interrumpe

```bash
python main.py                # arranca, procesa CHUNK 1, llega a CHUNK 2 y se interrumpe (Ctrl+C, timeout, crash)
python main.py                # vuelve a correr — el state detecta que CHUNK 1 ya está done y arranca desde CHUNK 2
```

Verás en consola:
```
[Checkpoint] State previo encontrado: 1/3 chunks ya completados. Reanudando.
CHUNK 1/3 ya completado el 2026-05-22T10:30:00 — saltando.
CHUNK 2/3 ... (procesa)
CHUNK 3/3 ... (procesa)
```

### Si modificas el Excel entre corridas

El pipeline detecta el cambio por hash SHA256 y **automáticamente descarta el state previo**:

```
[Checkpoint] Input file cambió (hash mismatch). State descartado.
[Checkpoint] Nuevo state inicializado en Data-bases/Estado/
CHUNK 1/...  ← arranca desde cero
```

### Si quieres forzar fresh run sin importar el state

```bash
python main.py --fresh
```

Borra el state file y arranca desde CHUNK 1.

### Si solo quieres regenerar el reporte sin tocar SAP

Útil si:
- Ya completaste el procesamiento y quieres regenerar el xlsx
- Quieres inspeccionar el state sin re-ejecutar

```bash
python main.py --report-only
```

Lee el state, consolida todos los chunks y emite el reporte xlsx. **No hace login a SAP.**

### Si quedaron tickets fallidos y quieres reintentar SOLO esos

`--retry-failed` está diseñado para el caso en que la corrida normal terminó "exitosa" pero algunos batches (típicamente Batch 5 BOL o Batch 6 Orders) dejaron tickets en la lista de fallidos. Sin esta bandera, re-correr `python main.py` saltaba esos chunks porque ya estaban marcados como completados.

```bash
python main.py --retry-failed
```

**Qué hace exactamente:**

1. Carga el state previo (detecta `mode` automáticamente: normal o manual).
2. Recorre todos los chunks ya completados y arma `{batch_id: [tickets fallidos]}`.
3. Para cada chunk con fallidos, arranca el retry en el **batch más temprano** que tenga fallidos y encadena los batches siguientes en cascada.
4. **NO vuelve a leer VL06F desde SAP** — usa los datos cacheados en `state["chunks"][i]["vl06f"]`.
5. Mergea el resultado: los tickets que ahora pasan se mueven de `fallidos` a `exitosos` en el state. Los que siguen fallando se mantienen.
6. Persiste el chunk actualizado y emite reporte consolidado final.

**Cuándo usarlo:**
- Después de una corrida con fallos parciales en Batch 5/6/7/8 (no quieres reprocesar los 100 tickets del chunk, solo los 3-4 que fallaron).
- Cuando SAP estaba lento/con locks y querés un segundo intento sin gastar tiempo en lo ya hecho.
- Como alternativa a `--fresh` cuando lo que falló es una porción pequeña.

**Cuándo NO usarlo:**
- Si quieres reprocesar TODO desde cero — usa `--fresh`.
- Si el state no existe o el Excel cambió (hash mismatch) — el retry abortará con error.
- Si quieres ejecutar batches que nunca corrieron en la corrida original — usa `--batches X` sin `--retry-failed`.

**Ejemplo de salida:**

```
  [Retry] State detectado en modo 'normal' con 3/3 chunks completados.
  [Retry] 2 chunk(s) con fallidos a reintentar: [0, 2]

████████████████████████████████████████████████████████████
  RETRY CHUNK 1/3 — 100 tickets originales
████████████████████████████████████████████████████████████
    Reintentando desde Batch 5 con 3 tickets.
  Batch 5 (retry): BOL Deletion (per-ticket)...
  ...
  [Checkpoint] Chunk 1/3 actualizado (45s).
```

### Estructura del state file

El state es un JSON human-readable que puedes abrir para inspeccionar:

```json
{
  "version": 1,
  "run_id": "2026-05-22T10:00:00_pid12345",
  "input_file": "Billing USA -May 22 2026.xlsx",
  "input_file_hash": "sha256:abc123...",
  "chunk_size": 100,
  "total_tickets": 250,
  "total_chunks": 3,
  "chunks": {
    "0": { "duration_seconds": 1820, "tickets": [...], "resultados": {...}, ... },
    "1": { ... },
    "2": { ... }
  }
}
```

---

## Reglas de negocio

Las reglas están documentadas en `code/rules.json`. Las más importantes:

| ID | Regla |
|----|-------|
| R001 | Si `ZZVBELN_IL` en VL06F empieza por `"7"`, el ticket tiene factura y **no se procesa** en Batch 1 |
| R002 | Solo los tickets que pasan **todos** los batches anteriores avanzan al Paso 7 (ZCMR) |

---

## Sesiones SAP

El pipeline intenta abrir 3 sesiones simultáneas:

| Sesión | Responsabilidad |
|--------|----------------|
| session1 | VL06F — lectura y verificaciones |
| session2 | Operaciones — VF11, VI05, VT02N, VL09 |
| session3 | Órdenes — ZCMR, VA02, ME22N |

Si SAP no permite abrir sesiones adicionales (límite de sesiones ocupado), el pipeline continúa con 1 sola sesión mostrando una advertencia.

---

## Solución de problemas

| Error / Síntoma | Causa | Solución |
|-----------------|-------|----------|
| `No se encontraron archivos .xlsx en:` | La carpeta de Entradas está vacía o no tiene subcarpeta | Crear la subcarpeta del mes y copiar el Excel |
| `Se esperaba 'ZCMR' en fila 3 columna C` | El formato del Excel cambió | Verificar que la columna C fila 3 tenga el encabezado `ZCMR` |
| `No se pudo establecer sesión con SAP` | SAP GUI no está abierto | Abrir SAP GUI manualmente antes de correr el pipeline |
| `ADVERTENCIA: No se pudieron abrir las 3 sesiones` | SAP tiene el límite de sesiones ocupado | Cerrar sesiones SAP innecesarias y volver a ejecutar |
| `ModuleNotFoundError: win32com` | `pywin32` no instalado en el entorno | `pip install pywin32` con `mi_entorno` activo |
| `[Reporte] No se pudo guardar el xlsx` | La carpeta `Salidas/<mes>/` no existe | Crear la carpeta de salida del mes correspondiente |
| `[CHUNK N] Error crítico: TIME_OUT` | VL09 superó el work process timeout (>10 min) | Reducir `--chunk-size` (ej: 50). El chunk fallido se reintenta al re-correr |
| `[Checkpoint] Input file cambió (hash mismatch)` | Modificaste el Excel entre corridas | Comportamiento esperado: descarta state y arranca fresh |
| Pipeline reprocesa chunks ya hechos | El state file fue borrado o el Excel cambió | Si fue accidental, verifica `Data-bases/Estado/`. Si fue intencional, OK |
| Quieres regenerar solo el reporte xlsx | No quieres volver a tocar SAP | `python main.py --report-only` |
| Quieres limpiar el progreso y arrancar de cero | Estado previo no es útil | `python main.py --fresh` o borrar el `_state.json` manualmente |
| Batch 5/6 dejó fallidos y `python main.py` no los reintenta | El chunk se marcó como completo aunque hubo fallos parciales | `python main.py --retry-failed` |
| `Error: --retry-failed requiere un state previo válido` | No hay state, o el Excel cambió (hash mismatch) | Correr `python main.py` primero, o restaurar el Excel original |
| `Error: --retry-failed y --fresh son mutuamente excluyentes` | Combinaste banderas contradictorias | Decidir: `--fresh` reprocesa todo desde cero; `--retry-failed` solo los fallidos |
| `No hay tickets fallidos en el state ... Nada que reintentar.` | El state ya está 100% exitoso para los batches seleccionados | No hay nada que hacer — el retry termina sin tocar SAP |
