# Architecture — Proyecto DIL

## Sesiones SAP

El pipeline abre **3 sesiones SAP simultáneas** desde la misma conexión, cada una con una responsabilidad exclusiva. Esto evita que las operaciones de eliminación interrumpan las lecturas de VL06F.

```
┌─────────────────────────────────────────────────────────────────┐
│  session1 — VL06F                                               │
│  Responsabilidad: lectura de datos y verificaciones             │
│  Transacciones: VL06F                                           │
│  Cuándo actúa: al inicio (lectura) y después de cada batch      │
│                (verificación). También ejecuta el Paso 6 (BOL). │
└─────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────┐
│  session2 — Operaciones                                         │
│  Responsabilidad: eliminar documentos en el orden correcto      │
│  Transacciones: VF11 · VI05 · VT02N · VL09                      │
│  Cuándo actúa: en cada batch. Nunca navega a VL06F.             │
└─────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────┐
│  session3 — Órdenes                                             │
│  Responsabilidad: leer ZCMR y eliminar/cancelar orders          │
│  Transacciones: ZCMR · VA02 · ME22N                            │
│  Cuándo actúa: Batch 6 (eliminación inicial) y Batch 8 (retry). │
└─────────────────────────────────────────────────────────────────┘
```

> Si SAP no permite abrir sesiones adicionales, `sap_login.py` asigna
> `session2 = session1` y `session3 = session1` como fallback y el
> pipeline continúa con una sola sesión.

---

## Flujo completo (chunked)

```
1. DESCARGA
   Google Drive → Desktop/DIL/May/archivo.xlsx

2. LECTURA
   archivo.xlsx → columna ZCMR → lista de tickets

3. LOGIN SAP
   sap_login.py → abre session1, session2, session3

4. CHUNKING
   Los tickets se parten en chunks de 100 (configurable con --chunk-size).
   Cada chunk se procesa COMPLETO por todos los batches antes de pasar al
   siguiente. Esto evita timeouts en VL09 y limita locks en SAP a un grupo
   manejable. Las sesiones SAP se REUSAN entre chunks (no se vuelve a logear).

   FOR each chunk of 100 tickets:

   4.1 LECTURA VL06F  [session1]
       Una consulta para los ~100 tickets del chunk actual.
       Retorna: {ticket: {billing_doc, shpt_cst, shipment, wbstk, delivery, block}}

   4.2 BATCH 1 — Billing Documents (VF11)
       ├── [session2] Eliminar billing docs en una llamada (max 14/página).
       └── [session1 / VL06F] verify bulk → exitosos continúan

   4.3 BATCH 2 — Shipment Cost (VI05)
       ├── [session2] Per-ticket con navegación limpia + estado=C handling.
       └── [session1 / VL06F] verify bulk

   4.4 BATCH 3 — Shipment Number (VT02N)
       ├── [session2] Per-ticket cargar → desasignar → guardar.
       └── [session1 / VL06F] verify bulk

   4.5 BATCH 4 — Reverse PGI (VL09)
       ├── [session2] Multi-valor con todos los deliveries del chunk →
       │     SelectAll → Reversar. Solo WBSTK != "A".
       └── [session1 / VL06F] verify bulk

   4.6 BATCH 5 — BOL Deletion (VL06F)
       [session1] Per-ticket. Filtra por delivery → btn[13]/[14] → confirm.

   4.7 BATCH 6 — ZCMR Orders (ZCMR + VA02 / ME22N)
       [session3] Lee orders en memoria → por cada uno:
       ├── intercompany → VA02 línea específica o borrado total
       └── intracompany (order[:2]=='47') → ME22N
       Captura "subsequent document" errors como fallidos.

   4.8 BATCH 7 — Verificación Final (ZSD_DEL_TICKETS)
       [session1] Confirma qué orders quedaron pendientes. Corrige Batch 6
       si encuentra tickets que SAP reportó éxito pero ZSD muestra pending.

   4.9 BATCH 8 — Order Cancellation retry (VA02)
       [session3] Solo si Batch 7 detectó pendientes. Cancela vía menú.

   4.10 ACUMULACIÓN
        Los exitosos/fallidos del chunk se mergean al diccionario consolidado
        que mantiene resultados de TODOS los chunks juntos.

5. REPORTE CONSOLIDADO FINAL
   Una sola vez al terminar todos los chunks: print_report + write_report_xlsx
   con datos agregados de todos los chunks juntos.
```

### Por qué chunking

SAP VL09 (reverse PGI) tiene un límite práctico de ~60-100 tickets por bulk:
- Timeout del work process de diálogo (`rdisp/max_wprun_time` = 600 seg default)
- Update task queue overflow
- Lock conflicts entre múltiples deliveries

Chunks de 100 mantienen cada operación dentro del rango seguro y reducen el
blast radius de un fallo: si un chunk falla por timeout, los demás siguen.

---

## Checkpointing (resume después de interrupciones)

El pipeline persiste el progreso por chunk en disco para sobrevivir
interrupciones (timeout, kill, crash, reinicio). Al re-correr, salta los
chunks ya completados.

```
Data-bases/Estado/{nombre_excel}_state.json
├── version, run_id, started_at
├── input_file, input_file_hash (SHA256)
├── chunk_size, total_tickets, total_chunks
├── batches_to_run
└── chunks: {
      "0": { tickets, resultados, vl06f, duración, ... },
      "1": { ... },
      ...
    }
```

### Reglas del state

| Evento | Comportamiento |
|--------|----------------|
| Primera corrida | Crea `_state.json` con hash del Excel; agrega chunks a medida que completan |
| Re-corrida con Excel sin cambios | Carga state, **skip chunks completados**, continúa desde el siguiente |
| Re-corrida con Excel modificado | Hash mismatch → **descarta state** y arranca fresh automáticamente |
| `--fresh` flag | Elimina state file antes de arrancar (forzar fresh) |
| `--report-only` flag | NO toca SAP; solo consolida state existente y emite reporte xlsx |
| Crash mid-chunk | El chunk NO se guarda en state — al re-correr se reintenta completo |
| Escritura del state | Atómica (write a `.json.tmp` + rename) — nunca queda corrupto |

### Idempotencia a nivel SAP

Las operaciones SAP **no son idempotentes** (un billing doc se borra una vez),
pero el pipeline ya maneja esto correctamente: las verifies clasifican como
**exitoso** a los tickets que ya no están en VL06F. Por lo tanto, el
checkpointing previene **re-trabajo innecesario**, no errores de double-delete.

---

## Ejecución parcial del pipeline

El pipeline acepta los siguientes argumentos para control fino:

```bash
python main.py                       # default: todos los batches, chunks de 100
python main.py --batches 2,3,4       # solo Batches 2, 3, 4
python main.py --batches 6-8         # desde Batch 6 hasta el final
python main.py --batches 1,4,6-8     # mixto
python main.py --chunk-size 50       # chunks más pequeños (más seguro/lento)
python main.py --chunk-size 200      # chunks más grandes (riesgo de timeout)
python main.py --fresh               # ignora state previo, arranca de cero
python main.py --report-only         # solo regenera reporte xlsx, NO toca SAP
python main.py --batches 6-8 --fresh # combinable
```

| Flag | Default | Descripción |
|------|---------|-------------|
| `--batches` | `all` | Subset de batches a correr. Formato: `1,2,3` o `6-8` o `1,4,6-8` |
| `--chunk-size` | `100` | Tickets por chunk. Límite práctico VL09 es 60-100 |
| `--fresh` | `false` | Elimina state previo antes de arrancar (forza fresh run) |
| `--report-only` | `false` | No login SAP. Solo consolida state existente y emite xlsx |

**Comportamiento clave:** si no se selecciona ningún batch 1-5, el pipeline
**omite la lectura de VL06F** y pasa todos los tickets del Excel directamente
a Batch 6. Esto permite re-procesar órdenes pendientes cuando los tickets ya
fueron eliminados de VL06F en una corrida anterior.

---

## Reglas del flujo

- Los tickets son **independientes entre sí** — el fallo de uno no detiene los demás
- El orden de los batches es **obligatorio** — SAP bloquea pasos posteriores si los anteriores no están completos
- Cada batch opera **solo sobre los tickets que pasaron el batch anterior**
- Si una operación bulk falla completamente, la verificación (VL06F) clasifica correctamente los tickets afectados como fallidos
- Si la verificación falla, los tickets quedan como fallidos por seguridad (nunca falso positivo)

---

## Estructura de archivos

```
Proyecto DIL/
├── code/
│   ├── main.py             # Orquestador — chunks + checkpoint + reporte
│   ├── sap_login.py        # Login y apertura de las 3 sesiones SAP
│   ├── sap_utils.py        # Utilidades compartidas (_navigate_to, _wait_ready, etc.)
│   ├── sap_vl06f.py        # session1: read_vl06f_data, delete_bol
│   ├── sap_batches.py      # session2: VF11, VI05, VT02N, VL09
│   ├── sap_orders.py       # session3: ZCMR, VA02, ME22N
│   ├── verifications.py    # session1: verificaciones bulk post-batch
│   ├── checkpoint.py       # Persiste/carga state JSON por chunk (resume)
│   ├── path.py             # Resuelve rutas de Entradas/Salidas/Estado
│   ├── excel_reader.py     # Lectura de columna ZCMR del Excel
│   ├── report_writer.py    # Escribe el xlsx final con sheets por batch
│   ├── sap_operations.py   # Re-exports de compatibilidad (no usar directamente)
│   ├── drive_download.py   # Descarga desde Google Drive (pendiente)
│   └── credentials.json    # Credenciales SAP (nunca compartir)
├── Data-bases/
│   ├── Entradas/<mes>/     # Excel de input (columna ZCMR)
│   ├── Salidas/<mes>/      # Reporte xlsx generado
│   └── Estado/             # State files JSON para resume (auto-creada)
└── Documentation/
    ├── Architecture.md
    ├── Entendimiento de modulos.md
    ├── Manual.md
    └── how to use it.md
```

---

## Módulos

### `sap_login.py`
Maneja login y apertura de las 3 sesiones. Si SAP no permite múltiples sesiones,
asigna `session2 = session3 = session1` como fallback automático.

### `sap_utils.py`
Utilidades compartidas usadas por todos los módulos:
- `_navigate_to(session, t_code)` — navegar a cualquier transacción SAP
- `_go_back(session, times)` — presionar F3 N veces
- `_wait_ready(session)` — esperar a que SAP termine de procesar
- `_enter_multi_values(session, table_id, values)` — llenar popup de selección múltiple
- `_POPUP_TABLE` — ID del popup de selección múltiple (constante compartida)

### `sap_vl06f.py` — session1
- `read_vl06f_data(session, tickets)` — lectura por chunk
- `delete_bol(session, delivery)` — eliminación de BOL (Batch 5)

### `sap_batches.py` — session2
- `delete_billing_documents_bulk(session, docs)` — VF11 bulk (Batch 1)
- `delete_shipment_costs_all(session, costs)` — VI05 stay-in-transaction (Batch 2)
- `delete_shipment_numbers_all(session, shipments)` — VT02N stay-in-transaction (Batch 3)
- `reverse_pgi_bulk(session, deliveries)` — VL09 bulk (Batch 4)

### `sap_orders.py` — session3
- `delete_orders_from_zcmr(session, tickets)` — orquestador ZCMR (Batch 6)
- `_read_zcmr_orders(session, tickets)` — lee todo en memoria antes de navegar
- `_delete_intercompany_order_va02(session, order, delivery)` — VA02 (Batch 6)
- `_delete_intracompany_order_me22n(session, order)` — ME22N (Batch 6)
- `cancel_va02_order(session, order)` — cancelación retry en VA02 (Batch 8)

### `verifications.py` — session1
- `verify_billing_documents_bulk(session, tickets)` → `(exitosos, fallidos)`
- `verify_shipment_costs_bulk(session, tickets)` → `(exitosos, fallidos)`
- `verify_shipment_numbers_bulk(session, tickets)` → `(exitosos, fallidos)`
- `verify_pgi_reversed_bulk(session, tickets)` → `(exitosos, fallidos)`
- `verify_zsd_del_tickets(session, tickets)` → `(exitosos, fallidos, ticket_to_order)` (Batch 7)

### `checkpoint.py` — persistencia
- `load_state(input_file)` → state dict o None (None si hash mismatch)
- `init_state(input_file, tickets, chunk_size, batches_to_run)` → state nuevo
- `is_chunk_completed(state, chunk_idx)` → bool
- `save_chunk_result(state, chunk_idx, tickets, result, duration, input_file)` → escritura atómica
- `consolidate_state(state)` → tupla `(resultados, vl06f, zcmr_failures, ticket_to_order, cancel_failures, no_enc)` acumulada de todos los chunks
- `clear_state(input_file)` — borra el state file (para `--fresh`)

### `main.py`
Orquesta el flujo completo: parsea CLI, lee tickets, login SAP, carga/crea
state, itera chunks (skip los completados), llama `process_chunk()`, persiste
con `save_chunk_result`, y al final consolida con `consolidate_state` para
emitir reporte único.

---

## Reporte final

El reporte se genera UNA sola vez al terminar TODOS los chunks (consolidado).

```
████████████████████████████████████████████████████████████
  CHUNK 1/3  —  100 tickets (115232482 ... 300791882)
████████████████████████████████████████████████████████████
  Batch 1: Billing Documents (bulk)...
  ...
  [Checkpoint] Chunk 1/3 guardado (1820s).

████████████████████████████████████████████████████████████
  CHUNK 2/3  —  100 tickets (336810852 ... 364175285)
████████████████████████████████████████████████████████████
  ...
  [Checkpoint] Chunk 2/3 guardado (1750s).

████████████████████████████████████████████████████████████
  CHUNK 3/3  —  50 tickets (399206199 ... 91266389)
████████████████████████████████████████████████████████████
  ...
  [Checkpoint] Chunk 3/3 guardado (920s).

═════════════════════════════════════════════════════════════
  REPORTE CONSOLIDADO — TODOS LOS CHUNKS
═════════════════════════════════════════════════════════════
  REPORTE DIL — 2026-05-22
  Tickets en Excel: 250

  BATCH 1 — Billing Documents
    Exitosos (250): T001, T002, T003, ...
    Fallidos   (0): —

  BATCH 2 — Shipment Cost
    Exitosos (248): T001, T002, T004, ...
    Fallidos   (2): T003, T067

  BATCH 3 — Shipment Number
    Exitosos (248): T001, T002, T004, ...
    Fallidos   (0): —

  BATCH 4 — Reverse PGI
    Exitosos (247): T001, T002, T004, ...
    Fallidos   (1): T112

  BATCH 5 — BOL Deletion
    Exitosos (247): T001, T002, T004, ...
    Fallidos   (0): —

  BATCH 6 — ZCMR Orders
    Exitosos (245): T001, T002, T004, ...
    Fallidos   (2): T089, T201

  BATCH 7 — Verificación Final ZSD
    Exitosos (245): T001, T002, T004, ...
    Fallidos   (0): —

  BATCH 8 — Order Cancellation
    Exitosos (2): T089, T201
    Fallidos  (0): —

═════════════════════════════════════════════════════════════
  Completados al 100%:  247/250 procesados
  Con fallos parciales:  3/250
  Revisar manualmente:  T003, T067, T112
═════════════════════════════════════════════════════════════
```

> **Batch 8** solo procesa los tickets que **Batch 7** detectó con orders
> pendientes en ZSD_DEL_TICKETS. En el ejemplo, T089 y T201 fueron
> recuperados por Batch 8 después de fallar Batch 6.
