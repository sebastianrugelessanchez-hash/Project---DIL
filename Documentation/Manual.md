# Manual — DIL Pipeline Modules

This document explains each Python module under `code/` and the key functions exposed by each. Use it as a quick reference for what each piece does, what arguments it takes, and what it returns.

---

## Project Structure

```
Proyecto DIL/
├── code/
│   ├── main.py             # Orchestrator — chunking + checkpointing + report
│   ├── sap_login.py        # Login + opens 3 parallel SAP sessions
│   ├── sap_utils.py        # Shared helpers (_navigate_to, _wait_ready, clipboard, etc.)
│   ├── sap_vl06f.py        # session1: read VL06F data + BOL deletion (Batch 5)
│   ├── sap_batches.py      # session2: VF11, VI05, VT02N, VL09 (Batches 1-4)
│   ├── sap_orders.py       # session3: ZCMR + VA02 + ME22N (Batches 6, 8)
│   ├── verifications.py    # session1: bulk verifies after each batch + ZSD (Batch 7)
│   ├── checkpoint.py       # State JSON persistence (resume after interrupts)
│   ├── path.py             # Resolves input/output/state paths
│   ├── excel_reader.py     # Reads ZCMR column from Excel
│   ├── report_writer.py    # Builds final xlsx report
│   └── credentials.json    # SAP login credentials (never share)
├── Data-bases/
│   ├── Entradas/<mes>/     # Input Excel
│   ├── Salidas/<mes>/      # Output xlsx report
│   └── Estado/             # Checkpoint state JSON (auto-created)
└── Documentation/
```

---

## Pipeline Flow (overview)

The pipeline processes tickets in **chunks of 100** (configurable). Each chunk runs the full sequence below; one consolidated report is emitted at the end of all chunks.

```
Read Excel → Login SAP → Load/init state → FOR each chunk:
    Read VL06F → Batch 1 → 2 → 3 → 4 → 5 → 6 → 7 → 8 → Save state
→ Consolidate state → Print/write report
```

| Batch | Transaction | Module | Purpose |
|-------|-------------|--------|---------|
| 1 | VF11 | `sap_batches.py` | Cancel billing documents |
| 2 | VI05 | `sap_batches.py` | Delete shipment costs |
| 3 | VT02N | `sap_batches.py` | Unassign / delete shipments |
| 4 | VL09 | `sap_batches.py` | Reverse PGI |
| 5 | VL06F | `sap_vl06f.py` | Delete BOL |
| 6 | ZCMR/VA02/ME22N | `sap_orders.py` | Delete orders |
| 7 | ZSD_DEL_TICKETS | `verifications.py` | Verify pending orders |
| 8 | VA02 | `sap_orders.py` | Cancel orders retry |

---

## `main.py`

Orchestrator: parses CLI, reads Excel, opens SAP sessions, runs chunks, persists state, and emits the final report.

### CLI flags

| Flag | Default | Description |
|------|---------|-------------|
| `--batches` | `all` | Subset of batches to run (`1,2,3` or `6-8` or `all`) |
| `--chunk-size` | `100` | Tickets per chunk |
| `--fresh` | `false` | Discard existing state and start from chunk 0 |
| `--report-only` | `false` | Do NOT touch SAP. Only consolidate state → emit xlsx |

### Key functions

- **`parse_batch_spec(spec)`** — Parses `'1,2,3'` / `'6-8'` / `'all'` / `'1,4,6-8'` into a set of batch numbers. Raises `ValueError` for invalid batches.
- **`parse_args()`** — argparse setup, returns the namespace.
- **`print_report(resultados, total_tickets, tickets_no_encontrados)`** — Prints the consolidated console report.
- **`process_chunk(tickets, session1, session2, session3, batches_to_run, needs_vl06f)`** — Runs Batches 1→8 for a single chunk. Returns the tuple `(resultados, vl06f, zcmr_failures, ticket_to_order, cancel_failures, tickets_no_encontrados)`.
- **`_emit_consolidated_report(state)`** — Calls `consolidate_state(state)` then `print_report` + `write_report_xlsx`.
- **`main()`** — Entry point. Coordinates the entire pipeline.

---

## `sap_login.py`

Handles SAP GUI startup, login, and opens up to 3 parallel sessions (session1, session2, session3) under the same connection. Falls back to a single session if SAP rejects additional ones.

### Configuration constants (top of file)

| Constant | Value | Description |
|---|---|---|
| `CREDENTIALS_FILE` | `credentials.json` | Path to credentials JSON |
| `SAP_PATH` | `C:\...\saplogon.exe` | SAP Logon executable |
| `SAP_SYSTEM` | `PRD - ECC Production` | System name as shown in SAP Logon |
| `CLIENT` | `900` | SAP client number |
| `LANGUAGE` | `EN` | Login language |

### `load_credentials(filepath)`

Reads `credentials.json` and returns a dict with `username` / `password`. Raises with a clear message on missing file, invalid JSON, or missing keys.

```json
{ "username": "your_user", "password": "your_password" }
```

### `class SapAutomation`

Main class. `__init__(credentials_path)` loads creds and prepares empty placeholders. The full flow is `launch_sap() → open_session() → login() → open_additional_sessions()`.

| Method | Purpose |
|--------|---------|
| `launch_sap()` | Connects to SAP COM (`GetObject("SAPGUI")`) or launches `saplogon.exe` if not running |
| `_wait_for_sap(retries, delay)` | Polls until the COM object is available; default 10 attempts × 2 sec |
| `open_session(retries, delay)` | `OpenConnection(SAP_SYSTEM)`, retrieves `connection.Children[0]`, maximizes window |
| `login()` | Fills RSYST fields, submits Enter, reads status bar to detect E/W/A errors |
| `_wait_for_session_ready(timeout, poll)` | Polls `session.Busy` instead of fixed sleep |
| `open_additional_sessions()` | Opens session2 and session3 via `CreateSession()`; falls back to `session1` if denied |
| `_get_status_message()` | Returns `(text, message_type)` where type is `E`/`W`/`A`/`S`/`I`/`""` |
| `run()` | Orchestrates the full sequence |

After `run()` completes:
- `sap.session` → session1 (VL06F)
- `sap.session2` → session2 (operations)
- `sap.session3` → session3 (orders)

---

## `sap_utils.py`

Shared utilities used by every SAP module. Stateless — every function takes `session` as an argument.

### Constants

- **`_POPUP_TABLE`** — Full path to the multi-value selection popup table in `wnd[1]`.

### Functions

| Function | Purpose |
|----------|---------|
| `_normalize_ticket(value)` | Strips whitespace + leading zeros for ticket matching (e.g., `'0044688706'` → `'44688706'`) |
| `_navigate_to(session, t_code)` | Types `/n<tcode>` in the OK code field and sends Enter |
| `_go_back(session, times=1)` | Presses F3 N times with small sleep between |
| `_wait_ready(session, timeout=10)` | Polls `session.Busy` until idle |
| `_wnd_exists(session, wnd_id)` | True if `findById(wnd_id)` succeeds |
| `_find_popup_wnd(session)` | Returns `'wnd[2]'` or `'wnd[1]'` whichever has the multi-value tabstrip |
| `_popup_table(wnd)` | Builds the multi-value table path inside the given popup window |
| `_set_clipboard_text(text)` | Sets Windows clipboard text (CF_UNICODETEXT) for the "Upload from Clipboard" pattern |
| `_popup_table_row_count(session, table_id)` | (Diagnostic only; popup is virtualized so result is capped at ~8) |
| `_enter_multi_values(session, table_id, values)` | Loads N values into a multi-value popup via clipboard + `btn[24]` (Upload from Clipboard). Falls back to scroll-based legacy method if upload fails |
| `_enter_multi_values_legacy(session, table_id, values)` | Fallback: writes values row-by-row scrolling every 8 entries (only reliable ≤16 values) |

---

## `sap_vl06f.py` — session1

VL06F reads and BOL deletion (Batch 5).

### `read_vl06f_data(session, tickets) → dict`

Navigates to VL06F, filters by the given tickets via multi-value popup, presses F8, applies the `/02C BOL STATUS` layout, materializes all rows via force-scroll, and reads the grid.

Returns: `{vbeln: {billing_doc, shpt_cst, shipment, wbstk, block, delivery, invoice_il}}`

Notes:
- Returns `{}` if `tickets` is empty (prevents F8 with no filter → freeze).
- Logs `[VL06F] Columnas disponibles` and `Column map final` for diagnostics.
- Applies rule R001: if `invoice_il` starts with `'7'`, `billing_doc` is set to `""`.

### `delete_bol(session, delivery)`

Per-ticket BOL deletion in VL06F. Navigates to VL06F, filters by the single delivery, presses F8, selects row 0, calls `btn[13]` + `btn[14]`, confirms popups, refreshes. Used by Batch 5.

### Internal helpers

- `_vl06f_delivery_filter(session, tickets)` — opens IT_VBELN popup, clears with `btn[16]`, uploads tickets, closes.
- `_select_bol_layout(session)` — Ctrl+F9 dialog → finds and selects layout `/02C`; logs available layouts if not found.
- `_build_vl06f_column_map(grid)` — detects technical column names (VBELN, ZZVBELN, etc.) from the actual `ColumnOrder`.

---

## `sap_batches.py` — session2

Batches 1-4 (VF11, VI05, VT02N, VL09). All functions take `session2` and operate on lists from the current chunk.

### Batch 1 — `delete_billing_documents_bulk(session, billing_docs)`

VF11 bulk cancel. Fills the entry table (14 docs per page; `btn[7]` for additional pages), sets today's date, saves with `btn[11]`, handles confirmation popups, reads status bar.

### Batch 2 — `delete_shipment_costs_all(session, shpt_csts)`

VI05 per-ticket with clean navigation each iteration. Each ticket runs `_process_single_shpt_cst()`:

1. `_vi05_setup_selection_screen()` — fresh nav to VI05 + sets date range
2. Open S_FKNUM multi-value popup, enter ticket, F8
3. Read status bar
4. Read `estado` from `lbl[33,5]`
5. If `estado == "C"`: call `_cambiar_estado_transferencia()` (mark SLSTOR=True + set STDAT date)
6. Call `_eliminar_shpt_cst()` (select chk[1,4], edit, delete, confirm)
7. Capture SAP status bar message for diagnostics

`_cambiar_estado_transferencia(session, fecha)` tries multiple date formats (MM/DD/YYYY, DD.MM.YYYY, DD/MM/YYYY, etc.) until one is accepted. Logs which one worked.

### Batch 3 — `delete_shipment_numbers_all(session, shipments)`

VT02N per-shipment. Each iteration runs `_process_single_shipment()`:

1. Fresh nav to VT02N
2. Enter VTTK-TKNUM, Enter
3. Click STABF (cancel transport) + STDIS (cancel planning)
4. Open planning screen (F7)
5. Select node "2" in tree, press button `MM_UNAS  10001` on the toolbar (`shellcont[1]/shell[0]`)
6. Save, confirm popup, capture status bar

### Batch 4 — `reverse_pgi_bulk(session, deliveries)`

VL09 bulk reverse. Opens IT_VBELN multi-value, uploads ALL deliveries of the chunk via clipboard, F8, SelectAll on grid, presses `btn[5]` (Reverse), confirms popups, refreshes. Reads status bar to detect timeouts.

### Per-ticket variants

`delete_billing_document`, `delete_shipment_cost`, `delete_shipment_number`, `reverse_pgi` exist for single-ticket use but **the pipeline does not call them** — only the `_all` / `_bulk` variants are used.

---

## `sap_orders.py` — session3

Order deletion via ZCMR and order cancellation via VA02 (Batches 6 and 8).

### `delete_orders_from_zcmr(session, tickets) → dict[ticket, list[orders_failed]]`

Batch 6 orchestrator. Two phases:

1. `_read_zcmr_orders(session, tickets)` — navigates to ZCMR, filters by tickets, reads grid + sub-grids to collect a list of `{ticket, order, delivery, is_intracompany}`. Dedups by `(order, delivery)`.
2. For each order:
   - `is_intracompany` (`order[:2] == "47"`) → `_delete_intracompany_order_me22n(session, order)`
   - Otherwise → `_delete_intercompany_order_va02(session, order, delivery)`

Returns a dict of `ticket → [orders_that_failed]`.

### `_delete_intercompany_order_va02(session, order, delivery)`

VA02. Navigates, enters order. If SAP shows an "items list" popup, deletes only the line matching the delivery (via `btnBT_POLO`). If no popup, deletes the entire order via `menu[0]/menu[11]`. Calls `_check_va02_error_after_action()` to catch "subsequent document" / "cannot be deleted" errors.

### `_delete_intracompany_order_me22n(session, order)`

ME22N. Other PO → enter EBELN → select row in `tblSAPLMEGUITC_1211` → `btnDELETE` → save → confirm.

### `cancel_va02_order(session, order)`

Batch 8 retry. Cancels an order via VA02 → menu Delete → confirm. Used when Batch 7 detects orders still pending in ZSD_DEL_TICKETS.

### `_check_va02_error_after_action(session, ...)`

Internal helper that reads popup text and status bar for keywords like "cannot be deleted", "subsequent document", "no se puede borrar". Raises `RuntimeError` with the SAP message if found — so the calling batch reports a real failure instead of silent success.

---

## `verifications.py` — session1

Bulk verifications after each batch + final verification via ZSD_DEL_TICKETS.

### `_classify(tickets, data, field, success_value="") → (exitosos, fallidos)`

Internal helper. For each ticket:
- Not in `data` (i.e., no longer in VL06F) → **exitoso** (assume already processed)
- In `data` and `data[t][field] == success_value` → **exitoso**
- Otherwise → **fallido**

### Bulk verifies (one VL06F call each)

| Function | Field checked | Success value |
|----------|---------------|---------------|
| `verify_billing_documents_bulk(session, tickets)` | `billing_doc` | `""` |
| `verify_shipment_costs_bulk(session, tickets)` | `shpt_cst` | `""` |
| `verify_shipment_numbers_bulk(session, tickets)` | `shipment` | `""` |
| `verify_pgi_reversed_bulk(session, tickets)` | `wbstk` | `"A"` |

All return `(exitosos, fallidos)`. Return `([], [])` immediately if `tickets` is empty.

### `verify_zsd_del_tickets(session, tickets) → (exitosos, fallidos, ticket_to_order)`

Batch 7. Navigates to ZSD_DEL_TICKETS, filters by P_TICKET, F8, reads the grid.

- For each row, extracts `(ticket, order)` from the same row.
- Filters out values that don't look like real Sales Orders (≥10 digits, all numeric, not all-zeros) — prevents picking up "Order Code" (7-digit code) instead of "Order" (10-digit SAP sales order).
- Returns `ticket_to_order` mapping for Batch 8 retry. Tickets in this dict are marked as **fallidos** in Batch 7.

**Fallback:** if the grid has orders but cannot map them to tickets (e.g., TICKET column hidden in layout), marks ALL tickets as failed and distributes detected orders cyclically — safer than reporting false success.

---

## `checkpoint.py` — state persistence

JSON-based progress tracking per chunk. Allows resume after interruptions.

### `STATE_DIR`

Default state directory: `<BASE_DIR>/Data-bases/Estado/`. Auto-created.

### Functions

| Function | Purpose |
|----------|---------|
| `compute_file_hash(path) → str` | SHA256 of file content (streaming) |
| `state_path_for_input(input_file) → Path` | Returns `Estado/{stem}_state.json` |
| `load_state(input_file) → dict \| None` | Loads state if hash matches the input file; returns `None` if not found or hash mismatch |
| `init_state(input_file, tickets, chunk_size, batches_to_run) → dict` | Creates a fresh state dict |
| `is_chunk_completed(state, chunk_idx) → bool` | O(1) check |
| `save_chunk_result(state, chunk_idx, tickets, result, duration_seconds, input_file)` | Persists the chunk atomically (write `.json.tmp` → rename) |
| `consolidate_state(state) → tuple` | Merges all chunks into the tuple `(resultados, vl06f, zcmr_failures, ticket_to_order, cancel_failures, tickets_no_encontrados)` |
| `clear_state(input_file)` | Deletes the state file (used by `--fresh`) |

### State schema (excerpt)

```json
{
  "version": 1,
  "run_id": "2026-05-22T10:00:00_pid12345",
  "input_file": "Billing USA -May 22 2026.xlsx",
  "input_file_hash": "sha256:...",
  "chunk_size": 100,
  "total_tickets": 250,
  "total_chunks": 3,
  "chunks": {
    "0": {
      "tickets": [...],
      "completed_at": "2026-05-22T10:30:00",
      "duration_seconds": 1820,
      "resultados": { "BATCH 1 — ...": [[exit...], [fail...]], ... },
      "vl06f": {...},
      "zcmr_failures": {...},
      "ticket_to_order": {...},
      "cancel_failures": {...},
      "tickets_no_encontrados": []
    }
  }
}
```

---

## `path.py`

Resolves filesystem paths for inputs, outputs, and state.

| Constant | Path |
|----------|------|
| `BASE_DIR` | `<repo root>` (parent of `code/`) |
| `ENTRADAS_DIR` | `<BASE_DIR>/Data-bases/Entradas` |
| `SALIDAS_DIR` | `<BASE_DIR>/Data-bases/Salidas` |

### Functions

- **`_latest_subdir(parent) → Path`** — Returns the most recently modified subdirectory.
- **`_xlsx_files(folder) → list[Path]`** — Lists `.xlsx` files in the folder, excluding Office lock files (`~$...xlsx`), sorted by mtime descending.
- **`get_billing_file(filename=None) → Path`** — Returns the input Excel. If `filename` is given, looks for it in the latest Entradas subfolder. Otherwise returns the most recent `.xlsx`.
- **`get_report_file() → Path`** — Returns the output xlsx path. Returns the most recent existing report in the latest Salidas subfolder, or generates a name like `Reporte DIL - May 22 2026.xlsx` if none exists.

---

## `excel_reader.py`

### `read_zcmr(filepath=None) → list[str]`

Reads the ZCMR column (column C, header on row 3) of the input Excel.

- If `filepath` is omitted, uses `get_billing_file()`.
- Validates that cell C3 says `ZCMR`. Raises `ValueError` otherwise.
- Skips empty cells and non-numeric values.
- Deduplicates while preserving insertion order.
- Normalizes each ticket via `_normalize_ticket` (strip whitespace, strip leading zeros).

Returns a list of ticket strings.

---

## `report_writer.py`

### `write_report_xlsx(resultados, total_tickets, vl06f, zcmr_failures, ticket_to_order, cancel_failures, tickets_no_encontrados)`

Builds the final xlsx report using openpyxl.

### Sheets generated

| Sheet | Content |
|-------|---------|
| `Resumen` | Counts of exitosos/fallidos per batch + summary stats |
| `Billing Documents` | Tickets that failed Batch 1, with their `N° Billing Document` |
| `Shipment Cost` | Tickets that failed Batch 2, with their `N° Shipment Cost` |
| `Shipment Number` | Tickets that failed Batch 3, with their `N° Shipment Number` |
| `Reverse PGI` | Tickets that failed Batch 4, with their `N° Delivery` |
| `BOL Deletion` | Tickets that failed Batch 5, with their `N° Delivery` |
| `ZCMR Orders` | Tickets that failed Batch 6, with their failed order numbers |
| `Verificación Final ZSD` | Tickets that failed Batch 7, with the pending `N° Order` |
| `Order Cancellation` | Tickets where Batch 8 retry failed, with the `N° Order` |
| `No encontrados en VL06F` | Tickets that VL06F did not return at all (require investigation) |

Sheets for batches with **zero failures are omitted** to keep the file clean. The Resumen sheet always exists.

The file path is resolved via `get_report_file()`.

---

## Entry Point

```python
if __name__ == "__main__":
    main()
```

Runs only when executing `python main.py` directly. Use the CLI flags described in `main.py` above to control behavior.
