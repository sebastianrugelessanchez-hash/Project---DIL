# Project DIL — Autonomous SAP Ticket-Decommissioning Pipeline

> Drives the SAP ECC GUI to fully decommission transport tickets end-to-end — no manual steps, no silent data loss.

---

## Overview

Project DIL automates a process previously performed manually in **SAP ECC (production)**: given a list of transport tickets, it deletes or reverses all associated documents in the mandatory teardown order. Think of it as a controlled **teardown chain** — each document type must be removed before the next one can be removed.

The bot drives the real SAP GUI through **SAP GUI Scripting (COM)**, runs **three parallel SAP sessions**, processes tickets in **resilient chunks**, verifies every step against SAP itself, recovers from interruptions, and emits an auditable XLSX + log report.

**Validated in production (June 2026):** a 200-ticket batch was decommissioned end-to-end with no manual intervention.

---

## Key Features

- **Verify-then-decide order deletion** — inspects each order line-by-line before deleting; never touches lines outside scope, even in mixed orders.
- **Order-level idempotency** — re-running never fails on orders already deleted in a previous run.
- **Batch 8 de-duplication** — dozens of tickets sharing the same order are processed once; result is propagated to all tickets.
- **Resilient grid-column matching** — resolves SAP column names against a candidate list and logs real names on first use, surviving layout changes.
- **File-based observability** — every run produces a timestamped UTF-8 log (`Data-bases/Logs/`) mirroring all console output.
- **Checkpoint & resume** — progress is persisted per chunk to disk; on re-run completed chunks are skipped automatically.
- **Read-only diagnostics** — `diagnose_grids.py` connects to a live SAP session and dumps real column names / layouts without modifying anything.
- **Bounded popup handling** — popup cascade after order save has a max-count, absolute deadline, and repeat-detection guard; never hangs the pipeline.

---

## Architecture

### SAP Sessions

The pipeline opens **three simultaneous SAP sessions** from the same connection, each with an exclusive responsibility:

| Session | Responsibility |
|---------|---------------|
| `session1` | VL06F reads, BOL/delivery deletion, ZSD ground-truth verification |
| `session2` | Bulk document reversal / deletion (VF11, VI05, VT02N, VL09) |
| `session3` | Order deletion — ZCMR, VA02, ME22N |

> **Fallback:** if SAP does not allow additional sessions, `sap_login.py` assigns `session2 = session3 = session1` and the pipeline continues with a single session.

### Why Chunking

Tickets are split into **chunks of 100** (configurable). SAP VL09 (reverse PGI) has a practical limit of ~60–100 tickets per bulk run due to:
- Dialog work-process timeout (`rdisp/max_wprun_time = 600 s` default)
- Update task queue overflow
- Lock conflicts between multiple deliveries

A single VL06F query loads every document number for the chunk into memory; all batches read from that snapshot instead of re-querying SAP — a key performance optimization.

---

## End-to-End Flow

```
FOR each chunk of ~100 tickets:
  4.1  READ VL06F              [session1]  one query → all document numbers in memory
  4.2  BATCH 1  VF11           [session2]  bulk-reverse billing documents       + verify
  4.3  BATCH 2  VI05           [session2]  delete shipment costs (unlock C)     + verify
  4.4  BATCH 3  VT02N          [session2]  un-assign + delete shipment          + verify
  4.5  BATCH 4  VL09           [session2]  bulk-reverse PGI (WBSTK != 'A')      + verify
  4.6  BATCH 5  VL06F          [session1]  delete BOL / delivery (per ticket)
  4.7  BATCH 6  ZCMR→VA02/ME22N [session3] delete orders (verify-then-decide)
  4.8  BATCH 7  ZSD_DEL_TICKETS [session1] ground-truth verification
  4.9  BATCH 8  VA02/ME22N     [session3]  retry pending orders (de-duplicated)
  4.10 ACCUMULATE chunk results into the consolidated dictionary

AFTER all chunks: consolidated report (console + xlsx)
```

### Flow Rules

- Tickets are **independent** — one ticket's failure never stops the others.
- **Batch order is mandatory** — SAP blocks later steps if earlier documents remain.
- Each batch operates only on tickets that **passed the previous batch**.
- If a bulk operation fails, the VL06F verification still classifies affected tickets correctly as failed.
- If verification fails, tickets are marked failed for safety — never a false positive.

---

## Safe Order Deletion: Verify-Then-Decide

A single order frequently mixes lines that must be deleted with lines that must be kept. The engine inspects each order's real state and decides **per line** what is safe to delete.

**Decision logic (per order):**

1. **Read** — open the order once and read its line items (`PO Number BSTKD_E` = ticket, zero-padded).
2. **Classify** — match every line against the in-scope ticket set: `in-scope` vs. `to-keep`.
3. **All lines in scope** — delete all selected lines; SAP removes the empty order.
4. **Mixed order** — select and delete **only** the in-scope lines; the rest persist untouched.
5. **No in-scope lines** — lines already gone → treated as success (idempotent).
6. **Fail safe** — if a line cannot be positively classified, the bot **stops rather than guessing** — it never deletes data it is unsure about.

---

## Project Structure

```
Proyecto DIL/
├── code/
│   ├── main.py               # Orchestrator: chunks + checkpoint + dedup Batch 8 + report
│   ├── sap_login.py          # Login + opens the 3 SAP sessions
│   ├── sap_utils.py          # Shared helpers (_navigate_to, _wait_ready, multi-value)
│   ├── sap_vl06f.py          # session1: read_vl06f_data, delete_bol
│   ├── sap_batches.py        # session2: VF11, VI05, VT02N, VL09
│   ├── sap_orders.py         # session3: ZCMR, VA02, ME22N + verify-then-decide engine
│   ├── verifications.py      # session1: bulk post-batch verification + ZSD ground truth
│   ├── checkpoint.py         # Per-chunk JSON state (resume)
│   ├── excel_reader.py       # Reads tickets (and the 'Manual Orders' sheet)
│   ├── report_writer.py      # Final XLSX report
│   ├── path.py               # Resolves Entradas / Salidas / Estado paths
│   ├── log_util.py           # Console + file logging (Tee, encoding-safe for cp1252)
│   ├── diagnose_grids.py     # Read-only SAP grid/layout diagnostics
│   ├── rules.json            # Configurable pipeline rules
│   └── credentials.example.json  # Credentials template (copy → credentials.json)
├── Documentation/
│   ├── Architecture (EN).docx
│   ├── Architecture.docx
│   ├── Architecture.md
│   ├── Manual.md
│   └── ...
└── .gitignore                # Data-bases/, Macro/, credentials.json excluded
```

> `Data-bases/` (input Excel, generated reports, logs, state) and `Macro/` are gitignored and must exist locally.

---

## Requirements

- **OS:** Windows (SAP GUI Scripting uses Windows COM automation)
- **SAP:** SAP ECC with GUI Scripting enabled (`sapgui/user_scripting = TRUE`)
- **Python:** 3.10 or higher

### Python Dependencies

```bash
pip install pywin32 openpyxl python-docx
```

| Package | Purpose |
|---------|---------|
| `pywin32` | SAP GUI Scripting via COM (`win32com.client`) |
| `openpyxl` | Read input Excel tickets + write XLSX report |
| `python-docx` | Architecture/case-study document generation |

---

## Setup

```bash
# 1. Clone the repository
git clone git@github.com:sebastianrugelessanchez-hash/Project---DIL.git
cd Project---DIL

# 2. Install dependencies
pip install pywin32 openpyxl python-docx

# 3. Configure SAP credentials
copy code\credentials.example.json code\credentials.json
```

Edit `code/credentials.json` and fill in your SAP username and password:

```json
{ "username": "YOUR_SAP_USER", "password": "YOUR_SAP_PASSWORD" }
```

```
# 4. Create the local data folders (gitignored)
mkdir -p "Data-bases/Entradas" "Data-bases/Salidas" "Data-bases/Estado" "Data-bases/Logs"
```

Place your input Excel file (with ticket numbers) in `Data-bases/Entradas/<month>/`.

---

## Usage

### Run the pipeline

```bash
cd code
python main.py
```

The orchestrator will:
1. Prompt for (or auto-detect) the input Excel file.
2. Split tickets into chunks of 100.
3. Process all 8 batches per chunk, with checkpoint saves.
4. Generate a consolidated XLSX report and timestamped log.

### Diagnostics (read-only)

Connect to a live SAP session to inspect real column names without modifying anything:

```bash
python diagnose_grids.py zcmr <ticket_number>
python diagnose_grids.py zsd <ticket_number>
python diagnose_grids.py vl06f <ticket_number>
python diagnose_grids.py all <ticket_number>
python diagnose_grids.py layouts
```

---

## Outputs

| Path | Content |
|------|---------|
| `Data-bases/Salidas/<month>/` | Final XLSX report (per-batch success/fail + summary) |
| `Data-bases/Logs/dil_run_YYYYMMDD_HHMMSS.log` | Full UTF-8 trace of every decision and SAP message |
| `Data-bases/Estado/<name>_state.json` | Chunk-level checkpoint for resume |

### Report Summary Format

```
===================================================
  CONSOLIDATED REPORT - ALL CHUNKS
===================================================
  BATCH 6 - ZCMR Orders     Successful: ...   Failed: ...
  BATCH 7 - Final ZSD        Successful: ...   Failed: ...
  BATCH 8 - Order Cancel     Successful: ...   Failed: ...
  ---------------------------------------------------
  100% completed:  N/total      Review manually:  ...
===================================================
```

---

## Checkpoint & Resume

If the pipeline is interrupted (timeout, kill, crash, reboot), re-running it automatically resumes from the last completed chunk. State is persisted in:

```
Data-bases/Estado/{excel_name}_state.json
```

The state file tracks: `run_id`, `input_file_hash` (SHA256), `chunk_size`, `total_tickets`, and per-chunk results including VL06F document snapshots.

---

## Partial Execution

If no batch 1–5 is selected, the pipeline skips the VL06F read and passes all tickets straight to Batch 6. This is useful to re-process pending orders after tickets were already removed from VL06F in a previous run.
