"""
Checkpointing por chunk para el pipeline DIL.

Persiste el progreso por chunk en disco (JSON) para permitir resume después de
interrupciones (timeout SAP, kill, crash). Atado al hash del Excel de entrada
para invalidar el state si el archivo cambia.

Uso típico en main.py:

    input_file = get_billing_file()
    state = load_state(input_file)
    if state is None:
        state = init_state(input_file, tickets, chunk_size, batches_to_run)

    for chunk_idx in range(n_chunks):
        if is_chunk_completed(state, chunk_idx):
            continue
        result = process_chunk(...)
        save_chunk_result(state, chunk_idx, chunk, result, duration, input_file)

    consolidated = consolidate_state(state)
"""
import hashlib
import json
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

from path import BASE_DIR


STATE_DIR = BASE_DIR / "Data-bases" / "Estado"


def compute_file_hash(path: Path) -> str:
    """SHA256 del contenido del archivo (streaming para archivos grandes)."""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(65536), b""):
            h.update(block)
    return h.hexdigest()


def state_path_for_input(input_file: Path, state_dir: Path = STATE_DIR) -> Path:
    """Genera la ruta del state file basada en el nombre del input Excel."""
    state_dir.mkdir(parents=True, exist_ok=True)
    return state_dir / f"{input_file.stem}_state.json"


def load_state(input_file: Path, expected_mode: str = "normal",
               state_dir: Path = STATE_DIR) -> Optional[dict]:
    """
    Carga el state file si existe Y:
      1. el hash del input matchea (Excel sin cambios)
      2. el mode coincide con el modo actual ('normal' vs 'manual')

    Retorna None si no hay state o si alguna validación falla.

    El mode separation es crítico: si la corrida previa fue 'normal' (procesando
    tickets de ZCMR/Ticket Number) y ahora corres '--manual-only' (procesando
    pares de Manual Orders), los tickets son DISTINTOS aunque el Excel sea el
    mismo. Sin esta separación, el state previo del modo normal saltaría
    chunks del modo manual y viceversa.
    """
    state_file = state_path_for_input(input_file, state_dir)
    if not state_file.exists():
        return None
    try:
        state = json.loads(state_file.read_text(encoding="utf-8"))
    except Exception as e:
        print(f"  [Checkpoint] Error leyendo state: {e}", file=sys.stderr)
        return None

    actual_hash = compute_file_hash(input_file)
    if state.get("input_file_hash") != actual_hash:
        print(f"  [Checkpoint] Input file cambió (hash mismatch). State descartado.")
        return None

    # Backwards-compatible: si el state viejo no tiene 'mode', asumimos 'normal'
    state_mode = state.get("mode", "normal")
    if state_mode != expected_mode:
        print(f"  [Checkpoint] Mode cambió ({state_mode!r} -> {expected_mode!r}). "
              f"State descartado (re-procesando desde chunk 0).")
        return None

    return state


def init_state(input_file: Path, tickets: list, chunk_size: int,
               batches_to_run: set, mode: str = "normal") -> dict:
    """
    Crea un state nuevo (cuando no hay resume).

    mode: 'normal' para procesamiento ZCMR/Ticket Number con Batches 0-8,
          'manual' para procesamiento de pares (ticket, order) con Batches 6-8.
    """
    n_chunks = (len(tickets) + chunk_size - 1) // chunk_size
    now = datetime.now().isoformat()
    return {
        "version": 1,
        "run_id": f"{now}_pid{os.getpid()}",
        "started_at": now,
        "last_updated_at": now,
        "mode": mode,
        "input_file": input_file.name,
        "input_file_hash": compute_file_hash(input_file),
        "input_file_mtime": input_file.stat().st_mtime,
        "chunk_size": chunk_size,
        "total_tickets": len(tickets),
        "total_chunks": n_chunks,
        "batches_to_run": sorted(batches_to_run),
        "chunks": {},
    }


def is_chunk_completed(state: dict, chunk_idx: int) -> bool:
    """True si el chunk_idx ya está registrado como completado en el state."""
    return str(chunk_idx) in state.get("chunks", {})


def save_chunk_result(state: dict, chunk_idx: int, tickets: list,
                      result: tuple, duration_seconds: float,
                      input_file: Path, state_dir: Path = STATE_DIR) -> None:
    """
    Persiste el resultado de un chunk al state file de forma atómica.

    result debe ser la tupla retornada por process_chunk:
        (resultados, vl06f, zcmr_failures, ticket_to_order,
         cancel_failures, tickets_no_encontrados, tickets_con_factura,
         order_tracking)

    Acepta tuplas de 7 elementos (sin order_tracking) por retro-compatibilidad
    con código antiguo; en ese caso order_tracking se persiste como dict vacío.
    """
    if len(result) == 8:
        (resultados, vl06f, zcmr_failures, ticket_to_order, cancel_failures,
         no_enc, con_factura, order_tracking) = result
    else:
        (resultados, vl06f, zcmr_failures, ticket_to_order, cancel_failures,
         no_enc, con_factura) = result
        order_tracking = {}

    state["chunks"][str(chunk_idx)] = {
        "chunk_idx": chunk_idx,
        "tickets": list(tickets),
        "completed_at": datetime.now().isoformat(),
        "duration_seconds": round(duration_seconds, 2),
        "resultados": {k: [list(v[0]), list(v[1])] for k, v in resultados.items()},
        "vl06f": vl06f,
        "zcmr_failures": {k: list(v) for k, v in zcmr_failures.items()},
        "ticket_to_order": ticket_to_order,
        "cancel_failures": cancel_failures,
        "tickets_no_encontrados": list(no_enc),
        "tickets_con_factura": list(con_factura),
        "order_tracking": dict(order_tracking),
    }
    state["last_updated_at"] = datetime.now().isoformat()

    # Escritura atómica: write a tmp + rename. Path.replace() es atómico en
    # POSIX y Windows (sobreescribe el destino si existe).
    state_file = state_path_for_input(input_file, state_dir)
    tmp = state_file.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(state_file)


def consolidate_state(state: dict) -> tuple:
    """
    Combina todos los chunks completados en una sola tupla con el formato que
    espera el reporter (idéntico al return de process_chunk pero acumulado).

    Returns:
        (resultados, vl06f, zcmr_failures, ticket_to_order,
         cancel_failures, tickets_no_encontrados, tickets_con_factura,
         order_tracking)
    """
    resultados_total: dict = {}
    vl06f_total: dict = {}
    zcmr_failures_total: dict = {}
    ticket_to_order_total: dict = {}
    cancel_failures_total: dict = {}
    no_enc_total: list = []
    con_factura_total: list = []
    order_tracking_total: dict = {}

    for _, chunk_data in sorted(state["chunks"].items(), key=lambda x: int(x[0])):
        for batch_name, (exitosos, fallidos) in chunk_data["resultados"].items():
            if batch_name not in resultados_total:
                resultados_total[batch_name] = ([], [])
            resultados_total[batch_name] = (
                resultados_total[batch_name][0] + list(exitosos),
                resultados_total[batch_name][1] + list(fallidos),
            )
        vl06f_total.update(chunk_data["vl06f"])
        zcmr_failures_total.update(chunk_data["zcmr_failures"])
        ticket_to_order_total.update(chunk_data["ticket_to_order"])
        cancel_failures_total.update(chunk_data["cancel_failures"])
        no_enc_total.extend(chunk_data["tickets_no_encontrados"])
        con_factura_total.extend(chunk_data.get("tickets_con_factura", []))
        # order_tracking puede no existir en states viejos — fallback a {}
        order_tracking_total.update(chunk_data.get("order_tracking", {}))

    return (resultados_total, vl06f_total, zcmr_failures_total,
            ticket_to_order_total, cancel_failures_total, no_enc_total,
            con_factura_total, order_tracking_total)


def clear_state(input_file: Path, state_dir: Path = STATE_DIR) -> None:
    """Elimina el state file (para --fresh)."""
    state_file = state_path_for_input(input_file, state_dir)
    if state_file.exists():
        state_file.unlink()
        print(f"  [Checkpoint] State eliminado: {state_file}")
    else:
        print(f"  [Checkpoint] No había state previo que eliminar.")


BATCH_NAME_TO_ID = {
    "BATCH 1 — Billing Documents": 1,
    "BATCH 2 — Shipment Cost": 2,
    "BATCH 3 — Shipment Number": 3,
    "BATCH 4 — Reverse PGI": 4,
    "BATCH 5 — BOL Deletion": 5,
    "BATCH 6 — ZCMR Orders": 6,
    "BATCH 7 — Verificación Final ZSD": 7,
    "BATCH 8 — Order Cancellation": 8,
}


def get_failed_tickets_per_batch(chunk_state: dict,
                                  batches_filter: Optional[set] = None) -> dict:
    """
    Extrae los tickets fallidos por batch del state de un chunk.

    Args:
        chunk_state: dict de state["chunks"][str(idx)].
        batches_filter: set de IDs de batch (1-8) a considerar. None = todos.

    Returns:
        dict { batch_id (int): [tickets fallidos] } solo para batches con fallidos.
    """
    out: dict = {}
    for name, par in chunk_state.get("resultados", {}).items():
        bid = BATCH_NAME_TO_ID.get(name)
        if bid is None:
            continue
        if batches_filter and bid not in batches_filter:
            continue
        # par es [exitosos, fallidos] tras la deserialización del JSON
        fallidos = par[1] if len(par) > 1 else []
        if fallidos:
            out[bid] = list(fallidos)
    return out


def has_pending_work(chunk_state: dict,
                     batches_filter: Optional[set] = None) -> bool:
    """
    True si el chunk tiene fallidos en algún batch O tickets que la lectura
    inicial de VL06F no detectó. Útil para que --retry-failed identifique
    chunks con trabajo pendiente sin ignorar el blind spot histórico de
    tickets_no_encontrados.
    """
    if get_failed_tickets_per_batch(chunk_state, batches_filter):
        return True
    return bool(chunk_state.get("tickets_no_encontrados"))
