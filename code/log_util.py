"""
Observabilidad: redirección de stdout/stderr a consola + archivo de log.

Centraliza el 'tee' que antes estaba duplicado en main.py y diagnose_grids.py.
Robusto a encoding: la consola de Windows (cp1252) no puede codificar algunos
caracteres (flechas, checks y similares); en vez de tragarse la línea entera
(pérdida silenciosa), reemplaza el caracter problemático SOLO en la consola.
El archivo de log se abre en UTF-8, así que ahí los caracteres quedan intactos.
"""
import sys
import time

from path import BASE_DIR


class _Tee:
    """Duplica la escritura a varios streams (consola + archivo)."""

    def __init__(self, *streams):
        self._streams = [s for s in streams if s is not None]

    def write(self, data):
        for s in self._streams:
            try:
                s.write(data)
            except UnicodeEncodeError:
                # La consola no puede codificar algún caracter: reemplazar en
                # lugar de perder la línea completa.
                enc = getattr(s, "encoding", None) or "ascii"
                try:
                    s.write(data.encode(enc, "replace").decode(enc, "replace"))
                except Exception:
                    pass
            except Exception:
                pass
            try:
                s.flush()
            except Exception:
                pass

    def flush(self):
        for s in self._streams:
            try:
                s.flush()
            except Exception:
                pass

    def isatty(self):
        for s in self._streams:
            try:
                return s.isatty()
            except Exception:
                continue
        return False


def setup_logging(prefix: str = "dil_run"):
    """
    Redirige stdout y stderr a la consola Y a un archivo timestamped en
    Data-bases/Logs/<prefix>_<ts>.log. No modifica ningún print existente.
    Retorna la ruta del log (o None si no se pudo crear).
    """
    try:
        logs_dir = BASE_DIR / "Data-bases" / "Logs"
        logs_dir.mkdir(parents=True, exist_ok=True)
        ts = time.strftime("%Y%m%d_%H%M%S")
        log_path = logs_dir / f"{prefix}_{ts}.log"
        log_file = open(log_path, "a", encoding="utf-8")

        # Evitar que la consola tire UnicodeEncodeError y se pierdan líneas.
        for stream in (sys.stdout, sys.stderr):
            try:
                stream.reconfigure(errors="replace")
            except Exception:
                pass

        # Usar sys.stdout/err ACTUALES (la consola real), no sys.__stdout__,
        # que puede ser un pipe con encoding estricto.
        sys.stdout = _Tee(sys.stdout, log_file)
        sys.stderr = _Tee(sys.stderr, log_file)
        print(f"[Log] Salida guardándose en: {log_path}")
        return log_path
    except Exception as e:
        print(f"[Log] WARNING: no se pudo iniciar logging a archivo: {e}",
              file=sys.stderr)
        return None
