from datetime import date
from pathlib import Path

# Raíz del proyecto
BASE_DIR = Path(__file__).resolve().parent.parent

# Carpeta raíz de entradas (contiene subcarpetas por mes, e.g. "May-2026")
ENTRADAS_DIR = BASE_DIR / "Data-bases" / "Entradas"

# Carpeta raíz de salidas (misma estructura: subcarpetas por mes)
SALIDAS_DIR = BASE_DIR / "Data-bases" / "Salidas"


def _latest_subdir(parent: Path) -> Path:
    """Retorna la subcarpeta más reciente (por fecha de modificación) dentro de parent."""
    subdirs = sorted(
        [p for p in parent.iterdir() if p.is_dir()],
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not subdirs:
        raise FileNotFoundError(f"No se encontraron subcarpetas en: {parent}")
    return subdirs[0]


def _xlsx_files(folder: Path) -> list:
    """Lista .xlsx ordenados por mtime descendente, excluyendo lock files de Office (~$...)."""
    return sorted(
        [p for p in folder.glob("*.xlsx") if not p.name.startswith("~$")],
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )


def get_report_file() -> Path:
    """
    Retorna la ruta donde guardar el reporte xlsx de salida.

    - Entra a la subcarpeta más reciente de SALIDAS_DIR.
    - Devuelve la ruta del archivo más reciente si ya existe alguno,
      o genera uno nuevo con la fecha de hoy si la carpeta está vacía.

    Lanza FileNotFoundError si no existe ninguna subcarpeta en SALIDAS_DIR.
    """
    folder = _latest_subdir(SALIDAS_DIR)
    xlsx_files = _xlsx_files(folder)
    if xlsx_files:
        return xlsx_files[0]
    hoy = date.today().strftime("%b %d %Y")
    return folder / f"Reporte DIL - {hoy}.xlsx"


def get_billing_file(filename: str = None) -> Path:
    """
    Retorna la ruta al archivo Excel de billing.

    - Si se pasa `filename`, lo busca dentro de la subcarpeta más reciente de ENTRADAS_DIR.
    - Si no se pasa nada, entra a la subcarpeta más reciente y retorna el .xlsx más reciente
      (ignorando lock files de Office `~$...`).

    Lanza FileNotFoundError si no encuentra subcarpetas o archivos .xlsx.
    """
    folder = _latest_subdir(ENTRADAS_DIR)

    if filename:
        path = folder / filename
        if not path.exists():
            raise FileNotFoundError(f"Archivo no encontrado: {path}")
        return path

    xlsx_files = _xlsx_files(folder)
    if not xlsx_files:
        raise FileNotFoundError(f"No se encontraron archivos .xlsx en: {folder}")

    return xlsx_files[0]
