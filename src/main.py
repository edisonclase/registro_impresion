from __future__ import annotations

from pathlib import Path

from .config import get_settings
from .export_plan import save_reports
from .workbook_inspector import inspect_workbook


def ensure_directories(*paths: Path) -> None:
    for path in paths:
        path.mkdir(parents=True, exist_ok=True)


def main() -> None:
    settings = get_settings()

    ensure_directories(
        settings.output_xlsx_dir,
        settings.output_pdf_dir,
        settings.temp_dir,
    )

    workbook_path = settings.target_workbook_path
    pdf_path = settings.reference_pdf_path

    if not workbook_path.exists():
        raise FileNotFoundError(f"No se encontró el Excel objetivo: {workbook_path}")

    if not pdf_path.exists():
        raise FileNotFoundError(f"No se encontró el PDF de referencia: {pdf_path}")

    result = inspect_workbook(
        workbook_path=workbook_path,
        stop_sheet_name=settings.stop_sheet_name,
    )

    json_path, txt_path = save_reports(result, settings.temp_dir)

    print("\n=== INSPECCIÓN DEL LIBRO ===")
    print(f"Libro: {result.workbook_name}")
    print(f"Total de hojas: {result.total_sheets}")
    print(f"Hoja de corte detectada: {result.detected_stop_sheet}")
    print(f"Hojas válidas para impresión: {result.printable_sheet_count}")

    print("\n=== RESUMEN POR TIPO ===")
    for kind, count in sorted(result.summary_by_kind.items()):
        print(f"- {kind}: {count}")

    print("\n=== PRIMERAS HOJAS DE CALIFICACIONES DETECTADAS ===")
    count = 0
    for item in result.sheet_reports:
        if item.kind == "calificaciones":
            print(
                f"{item.index:>3}. {item.name} | filas={item.max_row} | "
                f"columnas={item.max_column} | orientación={item.orientation}"
            )
            count += 1
            if count == 15:
                break

    print(f"\nJSON detallado: {json_path}")
    print(f"Plan de impresión: {txt_path}")


if __name__ == "__main__":
    main()