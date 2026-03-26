from __future__ import annotations

import json
from pathlib import Path

from .config import get_settings
from .workbook_inspector import inspect_workbook, inspection_result_to_dict


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

    output_json = settings.temp_dir / "inspection_2A.json"
    output_json.write_text(
        json.dumps(inspection_result_to_dict(result), indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

    print("\n=== INSPECCIÓN DEL LIBRO ===")
    print(f"Libro: {result.workbook_name}")
    print(f"Total de hojas: {result.total_sheets}")
    print(f"Hoja de corte configurada: {result.stop_sheet_configured}")
    print(f"Hoja de corte detectada: {result.detected_stop_sheet}")
    print(f"Hojas válidas para impresión: {result.printable_sheet_count}")
    print(f"JSON generado en: {output_json}")

    print("\n=== PRIMERAS HOJAS VÁLIDAS ===")
    for item in result.sheet_reports[:15]:
        print(
            f"{item.index:>3}. {item.name} | tipo={item.kind} | "
            f"filas={item.max_row} | columnas={item.max_column} | "
            f"orientación={item.orientation} | print_area={item.print_area}"
        )

    if result.excluded_sheet_names:
        print("\n=== HOJAS EXCLUIDAS DESDE EL CORTE ===")
        for name in result.excluded_sheet_names[:15]:
            print(f"- {name}")


if __name__ == "__main__":
    main()