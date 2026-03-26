from __future__ import annotations

import json
from pathlib import Path

from .config import get_settings
from .print_setup import prepare_print_workbook


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

    output_workbook = settings.output_xlsx_dir / "2A_preparado_para_impresion.xlsx"

    actions_log = prepare_print_workbook(
        source_path=workbook_path,
        output_path=output_workbook,
        stop_sheet_name="REPORTE CALIFICACI",
    )

    actions_json = settings.temp_dir / "2A_preparado_para_impresion_log.json"
    actions_json.write_text(
        json.dumps(
            [
                {
                    "sheet_name": item.sheet_name,
                    "detected_kind": item.detected_kind,
                    "action_taken": item.action_taken,
                }
                for item in actions_log
            ],
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    print("\n=== COPIA DE TRABAJO GENERADA ===")
    print(f"Archivo generado: {output_workbook}")

    print("\n=== PRIMERAS ACCIONES APLICADAS ===")
    shown = 0
    for item in actions_log:
        if item.detected_kind in {"competencias", "asistencia_asignatura", "datos_estudiante", "datos_centro"}:
            print(f"- {item.sheet_name} | tipo={item.detected_kind}")
            for action in item.action_taken:
                print(f"    * {action}")
            shown += 1
            if shown >= 12:
                break

    print(f"\nLog generado: {actions_json}")


if __name__ == "__main__":
    main()