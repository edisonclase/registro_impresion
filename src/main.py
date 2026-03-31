from __future__ import annotations

import json
import re
from pathlib import Path

from .config import get_settings
from .print_setup import prepare_print_workbook


def ensure_directories(*paths: Path) -> None:
    for path in paths:
        path.mkdir(parents=True, exist_ok=True)


def build_output_stem(workbook_filename: str) -> str:
    """
    Genera un nombre base limpio a partir del archivo objetivo.
    Ejemplos:
    - 2A SEGUNDO A ...xlsx -> 2A_preparado_para_impresion_v4
    - 2B SEGUNDO B ...xlsx -> 2B_preparado_para_impresion_v4
    - 3A Tercero A ...xlsx -> 3A_preparado_para_impresion_v4
    """
    base_name = Path(workbook_filename).stem.strip()

    match = re.match(r"^([A-Za-z0-9]+)", base_name)
    prefix = match.group(1).upper() if match else "REGISTRO"

    return f"{prefix}_preparado_para_impresion_v9"


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

    output_stem = build_output_stem(settings.target_workbook)

    output_workbook = settings.output_xlsx_dir / f"{output_stem}.xlsx"

    actions_log = prepare_print_workbook(
        source_path=workbook_path,
        output_path=output_workbook,
        stop_sheet_name=settings.stop_sheet_name,
    )

    actions_json = settings.temp_dir / f"{output_stem}_log.json"
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
    print(f"Libro origen: {workbook_path.name}")
    print(f"Archivo generado: {output_workbook}")

    print("\n=== HOJAS CLAVE AJUSTADAS ===")
    shown = 0
    for item in actions_log:
        if item.detected_kind in {
            "asistencia_asignatura",
            "competencias",
            "cf",
            "ecap",
            "datos_estudiante",
            "datos_centro",
        }:
            print(f"- {item.sheet_name} | tipo={item.detected_kind}")
            for action in item.action_taken[:12]:
                print(f"    * {action}")
            shown += 1
            if shown >= 14:
                break

    print(f"\nLog generado: {actions_json}")


if __name__ == "__main__":
    main()