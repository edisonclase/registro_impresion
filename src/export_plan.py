from __future__ import annotations

import json
from pathlib import Path

from .workbook_inspector import WorkbookInspectionResult


def build_text_report(result: WorkbookInspectionResult) -> str:
    lines: list[str] = []

    lines.append("PLAN INICIAL DE IMPRESIÓN - 2A")
    lines.append("=" * 60)
    lines.append(f"Libro: {result.workbook_name}")
    lines.append(f"Ruta: {result.workbook_path}")
    lines.append(f"Total de hojas del libro: {result.total_sheets}")
    lines.append(f"Hoja de corte detectada: {result.detected_stop_sheet}")
    lines.append(f"Hojas válidas antes del corte: {result.printable_sheet_count}")
    lines.append("")

    lines.append("RESUMEN POR TIPO")
    lines.append("-" * 60)
    for kind, count in sorted(result.summary_by_kind.items()):
        lines.append(f"{kind}: {count}")
    lines.append("")

    lines.append("HOJAS CANDIDATAS A REVISIÓN ESPECIAL")
    lines.append("-" * 60)
    for item in result.sheet_reports:
        if item.kind == "calificaciones" or not item.likely_printable:
            lines.append(
                f"{item.index:>3}. {item.name} | tipo={item.kind} | "
                f"filas={item.max_row} | columnas={item.max_column} | "
                f"imprimible={item.likely_printable}"
            )
            for note in item.notes:
                lines.append(f"     - {note}")
    lines.append("")

    lines.append("PRÓXIMAS REGLAS TÉCNICAS")
    lines.append("-" * 60)
    lines.append("1. No modificar textos oficiales.")
    lines.append("2. Mantener celdas vacías visibles dentro de las tablas.")
    lines.append("3. En hojas de calificaciones: excluir nombres, conservar número, encabezados y notas.")
    lines.append("4. Preparar impresión horizontal en carta para hojas amplias de calificaciones.")
    lines.append("5. Ajustar áreas de impresión sin alterar contenido.")
    lines.append("")

    return "\n".join(lines)


def save_reports(result: WorkbookInspectionResult, temp_dir: Path) -> tuple[Path, Path]:
    temp_dir.mkdir(parents=True, exist_ok=True)

    json_path = temp_dir / "inspection_2A_detailed.json"
    txt_path = temp_dir / "plan_inicial_impresion_2A.txt"

    from .workbook_inspector import inspection_result_to_dict

    json_path.write_text(
        json.dumps(inspection_result_to_dict(result), indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    txt_path.write_text(
        build_text_report(result),
        encoding="utf-8",
    )

    return json_path, txt_path