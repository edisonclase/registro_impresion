from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet


@dataclass
class SheetPrintAction:
    sheet_name: str
    detected_kind: str
    action_taken: list[str]


def cell_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip().upper()


def safe_last_col_letter(ws: Worksheet) -> str:
    """
    Devuelve la letra de la última columna usada sin depender
    de una celda concreta, evitando errores con MergedCell.
    """
    max_col = max(1, ws.max_column or 1)
    return get_column_letter(max_col)


def detect_sheet_kind_by_content(ws: Worksheet) -> str:
    """
    Detecta el tipo de hoja usando el contenido visible,
    no solo el nombre de la pestaña.
    """
    a1 = cell_text(ws["A1"].value)
    b1 = cell_text(ws["B1"].value)
    d1 = cell_text(ws["D1"].value)

    a2 = cell_text(ws["A2"].value)
    b2 = cell_text(ws["B2"].value)

    d4 = cell_text(ws["D4"].value)
    c6 = cell_text(ws["C6"].value)

    if "DATOS DEL CENTRO" in b1:
        return "datos_centro"

    if "DATOS DEL ESTUDIANTE" in b1:
        return "datos_estudiante"

    if "COMPETENCIA" in a2 and "NOMBRE" in b2:
        return "competencias"

    if "DÍAS TRABAJADOS" in d4 and "NOMBRE" in c6:
        return "asistencia_asignatura"

    if "DÍAS TRABAJADOS" in d4 and "NOMBRE" in c6 and ws.max_column <= 50:
        return "asistencia_print"

    if "COMPLETIVO" in a1 or "COMPLETIVO" in b1 or "COMPLETIVO" in d1:
        return "completivo"

    if "EXTRAORDINARIO" in a1 or "EXTRAORDINARIO" in b1 or "EXTRAORDINARIO" in d1:
        return "extraordinario"

    return "general"


def set_common_page_setup_letter(ws: Worksheet, landscape: bool) -> None:
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.orientation = "landscape" if landscape else "portrait"

    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1 if landscape else 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.page_margins.left = 0.20
    ws.page_margins.right = 0.20
    ws.page_margins.top = 0.30
    ws.page_margins.bottom = 0.30
    ws.page_margins.header = 0.15
    ws.page_margins.footer = 0.15

    ws.print_options.horizontalCentered = True
    ws.print_options.verticalCentered = False


def configure_competency_sheet_for_print(ws: Worksheet) -> list[str]:
    actions: list[str] = []

    set_common_page_setup_letter(ws, landscape=True)
    actions.append("orientación horizontal en carta")
    actions.append("ajuste a 1 página de ancho")

    ws.column_dimensions["B"].hidden = True
    actions.append("columna B oculta para no imprimir nombres")

    last_col_letter = safe_last_col_letter(ws)
    ws.print_area = f"A1:{last_col_letter}{ws.max_row}"
    actions.append(f"área de impresión definida: A1:{last_col_letter}{ws.max_row}")

    ws.print_title_rows = "$1:$4"
    actions.append("filas 1 a 4 repetidas como encabezado")

    return actions


def configure_text_sheet_for_print(ws: Worksheet) -> list[str]:
    actions: list[str] = []
    set_common_page_setup_letter(ws, landscape=False)
    actions.append("orientación vertical en carta")

    last_col_letter = safe_last_col_letter(ws)
    ws.print_area = f"A1:{last_col_letter}{ws.max_row}"
    actions.append(f"área de impresión definida: A1:{last_col_letter}{ws.max_row}")
    return actions


def configure_attendance_sheet_for_print(ws: Worksheet) -> list[str]:
    actions: list[str] = []
    set_common_page_setup_letter(ws, landscape=True)
    actions.append("orientación horizontal en carta")
    actions.append("ajuste a 1 página de ancho")

    last_col_letter = safe_last_col_letter(ws)
    ws.print_area = f"A1:{last_col_letter}{ws.max_row}"
    actions.append(f"área de impresión definida: A1:{last_col_letter}{ws.max_row}")

    ws.print_title_rows = "$1:$6"
    actions.append("filas 1 a 6 repetidas como encabezado")
    return actions


def prepare_print_workbook(
    source_path: Path,
    output_path: Path,
    stop_sheet_name: Optional[str] = None,
) -> list[SheetPrintAction]:
    wb = load_workbook(source_path)

    actions_log: list[SheetPrintAction] = []
    stop_detected = False

    for ws in wb.worksheets:
        sheet_name_upper = (ws.title or "").upper()

        if stop_sheet_name and stop_sheet_name.upper() in sheet_name_upper:
            stop_detected = True

        if stop_detected:
            actions_log.append(
                SheetPrintAction(
                    sheet_name=ws.title,
                    detected_kind="excluida_post_corte",
                    action_taken=["sin cambios en esta fase"],
                )
            )
            continue

        kind = detect_sheet_kind_by_content(ws)

        if kind == "competencias":
            actions = configure_competency_sheet_for_print(ws)
        elif kind in {"datos_centro", "datos_estudiante", "completivo", "extraordinario"}:
            actions = configure_text_sheet_for_print(ws)
        elif kind in {"asistencia_asignatura", "asistencia_print"}:
            actions = configure_attendance_sheet_for_print(ws)
        else:
            actions = ["sin cambios automáticos en esta fase"]

        actions_log.append(
            SheetPrintAction(
                sheet_name=ws.title,
                detected_kind=kind,
                action_taken=actions,
            )
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)

    return actions_log