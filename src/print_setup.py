from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from copy import copy

from openpyxl import load_workbook
from openpyxl.styles import Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet


@dataclass
class SheetPrintAction:
    sheet_name: str
    detected_kind: str
    action_taken: list[str]


THIN_BLACK_SIDE = Side(style="thin", color="FF000000")
THIN_BLACK_BORDER = Border(
    left=THIN_BLACK_SIDE,
    right=THIN_BLACK_SIDE,
    top=THIN_BLACK_SIDE,
    bottom=THIN_BLACK_SIDE,
)

WHITE_FILL = PatternFill(
    fill_type="solid",
    fgColor="FFFFFFFF",
    bgColor="FFFFFFFF",
)


def cell_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip().upper()


def safe_last_col_letter(ws: Worksheet) -> str:
    max_col = max(1, ws.max_column or 1)
    return get_column_letter(max_col)


def clear_conditional_formatting(ws: Worksheet) -> list[str]:
    """
    Elimina reglas de formato condicional.
    Esto es importante porque Google Sheets / Excel online pueden
    volver a pintar celdas en rojo aunque el fill y la fuente ya hayan sido cambiados.
    """
    removed = 0

    try:
        removed = len(ws.conditional_formatting)
    except Exception:
        removed = 0

    try:
        ws.conditional_formatting._cf_rules.clear()
    except Exception:
        pass

    if removed:
        return [f"formato condicional eliminado: {removed} reglas"]
    return ["sin formato condicional que eliminar"]


def font_to_times_new_roman_black(original_font: Font) -> Font:
    """
    Fuerza fuente Times New Roman 12 en negro.
    """
    font_copy = copy(original_font)
    font_copy.name = "Times New Roman"
    font_copy.size = 12
    font_copy.color = "FF000000"
    return font_copy


def white_fill() -> PatternFill:
    return copy(WHITE_FILL)


def normalize_cell_visual_style(cell) -> tuple[bool, bool]:
    """
    Reglas visuales globales:
    - Times New Roman 12
    - texto/números en negro
    - rellenos a blanco
    """
    font_changed = False
    fill_changed = False

    try:
        cell.font = font_to_times_new_roman_black(cell.font)
        font_changed = True
    except Exception:
        pass

    try:
        fill = cell.fill
        fill_type = getattr(fill, "fill_type", None)
        fg_rgb = getattr(getattr(fill, "fgColor", None), "rgb", None)
        bg_rgb = getattr(getattr(fill, "bgColor", None), "rgb", None)

        has_colored_fill = (
            fill_type is not None
            or fg_rgb not in (None, "00000000", "00FFFFFF", "FFFFFFFF")
            or bg_rgb not in (None, "00000000", "00FFFFFF", "FFFFFFFF")
        )

        if has_colored_fill:
            cell.fill = white_fill()
            fill_changed = True
    except Exception:
        pass

    return font_changed, fill_changed


def normalize_sheet_visual_style(ws: Worksheet) -> list[str]:
    changed_fill_count = 0
    changed_font_count = 0

    for row in ws.iter_rows():
        for cell in row:
            if cell is None:
                continue

            font_changed, fill_changed = normalize_cell_visual_style(cell)

            if font_changed:
                changed_font_count += 1
            if fill_changed:
                changed_fill_count += 1

    notes = ["fuente global forzada a Times New Roman 12 y color negro"]
    if changed_fill_count:
        notes.append(f"rellenos forzados a blanco: {changed_fill_count}")
    if changed_font_count:
        notes.append(f"fuentes forzadas a negro: {changed_font_count}")
    return notes


def autosize_columns_by_content(
    ws: Worksheet,
    *,
    min_width: float = 8.5,
    max_width: float = 45.0,
) -> list[str]:
    changed = 0

    for col_idx in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        max_len = 0

        for row_idx in range(1, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            value = cell.value
            if value is None:
                continue

            text = str(value).replace("\n", " ").strip()
            if not text:
                continue

            max_len = max(max_len, len(text))

        if max_len == 0:
            continue

        proposed = min(max(min_width, max_len * 0.95), max_width)

        current_width = ws.column_dimensions[col_letter].width
        if current_width is None or proposed > current_width:
            ws.column_dimensions[col_letter].width = proposed
            changed += 1

    return [f"ancho de columnas ajustado según contenido: {changed} columnas"]


def complete_used_range_borders(ws: Worksheet) -> list[str]:
    applied = 0
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            if cell is None:
                continue
            try:
                cell.border = copy(THIN_BLACK_BORDER)
                applied += 1
            except Exception:
                pass

    return [f"bordes completos aplicados en rango usado: {applied} celdas"]


def detect_sheet_kind_by_content(ws: Worksheet) -> str:
    title = cell_text(ws.title)

    a1 = cell_text(ws["A1"].value)
    b1 = cell_text(ws["B1"].value)
    d1 = cell_text(ws["D1"].value)

    a2 = cell_text(ws["A2"].value)
    b2 = cell_text(ws["B2"].value)

    d4 = cell_text(ws["D4"].value)
    c6 = cell_text(ws["C6"].value)

    row4 = " | ".join(cell_text(ws.cell(4, c).value) for c in range(1, min(ws.max_column, 12) + 1))

    if "DATOS DEL CENTRO" in b1:
        return "datos_centro"

    if "DATOS DEL ESTUDIANTE" in b1:
        return "datos_estudiante"

    if title.startswith("ECAP"):
        return "ecap"

    if title.startswith("CF-") or "C.F." in row4:
        return "cf"

    if "COMPETENCIA" in a2 and "NOMBRE" in b2:
        return "competencias"

    if "DÍAS TRABAJADOS" in d4 and "NOMBRE" in c6:
        return "asistencia_asignatura"

    if "COMPLETIVO" in a1 or "COMPLETIVO" in b1 or "COMPLETIVO" in d1:
        return "completivo"

    if "EXTRAORDINARIO" in a1 or "EXTRAORDINARIO" in b1 or "EXTRAORDINARIO" in d1:
        return "extraordinario"

    return "general"


def set_common_page_setup_letter_portrait(ws: Worksheet) -> None:
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.page_margins.left = 0.20
    ws.page_margins.right = 0.20
    ws.page_margins.top = 0.30
    ws.page_margins.bottom = 0.30
    ws.page_margins.header = 0.15
    ws.page_margins.footer = 0.15

    ws.print_options.horizontalCentered = True
    ws.print_options.verticalCentered = False


def set_print_area_full_used_range(ws: Worksheet) -> str:
    last_col_letter = safe_last_col_letter(ws)
    ws.print_area = f"A1:{last_col_letter}{ws.max_row}"
    return f"A1:{last_col_letter}{ws.max_row}"


def column_has_header_text(ws: Worksheet, col_idx: int, texts: set[str], max_header_row: int = 8) -> bool:
    normalized_targets = {t.replace(" ", "") for t in texts}

    for row_idx in range(1, min(ws.max_row, max_header_row) + 1):
        value = ws.cell(row=row_idx, column=col_idx).value
        text = cell_text(value)
        compact = text.replace(" ", "")
        if text in texts or compact in normalized_targets:
            return True
    return False


def find_cf_attendance_columns(ws: Worksheet, max_header_row: int = 8) -> tuple[list[int], list[int]]:
    """
    Detecta en hojas CF las columnas del bloque ASISTENCIA.
    Debe dejar visible solo la columna % ANUAL / %ANUAL.
    """
    attendance_cols: list[int] = []
    annual_cols: list[int] = []

    attendance_markers = {
        "P1",
        "P2",
        "P3",
        "P4",
        "% ANUAL",
        "%ANUAL",
        "ASISTENCIA",
    }

    annual_markers = {
        "% ANUAL",
        "%ANUAL",
    }

    normalized_attendance = {m.replace(" ", "") for m in attendance_markers}
    normalized_annual = {m.replace(" ", "") for m in annual_markers}

    for col_idx in range(1, ws.max_column + 1):
        matched_texts: set[str] = set()

        for row_idx in range(1, min(ws.max_row, max_header_row) + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            text = cell_text(value)
            compact = text.replace(" ", "")
            if text:
                matched_texts.add(text)
                matched_texts.add(compact)

        if matched_texts.intersection(attendance_markers.union(normalized_attendance)):
            attendance_cols.append(col_idx)

        if matched_texts.intersection(annual_markers.union(normalized_annual)):
            annual_cols.append(col_idx)

    return attendance_cols, annual_cols


def hide_cf_attendance_details_except_annual(ws: Worksheet) -> list[str]:
    """
    En CF deja visible solo la columna % ANUAL del bloque de asistencia.
    Mantiene visible el encabezado general ASISTENCIA porque está en celdas combinadas,
    pero oculta las columnas P1, P2, P3 y P4.
    """
    actions: list[str] = []

    attendance_cols, annual_cols = find_cf_attendance_columns(ws)

    if not attendance_cols:
        return ["no se detectó bloque de asistencia en hoja CF"]

    if not annual_cols:
        return ["se detectó bloque de asistencia, pero no se encontró columna % ANUAL"]

    annual_set = set(annual_cols)
    hidden_count = 0

    for col_idx in attendance_cols:
        if col_idx not in annual_set:
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].hidden = True
            hidden_count += 1

    visible_letters = [get_column_letter(c) for c in annual_cols]
    actions.append(f"bloque de asistencia depurado en CF: {hidden_count} columnas ocultas")
    actions.append(f"columna(s) % ANUAL visibles: {', '.join(visible_letters)}")
    return actions


def configure_competency_sheet_for_print(ws: Worksheet) -> list[str]:
    actions: list[str] = []

    set_common_page_setup_letter_portrait(ws)
    actions.append("orientación vertical en carta")
    actions.append("ajuste a 1 página de ancho")

    ws.column_dimensions["B"].hidden = True
    actions.append("columna B oculta para no imprimir nombres")

    print_area = set_print_area_full_used_range(ws)
    actions.append(f"área de impresión definida: {print_area}")

    ws.print_title_rows = "$1:$4"
    actions.append("filas 1 a 4 repetidas como encabezado")

    actions.extend(autosize_columns_by_content(ws, min_width=8.5, max_width=30.0))
    return actions


def configure_cf_sheet_for_print(ws: Worksheet) -> list[str]:
    actions: list[str] = []

    set_common_page_setup_letter_portrait(ws)
    actions.append("orientación vertical en carta")
    actions.append("ajuste a 1 página de ancho")

    ws.column_dimensions["B"].hidden = True
    ws.column_dimensions["D"].hidden = True
    actions.append("columnas B y D ocultas para no imprimir ID ni nombre")

    actions.extend(hide_cf_attendance_details_except_annual(ws))

    print_area = set_print_area_full_used_range(ws)
    actions.append(f"área de impresión definida: {print_area}")

    ws.print_title_rows = "$1:$6"
    actions.append("filas 1 a 6 repetidas como encabezado")

    actions.extend(autosize_columns_by_content(ws, min_width=8.5, max_width=24.0))
    return actions


def configure_attendance_sheet_for_print(ws: Worksheet) -> list[str]:
    actions: list[str] = []

    set_common_page_setup_letter_portrait(ws)
    actions.append("orientación vertical en carta")
    actions.append("ajuste a 1 página de ancho")

    for col in ["A", "C", "AA", "AB", "AC", "AD"]:
        ws.column_dimensions[col].hidden = True

    for col in ["AF", "BD", "BE", "BF", "BG"]:
        ws.column_dimensions[col].hidden = True

    actions.append("columnas ocultas en asistencia: ID, nombre, TA, %A, TE, %E")

    print_area = set_print_area_full_used_range(ws)
    actions.append(f"área de impresión definida: {print_area}")

    ws.print_title_rows = "$1:$6"
    actions.append("filas 1 a 6 repetidas como encabezado")

    actions.extend(autosize_columns_by_content(ws, min_width=4.5, max_width=18.0))
    return actions


def configure_text_sheet_for_print(ws: Worksheet) -> list[str]:
    actions: list[str] = []
    set_common_page_setup_letter_portrait(ws)
    actions.append("orientación vertical en carta")

    print_area = set_print_area_full_used_range(ws)
    actions.append(f"área de impresión definida: {print_area}")

    actions.extend(autosize_columns_by_content(ws, min_width=8.5, max_width=28.0))
    return actions


def configure_ecap_sheet_for_print(ws: Worksheet) -> list[str]:
    actions: list[str] = []
    set_common_page_setup_letter_portrait(ws)
    actions.append("orientación vertical en carta")

    print_area = set_print_area_full_used_range(ws)
    actions.append(f"área de impresión definida: {print_area}")

    actions.extend(complete_used_range_borders(ws))
    actions.extend(autosize_columns_by_content(ws, min_width=12.0, max_width=42.0))
    return actions


def prepare_print_workbook(
    source_path: Path,
    output_path: Path,
    stop_sheet_name: Optional[str] = None,
) -> list[SheetPrintAction]:
    wb = load_workbook(source_path)

    actions_log: list[SheetPrintAction] = []
    stop_detected = False

    stop_tokens = ["REPORTE CALIFICACI"]
    if stop_sheet_name:
        stop_tokens.append(stop_sheet_name.upper())

    for ws in wb.worksheets:
        title_upper = (ws.title or "").upper()

        if any(token in title_upper for token in stop_tokens):
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
        actions: list[str] = []

        actions.extend(clear_conditional_formatting(ws))
        actions.extend(normalize_sheet_visual_style(ws))

        if kind == "competencias":
            actions.extend(configure_competency_sheet_for_print(ws))
        elif kind == "cf":
            actions.extend(configure_cf_sheet_for_print(ws))
        elif kind == "ecap":
            actions.extend(configure_ecap_sheet_for_print(ws))
        elif kind == "asistencia_asignatura":
            actions.extend(configure_attendance_sheet_for_print(ws))
        else:
            actions.extend(configure_text_sheet_for_print(ws))

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