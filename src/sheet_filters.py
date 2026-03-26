from __future__ import annotations

import re
import unicodedata
from typing import Iterable, Optional


def normalize_text(value: str) -> str:
    if value is None:
        return ""

    value = str(value).strip().upper()
    value = unicodedata.normalize("NFKD", value)
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = re.sub(r"\s+", " ", value)
    return value


def is_stop_sheet(sheet_name: str, configured_stop_name: str) -> bool:
    current = normalize_text(sheet_name)
    configured = normalize_text(configured_stop_name)

    if configured and configured in current:
        return True

    fallback_patterns = [
        "REPORTE CALIFICACI",
        "REPORTES DE CALIFICACIONES",
        "REPORTE DE CALIFICACIONES",
    ]

    return any(pattern in current for pattern in fallback_patterns)


def get_printable_sheet_names(sheet_names: Iterable[str], configured_stop_name: str) -> list[str]:
    printable: list[str] = []

    for sheet_name in sheet_names:
        if is_stop_sheet(sheet_name, configured_stop_name):
            break
        printable.append(sheet_name)

    return printable


def find_stop_sheet(sheet_names: Iterable[str], configured_stop_name: str) -> Optional[str]:
    for sheet_name in sheet_names:
        if is_stop_sheet(sheet_name, configured_stop_name):
            return sheet_name
    return None


def is_student_data_sheet(sheet_name: str) -> bool:
    n = normalize_text(sheet_name)
    return "DATOS DEL ESTUDIANTE" in n


def is_center_data_sheet(sheet_name: str) -> bool:
    n = normalize_text(sheet_name)
    return "DATOS DEL CENTRO" in n


def is_completivo_sheet(sheet_name: str) -> bool:
    n = normalize_text(sheet_name)
    return "COMPLETIVO" in n


def is_extraordinario_sheet(sheet_name: str) -> bool:
    n = normalize_text(sheet_name)
    return "EXTRAORDINARIO" in n


def is_acta_sheet(sheet_name: str) -> bool:
    n = normalize_text(sheet_name)
    return "ACTA" in n


def is_attendance_sheet(sheet_name: str) -> bool:
    n = normalize_text(sheet_name)
    return "ASIT" in n or "ASIST" in n


def looks_like_grade_sheet_by_name(sheet_name: str) -> bool:
    """
    Hojas como ALE216, MAT218, SOC220, etc.
    Suelen ser hojas amplias con calificaciones.
    """
    n = normalize_text(sheet_name)
    return bool(re.fullmatch(r"[A-ZÑ]{2,8}\d{2,5}", n))


def classify_sheet(sheet_name: str) -> str:
    if is_center_data_sheet(sheet_name):
        return "datos_centro"

    if is_student_data_sheet(sheet_name):
        return "datos_estudiante"

    if is_attendance_sheet(sheet_name):
        return "asistencia"

    if is_completivo_sheet(sheet_name):
        return "completivo"

    if is_extraordinario_sheet(sheet_name):
        return "extraordinario"

    if is_acta_sheet(sheet_name):
        return "acta"

    if looks_like_grade_sheet_by_name(sheet_name):
        return "calificaciones"

    return "general"