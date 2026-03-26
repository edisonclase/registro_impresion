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
    """
    Detecta la hoja de corte de manera flexible.
    Soporta variaciones como:
    - Reportes de Calificaciones
    - Reporte de Calificaciones
    - REPORTE CALIFICACI
    """
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


def classify_sheet(sheet_name: str) -> str:
    normalized = normalize_text(sheet_name)

    if "DATOS DEL CENTRO" in normalized:
        return "datos_centro"

    if "DATOS DEL ESTUDIANTE" in normalized:
        return "datos_estudiante"

    if "COMPLETIVO" in normalized:
        return "completivo"

    if "EXTRAORDINARIO" in normalized:
        return "extraordinario"

    if "ACTA" in normalized:
        return "acta"

    if "ASIST" in normalized:
        return "asistencia"

    # Luego afinaremos esta clasificación con reglas más precisas
    return "general"