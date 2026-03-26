from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv


BASE_DIR = Path(__file__).resolve().parent.parent
load_dotenv(BASE_DIR / ".env")


def _resolve_path(value: str) -> Path:
    path = Path(value)
    if not path.is_absolute():
        path = BASE_DIR / path
    return path.resolve()


@dataclass(frozen=True)
class Settings:
    project_name: str
    python_env: str
    input_excel_dir: Path
    input_pdf_dir: Path
    output_xlsx_dir: Path
    output_pdf_dir: Path
    temp_dir: Path
    log_level: str
    stop_sheet_name: str
    target_workbook: str
    reference_pdf: str

    @property
    def target_workbook_path(self) -> Path:
        return (self.input_excel_dir / self.target_workbook).resolve()

    @property
    def reference_pdf_path(self) -> Path:
        return (self.input_pdf_dir / self.reference_pdf).resolve()


def get_settings() -> Settings:
    return Settings(
        project_name=os.getenv("PROJECT_NAME", "registro_impresion"),
        python_env=os.getenv("PYTHON_ENV", "development"),
        input_excel_dir=_resolve_path(os.getenv("INPUT_EXCEL_DIR", "data/input/excel")),
        input_pdf_dir=_resolve_path(os.getenv("INPUT_PDF_DIR", "data/input/pdf_referencia")),
        output_xlsx_dir=_resolve_path(os.getenv("OUTPUT_XLSX_DIR", "data/output/xlsx_ajustados")),
        output_pdf_dir=_resolve_path(os.getenv("OUTPUT_PDF_DIR", "data/output/pdf_generados")),
        temp_dir=_resolve_path(os.getenv("TEMP_DIR", "data/temp")),
        log_level=os.getenv("LOG_LEVEL", "INFO").upper(),
        stop_sheet_name=os.getenv("STOP_SHEET_NAME", "Reportes de Calificaciones"),
        target_workbook=os.getenv(
            "TARGET_WORKBOOK",
            "2A SEGUNDO A 2025 - 2026 REGISTRO DE GRADO CEJOMA ORD. 04 2023.xlsx",
        ),
        reference_pdf=os.getenv(
            "REFERENCE_PDF",
            "Registro-2do-Grado-Sec-General-1-1.pdf",
        ),
    )