from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime


@dataclass(frozen=True)
class InvoiceTarget:
    sheet_name: str
    row: int
    status_col: int
    motivo_col: int


@dataclass(frozen=True)
class InvoiceTask:
    invoice: str
    targets: list[InvoiceTarget]


@dataclass(frozen=True)
class InvoiceResult:
    status: str
    motivo: str = ""


@dataclass
class ExcelPlan:
    file_path: str
    sheet_names: list[str]
    dt_from: datetime
    tasks: list[InvoiceTask]