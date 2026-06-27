from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(slots=True)
class CellRecord:
    sheet: str
    address: str
    value: Any
    formula: str | None = None
    data_type: str | None = None

    @property
    def key(self) -> str:
        return f"{self.sheet}!{self.address}"


@dataclass(slots=True)
class SheetSummary:
    name: str
    max_row: int
    max_col: int
    formula_count: int = 0
    value_count: int = 0
    text_count: int = 0
    numeric_count: int = 0


@dataclass(slots=True)
class Issue:
    title: str
    detail: str
    priority: str
    sheet: str
    cells: list[str]
    detector: str
    evidence: dict[str, Any] = field(default_factory=dict)


@dataclass(slots=True)
class FormulaBlock:
    sheet: str
    orientation: str
    anchor: str
    cells: list[str]
    normalized_formulas: list[str]
    dominant_pattern: str | None = None


@dataclass(slots=True)
class WorkbookModel:
    filename: str
    sheets: list[SheetSummary]
    cells: dict[str, CellRecord]
    formulas: dict[str, CellRecord]
    labels: dict[str, str]
    named_ranges: dict[str, str]
    workbook_type_hint: str = "general"

    @property
    def total_formulas(self) -> int:
        return len(self.formulas)

    @property
    def total_sheets(self) -> int:
        return len(self.sheets)
