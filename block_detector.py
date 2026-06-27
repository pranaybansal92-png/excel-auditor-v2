from __future__ import annotations

import re
from collections import defaultdict
from itertools import groupby

from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

from models import FormulaBlock, WorkbookModel

CELL_RE = re.compile(r"(?:(?:'[^']+'|[A-Za-z0-9_]+)!)?\$?([A-Z]{1,3})\$?(\d{1,7})")


def _normalize_formula(formula: str, base_address: str) -> str:
    base_col_letter, base_row = coordinate_from_string(base_address)
    base_col = column_index_from_string(base_col_letter)

    def replace_ref(match: re.Match[str]) -> str:
        col_letter, row_text = match.groups()
        col = column_index_from_string(col_letter)
        row = int(row_text)
        return f"R[{row - base_row}]C[{col - base_col}]"

    return CELL_RE.sub(replace_ref, formula)


def detect_formula_blocks(workbook: WorkbookModel, min_block_size: int = 4) -> list[FormulaBlock]:
    row_groups: dict[tuple[str, int], list[str]] = defaultdict(list)
    col_groups: dict[tuple[str, int], list[str]] = defaultdict(list)

    for key, cell in workbook.formulas.items():
        col_letter, row_idx = coordinate_from_string(cell.address)
        col_idx = column_index_from_string(col_letter)
        row_groups[(cell.sheet, row_idx)].append(cell.address)
        col_groups[(cell.sheet, col_idx)].append(cell.address)

    blocks: list[FormulaBlock] = []
    blocks.extend(_build_blocks(workbook, row_groups, "row", min_block_size))
    blocks.extend(_build_blocks(workbook, col_groups, "column", min_block_size))
    return blocks


def _build_blocks(
    workbook: WorkbookModel,
    groups: dict[tuple[str, int], list[str]],
    orientation: str,
    min_block_size: int,
) -> list[FormulaBlock]:
    blocks: list[FormulaBlock] = []
    for (sheet, _), addresses in groups.items():
        ordered = sorted(addresses, key=_sort_key)
        for segment in _contiguous_segments(ordered, orientation):
            if len(segment) < min_block_size:
                continue
            normalized = [
                _normalize_formula(workbook.formulas[f"{sheet}!{address}"].formula or "", address)
                for address in segment
            ]
            blocks.append(
                FormulaBlock(
                    sheet=sheet,
                    orientation=orientation,
                    anchor=segment[0],
                    cells=segment,
                    normalized_formulas=normalized,
                )
            )
    return blocks


def _sort_key(address: str) -> tuple[int, int]:
    col_letter, row_idx = coordinate_from_string(address)
    return row_idx, column_index_from_string(col_letter)


def _contiguous_segments(addresses: list[str], orientation: str) -> list[list[str]]:
    if not addresses:
        return []

    indexed: list[tuple[int, str]] = []
    for address in addresses:
        col_letter, row_idx = coordinate_from_string(address)
        col_idx = column_index_from_string(col_letter)
        axis = col_idx if orientation == "row" else row_idx
        indexed.append((axis, address))

    segments: list[list[str]] = []
    current: list[str] = [indexed[0][1]]
    previous_axis = indexed[0][0]
    for axis, address in indexed[1:]:
        if axis == previous_axis + 1:
            current.append(address)
        else:
            segments.append(current)
            current = [address]
        previous_axis = axis
    segments.append(current)
    return segments
