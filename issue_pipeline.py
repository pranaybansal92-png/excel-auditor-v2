from __future__ import annotations

from block_detector import detect_formula_blocks
from detectors import (
    detect_balance_check_issues,
    detect_divide_by_zero,
    detect_dsra_logic_issues,
    detect_formula_inconsistencies,
    detect_hardcoded_numbers,
    detect_tax_anomalies,
    detect_label_semantic_issues,
    detect_rollforward_issues,
    detect_text_math,
)
from models import Issue, WorkbookModel

_PRIORITY_ORDER = {"High": 0, "Medium": 1, "Low": 2}


def run_issue_pipeline(workbook: WorkbookModel) -> tuple[list[Issue], list]:
    blocks = detect_formula_blocks(workbook)
    issues: list[Issue] = []
    issues.extend(detect_formula_inconsistencies(workbook, blocks))
    issues.extend(detect_divide_by_zero(workbook))
    issues.extend(detect_hardcoded_numbers(workbook))
    issues.extend(detect_text_math(workbook))
    issues.extend(detect_balance_check_issues(workbook))
    issues.extend(detect_label_semantic_issues(workbook))
    issues.extend(detect_rollforward_issues(workbook))
    issues.extend(detect_tax_anomalies(workbook))
    issues.extend(detect_dsra_logic_issues(workbook))
    deduped = _dedupe_issues(issues)
    deduped.sort(key=lambda item: (_PRIORITY_ORDER.get(item.priority, 99), item.sheet, item.cells[0]))
    return deduped, blocks


def _dedupe_issues(issues: list[Issue]) -> list[Issue]:
    by_key: dict[tuple[str, str, tuple[str, ...]], Issue] = {}
    for issue in issues:
        key = (issue.detector, issue.sheet, tuple(issue.cells))
        by_key[key] = issue
    return list(by_key.values())
