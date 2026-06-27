from __future__ import annotations

from collections import Counter

from dependency_graph import trace_precedents
from models import Issue, WorkbookModel


def search_labels(workbook: WorkbookModel, query: str, limit: int = 20) -> list[dict]:
    query_lower = query.lower()
    matches = []
    for key, label in workbook.labels.items():
        if query_lower in label.lower():
            matches.append({"cell": key, "label": label})
    return matches[:limit]


def trace_cell(workbook: WorkbookModel, graph, target: str) -> dict:
    precedents = trace_precedents(graph, target, depth=3)
    formula = workbook.formulas.get(target)
    return {
        "target": target,
        "formula": formula.formula if formula else None,
        "precedents": precedents,
    }


def list_issues(issues: list[Issue], priority: str | None = None) -> list[dict]:
    rows = []
    for issue in issues:
        if priority and issue.priority.lower() != priority.lower():
            continue
        rows.append(
            {
                "title": issue.title,
                "priority": issue.priority,
                "sheet": issue.sheet,
                "cells": ", ".join(issue.cells),
                "detail": issue.detail,
            }
        )
    return rows


def summarize_workbook(workbook: WorkbookModel, issues: list[Issue]) -> dict:
    by_detector = Counter(issue.detector for issue in issues)
    by_priority = Counter(issue.priority for issue in issues)
    return {
        "filename": workbook.filename,
        "type_hint": workbook.workbook_type_hint,
        "sheet_count": workbook.total_sheets,
        "formula_count": workbook.total_formulas,
        "issue_count": len(issues),
        "issue_priority_counts": dict(by_priority),
        "issue_detector_counts": dict(by_detector),
    }


def summarize_sheets(workbook: WorkbookModel, limit: int = 12) -> list[dict]:
    ranked = sorted(workbook.sheets, key=lambda item: item.formula_count, reverse=True)
    return [
        {
            "sheet": sheet.name,
            "rows": sheet.max_row,
            "cols": sheet.max_col,
            "formulas": sheet.formula_count,
            "text": sheet.text_count,
            "numeric": sheet.numeric_count,
        }
        for sheet in ranked[:limit]
    ]


def list_issues_for_sheet(issues: list[Issue], sheet_name: str, limit: int = 20) -> list[dict]:
    rows = []
    for issue in issues:
        if issue.sheet.lower() != sheet_name.lower():
            continue
        rows.append(
            {
                "priority": issue.priority,
                "title": issue.title,
                "cells": ", ".join(issue.cells),
                "detector": issue.detector,
                "detail": issue.detail,
            }
        )
    return rows[:limit]


def summarize_issues_for_sheet(issues: list[Issue], sheet_name: str) -> dict:
    matches = [issue for issue in issues if issue.sheet.lower() == sheet_name.lower()]
    by_priority = Counter(issue.priority for issue in matches)
    by_detector = Counter(issue.detector for issue in matches)
    top_titles = Counter(issue.title for issue in matches).most_common(5)
    return {
        "sheet": sheet_name,
        "count": len(matches),
        "priority_counts": dict(by_priority),
        "detector_counts": dict(by_detector),
        "top_titles": top_titles,
    }


def list_issues_for_detector(issues: list[Issue], detector: str, limit: int = 20) -> list[dict]:
    rows = []
    for issue in issues:
        if issue.detector.lower() != detector.lower():
            continue
        rows.append(
            {
                "priority": issue.priority,
                "sheet": issue.sheet,
                "cells": ", ".join(issue.cells),
                "title": issue.title,
                "detail": issue.detail,
            }
        )
    return rows[:limit]


def detect_sheet_mentions(workbook: WorkbookModel, text: str) -> list[str]:
    lowered = text.lower()
    return [sheet.name for sheet in workbook.sheets if sheet.name.lower() in lowered]


def explain_formula(workbook: WorkbookModel, target: str) -> str:
    cell = workbook.formulas.get(target)
    if not cell:
        return "No formula found for that cell."
    return f"{target} contains {cell.formula}. The next step is to trace its precedents and compare it with peer formulas."


def get_cell_details(workbook: WorkbookModel, target: str) -> dict:
    cell = workbook.cells.get(target)
    if not cell:
        return {"target": target, "found": False}

    return {
        "target": target,
        "found": True,
        "sheet": cell.sheet,
        "address": cell.address,
        "value": cell.value,
        "formula": cell.formula,
        "data_type": cell.data_type,
    }


def find_direct_issue_hits(issues: list[Issue], target: str) -> list[dict]:
    hits: list[dict] = []
    for issue in issues:
        issue_targets = [f"{issue.sheet}!{cell}" for cell in issue.cells]
        if target not in issue_targets:
            continue
        hits.append(
            {
                "priority": issue.priority,
                "title": issue.title,
                "detail": issue.detail,
                "detector": issue.detector,
                "cells": issue_targets,
                "evidence": issue.evidence,
            }
        )
    return hits
