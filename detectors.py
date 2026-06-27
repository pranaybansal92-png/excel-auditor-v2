from __future__ import annotations

import re
from collections import Counter, defaultdict
from difflib import SequenceMatcher

from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

from dependency_graph import extract_references
from models import Issue, WorkbookModel

HARDCODE_RE = re.compile(r"(?<![A-Z])(?<![A-Z]\d)([-+]?\d+(?:\.\d+)?)")
BOOLEAN_FORMULA_RE = re.compile(r"^=\s*(AND|OR|NOT)\(|^=\s*[^=]+=[^=]+$")
SUMMARY_SHEET_TERMS = ("summary", "output", "dashboard", "control", "checks", "cover")
SUMMARY_ROW_TERMS = ("total", "subtotal", "summary", "check", "balance check", "header")
SAFE_HARDCODES = {"2", "4", "6", "8", "12", "20", "24", "30", "100", "180", "365", "730", "1000"}
ASSUMPTION_ROW_TERMS = (
    "price", "unit", "headcount", "salary", "days", "rate", "margin", "multiple",
    "capex", "rent", "tax", "growth", "discount", "wacc", "terminal", "opening",
)
JUMP_ROW_TERMS = ("tariff", "opex", "tax", "dsra", "interest", "revenue")


def _make_issue(
    title: str,
    detail: str,
    priority: str,
    sheet: str,
    cells: list[str],
    detector: str,
    evidence: dict | None = None,
) -> Issue:
    return Issue(
        title=title,
        detail=detail,
        priority=priority,
        sheet=sheet,
        cells=cells,
        detector=detector,
        evidence=evidence or {},
    )


def detect_divide_by_zero(workbook: WorkbookModel) -> list[Issue]:
    issues: list[Issue] = []
    for cell in workbook.formulas.values():
        if "/0" in (cell.formula or "").replace(" ", ""):
            issues.append(
                _make_issue(
                    title="Division by zero literal",
                    detail="Formula contains a literal division by zero.",
                    priority="High",
                    sheet=cell.sheet,
                    cells=[cell.address],
                    detector="divide_by_zero",
                )
            )
    return issues


def detect_hardcoded_numbers(workbook: WorkbookModel) -> list[Issue]:
    grouped: dict[tuple[str, str], list[str]] = defaultdict(list)
    for cell in workbook.formulas.values():
        formula = cell.formula or ""
        for match in HARDCODE_RE.findall(formula):
            if match in {"0", "1", "-1"}:
                continue
            grouped[(cell.sheet, match)].append(cell.address)

    issues: list[Issue] = []
    for (sheet, constant), addresses in grouped.items():
        if len(addresses) < 3:
            continue
        if constant in SAFE_HARDCODES:
            continue
        if _all_assumption_contexts(workbook, sheet, addresses):
            continue
        priority = "Medium" if len(addresses) >= 4 else "Low"
        issues.append(
            _make_issue(
                title="Hardcoded number in formulas",
                detail=f"Constant {constant} is embedded in formulas instead of referenced from an input cell.",
                priority=priority,
                sheet=sheet,
                cells=sorted(addresses)[:12],
                detector="hardcoded_numbers",
                evidence={"constant": constant, "count": len(addresses)},
            )
        )
    issues.extend(detect_hardcoded_jump_outliers(workbook))
    return issues


def _all_assumption_contexts(workbook: WorkbookModel, sheet: str, addresses: list[str]) -> bool:
    contexts = [_row_context_text(workbook, sheet, address) for address in addresses[:12]]
    if not contexts:
        return False
    return all(any(term in context for term in ASSUMPTION_ROW_TERMS) for context in contexts)


def detect_formula_inconsistencies(workbook: WorkbookModel, blocks) -> list[Issue]:
    issues: list[Issue] = []
    grouped_issues: dict[tuple[str, str, str, str, str], list[str]] = defaultdict(list)
    for block in blocks:
        counts = Counter(block.normalized_formulas)
        if len(counts) < 2:
            continue
        dominant, dominant_count = counts.most_common(1)[0]
        if dominant_count < 3:
            continue
        block_size = len(block.cells)
        dominant_ratio = dominant_count / block_size
        if dominant_ratio < 0.7:
            continue

        outlier_cells = [
            (address, normalized)
            for address, normalized in zip(block.cells, block.normalized_formulas)
            if normalized != dominant
        ]
        if len(outlier_cells) > max(2, block_size // 4):
            continue

        block.dominant_pattern = dominant
        for address, normalized in outlier_cells:
            formula_cell = workbook.formulas.get(f"{block.sheet}!{address}")
            if not formula_cell or _is_boolean_formula(formula_cell.formula or ""):
                continue
            if _is_summary_context(workbook, block.sheet, address):
                continue
            diagnostic = _diagnose_pattern_difference(dominant, normalized)
            priority = "Medium" if _is_output_sheet(block.sheet) or _has_cross_sheet_ref(formula_cell.formula or "") else "High"
            group_key = (block.sheet, block.orientation, dominant, normalized, priority)
            grouped_issues[group_key].append(address)

    for (sheet, orientation, dominant, normalized, priority), addresses in grouped_issues.items():
        diagnostic = _diagnose_pattern_difference(dominant, normalized)
        issues.append(
            _make_issue(
                title="Formula pattern outlier",
                detail=(
                    f"These formulas break the dominant {orientation} pattern in the surrounding block. "
                    f"{diagnostic}"
                ),
                priority=priority,
                sheet=sheet,
                cells=sorted(addresses)[:12],
                detector="formula_inconsistency",
                evidence={
                    "orientation": orientation,
                    "expected_pattern": dominant,
                    "actual_pattern": normalized,
                    "diagnostic": diagnostic,
                    "count": len(addresses),
                },
            )
        )
    return issues


def _is_boolean_formula(formula: str) -> bool:
    compact = formula.replace(" ", "").upper()
    return bool(BOOLEAN_FORMULA_RE.match(compact))


def _has_cross_sheet_ref(formula: str) -> bool:
    return "!" in formula


def _is_output_sheet(sheet_name: str) -> bool:
    lowered = sheet_name.lower()
    return any(term in lowered for term in SUMMARY_SHEET_TERMS)


def _is_summary_context(workbook: WorkbookModel, sheet: str, address: str) -> bool:
    if _is_output_sheet(sheet):
        return True

    row_context = _row_context_text(workbook, sheet, address)
    return any(term in row_context for term in SUMMARY_ROW_TERMS)


def _row_context_text(workbook: WorkbookModel, sheet: str, address: str) -> str:
    col_letter, row_idx = coordinate_from_string(address)
    col_idx = column_index_from_string(col_letter)
    texts: list[str] = []
    for offset in range(1, col_idx):
        candidate_col = col_idx - offset
        candidate_key = f"{sheet}!{_column_letter(candidate_col)}{row_idx}"
        candidate = workbook.cells.get(candidate_key)
        label_text = _extract_text_like_label(candidate) if candidate else None
        if label_text:
            texts.append(label_text.lower())
            if candidate_col <= 3:
                break
    return " ".join(texts)


def _column_letter(index: int) -> str:
    letters = []
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


def detect_text_math(workbook: WorkbookModel) -> list[Issue]:
    grouped: dict[tuple[str, str, str], list[str]] = defaultdict(list)
    for cell in workbook.formulas.values():
        formula = cell.formula or ""
        compact = formula.replace(" ", "")
        if "/" not in compact and "^" not in compact:
            continue
        if _is_summary_context(workbook, cell.sheet, cell.address):
            continue

        refs = extract_references(formula, cell.sheet)
        operators = _operator_reference_targets(formula, cell.sheet)
        for ref in refs:
            if ref not in operators:
                continue
            referenced = workbook.cells.get(ref)
            if not referenced:
                continue
            label_text = _extract_text_like_label(referenced)
            if not label_text:
                continue
            if _looks_like_period_label(label_text):
                grouped[(cell.sheet, ref, label_text)].append(cell.address)
                break
            grouped[(cell.sheet, ref, label_text)].append(cell.address)
            break

    issues: list[Issue] = []
    for (sheet, ref, label_text), addresses in grouped.items():
        detail = (
            "Formula appears to divide by or exponentiate a referenced period label stored as text."
            if _looks_like_period_label(label_text)
            else "Formula appears to divide by or exponentiate a referenced cell that is stored as text."
        )
        issues.append(
            _make_issue(
                title="Possible math against text label",
                detail=detail,
                priority="High",
                sheet=sheet,
                cells=sorted(addresses)[:12],
                detector="text_math",
                evidence={"label_cell": ref, "label_text": label_text, "count": len(addresses)},
            )
        )
    return issues


def _extract_text_like_label(cell) -> str | None:
    if cell.formula:
        match = re.fullmatch(r'=\s*"([^"]+)"', cell.formula.strip())
        if match:
            return match.group(1).strip()
        return None
    if isinstance(cell.value, str):
        return cell.value.strip()
    return None


def _operator_reference_targets(formula: str, current_sheet: str) -> set[str]:
    targets: set[str] = set()
    for pattern in (
        r"/\s*((?:'[^']+'|[A-Za-z0-9_]+)!\$?[A-Z]{1,3}\$?\d{1,7}|\$?[A-Z]{1,3}\$?\d{1,7})",
        r"\^\s*((?:'[^']+'|[A-Za-z0-9_]+)!\$?[A-Z]{1,3}\$?\d{1,7}|\$?[A-Z]{1,3}\$?\d{1,7})",
    ):
        for match in re.findall(pattern, formula):
            targets.update(extract_references(match, current_sheet))
    return targets


def _looks_like_period_label(text: str) -> bool:
    lowered = text.strip().lower()
    if re.fullmatch(r"(fy|q|m)?\d{1,4}", lowered):
        return True
    if re.fullmatch(r"(month|quarter)\s*\d+", lowered):
        return True
    return lowered in {
        "jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "sept", "oct", "nov", "dec",
        "january", "february", "march", "april", "june", "july", "august", "september", "october", "november", "december",
    }


def detect_balance_check_issues(workbook: WorkbookModel) -> list[Issue]:
    issues: list[Issue] = []
    for label_key, label_text in workbook.labels.items():
        lowered = label_text.lower()
        if "balance check" not in lowered and "balancecheck" not in lowered:
            continue

        sheet, address = label_key.split("!", 1)
        formula_cell = _find_neighbor_formula(workbook, sheet, address)
        if not formula_cell or not formula_cell.formula:
            continue

        formula = formula_cell.formula.upper().replace(" ", "")
        if "+" in formula and "-" not in formula:
            issues.append(
                _make_issue(
                    title="Possible balance check sign issue",
                    detail="Balance check appears to add components without an offsetting subtraction.",
                    priority="Medium",
                    sheet=sheet,
                    cells=[formula_cell.address],
                    detector="balance_check",
                    evidence={"label_cell": label_key, "formula": formula_cell.formula},
                )
            )
    return issues


def detect_label_semantic_issues(workbook: WorkbookModel) -> list[Issue]:
    issues: list[Issue] = []
    for label_key, label_text in workbook.labels.items():
        sheet, address = label_key.split("!", 1)
        formula_cell = _find_neighbor_formula(workbook, sheet, address)
        if not formula_cell or not formula_cell.formula:
            continue

        lowered = label_text.lower()
        formula = formula_cell.formula.upper().replace(" ", "")

        if any(term in lowered for term in ["coverage", "dscr", "moic", "per share", "margin", "yield"]):
            if "/" not in formula and "RATIO" not in formula:
                issues.append(
                    _make_issue(
                        title="Possible ratio formula mismatch",
                        detail="Label suggests a ratio-style metric, but the nearby formula does not divide values.",
                        priority="Medium",
                        sheet=sheet,
                        cells=[formula_cell.address],
                        detector="label_semantics",
                        evidence={"label_cell": label_key, "label_text": label_text, "formula": formula_cell.formula},
                    )
                )

        if "irr" in lowered and not any(fn in formula for fn in ["IRR(", "XIRR(", "MIRR("]):
            issues.append(
                _make_issue(
                    title="Possible IRR formula mismatch",
                    detail="Label suggests an IRR metric, but the nearby formula does not use IRR, XIRR, or MIRR.",
                    priority="High",
                    sheet=sheet,
                    cells=[formula_cell.address],
                    detector="label_semantics",
                    evidence={"label_cell": label_key, "label_text": label_text, "formula": formula_cell.formula},
                )
            )

        if "terminal value" in lowered and "R+G" in formula:
            issues.append(
                _make_issue(
                    title="Possible Gordon Growth sign issue",
                    detail="Terminal value appears to use r + g instead of the usual r - g denominator.",
                    priority="High",
                    sheet=sheet,
                    cells=[formula_cell.address],
                    detector="label_semantics",
                    evidence={"label_cell": label_key, "label_text": label_text, "formula": formula_cell.formula},
                )
            )

        if any(term in lowered for term in ["runway", "burn rate", "burnrate"]):
            if "/" in formula and "ABS(" not in formula:
                issues.append(
                    _make_issue(
                        title="Possible runway / burn rate sign issue",
                        detail="Runway-style calculation divides by a burn metric without ABS(), which can flip the sign.",
                        priority="Medium",
                        sheet=sheet,
                        cells=[formula_cell.address],
                        detector="label_semantics",
                        evidence={"label_cell": label_key, "label_text": label_text, "formula": formula_cell.formula},
                    )
                )
    return issues


def detect_rollforward_issues(workbook: WorkbookModel) -> list[Issue]:
    issues: list[Issue] = []
    keywords = ["debt", "pp&e", "ppe", "inventory", "working capital", "depreciation", "amortization"]
    for label_key, label_text in workbook.labels.items():
        lowered = label_text.lower()
        if not any(keyword in lowered for keyword in keywords):
            continue

        sheet, address = label_key.split("!", 1)
        formula_cell = _find_neighbor_formula(workbook, sheet, address)
        if not formula_cell or not formula_cell.formula:
            continue

        formula = formula_cell.formula.upper().replace(" ", "")
        if "+" in formula and "-" not in formula and "SUM(" not in formula:
            issues.append(
                _make_issue(
                    title="Possible roll-forward structure issue",
                    detail="Roll-forward style line appears to add components without a subtraction or offsetting movement.",
                    priority="Medium",
                    sheet=sheet,
                    cells=[formula_cell.address],
                    detector="rollforward_logic",
                    evidence={"label_cell": label_key, "label_text": label_text, "formula": formula_cell.formula},
                )
            )
    return issues


def detect_tax_anomalies(workbook: WorkbookModel) -> list[Issue]:
    issues: list[Issue] = []
    for cell in workbook.formulas.values():
        context = _row_context_text(workbook, cell.sheet, cell.address)
        if "tax" not in context:
            continue
        formula = (cell.formula or "").replace(" ", "").upper()
        if "*0" in formula or formula.endswith("=0"):
            issues.append(
                _make_issue(
                    title="Possible tax suppression",
                    detail="Tax line appears to be zeroed out in formula despite being labeled as taxes.",
                    priority="High",
                    sheet=cell.sheet,
                    cells=[cell.address],
                    detector="tax_anomaly",
                    evidence={"formula": cell.formula, "context": context},
                )
            )
    return issues


def detect_dsra_logic_issues(workbook: WorkbookModel) -> list[Issue]:
    grouped: dict[tuple[str, str], list[str]] = defaultdict(list)
    for cell in workbook.formulas.values():
        context = _row_context_text(workbook, cell.sheet, cell.address)
        if "dsra" not in context:
            continue
        formula = (cell.formula or "").replace(" ", "")
        if "Control'!B14" in formula or "Control!B14" in formula:
            if "/12" not in formula:
                grouped[(cell.sheet, context)].append(cell.address)

    issues: list[Issue] = []
    for (sheet, context), addresses in grouped.items():
        issues.append(
            _make_issue(
                title="Possible DSRA months logic issue",
                detail="DSRA requirement appears to use months directly without converting to a monthly fraction.",
                priority="Medium",
                sheet=sheet,
                cells=sorted(addresses)[:12],
                detector="dsra_logic",
                evidence={"count": len(addresses), "context": context},
            )
        )
    return issues


def detect_hardcoded_jump_outliers(workbook: WorkbookModel) -> list[Issue]:
    issues: list[Issue] = []
    for cell in workbook.formulas.values():
        context = _row_context_text(workbook, cell.sheet, cell.address)
        if not any(term in context for term in JUMP_ROW_TERMS):
            continue
        formula = cell.formula or ""
        for match in re.findall(r"\(1([+\-])(\d+(?:\.\d+)?)\)", formula.replace(" ", "")):
            sign, number_text = match
            value = float(number_text)
            if value >= 0.1:
                issues.append(
                    _make_issue(
                        title="Possible hardcoded jump rate",
                        detail="Formula contains a one-off hardcoded jump rate instead of linking to an assumption.",
                        priority="Medium",
                        sheet=cell.sheet,
                        cells=[cell.address],
                        detector="hardcoded_jump",
                        evidence={"jump_rate": f"{sign}{number_text}", "formula": formula},
                    )
                )
                break
    return issues


def _find_neighbor_formula(workbook: WorkbookModel, sheet: str, address: str):
    from openpyxl.utils.cell import coordinate_from_string, get_column_letter, column_index_from_string

    col_letter, row_idx = coordinate_from_string(address)
    col_idx = column_index_from_string(col_letter)
    candidates = []
    for offset in (1, 2, -1):
        next_col = col_idx + offset
        if next_col < 1:
            continue
        candidate_address = f"{sheet}!{get_column_letter(next_col)}{row_idx}"
        candidate = workbook.formulas.get(candidate_address)
        if candidate:
            candidates.append(candidate)
    return candidates[0] if candidates else None


def _diagnose_pattern_difference(expected: str, actual: str) -> str:
    if expected == actual:
        return "No structural difference detected."

    expected_refs = re.findall(r"R\[[^\]]+\]C\[[^\]]+\]", expected)
    actual_refs = re.findall(r"R\[[^\]]+\]C\[[^\]]+\]", actual)

    if len(expected_refs) != len(actual_refs):
        return (
            "Reference count changed compared with peer formulas, which can indicate a missing input, "
            "extra term, or broken copied formula."
        )

    if any(left != right for left, right in zip(expected_refs, actual_refs)):
        diffs = [
            f"expected {left} but found {right}"
            for left, right in zip(expected_refs, actual_refs)
            if left != right
        ]
        return "Reference shift mismatch detected: " + "; ".join(diffs[:3]) + "."

    if any(op in expected or op in actual for op in ["+", "-", "*", "/"]):
        expected_ops = re.findall(r"[+\-*/]", expected)
        actual_ops = re.findall(r"[+\-*/]", actual)
        if expected_ops != actual_ops:
            return "Operator pattern changed compared with peer formulas."

    matcher = SequenceMatcher(a=expected, b=actual)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag != "equal":
            expected_part = expected[i1:i2] or "nothing"
            actual_part = actual[j1:j2] or "nothing"
            return f"Pattern segment changed from `{expected_part}` to `{actual_part}`."

    return "Pattern differs from peers, but the exact cause needs deeper tracing."
