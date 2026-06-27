from __future__ import annotations

from ai_tools import (
    detect_sheet_mentions,
    explain_formula,
    find_direct_issue_hits,
    get_cell_details,
    list_issues,
    list_issues_for_detector,
    list_issues_for_sheet,
    search_labels,
    summarize_sheets,
    summarize_issues_for_sheet,
    summarize_workbook,
    trace_cell,
)


def answer_question(question: str, workbook, graph, issues) -> str:
    text = question.strip()
    text_lower = text.lower()
    mentioned_sheets = detect_sheet_mentions(workbook, text)

    cell_match = None
    for token in text.replace(",", " ").replace("?", " ").split():
        cleaned = token.strip()
        if "!" in cleaned:
            cell_match = cleaned
            break
        if len(cleaned) >= 2 and cleaned[0].isalpha() and any(ch.isdigit() for ch in cleaned):
            cell_match = cleaned
            break

    if "high priority" in text_lower:
        matches = list_issues(issues, priority="High")
        if not matches:
            return "No high priority issues were found in this pass."
        return "\n".join(
            f"- {row['sheet']} {row['cells']}: {row['title']}"
            for row in matches[:10]
        )

    if "how many sheet" in text_lower:
        return f"The workbook has {workbook.total_sheets} sheets."

    if mentioned_sheets and ("summary" in text_lower or "summarize" in text_lower or "sheet" in text_lower):
        sheet_name = mentioned_sheets[0]
        rows = [row for row in summarize_sheets(workbook, limit=50) if row["sheet"] == sheet_name]
        if rows:
            row = rows[0]
            return (
                f"{sheet_name} has {row['rows']} rows, {row['cols']} columns, "
                f"{row['formulas']} formulas, {row['text']} text cells, and {row['numeric']} numeric cells."
            )

    if "summary" in text_lower or "summarize" in text_lower or "overview" in text_lower:
        summary = summarize_workbook(workbook, issues)
        return (
            f"{summary['filename']} looks like a {summary['type_hint']} model with "
            f"{summary['sheet_count']} sheets, {summary['formula_count']} formulas, "
            f"and {summary['issue_count']} flagged issues."
        )

    if "most formulas" in text_lower or "largest sheets" in text_lower:
        rows = summarize_sheets(workbook, limit=5)
        return "\n".join(
            f"- {row['sheet']}: {row['formulas']} formulas"
            for row in rows
        )

    if "conclusion" in text_lower or "bottom line" in text_lower or "takeaway" in text_lower:
        summary = summarize_workbook(workbook, issues)
        top_high = list_issues(issues, priority="High")[:3]
        if top_high:
            risk_line = "; ".join(f"{row['sheet']} {row['cells']}: {row['title']}" for row in top_high)
            return (
                f"This looks like a {summary['type_hint']} model with {summary['formula_count']} formulas. "
                f"The main concern is concentrated in these flagged areas: {risk_line}."
            )
        return (
            f"This looks like a {summary['type_hint']} model with {summary['formula_count']} formulas "
            "and no current high-priority issues."
        )

    if "detector" in text_lower:
        for detector in ["formula_inconsistency", "hardcoded_numbers", "hardcoded_jump", "text_math", "tax_anomaly", "balance_check", "dsra_logic", "label_semantics"]:
            if detector in text_lower:
                matches = list_issues_for_detector(issues, detector)
                if not matches:
                    return f"No issues found for detector `{detector}`."
                return "\n".join(
                    f"- {row['sheet']} {row['cells']}: {row['title']}"
                    for row in matches[:10]
                )

    if "trace" in text_lower and cell_match:
        target = cell_match if "!" in cell_match else f"{mentioned_sheets[0]}!{cell_match}" if mentioned_sheets else ""
        if target:
            traced = trace_cell(workbook, graph, target)
            return f"{traced['target']} depends on {len(traced['precedents'])} precedent cells: {', '.join(traced['precedents'][:12]) or 'none found'}."

    if "explain" in text_lower and cell_match:
        target = cell_match if "!" in cell_match else f"{mentioned_sheets[0]}!{cell_match}" if mentioned_sheets else ""
        if target:
            direct_hits = find_direct_issue_hits(issues, target)
            if direct_hits:
                top_hit = direct_hits[0]
                return (
                    f"{target} is flagged as `{top_hit['title']}`. "
                    f"{top_hit['detail']} Detector: {top_hit['detector']}."
                )
            return explain_formula(workbook, target)

    if "show" in text_lower and cell_match:
        target = cell_match if "!" in cell_match else f"{mentioned_sheets[0]}!{cell_match}" if mentioned_sheets else ""
        if target:
            details = get_cell_details(workbook, target)
            if not details["found"]:
                return "I could not find that cell in the workbook map."
            return f"{target}: value={details['value']} formula={details['formula'] or 'none'} type={details['data_type']}"

    if mentioned_sheets and ("issue" in text_lower or "problem" in text_lower or "flag" in text_lower):
        if "explain" in text_lower or "why" in text_lower:
            sheet_name = mentioned_sheets[0]
            summary = summarize_issues_for_sheet(issues, sheet_name)
            if not summary["count"]:
                return f"No issues found on {sheet_name}."
            top_titles = ", ".join(title for title, _ in summary["top_titles"][:3])
            return (
                f"{sheet_name} has {summary['count']} flagged issues. "
                f"Priority mix: {summary['priority_counts']}. "
                f"Most common issue types: {summary['detector_counts']}. "
                f"Main findings include: {top_titles}."
            )
        matches = list_issues_for_sheet(issues, mentioned_sheets[0])
        if not matches:
            return f"No issues found on {mentioned_sheets[0]}."
        return "\n".join(
            f"- {row['priority']} {row['cells']}: {row['title']}"
            for row in matches[:12]
        )

    if mentioned_sheets and ("error" in text_lower or "errors" in text_lower):
        matches = list_issues_for_sheet(issues, mentioned_sheets[0])
        if not matches:
            return f"No issues found on {mentioned_sheets[0]}."
        return "\n".join(
            f"- {row['priority']} {row['cells']}: {row['title']}"
            for row in matches[:12]
        )

    label_hits = search_labels(workbook, text)
    if label_hits:
        preview = ", ".join(hit["cell"] for hit in label_hits[:5])
        return f"I found matching labels at {preview}."

    return (
        "I can help with workbook overview, sheet summaries, issue lookup by sheet, label search, and cell tracing."
    )
