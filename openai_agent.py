from __future__ import annotations

import json
import os

from ai_tools import (
    detect_sheet_mentions,
    explain_formula,
    find_direct_issue_hits,
    get_cell_details,
    list_issues,
    list_issues_for_sheet,
    search_labels,
    summarize_sheets,
    summarize_workbook,
    trace_cell,
)


SYSTEM_PROMPT = """
You are an Excel financial model audit assistant.
Answer only using the supplied workbook context and tool results.
Be concrete, cite cell references, and avoid guessing when evidence is missing.
Prefer concise analyst-style answers with:
- direct answer first
- supporting evidence second
- exact sheet/cell references when available
If the user asks for a conclusion, summarize the model's main risk areas and what to inspect next.
""".strip()


def is_openai_enabled(api_key: str | None = None) -> bool:
    if not (api_key or os.getenv("OPENAI_API_KEY")):
        return False
    try:
        from openai import OpenAI  # noqa: F401
    except ImportError:
        return False
    return True


def answer_with_openai(question: str, workbook, graph, issues, *, api_key: str | None = None, focus_sheet: str | None = None, model: str = "gpt-5.4-mini") -> str:
    from openai import OpenAI

    client = OpenAI(api_key=api_key or os.getenv("OPENAI_API_KEY"))
    tool_context = {
        "workbook": summarize_workbook(workbook, issues),
        "sheet_summaries": summarize_sheets(workbook, limit=12),
        "high_priority_issues": list_issues(issues, priority="High")[:10],
        "medium_priority_issues": list_issues(issues, priority="Medium")[:10],
        "label_matches": search_labels(workbook, question, limit=10),
    }

    if focus_sheet and focus_sheet != "None":
        tool_context["focus_sheet"] = focus_sheet
        tool_context["focus_sheet_issues"] = list_issues_for_sheet(issues, focus_sheet, limit=15)

    mentioned_sheets = detect_sheet_mentions(workbook, question)
    if mentioned_sheets:
        tool_context["sheet_issue_context"] = {
            sheet_name: list_issues_for_sheet(issues, sheet_name, limit=12)
            for sheet_name in mentioned_sheets[:3]
        }

    lower_question = question.lower()
    traced = None
    if "!" in question and any(keyword in lower_question for keyword in ["trace", "why", "where", "explain"]):
        target = next((token.strip(" ,.?") for token in question.split() if "!" in token), "")
        if target:
            traced = trace_cell(workbook, graph, target)
            tool_context["trace"] = traced
            tool_context["formula_explanation"] = explain_formula(workbook, target)
            tool_context["cell_details"] = get_cell_details(workbook, target)
            tool_context["direct_issue_hits"] = find_direct_issue_hits(issues, target)

    response = client.responses.create(
        model=model,
        input=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": [
                    {"type": "input_text", "text": question},
                    {
                        "type": "input_text",
                        "text": "Workbook context:\n" + json.dumps(tool_context, indent=2),
                    },
                ],
            },
        ],
    )
    return response.output_text
