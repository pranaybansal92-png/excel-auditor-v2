from __future__ import annotations

import os

import streamlit as st

from ai_tools import find_direct_issue_hits, get_cell_details, summarize_sheets
from chat_agent import answer_question
from dependency_graph import build_dependency_graph, trace_dependents, trace_precedents
from issue_pipeline import run_issue_pipeline
from openai_agent import answer_with_openai, is_openai_enabled
from workbook_parser import parse_workbook

st.set_page_config(page_title="Excel Auditor V2", layout="wide")


def _priority_counts(issues):
    counts = {"High": 0, "Medium": 0, "Low": 0}
    for issue in issues:
        counts[issue.priority] = counts.get(issue.priority, 0) + 1
    return counts


def _check_app_access() -> bool:
    required_password = os.getenv("APP_PASSWORD", "").strip()
    if not required_password:
        return True

    if st.session_state.get("app_unlocked"):
        return True

    st.title("Excel Auditor V2")
    st.caption("Private beta access")
    password = st.text_input("Enter app password", type="password")
    if st.button("Unlock"):
        if password == required_password:
            st.session_state["app_unlocked"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False


def main() -> None:
    if not _check_app_access():
        return

    st.title("Excel Auditor V2")
    st.caption("Structure-first auditing and AI-ready workbook navigation.")

    with st.sidebar:
        st.subheader("AI Settings")
        api_key = st.text_input("OpenAI API Key", type="password")
        model_choice = st.selectbox("Model", ["gpt-5.4-mini", "gpt-5.4", "gpt-5.2"], index=0)

    uploaded = st.file_uploader("Upload an Excel workbook", type=["xlsx"])
    if not uploaded:
        st.info("Upload a workbook to generate a structural map, issue list, and AI-ready analysis layer.")
        return

    with st.spinner("Parsing workbook..."):
        workbook = parse_workbook(uploaded, uploaded.name)
        graph = build_dependency_graph(workbook)
        issues, blocks = run_issue_pipeline(workbook)

    counts = _priority_counts(issues)
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Sheets", workbook.total_sheets)
    c2.metric("Formulas", workbook.total_formulas)
    c3.metric("High", counts["High"])
    c4.metric("Medium", counts["Medium"])
    c5.metric("Low", counts["Low"])

    st.subheader("Executive Summary")
    st.write(
        f"This workbook looks most like a **{workbook.workbook_type_hint}** model. "
        f"The current starter engine found **{len(issues)}** issues across **{len(blocks)}** formula blocks. "
        "V2 is structured to improve hidden-formula detection and support AI-guided navigation next."
    )

    tab_summary, tab_issues, tab_trace, tab_chat = st.tabs(
        ["Model Summary", "Detected Issues", "Trace a Cell", "AI Navigator"]
    )

    with tab_summary:
        rows = [
            {
                "Sheet": sheet.name,
                "Rows": sheet.max_row,
                "Cols": sheet.max_col,
                "Formulas": sheet.formula_count,
                "Text": sheet.text_count,
                "Numeric": sheet.numeric_count,
            }
            for sheet in workbook.sheets
        ]
        st.dataframe(rows, use_container_width=True)
        st.write(
            {
                "named_ranges": len(workbook.named_ranges),
                "dependency_nodes": graph.number_of_nodes(),
                "dependency_edges": graph.number_of_edges(),
                "formula_blocks": len(blocks),
            }
        )

    with tab_issues:
        selected_priority = st.selectbox("Priority filter", ["All", "High", "Medium", "Low"], index=0)
        issue_rows = [
            {
                "Priority": issue.priority,
                "Title": issue.title,
                "Sheet": issue.sheet,
                "Cells": ", ".join(issue.cells),
                "Detector": issue.detector,
                "Detail": issue.detail,
            }
            for issue in issues
            if selected_priority == "All" or issue.priority == selected_priority
        ]
        st.dataframe(issue_rows, use_container_width=True)

    with tab_trace:
        default_target = next(iter(workbook.formulas.keys()), "")
        target = st.text_input("Cell to inspect", value=default_target, key="trace_target")
        if target:
            details = get_cell_details(workbook, target)
            if not details["found"]:
                st.warning("That cell was not found in the workbook map.")
            else:
                st.write(
                    {
                        "target": details["target"],
                        "value": details["value"],
                        "formula": details["formula"],
                        "type": details["data_type"],
                    }
                )

                precedents = trace_precedents(graph, target, depth=3)
                dependents = trace_dependents(graph, target, depth=2)
                related_issues = find_direct_issue_hits(issues, target)

                col1, col2 = st.columns(2)
                with col1:
                    st.subheader("Precedents")
                    st.dataframe([{"Cell": item} for item in precedents] or [{"Cell": "None found"}], use_container_width=True)
                with col2:
                    st.subheader("Dependents")
                    st.dataframe([{"Cell": item} for item in dependents] or [{"Cell": "None found"}], use_container_width=True)

                st.subheader("Related Issues")
                st.dataframe(
                    related_issues
                    or [{"priority": "-", "title": "No direct issues on this cell", "cells": [target], "detail": "", "detector": "", "evidence": {}}],
                    use_container_width=True,
                )

    with tab_chat:
        if is_openai_enabled(api_key):
            st.success("OpenAI chat is enabled for grounded workbook Q&A.")
        else:
            st.info("Add an OpenAI API key in the sidebar to enable real AI chat. Fallback local Q&A is active for now.")

        sheet_options = [sheet.name for sheet in workbook.sheets]
        selected_sheet = st.selectbox("Navigator focus sheet", ["None"] + sheet_options, index=0)
        suggestion_col1, suggestion_col2, suggestion_col3 = st.columns(3)
        quick_prompt = None
        with suggestion_col1:
            if st.button("Top Issues"):
                quick_prompt = "Show the high priority issues."
        with suggestion_col2:
            if st.button("Largest Sheets"):
                quick_prompt = "Which sheets have the most formulas?"
        with suggestion_col3:
            if st.button("Workbook Summary"):
                quick_prompt = "Give me a workbook summary."

        if selected_sheet != "None":
            action_col1, action_col2, action_col3 = st.columns(3)
            with action_col1:
                if st.button("Explain Focus Sheet"):
                    quick_prompt = f"Explain issues in {selected_sheet}"
            with action_col2:
                if st.button("List Focus Issues"):
                    quick_prompt = f"Show issues in {selected_sheet}"
            with action_col3:
                if st.button("Summarize Focus Sheet"):
                    quick_prompt = f"Summarize {selected_sheet}"

        if selected_sheet != "None":
            st.caption(
                f"Try prompts like: `show issues in {selected_sheet}`, "
                f"`summarize {selected_sheet}`, or `trace {selected_sheet}!B10`."
            )

        default_question = quick_prompt or st.session_state.get("navigator_question", "")
        question = st.text_input(
            "Ask about the workbook",
            value=default_question,
            key="navigator_question",
            placeholder="Show issues in Debt Schedule or trace Outputs!F42",
        )
        ask_clicked = st.button("Ask")

        st.subheader("Sheet Navigator")
        st.dataframe(summarize_sheets(workbook, limit=12), use_container_width=True)

        if ask_clicked and question:
            if is_openai_enabled(api_key):
                st.write(
                    answer_with_openai(
                        question,
                        workbook,
                        graph,
                        issues,
                        api_key=api_key,
                        focus_sheet=selected_sheet,
                        model=model_choice,
                    )
                )
            else:
                st.write(answer_question(question, workbook, graph, issues))


if __name__ == "__main__":
    main()
