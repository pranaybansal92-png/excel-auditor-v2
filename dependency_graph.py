from __future__ import annotations

import re

import networkx as nx

from models import WorkbookModel

REF_RE = re.compile(
    r"(?:(?:'[^']+'|[A-Za-z0-9_]+)!)?\$?[A-Z]{1,3}\$?\d{1,7}"
)


def build_dependency_graph(workbook: WorkbookModel) -> nx.DiGraph:
    graph = nx.DiGraph()
    for key, cell in workbook.formulas.items():
        graph.add_node(key)
        for ref in extract_references(cell.formula or "", cell.sheet):
            graph.add_edge(ref, key)
    return graph


def extract_references(formula: str, current_sheet: str) -> list[str]:
    refs: list[str] = []
    for match in REF_RE.findall(formula):
        if "!" in match:
            sheet_part, cell_part = match.split("!", 1)
            sheet_name = sheet_part.strip("'")
        else:
            sheet_name = current_sheet
            cell_part = match
        refs.append(f"{sheet_name}!{cell_part.replace('$', '')}")
    return refs


def trace_precedents(graph: nx.DiGraph, target: str, depth: int = 2) -> list[str]:
    if target not in graph:
        return []
    seen: set[str] = set()
    frontier = {target}
    for _ in range(depth):
        next_frontier: set[str] = set()
        for node in frontier:
            for predecessor in graph.predecessors(node):
                if predecessor not in seen:
                    seen.add(predecessor)
                    next_frontier.add(predecessor)
        frontier = next_frontier
        if not frontier:
            break
    return sorted(seen)


def trace_dependents(graph: nx.DiGraph, target: str, depth: int = 2) -> list[str]:
    if target not in graph:
        return []
    seen: set[str] = set()
    frontier = {target}
    for _ in range(depth):
        next_frontier: set[str] = set()
        for node in frontier:
            for dependent in graph.successors(node):
                if dependent not in seen:
                    seen.add(dependent)
                    next_frontier.add(dependent)
        frontier = next_frontier
        if not frontier:
            break
    return sorted(seen)
