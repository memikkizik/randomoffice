"""Dependency analysis for spreadsheet formulas.

Extracts cell references, expands cell ranges, detects direct and indirect
circular dependencies, and provides deterministic topological ordering.
"""

from __future__ import annotations

from collections import defaultdict, deque
from dataclasses import dataclass, field
import re
from typing import Iterable


MAX_RANGE_EXPANSION_CELLS = 10_000

# Matches ranges like 'Sheet1'!A1:B2 or A1:B2
_RANGE_REFERENCE = re.compile(
    r"(?<![\[\w])(?:(?P<sheet>'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?"
    r"(?P<start>\$?[A-Z]{1,3}\$?[1-9]\d*):(?P<end>\$?[A-Z]{1,3}\$?[1-9]\d*)",
    re.IGNORECASE,
)

# Matches single cell like 'Sheet1'!A1 or A1, NOT followed by ( (function call) or : (range)
_SINGLE_CELL_REFERENCE = re.compile(
    r"(?<![\[\w])(?:(?P<sheet>'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?"
    r"(?P<cell>\$?[A-Z]{1,3}\$?[1-9]\d*)(?!\s*[\(:])",
    re.IGNORECASE,
)


class FormulaCycleError(ValueError):
    """Raised when formula dependencies cannot be ordered because of a cycle."""

    def __init__(self, nodes: Iterable[str]) -> None:
        self.nodes = tuple(sorted(set(nodes)))
        super().__init__(f"Circular formula reference: {', '.join(self.nodes)}")


class RangeExpansionLimitError(ValueError):
    """Raised when expanding a spreadsheet range would exceed the safety limit."""


def _col_to_num(col: str) -> int:
    num = 0
    for char in col.upper():
        num = num * 26 + (ord(char) - ord("A") + 1)
    return num


def _num_to_col(num: int) -> str:
    col = ""
    while num > 0:
        num, remainder = divmod(num - 1, 26)
        col = chr(ord("A") + remainder) + col
    return col


def _parse_cell_coord(cell: str) -> tuple[int, int]:
    clean = cell.replace("$", "").upper()
    match = re.fullmatch(r"([A-Z]{1,3})([1-9]\d*)", clean)
    if not match:
        raise ValueError(f"Invalid cell coordinates: {cell}")
    return _col_to_num(match.group(1)), int(match.group(2))


def _normalise_sheet(sheet: str | None) -> str:
    if not sheet:
        return ""
    if sheet.startswith("'") and sheet.endswith("'"):
        sheet = sheet[1:-1].replace("''", "'")
    return sheet.casefold()


def _normalise_cell(cell: str) -> str:
    return cell.replace("$", "").upper()


def _normalise_ref_key(cell: str, current_sheet: str | None = None) -> str:
    """Normalize a reference or cell key to canonical form."""
    if "!" in cell:
        sheet, raw_cell = cell.split("!", 1)
        norm_sheet = _normalise_sheet(sheet)
        norm_cell = _normalise_cell(raw_cell)
        if norm_sheet and norm_sheet != _normalise_sheet(current_sheet):
            return f"{norm_sheet}!{norm_cell}"
        return norm_cell
    return _normalise_cell(cell)


def _expand_range(
    start: str,
    end: str,
    prefix: str,
    max_cells: int = MAX_RANGE_EXPANSION_CELLS,
) -> set[str]:
    c1, r1 = _parse_cell_coord(start)
    c2, r2 = _parse_cell_coord(end)
    min_c, max_c = min(c1, c2), max(c1, c2)
    min_r, max_r = min(r1, r2), max(r1, r2)

    total_cells = (max_c - min_c + 1) * (max_r - min_r + 1)
    if total_cells > max_cells:
        raise RangeExpansionLimitError(
            f"Range {start}:{end} ({total_cells} cells) exceeds limit of {max_cells} cells"
        )

    cells = set()
    for col in range(min_c, max_c + 1):
        col_str = _num_to_col(col)
        for row in range(min_r, max_r + 1):
            cell_name = f"{col_str}{row}"
            cells.add(f"{prefix}{cell_name}" if prefix else cell_name)
    return cells


def extract_references(
    formula: str,
    *,
    current_sheet: str | None = None,
    max_range_cells: int = MAX_RANGE_EXPANSION_CELLS,
) -> set[str]:
    """Extract normalized cell references from one formula.

    - Expands ranges (e.g. A1:B2 -> A1, A2, B1, B2) with bounds check.
    - Excludes function call tokens (e.g. LOG10(A1) extracts only A1).
    - Ignores external workbook references (e.g. [Book.xlsx]Sheet!A1).
    - Ignores string literals.
    - Normalizes sheet names case-insensitively and cell coordinates to uppercase.
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return set()

    # 1. Strip quoted strings
    expression = re.sub(r'"(?:[^"]|"")*"', "", formula)

    # 2. Mask out external workbook references like [Book.xlsx]Sheet!A1:B2
    expression = re.sub(
        r"\[[^\]]+\](?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!"
        r"\$?[A-Z]{1,3}\$?[1-9]\d*(?::\$?[A-Z]{1,3}\$?[1-9]\d*)?",
        "",
        expression,
    )

    # 3. Strip function names followed by '(' (e.g. LOG10(, SUM(, SIN()
    # Replace with '(' so cell references inside are not confused with function names
    expression = re.sub(r"(?<![!\w])[A-Za-z_][\w.]*\s*\(", "(", expression)

    current = _normalise_sheet(current_sheet)
    references: set[str] = set()

    # 4. Match and expand range expressions first
    def handle_range(match: re.Match) -> str:
        sheet = _normalise_sheet(match.group("sheet"))
        prefix = f"{sheet}!" if sheet and sheet != current else ""
        expanded = _expand_range(match.group("start"), match.group("end"), prefix, max_cells=max_range_cells)
        references.update(expanded)
        return " " * len(match.group(0))

    expression_without_ranges = _RANGE_REFERENCE.sub(handle_range, expression)

    # 5. Match remaining single cell references
    for match in _SINGLE_CELL_REFERENCE.finditer(expression_without_ranges):
        sheet = _normalise_sheet(match.group("sheet"))
        cell = _normalise_cell(match.group("cell"))
        prefix = f"{sheet}!" if sheet and sheet != current else ""
        references.add(f"{prefix}{cell}")

    return references


@dataclass
class DependencyGraph:
    """Directed formula dependency graph.

    Edges point from a referenced cell to the formula cell that depends on it.
    Non-formula cells may be included as roots; this makes the resulting order
    useful to calculation engines.
    """

    dependencies: dict[str, set[str]] = field(default_factory=dict)

    @classmethod
    def from_formulas(
        cls,
        formulas: dict[str, str],
        *,
        current_sheet: str | None = None,
        max_range_cells: int = MAX_RANGE_EXPANSION_CELLS,
    ) -> "DependencyGraph":
        """Build a graph from ``cell -> formula`` mappings.

        Direct self-references (including a cell located inside its own referenced range)
        raise FormulaCycleError immediately.
        """
        normalized: dict[str, set[str]] = {}
        for cell, formula in formulas.items():
            target = _normalise_ref_key(cell, current_sheet=current_sheet)
            refs = extract_references(
                formula,
                current_sheet=current_sheet,
                max_range_cells=max_range_cells,
            )
            if target in refs:
                raise FormulaCycleError([target])
            normalized[target] = refs
        return cls(normalized)

    def topological_order(self) -> list[str]:
        """Return dependencies before dependants, or raise on a circular graph."""
        nodes = set(self.dependencies)
        for refs in self.dependencies.values():
            nodes.update(refs)

        outgoing: dict[str, set[str]] = defaultdict(set)
        indegree = {node: 0 for node in nodes}

        for target, refs in self.dependencies.items():
            for source in refs:
                if target not in outgoing[source]:
                    outgoing[source].add(target)
                    indegree[target] += 1

        ready = deque(sorted(node for node, degree in indegree.items() if degree == 0))
        ordered: list[str] = []
        while ready:
            node = ready.popleft()
            ordered.append(node)
            for target in sorted(outgoing[node]):
                indegree[target] -= 1
                if indegree[target] == 0:
                    ready.append(target)

        if len(ordered) != len(nodes):
            cycle_nodes = [node for node, degree in indegree.items() if degree > 0]
            raise FormulaCycleError(cycle_nodes)
        return ordered

    def dependencies_for(self, cell: str, *, current_sheet: str | None = None) -> frozenset[str]:
        """Return the direct references of one cell."""
        norm = _normalise_ref_key(cell, current_sheet=current_sheet)
        return frozenset(self.dependencies.get(norm, set()))
