import pytest

from randomoffice.engine.dependency_graph import (
    DependencyGraph,
    FormulaCycleError,
    RangeExpansionLimitError,
    extract_references,
)


def test_extract_references_expands_ranges_and_ignores_strings() -> None:
    # A1:B2 expands to A1, A2, B1, B2; "A99" literal is ignored; Budget!$C$4 is normalized
    refs = extract_references('=SUM(A1:B2)+"A99"+Budget!$C$4', current_sheet="Sheet1")
    assert refs == {
        "A1",
        "A2",
        "B1",
        "B2",
        "budget!C4",
    }


def test_extract_references_distinguishes_functions_from_cells() -> None:
    # LOG10 is a function call followed by '(', not a cell reference
    refs = extract_references("=LOG10(A1)+SIN(B1)")
    assert refs == {"A1", "B1"}
    assert "LOG10" not in refs
    assert "SIN" not in refs


def test_extract_references_ignores_external_workbook_references() -> None:
    # External workbook reference should not be quietly converted to a local cell
    refs = extract_references("=[OtherBook.xlsx]Data!A1 + C1")
    assert refs == {"C1"}


def test_extract_references_ignores_external_workbook_ranges() -> None:
    refs = extract_references("=SUM([OtherBook.xlsx]Data!A1:B2)+C1")

    assert refs == {"C1"}


def test_extract_references_range_size_limit() -> None:
    # Extremely large range should trigger protection
    with pytest.raises(RangeExpansionLimitError, match="exceeds limit"):
        extract_references("=SUM(A1:Z5000)", max_range_cells=100)


def test_interior_range_cycle_detected() -> None:
    # A2 inside SUM(A1:A3) is a circular dependency
    with pytest.raises(FormulaCycleError, match="A2"):
        DependencyGraph.from_formulas({"A2": "=SUM(A1:A3)"})


def test_topological_order_puts_formula_inputs_first() -> None:
    graph = DependencyGraph.from_formulas(
        {
            "B1": "=A1*2",
            "C1": "=B1+5",
            "D1": "=SUM(A1:C1)",
        }
    )

    order = graph.topological_order()

    assert order.index("A1") < order.index("B1")
    assert order.index("B1") < order.index("C1")
    assert order.index("C1") < order.index("D1")


def test_topological_order_reports_indirect_cycles() -> None:
    graph = DependencyGraph.from_formulas(
        {"A1": "=B1+1", "B1": "=C1+1", "C1": "=A1+1"}
    )

    with pytest.raises(FormulaCycleError, match="Circular formula reference"):
        graph.topological_order()


def test_indirect_range_cycle_detected() -> None:
    # A1 depends on B1, while B1 depends on A1:A2 (which contains A1)
    graph = DependencyGraph.from_formulas(
        {
            "A1": "=B1+10",
            "B1": "=SUM(A1:A2)",
        }
    )
    with pytest.raises(FormulaCycleError):
        graph.topological_order()


def test_cross_sheet_normalization() -> None:
    graph = DependencyGraph.from_formulas(
        {"B1": "='Sheet1'!a1 + 'OTHER'!C2"},
        current_sheet="sheet1",
    )
    # Sheet1!a1 normalized to A1 (current sheet), other!C2 normalized with sheet prefix
    assert graph.dependencies_for("B1", current_sheet="Sheet1") == {"A1", "other!C2"}
    assert graph.dependencies_for("b1", current_sheet="Sheet1") == {"A1", "other!C2"}


def test_dependencies_for_unknown_cell_is_empty() -> None:
    graph = DependencyGraph.from_formulas({"B2": "=A1"})

    assert graph.dependencies_for("C3") == frozenset()
