"""Shared dynamic row/column geometry for the COMPARE LINE Excel report.

The template stores machine and process items as two-row blocks. Both the
report generator and chart generator use this module so their anchors cannot
drift apart when the number of lines or rows changes.
"""

from __future__ import annotations

from typing import Dict


MACHINE_FIRST_ROW = 17
PROCESS_CHART_TOP_N = 15


def _non_negative_int(name: str, value: int) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise ValueError(f"{name} harus berupa bilangan bulat >= 0; diterima: {value!r}.")
    return value


def calculate_excel_layout(
    machine_count: int,
    process_count: int,
    line_count: int,
    chart_process_count: int | None = None,
) -> Dict[str, int]:
    """Return all dynamic report and chart anchors for a given comparison size."""
    machine_count = _non_negative_int("machine_count", machine_count)
    process_count = _non_negative_int("process_count", process_count)
    line_count = _non_negative_int("line_count", line_count)
    if line_count < 2:
        raise ValueError("line_count minimal 2 agar perbandingan antar-line bermakna.")

    allocated_process_count = max(1, process_count)
    if chart_process_count is None:
        chart_process_count = max(1, min(process_count, PROCESS_CHART_TOP_N))
    else:
        chart_process_count = _non_negative_int("chart_process_count", chart_process_count)
        chart_process_count = max(1, chart_process_count)
    if chart_process_count > allocated_process_count:
        raise ValueError(
            "chart_process_count tidak boleh melebihi allocated_process_count."
        )

    machine_total_row = MACHINE_FIRST_ROW + (2 * machine_count)
    section3_row = machine_total_row + 3
    section3_header_row = section3_row + 1
    process_start_row = section3_header_row + 1
    process_end_row = process_start_row + (2 * allocated_process_count) - 1
    process_chart_end_row = process_start_row + (2 * chart_process_count) - 1

    # C onward are the dynamic line-value columns. The two-column chart area
    # and one note column are always kept to their right.
    chart_col = 3 + line_count
    note_col = chart_col + 2
    helper_col = note_col + 3

    return {
        "machine_first_row": MACHINE_FIRST_ROW,
        "machine_total_row": machine_total_row,
        "machine_chart_end_row": machine_total_row + 2,
        "section3_row": section3_row,
        "section3_header_row": section3_header_row,
        "process_start_row": process_start_row,
        "process_end_row": process_end_row,
        "process_chart_end_row": process_chart_end_row,
        "allocated_process_count": allocated_process_count,
        "chart_process_count": chart_process_count,
        "chart_col": chart_col,
        "note_col": note_col,
        "helper_col": helper_col,
        "last_column": note_col,
    }
