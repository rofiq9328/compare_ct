"""Shared dynamic row/column geometry for the COMPARE LINE Excel report.

The report uses two-row data blocks. Each original chart remains inside the
template's two-column graph area and follows the vertical extent of its section.
"""

from __future__ import annotations

from typing import Dict


MACHINE_FIRST_ROW = 17
CHART_COLUMN_SPAN = 2
# Chart objects align with the first data row, not the section title/header.
SUMMARY_CHART_ANCHOR_ROW = 4
SUMMARY_CHART_START_ROW = 4
SUMMARY_CHART_END_ROW = 13
MACHINE_CHART_ANCHOR_ROW = MACHINE_FIRST_ROW
MACHINE_CHART_START_ROW = MACHINE_FIRST_ROW


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
    """Return the dynamic table and chart anchors for one comparison workbook."""
    machine_count = _non_negative_int("machine_count", machine_count)
    process_count = _non_negative_int("process_count", process_count)
    line_count = _non_negative_int("line_count", line_count)
    if line_count < 2:
        raise ValueError("line_count minimal 2 agar perbandingan antar-line bermakna.")

    allocated_process_count = max(1, process_count)
    if chart_process_count is None:
        chart_process_count = max(1, process_count)
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

    # C onward are the dynamic line-value columns. Exactly two columns are
    # reserved for the template chart, followed by the notes column.
    chart_col = 3 + line_count
    chart_last_col = chart_col + CHART_COLUMN_SPAN - 1
    note_col = chart_last_col + 1

    return {
        "machine_first_row": MACHINE_FIRST_ROW,
        "machine_total_row": machine_total_row,
        "summary_chart_anchor_row": SUMMARY_CHART_ANCHOR_ROW,
        "summary_chart_start_row": SUMMARY_CHART_START_ROW,
        "summary_chart_end_row": SUMMARY_CHART_END_ROW,
        "machine_chart_anchor_row": MACHINE_CHART_ANCHOR_ROW,
        "machine_chart_start_row": MACHINE_CHART_START_ROW,
        # The total-machine block is two rows high; stop at its last row.
        "machine_chart_end_row": machine_total_row + 1,
        "section3_row": section3_row,
        "section3_header_row": section3_header_row,
        "process_start_row": process_start_row,
        "process_end_row": process_end_row,
        "process_chart_anchor_row": process_start_row,
        "process_chart_start_row": process_start_row,
        "process_chart_end_row": process_chart_end_row,
        "allocated_process_count": allocated_process_count,
        "chart_process_count": chart_process_count,
        "chart_col": chart_col,
        "chart_last_col": chart_last_col,
        "chart_column_span": CHART_COLUMN_SPAN,
        "note_col": note_col,
        "last_column": note_col,
    }
