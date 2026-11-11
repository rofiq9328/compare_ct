"""Update the three original charts in the COMPARE LINE template.

Run from the COMPARE_CT project root after Stage 6:
    py -m excel.chart_generator

Optional:
    py -m excel.chart_generator "output\\COMPARE_LINE.xlsx" \\
        --comparison "logs\\comparison_data.json"

Requires Windows, Microsoft Excel desktop, and pywin32.
"""
from __future__ import annotations

import argparse
import json
import math
import sys
import traceback
from pathlib import Path
from typing import Any, Dict, List, Mapping, Sequence

from excel.layout import CHART_COLUMN_SPAN, calculate_excel_layout

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_WORKBOOK = PROJECT_ROOT / "output" / "COMPARE_LINE.xlsx"
DEFAULT_COMPARISON = PROJECT_ROOT / "logs" / "comparison_data.json"

XL_COLUMN_CLUSTERED = 51  # retained for compatibility with earlier Stage 7 files
XL_BAR_CLUSTERED = 57
XL_VALUE = 2
XL_PRIMARY = 1


class ChartGeneratorError(ValueError):
    pass


def _read_json(path: Path) -> Dict[str, Any]:
    if not path.is_file():
        raise FileNotFoundError(f"comparison_data.json tidak ditemukan: {path}")
    try:
        value = json.loads(path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        raise ChartGeneratorError(
            f"comparison_data.json bukan JSON valid: {exc}"
        ) from exc
    if not isinstance(value, dict):
        raise ChartGeneratorError(
            "comparison_data.json harus berupa object JSON."
        )
    if (
        value.get("schema_version") != 1
        or value.get("comparison_type") != "dynamic_multi_line"
    ):
        raise ChartGeneratorError(
            "Format comparison_data.json tidak didukung."
        )
    return value


def _safe_number(value: Any) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return 0.0
    if not math.isfinite(float(value)):
        return 0.0
    return float(value)


def _build_chart_data(data: Mapping[str, Any]) -> Dict[str, Any]:
    line_columns = data.get("line_columns") or []
    labels = [
        str(item.get("line_label") or "").strip()
        for item in line_columns
    ]
    if len(labels) < 2 or any(not label for label in labels):
        raise ChartGeneratorError(
            "line_columns harus berisi minimal dua line dengan label valid."
        )
    if len(set(labels)) != len(labels):
        raise ChartGeneratorError("line_label tidak boleh duplikat.")

    summary_rows = {
        str(row.get("key") or ""): row
        for row in data.get("summary_rows") or []
    }
    # Keep the original five summary categories from the workbook template.
    # Units are displayed beside values in each LINE column, not on chart axes.
    summary_order = (
        ("operator_allocations", "Jumlah Operator"),
        ("takt_time", "Takt Time"),
        ("target_per_hour", "Target Output / Jam"),
        ("potential_per_hour", "Potensial Output / Jam"),
        ("actual_per_hour", "Aktual Output / Jam"),
    )

    summary_matrix: List[List[Any]] = [["VARIABEL"] + labels]
    for key, fallback_label in summary_order:
        if key not in summary_rows:
            raise ChartGeneratorError(
                f"Summary tidak memiliki key {key}."
            )
        row = summary_rows[key]
        values = row.get("values_by_line") or {}
        label = str(row.get("label") or fallback_label)
        summary_matrix.append(
            [label] + [_safe_number(values.get(line)) for line in labels]
        )

    machine_rows = data.get("machine_rows") or []
    machine_matrix: List[List[Any]] = [["JENIS MESIN"] + labels]
    for row in machine_rows:
        machine = str(row.get("machine") or "").strip()
        if not machine:
            continue
        values = row.get("values_by_line") or {}
        machine_matrix.append(
            [machine] + [_safe_number(values.get(line)) for line in labels]
        )
    if len(machine_matrix) == 1:
        raise ChartGeneratorError(
            "Tidak ada data jenis mesin untuk grafik."
        )

    total_row = data.get("machine_total_row") or {}
    total_values = total_row.get("values_by_line") or {}
    machine_matrix.append(
        ["TOTAL MESIN"] + [
            _safe_number(total_values.get(line)) for line in labels
        ]
    )

    process_rows = []
    for row in data.get("process_rows") or []:
        process = str(row.get("process") or "").strip()
        if not process:
            continue
        values = row.get("operator_allocations_by_line") or {}
        per_line = [_safe_number(values.get(line)) for line in labels]
        # Preserve the report table's order and keep zero-allocation rows so
        # every visible process variable is represented in the chart.
        process_rows.append((process, per_line))

    process_matrix: List[List[Any]] = [["PROSES"] + labels]
    for process, per_line in process_rows:
        process_matrix.append([process] + per_line)
    if len(process_matrix) == 1:
        process_matrix.append(
            ["Tidak ada alokasi"] + [0 for _ in labels]
        )

    return {
        "labels": labels,
        "summary_matrix": summary_matrix,
        "machine_matrix": machine_matrix,
        "process_matrix": process_matrix,
        "machine_count": len(machine_rows),
        "process_count": len(process_rows),
        "process_chart_count": len(process_matrix) - 1,
    }


def _get_chart_objects(ws: Any) -> List[Any]:
    collection = ws.ChartObjects()
    count = int(collection.Count)
    if count != 3:
        raise ChartGeneratorError(
            "Template harus memiliki tepat 3 grafik bawaan "
            f"(Summary, Mesin, Operator); ditemukan {count}. "
            "Jalankan ulang Tahap 6 menggunakan report_generator.py "
            "versi revisi agar grafik template tidak terhapus."
        )
    objects = [collection.Item(index) for index in range(1, count + 1)]
    # Template order by vertical position: Summary, Mesin, Operator.
    return sorted(objects, key=lambda item: float(item.Top))


def _excel_column_letter(column: int) -> str:
    """Convert a 1-based Excel column number to its A1 column label."""
    if isinstance(column, bool) or not isinstance(column, int) or column < 1:
        raise ValueError(f"Nomor kolom Excel tidak valid: {column!r}.")
    result = ""
    while column:
        column, remainder = divmod(column - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _qualified_cell_reference(ws: Any, row: int, column: int) -> str:
    """Return an absolute A1 reference to a visible report cell."""
    if isinstance(row, bool) or not isinstance(row, int) or row < 1:
        raise ValueError(f"Nomor baris Excel tidak valid: {row!r}.")
    if isinstance(column, bool) or not isinstance(column, int) or column < 1:
        raise ValueError(f"Nomor kolom Excel tidak valid: {column!r}.")
    sheet_name = str(ws.Name).replace("'", "''")
    return "'" + sheet_name + "'!$" + _excel_column_letter(column) + "$" + str(row)


def _build_series_formula(
    ws: Any,
    header_row: int,
    value_column: int,
    category_column: int,
    data_rows: Sequence[int],
    series_order: int,
) -> str:
    """Build a live SERIES formula using one visible table cell per variable."""
    ordered_rows = list(data_rows)
    if not ordered_rows:
        raise ChartGeneratorError("Baris sumber grafik dari tabel kosong.")
    if series_order < 1:
        raise ValueError("series_order minimal 1.")

    name_ref = _qualified_cell_reference(ws, header_row, value_column)
    category_refs = ",".join(
        _qualified_cell_reference(ws, row, category_column)
        for row in ordered_rows
    )
    value_refs = ",".join(
        _qualified_cell_reference(ws, row, value_column)
        for row in ordered_rows
    )

    return (
        f"=SERIES({name_ref},"
        f"({category_refs}),"
        f"({value_refs}),"
        f"{series_order})"
    )


def _build_table_row_plan(
    layout: Mapping[str, int],
    machine_count: int,
    process_count: int,
) -> Dict[str, tuple[int, ...]]:
    """Return the actual visible report rows used by each chart."""
    summary_rows = (4, 6, 8, 10, 12)

    machine_rows = tuple(
        layout["machine_first_row"] + (2 * index)
        for index in range(machine_count)
    ) + (layout["machine_total_row"],)

    # The report reserves one placeholder block when there are no processes.
    process_rows = tuple(
        layout["process_start_row"] + (2 * index)
        for index in range(max(1, process_count))
    )

    return {
        "summary": summary_rows,
        "machines": machine_rows,
        "processes": process_rows,
    }


def _update_chart_series_from_table(
    chart_object: Any,
    ws: Any,
    labels: Sequence[str],
    header_row: int,
    line_start_col: int,
    category_col: int,
    data_rows: Sequence[int],
) -> None:
    """Bind every chart series directly to the visible table cells."""
    if not data_rows:
        raise ChartGeneratorError("Tidak ada baris tabel untuk sumber grafik.")

    chart = chart_object.Chart
    chart.PlotVisibleOnly = True
    chart.PlotBy = 2  # xlColumns: one series for each compared LINE

    series_collection = chart.SeriesCollection()

    # Retain existing template series (and their formatting), add/remove only
    # the number needed for the current line count.
    while int(series_collection.Count) > len(labels):
        series_collection.Item(int(series_collection.Count)).Delete()

    while int(series_collection.Count) < len(labels):
        series_collection.NewSeries()

    extra_colors = [
        0x70AD47,  # green
        0xED7D31,  # orange
        0xA5A5A5,  # gray
        0x7030A0,  # purple
        0xC00000,  # dark red
        0x4472C4,  # blue
        0x00B0F0,  # cyan
        0xC55A11,  # brown-orange
        0x548235,  # dark green
        0xBF9000,  # gold
        0x8064A2,  # violet
        0x31869B,  # teal
    ]

    def extra_color(index: int) -> int:
        if index < len(extra_colors):
            return extra_colors[index]
        import colorsys
        hue = ((index - len(extra_colors)) * 0.61803398875 + 0.13) % 1.0
        red, green, blue = colorsys.hsv_to_rgb(hue, 0.68, 0.82)
        return int(red * 255) | (int(green * 255) << 8) | (int(blue * 255) << 16)

    for index, label in enumerate(labels):
        value_col = line_start_col + index
        header_value = ws.Cells(header_row, value_col).Value
        if header_value is None or not str(header_value).strip():
            raise ChartGeneratorError(
                f"Header LINE pada tabel kosong: baris {header_row}, kolom {value_col}."
            )

        series = series_collection.Item(index + 1)

        # Reference exactly one top-left cell from each merged two-row item.
        # This keeps category count equal to variable count and keeps every
        # point live-linked to the visible report table.
        series.Formula = _build_series_formula(
            ws,
            header_row=header_row,
            value_column=value_col,
            category_column=category_col,
            data_rows=data_rows,
            series_order=index + 1,
        )

        # Keep the first two series' exact template colors. Additional lines
        # receive distinct colors while retaining the chart's native style.
        if index >= 2:
            color = extra_color(index - 2)
            try:
                series.Format.Fill.ForeColor.RGB = color
                series.Format.Line.ForeColor.RGB = color
            except Exception:
                pass


def _chart_row_bounds(data_rows: Sequence[int], block_height: int = 2) -> tuple[int, int]:
    """Return exact worksheet rows occupied by the chart's table variables."""
    rows = list(data_rows)
    if not rows:
        raise ChartGeneratorError("Tidak ada baris tabel untuk menentukan ukuran grafik.")
    if block_height < 1:
        raise ValueError("block_height minimal 1.")
    if any(current - previous != block_height for previous, current in zip(rows, rows[1:])):
        raise ChartGeneratorError("Baris chart tidak mengikuti blok tabel dua baris.")
    return rows[0], rows[-1] + block_height - 1


def _position_chart(
    chart_object: Any,
    ws: Any,
    chart_col: int,
    first_row: int,
    last_row: int,
) -> float:
    """Fit chart exactly over the two chart columns and its variable table rows."""
    if first_row < 1 or last_row < first_row:
        raise ChartGeneratorError(
            f"Rentang baris grafik tidak valid: {first_row}:{last_row}."
        )

    target_range = ws.Range(
        ws.Cells(first_row, chart_col),
        ws.Cells(last_row, chart_col + CHART_COLUMN_SPAN - 1),
    )

    # xlMove keeps the chart attached to cells without automatic resize.
    # This run then explicitly sets all edges to the actual table range.
    chart_object.Placement = 2
    chart_object.Left = float(target_range.Left)
    chart_object.Top = float(target_range.Top)
    chart_object.Width = float(target_range.Width)
    chart_object.Height = float(target_range.Height)
    return float(chart_object.Height)


def generate_charts(
    workbook_path: str | Path = DEFAULT_WORKBOOK,
    comparison_path: str | Path = DEFAULT_COMPARISON,
) -> Dict[str, Any]:
    workbook_file = Path(workbook_path).expanduser().resolve()
    comparison_file = Path(comparison_path).expanduser().resolve()

    if not workbook_file.is_file():
        raise FileNotFoundError(
            f"Workbook Tahap 6 tidak ditemukan: {workbook_file}"
        )

    data = _read_json(comparison_file)
    chart_data = _build_chart_data(data)

    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 belum terpasang. Jalankan: py -m pip install pywin32"
        ) from exc

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    stage = "Inisialisasi Excel"

    try:
        stage = "Membuka Excel COM"
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        stage = "Membuka workbook Tahap 6"
        workbook = excel.Workbooks.Open(
            str(workbook_file), UpdateLinks=0, ReadOnly=False
        )

        stage = "Memeriksa worksheet"
        if workbook.Worksheets.Count != 1:
            raise ChartGeneratorError(
                "Workbook harus hanya memiliki satu sheet; "
                f"ditemukan {workbook.Worksheets.Count}."
            )
        ws = workbook.Worksheets(1)
        sheet_name = str(ws.Name)
        if sheet_name.strip().lower() != "compare line":
            raise ChartGeneratorError(
                f"Nama sheet harus 'compare line'; ditemukan {sheet_name!r}."
            )

        labels = chart_data["labels"]
        line_count = len(labels)
        machine_count = chart_data["machine_count"]
        process_count = chart_data["process_count"]

        # Use the same dynamic geometry as Stage 6.
        chart_process_count = chart_data["process_chart_count"]
        layout = calculate_excel_layout(
            machine_count, process_count, line_count,
            chart_process_count=chart_process_count,
        )
        machine_total_row = layout["machine_total_row"]
        section3_row = layout["section3_row"]
        section3_header_row = layout["section3_header_row"]
        process_start_row = layout["process_start_row"]
        process_end_row = layout["process_end_row"]
        chart_col = layout["chart_col"]
        note_col = layout["note_col"]

        # Each SERIES formula references exactly one visible table cell per
        # variable, skipping the blank second row of each merged two-row block.
        row_plan = _build_table_row_plan(
            layout, machine_count, process_count
        )

        stage = "Membaca tiga grafik bawaan template"
        summary_chart, machine_chart, process_chart = _get_chart_objects(ws)

        stage = "Memperbarui grafik Summary dari tabel"
        _update_chart_series_from_table(
            summary_chart, ws, labels,
            header_row=3,
            line_start_col=3,
            category_col=2,
            data_rows=row_plan["summary"],
        )
        summary_chart_rows = _chart_row_bounds(row_plan["summary"])
        summary_height = _position_chart(
            summary_chart, ws, chart_col,
            first_row=summary_chart_rows[0],
            last_row=summary_chart_rows[1],
        )

        stage = "Memperbarui grafik Penggunaan Mesin dari tabel"
        _update_chart_series_from_table(
            machine_chart, ws, labels,
            header_row=16,
            line_start_col=3,
            category_col=2,
            data_rows=row_plan["machines"],
        )
        machine_chart_rows = _chart_row_bounds(row_plan["machines"])
        machine_height = _position_chart(
            machine_chart, ws, chart_col,
            first_row=machine_chart_rows[0],
            last_row=machine_chart_rows[1],
        )

        stage = "Memperbarui grafik Penggunaan Operator dari tabel"
        _update_chart_series_from_table(
            process_chart, ws, labels,
            header_row=section3_header_row,
            line_start_col=3,
            category_col=2,
            data_rows=row_plan["processes"],
        )
        operator_chart_rows = _chart_row_bounds(row_plan["processes"])
        operator_height = _position_chart(
            process_chart, ws, chart_col,
            first_row=operator_chart_rows[0],
            last_row=operator_chart_rows[1],
        )

        stage = "Menyimpan workbook"
        ws.Activate()
        ws.Range("B1").Select()
        workbook.Save()

        stage = "Menutup workbook"
        workbook.Close(SaveChanges=False)
        workbook = None

        return {
            "output": str(workbook_file),
            "sheet": sheet_name,
            "line_count": line_count,
            "machine_count": machine_count,
            "process_count": process_count,
            "process_chart_count": chart_process_count,
            "chart_count": 3,
            "chart_mode": "TABEL_LAPORAN_LANGSUNG",
            "chart_column_span": CHART_COLUMN_SPAN,
            "chart_data_layout": "ONE_VARIABLE_ROW_PER_CATEGORY",
            "chart_anchor_rows": {
                "summary": summary_chart_rows[0],
                "machines": machine_chart_rows[0],
                "operators": operator_chart_rows[0],
            },
            "chart_end_rows": {
                "summary": summary_chart_rows[1],
                "machines": machine_chart_rows[1],
                "operators": operator_chart_rows[1],
            },
            "chart_heights": {
                "summary": summary_height,
                "machines": machine_height,
                "operators": operator_height,
            },
            "machine_category_count": len(row_plan["machines"]),
            "operator_category_count": len(row_plan["processes"]),
            "status": "PASS",
        }

    except Exception as exc:
        raise RuntimeError(
            f"Gagal pada tahap: {stage}\nPenyebab: {exc}"
        ) from exc

    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.DisplayAlerts = False
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def format_console_report(result: Mapping[str, Any]) -> str:
    return "\n".join([
        "=" * 72,
        "COMPARE_CT — TAHAP 7: UPDATE GRAFIK TEMPLATE",
        "=" * 72,
        f"OUTPUT              : {result['output']}",
        f"SHEET               : {result['sheet']}",
        f"JUMLAH LINE         : {result['line_count']}",
        f"JENIS MESIN         : {result['machine_count']}",
        f"BARIS PROSES        : {result['process_count']}",
        f"PROSES DI GRAFIK    : {result['process_chart_count']} (SELURUH PROSES)",
        f"JUMLAH GRAFIK       : {result['chart_count']}",
        f"MODE GRAFIK         : {result['chart_mode']}",
        f"LEBAR GRAFIK        : {result['chart_column_span']} kolom",
        f"KATEGORI MESIN      : {result['machine_category_count']}",
        f"KATEGORI OPERATOR   : {result['operator_category_count']}",
        f"SUMMARY             : BARIS {result['chart_anchor_rows']['summary']}-"
        f"{result['chart_end_rows']['summary']} / {result['chart_heights']['summary']:.1f} pt",
        f"MESIN               : BARIS {result['chart_anchor_rows']['machines']}-"
        f"{result['chart_end_rows']['machines']} / {result['chart_heights']['machines']:.1f} pt",
        f"OPERATOR            : BARIS {result['chart_anchor_rows']['operators']}-"
        f"{result['chart_end_rows']['operators']} / {result['chart_heights']['operators']:.1f} pt",
        f"STATUS              : {result['status']}",
        "=" * 72,
    ])


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Perbarui data tiga grafik bawaan template COMPARE LINE "
            "tanpa mengganti desainnya."
        )
    )
    parser.add_argument(
        "workbook", nargs="?", default=str(DEFAULT_WORKBOOK)
    )
    parser.add_argument(
        "--comparison", default=str(DEFAULT_COMPARISON)
    )
    args = parser.parse_args(argv)

    try:
        result = generate_charts(args.workbook, args.comparison)
    except Exception as exc:
        print("\n" + "=" * 72)
        print("COMPARE_CT — TAHAP 7: UPDATE GRAFIK TEMPLATE")
        print("=" * 72)
        print("STATUS              : ERROR")
        print(f"PENYEBAB            : {exc}")
        print("\nDETAIL TRACEBACK:")
        traceback.print_exc()
        print("=" * 72)
        return 1

    print(format_console_report(result))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())