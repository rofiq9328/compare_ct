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

from excel.layout import calculate_excel_layout

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


def _col_letter(number: int) -> str:
    result = ""
    while number:
        number, rem = divmod(number - 1, 26)
        result = chr(65 + rem) + result
    return result


def _write_matrix(
    ws: Any,
    start_row: int,
    start_col: int,
    matrix: Sequence[Sequence[Any]],
) -> tuple[int, int]:
    rows = len(matrix)
    cols = max((len(row) for row in matrix), default=0)
    if not rows or not cols:
        raise ChartGeneratorError("Data sumber grafik kosong.")

    normalized = [
        list(row) + [None] * (cols - len(row))
        for row in matrix
    ]
    ws.Range(
        ws.Cells(start_row, start_col),
        ws.Cells(start_row + rows - 1, start_col + cols - 1),
    ).Value = tuple(tuple(row) for row in normalized)
    return start_row + 1, start_row + rows - 1


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
    # Only compare output-rate KPIs in one chart. Operator count and takt
    # time use different units and must not share the same value axis.
    summary_order = (
        ("target_per_hour", "Target"),
        ("potential_per_hour", "Potensial"),
        ("actual_per_hour", "Aktual"),
    )

    summary_matrix: List[List[Any]] = [["OUTPUT PER JAM (PCS/JAM)"] + labels]
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
    machine_matrix: List[List[Any]] = [["JENIS MESIN (UNIT)"] + labels]
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

    # TOTAL MESIN remains a table KPI, not a machine category. Excluding it
    # keeps the chart scale comparable between individual machine types.

    process_ranked = []
    for row in data.get("process_rows") or []:
        process = str(row.get("process") or "").strip()
        if not process:
            continue
        values = row.get("operator_allocations_by_line") or {}
        per_line = [_safe_number(values.get(line)) for line in labels]
        total = sum(per_line)
        if total > 0:
            process_ranked.append((total, process, per_line))

    process_ranked.sort(
        key=lambda item: (-item[0], item[1].casefold())
    )
    top_processes = process_ranked[:15]
    process_matrix: List[List[Any]] = [["PROSES (TOP 15) - OPERATOR (OPT)"] + labels]
    for _, process, per_line in top_processes:
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
        "process_count": len(data.get("process_rows") or []),
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


def _update_chart_series(
    chart_object: Any,
    ws: Any,
    labels: Sequence[str],
    helper_col: int,
    header_row: int,
    first_data_row: int,
    last_data_row: int,
) -> None:
    chart = chart_object.Chart
    chart.PlotVisibleOnly = False  # source table is stored in hidden helper columns

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
        # Keep generating distinct colors for comparisons with many lines.
        import colorsys
        hue = ((index - len(extra_colors)) * 0.61803398875 + 0.13) % 1.0
        red, green, blue = colorsys.hsv_to_rgb(hue, 0.68, 0.82)
        return int(red * 255) | (int(green * 255) << 8) | (int(blue * 255) << 16)

    for index, label in enumerate(labels):
        series = series_collection.Item(index + 1)
        value_col = helper_col + index + 1
        category_range = ws.Range(
            ws.Cells(first_data_row, helper_col),
            ws.Cells(last_data_row, helper_col),
        )
        value_range = ws.Range(
            ws.Cells(first_data_row, value_col),
            ws.Cells(last_data_row, value_col),
        )
        series.Name = label
        series.XValues = category_range
        series.Values = value_range

        # Keep the first two series' exact template colors. Additional lines
        # receive distinct colors while retaining the chart's native style.
        if index >= 2:
            color = extra_color(index - 2)
            try:
                series.Format.Fill.ForeColor.RGB = color
                series.Format.Line.ForeColor.RGB = color
            except Exception:
                pass


def _position_chart(
    chart_object: Any,
    ws: Any,
    chart_col: int,
    first_row: int,
    last_row: int,
) -> None:
    left_cell = ws.Cells(first_row, chart_col)
    top = left_cell.Top
    left = left_cell.Left
    width = ws.Range(
        ws.Cells(first_row, chart_col),
        ws.Cells(first_row, chart_col + 1),
    ).Width
    height = ws.Range(
        ws.Cells(first_row, chart_col),
        ws.Cells(last_row, chart_col + 1),
    ).Height

    chart_object.Left = left
    chart_object.Top = top
    chart_object.Width = width
    chart_object.Height = height


def _set_chart_titles(chart_object: Any, title: str, value_axis_title: str) -> None:
    """Set meaningful chart and value-axis titles without replacing template styling."""
    chart = chart_object.Chart
    chart.HasTitle = True
    chart.ChartTitle.Text = title
    try:
        value_axis = chart.Axes(XL_VALUE, XL_PRIMARY)
        value_axis.HasTitle = True
        value_axis.AxisTitle.Text = value_axis_title
        value_axis.TickLabels.NumberFormat = "0.##"
    except Exception:
        # Some template chart types may not expose a primary value axis.
        pass


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

        # Compact helper tables are placed out of sight to the right of the
        # report. They feed the existing template charts without rebuilding them.
        helper_col = layout["helper_col"]
        summary_header, summary_end = _write_matrix(
            ws, 1, helper_col, chart_data["summary_matrix"]
        )
        machine_start_row = 10
        machine_header, machine_end = _write_matrix(
            ws, machine_start_row, helper_col,
            chart_data["machine_matrix"],
        )
        process_start_helper = machine_start_row + len(
            chart_data["machine_matrix"]
        ) + 3
        process_header, process_end = _write_matrix(
            ws, process_start_helper, helper_col,
            chart_data["process_matrix"],
        )

        helper_last_col = helper_col + line_count
        ws.Columns(
            f"{_col_letter(helper_col)}:{_col_letter(helper_last_col)}"
        ).Hidden = True

        stage = "Membaca tiga grafik bawaan template"
        summary_chart, machine_chart, process_chart = _get_chart_objects(ws)

        stage = "Memperbarui grafik Summary"
        _update_chart_series(
            summary_chart, ws, labels, helper_col,
            header_row=1,
            first_data_row=2,
            last_data_row=summary_end,
        )
        _position_chart(
            summary_chart, ws, chart_col,
            first_row=3,
            last_row=14,
        )
        _set_chart_titles(
            summary_chart,
            "PERBANDINGAN OUTPUT PER JAM",
            "Output (pcs/jam)",
        )

        stage = "Memperbarui grafik Penggunaan Mesin"
        _update_chart_series(
            machine_chart, ws, labels, helper_col,
            header_row=machine_start_row,
            first_data_row=machine_header,
            last_data_row=machine_end,
        )
        _position_chart(
            machine_chart, ws, chart_col,
            first_row=15,
            last_row=machine_total_row + 2,
        )
        _set_chart_titles(
            machine_chart,
            "PENGGUNAAN JENIS MESIN PER LINE",
            "Jumlah mesin (unit)",
        )

        stage = "Memperbarui grafik Penggunaan Operator"
        _update_chart_series(
            process_chart, ws, labels, helper_col,
            header_row=process_start_helper,
            first_data_row=process_header,
            last_data_row=process_end,
        )
        _position_chart(
            process_chart, ws, chart_col,
            first_row=section3_row,
            # The process table has two rows per item; the final occupied
            # row is start + (2 * count) - 1, not start + (2 * count).
            last_row=process_start_row + (2 * chart_process_count) - 1,
        )
        _set_chart_titles(
            process_chart,
            "TOP 15 PENGGUNAAN OPERATOR PER PROSES",
            "Jumlah operator (OPT)",
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
            "chart_mode": "TEMPLATE_ASLI_DIPERTAHANKAN",
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
        f"PROSES DI GRAFIK    : {result['process_chart_count']} (TOP 15)",
        f"JUMLAH GRAFIK       : {result['chart_count']}",
        f"MODE GRAFIK         : {result['chart_mode']}",
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