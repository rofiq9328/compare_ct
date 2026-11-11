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

    process_ranked = []
    for row in data.get("process_rows") or []:
        process = str(row.get("process") or "").strip()
        if not process:
            continue
        values = row.get("operator_allocations_by_line") or {}
        per_line = [_safe_number(values.get(line)) for line in labels]
        total = sum(per_line)
        # Keep zero-allocation rows as well so the chart has one category for
        # every process shown in the report table.
        process_ranked.append((total, process, per_line))

    process_ranked.sort(
        key=lambda item: (-item[0], item[1].casefold())
    )
    # Include every process row, not just the first 15. The chart height grows
    # when the detail process data gets longer.
    process_matrix: List[List[Any]] = [["PROSES"] + labels]
    for _, process, per_line in process_ranked:
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
        "process_count": len(process_ranked),
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


def _union_cell_range(
    excel: Any,
    ws: Any,
    column: int,
    rows: Sequence[int],
) -> Any:
    """Build a live, non-contiguous Excel range from the visible report table."""
    ordered_rows = list(rows)
    if not ordered_rows:
        raise ChartGeneratorError("Baris sumber grafik dari tabel kosong.")

    source_range = ws.Cells(ordered_rows[0], column)
    for row_number in ordered_rows[1:]:
        source_range = excel.Union(
            source_range,
            ws.Cells(row_number, column),
        )
    return source_range


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


def _worksheet_cell_reference(ws: Any, row: int, column: int) -> str:
    """Return a live R1C1 reference to a cell on the report worksheet."""
    if isinstance(row, bool) or not isinstance(row, int) or row < 1:
        raise ValueError(f"Nomor baris Excel tidak valid: {row!r}.")
    if isinstance(column, bool) or not isinstance(column, int) or column < 1:
        raise ValueError(f"Nomor kolom Excel tidak valid: {column!r}.")

    sheet_name = str(ws.Name).replace("'", "''")
    return f"='{sheet_name}'!R{row}C{column}"


def _update_chart_series_from_table(
    excel: Any,
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

    category_range = _union_cell_range(excel, ws, category_col, data_rows)

    for index, label in enumerate(labels):
        value_col = line_start_col + index
        header_value = ws.Cells(header_row, value_col).Value
        if header_value is None or not str(header_value).strip():
            raise ChartGeneratorError(
                f"Header LINE pada tabel kosong: baris {header_row}, kolom {value_col}."
            )

        series = series_collection.Item(index + 1)
        value_range = _union_cell_range(excel, ws, value_col, data_rows)

        # The legend name, category labels, and values all stay linked to the
        # visible report cells. Editing a LINE header also updates the legend.
        series.Name = _worksheet_cell_reference(ws, header_row, value_col)
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
    """Fit the chart exactly inside its two template columns and row block."""
    if last_row < first_row:
        raise ChartGeneratorError(
            f"Rentang baris grafik tidak valid: {first_row}:{last_row}."
        )

    # Template graph area is always two columns wide.
    anchor_range = ws.Range(
        ws.Cells(first_row, chart_col),
        ws.Cells(last_row, chart_col + CHART_COLUMN_SPAN - 1),
    )
    chart_object.Left = float(anchor_range.Left)
    chart_object.Top = float(anchor_range.Top)
    chart_object.Width = float(anchor_range.Width)
    chart_object.Height = float(anchor_range.Height)
    # Keep the object aligned to the cells if rows/columns are resized.
    chart_object.Placement = 1  # xlMoveAndSize


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

        # All chart series link to the report's visible cells. No hidden helper
        # table is written; the category/value ranges are non-contiguous because
        # each template data item occupies a two-row merged block.
        row_plan = _build_table_row_plan(
            layout, machine_count, process_count
        )

        stage = "Membaca tiga grafik bawaan template"
        summary_chart, machine_chart, process_chart = _get_chart_objects(ws)

        stage = "Memperbarui grafik Summary dari tabel"
        _update_chart_series_from_table(
            excel, summary_chart, ws, labels,
            header_row=3,
            line_start_col=3,
            category_col=2,
            data_rows=row_plan["summary"],
        )
        _position_chart(
            summary_chart, ws, chart_col,
            first_row=layout["summary_chart_start_row"],
            last_row=layout["summary_chart_end_row"],
        )

        stage = "Memperbarui grafik Penggunaan Mesin dari tabel"
        _update_chart_series_from_table(
            excel, machine_chart, ws, labels,
            header_row=16,
            line_start_col=3,
            category_col=2,
            data_rows=row_plan["machines"],
        )
        _position_chart(
            machine_chart, ws, chart_col,
            first_row=layout["machine_chart_start_row"],
            last_row=layout["machine_chart_end_row"],
        )

        stage = "Memperbarui grafik Penggunaan Operator dari tabel"
        _update_chart_series_from_table(
            excel, process_chart, ws, labels,
            header_row=section3_header_row,
            line_start_col=3,
            category_col=2,
            data_rows=row_plan["processes"],
        )
        _position_chart(
            process_chart, ws, chart_col,
            first_row=layout["process_chart_start_row"],
            last_row=layout["process_chart_end_row"],
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