
"""Add dynamic charts to the single-sheet COMPARE LINE workbook.

Run from the COMPARE_CT project root:
    py -m excel.chart_generator

Optional:
    py -m excel.chart_generator "output\\COMPARE_LINE.xlsx" \
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

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_WORKBOOK = PROJECT_ROOT / "output" / "COMPARE_LINE.xlsx"
DEFAULT_COMPARISON = PROJECT_ROOT / "logs" / "comparison_data.json"

# Excel constants
XL_COLUMN_CLUSTERED = 51
XL_BAR_CLUSTERED = 57
XL_LEGEND_POSITION_BOTTOM = -4107
XL_LABEL_POSITION_OUTSIDE_END = 2
XL_CATEGORY = 1
XL_VALUE = 2
XL_PRIMARY = 1
XL_ROW_FIELD = 1
XL_COLUMN_FIELD = 2


class ChartGeneratorError(ValueError):
    pass


def _read_json(path: Path) -> Dict[str, Any]:
    if not path.is_file():
        raise FileNotFoundError(
            f"comparison_data.json tidak ditemukan: {path}"
        )

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
) -> Any:

    rows = len(matrix)
    cols = max((len(row) for row in matrix), default=0)

    if not rows or not cols:
        raise ChartGeneratorError("Data sumber grafik kosong.")

    normalized = [
        list(row) + [None] * (cols - len(row))
        for row in matrix
    ]

    target = ws.Range(
        ws.Cells(start_row, start_col),
        ws.Cells(
            start_row + rows - 1,
            start_col + cols - 1
        ),
    )

    target.Value = tuple(
        tuple(row) for row in normalized
    )

    return target


def _add_chart(
    excel: Any,
    ws: Any,
    source_range: Any,
    anchor_row: int,
    anchor_col: int,
    width: float,
    height: float,
    chart_type: int,
    title: str,
    *,
    top_n_note: str = "",
) -> Any:

    left = ws.Cells(anchor_row, anchor_col).Left
    top = ws.Cells(anchor_row, anchor_col).Top

    chart_object = ws.ChartObjects().Add(
        left,
        top,
        width,
        height
    )

    chart = chart_object.Chart

    chart.ChartType = chart_type

    chart.SetSourceData(
        Source=source_range,
        PlotBy=XL_COLUMN_FIELD
    )

    chart.HasTitle = True
    chart.ChartTitle.Text = title

    chart.HasLegend = True
    chart.Legend.Position = XL_LEGEND_POSITION_BOTTOM

    chart.DisplayBlanksAs = 0

    try:
        chart.Axes(XL_CATEGORY).TickLabels.Font.Size = 8
        chart.Axes(XL_VALUE).TickLabels.Font.Size = 8
        chart.Axes(XL_VALUE).HasMajorGridlines = True

        chart.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 11
        chart.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = True

    except Exception:
        pass

    try:
        chart.ChartArea.Format.Line.Visible = False
        chart.PlotArea.Format.Line.Visible = False

    except Exception:
        pass

    if top_n_note:
        try:
            chart.HasTitle = True
            chart.ChartTitle.Text = title + "\n" + top_n_note
            chart.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 10

        except Exception:
            pass

    return chart_object


def _build_chart_data(
    data: Mapping[str, Any]
) -> Dict[str, Any]:

    columns = data.get("line_columns") or []

    labels = [
        str(item.get("line_label") or "").strip()
        for item in columns
    ]

    if len(labels) < 2 or any(not label for label in labels):
        raise ChartGeneratorError(
            "line_columns harus berisi minimal dua line dengan label valid."
        )

    summary = {
        str(row.get("key")): row
        for row in data.get("summary_rows") or []
    }

    required = (
        "target_per_hour",
        "potential_per_hour",
        "actual_per_hour",
    )

    for key in required:
        if key not in summary:
            raise ChartGeneratorError(
                f"Summary tidak memiliki key {key}."
            )

    summary_matrix: List[List[Any]] = [
        ["OUTPUT / JAM (PCS)"] + labels
    ]

    summary_labels = {
        "target_per_hour": "Target",
        "potential_per_hour": "Potensial",
        "actual_per_hour": "Aktual",
    }

    for key in required:
        values = summary[key].get("values_by_line") or {}

        summary_matrix.append(
            [
                summary_labels[key]
            ] + [
                _safe_number(values.get(label))
                for label in labels
            ]
        )

    machine_matrix: List[List[Any]] = [
        ["JENIS MESIN"] + labels
    ]

    for row in data.get("machine_rows") or []:
        machine = str(row.get("machine") or "").strip()

        if not machine:
            continue

        values = row.get("values_by_line") or {}

        machine_matrix.append(
            [machine] + [
                _safe_number(values.get(label))
                for label in labels
            ]
        )

    if len(machine_matrix) == 1:
        raise ChartGeneratorError(
            "Tidak ada data jenis mesin untuk grafik."
        )

    process_rows = []

    for row in data.get("process_rows") or []:
        process = str(row.get("process") or "").strip()

        if not process:
            continue

        values = row.get("operator_allocations_by_line") or {}

        per_line = [
            _safe_number(values.get(label))
            for label in labels
        ]

        total = sum(per_line)

        if total > 0:
            process_rows.append(
                (total, process, per_line)
            )

    process_rows.sort(
        key=lambda item: (
            -item[0],
            item[1].casefold()
        )
    )

    top_processes = process_rows[:15]

    process_matrix: List[List[Any]] = [
        ["PROSES (TOP 15)"] + labels
    ]

    for _, process, per_line in top_processes:
        process_matrix.append(
            [process] + per_line
        )

    if len(process_matrix) == 1:
        process_matrix.append(
            ["Tidak ada alokasi"] + [0 for _ in labels]
        )

    return {
        "labels": labels,
        "summary_matrix": summary_matrix,
        "machine_matrix": machine_matrix,
        "process_matrix": process_matrix,
        "machine_count": len(machine_matrix) - 1,
        "process_chart_count": len(process_matrix) - 1,
    }


def generate_charts(
    workbook_path: str | Path = DEFAULT_WORKBOOK,
    comparison_path: str | Path = DEFAULT_COMPARISON,
) -> Dict[str, Any]:

    current_stage = "Persiapan file"

    workbook_file = Path(workbook_path).expanduser().resolve()
    comparison_file = Path(comparison_path).expanduser().resolve()

    if not workbook_file.is_file():
        raise FileNotFoundError(
            f"Workbook Tahap 6 tidak ditemukan: {workbook_file}"
        )

    current_stage = "Membaca comparison_data.json"

    data = _read_json(comparison_file)
    chart_data = _build_chart_data(data)

    try:
        import pythoncom
        import win32com.client

    except ImportError as exc:
        raise RuntimeError(
            "pywin32 belum terpasang. Jalankan: "
            "py -m pip install pywin32"
        ) from exc

    pythoncom.CoInitialize()

    excel = None
    workbook = None

    try:
        current_stage = "Membuka Excel COM"

        excel = win32com.client.DispatchEx(
            "Excel.Application"
        )

        # Tidak mengatur Visible=False karena pada beberapa
        # instalasi Excel pengaturan ini memicu OLE error.
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        current_stage = "Membuka workbook"

        workbook = excel.Workbooks.Open(
            str(workbook_file),
            UpdateLinks=0,
            ReadOnly=False
        )

        current_stage = "Memeriksa jumlah worksheet"

        if workbook.Worksheets.Count != 1:
            raise ChartGeneratorError(
                "Workbook harus hanya memiliki satu sheet; "
                f"ditemukan {workbook.Worksheets.Count}."
            )

        current_stage = "Memeriksa nama worksheet"

        ws = workbook.Worksheets(1)

        sheet_name = str(ws.Name)

        if sheet_name.strip().lower() != "compare line":
            raise ChartGeneratorError(
                "Nama sheet harus 'compare line'; "
                f"ditemukan {sheet_name!r}."
            )

        labels = chart_data["labels"]

        line_count = len(labels)
        machine_count = chart_data["machine_count"]

        process_count = len(
            data.get("process_rows") or []
        )

        # Posisi dinamis sesuai struktur Tahap 6.
        machine_first_row = 17

        machine_total_row = (
            machine_first_row + 2 * machine_count
        )

        section3_row = machine_total_row + 3

        section3_header_row = section3_row + 1

        chart_col = 3 + line_count

        note_col = chart_col + 2

        # Mengatur lebar area grafik.
        current_stage = "Mengatur lebar kolom grafik"

        ws.Columns(
            f"{_col_letter(chart_col)}:"
            f"{_col_letter(chart_col + 1)}"
        ).ColumnWidth = 24

        # Data bantuan grafik disimpan pada sheet yang sama.
        current_stage = "Menentukan posisi helper data"

        helper_col = note_col + 3

        helper_start_rows = {
            "summary": 1,
            "machine": 8,
            "process": (
                8 + len(chart_data["machine_matrix"]) + 3
            ),
        }

        # Menulis data summary.
        current_stage = "Menulis helper data summary"

        summary_range = _write_matrix(
            ws,
            helper_start_rows["summary"],
            helper_col,
            chart_data["summary_matrix"],
        )

        # Menulis data mesin.
        current_stage = "Menulis helper data machine"

        machine_range = _write_matrix(
            ws,
            helper_start_rows["machine"],
            helper_col,
            chart_data["machine_matrix"],
        )

        # Menulis data proses.
        current_stage = "Menulis helper data process"

        process_range = _write_matrix(
            ws,
            helper_start_rows["process"],
            helper_col,
            chart_data["process_matrix"],
        )

        # Menyembunyikan kolom helper.
        current_stage = "Menyembunyikan kolom helper"

        ws.Columns(
            f"{_col_letter(helper_col)}:"
            f"{_col_letter(helper_col + line_count)}"
        ).Hidden = True

        # Menghapus grafik sebelumnya.
        current_stage = "Menghapus grafik lama"

        try:
            chart_objects = ws.ChartObjects()

            if chart_objects.Count > 0:
                chart_objects.Delete()

        except Exception:
            pass

        chart_width = 315

        # Grafik 1: Output per jam.
        current_stage = "Membuat grafik output per jam"

        summary_chart = _add_chart(
            excel,
            ws,
            summary_range,
            3,
            chart_col,
            chart_width,
            205,
            XL_COLUMN_CLUSTERED,
            "PERBANDINGAN OUTPUT PER JAM",
        )

        # Grafik 2: Penggunaan jenis mesin.
        current_stage = "Membuat grafik penggunaan mesin"

        machine_chart = _add_chart(
            excel,
            ws,
            machine_range,
            16,
            chart_col,
            chart_width,
            245,
            XL_BAR_CLUSTERED,
            "PENGGUNAAN JENIS MESIN PER LINE",
        )

        # Grafik 3: Penggunaan operator per proses.
        current_stage = "Membuat grafik penggunaan operator"

        process_chart = _add_chart(
            excel,
            ws,
            process_range,
            section3_header_row,
            chart_col,
            chart_width,
            360,
            XL_BAR_CLUSTERED,
            "PENGGUNAAN OPERATOR PER PROSES",
            top_n_note=(
                "Top 15 proses; tabel di kiri tetap "
                "menampilkan seluruh proses."
            ),
        )

        # Data labels untuk grafik pertama dan kedua.
        current_stage = "Mengatur data labels grafik"

        for chart_object in (
            summary_chart,
            machine_chart,
        ):
            try:
                chart = chart_object.Chart

                for series_index in range(
                    1,
                    chart.SeriesCollection().Count + 1
                ):
                    chart.SeriesCollection(
                        series_index
                    ).ApplyDataLabels()

                    chart.SeriesCollection(
                        series_index
                    ).DataLabels().Font.Size = 8

            except Exception:
                pass

        # Data labels grafik proses.
        try:
            for series_index in range(
                1,
                process_chart.Chart.SeriesCollection().Count + 1
            ):
                process_chart.Chart.SeriesCollection(
                    series_index
                ).ApplyDataLabels()

                process_chart.Chart.SeriesCollection(
                    series_index
                ).DataLabels().Font.Size = 8

        except Exception:
            pass

        # Simpan workbook.
        current_stage = "Menyimpan workbook"

        ws.Activate()
        ws.Range("B1").Select()

        workbook.Save()

        # Tutup workbook.
        current_stage = "Menutup workbook"

        workbook.Close(SaveChanges=False)
        workbook = None

        return {
            "output": str(workbook_file),
            "sheet": sheet_name,
            "line_count": line_count,
            "machine_count": machine_count,
            "process_count": process_count,
            "process_chart_count": chart_data[
                "process_chart_count"
            ],
            "chart_count": 3,
            "status": "PASS",
        }

    except Exception as exc:
        raise RuntimeError(
            f"Gagal pada tahap: {current_stage}\n"
            f"Penyebab: {exc}"
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


def format_console_report(
    result: Mapping[str, Any]
) -> str:

    return "\n".join([
        "=" * 72,
        "COMPARE_CT — TAHAP 7: CHART GENERATOR",
        "=" * 72,
        f"OUTPUT              : {result['output']}",
        f"SHEET               : {result['sheet']}",
        f"JUMLAH LINE         : {result['line_count']}",
        f"JENIS MESIN         : {result['machine_count']}",
        f"BARIS PROSES        : {result['process_count']}",
        (
            "PROSES DI GRAFIK    : "
            f"{result['process_chart_count']} (TOP 15)"
        ),
        f"JUMLAH GRAFIK       : {result['chart_count']}",
        f"STATUS              : {result['status']}",
        "=" * 72,
    ])


def main(
    argv: Sequence[str] | None = None
) -> int:

    parser = argparse.ArgumentParser(
        description=(
            "Tambahkan tiga grafik dinamis "
            "ke workbook COMPARE LINE."
        )
    )

    parser.add_argument(
        "workbook",
        nargs="?",
        default=str(DEFAULT_WORKBOOK),
    )

    parser.add_argument(
        "--comparison",
        default=str(DEFAULT_COMPARISON),
    )

    args = parser.parse_args(argv)

    try:
        result = generate_charts(
            args.workbook,
            args.comparison,
        )

    except Exception as exc:
        print("\n" + "=" * 72)
        print("COMPARE_CT — TAHAP 7: CHART GENERATOR")
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
