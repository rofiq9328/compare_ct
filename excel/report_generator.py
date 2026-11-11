"""Generate a single-sheet COMPARE LINE workbook from comparison and analysis JSON.

Excel automation uses Microsoft Excel through pywin32 so the supplied template's
formatting is retained. This stage prepares report data and removes stale sample
charts; the chart objects are added by the next chart-generator stage.

Run from the project root:
    py -m excel.report_generator

Optional explicit paths:
    py -m excel.report_generator \
        "logs\\comparison_data.json" "logs\\analysis_data.json" \
        --template "templates\\template_compare.xlsx" \
        --output "output\\COMPARE_LINE.xlsx"
"""

from __future__ import annotations

import argparse
import json
import math
import os
import sys
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Sequence, Tuple


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_COMPARISON_PATH = PROJECT_ROOT / "logs" / "comparison_data.json"
DEFAULT_ANALYSIS_PATH = PROJECT_ROOT / "logs" / "analysis_data.json"
DEFAULT_TEMPLATE_PATH = PROJECT_ROOT / "templates" / "template_compare.xlsx"
DEFAULT_OUTPUT_PATH = PROJECT_ROOT / "output" / "COMPARE_LINE.xlsx"

SUMMARY_ROW_KEYS: Tuple[str, ...] = (
    "operator_allocations",
    "takt_time",
    "target_per_hour",
    "potential_per_hour",
    "actual_per_hour",
)
SUMMARY_TEMPLATE_ROWS = (4, 6, 8, 10, 12)
MACHINE_FIRST_ROW = 17
MACHINE_TEMPLATE_CAPACITY = 6
MACHINE_TOTAL_TEMPLATE_ROW = 29

# Excel constants; keeping them local avoids requiring win32com at import time.
XL_OPENXML_WORKBOOK = 51
XL_PASTE_FORMATS = -4122
XL_CENTER = -4108
XL_LEFT = -4131
XL_TOP = -4160
XL_CONTINUOUS = 1
XL_THIN = 2
XL_EDGE_BOTTOM = 9


class ReportGeneratorError(ValueError):
    """Raised when the report data or template cannot be used safely."""


def read_json_object(path: str | Path, label: str) -> Dict[str, Any]:
    file_path = Path(path).expanduser()
    if not file_path.is_absolute():
        file_path = (Path.cwd() / file_path).resolve()
    if not file_path.is_file():
        raise FileNotFoundError(f"{label} tidak ditemukan: {file_path}")
    try:
        data = json.loads(file_path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        raise ReportGeneratorError(f"{label} bukan JSON valid: {exc}") from exc
    if not isinstance(data, dict):
        raise ReportGeneratorError(f"Isi {label} harus berupa object JSON.")
    return data


def _is_number(value: Any) -> bool:
    return (
        isinstance(value, (int, float))
        and not isinstance(value, bool)
        and math.isfinite(float(value))
    )


def _as_count(value: Any, context: str) -> int:
    if not _is_number(value) or float(value) < 0 or not float(value).is_integer():
        raise ReportGeneratorError(f"Nilai jumlah tidak valid pada {context}: {value!r}.")
    return int(value)


def _value_for_line(values_by_line: Mapping[str, Any], label: str, context: str) -> float | int | None:
    value = values_by_line.get(label)
    if value is None or value == "":
        return None
    if not _is_number(value):
        raise ReportGeneratorError(f"Nilai {context} untuk {label} bukan angka: {value!r}.")
    return value


def validate_report_inputs(
    comparison: Mapping[str, Any],
    analysis: Mapping[str, Any],
) -> Dict[str, Any]:
    """Validate schemas, line order, and allocation reconciliation before Excel opens."""
    if comparison.get("schema_version") != 1:
        raise ReportGeneratorError("schema_version comparison_data.json harus 1.")
    if comparison.get("comparison_type") != "dynamic_multi_line":
        raise ReportGeneratorError("comparison_type harus 'dynamic_multi_line'.")

    columns = comparison.get("line_columns")
    if not isinstance(columns, list) or len(columns) < 2:
        raise ReportGeneratorError("line_columns harus berisi minimal dua line.")
    if comparison.get("line_count") != len(columns):
        raise ReportGeneratorError("line_count tidak sama dengan jumlah line_columns.")

    labels: List[str] = []
    for index, item in enumerate(columns, start=1):
        if not isinstance(item, dict):
            raise ReportGeneratorError(f"line_columns item ke-{index} bukan object.")
        label = str(item.get("line_label") or "").strip()
        if not label:
            raise ReportGeneratorError(f"line_label item ke-{index} kosong.")
        if label in labels:
            raise ReportGeneratorError(f"line_label duplikat: {label}.")
        labels.append(label)

    if comparison.get("quality", {}).get("valid") is not True:
        raise ReportGeneratorError("comparison_data.json belum berstatus quality.valid = true.")
    if analysis.get("schema_version") != 1:
        raise ReportGeneratorError("schema_version analysis_data.json harus 1.")
    if analysis.get("quality", {}).get("valid") is not True:
        raise ReportGeneratorError("analysis_data.json belum berstatus quality.valid = true.")
    if analysis.get("quality", {}).get("source_comparison_quality_valid") is not True:
        raise ReportGeneratorError("Analysis dibuat dari comparison_data yang tidak valid.")
    if analysis.get("line_count") != len(labels):
        raise ReportGeneratorError("line_count analysis tidak sama dengan comparison_data.")
    if list(analysis.get("line_labels") or []) != labels:
        raise ReportGeneratorError("Urutan line_labels pada analysis tidak sama dengan comparison_data.")

    summary_rows = comparison.get("summary_rows")
    machine_rows = comparison.get("machine_rows")
    process_rows = comparison.get("process_rows")
    if not isinstance(summary_rows, list) or not isinstance(machine_rows, list) or not isinstance(process_rows, list):
        raise ReportGeneratorError("summary_rows, machine_rows, dan process_rows harus berupa list.")

    summary_by_key: Dict[str, Mapping[str, Any]] = {}
    for row in summary_rows:
        if not isinstance(row, dict):
            raise ReportGeneratorError("Ada item summary_rows yang bukan object.")
        key = str(row.get("key") or "")
        if not key or key in summary_by_key:
            raise ReportGeneratorError(f"Key summary kosong atau duplikat: {key!r}.")
        summary_by_key[key] = row
        values = row.get("values_by_line") or {}
        if not isinstance(values, dict):
            raise ReportGeneratorError(f"values_by_line pada summary {key} harus object.")
        for label in labels:
            _value_for_line(values, label, f"summary {key}")
    missing_summary = [key for key in SUMMARY_ROW_KEYS if key not in summary_by_key]
    if missing_summary:
        raise ReportGeneratorError(f"Summary belum lengkap: {', '.join(missing_summary)}.")

    total_row = comparison.get("machine_total_row") or {}
    total_values = total_row.get("values_by_line") or {}
    machine_sum_by_line = {label: 0 for label in labels}
    process_operator_sum_by_line = {label: 0 for label in labels}
    process_machine_sum_by_line = {label: 0 for label in labels}

    for index, row in enumerate(machine_rows, start=1):
        if not isinstance(row, dict) or not str(row.get("machine") or "").strip():
            raise ReportGeneratorError(f"machine_rows item ke-{index} tidak memiliki nama mesin.")
        values = row.get("values_by_line") or {}
        if not isinstance(values, dict):
            raise ReportGeneratorError(f"values_by_line pada mesin ke-{index} harus object.")
        for label in labels:
            machine_sum_by_line[label] += _as_count(
                values.get(label, 0) or 0, f"mesin {row.get('machine')} / {label}"
            )

    for index, row in enumerate(process_rows, start=1):
        if not isinstance(row, dict) or not str(row.get("process") or "").strip():
            raise ReportGeneratorError(f"process_rows item ke-{index} tidak memiliki nama proses.")
        op_values = row.get("operator_allocations_by_line") or {}
        mc_values = row.get("machine_allocations_by_line") or {}
        if not isinstance(op_values, dict) or not isinstance(mc_values, dict):
            raise ReportGeneratorError(f"Alokasi proses ke-{index} harus berupa object per line.")
        for label in labels:
            process_operator_sum_by_line[label] += _as_count(
                op_values.get(label, 0) or 0, f"operator proses {row.get('process')} / {label}"
            )
            process_machine_sum_by_line[label] += _as_count(
                mc_values.get(label, 0) or 0, f"mesin proses {row.get('process')} / {label}"
            )

    line_audits = comparison.get("quality", {}).get("line_audits") or []
    audit_by_line = {
        str(item.get("line_label")): item
        for item in line_audits
        if isinstance(item, dict) and item.get("line_label")
    }
    missing_audits = [label for label in labels if label not in audit_by_line]
    if missing_audits:
        raise ReportGeneratorError(f"Audit tidak tersedia untuk: {', '.join(missing_audits)}.")

    for label in labels:
        audit = audit_by_line[label]
        summary_operator = _as_count(
            summary_by_key["operator_allocations"]["values_by_line"].get(label),
            f"summary operator {label}",
        )
        summary_machine = _as_count(total_values.get(label), f"total mesin {label}")
        if summary_operator != process_operator_sum_by_line[label]:
            raise ReportGeneratorError(
                f"Total operator {label} tidak cocok: summary {summary_operator}, "
                f"jumlah detail proses {process_operator_sum_by_line[label]}."
            )
        if summary_machine != machine_sum_by_line[label]:
            raise ReportGeneratorError(
                f"Total mesin {label} tidak cocok: total row {summary_machine}, "
                f"jumlah katalog mesin {machine_sum_by_line[label]}."
            )
        if summary_machine != process_machine_sum_by_line[label]:
            raise ReportGeneratorError(
                f"Total mesin {label} tidak cocok dengan alokasi proses: "
                f"total {summary_machine}, detail {process_machine_sum_by_line[label]}."
            )
        if audit.get("operator_match") is not True or audit.get("machine_match") is not True:
            raise ReportGeneratorError(f"Audit rekonsiliasi belum PASS untuk {label}.")
        for audit_key, expected in (
            ("summary_operator_allocations", summary_operator),
            ("process_operator_allocations", process_operator_sum_by_line[label]),
            ("summary_machine_allocations", summary_machine),
            ("machine_catalog_total", machine_sum_by_line[label]),
            ("process_machine_allocations", process_machine_sum_by_line[label]),
        ):
            if _as_count(audit.get(audit_key), f"audit {audit_key} {label}") != expected:
                raise ReportGeneratorError(f"Nilai audit {audit_key} tidak konsisten untuk {label}.")

    if len(process_rows) != int(analysis.get("process_analysis", {}).get("process_row_count", len(process_rows))):
        raise ReportGeneratorError("Jumlah process_rows tidak sama dengan analysis.process_row_count.")
    if len(machine_rows) != int(analysis.get("machine_analysis", {}).get("machine_type_count", len(machine_rows))):
        raise ReportGeneratorError("Jumlah machine_rows tidak sama dengan analysis.machine_type_count.")

    comparison_review = int(comparison.get("quality", {}).get("process_rows_requiring_review", 0) or 0)
    analysis_review = int(analysis.get("quality", {}).get("process_matching_review_count", 0) or 0)
    comparison_candidates = int(comparison.get("quality", {}).get("pending_process_match_candidates", 0) or 0)
    analysis_candidates = int(analysis.get("quality", {}).get("process_matching_candidate_count", 0) or 0)
    review_status_count = sum(
        1 for row in process_rows
        if str(row.get("matching_status") or "").upper() == "REVIEW_REQUIRED"
    )
    if review_status_count != comparison_review:
        raise ReportGeneratorError(
            f"Jumlah status REVIEW_REQUIRED ({review_status_count}) tidak sama dengan quality ({comparison_review})."
        )
    if comparison_review != analysis_review or comparison_candidates != analysis_candidates:
        raise ReportGeneratorError("Jumlah review/kandidat pada analysis tidak sama dengan comparison_data.")

    return {
        "line_labels": labels,
        "summary_by_key": summary_by_key,
        "machine_rows": machine_rows,
        "process_rows": process_rows,
        "machine_total_row": total_row,
        "line_audits": audit_by_line,
        "summary_differences": analysis.get("summary_differences") or [],
        "line_kpis": analysis.get("line_kpis") or {},
        "review_count": comparison_review,
        "candidate_count": comparison_candidates,
    }


def _column_letter(column_number: int) -> str:
    """Convert a 1-based Excel column number to its letter representation."""
    result = ""
    number = int(column_number)
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _format_delta(delta: Any, percent: Any, unit: str) -> str:
    if not _is_number(delta):
        return ""
    sign = "+" if float(delta) > 0 else ""
    numeric = f"{sign}{float(delta):g}"
    text = f"{numeric} {unit}".strip()
    if _is_number(percent):
        pct_sign = "+" if float(percent) > 0 else ""
        text += f" ({pct_sign}{float(percent):.2f}%)"
    return text


def _summary_note(
    row_key: str,
    unit: str,
    labels: Sequence[str],
    summary_differences: Sequence[Mapping[str, Any]],
    line_kpis: Mapping[str, Mapping[str, Any]],
) -> str:
    difference = next((item for item in summary_differences if item.get("key") == row_key), None)
    parts: List[str] = []
    if difference:
        reference = difference.get("reference_line")
        values = difference.get("values_by_line") or {}
        for label in labels:
            if label == reference:
                continue
            item = values.get(label) or {}
            delta_text = _format_delta(
                item.get("delta_vs_reference"), item.get("delta_vs_reference_pct"), unit
            )
            if delta_text:
                parts.append(f"Selisih vs {reference}: {delta_text}.")
    if not parts:
        values = []
        for label in labels:
            value = (line_kpis.get(label) or {}).get(
                "operator_allocations" if row_key == "operator_allocations" else
                "takt_time_seconds" if row_key == "takt_time" else row_key
            )
            if _is_number(value):
                values.append(float(value))
        if values and all(value == values[0] for value in values):
            parts.append("Tidak ada selisih antar-line.")

    if row_key == "actual_per_hour":
        for label in labels:
            kpi = line_kpis.get(label) or {}
            actual = kpi.get("actual_per_hour")
            target = kpi.get("target_per_hour")
            pct = kpi.get("actual_vs_target_pct")
            gap = kpi.get("target_gap_pcs_per_hour")
            if _is_number(actual) and _is_number(target) and _is_number(pct):
                sentence = f"{label}: {actual:g}/{target:g} pcs/jam ({pct:.2f}% target)"
                if _is_number(gap) and float(gap) > 0:
                    sentence += f", kurang {gap:g} pcs/jam"
                parts.append(sentence + ".")
    elif row_key == "potential_per_hour":
        for label in labels:
            kpi = line_kpis.get(label) or {}
            potential = kpi.get("potential_per_hour")
            target = kpi.get("target_per_hour")
            gap = kpi.get("potential_gap_to_target_pcs_per_hour")
            if _is_number(potential) and _is_number(target) and _is_number(gap):
                if float(gap) > 0:
                    parts.append(
                        f"{label}: potensial {potential:g} masih {gap:g} pcs/jam di bawah target {target:g}."
                    )
    return " ".join(parts)


def _machine_note(row: Mapping[str, Any], labels: Sequence[str]) -> str:
    present = list(row.get("present_in_lines") or [])
    if len(present) == len(labels):
        note = "Digunakan pada seluruh line."
    elif present:
        note = "Hanya tercatat pada: " + ", ".join(present) + "."
    else:
        note = "Tidak ada alokasi pada line yang dibandingkan."
    if row.get("alias_applied"):
        note += " Alias mesin diterapkan."
    return note


def _process_note(row: Mapping[str, Any], labels: Sequence[str]) -> str:
    status = str(row.get("matching_status") or "").upper()
    parts: List[str] = []
    candidates = row.get("review_candidates") or []

    if status == "APPROVED_ALIAS":
        parts.append("ALIAS DISETUJUI.")
    elif status == "REVIEW_REQUIRED":
        parts.append(
            f"REVIEW_REQUIRED: {len(candidates)} kandidat; belum dikonfirmasi dan tidak dianggap proses yang sama."
        )
    elif status == "ONLY_IN_LISTED_LINES":
        present = list(row.get("present_in_lines") or [])
        parts.append("Hanya tercatat pada: " + (", ".join(present) if present else "tidak diketahui") + ".")
    elif status:
        parts.append(status + ".")

    groups_by_line = row.get("groups_by_line") or {}
    normalized_groups = {
        label: [str(group) for group in (groups_by_line.get(label) or []) if str(group).strip()]
        for label in labels
    }
    non_empty_group_sets = {tuple(value) for value in normalized_groups.values() if value}
    if len(non_empty_group_sets) > 1:
        details = []
        for label in labels:
            groups = normalized_groups[label]
            if groups:
                details.append(f"{label}: {' / '.join(groups)}")
        parts.append("Group berbeda: " + "; ".join(details) + ".")
    else:
        first_groups = next(iter(non_empty_group_sets), ())
        if first_groups:
            parts.append("Group: " + " / ".join(first_groups) + ".")

    return " ".join(parts)


def _set_range_value(ws: Any, row: int, column: int, value: Any) -> None:
    ws.Cells(row, column).Value = value


def _copy_row_formats(excel: Any, ws: Any, source_row: int, start_row: int, end_row: int) -> None:
    if end_row <= start_row:
        return
    ws.Rows(source_row).Copy()
    ws.Rows(f"{start_row}:{end_row}").PasteSpecial(Paste=XL_PASTE_FORMATS)
    excel.CutCopyMode = False


def _prepare_template_structure(excel: Any, ws: Any, machine_count: int, line_count: int) -> Dict[str, int]:
    """Resize template's machine rows and value/chart/note columns."""
    # Machine rows in the original template have capacity for six machine types.
    if machine_count > MACHINE_TEMPLATE_CAPACITY:
        extra_rows = 2 * (machine_count - MACHINE_TEMPLATE_CAPACITY)
        ws.Rows(f"{MACHINE_TOTAL_TEMPLATE_ROW}:{MACHINE_TOTAL_TEMPLATE_ROW + extra_rows - 1}").Insert()
        # Match inserted machine data rows to the existing machine data style.
        for index in range(MACHINE_TEMPLATE_CAPACITY, machine_count):
            target_row = MACHINE_FIRST_ROW + 2 * index
            ws.Rows(MACHINE_FIRST_ROW + 2 * (MACHINE_TEMPLATE_CAPACITY - 1)).Copy()
            ws.Rows(target_row).PasteSpecial(Paste=XL_PASTE_FORMATS)
            excel.CutCopyMode = False
    elif machine_count < MACHINE_TEMPLATE_CAPACITY:
        # Retain one blank spacer before TOTAL MESIN, but remove unused machine slots.
        desired_total_row = MACHINE_FIRST_ROW + (2 * machine_count)
        if desired_total_row < MACHINE_TOTAL_TEMPLATE_ROW:
            delete_start = desired_total_row
            delete_end = MACHINE_TOTAL_TEMPLATE_ROW - 1
            ws.Rows(f"{delete_start}:{delete_end}").Delete()

    total_row = MACHINE_FIRST_ROW + (2 * machine_count)
    section3_row = total_row + 3
    section3_header_row = section3_row + 1
    process_start_row = section3_header_row + 1

    # There are two line-value columns in the base template. Insert more before chart area E.
    extra_columns = line_count - 2
    if extra_columns > 0:
        ws.Columns(f"E:{_column_letter(4 + extra_columns)}").Insert()
        # Copy value-column formatting to newly inserted line columns.
        ws.Columns("D:D").Copy()
        ws.Columns(f"E:{_column_letter(4 + extra_columns)}").PasteSpecial(Paste=XL_PASTE_FORMATS)
        excel.CutCopyMode = False

    chart_col = 3 + line_count
    note_col = chart_col + 2

    # Delete old template charts: they contain example data and will be rebuilt next stage.
    try:
        ws.ChartObjects().Delete()
    except Exception:
        # Some Excel versions return an empty collection that cannot be deleted.
        pass

    # Clear template values only; all existing cell formatting is kept.
    last_used_row = max(process_start_row + 100, int(ws.UsedRange.Rows.Count) + 10)
    ws.Range(ws.Cells(1, 2), ws.Cells(last_used_row, note_col)).ClearContents()

    return {
        "total_row": total_row,
        "section3_row": section3_row,
        "section3_header_row": section3_header_row,
        "process_start_row": process_start_row,
        "chart_col": chart_col,
        "note_col": note_col,
        "last_column": note_col,
    }


def _populate_workbook(excel: Any, ws: Any, data: Mapping[str, Any]) -> Dict[str, Any]:
    labels: Sequence[str] = data["line_labels"]
    summary_by_key: Mapping[str, Mapping[str, Any]] = data["summary_by_key"]
    machine_rows: Sequence[Mapping[str, Any]] = data["machine_rows"]
    process_rows: Sequence[Mapping[str, Any]] = data["process_rows"]
    line_columns_start = 3  # C
    layout = _prepare_template_structure(excel, ws, len(machine_rows), len(labels))
    chart_col = layout["chart_col"]
    note_col = layout["note_col"]
    last_col = layout["last_column"]
    section3_row = layout["section3_row"]
    section3_header_row = layout["section3_header_row"]
    process_start = layout["process_start_row"]
    total_row = layout["total_row"]

    # Workbook and section headings.
    ws.Cells(1, 2).Value = "COMPARE LINE"
    ws.Cells(2, 2).Value = "I. SUMMARY"
    ws.Cells(15, 2).Value = "II. PENGGUNAAN MESIN"
    ws.Cells(section3_row, 2).Value = "III. PENGGUNAAN OPERATOR"

    # Header rows for each of the three sections.
    for header_row, variable_label in (
        (3, "VARIABEL"),
        (16, "VARIABEL / PARAMETER"),
        (section3_header_row, "PROSES"),
    ):
        ws.Cells(header_row, 2).Value = variable_label
        for index, label in enumerate(labels):
            ws.Cells(header_row, line_columns_start + index).Value = label
        ws.Cells(header_row, chart_col).Value = "GRAFIK"
        ws.Cells(header_row, note_col).Value = "KETERANGAN"

    # I. SUMMARY.
    summary_key_to_label = {
        "operator_allocations": "Jumlah Operator",
        "takt_time": "Takt Time",
        "target_per_hour": "Target Output / Jam",
        "potential_per_hour": "Potensial Output / Jam",
        "actual_per_hour": "Aktual Output / Jam",
    }
    summary_units = {
        "operator_allocations": "OPT",
        "takt_time": "detik",
        "target_per_hour": "pcs",
        "potential_per_hour": "pcs",
        "actual_per_hour": "pcs",
    }
    for key, row_number in zip(SUMMARY_ROW_KEYS, SUMMARY_TEMPLATE_ROWS):
        item = summary_by_key[key]
        ws.Cells(row_number, 2).Value = item.get("label") or summary_key_to_label[key]
        values = item.get("values_by_line") or {}
        for index, label in enumerate(labels):
            value = values.get(label)
            _set_range_value(ws, row_number, line_columns_start + index, value)
        ws.Cells(row_number, note_col).Value = _summary_note(
            key,
            str(item.get("unit") or summary_units[key]),
            labels,
            data["summary_differences"],
            data["line_kpis"],
        )

    # II. PENGGUNAAN MESIN.
    machine_data_rows = [MACHINE_FIRST_ROW + 2 * index for index in range(len(machine_rows))]
    for index, (item, row_number) in enumerate(zip(machine_rows, machine_data_rows)):
        ws.Cells(row_number, 2).Value = item.get("machine")
        values = item.get("values_by_line") or {}
        for line_index, label in enumerate(labels):
            value = values.get(label, 0)
            ws.Cells(row_number, line_columns_start + line_index).Value = value
        ws.Cells(row_number, note_col).Value = _machine_note(item, labels)
        if item.get("alias_applied"):
            ws.Cells(row_number, note_col).Font.Color = 0x006600

    ws.Cells(total_row, 2).Value = "TOTAL MESIN"
    total_values = data["machine_total_row"].get("values_by_line") or {}
    for index, label in enumerate(labels):
        ws.Cells(total_row, line_columns_start + index).Value = total_values.get(label, 0)
    ws.Cells(total_row, note_col).Value = "Total alokasi mesin; direkonsiliasi dengan katalog mesin dan detail proses."

    # III. PENGGUNAAN OPERATOR. Make the detailed process list compact (one process per row).
    if process_rows:
        process_end = process_start + len(process_rows) - 1
        _copy_row_formats(excel, ws, process_start, process_start + 1, process_end)
        for index, item in enumerate(process_rows):
            row_number = process_start + index
            ws.Cells(row_number, 2).Value = item.get("process")
            op_values = item.get("operator_allocations_by_line") or {}
            for line_index, label in enumerate(labels):
                ws.Cells(row_number, line_columns_start + line_index).Value = op_values.get(label, 0)
            note = _process_note(item, labels)
            ws.Cells(row_number, note_col).Value = note
            status = str(item.get("matching_status") or "").upper()
            if status == "REVIEW_REQUIRED":
                ws.Cells(row_number, note_col).Interior.Color = 0xD9EAF7  # pale blue review marker
                ws.Cells(row_number, note_col).Font.Color = 0x7F6000
            elif status == "APPROVED_ALIAS":
                ws.Cells(row_number, note_col).Font.Color = 0x006100
            elif status == "ONLY_IN_LISTED_LINES":
                ws.Cells(row_number, note_col).Font.Color = 0x666666
            ws.Rows(row_number).RowHeight = 30 if len(str(item.get("process") or "")) > 42 or len(note) > 85 else 22
    else:
        process_end = process_start
        ws.Cells(process_start, 2).Value = "Tidak ada detail proses."

    # Readability and alignment.
    ws.Columns("A:A").ColumnWidth = 3
    ws.Columns("B:B").ColumnWidth = 42
    for index in range(len(labels)):
        col = _column_letter(line_columns_start + index)
        ws.Columns(f"{col}:{col}").ColumnWidth = 17
    ws.Columns(f"{_column_letter(chart_col)}:{_column_letter(chart_col + 1)}").ColumnWidth = 12
    ws.Columns(f"{_column_letter(note_col)}:{_column_letter(note_col)}").ColumnWidth = 48

    # Apply consistent wrap/alignment to labels, values, and notes.
    ws.Range(ws.Cells(3, 2), ws.Cells(process_end, note_col)).VerticalAlignment = XL_TOP
    ws.Range(ws.Cells(3, 2), ws.Cells(3, note_col)).WrapText = True
    ws.Range(ws.Cells(4, 2), ws.Cells(12, 2)).WrapText = True
    ws.Range(ws.Cells(16, 2), ws.Cells(16, note_col)).WrapText = True
    ws.Range(ws.Cells(section3_header_row, 2), ws.Cells(section3_header_row, note_col)).WrapText = True
    for header_row in (3, 16, section3_header_row):
        ws.Range(ws.Cells(header_row, 2), ws.Cells(header_row, note_col)).HorizontalAlignment = XL_CENTER
        ws.Rows(header_row).RowHeight = 32
    ws.Range(ws.Cells(4, line_columns_start), ws.Cells(12, line_columns_start + len(labels) - 1)).HorizontalAlignment = XL_CENTER
    if machine_data_rows:
        ws.Range(
            ws.Cells(MACHINE_FIRST_ROW, line_columns_start),
            ws.Cells(machine_data_rows[-1], line_columns_start + len(labels) - 1),
        ).HorizontalAlignment = XL_CENTER
    if process_rows:
        ws.Range(
            ws.Cells(process_start, 2),
            ws.Cells(process_end, 2),
        ).WrapText = True
        ws.Range(
            ws.Cells(process_start, line_columns_start),
            ws.Cells(process_end, line_columns_start + len(labels) - 1),
        ).HorizontalAlignment = XL_CENTER

    # Number formats: only potential output needs decimals; other counts are integers.
    for key, row_number in zip(SUMMARY_ROW_KEYS, SUMMARY_TEMPLATE_ROWS):
        fmt = "0.00" if key == "potential_per_hour" else "0"
        ws.Range(
            ws.Cells(row_number, line_columns_start),
            ws.Cells(row_number, line_columns_start + len(labels) - 1),
        ).NumberFormat = fmt
    if machine_data_rows:
        ws.Range(
            ws.Cells(MACHINE_FIRST_ROW, line_columns_start),
            ws.Cells(machine_data_rows[-1], line_columns_start + len(labels) - 1),
        ).NumberFormat = "0"
    ws.Range(
        ws.Cells(total_row, line_columns_start),
        ws.Cells(total_row, line_columns_start + len(labels) - 1),
    ).NumberFormat = "0"
    if process_rows:
        ws.Range(
            ws.Cells(process_start, line_columns_start),
            ws.Cells(process_end, line_columns_start + len(labels) - 1),
        ).NumberFormat = "0"

    # Make note cells legible and keep data visible without autofitting excessively.
    for row_number in list(SUMMARY_TEMPLATE_ROWS) + machine_data_rows + [total_row]:
        ws.Cells(row_number, 2).WrapText = True
        ws.Cells(row_number, note_col).WrapText = True
        ws.Cells(row_number, note_col).VerticalAlignment = XL_TOP
        if row_number in SUMMARY_TEMPLATE_ROWS:
            ws.Rows(row_number).RowHeight = 28
        elif row_number in machine_data_rows:
            ws.Rows(row_number).RowHeight = 26
    if process_rows:
        notes_range = ws.Range(ws.Cells(process_start, note_col), ws.Cells(process_end, note_col))
        notes_range.WrapText = True
        notes_range.VerticalAlignment = XL_TOP
        # Thin separators make the long process list easy to scan.
        try:
            border = notes_range.Borders(XL_EDGE_BOTTOM)
            border.LineStyle = XL_CONTINUOUS
            border.Weight = XL_THIN
        except Exception:
            pass

    # Preserve the total row's emphasis even if the template's row was shifted.
    ws.Range(ws.Cells(total_row, 2), ws.Cells(total_row, note_col)).Font.Bold = True

    return {
        **layout,
        "machine_data_rows": machine_data_rows,
        "process_end_row": process_end,
        "process_count": len(process_rows),
        "machine_count": len(machine_rows),
        "line_count": len(labels),
    }


def generate_report(
    comparison_path: str | Path = DEFAULT_COMPARISON_PATH,
    analysis_path: str | Path = DEFAULT_ANALYSIS_PATH,
    template_path: str | Path = DEFAULT_TEMPLATE_PATH,
    output_path: str | Path = DEFAULT_OUTPUT_PATH,
) -> Dict[str, Any]:
    """Create the COMPARE LINE workbook; requires Windows + Microsoft Excel + pywin32."""
    comparison_file = Path(comparison_path).expanduser().resolve()
    analysis_file = Path(analysis_path).expanduser().resolve()
    template_file = Path(template_path).expanduser().resolve()
    output_file = Path(output_path).expanduser().resolve()

    comparison = read_json_object(comparison_file, "comparison_data.json")
    analysis = read_json_object(analysis_file, "analysis_data.json")
    data = validate_report_inputs(comparison, analysis)

    if not template_file.is_file():
        raise FileNotFoundError(f"Template Excel tidak ditemukan: {template_file}")
    output_file.parent.mkdir(parents=True, exist_ok=True)

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
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(template_file), UpdateLinks=0, ReadOnly=False)
        if workbook.Worksheets.Count != 1:
            raise ReportGeneratorError(
                f"Template harus hanya memiliki satu sheet; ditemukan {workbook.Worksheets.Count}."
            )

        ws = workbook.Worksheets(1)
        if str(ws.Name).strip().lower() != "compare line":
            raise ReportGeneratorError(
                f"Nama sheet template harus 'compare line'; ditemukan {ws.Name!r}."
            )

        report_info = _populate_workbook(excel, ws, data)
        ws.Activate()
        ws.Range("B1").Select()
        if output_file.exists():
            output_file.unlink()
        workbook.SaveAs(str(output_file), FileFormat=XL_OPENXML_WORKBOOK)
        workbook.Close(SaveChanges=False)
        workbook = None

        return {
            "output": str(output_file),
            "sheet": "compare line",
            "line_count": report_info["line_count"],
            "machine_count": report_info["machine_count"],
            "process_count": report_info["process_count"],
            "review_required": data["review_count"],
            "pending_candidates": data["candidate_count"],
            "chart_status": "CHART_BELUM_DIBUAT_TAHAP_7",
            "status": "PASS",
        }
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
    return "\n".join(
        [
            "=" * 72,
            "COMPARE_CT — TAHAP 6: EXCEL REPORT GENERATOR",
            "=" * 72,
            f"OUTPUT              : {result['output']}",
            f"SHEET               : {result['sheet']}",
            f"JUMLAH LINE         : {result['line_count']}",
            f"JENIS MESIN         : {result['machine_count']}",
            f"BARIS PROSES        : {result['process_count']}",
            f"PROSES REVIEW       : {result['review_required']}",
            f"KANDIDAT PENDING    : {result['pending_candidates']}",
            f"GRAFIK              : {result['chart_status']}",
            f"STATUS              : {result['status']}",
            "=" * 72,
        ]
    )


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Buat workbook COMPARE LINE dari comparison_data.json dan analysis_data.json."
    )
    parser.add_argument("comparison", nargs="?", default=str(DEFAULT_COMPARISON_PATH))
    parser.add_argument("analysis", nargs="?", default=str(DEFAULT_ANALYSIS_PATH))
    parser.add_argument("--template", default=str(DEFAULT_TEMPLATE_PATH))
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT_PATH))
    args = parser.parse_args(argv)

    try:
        result = generate_report(
            comparison_path=args.comparison,
            analysis_path=args.analysis,
            template_path=args.template,
            output_path=args.output,
        )
    except Exception as exc:
        print("STATUS              : ERROR")
        print(f"PENYEBAB            : {exc}")
        return 1

    print(format_console_report(result))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
