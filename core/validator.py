"""Validation layer for COMPARE_CT source workbooks.

This module validates the structured result returned by ``core.reader``.
It deliberately does not edit the reader or modify source workbooks.

Run from the project root on Windows with Microsoft Excel and pywin32:
    py -m core.validator "input\\LINE37.xlsx" "input\\LINE61.xlsx"
"""

from __future__ import annotations

import argparse
import math
import sys
from collections import Counter
from pathlib import Path
from typing import Any, Dict, Iterable, List, Sequence, Tuple

from .reader import SHEET_NAME, read_many


class ValidationError(RuntimeError):
    """Raised when a read result fails structural/data validation."""


def _is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(float(value))


def _as_int(value: Any) -> int | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float) and math.isfinite(value) and value.is_integer():
        return int(value)
    return None


def _same_value(left: Any, right: Any) -> bool:
    """Compare numeric Excel values safely while retaining exact text checks."""
    if _is_number(left) and _is_number(right):
        return math.isclose(float(left), float(right), rel_tol=1e-9, abs_tol=1e-9)
    return left == right


def _record_rows_by_source(rows: Sequence[Dict[str, Any]]) -> Dict[int, Dict[str, Any]]:
    result: Dict[int, Dict[str, Any]] = {}
    for row in rows:
        source_row = _as_int(row.get("source_row"))
        if source_row is not None and source_row not in result:
            result[source_row] = row
    return result


def validate_result(result: Dict[str, Any]) -> Dict[str, Any]:
    """Validate one parsed workbook result and return a detailed report.

    Warnings from the reader are retained as warnings. Structural mismatches,
    invalid metadata, and count inconsistencies are reported as errors.
    """
    errors: List[str] = []
    warnings: List[str] = []
    checks: List[Dict[str, Any]] = []

    def check(
        name: str,
        passed: bool,
        detail: str = "",
        success_detail: str = "",
    ) -> None:
        shown_detail = success_detail if passed else detail
        checks.append({"name": name, "passed": bool(passed), "detail": shown_detail})
        if not passed:
            errors.append(f"{name}: {detail or 'pemeriksaan gagal'}")

    if not isinstance(result, dict):
        return {
            "valid": False,
            "status": "FAIL",
            "source_file": "",
            "line_label": "",
            "checks": [],
            "errors": ["Hasil reader bukan dictionary."],
            "warnings": [],
        }

    source_file = str(result.get("source_file") or "")
    line_label = str(result.get("line_label") or "LINE TIDAK DIKETAHUI")
    summary = result.get("summary")
    rows = result.get("rows")
    operator_allocations = result.get("operator_allocations")
    machine_allocations = result.get("machine_allocations")
    machine_counts = result.get("machine_counts")
    process_summary = result.get("process_summary")

    # 1. Workbook/sheet identity and basic metadata.
    check("SOURCE_FILE", bool(source_file), "Nama/path file sumber kosong.")
    check(
        "SHEET_NAME",
        result.get("sheet_name") == SHEET_NAME,
        f"Sheet harus '{SHEET_NAME}', terbaca: {result.get('sheet_name')!r}.",
    )

    line_number = result.get("line_number")
    check(
        "LINE_NUMBER",
        _is_number(line_number) and float(line_number) > 0,
        f"Nomor line harus numerik dan positif, terbaca: {line_number!r}.",
    )
    check("STYLE", bool(str(result.get("style") or "").strip()), "Style pada F5 kosong.")

    # 2. Row boundaries.
    data_start = _as_int(result.get("data_start_row"))
    data_end = _as_int(result.get("data_end_row"))
    total_row = _as_int(result.get("total_row"))
    bounds_ok = (
        data_start is not None
        and data_end is not None
        and total_row is not None
        and data_start >= 1
        and total_row > data_start
        and data_end == total_row - 1
    )
    check(
        "ROW_BOUNDARIES",
        bounds_ok,
        f"Batas baris tidak konsisten: mulai={data_start}, akhir={data_end}, TOTAL={total_row}.",
    )

    if not isinstance(rows, list):
        check("ROWS_STRUCTURE", False, "Field 'rows' harus berupa list.")
        rows = []
    else:
        check("ROWS_STRUCTURE", True, f"{len(rows)} record terbaca.")

    source_row_numbers: List[int] = []
    malformed_rows = 0
    for row in rows:
        if not isinstance(row, dict):
            malformed_rows += 1
            continue
        source_row = _as_int(row.get("source_row"))
        if source_row is None:
            malformed_rows += 1
        else:
            source_row_numbers.append(source_row)
            if data_start is not None and data_end is not None:
                if not (data_start <= source_row <= data_end):
                    errors.append(
                        f"SOURCE_ROW_RANGE: baris {source_row} berada di luar rentang data "
                        f"{data_start}-{data_end}."
                    )
    check("ROW_RECORDS", malformed_rows == 0, f"{malformed_rows} record memiliki struktur/baris tidak valid.")
    check(
        "SOURCE_ROW_UNIQUE",
        len(source_row_numbers) == len(set(source_row_numbers)),
        "Ada source_row yang tercatat lebih dari satu kali." if len(source_row_numbers) != len(set(source_row_numbers)) else "",
    )
    check(
        "SOURCE_ROW_RANGE",
        not any(
            data_start is not None and data_end is not None and not (data_start <= n <= data_end)
            for n in source_row_numbers
        ),
        f"Seluruh record harus berada di baris {data_start}-{data_end}.",
    )

    # 3. Summary fields and numeric metadata.
    if not isinstance(summary, dict):
        check("SUMMARY_STRUCTURE", False, "Field 'summary' tidak ditemukan atau bukan dictionary.")
        summary = {}
    else:
        check("SUMMARY_STRUCTURE", True, "Field summary tersedia.")

    numeric_fields = {
        "target_per_hour": "TARGET_PER_HOUR",
        "potential_per_hour": "POTENTIAL_PER_HOUR",
        "actual_per_hour": "ACTUAL_PER_HOUR",
        "takt_time": "TAKT_TIME",
    }
    for field, label in numeric_fields.items():
        value = summary.get(field)
        minimum_ok = _is_number(value) and float(value) >= 0
        if field == "takt_time":
            minimum_ok = _is_number(value) and float(value) > 0
        check(label, minimum_ok, f"Nilai {field} tidak valid: {value!r}.")

    # 4. Allocation lists must agree with the physical rows parsed by reader.
    if not isinstance(operator_allocations, list):
        check("OPERATOR_LIST", False, "Field 'operator_allocations' harus berupa list.")
        operator_allocations = []
    else:
        check("OPERATOR_LIST", True, f"{len(operator_allocations)} alokasi operator.")

    if not isinstance(machine_allocations, list):
        check("MACHINE_LIST", False, "Field 'machine_allocations' harus berupa list.")
        machine_allocations = []
    else:
        check("MACHINE_LIST", True, f"{len(machine_allocations)} alokasi mesin.")

    expected_operator_rows = sorted(
        source_row
        for row in rows
        if isinstance(row, dict)
        and str(row.get("operator_name") or "").strip()
        and (source_row := _as_int(row.get("source_row"))) is not None
    )
    actual_operator_rows = sorted(
        _as_int(item.get("source_row"))
        for item in operator_allocations
        if isinstance(item, dict) and _as_int(item.get("source_row")) is not None
    )
    operator_records_well_formed = all(
        isinstance(item, dict)
        and _as_int(item.get("source_row")) is not None
        and bool(str(item.get("operator_name") or "").strip())
        for item in operator_allocations
    )
    check("OPERATOR_RECORDS", operator_records_well_formed, "Ada alokasi operator tanpa baris sumber/nama operator yang valid.")
    check(
        "OPERATOR_ROW_MATCH",
        expected_operator_rows == actual_operator_rows,
        f"Baris operator dari rows ({len(expected_operator_rows)}) tidak sama dengan daftar alokasi ({len(actual_operator_rows)}).",
    )
    check(
        "OPERATOR_SUMMARY_COUNT",
        _as_int(summary.get("operator_allocations")) == len(operator_allocations),
        f"Summary={summary.get('operator_allocations')!r}, daftar alokasi={len(operator_allocations)}.",
    )

    expected_machine_rows = sorted(
        source_row
        for row in rows
        if isinstance(row, dict)
        and str(row.get("machine_code") or "").strip()
        and (source_row := _as_int(row.get("source_row"))) is not None
    )
    actual_machine_rows = sorted(
        _as_int(item.get("source_row"))
        for item in machine_allocations
        if isinstance(item, dict) and _as_int(item.get("source_row")) is not None
    )
    machine_records_well_formed = all(
        isinstance(item, dict)
        and _as_int(item.get("source_row")) is not None
        and bool(str(item.get("machine_code") or "").strip())
        for item in machine_allocations
    )
    check("MACHINE_RECORDS", machine_records_well_formed, "Ada alokasi mesin tanpa baris sumber/kode mesin yang valid.")
    check(
        "MACHINE_ROW_MATCH",
        expected_machine_rows == actual_machine_rows,
        f"Baris mesin dari rows ({len(expected_machine_rows)}) tidak sama dengan daftar alokasi ({len(actual_machine_rows)}).",
    )
    check(
        "MACHINE_SUMMARY_COUNT",
        _as_int(summary.get("machine_allocations")) == len(machine_allocations),
        f"Summary={summary.get('machine_allocations')!r}, daftar alokasi={len(machine_allocations)}.",
    )

    # 5. Machine frequency map must match machine allocation list.
    if not isinstance(machine_counts, dict):
        check("MACHINE_COUNTS_STRUCTURE", False, "Field 'machine_counts' harus berupa dictionary.")
        machine_counts = {}
    else:
        check("MACHINE_COUNTS_STRUCTURE", True, f"{len(machine_counts)} tipe mesin unik.")

    counted_machines = Counter(
        str(item.get("machine_code") or "").strip()
        for item in machine_allocations
        if isinstance(item, dict) and str(item.get("machine_code") or "").strip()
    )
    normalized_machine_counts = {
        str(key).strip(): _as_int(value) for key, value in machine_counts.items()
    }
    machine_frequency_ok = (
        all(value is not None and value >= 0 for value in normalized_machine_counts.values())
        and dict(counted_machines) == normalized_machine_counts
    )
    check(
        "MACHINE_FREQUENCY_MATCH",
        machine_frequency_ok,
        "Frekuensi tipe mesin tidak sama dengan rincian alokasi mesin." if not machine_frequency_ok else f"{len(counted_machines)} tipe mesin cocok.",
    )

    # 6. Process summary/count consistency.
    if not isinstance(process_summary, list):
        check("PROCESS_SUMMARY_STRUCTURE", False, "Field 'process_summary' harus berupa list.")
        process_summary = []
    else:
        check("PROCESS_SUMMARY_STRUCTURE", True, f"{len(process_summary)} proses unik/terkelompok.")

    named_process_rows = [
        row for row in rows
        if isinstance(row, dict)
        and bool(str(row.get("process_name") or "").strip())
        and not bool(row.get("is_continuation_row"))
    ]
    check(
        "PROCESS_SUMMARY_COUNT",
        _as_int(summary.get("process_count")) == len(named_process_rows),
        f"Summary={summary.get('process_count')!r}, baris proses bernama={len(named_process_rows)}.",
    )

    process_keys: List[str] = []
    process_summary_well_formed = True
    for item in process_summary:
        if not isinstance(item, dict):
            process_summary_well_formed = False
            continue
        key = str(item.get("process_key") or "").strip()
        name = str(item.get("process_name") or "").strip()
        if not key or not name:
            process_summary_well_formed = False
        process_keys.append(key)
    check("PROCESS_SUMMARY_RECORDS", process_summary_well_formed, "Ada proses summary tanpa nama atau process_key.")
    check(
        "PROCESS_KEY_UNIQUE",
        len(process_keys) == len(set(process_keys)),
        "Ada process_key duplikat di process_summary." if len(process_keys) != len(set(process_keys)) else "",
    )

    # Verify per-process allocation totals against the raw row records.
    process_allocations_ok = True
    rows_by_source = _record_rows_by_source(rows)
    for item in process_summary:
        if not isinstance(item, dict):
            continue
        process_key = str(item.get("process_key") or "")
        expected_ops = sum(
            1 for row in rows
            if isinstance(row, dict)
            and row.get("process_key") == process_key
            and str(row.get("operator_name") or "").strip()
        )
        expected_machines = sum(
            1 for row in rows
            if isinstance(row, dict)
            and row.get("process_key") == process_key
            and str(row.get("machine_code") or "").strip()
        )
        if _as_int(item.get("operator_allocations")) != expected_ops:
            process_allocations_ok = False
        if _as_int(item.get("machine_allocations")) != expected_machines:
            process_allocations_ok = False
    check(
        "PROCESS_ALLOCATION_TOTALS",
        process_allocations_ok,
        "Jumlah alokasi mesin/operator dalam process_summary tidak cocok dengan rows.",
    )

    # 7. Continuation rows and inherited process references.
    continuation_rows = result.get("continuation_operator_rows")
    if not isinstance(continuation_rows, list):
        check("CONTINUATION_LIST", False, "Field 'continuation_operator_rows' harus berupa list.")
        continuation_rows = []
    else:
        check("CONTINUATION_LIST", True, f"{len(continuation_rows)} baris lanjutan operator.")

    expected_continuation_rows = sorted(
        _as_int(item.get("source_row"))
        for item in operator_allocations
        if isinstance(item, dict)
        and _as_int(item.get("source_row")) is not None
        and bool(rows_by_source.get(_as_int(item.get("source_row")), {}).get("is_continuation_row"))
    )
    actual_continuation_rows = sorted(
        _as_int(value) for value in continuation_rows if _as_int(value) is not None
    )
    check(
        "CONTINUATION_ROW_MATCH",
        expected_continuation_rows == actual_continuation_rows,
        f"Baris lanjutan yang terdeteksi ({actual_continuation_rows}) tidak cocok dengan rows ({expected_continuation_rows}).",
    )

    # 8. Preserve warnings generated by the reader; do not silently ignore them.
    reader_warnings = result.get("warnings")
    if isinstance(reader_warnings, list):
        warnings.extend(str(item) for item in reader_warnings if str(item).strip())
    else:
        warnings.append("Field warnings dari reader tidak berbentuk list.")

    # Warn, rather than fail, if an allocation has no inherited process label.
    for allocation in operator_allocations:
        if isinstance(allocation, dict) and not str(allocation.get("process_name") or "").strip():
            warnings.append(
                f"Baris {allocation.get('source_row')}: alokasi operator belum memiliki nama proses induk."
            )
    for allocation in machine_allocations:
        if isinstance(allocation, dict) and not str(allocation.get("process_name") or "").strip():
            warnings.append(
                f"Baris {allocation.get('source_row')}: alokasi mesin belum memiliki nama proses induk."
            )

    valid = len(errors) == 0
    status = "FAIL" if not valid else ("PASS WITH WARNING" if warnings else "PASS")
    return {
        "valid": valid,
        "status": status,
        "source_file": source_file,
        "line_label": line_label,
        "checks": checks,
        "errors": errors,
        "warnings": warnings,
    }


def validate_many(results: Sequence[Dict[str, Any]]) -> Dict[str, Any]:
    """Validate all reader results and ensure each line number is unique."""
    reports = [validate_result(result) for result in results]
    errors: List[str] = []
    warnings: List[str] = []

    if len(results) < 2:
        errors.append("Minimal dua file sumber diperlukan untuk komparasi.")

    line_numbers: List[Any] = [item.get("line_number") for item in results]
    valid_line_numbers = [value for value in line_numbers if _is_number(value)]
    duplicates = [value for value, count in Counter(valid_line_numbers).items() if count > 1]
    if duplicates:
        errors.append(f"Nomor line duplikat: {duplicates}.")

    for report in reports:
        errors.extend(f"{report['line_label']} — {message}" for message in report["errors"])
        warnings.extend(f"{report['line_label']} — {message}" for message in report["warnings"])

    valid = not errors
    status = "FAIL" if not valid else ("PASS WITH WARNING" if warnings else "PASS")
    return {
        "valid": valid,
        "status": status,
        "file_count": len(results),
        "reports": reports,
        "errors": errors,
        "warnings": warnings,
    }


def format_validation_report(results: Sequence[Dict[str, Any]], report: Dict[str, Any]) -> str:
    """Format per-file checks and overall validation status for CMD."""
    output: List[str] = []
    output.append("=" * 72)
    output.append("COMPARE_CT — TAHAP 2: VALIDASI SUMBER LINE BALANCING")
    output.append("=" * 72)

    for index, (result, file_report) in enumerate(zip(results, report["reports"]), start=1):
        if index > 1:
            output.append("")
            output.append("-" * 72)
        output.append(f"FILE   : {Path(file_report['source_file']).name}")
        output.append(f"LINE   : {file_report['line_label']}")
        output.append(f"STATUS : {file_report['status']}")
        for item in file_report["checks"]:
            mark = "PASS" if item["passed"] else "FAIL"
            detail = f" — {item['detail']}" if item.get("detail") else ""
            output.append(f"  [{mark}] {item['name']}{detail}")
        for warning in file_report["warnings"]:
            output.append(f"  [WARN] {warning}")
        for error in file_report["errors"]:
            output.append(f"  [FAIL] {error}")

    output.append("")
    output.append("=" * 72)
    output.append(f"HASIL AKHIR : {report['status']}")
    output.append(f"JUMLAH FILE : {report['file_count']}")
    output.append(f"ERROR       : {len(report['errors'])}")
    output.append(f"PERINGATAN  : {len(report['warnings'])}")
    if report["errors"]:
        output.append("DETAIL ERROR:")
        output.extend(f"  - {message}" for message in report["errors"])
    output.append("=" * 72)
    return "\n".join(output)


def run_validation(file_paths: Sequence[str | Path]) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """Read source files through reader.py, then validate the parsed results."""
    results = read_many(file_paths)
    report = validate_many(results)
    return results, report


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Membaca dan memvalidasi 2 atau lebih file sumber Line Balancing."
    )
    parser.add_argument("files", nargs="+", help="File Excel sumber (minimal 2)")
    args = parser.parse_args(argv)

    try:
        results, report = run_validation(args.files)
        print(format_validation_report(results, report))
        return 0 if report["valid"] else 2
    except Exception as exc:
        print(f"[GAGAL] Validasi tidak dapat dijalankan: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
