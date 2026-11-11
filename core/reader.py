"""Read only the 'Line Balancing' worksheet from Line Balancing workbooks.

This module deliberately ignores all other worksheets. It is the first-stage
reader for AUTO LINE BALANCING COMPARISON.

Windows runtime requirements:
    Python 3.10+
    Microsoft Excel installed
    pywin32 (py -m pip install pywin32)

Example:
    from core.reader import read_line_balancing
    result = read_line_balancing(r"D:\\INPUT\\LINE_37.xlsx")
    print(result["summary"])
"""

from __future__ import annotations

import re
import unicodedata
from collections import Counter, OrderedDict
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple


SHEET_NAME = "Line Balancing"
DATA_START_ROW = 12
TOTAL_MARKER = "TOTAL"

# Excel column positions (1-based) on the source worksheet.
COL_GROUP = 2          # B: group / section
COL_PROCESS_NO = 4     # D: process number
COL_PROCESS_NAME = 5   # E: process name
COL_MACHINE = 8        # H: machine code
COL_OPERATOR = 9       # I: operator allocation
COL_CT = 10            # J: cycle time
COL_AVERAGE_CT = 11    # K: average cycle time
COL_PROCESS_POT = 12  # L: process potential
COL_OPERATOR_POT = 13 # M: operator potential
COL_REMARKS = 15       # O: remarks


class LineBalancingReadError(RuntimeError):
    """Raised when a source workbook cannot be read safely."""


def clean_text(value: Any) -> str:
    """Clean Excel text artifacts while preserving the original wording."""
    if value is None:
        return ""
    text = str(value)
    text = text.replace("_x000D_", " ")
    text = text.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_key(value: Any) -> str:
    """Create a conservative matching key; semantic matching is a later stage."""
    text = clean_text(value)
    if not text:
        return ""
    text = unicodedata.normalize("NFKC", text).upper()
    # Normalize typographic variants and spacing, but do not remove meaningful
    # words or automatically equate semantically different process names.
    text = text.replace("–", "-").replace("—", "-").replace("×", "X")
    text = re.sub(r"\s*([+/&-])\s*", r"\1", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _is_empty(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _to_number(value: Any) -> Optional[float | int]:
    """Return a numeric Excel value as int/float, otherwise None."""
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float)):
        if isinstance(value, float) and value.is_integer():
            return int(value)
        return value
    if isinstance(value, str):
        text = value.strip().replace(",", ".")
        try:
            number = float(text)
            return int(number) if number.is_integer() else number
        except ValueError:
            return None
    return None


def _cell(matrix: Sequence[Sequence[Any]], row_1based: int, col_1based: int) -> Any:
    """Get a 1-based cell from a matrix whose top-left is A1."""
    try:
        return matrix[row_1based - 1][col_1based - 1]
    except (IndexError, TypeError):
        return None


def _find_total_row(matrix: Sequence[Sequence[Any]]) -> int:
    """Find the first exact TOTAL marker in column B from row 12 onward."""
    for row_number in range(DATA_START_ROW, len(matrix) + 1):
        marker = clean_text(_cell(matrix, row_number, COL_GROUP)).upper()
        if marker == TOTAL_MARKER:
            return row_number
    raise LineBalancingReadError(
        "Baris TOTAL tidak ditemukan di kolom B mulai baris "
        f"{DATA_START_ROW}. File tidak diproses agar data tidak terpotong."
    )


def parse_line_balancing_matrix(matrix: Sequence[Sequence[Any]], source_path: str = "") -> Dict[str, Any]:
    """Parse an A1-based worksheet matrix. Useful for deterministic testing.

    Only rows from 12 through the row before TOTAL are interpreted as source
    data. Metadata is read from F4, F5, H4:H6, and F7.
    """
    total_row = _find_total_row(matrix)

    line_number = _to_number(_cell(matrix, 4, 6))
    style = clean_text(_cell(matrix, 5, 6))
    target_per_hour = _to_number(_cell(matrix, 4, 8))
    potential_per_hour = _to_number(_cell(matrix, 5, 8))
    actual_per_hour = _to_number(_cell(matrix, 6, 8))
    takt_time = _to_number(_cell(matrix, 7, 6))

    raw_rows: List[Dict[str, Any]] = []
    operator_allocations: List[Dict[str, Any]] = []
    machine_allocations: List[Dict[str, Any]] = []
    process_summary: "OrderedDict[str, Dict[str, Any]]" = OrderedDict()
    machine_counts: Counter[str] = Counter()
    warnings: List[str] = []

    current_group = ""
    current_process_no: Any = None
    current_process_name = ""

    for row_number in range(DATA_START_ROW, total_row):
        raw_group = clean_text(_cell(matrix, row_number, COL_GROUP))
        raw_process_no = _cell(matrix, row_number, COL_PROCESS_NO)
        raw_process_name = clean_text(_cell(matrix, row_number, COL_PROCESS_NAME))
        machine_code = clean_text(_cell(matrix, row_number, COL_MACHINE))
        operator_name = clean_text(_cell(matrix, row_number, COL_OPERATOR))
        ct = _to_number(_cell(matrix, row_number, COL_CT))
        average_ct = _to_number(_cell(matrix, row_number, COL_AVERAGE_CT))
        process_potential = _to_number(_cell(matrix, row_number, COL_PROCESS_POT))
        operator_potential = _to_number(_cell(matrix, row_number, COL_OPERATOR_POT))
        remarks = clean_text(_cell(matrix, row_number, COL_REMARKS))

        # A group label applies to subsequent rows until another label appears.
        if raw_group:
            current_group = raw_group

        # A blank process cell on a secondary allocation row means continuation
        # of the previous process, not a new/unnamed process.
        if not _is_empty(raw_process_no):
            current_process_no = raw_process_no
        if raw_process_name:
            current_process_name = raw_process_name

        if not any((raw_group, not _is_empty(raw_process_no), raw_process_name,
                    machine_code, operator_name, ct is not None,
                    average_ct is not None, process_potential is not None,
                    operator_potential is not None, remarks)):
            continue

        process_key = normalize_key(current_process_name)
        record = {
            "source_row": row_number,
            "group": current_group,
            "process_no": current_process_no,
            "process_name": current_process_name,
            "process_key": process_key,
            "machine_code": machine_code,
            "machine_key": normalize_key(machine_code),
            "operator_name": operator_name,
            "operator_key": normalize_key(operator_name),
            "ct": ct,
            "average_ct": average_ct,
            "process_potential": process_potential,
            "operator_potential": operator_potential,
            "remarks": remarks,
            "is_continuation_row": not bool(raw_process_name),
        }
        raw_rows.append(record)

        # Keep every named process in the result, including processes with zero
        # operator allocation, so the comparison stage can show a real zero.
        if raw_process_name:
            process_summary.setdefault(process_key, {
                "process_name": current_process_name,
                "process_key": process_key,
                "groups": [],
                "process_numbers": [],
                "operator_allocations": 0,
                "machine_allocations": 0,
            })
            summary_item = process_summary[process_key]
            if current_group and current_group not in summary_item["groups"]:
                summary_item["groups"].append(current_group)
            if not _is_empty(current_process_no) and current_process_no not in summary_item["process_numbers"]:
                summary_item["process_numbers"].append(current_process_no)

        if machine_code:
            machine_counts[machine_code] += 1
            machine_allocations.append({
                "source_row": row_number,
                "group": current_group,
                "process_no": current_process_no,
                "process_name": current_process_name,
                "process_key": process_key,
                "machine_code": machine_code,
            })
            if process_key in process_summary:
                process_summary[process_key]["machine_allocations"] += 1
            elif not process_key:
                warnings.append(f"Baris {row_number}: mesin terisi tetapi nama proses belum terbaca.")

        # KPI rule locked by the user: count each non-empty operator cell as one
        # allocation, even when the same person appears on another process.
        if operator_name:
            if not process_key:
                warnings.append(f"Baris {row_number}: operator '{operator_name}' tidak memiliki proses induk.")
            operator_allocations.append({
                "source_row": row_number,
                "group": current_group,
                "process_no": current_process_no,
                "process_name": current_process_name,
                "process_key": process_key,
                "machine_code": machine_code,
                "operator_name": operator_name,
            })
            if process_key in process_summary:
                process_summary[process_key]["operator_allocations"] += 1

    # Count rows with a named process separately from physical allocation rows.
    process_name_row_count = sum(1 for row in raw_rows if row["process_name"] and not row["is_continuation_row"])
    continuation_operator_rows = [
        item["source_row"] for item in operator_allocations
        if item["source_row"] in {
            row["source_row"] for row in raw_rows
            if not row["process_name"] or row["is_continuation_row"]
        }
    ]

    line_label = f"LINE {line_number}" if line_number is not None else "LINE TIDAK DIKETAHUI"
    if style:
        line_label += f" ({style})"

    return {
        "source_file": str(source_path),
        "sheet_name": SHEET_NAME,
        "line_number": line_number,
        "style": style,
        "line_label": line_label,
        "summary": {
            "target_per_hour": target_per_hour,
            "potential_per_hour": potential_per_hour,
            "actual_per_hour": actual_per_hour,
            "takt_time": takt_time,
            "operator_allocations": len(operator_allocations),
            "machine_allocations": len(machine_allocations),
            "process_count": process_name_row_count,
        },
        "total_row": total_row,
        "data_start_row": DATA_START_ROW,
        "data_end_row": total_row - 1,
        "rows": raw_rows,
        "operator_allocations": operator_allocations,
        "machine_allocations": machine_allocations,
        "machine_counts": dict(machine_counts),
        "process_summary": list(process_summary.values()),
        "continuation_operator_rows": continuation_operator_rows,
        "warnings": warnings,
    }


def _as_matrix(value: Any) -> List[List[Any]]:
    """Normalize Excel COM Range.Value2 to a 2D Python list."""
    if isinstance(value, tuple):
        if value and isinstance(value[0], tuple):
            return [list(row) for row in value]
        return [list(value)]
    return [[value]]


def read_line_balancing(file_path: str | Path) -> Dict[str, Any]:
    """Read one workbook through Excel COM, using only 'Line Balancing'.

    The workbook is opened read-only and never saved. All other worksheets are
    ignored. Requires Microsoft Excel and pywin32 on Windows.
    """
    path = Path(file_path).expanduser().resolve()
    if not path.is_file():
        raise FileNotFoundError(f"File sumber tidak ditemukan: {path}")
    if path.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
        raise LineBalancingReadError(f"Format file tidak didukung: {path.suffix}")

    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise LineBalancingReadError(
            "Modul pywin32 belum terpasang. Jalankan: py -m pip install pywin32"
        ) from exc

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False

        workbook = excel.Workbooks.Open(
            str(path),
            UpdateLinks=0,
            ReadOnly=True,
            AddToMru=False,
            Notify=False,
        )

        try:
            worksheet = workbook.Worksheets(SHEET_NAME)
        except Exception as exc:
            raise LineBalancingReadError(
                f"Sheet '{SHEET_NAME}' tidak ditemukan pada file: {path.name}"
            ) from exc

        used_range = worksheet.UsedRange
        last_used_row = int(used_range.Row + used_range.Rows.Count - 1)
        if last_used_row < DATA_START_ROW:
            raise LineBalancingReadError(
                f"Sheet '{SHEET_NAME}' tidak memiliki data mulai baris {DATA_START_ROW}."
            )

        # Read A1:O(last used row) in one COM call to reduce overhead.
        matrix = _as_matrix(worksheet.Range(f"A1:O{last_used_row}").Value2)
        result = parse_line_balancing_matrix(matrix, str(path))
        result["workbook_sheets_ignored"] = True
        return result

    except LineBalancingReadError:
        raise
    except Exception as exc:
        raise LineBalancingReadError(
            f"Gagal membaca '{path.name}' pada sheet '{SHEET_NAME}': {exc}"
        ) from exc
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


def read_many(file_paths: Sequence[str | Path]) -> List[Dict[str, Any]]:
    """Read two or more source workbooks; stop on the first invalid workbook."""
    if len(file_paths) < 2:
        raise ValueError("Minimal pilih 2 file Line Balancing untuk komparasi.")
    results: List[Dict[str, Any]] = []
    seen_lines = set()
    for path in file_paths:
        item = read_line_balancing(path)
        line_number = item.get("line_number")
        if line_number in seen_lines:
            raise LineBalancingReadError(
                f"Nomor line {line_number} terdeteksi lebih dari satu kali. "
                "Periksa file yang dipilih."
            )
        seen_lines.add(line_number)
        results.append(item)
    return results


def format_audit(result: Dict[str, Any]) -> str:
    """Create a concise validation report for console/log output."""
    summary = result["summary"]
    lines = [
        f"FILE              : {Path(result['source_file']).name}",
        f"SHEET             : {result['sheet_name']}",
        f"LINE              : {result['line_label']}",
        f"BARIS DATA        : {result['data_start_row']} - {result['data_end_row']}",
        f"BARIS TOTAL       : {result['total_row']}",
        f"JUMLAH PROSES     : {summary['process_count']}",
        f"ALOKASI MESIN     : {summary['machine_allocations']}",
        f"ALOKASI OPERATOR  : {summary['operator_allocations']}",
        f"TARGET/JAM        : {summary['target_per_hour']}",
        f"POTENSIAL/JAM     : {summary['potential_per_hour']}",
        f"AKTUAL/JAM        : {summary['actual_per_hour']}",
        f"TAKT TIME         : {summary['takt_time']}",
        f"JENIS MESIN       : {len(result['machine_counts'])}",
        f"PROSES UNIK       : {len(result['process_summary'])}",
        f"BARIS LANJUTAN OPT: {len(result['continuation_operator_rows'])}",
        f"PERINGATAN        : {len(result['warnings'])}",
    ]
    for warning in result["warnings"]:
        lines.append(f"  - {warning}")
    return "\n".join(lines)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="Validasi pembacaan sheet Line Balancing saja."
    )
    parser.add_argument("files", nargs="+", help="File Excel sumber (minimal 2)")
    args = parser.parse_args()

    try:
        results = read_many(args.files)
        for index, result in enumerate(results, start=1):
            if index > 1:
                print("\n" + "=" * 64 + "\n")
            print(format_audit(result))
    except Exception as exc:
        raise SystemExit(f"[GAGAL] {exc}")
