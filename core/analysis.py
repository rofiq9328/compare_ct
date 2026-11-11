"""Analyze a dynamic comparison dataset produced by core.comparison.

This module is intentionally conservative:
- it compares only rows already matched by exact normalization or approved aliases;
- it never merges REVIEW_REQUIRED processes;
- it reports allocation differences as observations, not as proven bottlenecks;
- cycle-time bottleneck analysis is marked unavailable because the current
  comparison_data.json schema does not yet carry process-level cycle-time data.

Run from the project root:
    py -m core.analysis "logs\\comparison_data.json"

Optional:
    py -m core.analysis INPUT --output logs/analysis_data.json
"""

from __future__ import annotations

import argparse
import json
import math
import os
import sys
from pathlib import Path
from typing import Any, Dict, Mapping, Sequence

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUT_PATH = PROJECT_ROOT / "logs" / "comparison_data.json"
DEFAULT_OUTPUT_PATH = PROJECT_ROOT / "logs" / "analysis_data.json"


class AnalysisError(ValueError):
    """Raised when comparison data cannot be analyzed safely."""


def _read_json_object(path: str | Path) -> Dict[str, Any]:
    file_path = Path(path).expanduser()
    if not file_path.is_absolute():
        file_path = (Path.cwd() / file_path).resolve()
    if not file_path.is_file():
        raise FileNotFoundError(f"File comparison JSON tidak ditemukan: {file_path}")
    try:
        data = json.loads(file_path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        raise AnalysisError(f"JSON tidak valid pada {file_path}: {exc}") from exc
    if not isinstance(data, dict):
        raise AnalysisError("Isi comparison_data.json harus berupa JSON object.")
    return data


def _is_number(value: Any) -> bool:
    return (
        isinstance(value, (int, float))
        and not isinstance(value, bool)
        and math.isfinite(float(value))
    )


def _as_number(value: Any, field: str, context: str) -> float | None:
    if value is None or value == "":
        return None
    if not _is_number(value):
        raise AnalysisError(f"Nilai {field} pada {context} bukan angka valid: {value!r}.")
    return float(value)


def _as_non_negative_int(value: Any, field: str, context: str) -> int:
    if value is None:
        return 0
    if not _is_number(value) or float(value) < 0 or not float(value).is_integer():
        raise AnalysisError(
            f"Nilai {field} pada {context} harus bilangan bulat >= 0; diterima {value!r}."
        )
    return int(value)


def _round(value: float | None, digits: int = 2) -> float | None:
    if value is None:
        return None
    result = round(float(value), digits)
    return 0.0 if result == 0 else result


def _percent(numerator: float | None, denominator: float | None) -> float | None:
    if numerator is None or denominator is None or denominator == 0:
        return None
    return numerator / denominator * 100.0


def _validate_input(data: Mapping[str, Any]) -> list[str]:
    if data.get("schema_version") != 1:
        raise AnalysisError(
            f"schema_version tidak didukung: {data.get('schema_version')!r}; diperlukan 1."
        )
    if data.get("comparison_type") != "dynamic_multi_line":
        raise AnalysisError("comparison_type bukan 'dynamic_multi_line'.")

    line_columns = data.get("line_columns")
    if not isinstance(line_columns, list) or len(line_columns) < 2:
        raise AnalysisError("line_columns harus berupa list dengan minimal 2 line.")
    if data.get("line_count") != len(line_columns):
        raise AnalysisError(
            f"line_count={data.get('line_count')!r}, tetapi line_columns berisi {len(line_columns)} line."
        )

    labels: list[str] = []
    for index, item in enumerate(line_columns, start=1):
        if not isinstance(item, dict):
            raise AnalysisError(f"line_columns ke-{index} bukan object.")
        label = str(item.get("line_label") or "").strip()
        if not label:
            raise AnalysisError(f"line_label pada line ke-{index} kosong.")
        if label in labels:
            raise AnalysisError(f"line_label duplikat: {label}.")
        labels.append(label)

    quality = data.get("quality")
    if not isinstance(quality, dict) or quality.get("valid") is not True:
        raise AnalysisError("Quality gate comparison tidak PASS; analisis dihentikan.")
    if quality.get("errors"):
        raise AnalysisError("Comparison masih memiliki quality.errors; analisis dihentikan.")
    if quality.get("operator_count_basis") != (
        "alokasi operator berdasarkan jumlah sel nama operator terisi"
    ):
        raise AnalysisError("Basis hitung operator tidak sesuai aturan project.")

    summary_rows = data.get("summary_rows")
    machine_rows = data.get("machine_rows")
    process_rows = data.get("process_rows")
    if not isinstance(summary_rows, list) or not summary_rows:
        raise AnalysisError("summary_rows tidak tersedia atau kosong.")
    if not isinstance(machine_rows, list):
        raise AnalysisError("machine_rows harus berupa list.")
    if not isinstance(process_rows, list):
        raise AnalysisError("process_rows harus berupa list.")

    for row_index, row in enumerate(summary_rows, start=1):
        if not isinstance(row, dict):
            raise AnalysisError(f"summary_rows ke-{row_index} bukan object.")
        values = row.get("values_by_line")
        if not isinstance(values, dict):
            raise AnalysisError(f"summary_rows ke-{row_index} tidak memiliki values_by_line.")
        for label in labels:
            if label not in values:
                raise AnalysisError(f"SUMMARY {row.get('label')!r} tidak memiliki nilai untuk {label}.")
            _as_number(values[label], row.get("key", row.get("label", "summary")), label)

    for collection_name, collection in (("machine_rows", machine_rows), ("process_rows", process_rows)):
        for row_index, row in enumerate(collection, start=1):
            if not isinstance(row, dict):
                raise AnalysisError(f"{collection_name} ke-{row_index} bukan object.")
            map_fields = (
                ("values_by_line",)
                if collection_name == "machine_rows"
                else ("operator_allocations_by_line", "machine_allocations_by_line")
            )
            for field in map_fields:
                values = row.get(field)
                if not isinstance(values, dict):
                    raise AnalysisError(f"{collection_name} ke-{row_index} tidak memiliki {field}.")
                for label in labels:
                    if label not in values:
                        raise AnalysisError(
                            f"{collection_name} ke-{row_index} {field} tidak memiliki {label}."
                        )
                    _as_non_negative_int(values[label], field, label)
            present = row.get("present_in_lines", [])
            if not isinstance(present, list) or any(label not in labels for label in present):
                raise AnalysisError(f"present_in_lines tidak valid pada {collection_name} ke-{row_index}.")

    return labels


def _summary_map(summary_rows: Sequence[Mapping[str, Any]]) -> Dict[str, Mapping[str, Any]]:
    result: Dict[str, Mapping[str, Any]] = {}
    for row in summary_rows:
        key = str(row.get("key") or "").strip()
        if not key:
            raise AnalysisError("Terdapat SUMMARY row tanpa key.")
        if key in result:
            raise AnalysisError(f"SUMMARY key duplikat: {key}.")
        result[key] = row
    return result


def _line_kpis(
    summary_by_key: Mapping[str, Mapping[str, Any]], labels: Sequence[str]
) -> Dict[str, Dict[str, Any]]:
    required = ("target_per_hour", "potential_per_hour", "actual_per_hour", "takt_time", "operator_allocations")
    for key in required:
        if key not in summary_by_key:
            raise AnalysisError(f"SUMMARY tidak memiliki KPI wajib: {key}.")

    output: Dict[str, Dict[str, Any]] = {}
    for label in labels:
        target = _as_number(summary_by_key["target_per_hour"]["values_by_line"][label], "target_per_hour", label)
        potential = _as_number(summary_by_key["potential_per_hour"]["values_by_line"][label], "potential_per_hour", label)
        actual = _as_number(summary_by_key["actual_per_hour"]["values_by_line"][label], "actual_per_hour", label)
        takt = _as_number(summary_by_key["takt_time"]["values_by_line"][label], "takt_time", label)
        operators = _as_non_negative_int(
            summary_by_key["operator_allocations"]["values_by_line"][label], "operator_allocations", label
        )

        target_gap = None if target is None or actual is None else target - actual
        potential_gap = None if target is None or potential is None else target - potential
        if target is None or actual is None:
            production_status = "DATA_TIDAK_LENGKAP"
        elif actual < target:
            production_status = "AKTUAL_DI_BAWAH_TARGET"
        else:
            production_status = "AKTUAL_MENCAPAI_ATAU_MELEBIHI_TARGET"

        output[label] = {
            "target_per_hour": _round(target, 4),
            "potential_per_hour": _round(potential, 4),
            "actual_per_hour": _round(actual, 4),
            "takt_time_seconds": _round(takt, 4),
            "operator_allocations": operators,
            "actual_vs_target_pct": _round(_percent(actual, target), 2),
            "actual_vs_potential_pct": _round(_percent(actual, potential), 2),
            "potential_vs_target_pct": _round(_percent(potential, target), 2),
            "target_gap_pcs_per_hour": _round(target_gap, 2),
            "potential_gap_to_target_pcs_per_hour": _round(potential_gap, 2),
            "production_status": production_status,
        }
    return output


def _summary_differences(
    summary_by_key: Mapping[str, Mapping[str, Any]], labels: Sequence[str]
) -> list[Dict[str, Any]]:
    reference = labels[0]
    differences: list[Dict[str, Any]] = []
    for key, row in summary_by_key.items():
        values = row["values_by_line"]
        ref_value = _as_number(values[reference], key, reference)
        per_line: Dict[str, Any] = {}
        for label in labels:
            value = _as_number(values[label], key, label)
            delta = None if value is None or ref_value is None else value - ref_value
            delta_pct = _percent(delta, abs(ref_value)) if delta is not None else None
            per_line[label] = {
                "value": _round(value, 4),
                "delta_vs_reference": _round(delta, 4),
                "delta_vs_reference_pct": _round(delta_pct, 2),
            }
        differences.append({
            "key": key,
            "label": row.get("label", key),
            "unit": row.get("unit", ""),
            "reference_line": reference,
            "values_by_line": per_line,
        })
    return differences


def _analyze_machine_rows(
    rows: Sequence[Mapping[str, Any]], labels: Sequence[str]
) -> Dict[str, Any]:
    differing: list[Dict[str, Any]] = []
    line_exclusive: list[Dict[str, Any]] = []
    for row in rows:
        values = {label: _as_non_negative_int(row["values_by_line"][label], "machine count", label) for label in labels}
        present = list(row.get("present_in_lines") or [])
        nonzero = [label for label in labels if values[label] > 0]
        observed_values = [values[label] for label in labels]
        spread = max(observed_values) - min(observed_values) if observed_values else 0
        item = {
            "machine": row.get("machine"),
            "canonical_key": row.get("canonical_key"),
            "values_by_line": values,
            "present_in_lines": present,
            "lines_with_allocation": nonzero,
            "allocation_spread": spread,
            "alias_applied": bool(row.get("alias_applied", False)),
        }
        if spread > 0:
            differing.append(item)
        if len(nonzero) == 1:
            line_exclusive.append(item)
    return {
        "machine_type_count": len(rows),
        "machine_types_with_allocation_difference": differing,
        "machine_types_allocated_in_one_line_only": line_exclusive,
    }


def _analyze_process_rows(
    rows: Sequence[Mapping[str, Any]], labels: Sequence[str]
) -> Dict[str, Any]:
    matched_allocation_differences: list[Dict[str, Any]] = []
    processes_only_in_lines: list[Dict[str, Any]] = []
    review_required: list[Dict[str, Any]] = []
    group_variations: list[Dict[str, Any]] = []

    for row in rows:
        status = str(row.get("matching_status") or "UNKNOWN")
        present = list(row.get("present_in_lines") or [])
        operator_values = {
            label: _as_non_negative_int(row["operator_allocations_by_line"][label], "operator allocation", label)
            for label in labels
        }
        machine_values = {
            label: _as_non_negative_int(row["machine_allocations_by_line"][label], "process machine allocation", label)
            for label in labels
        }
        if status == "REVIEW_REQUIRED":
            review_required.append({
                "process": row.get("process"),
                "source_names": list(row.get("source_names") or []),
                "present_in_lines": present,
                "review_candidates": list(row.get("review_candidates") or []),
                "note": "Belum dikonfirmasi; tidak dianalisis sebagai proses yang sama antar-line.",
            })
            continue

        if len(present) == 1:
            processes_only_in_lines.append({
                "process": row.get("process"),
                "canonical_key": row.get("canonical_key"),
                "present_in_line": present[0],
                "operator_allocations": operator_values[present[0]],
                "machine_allocations": machine_values[present[0]],
                "matching_status": status,
            })
        elif len(present) > 1:
            op_values_present = [operator_values[label] for label in present]
            mc_values_present = [machine_values[label] for label in present]
            op_spread = max(op_values_present) - min(op_values_present)
            mc_spread = max(mc_values_present) - min(mc_values_present)
            if op_spread > 0 or mc_spread > 0:
                matched_allocation_differences.append({
                    "process": row.get("process"),
                    "canonical_key": row.get("canonical_key"),
                    "matching_status": status,
                    "present_in_lines": present,
                    "operator_allocations_by_line": operator_values,
                    "operator_allocation_spread": op_spread,
                    "machine_allocations_by_line": machine_values,
                    "machine_allocation_spread": mc_spread,
                    "groups_by_line": row.get("groups_by_line") or {},
                    "note": "Perbedaan alokasi teramati; bukan bukti bottleneck atau penyebab produktivitas.",
                })

        groups_by_line = row.get("groups_by_line") or {}
        nonempty_groups = {
            label: list(groups_by_line.get(label) or [])
            for label in present
            if groups_by_line.get(label)
        }
        unique_group_sets = {tuple(value) for value in nonempty_groups.values()}
        if len(unique_group_sets) > 1 and len(nonempty_groups) > 1:
            group_variations.append({
                "process": row.get("process"),
                "groups_by_line": nonempty_groups,
                "note": "Group sumber berbeda; periksa apakah perbedaan hanya penamaan atau memang alur kerja berbeda.",
            })

    return {
        "process_row_count": len(rows),
        "matched_processes_with_allocation_difference": matched_allocation_differences,
        "processes_present_in_one_line_only": processes_only_in_lines,
        "processes_requiring_matching_review": review_required,
        "process_group_variations": group_variations,
    }


def build_analysis(comparison: Mapping[str, Any]) -> Dict[str, Any]:
    """Build conservative analysis and report-ready observations."""
    labels = _validate_input(comparison)
    summary_by_key = _summary_map(comparison["summary_rows"])
    kpis = _line_kpis(summary_by_key, labels)
    summary_differences = _summary_differences(summary_by_key, labels)
    machine_analysis = _analyze_machine_rows(comparison["machine_rows"], labels)
    process_analysis = _analyze_process_rows(comparison["process_rows"], labels)

    observations: list[Dict[str, Any]] = []
    for label in labels:
        kpi = kpis[label]
        if kpi["production_status"] == "AKTUAL_DI_BAWAH_TARGET":
            observations.append({
                "type": "ACTUAL_BELOW_TARGET",
                "severity": "INFO",
                "line": label,
                "message": (
                    f"Aktual {kpi['actual_per_hour']} pcs/jam berada {kpi['target_gap_pcs_per_hour']} "
                    f"pcs/jam di bawah target {kpi['target_per_hour']} pcs/jam."
                ),
            })
        if (
            kpi["potential_per_hour"] is not None
            and kpi["target_per_hour"] is not None
            and kpi["potential_per_hour"] < kpi["target_per_hour"]
        ):
            observations.append({
                "type": "POTENTIAL_BELOW_TARGET",
                "severity": "INFO",
                "line": label,
                "message": (
                    f"Potensial {kpi['potential_per_hour']} pcs/jam masih {kpi['potential_gap_to_target_pcs_per_hour']} "
                    f"pcs/jam di bawah target {kpi['target_per_hour']} pcs/jam."
                ),
            })

    for item in process_analysis["processes_present_in_one_line_only"]:
        observations.append({
            "type": "PROCESS_ONLY_IN_ONE_LINE",
            "severity": "REVIEW",
            "line": item["present_in_line"],
            "message": (
                f"Proses '{item['process']}' hanya tercatat pada {item['present_in_line']}. "
                "Pastikan memang tidak digunakan di line lain, bukan akibat perbedaan nama yang belum terpetakan."
            ),
        })

    for item in machine_analysis["machine_types_allocated_in_one_line_only"]:
        observations.append({
            "type": "MACHINE_ONLY_IN_ONE_LINE",
            "severity": "INFO",
            "line": item["lines_with_allocation"][0],
            "message": (
                f"Tipe mesin '{item['machine']}' hanya memiliki alokasi pada "
                f"{item['lines_with_allocation'][0]} (jumlah {item['values_by_line'][item['lines_with_allocation'][0]]})."
            ),
        })

    for item in process_analysis["matched_processes_with_allocation_difference"]:
        observations.append({
            "type": "PROCESS_ALLOCATION_DIFFERENCE",
            "severity": "REVIEW",
            "line": None,
            "message": (
                f"Proses '{item['process']}' memiliki rentang alokasi operator {item['operator_allocation_spread']} "
                f"dan rentang alokasi mesin {item['machine_allocation_spread']} antar-line yang telah cocok."
            ),
        })

    review_count = len(process_analysis["processes_requiring_matching_review"])
    candidate_count = int(
        (comparison.get("quality") or {}).get("pending_process_match_candidates", 0) or 0
    )
    if review_count or candidate_count:
        observations.append({
            "type": "UNCONFIRMED_PROCESS_MATCHES",
            "severity": "WARNING",
            "line": None,
            "message": (
                f"Masih ada {candidate_count} kandidat pencocokan dan {review_count} baris proses "
                "yang memerlukan review. Proses tersebut tidak diperlakukan sebagai proses yang sama."
            ),
        })

    quality_warnings = list((comparison.get("quality") or {}).get("warnings") or [])
    quality = {
        "valid": True,
        "source_comparison_quality_valid": True,
        "warnings_inherited_from_comparison": quality_warnings,
        "process_matching_review_count": review_count,
        "process_matching_candidate_count": candidate_count,
        "review_required_processes_excluded_from_semantic_comparison": True,
        "bottleneck_assessment": "NOT_ASSESSED_PROCESS_CYCLE_TIME_NOT_AVAILABLE_IN_COMPARISON_DATA",
        "limitations": [
            "Analisis bottleneck per proses belum dapat disimpulkan karena comparison_data.json saat ini belum memuat cycle time tiap proses.",
            "Perbedaan alokasi operator atau mesin adalah temuan deskriptif, bukan bukti penyebab bottleneck.",
            "Proses yang belum dikonfirmasi tidak dianggap sama meskipun memiliki nama yang mirip.",
        ],
    }

    return {
        "schema_version": 1,
        "analysis_type": "conservative_multi_line_analysis",
        "line_count": len(labels),
        "line_labels": labels,
        "reference_line_for_deltas": labels[0],
        "line_kpis": kpis,
        "summary_differences": summary_differences,
        "machine_analysis": machine_analysis,
        "process_analysis": process_analysis,
        "observations": observations,
        "quality": quality,
    }


def format_analysis_report(result: Mapping[str, Any], output_path: Path) -> str:
    lines = [
        "=" * 78,
        "COMPARE_CT — TAHAP 5: ANALYSIS ENGINE",
        "=" * 78,
        f"JUMLAH LINE          : {result['line_count']}",
        f"LINE REFERENSI DELTA : {result['reference_line_for_deltas']}",
        f"JENIS MESIN           : {result['machine_analysis']['machine_type_count']}",
        f"BARIS PROSES          : {result['process_analysis']['process_row_count']}",
        f"PROSES ALOKASI BEDA   : {len(result['process_analysis']['matched_processes_with_allocation_difference'])}",
        f"PROSES SATU LINE      : {len(result['process_analysis']['processes_present_in_one_line_only'])}",
        f"PROSES REVIEW         : {len(result['process_analysis']['processes_requiring_matching_review'])}",
        "",
        "KPI PER LINE:",
    ]
    for label in result["line_labels"]:
        item = result["line_kpis"][label]
        actual_target = "N/A" if item["actual_vs_target_pct"] is None else f"{item['actual_vs_target_pct']:.2f}%"
        actual_potential = "N/A" if item["actual_vs_potential_pct"] is None else f"{item['actual_vs_potential_pct']:.2f}%"
        lines.append(
            f"  {label}: TARGET {item['target_per_hour']} | POTENSIAL {item['potential_per_hour']} | "
            f"AKTUAL {item['actual_per_hour']} pcs/jam | AKTUAL/TARGET {actual_target} | "
            f"AKTUAL/POTENSIAL {actual_potential} | OPT {item['operator_allocations']}"
        )
    lines.extend(["", "OBSERVASI:"])
    if result["observations"]:
        lines.extend(f"  - [{item['severity']}] {item['message']}" for item in result["observations"])
    else:
        lines.append("  - Tidak ada observasi dari aturan analisis saat ini.")
    lines.extend([
        "",
        "BATASAN ANALISIS:",
    ])
    lines.extend(f"  - {item}" for item in result["quality"]["limitations"])
    lines.extend([
        "",
        f"OUTPUT JSON          : {output_path}",
        "STATUS               : PASS",
        "BOTTLENECK           : BELUM DINILAI (CT PER PROSES BELUM TERSEDIA)",
        "=" * 78,
    ])
    return "\n".join(lines)


def run_analysis(
    comparison_path: str | Path = DEFAULT_INPUT_PATH,
    output_path: str | Path = DEFAULT_OUTPUT_PATH,
) -> Dict[str, Any]:
    comparison_file = Path(comparison_path).expanduser()
    if not comparison_file.is_absolute():
        comparison_file = (Path.cwd() / comparison_file).resolve()
    comparison = _read_json_object(comparison_file)
    result = build_analysis(comparison)

    output_file = Path(output_path).expanduser()
    if not output_file.is_absolute():
        output_file = (Path.cwd() / output_file).resolve()
    output_file.parent.mkdir(parents=True, exist_ok=True)
    temp_file = output_file.with_name(output_file.name + ".tmp")
    temp_file.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    os.replace(temp_file, output_file)
    print(format_analysis_report(result, output_file))
    return result


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Analisis konservatif dari logs/comparison_data.json; tidak menyimpulkan bottleneck tanpa CT proses."
    )
    parser.add_argument("comparison", nargs="?", default=str(DEFAULT_INPUT_PATH), help="Input comparison_data.json")
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT_PATH), help="Output analysis_data.json")
    args = parser.parse_args(argv)

    try:
        run_analysis(args.comparison, args.output)
        return 0
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        print(f"[GAGAL] Analysis engine dihentikan: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
