"""Build a dynamic, multi-line comparison dataset from normalization_preview.json.

This module does not read or modify Excel workbooks. It consumes the validated
output of core.normalizer and prepares a stable, line-dynamic structure for the
Excel report and chart generators in later stages.

Run from the project root:
    py -m core.comparison "logs\\normalization_preview.json"

Optional:
    py -m core.comparison PREVIEW --mapping config/mapping.json \
        --output logs/comparison_data.json
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import sys
from collections import OrderedDict
from pathlib import Path
from typing import Any, Dict, Mapping, Sequence, Tuple

from core.normalizer import DEFAULT_MAPPING_PATH, DEFAULT_OUTPUT_PATH, load_mapping, normalize_process_name


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_COMPARISON_OUTPUT = PROJECT_ROOT / "logs" / "comparison_data.json"
SUMMARY_FIELDS: Tuple[Tuple[str, str, str], ...] = (
    ("operator_allocations", "Jumlah Operator", "OPT"),
    ("takt_time", "Takt Time", "detik"),
    ("target_per_hour", "Target Output / Jam", "pcs"),
    ("potential_per_hour", "Potensial Output / Jam", "pcs"),
    ("actual_per_hour", "Aktual Output / Jam", "pcs"),
)


class ComparisonError(ValueError):
    """Raised when a normalized preview cannot be compared safely."""


def _read_json_object(path: str | Path) -> Dict[str, Any]:
    file_path = Path(path).expanduser()
    if not file_path.is_absolute():
        file_path = (Path.cwd() / file_path).resolve()
    if not file_path.is_file():
        raise FileNotFoundError(f"File JSON tidak ditemukan: {file_path}")
    try:
        data = json.loads(file_path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        raise ComparisonError(f"JSON tidak valid pada {file_path}: {exc}") from exc
    if not isinstance(data, dict):
        raise ComparisonError(f"Isi JSON harus berupa object: {file_path}")
    return data


def _is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(float(value))


def _number_or_none(value: Any, field: str, line_label: str) -> int | float | None:
    if value is None or value == "":
        return None
    if not _is_number(value):
        raise ComparisonError(
            f"Nilai {field!r} pada {line_label} bukan angka yang valid: {value!r}."
        )
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def _non_negative_count(value: Any, field: str, line_label: str) -> int:
    if value is None:
        return 0
    if not _is_number(value) or float(value) < 0:
        raise ComparisonError(
            f"Jumlah {field!r} pada {line_label} harus berupa angka >= 0; diterima: {value!r}."
        )
    if not float(value).is_integer():
        raise ComparisonError(
            f"Jumlah {field!r} pada {line_label} harus bilangan bulat; diterima: {value!r}."
        )
    return int(value)


def _pair_id(normalized_a: str, normalized_b: str) -> str:
    payload = json.dumps(
        sorted((normalized_a, normalized_b)),
        ensure_ascii=False,
        separators=(",", ":"),
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def validate_preview_structure(data: Mapping[str, Any]) -> list[str]:
    """Validate minimum structure and return the ordered line labels."""
    if data.get("schema_version") != 1:
        raise ComparisonError(
            f"Versi preview tidak didukung: {data.get('schema_version')!r}. Diperlukan schema_version=1."
        )
    raw_lines = data.get("lines")
    if not isinstance(raw_lines, list) or len(raw_lines) < 2:
        raise ComparisonError("Preview harus berisi minimal 2 line pada properti 'lines'.")
    if data.get("line_count") != len(raw_lines):
        raise ComparisonError(
            f"line_count={data.get('line_count')!r}, tetapi jumlah item 'lines'={len(raw_lines)}."
        )

    labels: list[str] = []
    line_numbers: set[int] = set()
    for index, line in enumerate(raw_lines, start=1):
        if not isinstance(line, dict):
            raise ComparisonError(f"Item lines ke-{index} bukan object.")
        label = str(line.get("line_label") or "").strip()
        if not label:
            raise ComparisonError(f"Nama line pada item ke-{index} kosong.")
        if label in labels:
            raise ComparisonError(f"Label line duplikat ditemukan: {label}.")
        labels.append(label)

        line_number = line.get("line_number")
        if line_number is not None:
            if not _is_number(line_number) or not float(line_number).is_integer():
                raise ComparisonError(f"Nomor line tidak valid pada {label}: {line_number!r}.")
            if int(line_number) in line_numbers:
                raise ComparisonError(f"Nomor line {int(line_number)} muncul lebih dari satu kali.")
            line_numbers.add(int(line_number))

        summary = line.get("summary")
        if not isinstance(summary, dict):
            raise ComparisonError(f"Ringkasan SUMMARY tidak ditemukan pada {label}.")
        for field, _caption, _unit in SUMMARY_FIELDS:
            _number_or_none(summary.get(field), field, label)
        for required_count in ("operator_allocations", "machine_allocations"):
            count = summary.get(required_count)
            if count is None or _non_negative_count(count, required_count, label) < 0:
                raise ComparisonError(f"Ringkasan {required_count} tidak valid pada {label}.")

    for catalog_name in ("process_catalog", "machine_catalog"):
        catalog = data.get(catalog_name)
        if not isinstance(catalog, list):
            raise ComparisonError(f"Preview tidak memiliki katalog '{catalog_name}' berbentuk list.")
        seen_keys: set[str] = set()
        for index, item in enumerate(catalog, start=1):
            if not isinstance(item, dict):
                raise ComparisonError(f"Item {catalog_name} ke-{index} bukan object.")
            key = str(item.get("canonical_key") or "").strip()
            name = str(item.get("canonical_name") or "").strip()
            if not key or not name:
                raise ComparisonError(f"Item {catalog_name} ke-{index} tidak memiliki canonical_key/name.")
            if key in seen_keys:
                raise ComparisonError(f"canonical_key duplikat pada {catalog_name}: {key}.")
            seen_keys.add(key)

            line_map_fields = (
                ("operator_allocations_by_line", "machine_allocations_by_line")
                if catalog_name == "process_catalog"
                else ("counts_by_line",)
            )
            for map_field in line_map_fields:
                line_map = item.get(map_field) or {}
                if not isinstance(line_map, dict):
                    raise ComparisonError(
                        f"{catalog_name}[{index}].{map_field} harus berupa object/dictionary."
                    )
                unknown_labels = sorted(set(line_map) - set(labels))
                if unknown_labels:
                    raise ComparisonError(
                        f"{catalog_name}[{index}].{map_field} memiliki label line yang tidak ada "
                        f"pada daftar sumber: {unknown_labels}."
                    )

    candidates = data.get("process_match_candidates", [])
    if not isinstance(candidates, list):
        raise ComparisonError("process_match_candidates harus berupa list.")
    return labels


def validate_mapping_sync(
    preview: Mapping[str, Any],
    mapping: Mapping[str, Any],
) -> None:
    """Refuse a preview that predates mapping decisions or alias changes.

    The user must rerun the normalizer after making review decisions. Otherwise
    this comparison could silently ignore recently approved process aliases.
    """
    stats = preview.get("mapping_stats") or {}
    if not isinstance(stats, dict):
        raise ComparisonError("mapping_stats pada preview tidak valid.")

    for stat_key, mapping_key in (
        ("process_aliases", "process_aliases"),
        ("machine_aliases", "machine_aliases"),
    ):
        preview_count = stats.get(stat_key)
        current_count = len(mapping.get(mapping_key) or {})
        if preview_count is not None and int(preview_count) != current_count:
            raise ComparisonError(
                "Preview belum sinkron dengan mapping.json: "
                f"jumlah {mapping_key} pada preview={preview_count}, saat ini={current_count}. "
                "Jalankan ulang normalizer.py setelah perubahan mapping."
            )

    saved_decisions = mapping.get("process_match_decisions") or {}
    if not isinstance(saved_decisions, dict):
        raise ComparisonError("process_match_decisions pada mapping.json tidak valid.")

    for candidate in preview.get("process_match_candidates") or []:
        if not isinstance(candidate, dict):
            continue
        norm_a = normalize_process_name(candidate.get("normalized_a") or candidate.get("process_a"))
        norm_b = normalize_process_name(candidate.get("normalized_b") or candidate.get("process_b"))
        if norm_a and norm_b and _pair_id(norm_a, norm_b) in saved_decisions:
            raise ComparisonError(
                "Preview masih memuat pasangan proses yang sudah diputuskan di mapping.json. "
                "Jalankan ulang normalizer.py agar alias SAME diterapkan dan pasangan DIFFERENT dikeluarkan."
            )


def _source_names_by_line(
    normalized_lines: Sequence[Mapping[str, Any]],
    canonical_key: str,
    line_labels: Sequence[str],
) -> Tuple[Dict[str, list[str]], Dict[str, list[Any]], Dict[str, list[str]]]:
    source_by_line: Dict[str, list[str]] = {label: [] for label in line_labels}
    numbers_by_line: Dict[str, list[Any]] = {label: [] for label in line_labels}
    groups_by_line: Dict[str, list[str]] = {label: [] for label in line_labels}
    for line in normalized_lines:
        label = line["line_label"]
        for process in line.get("processes", []) or []:
            if str(process.get("canonical_key") or "") != canonical_key:
                continue
            source_by_line[label] = list(process.get("source_names") or [])
            numbers_by_line[label] = list(process.get("process_numbers") or [])
            groups_by_line[label] = list(process.get("groups") or [])
    return source_by_line, numbers_by_line, groups_by_line


def build_comparison(preview: Mapping[str, Any]) -> Dict[str, Any]:
    """Build dynamic summary, machine, and operator-process comparison tables."""
    line_labels = validate_preview_structure(preview)
    normalized_lines = preview["lines"]
    line_by_label = {line["line_label"]: line for line in normalized_lines}

    # SUMMARY rows: keep the five KPI rows represented in template_compare.xlsx.
    summary_rows = []
    for field, caption, unit in SUMMARY_FIELDS:
        values = OrderedDict(
            (label, _number_or_none(line_by_label[label]["summary"].get(field), field, label))
            for label in line_labels
        )
        summary_rows.append({
            "key": field,
            "label": caption,
            "unit": unit,
            "values_by_line": dict(values),
        })

    # Machine comparison: missing machine types are represented by numeric 0.
    machine_rows = []
    machine_totals: OrderedDict[str, int] = OrderedDict((label, 0) for label in line_labels)
    for item in preview.get("machine_catalog", []) or []:
        counts = item.get("counts_by_line") or {}
        values: OrderedDict[str, int] = OrderedDict()
        for label in line_labels:
            count = _non_negative_count(counts.get(label, 0), "machine count", label)
            values[label] = count
            machine_totals[label] += count
        machine_rows.append({
            "canonical_key": item["canonical_key"],
            "machine": item["canonical_name"],
            "source_names": list(item.get("source_names") or []),
            "values_by_line": dict(values),
            "present_in_lines": list(item.get("lines") or []),
            "alias_applied": bool(item.get("alias_applied", False)),
        })

    # Process comparison. Do not infer a semantic match here; only use the
    # canonical catalog already produced from exact normalization/approved aliases.
    pending_candidates = list(preview.get("process_match_candidates") or [])
    pending_by_name: Dict[str, list[Dict[str, Any]]] = {}
    for candidate in pending_candidates:
        if not isinstance(candidate, dict):
            continue
        for field in ("normalized_a", "normalized_b"):
            key = normalize_process_name(candidate.get(field) or candidate.get(
                "process_a" if field == "normalized_a" else "process_b"
            ))
            if key:
                pending_by_name.setdefault(key, []).append(candidate)

    process_rows = []
    operator_totals: OrderedDict[str, int] = OrderedDict((label, 0) for label in line_labels)
    process_machine_totals: OrderedDict[str, int] = OrderedDict((label, 0) for label in line_labels)
    pending_catalog_rows = 0

    for item in preview.get("process_catalog", []) or []:
        canonical_key = item["canonical_key"]
        operator_map = item.get("operator_allocations_by_line") or {}
        machine_map = item.get("machine_allocations_by_line") or {}
        operator_values: OrderedDict[str, int] = OrderedDict()
        machine_values: OrderedDict[str, int] = OrderedDict()
        for label in line_labels:
            operator_count = _non_negative_count(operator_map.get(label, 0), "operator allocation", label)
            machine_count = _non_negative_count(machine_map.get(label, 0), "process machine allocation", label)
            operator_values[label] = operator_count
            machine_values[label] = machine_count
            operator_totals[label] += operator_count
            process_machine_totals[label] += machine_count

        source_by_line, numbers_by_line, groups_by_line = _source_names_by_line(
            normalized_lines, canonical_key, line_labels
        )
        source_names = list(item.get("source_names") or [])
        candidate_refs: list[Dict[str, Any]] = []
        for source_name in source_names:
            source_key = normalize_process_name(source_name)
            for candidate in pending_by_name.get(source_key, []):
                ref = {
                    "process_a": candidate.get("process_a"),
                    "process_b": candidate.get("process_b"),
                    "similarity": candidate.get("similarity"),
                    "confidence": candidate.get("confidence"),
                }
                if ref not in candidate_refs:
                    candidate_refs.append(ref)

        present_lines = list(item.get("lines") or [])
        alias_applied = bool(item.get("alias_applied", False))
        if candidate_refs:
            match_status = "REVIEW_REQUIRED"
            pending_catalog_rows += 1
        elif alias_applied:
            match_status = "APPROVED_ALIAS"
        elif len(present_lines) >= 2:
            match_status = "MATCHED_BY_NORMALIZATION"
        else:
            match_status = "ONLY_IN_LISTED_LINES"

        all_groups: list[str] = []
        for label in line_labels:
            for group in groups_by_line[label]:
                if group not in all_groups:
                    all_groups.append(group)

        process_rows.append({
            "canonical_key": canonical_key,
            "process": item["canonical_name"],
            "source_names": source_names,
            "source_names_by_line": source_by_line,
            "groups": all_groups,
            "groups_by_line": groups_by_line,
            "process_numbers_by_line": numbers_by_line,
            "operator_allocations_by_line": dict(operator_values),
            "machine_allocations_by_line": dict(machine_values),
            "present_in_lines": present_lines,
            "alias_applied": alias_applied,
            "matching_status": match_status,
            "review_candidates": candidate_refs,
        })

    # Reconcile all totals against source summary totals before allowing export.
    errors: list[str] = []
    warnings: list[str] = []
    line_audits = []
    for label in line_labels:
        summary = line_by_label[label]["summary"]
        expected_operators = _non_negative_count(summary.get("operator_allocations"), "operator_allocations", label)
        expected_machines = _non_negative_count(summary.get("machine_allocations"), "machine_allocations", label)
        actual_operators = operator_totals[label]
        actual_machines_from_catalog = machine_totals[label]
        actual_machines_from_processes = process_machine_totals[label]

        if actual_operators != expected_operators:
            errors.append(
                f"{label}: total alokasi operator pada katalog proses={actual_operators}, "
                f"tetapi SUMMARY sumber={expected_operators}."
            )
        if actual_machines_from_catalog != expected_machines:
            errors.append(
                f"{label}: total mesin pada katalog mesin={actual_machines_from_catalog}, "
                f"tetapi SUMMARY sumber={expected_machines}."
            )
        if actual_machines_from_processes != expected_machines:
            errors.append(
                f"{label}: total alokasi mesin pada katalog proses={actual_machines_from_processes}, "
                f"tetapi SUMMARY sumber={expected_machines}."
            )

        line_audits.append({
            "line_label": label,
            "summary_operator_allocations": expected_operators,
            "process_operator_allocations": actual_operators,
            "summary_machine_allocations": expected_machines,
            "machine_catalog_total": actual_machines_from_catalog,
            "process_machine_allocations": actual_machines_from_processes,
            "operator_match": actual_operators == expected_operators,
            "machine_match": actual_machines_from_catalog == expected_machines == actual_machines_from_processes,
        })

    if pending_candidates:
        warnings.append(
            f"Masih ada {len(pending_candidates)} kandidat nama proses yang belum dikonfirmasi. "
            "Baris terkait tetap dipisahkan; tidak ada penggabungan otomatis."
        )
    if pending_catalog_rows:
        warnings.append(
            f"{pending_catalog_rows} baris proses memiliki kandidat review yang belum diputuskan. "
            "Periksa kolom matching_status sebelum menggunakan laporan."
        )

    if errors:
        raise ComparisonError(
            "Rekonsiliasi komparasi gagal. Data tidak boleh diteruskan ke Excel.\n- "
            + "\n- ".join(errors)
        )

    return {
        "schema_version": 1,
        "comparison_type": "dynamic_multi_line",
        "line_count": len(line_labels),
        "line_columns": [
            {
                "line_label": line["line_label"],
                "line_number": line.get("line_number"),
                "style": line.get("style", ""),
                "source_file": line.get("source_file", ""),
            }
            for line in normalized_lines
        ],
        "summary_rows": summary_rows,
        "machine_rows": machine_rows,
        "machine_total_row": {
            "label": "TOTAL MESIN",
            "values_by_line": dict(machine_totals),
        },
        "process_rows": process_rows,
        "quality": {
            "valid": True,
            "errors": errors,
            "warnings": warnings,
            "line_audits": line_audits,
            "pending_process_match_candidates": len(pending_candidates),
            "process_rows_requiring_review": pending_catalog_rows,
            "unconfirmed_processes_are_not_merged": True,
            "operator_count_basis": "alokasi operator berdasarkan jumlah sel nama operator terisi",
        },
        "source_mapping_stats": dict(preview.get("mapping_stats") or {}),
    }


def format_comparison_report(result: Mapping[str, Any], output_path: Path) -> str:
    labels = [item["line_label"] for item in result["line_columns"]]
    lines = [
        "=" * 78,
        "COMPARE_CT — TAHAP 4: DYNAMIC COMPARISON ENGINE",
        "=" * 78,
        f"JUMLAH LINE       : {result['line_count']}",
        f"JUMLAH JENIS MESIN: {len(result['machine_rows'])}",
        f"JUMLAH PROSES     : {len(result['process_rows'])}",
        f"KANDIDAT REVIEW   : {result['quality']['pending_process_match_candidates']}",
        "",
        "REKAPITULASI PER LINE:",
    ]
    for audit in result["quality"]["line_audits"]:
        lines.append(
            f"  {audit['line_label']}: OPT {audit['process_operator_allocations']}/"
            f"{audit['summary_operator_allocations']} | MESIN {audit['machine_catalog_total']}/"
            f"{audit['summary_machine_allocations']} | "
            f"{'PASS' if audit['operator_match'] and audit['machine_match'] else 'FAIL'}"
        )
    lines.extend(["", "SUMMARY DINAMIS:"])
    for row in result["summary_rows"]:
        values = " | ".join(f"{label}: {row['values_by_line'][label]}" for label in labels)
        lines.append(f"  {row['label']}: {values} {row['unit']}")
    lines.extend([
        "",
        "PERINGATAN:",
    ])
    if result["quality"]["warnings"]:
        lines.extend(f"  - {warning}" for warning in result["quality"]["warnings"])
    else:
        lines.append("  - Tidak ada.")
    lines.extend([
        "",
        f"OUTPUT JSON       : {output_path}",
        "STATUS            : PASS",
        "CATATAN           : Baris dengan REVIEW_REQUIRED tetap terpisah.",
        "=" * 78,
    ])
    return "\n".join(lines)


def run_comparison(
    preview_path: str | Path = DEFAULT_OUTPUT_PATH,
    mapping_path: str | Path = DEFAULT_MAPPING_PATH,
    output_path: str | Path = DEFAULT_COMPARISON_OUTPUT,
) -> Dict[str, Any]:
    preview_file = Path(preview_path).expanduser()
    if not preview_file.is_absolute():
        preview_file = (Path.cwd() / preview_file).resolve()
    preview = _read_json_object(preview_file)

    mapping = load_mapping(mapping_path)
    validate_mapping_sync(preview, mapping)
    result = build_comparison(preview)

    output_file = Path(output_path).expanduser()
    if not output_file.is_absolute():
        output_file = (Path.cwd() / output_file).resolve()
    output_file.parent.mkdir(parents=True, exist_ok=True)
    temp_file = output_file.with_name(output_file.name + ".tmp")
    temp_file.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    os.replace(temp_file, output_file)
    print(format_comparison_report(result, output_file))
    return result


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Gabungkan katalog proses dan mesin menjadi struktur komparasi dinamis untuk 2+ line."
    )
    parser.add_argument(
        "preview", nargs="?", default=str(DEFAULT_OUTPUT_PATH),
        help="File logs/normalization_preview.json hasil normalizer.py",
    )
    parser.add_argument("--mapping", default=str(DEFAULT_MAPPING_PATH), help="Lokasi config/mapping.json")
    parser.add_argument("--output", default=str(DEFAULT_COMPARISON_OUTPUT), help="Lokasi output JSON komparasi")
    args = parser.parse_args(argv)

    try:
        run_comparison(args.preview, args.mapping, args.output)
        return 0
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        print(f"[GAGAL] Comparison engine dihentikan: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
