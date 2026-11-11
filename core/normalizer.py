"""Conservative normalisation and review suggestions for COMPARE_CT.

Stage 3 responsibilities:
- normalise process names without changing their meaning;
- apply only user-approved aliases from config/mapping.json;
- normalise machine-code spelling conservatively;
- suggest (but NEVER auto-merge) similar process names for human review.

Run from the project root on Windows with Excel + pywin32:
    py -m core.normalizer "input\\LINE37.xlsx" "input\\LINE61.xlsx"

Optional:
    py -m core.normalizer --mapping config/mapping.json --output logs/normalization_preview.json FILE1 FILE2
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import unicodedata
from collections import OrderedDict
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Sequence, Tuple

from core.reader import clean_text, read_many
from core.validator import validate_many, format_validation_report


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_MAPPING_PATH = PROJECT_ROOT / "config" / "mapping.json"
DEFAULT_OUTPUT_PATH = PROJECT_ROOT / "logs" / "normalization_preview.json"
MAPPING_SCHEMA_VERSION = 1


def normalize_process_name(value: Any) -> str:
    """Return a conservative comparison key for a process name.

    It removes Excel/control artifacts and normalises case, Unicode, spacing,
    common conjunction symbols and punctuation. It deliberately preserves all
    meaningful words, numbers, measurements and operation verbs.
    """
    text = clean_text(value)
    if not text:
        return ""

    text = unicodedata.normalize("NFKC", text).upper()
    text = text.replace("×", "X").replace("✕", "X")
    text = text.replace("&", " DAN ").replace("+", " DAN ")
    text = text.replace("–", "-").replace("—", "-")
    # Hyphens and slashes are separators in process descriptions, not machine IDs.
    text = text.replace("-", " ").replace("/", " ")
    # Keep letters/numbers; turn other punctuation into a separator.
    text = "".join(ch if ch.isalnum() or ch.isspace() else " " for ch in text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_machine_code(value: Any) -> str:
    """Normalise machine codes while preserving model suffixes and hyphens.

    Example: ' du - 1481 - 7 ' -> 'DU-1481-7'. The suffix '-7' is retained;
    this module does not assume it is equivalent to 'DU-1481'.
    """
    text = clean_text(value)
    if not text:
        return ""
    text = unicodedata.normalize("NFKC", text).upper()
    text = text.replace("–", "-").replace("—", "-")
    text = re.sub(r"\s*-\s*", "-", text)
    text = re.sub(r"\s+", "", text)
    return text.strip("-")


def _validate_alias_section(raw: Any, section_name: str, normalizer) -> Dict[str, str]:
    if raw is None:
        return {}
    if not isinstance(raw, dict):
        raise ValueError(f"'{section_name}' pada mapping.json harus berupa object/dictionary.")

    aliases: Dict[str, str] = {}
    original_keys: Dict[str, str] = {}
    for alias, canonical in raw.items():
        if not isinstance(alias, str) or not isinstance(canonical, str):
            raise ValueError(f"Semua alias dan canonical pada '{section_name}' harus berupa teks.")
        alias_key = normalizer(alias)
        canonical_name = clean_text(canonical)
        if not alias_key or not canonical_name:
            raise ValueError(f"Alias/canonical kosong ditemukan pada '{section_name}'.")
        if alias_key in aliases and aliases[alias_key] != canonical_name:
            raise ValueError(
                f"Konflik alias '{alias}' dan '{original_keys[alias_key]}' setelah normalisasi."
            )
        aliases[alias_key] = canonical_name
        original_keys[alias_key] = alias
    return aliases


def load_mapping(mapping_path: str | Path = DEFAULT_MAPPING_PATH) -> Dict[str, Any]:
    """Load and validate user-approved aliases and fuzzy-review settings."""
    path = Path(mapping_path)
    if not path.is_absolute():
        path = (Path.cwd() / path).resolve()
    if not path.exists():
        raise FileNotFoundError(
            f"File mapping tidak ditemukan: {path}. Buat config/mapping.json dari template Tahap 3."
        )

    try:
        raw = json.loads(path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"Format JSON mapping tidak valid: {exc}") from exc

    if not isinstance(raw, dict):
        raise ValueError("Isi mapping.json harus berupa object JSON.")
    version = raw.get("schema_version", MAPPING_SCHEMA_VERSION)
    if version != MAPPING_SCHEMA_VERSION:
        raise ValueError(
            f"schema_version mapping tidak didukung: {version}. Yang didukung: {MAPPING_SCHEMA_VERSION}."
        )

    fuzzy = raw.get("fuzzy_review", {})
    if not isinstance(fuzzy, dict):
        raise ValueError("'fuzzy_review' harus berupa object/dictionary.")
    try:
        min_score = float(fuzzy.get("minimum_score", 0.68))
        max_suggestions = int(fuzzy.get("max_suggestions_per_process", 3))
    except (TypeError, ValueError) as exc:
        raise ValueError("Pengaturan fuzzy_review harus berupa angka yang valid.") from exc
    if not 0.0 <= min_score <= 1.0:
        raise ValueError("fuzzy_review.minimum_score harus antara 0 dan 1.")
    if not 1 <= max_suggestions <= 20:
        raise ValueError("fuzzy_review.max_suggestions_per_process harus antara 1 dan 20.")

    raw_decisions = raw.get("process_match_decisions", {})
    if raw_decisions is None:
        raw_decisions = {}
    if not isinstance(raw_decisions, dict):
        raise ValueError("'process_match_decisions' pada mapping.json harus berupa object/dictionary.")

    process_match_decisions: Dict[str, Dict[str, Any]] = {}
    for pair_key, decision_record in raw_decisions.items():
        if not isinstance(pair_key, str) or not isinstance(decision_record, dict):
            raise ValueError("Record process_match_decisions pada mapping.json tidak valid.")
        decision = str(decision_record.get("decision") or "").upper()
        if decision not in {"SAME", "DIFFERENT"}:
            raise ValueError(f"Keputusan process_match_decisions tidak dikenal: {decision!r}.")
        normalized_a = normalize_process_name(
            decision_record.get("normalized_a") or decision_record.get("process_a")
        )
        normalized_b = normalize_process_name(
            decision_record.get("normalized_b") or decision_record.get("process_b")
        )
        if not normalized_a or not normalized_b or normalized_a == normalized_b:
            raise ValueError(f"Nama pasangan tidak valid pada process_match_decisions: {pair_key}.")
        process_match_decisions[pair_key] = dict(
            decision_record,
            decision=decision,
            normalized_a=normalized_a,
            normalized_b=normalized_b,
        )

    return {
        "schema_version": MAPPING_SCHEMA_VERSION,
        "process_aliases": _validate_alias_section(
            raw.get("process_aliases", {}), "process_aliases", normalize_process_name
        ),
        "machine_aliases": _validate_alias_section(
            raw.get("machine_aliases", {}), "machine_aliases", normalize_machine_code
        ),
        "process_match_decisions": process_match_decisions,
        "fuzzy_review": {
            "minimum_score": min_score,
            "max_suggestions_per_process": max_suggestions,
        },
        "mapping_path": str(path),
    }


def _resolve_alias(
    value: str,
    aliases: Mapping[str, str],
    normalizer,
    entity_label: str,
) -> Tuple[str, str, bool]:
    """Resolve alias chains and detect circular aliases."""
    original = clean_text(value)
    current = original
    visited: set[str] = set()
    changed = False
    while True:
        key = normalizer(current)
        if not key or key not in aliases:
            return current, normalizer(current), changed
        if key in visited:
            raise ValueError(f"Siklus alias terdeteksi pada {entity_label}: {original!r}.")
        visited.add(key)
        current = aliases[key]
        changed = True


def _similarity(left: str, right: str) -> float:
    """Blend character, sorted-token and token-overlap similarities."""
    if not left or not right:
        return 0.0
    if left == right:
        return 1.0
    char_ratio = SequenceMatcher(None, left, right).ratio()
    left_tokens = left.split()
    right_tokens = right.split()
    sorted_ratio = SequenceMatcher(
        None, " ".join(sorted(left_tokens)), " ".join(sorted(right_tokens))
    ).ratio()
    left_set, right_set = set(left_tokens), set(right_tokens)
    union = left_set | right_set
    jaccard = len(left_set & right_set) / len(union) if union else 0.0
    # Token Jaccard has lower weight to avoid treating long names with one
    # shared operation word as a strong match.
    return round(0.50 * char_ratio + 0.35 * sorted_ratio + 0.15 * jaccard, 4)


def _confidence(score: float) -> str:
    if score >= 0.90:
        return "SANGAT MIRIP"
    if score >= 0.82:
        return "MIRIP"
    if score >= 0.72:
        return "CUKUP MIRIP"
    return "KEMUNGKINAN RENDAH"


def _add_unique(target: List[Any], value: Any) -> None:
    if value not in target:
        target.append(value)


def normalize_results(
    results: Sequence[Dict[str, Any]],
    mapping: Dict[str, Any],
) -> Dict[str, Any]:
    """Normalize parsed reader results into per-line process/machine catalogs.

    Fuzzy candidates are review-only. No candidate is merged unless an explicit
    alias is added to mapping.json and the normalizer is run again.
    """
    process_aliases = mapping["process_aliases"]
    machine_aliases = mapping["machine_aliases"]
    normalized_lines: List[Dict[str, Any]] = []
    global_processes: "OrderedDict[str, Dict[str, Any]]" = OrderedDict()
    global_machines: "OrderedDict[str, Dict[str, Any]]" = OrderedDict()
    raw_process_occurrences: "OrderedDict[str, Dict[str, Any]]" = OrderedDict()

    for result in results:
        line_label = str(result.get("line_label") or "LINE TIDAK DIKETAHUI")
        line_number = result.get("line_number")
        style = str(result.get("style") or "")
        summary = dict(result.get("summary") or {})

        line_processes: "OrderedDict[str, Dict[str, Any]]" = OrderedDict()
        for item in result.get("process_summary", []) or []:
            if not isinstance(item, dict):
                continue
            source_name = clean_text(item.get("process_name"))
            source_key = normalize_process_name(source_name)
            if not source_key:
                continue
            canonical_name, canonical_key, alias_applied = _resolve_alias(
                source_name, process_aliases, normalize_process_name, "proses"
            )
            process = line_processes.setdefault(canonical_key, {
                "canonical_key": canonical_key,
                "canonical_name": canonical_name,
                "source_names": [],
                "operator_allocations": 0,
                "machine_allocations": 0,
                "groups": [],
                "process_numbers": [],
                "alias_applied": False,
            })
            _add_unique(process["source_names"], source_name)
            process["operator_allocations"] += int(item.get("operator_allocations") or 0)
            process["machine_allocations"] += int(item.get("machine_allocations") or 0)
            process["alias_applied"] = process["alias_applied"] or alias_applied
            for group in item.get("groups", []) or []:
                if clean_text(group):
                    _add_unique(process["groups"], clean_text(group))
            for process_number in item.get("process_numbers", []) or []:
                _add_unique(process["process_numbers"], process_number)

            global_item = global_processes.setdefault(canonical_key, {
                "canonical_key": canonical_key,
                "canonical_name": canonical_name,
                "source_names": [],
                "lines": [],
                "operator_allocations_by_line": OrderedDict(),
                "machine_allocations_by_line": OrderedDict(),
                "alias_applied": False,
            })
            _add_unique(global_item["source_names"], source_name)
            _add_unique(global_item["lines"], line_label)
            global_item["operator_allocations_by_line"][line_label] = (
                global_item["operator_allocations_by_line"].get(line_label, 0)
                + int(item.get("operator_allocations") or 0)
            )
            global_item["machine_allocations_by_line"][line_label] = (
                global_item["machine_allocations_by_line"].get(line_label, 0)
                + int(item.get("machine_allocations") or 0)
            )
            global_item["alias_applied"] = global_item["alias_applied"] or alias_applied

            occurrence = raw_process_occurrences.setdefault(source_key, {
                "source_key": source_key,
                "source_name": source_name,
                "canonical_key": canonical_key,
                "canonical_name": canonical_name,
                "lines": [],
                "alias_applied": alias_applied,
            })
            _add_unique(occurrence["lines"], line_label)

        line_machines: "OrderedDict[str, Dict[str, Any]]" = OrderedDict()
        for source_code, count in (result.get("machine_counts") or {}).items():
            source_code = clean_text(source_code)
            source_key = normalize_machine_code(source_code)
            if not source_key:
                continue
            canonical_name, canonical_key, alias_applied = _resolve_alias(
                source_code, machine_aliases, normalize_machine_code, "mesin"
            )
            machine = line_machines.setdefault(canonical_key, {
                "canonical_key": canonical_key,
                "canonical_name": canonical_name,
                "source_names": [],
                "count": 0,
                "alias_applied": False,
            })
            _add_unique(machine["source_names"], source_code)
            machine["count"] += int(count or 0)
            machine["alias_applied"] = machine["alias_applied"] or alias_applied

            global_machine = global_machines.setdefault(canonical_key, {
                "canonical_key": canonical_key,
                "canonical_name": canonical_name,
                "source_names": [],
                "counts_by_line": OrderedDict(),
                "lines": [],
                "alias_applied": False,
            })
            _add_unique(global_machine["source_names"], source_code)
            _add_unique(global_machine["lines"], line_label)
            global_machine["counts_by_line"][line_label] = (
                global_machine["counts_by_line"].get(line_label, 0) + int(count or 0)
            )
            global_machine["alias_applied"] = global_machine["alias_applied"] or alias_applied

        normalized_lines.append({
            "line_number": line_number,
            "line_label": line_label,
            "style": style,
            "source_file": result.get("source_file", ""),
            "summary": summary,
            "processes": list(line_processes.values()),
            "machines": list(line_machines.values()),
        })

    # Generate cross-line suggestions using unaliased source labels only.
    occurrences = list(raw_process_occurrences.values())
    min_score = float(mapping["fuzzy_review"]["minimum_score"])
    max_per_process = int(mapping["fuzzy_review"]["max_suggestions_per_process"])
    scored_pairs: List[Dict[str, Any]] = []
    reviewed_different_pairs = {
        key for key, record in mapping.get("process_match_decisions", {}).items()
        if str(record.get("decision", "")).upper() == "DIFFERENT"
    }
    for i, left in enumerate(occurrences):
        left_lines = set(left["lines"])
        for right in occurrences[i + 1:]:
            if left["canonical_key"] == right["canonical_key"]:
                continue
            pair_payload = json.dumps(
                sorted((left["source_key"], right["source_key"])),
                ensure_ascii=False, separators=(",", ":")
            )
            pair_key = hashlib.sha256(pair_payload.encode("utf-8")).hexdigest()
            if pair_key in reviewed_different_pairs:
                continue
            right_lines = set(right["lines"])
            # If two differently named processes coexist in the same line,
            # they are evidence of separate operations, not aliases. Only
            # compare names whose observed line sets do not overlap.
            if left_lines.intersection(right_lines):
                continue
            score = _similarity(left["source_key"], right["source_key"])
            if score < min_score:
                continue
            scored_pairs.append({
                "process_a": left["source_name"],
                "process_b": right["source_name"],
                "normalized_a": left["source_key"],
                "normalized_b": right["source_key"],
                "lines_a": sorted(left_lines),
                "lines_b": sorted(right_lines),
                "similarity": score,
                "confidence": _confidence(score),
                "decision": "REVIEW_REQUIRED",
                "auto_merged": False,
                "instruction": "Periksa arti pekerjaan sebelum menambahkan alias ke mapping.json.",
            })

    # Keep only the best N suggestions for each source process name.
    scored_pairs.sort(key=lambda item: (-item["similarity"], item["process_a"], item["process_b"]))
    retained: List[Dict[str, Any]] = []
    per_name_count: Dict[str, int] = {}
    for pair in scored_pairs:
        key_a = pair["normalized_a"]
        key_b = pair["normalized_b"]
        if per_name_count.get(key_a, 0) >= max_per_process:
            continue
        if per_name_count.get(key_b, 0) >= max_per_process:
            continue
        retained.append(pair)
        per_name_count[key_a] = per_name_count.get(key_a, 0) + 1
        per_name_count[key_b] = per_name_count.get(key_b, 0) + 1

    return {
        "schema_version": 1,
        "line_count": len(normalized_lines),
        "lines": normalized_lines,
        "process_catalog": list(global_processes.values()),
        "machine_catalog": list(global_machines.values()),
        "process_match_candidates": retained,
        "rules": {
            "normalization_only": "case, Unicode, spacing, punctuation and conjunction formatting",
            "approved_aliases_only": True,
            "fuzzy_matching_is_suggestion_only": True,
            "machine_fuzzy_matching": False,
            "operator_count_basis": "non-empty operator allocation rows from reader summary",
        },
        "mapping_stats": {
            "process_aliases": len(process_aliases),
            "machine_aliases": len(machine_aliases),
            "fuzzy_candidates": len(retained),
            "fuzzy_threshold": min_score,
        },
    }


def format_normalization_report(normalized: Dict[str, Any], output_path: Path) -> str:
    lines = [
        "=" * 72,
        "COMPARE_CT — TAHAP 3: NORMALISASI & SMART MATCHING REVIEW",
        "=" * 72,
    ]
    for line in normalized["lines"]:
        lines.extend([
            f"LINE              : {line['line_label']}",
            f"PROSES STANDAR    : {len(line['processes'])}",
            f"JENIS MESIN       : {len(line['machines'])}",
            f"ALOKASI OPERATOR  : {line['summary'].get('operator_allocations', 0)}",
            f"ALOKASI MESIN     : {line['summary'].get('machine_allocations', 0)}",
            "-" * 72,
        ])
    lines.extend([
        f"KATALOG PROSES    : {len(normalized['process_catalog'])}",
        f"KATALOG MESIN     : {len(normalized['machine_catalog'])}",
        f"ALIAS PROSES      : {normalized['mapping_stats']['process_aliases']}",
        f"ALIAS MESIN       : {normalized['mapping_stats']['machine_aliases']}",
        f"KANDIDAT REVIEW   : {normalized['mapping_stats']['fuzzy_candidates']}",
        f"FILE PREVIEW      : {output_path}",
        "",
        "CATATAN: Kandidat fuzzy TIDAK digabung otomatis. Semua pasangan harus diperiksa.",
    ])
    if normalized["process_match_candidates"]:
        lines.append("TOP KANDIDAT KEMIRIPAN:")
        for index, item in enumerate(normalized["process_match_candidates"][:15], start=1):
            lines.append(
                f"  {index:02}. [{item['similarity']:.0%}] {item['process_a']}  <=>  {item['process_b']}"
            )
    else:
        lines.append("Tidak ada kandidat fuzzy di atas ambang batas.")
    lines.append("=" * 72)
    return "\n".join(lines)


def run_normalization(
    file_paths: Sequence[str | Path],
    mapping_path: str | Path = DEFAULT_MAPPING_PATH,
) -> Tuple[List[Dict[str, Any]], Dict[str, Any], Dict[str, Any]]:
    """Read, validate and normalize multiple source files."""
    results = read_many(file_paths)
    validation = validate_many(results)
    if not validation["valid"]:
        raise ValueError(format_validation_report(results, validation))
    mapping = load_mapping(mapping_path)
    normalized = normalize_results(results, mapping)
    return results, validation, normalized


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Normalisasi proses/mesin dan sarankan kemiripan nama tanpa auto-merge."
    )
    parser.add_argument("files", nargs="+", help="File Excel sumber (minimal 2)")
    parser.add_argument("--mapping", default=str(DEFAULT_MAPPING_PATH), help="Lokasi config/mapping.json")
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT_PATH), help="Lokasi JSON preview")
    args = parser.parse_args(argv)

    output_path = Path(args.output)
    if not output_path.is_absolute():
        output_path = (Path.cwd() / output_path).resolve()

    try:
        results, validation, normalized = run_normalization(args.files, args.mapping)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(
            json.dumps(normalized, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(format_normalization_report(normalized, output_path))
        return 0
    except Exception as exc:
        print(f"[GAGAL] Normalisasi tidak dapat dijalankan: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
