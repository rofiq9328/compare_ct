"""Interactive approval workflow for process-name matching candidates.

Run from the COMPARE_CT project root:
    py -m core.matching_review "logs\\normalization_preview.json"

Decisions:
- S = SAME: save approved aliases (default canonical name = process A).
- B = DIFFERENT: remember this exact pair as different; do not suggest it again.
- T = DEFER: keep the pair pending for a later review.
- Q = QUIT: stop; decisions already made are saved immediately.

This module never changes the source Excel workbooks or the preview JSON.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Sequence, Tuple

from core.normalizer import (
    DEFAULT_MAPPING_PATH,
    DEFAULT_OUTPUT_PATH,
    load_mapping,
    normalize_process_name,
    _validate_alias_section,
)

PROJECT_ROOT = Path(__file__).resolve().parents[1]


def pair_id(normalized_a: str, normalized_b: str) -> str:
    """Create an order-independent stable ID for a pair of normalized names."""
    pair = sorted((normalized_a, normalized_b))
    payload = json.dumps(pair, ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _read_json_object(path: Path) -> Dict[str, Any]:
    if not path.exists():
        raise FileNotFoundError(f"File tidak ditemukan: {path}")
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"JSON tidak valid pada {path}: {exc}") from exc
    if not isinstance(data, dict):
        raise ValueError(f"Isi file harus berupa object JSON: {path}")
    return data


def _atomic_write_json(path: Path, data: Mapping[str, Any]) -> None:
    """Write JSON atomically to reduce the risk of a damaged mapping file."""
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".tmp")
    payload = json.dumps(data, ensure_ascii=False, indent=2) + "\n"
    temporary.write_text(payload, encoding="utf-8")
    os.replace(temporary, path)


def load_review_candidates(preview_path: str | Path) -> List[Dict[str, Any]]:
    path = Path(preview_path)
    if not path.is_absolute():
        path = (Path.cwd() / path).resolve()
    preview = _read_json_object(path)
    raw_candidates = preview.get("process_match_candidates")
    if not isinstance(raw_candidates, list):
        raise ValueError("Preview tidak memiliki daftar process_match_candidates yang valid.")

    candidates: List[Dict[str, Any]] = []
    seen: set[str] = set()
    for index, raw in enumerate(raw_candidates, start=1):
        if not isinstance(raw, dict):
            continue
        name_a = str(raw.get("process_a") or "").strip()
        name_b = str(raw.get("process_b") or "").strip()
        norm_a = normalize_process_name(raw.get("normalized_a") or name_a)
        norm_b = normalize_process_name(raw.get("normalized_b") or name_b)
        if not name_a or not name_b or not norm_a or not norm_b or norm_a == norm_b:
            continue
        key = pair_id(norm_a, norm_b)
        if key in seen:
            continue
        seen.add(key)
        item = dict(raw)
        item.update({
            "process_a": name_a,
            "process_b": name_b,
            "normalized_a": norm_a,
            "normalized_b": norm_b,
            "pair_id": key,
            "preview_index": index,
        })
        candidates.append(item)

    # Highest similarity first; preserve the preview's original order for ties.
    candidates.sort(key=lambda x: (-float(x.get("similarity") or 0), x["preview_index"]))
    return candidates


def _validated_decisions(raw_mapping: Mapping[str, Any]) -> Dict[str, Dict[str, Any]]:
    raw = raw_mapping.get("process_match_decisions", {})
    if raw is None:
        return {}
    if not isinstance(raw, dict):
        raise ValueError("'process_match_decisions' pada mapping.json harus berupa object.")

    decisions: Dict[str, Dict[str, Any]] = {}
    for key, record in raw.items():
        if not isinstance(key, str) or not isinstance(record, dict):
            raise ValueError("Record process_match_decisions tidak valid.")
        decision = str(record.get("decision") or "").upper()
        if decision not in {"SAME", "DIFFERENT"}:
            raise ValueError(f"Keputusan tersimpan tidak dikenal pada record {key}.")
        norm_a = normalize_process_name(record.get("normalized_a") or record.get("process_a"))
        norm_b = normalize_process_name(record.get("normalized_b") or record.get("process_b"))
        if not norm_a or not norm_b or norm_a == norm_b:
            raise ValueError(f"Nama pasangan tidak valid pada record keputusan {key}.")
        expected = pair_id(norm_a, norm_b)
        if key != expected:
            raise ValueError(f"ID pasangan tidak cocok pada record keputusan {key}.")
        decisions[key] = dict(record, decision=decision, normalized_a=norm_a, normalized_b=norm_b)
    return decisions


def pending_candidates(
    candidates: Sequence[Dict[str, Any]],
    mapping_path: str | Path = DEFAULT_MAPPING_PATH,
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """Remove pairs that have already received a SAME or DIFFERENT decision."""
    mapping_file = Path(mapping_path)
    if not mapping_file.is_absolute():
        mapping_file = (Path.cwd() / mapping_file).resolve()
    validated = load_mapping(mapping_file)
    raw_mapping = _read_json_object(mapping_file)
    decisions = _validated_decisions(raw_mapping)

    pending = [item for item in candidates if item["pair_id"] not in decisions]
    return pending, {
        "mapping_path": str(mapping_file),
        "decisions": decisions,
        "validated_mapping": validated,
    }


def save_decision(
    candidate: Mapping[str, Any],
    decision: str,
    mapping_path: str | Path = DEFAULT_MAPPING_PATH,
    canonical_name: str | None = None,
    note: str = "",
) -> Dict[str, Any]:
    """Persist a SAME or DIFFERENT decision and, for SAME, approved aliases."""
    decision = decision.strip().upper()
    if decision not in {"SAME", "DIFFERENT"}:
        raise ValueError("Keputusan yang dapat disimpan hanya SAME atau DIFFERENT.")

    path = Path(mapping_path)
    if not path.is_absolute():
        path = (Path.cwd() / path).resolve()
    raw = _read_json_object(path)

    name_a = str(candidate.get("process_a") or "").strip()
    name_b = str(candidate.get("process_b") or "").strip()
    norm_a = normalize_process_name(candidate.get("normalized_a") or name_a)
    norm_b = normalize_process_name(candidate.get("normalized_b") or name_b)
    if not name_a or not name_b or not norm_a or not norm_b or norm_a == norm_b:
        raise ValueError("Pasangan nama proses tidak valid atau sudah identik setelah normalisasi.")

    key = pair_id(norm_a, norm_b)
    if decision == "SAME":
        standard_name = (canonical_name or name_a).strip()
        standard_key = normalize_process_name(standard_name)
        if not standard_key:
            raise ValueError("Nama standar tidak boleh kosong.")

        raw_aliases = raw.setdefault("process_aliases", {})
        if not isinstance(raw_aliases, dict):
            raise ValueError("'process_aliases' pada mapping.json harus berupa object.")
        existing = _validate_alias_section(raw_aliases, "process_aliases", normalize_process_name)

        for source_name in (name_a, name_b):
            source_key = normalize_process_name(source_name)
            if source_key == standard_key:
                continue
            old_target = existing.get(source_key)
            if old_target and normalize_process_name(old_target) != standard_key:
                raise ValueError(
                    f"Konflik alias: '{source_name}' sudah dipetakan ke '{old_target}', "
                    f"tidak bisa dipindahkan otomatis ke '{standard_name}'."
                )
            raw_aliases[source_name] = standard_name

    decisions = raw.setdefault("process_match_decisions", {})
    if not isinstance(decisions, dict):
        raise ValueError("'process_match_decisions' pada mapping.json harus berupa object.")

    decisions[key] = {
        "decision": decision,
        "process_a": name_a,
        "process_b": name_b,
        "normalized_a": norm_a,
        "normalized_b": norm_b,
        "canonical_name": (canonical_name or name_a).strip() if decision == "SAME" else None,
        "note": note.strip(),
        "decided_at_utc": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
    }

    # Validate the complete resulting mapping before replacing the file.
    _validate_alias_section(raw.get("process_aliases", {}), "process_aliases", normalize_process_name)
    _validated_decisions(raw)
    _atomic_write_json(path, raw)
    return decisions[key]


def format_candidate(candidate: Mapping[str, Any], number: int, total: int) -> str:
    score = float(candidate.get("similarity") or 0)
    lines_a = ", ".join(candidate.get("lines_a") or []) or "-"
    lines_b = ", ".join(candidate.get("lines_b") or []) or "-"
    return "\n".join([
        "=" * 76,
        f"REVIEW PROSES {number}/{total}  |  KEMIRIPAN {score:.2%}  |  {candidate.get('confidence', '-')}",
        "-" * 76,
        f"A: {candidate['process_a']}",
        f"   Sumber: {lines_a}",
        "",
        f"B: {candidate['process_b']}",
        f"   Sumber: {lines_b}",
        "-" * 76,
        "S = SAMA  |  B = BERBEDA  |  T = TUNDA  |  Q = KELUAR",
    ])


def run_interactive_review(
    preview_path: str | Path = DEFAULT_OUTPUT_PATH,
    mapping_path: str | Path = DEFAULT_MAPPING_PATH,
    input_fn=input,
    output_fn=print,
) -> int:
    candidates = load_review_candidates(preview_path)
    pending, state = pending_candidates(candidates, mapping_path)
    decisions = state["decisions"]
    output_fn("=" * 76)
    output_fn("COMPARE_CT — REVIEW SMART MATCHING NAMA PROSES")
    output_fn("=" * 76)
    output_fn(f"Kandidat pada preview : {len(candidates)}")
    output_fn(f"Sudah diputuskan     : {len(candidates) - len(pending)}")
    output_fn(f"Menunggu review      : {len(pending)}")
    output_fn(f"Mapping              : {state['mapping_path']}")
    output_fn("")
    output_fn("Keputusan disimpan langsung setelah setiap pilihan SAME atau DIFFERENT.")
    output_fn("TUNDA tidak mengubah mapping dan akan muncul lagi pada review berikutnya.")

    saved = 0
    for index, candidate in enumerate(pending, start=1):
        output_fn(format_candidate(candidate, index, len(pending)))
        while True:
            choice = input_fn("Keputusan [S/B/T/Q]: ").strip().upper()
            if choice in {"S", "B", "T", "Q"}:
                break
            output_fn("Pilihan tidak dikenali. Masukkan S, B, T, atau Q.")

        if choice == "Q":
            output_fn("Review dihentikan. Keputusan sebelumnya sudah tersimpan.")
            return saved
        if choice == "T":
            output_fn("Ditunda; belum ada perubahan pada mapping.json.")
            continue

        if choice == "S":
            default_name = candidate["process_a"]
            canonical = input_fn(
                f"Nama standar [Enter = gunakan '{default_name}']: "
            ).strip() or default_name
            try:
                record = save_decision(candidate, "SAME", mapping_path, canonical)
            except Exception as exc:
                output_fn(f"[GAGAL SIMPAN] {exc}")
                continue
            output_fn(f"[TERSIMPAN] SAMA -> nama standar: {record['canonical_name']}")
        else:
            note = input_fn("Catatan (opsional, Enter untuk kosong): ").strip()
            try:
                save_decision(candidate, "DIFFERENT", mapping_path, note=note)
            except Exception as exc:
                output_fn(f"[GAGAL SIMPAN] {exc}")
                continue
            output_fn("[TERSIMPAN] BERBEDA -> pasangan ini tidak diminta lagi.")
        saved += 1

    output_fn("=" * 76)
    output_fn("REVIEW SELESAI")
    output_fn(f"Keputusan baru tersimpan : {saved}")
    output_fn(f"Masih tertunda           : {sum(1 for x in pending if pair_id(x['normalized_a'], x['normalized_b']) not in _validated_decisions(_read_json_object(Path(state['mapping_path']))))}")
    output_fn("Jalankan ulang normalizer untuk memperbarui katalog setelah alias SAME disetujui.")
    return saved


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Konfirmasi kandidat nama proses dan simpan keputusan ke mapping.json."
    )
    parser.add_argument("preview", nargs="?", default=str(DEFAULT_OUTPUT_PATH), help="File normalization_preview.json")
    parser.add_argument("--mapping", default=str(DEFAULT_MAPPING_PATH), help="Lokasi config/mapping.json")
    args = parser.parse_args(argv)
    try:
        run_interactive_review(args.preview, args.mapping)
        return 0
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        print(f"[GAGAL] Review tidak dapat dijalankan: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
