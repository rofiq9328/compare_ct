import copy
import unittest

from core.analysis import AnalysisError, build_analysis


class TestAnalysisEngine(unittest.TestCase):
    def setUp(self):
        l37 = "LINE 37 (A8MVW)"
        l61 = "LINE 61 (A8MVX)"
        self.data = {
            "schema_version": 1,
            "comparison_type": "dynamic_multi_line",
            "line_count": 2,
            "line_columns": [
                {"line_label": l37, "line_number": 37, "style": "A8MVW", "source_file": "l37.xlsx"},
                {"line_label": l61, "line_number": 61, "style": "A8MVX", "source_file": "l61.xlsx"},
            ],
            "summary_rows": [
                {"key": "operator_allocations", "label": "Jumlah Operator", "unit": "OPT", "values_by_line": {l37: 54, l61: 51}},
                {"key": "takt_time", "label": "Takt Time", "unit": "detik", "values_by_line": {l37: 39, l61: 41}},
                {"key": "target_per_hour", "label": "Target Output / Jam", "unit": "pcs", "values_by_line": {l37: 55, l61: 53}},
                {"key": "potential_per_hour", "label": "Potensial Output / Jam", "unit": "pcs", "values_by_line": {l37: 39.5859494315, l61: 28.424719391}},
                {"key": "actual_per_hour", "label": "Aktual Output / Jam", "unit": "pcs", "values_by_line": {l37: 30, l61: 20}},
            ],
            "machine_rows": [
                {"canonical_key": "DU-1481", "machine": "DU-1481", "values_by_line": {l37: 35, l61: 30}, "present_in_lines": [l37, l61], "alias_applied": False},
                {"canonical_key": "H-5150", "machine": "H-5150", "values_by_line": {l37: 1, l61: 0}, "present_in_lines": [l37], "alias_applied": False},
            ],
            "machine_total_row": {"label": "TOTAL MESIN", "values_by_line": {l37: 56, l61: 53}},
            "process_rows": [
                {
                    "canonical_key": "BINDING SIDE", "process": "Binding Side", "source_names": ["Binding Side"],
                    "present_in_lines": [l37, l61], "matching_status": "APPROVED_ALIAS", "alias_applied": True,
                    "operator_allocations_by_line": {l37: 2, l61: 3}, "machine_allocations_by_line": {l37: 1, l61: 1},
                    "groups_by_line": {l37: ["FRONT"], l61: ["FRONT"]}, "review_candidates": [],
                },
                {
                    "canonical_key": "BARTACK HANDLE", "process": "Bartack Handle", "source_names": ["Bartack Handle"],
                    "present_in_lines": [l37], "matching_status": "ONLY_IN_LISTED_LINES", "alias_applied": False,
                    "operator_allocations_by_line": {l37: 1, l61: 0}, "machine_allocations_by_line": {l37: 1, l61: 0},
                    "groups_by_line": {l37: ["BACK"], l61: []}, "review_candidates": [],
                },
                {
                    "canonical_key": "JOIN POCKET", "process": "Join Pocket B", "source_names": ["Join Pocket B"],
                    "present_in_lines": [l61], "matching_status": "REVIEW_REQUIRED", "alias_applied": False,
                    "operator_allocations_by_line": {l37: 0, l61: 1}, "machine_allocations_by_line": {l37: 0, l61: 1},
                    "groups_by_line": {l37: [], l61: ["FRONT"]},
                    "review_candidates": [{"process_a": "Join Pocket", "process_b": "Join Pocket B", "similarity": 0.94}],
                },
            ],
            "quality": {
                "valid": True, "errors": [], "warnings": ["Masih ada kandidat review."],
                "pending_process_match_candidates": 1, "process_rows_requiring_review": 1,
                "operator_count_basis": "alokasi operator berdasarkan jumlah sel nama operator terisi",
            },
        }

    def test_calculates_line_kpis(self):
        result = build_analysis(self.data)
        l37 = result["line_kpis"]["LINE 37 (A8MVW)"]
        self.assertEqual(l37["actual_vs_target_pct"], 54.55)
        self.assertEqual(l37["actual_vs_potential_pct"], 75.78)
        self.assertEqual(l37["target_gap_pcs_per_hour"], 25)
        self.assertEqual(l37["production_status"], "AKTUAL_DI_BAWAH_TARGET")

    def test_calculates_summary_deltas_against_first_line(self):
        result = build_analysis(self.data)
        actual = next(row for row in result["summary_differences"] if row["key"] == "actual_per_hour")
        self.assertEqual(actual["values_by_line"]["LINE 37 (A8MVW)"]["delta_vs_reference"], 0)
        self.assertEqual(actual["values_by_line"]["LINE 61 (A8MVX)"]["delta_vs_reference"], -10)

    def test_review_required_process_is_not_reported_as_semantically_unique(self):
        result = build_analysis(self.data)
        names = [item["process"] for item in result["process_analysis"]["processes_present_in_one_line_only"]]
        self.assertEqual(names, ["Bartack Handle"])
        self.assertEqual(len(result["process_analysis"]["processes_requiring_matching_review"]), 1)

    def test_reports_matched_process_allocation_difference(self):
        result = build_analysis(self.data)
        rows = result["process_analysis"]["matched_processes_with_allocation_difference"]
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["operator_allocation_spread"], 1)

    def test_reports_machine_allocated_in_one_line_only(self):
        result = build_analysis(self.data)
        rows = result["machine_analysis"]["machine_types_allocated_in_one_line_only"]
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["machine"], "H-5150")

    def test_marks_bottleneck_as_not_assessed_without_process_ct(self):
        result = build_analysis(self.data)
        self.assertTrue(result["quality"]["bottleneck_assessment"].startswith("NOT_ASSESSED"))

    def test_includes_review_observation(self):
        result = build_analysis(self.data)
        self.assertTrue(any(item["type"] == "UNCONFIRMED_PROCESS_MATCHES" for item in result["observations"]))

    def test_rejects_invalid_comparison_quality(self):
        broken = copy.deepcopy(self.data)
        broken["quality"]["valid"] = False
        with self.assertRaises(AnalysisError):
            build_analysis(broken)

    def test_rejects_operator_basis_drift(self):
        broken = copy.deepcopy(self.data)
        broken["quality"]["operator_count_basis"] = "unique operator names"
        with self.assertRaises(AnalysisError):
            build_analysis(broken)

    def test_rejects_missing_line_value(self):
        broken = copy.deepcopy(self.data)
        broken["machine_rows"][0]["values_by_line"].pop("LINE 61 (A8MVX)")
        with self.assertRaises(AnalysisError):
            build_analysis(broken)

    def test_supports_three_or_more_lines(self):
        expanded = copy.deepcopy(self.data)
        l15 = "LINE 15 (STYLE C)"
        expanded["line_count"] = 3
        expanded["line_columns"].append({"line_label": l15, "line_number": 15, "style": "STYLE C", "source_file": "l15.xlsx"})
        for row in expanded["summary_rows"]:
            row["values_by_line"][l15] = row["values_by_line"]["LINE 37 (A8MVW)"]
        for row in expanded["machine_rows"]:
            row["values_by_line"][l15] = 0
        for row in expanded["process_rows"]:
            row["operator_allocations_by_line"][l15] = 0
            row["machine_allocations_by_line"][l15] = 0
        result = build_analysis(expanded)
        self.assertEqual(result["line_count"], 3)
        self.assertIn(l15, result["line_kpis"])

    def test_group_variation_is_flagged_for_review(self):
        expanded = copy.deepcopy(self.data)
        expanded["process_rows"][0]["groups_by_line"]["LINE 61 (A8MVX)"] = ["BACK"]
        result = build_analysis(expanded)
        self.assertEqual(len(result["process_analysis"]["process_group_variations"]), 1)


if __name__ == "__main__":
    unittest.main()
