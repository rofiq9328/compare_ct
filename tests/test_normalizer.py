import unittest

from core.normalizer import (
    normalize_process_name, normalize_machine_code, _similarity,
    normalize_results, load_mapping, _resolve_alias,
)


class TestNormalizer(unittest.TestCase):
    def test_case_space_and_excel_artifact(self):
        self.assertEqual(
            normalize_process_name(" Jahit Dasar + Lapis Sayap AB BACK_x000D_ "),
            "JAHIT DASAR DAN LAPIS SAYAP AB BACK",
        )

    def test_ampersand_and_plus_are_conjunctions(self):
        self.assertEqual(
            normalize_process_name("JOIN LINING & MA KE ZIPPER"),
            normalize_process_name("join lining + MA ke zipper"),
        )

    def test_machine_model_suffix_is_preserved(self):
        self.assertEqual(normalize_machine_code(" du - 1481 - 7 "), "DU-1481-7")
        self.assertNotEqual(normalize_machine_code("DU-1481-7"), normalize_machine_code("DU-1481"))

    def test_fuzzy_similarity_is_suggestion_metric(self):
        self.assertGreater(_similarity("BACKTACK LIPAT WEBBING TANSION TAB AB", "BACKTACK LIPAT UJUNG WEBBING AB TENSION TAB"), 0.60)

    def test_aliases_only_merge_after_explicit_mapping(self):
        mapping = {
            "process_aliases": {normalize_process_name("JAHIT BINDING SIDE"): "JAHIT BINDING"},
            "machine_aliases": {},
            "fuzzy_review": {"minimum_score": 0.68, "max_suggestions_per_process": 3},
        }
        results = [
            {
                "line_label": "LINE 1 (STYLE A)", "line_number": 1, "style": "STYLE A",
                "summary": {"operator_allocations": 1, "machine_allocations": 1},
                "process_summary": [{"process_name": "JAHIT BINDING SIDE", "operator_allocations": 1, "machine_allocations": 1, "groups": [], "process_numbers": [1]}],
                "machine_counts": {"DU-1481": 1},
            },
            {
                "line_label": "LINE 2 (STYLE B)", "line_number": 2, "style": "STYLE B",
                "summary": {"operator_allocations": 1, "machine_allocations": 1},
                "process_summary": [{"process_name": "JAHIT BINDING", "operator_allocations": 1, "machine_allocations": 1, "groups": [], "process_numbers": [2]}],
                "machine_counts": {"DU-1481": 1},
            },
        ]
        normalized = normalize_results(results, mapping)
        self.assertEqual(len(normalized["process_catalog"]), 1)
        self.assertEqual(normalized["process_catalog"][0]["operator_allocations_by_line"]["LINE 1 (STYLE A)"], 1)
        self.assertEqual(normalized["process_catalog"][0]["operator_allocations_by_line"]["LINE 2 (STYLE B)"], 1)

    def test_fuzzy_candidates_are_not_auto_merged(self):
        mapping = {
            "process_aliases": {}, "machine_aliases": {},
            "fuzzy_review": {"minimum_score": 0.50, "max_suggestions_per_process": 3},
        }
        results = [
            {
                "line_label": "LINE 1", "line_number": 1, "style": "A", "summary": {},
                "process_summary": [{"process_name": "JAHIT BINDING SIDE", "operator_allocations": 1, "machine_allocations": 0, "groups": [], "process_numbers": [1]}],
                "machine_counts": {},
            },
            {
                "line_label": "LINE 2", "line_number": 2, "style": "B", "summary": {},
                "process_summary": [{"process_name": "JAHIT BINDING SID", "operator_allocations": 1, "machine_allocations": 0, "groups": [], "process_numbers": [2]}],
                "machine_counts": {},
            },
        ]
        normalized = normalize_results(results, mapping)
        self.assertEqual(len(normalized["process_catalog"]), 2)
        self.assertTrue(normalized["process_match_candidates"])
        self.assertTrue(all(not item["auto_merged"] for item in normalized["process_match_candidates"]))

    def test_alias_cycle_is_detected(self):
        aliases = {"A": "B", "B": "A"}
        with self.assertRaises(ValueError):
            _resolve_alias("A", aliases, lambda value: value, "proses")


if __name__ == "__main__":
    unittest.main()
