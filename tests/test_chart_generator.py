import unittest

from excel.chart_generator import _build_chart_data


class TestDynamicChartData(unittest.TestCase):
    def setUp(self):
        self.labels = [
            "LINE 37 (A8MVW)",
            "LINE 61 (A8MVX)",
            "LINE 52 (STYLE C)",
            "LINE 71 (STYLE D)",
        ]
        self.data = {
            "schema_version": 1,
            "comparison_type": "dynamic_multi_line",
            "line_count": len(self.labels),
            "line_columns": [{"line_label": label} for label in self.labels],
            "summary_rows": [
                {
                    "key": key,
                    "label": label,
                    "unit": unit,
                    "values_by_line": {line: value for line, value in zip(self.labels, values)},
                }
                for key, label, unit, values in (
                    ("operator_allocations", "Jumlah Operator", "OPT", [54, 51, 48, 50]),
                    ("takt_time", "Takt Time", "detik", [39, 41, 40, 38]),
                    ("target_per_hour", "Target Output / Jam", "pcs/jam", [55, 53, 52, 56]),
                    ("potential_per_hour", "Potensial Output / Jam", "pcs/jam", [45.5, 42.1, 40.2, 47.0]),
                    ("actual_per_hour", "Aktual Output / Jam", "pcs/jam", [30, 20, 24, 28]),
                )
            ],
            "machine_rows": [
                {
                    "machine": "DNL-150",
                    "values_by_line": {line: i + 1 for i, line in enumerate(self.labels)},
                },
                {
                    "machine": "DSC-246",
                    "values_by_line": {line: (i + 1) * 2 for i, line in enumerate(self.labels)},
                },
            ],
            "machine_total_row": {
                "values_by_line": {line: 99 for line in self.labels}
            },
            "process_rows": [
                {
                    "process": f"Proses {index:02d}",
                    "operator_allocations_by_line": {
                        line: (index % 3) + line_index
                        for line_index, line in enumerate(self.labels)
                    },
                }
                for index in range(1, 20)
            ],
        }

    def test_four_lines_are_each_a_separate_series_column(self):
        result = _build_chart_data(self.data)
        self.assertEqual(result["labels"], self.labels)
        self.assertEqual(len(result["summary_matrix"][0]), 5)
        self.assertEqual(len(result["machine_matrix"][0]), 5)
        self.assertEqual(len(result["process_matrix"][0]), 5)

    def test_summary_chart_only_contains_same_unit_output_rates(self):
        result = _build_chart_data(self.data)
        self.assertEqual(
            [row[0] for row in result["summary_matrix"]],
            ["OUTPUT PER JAM (PCS/JAM)", "Target Output / Jam", "Potensial Output / Jam", "Aktual Output / Jam"],
        )
        self.assertNotIn("Jumlah Operator", [row[0] for row in result["summary_matrix"]])
        self.assertNotIn("Takt Time", [row[0] for row in result["summary_matrix"]])

    def test_total_machine_is_not_a_machine_type_in_chart(self):
        result = _build_chart_data(self.data)
        categories = [row[0] for row in result["machine_matrix"][1:]]
        self.assertEqual(categories, ["DNL-150", "DSC-246"])
        self.assertNotIn("TOTAL MESIN", categories)

    def test_process_chart_is_top_fifteen_while_full_table_count_is_retained(self):
        result = _build_chart_data(self.data)
        self.assertEqual(result["process_count"], 19)
        self.assertEqual(result["process_chart_count"], 15)
        self.assertEqual(len(result["process_matrix"]), 16)


if __name__ == "__main__":
    unittest.main()
