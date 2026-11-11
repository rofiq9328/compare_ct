import unittest

from excel.chart_generator import (
    _build_chart_data,
    _build_table_row_plan,
    _worksheet_cell_reference,
)
from excel.layout import calculate_excel_layout


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

    def test_summary_chart_keeps_original_template_categories(self):
        result = _build_chart_data(self.data)
        self.assertEqual(
            [row[0] for row in result["summary_matrix"]],
            ["VARIABEL", "Jumlah Operator", "Takt Time", "Target Output / Jam",
             "Potensial Output / Jam", "Aktual Output / Jam"],
        )

    def test_machine_chart_keeps_total_row_from_original_template(self):
        result = _build_chart_data(self.data)
        categories = [row[0] for row in result["machine_matrix"][1:]]
        self.assertEqual(categories, ["DNL-150", "DSC-246", "TOTAL MESIN"])

    def test_process_chart_includes_every_process_row(self):
        result = _build_chart_data(self.data)
        self.assertEqual(result["process_count"], 19)
        self.assertEqual(result["process_chart_count"], 19)
        self.assertEqual(len(result["process_matrix"]), 20)

    def test_chart_sources_use_visible_table_row_positions(self):
        layout = calculate_excel_layout(
            machine_count=2,
            process_count=19,
            line_count=4,
        )
        plan = _build_table_row_plan(layout, machine_count=2, process_count=19)

        self.assertEqual(plan["summary"], (4, 6, 8, 10, 12))
        self.assertEqual(plan["machines"], (17, 19, layout["machine_total_row"]))
        self.assertEqual(plan["processes"][0], layout["process_start_row"])
        self.assertEqual(plan["processes"][-1], layout["process_end_row"])
        self.assertEqual(len(plan["processes"]), 19)


class TestChartWorksheetReferences(unittest.TestCase):
    def test_series_name_uses_a_live_header_cell_reference(self):
        class Worksheet:
            Name = "compare line"

        self.assertEqual(
            _worksheet_cell_reference(Worksheet(), row=3, column=4),
            "='compare line'!R3C4",
        )

    def test_series_name_reference_escapes_sheet_apostrophes(self):
        class Worksheet:
            Name = "Line's Report"

        self.assertEqual(
            _worksheet_cell_reference(Worksheet(), row=16, column=27),
            "='Line''s Report'!R16C27",
        )

    def test_series_name_reference_rejects_invalid_coordinates(self):
        class Worksheet:
            Name = "compare line"

        with self.assertRaises(ValueError):
            _worksheet_cell_reference(Worksheet(), row=0, column=3)
        with self.assertRaises(ValueError):
            _worksheet_cell_reference(Worksheet(), row=3, column=0)


if __name__ == "__main__":
    unittest.main()
