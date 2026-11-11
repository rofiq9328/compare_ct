import unittest

from excel.chart_generator import (
    _build_chart_data,
    _build_series_formula,
    _build_table_row_plan,
    _chart_row_bounds,
    _position_chart,
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

    def test_process_chart_keeps_the_same_order_as_the_report_table(self):
        rows = self.data["process_rows"]
        rows[0]["process"] = "Proses PERTAMA"
        rows[-1]["process"] = "Proses TERAKHIR"

        result = _build_chart_data(self.data)

        self.assertEqual(result["process_matrix"][1][0], "Proses PERTAMA")
        self.assertEqual(result["process_matrix"][-1][0], "Proses TERAKHIR")
        self.assertEqual(result["process_count"], len(rows))

    def test_each_variable_uses_all_compared_line_series(self):
        result = _build_chart_data(self.data)

        self.assertEqual(len(result["labels"]), 4)
        for matrix_name in ("summary_matrix", "machine_matrix", "process_matrix"):
            matrix = result[matrix_name]
            for variable_row in matrix[1:]:
                self.assertEqual(
                    len(variable_row) - 1,
                    len(result["labels"]),
                    msg=f"{matrix_name}: {variable_row[0]}",
                )

    def test_series_formula_uses_one_category_cell_per_variable(self):
        class Worksheet:
            Name = "compare line"

        formula = _build_series_formula(
            Worksheet(),
            header_row=16,
            value_column=4,
            category_column=2,
            data_rows=(17, 19, 21),
            series_order=2,
        )

        self.assertEqual(
            formula,
            "=SERIES('compare line'!$D$16,"
            "('compare line'!$B$17,'compare line'!$B$19,'compare line'!$B$21),"
            "('compare line'!$D$17,'compare line'!$D$19,'compare line'!$D$21),2)",
        )

    def test_series_formula_escapes_sheet_apostrophes(self):
        class Worksheet:
            Name = "Line's Report"

        formula = _build_series_formula(
            Worksheet(), 3, 27, 2, (4, 6), 1
        )
        self.assertIn("'Line''s Report'!$AA$3", formula)
        self.assertIn("'Line''s Report'!$B$4", formula)
        self.assertIn("'Line''s Report'!$AA$6", formula)

    def test_chart_bounds_cover_full_two_row_blocks(self):
        self.assertEqual(_chart_row_bounds((4, 6, 8, 10, 12)), (4, 13))
        self.assertEqual(_chart_row_bounds((17, 19, 21)), (17, 22))
        self.assertEqual(_chart_row_bounds((40, 42, 44)), (40, 45))

    def test_chart_bounds_reject_gaps_and_empty_rows(self):
        with self.assertRaises(Exception):
            _chart_row_bounds(())
        with self.assertRaises(Exception):
            _chart_row_bounds((17, 19, 22))

    def test_chart_is_sized_to_exact_table_range_and_two_chart_columns(self):
        class Cell:
            def __init__(self, row, column):
                self.Left = column * 50.0
                self.Top = row * 20.0
                self.Bottom = (row + 1) * 20.0

        class Range:
            def __init__(self, start, end):
                self.Left = start.Left
                self.Top = start.Top
                self.Width = (end.Left + 50.0) - start.Left
                self.Height = end.Bottom - start.Top

        class Worksheet:
            def Cells(self, row, column):
                return Cell(row, column)

            def Range(self, start, end):
                return Range(start, end)

        class ChartObject:
            def __init__(self):
                self.Left = 0.0
                self.Top = 0.0
                self.Width = 500.0
                self.Height = 275.0
                self.Placement = 1

        chart = ChartObject()
        height = _position_chart(
            chart, Worksheet(), chart_col=7, first_row=40, last_row=221
        )
        self.assertEqual(chart.Left, 350.0)
        self.assertEqual(chart.Top, 800.0)
        self.assertEqual(chart.Width, 100.0)
        self.assertEqual(chart.Height, 3640.0)
        self.assertEqual(height, 3640.0)
        self.assertEqual(chart.Placement, 2)

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
        # Every process occupies a two-row merged block. The chart reads
        # the first row of each block, while process_end_row is the second
        # row of the final block.
        self.assertEqual(
            plan["processes"][-1],
            layout["process_end_row"] - 1,
        )
        self.assertEqual(len(plan["processes"]), 19)
        self.assertTrue(
            all(
                row == layout["process_start_row"] + 2 * index
                for index, row in enumerate(plan["processes"])
            )
        )

    def test_chart_ranges_match_machine_and_operator_table_extents(self):
        layout = calculate_excel_layout(
            machine_count=9,
            process_count=91,
            line_count=2,
        )
        plan = _build_table_row_plan(layout, machine_count=9, process_count=91)

        self.assertEqual(_chart_row_bounds(plan["summary"]), (4, 13))
        self.assertEqual(_chart_row_bounds(plan["machines"]), (17, 36))
        self.assertEqual(_chart_row_bounds(plan["processes"]), (40, 221))
        self.assertEqual(len(plan["machines"]), 10)  # 9 machine types + total
        self.assertEqual(len(plan["processes"]), 91)


if __name__ == "__main__":
    unittest.main()
