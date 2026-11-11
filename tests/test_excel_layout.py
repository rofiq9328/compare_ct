import unittest

from excel.layout import calculate_excel_layout


class TestDynamicExcelLayout(unittest.TestCase):
    def test_line_columns_expand_without_overlapping_chart_and_note(self):
        for line_count in (2, 3, 4, 5, 8, 12):
            with self.subTest(line_count=line_count):
                layout = calculate_excel_layout(
                    machine_count=9,
                    process_count=27,
                    line_count=line_count,
                    chart_process_count=15,
                )
                self.assertEqual(layout["chart_col"], 3 + line_count)
                self.assertEqual(layout["note_col"], layout["chart_col"] + 2)
                self.assertEqual(layout["helper_col"], layout["note_col"] + 3)
                self.assertGreater(layout["chart_col"], 2 + line_count)

    def test_machine_and_process_blocks_expand_dynamically(self):
        cases = (
            (0, 0),
            (6, 10),
            (7, 11),
            (12, 40),
        )
        for machine_count, process_count in cases:
            with self.subTest(machine_count=machine_count, process_count=process_count):
                layout = calculate_excel_layout(
                    machine_count, process_count, line_count=4
                )
                expected_total = 17 + 2 * machine_count
                expected_process_start = expected_total + 5
                self.assertEqual(layout["machine_total_row"], expected_total)
                self.assertEqual(layout["section3_row"], expected_total + 3)
                self.assertEqual(layout["section3_header_row"], expected_total + 4)
                self.assertEqual(layout["process_start_row"], expected_process_start)
                self.assertEqual(
                    layout["process_end_row"],
                    expected_process_start + 2 * max(1, process_count) - 1,
                )

    def test_chart_uses_top_fifteen_but_table_keeps_all_process_rows(self):
        layout = calculate_excel_layout(
            machine_count=8,
            process_count=34,
            line_count=6,
            chart_process_count=15,
        )
        self.assertEqual(layout["allocated_process_count"], 34)
        self.assertEqual(
            layout["process_chart_end_row"],
            layout["process_start_row"] + 29,
        )
        self.assertLess(layout["process_chart_end_row"], layout["process_end_row"])

    def test_rejects_invalid_line_count(self):
        with self.assertRaises(ValueError):
            calculate_excel_layout(5, 12, 1)


if __name__ == "__main__":
    unittest.main()
