import unittest

import pandas as pd

from importer import ImportValidationError, compare_frames, validate_and_serialise


HEADERS = ["MINGGU KE", "AWAL MINGGU (SENIN)", "DOKTER", "SENIN TASK 4"]


def source_frame():
    return pd.DataFrame([[45, "2026-11-09", "dr. A", 0.5]], columns=HEADERS)


class ImporterTest(unittest.TestCase):
    def test_serialises_date_and_excel_time(self):
        result = validate_and_serialise(source_frame())
        self.assertEqual(
            result.iloc[0].to_dict(),
            {
                "MINGGU KE": "45",
                "AWAL MINGGU (SENIN)": "2026-11-09",
                "DOKTER": "dr. A",
                "SENIN TASK 4": "12:00:00",
            },
        )

    def test_compares_existing_and_new_by_date_and_doctor(self):
        source = validate_and_serialise(source_frame())
        destination = pd.DataFrame([["44", "2026-11-02", "dr. A", "12:00:00"]], columns=HEADERS)
        comparison = compare_frames(source, destination)
        self.assertEqual(len(comparison.new_rows), 1)
        self.assertEqual(comparison.existing_count, 0)

    def test_detects_changed_existing_row_as_conflict(self):
        source = validate_and_serialise(source_frame())
        destination = pd.DataFrame([["45", "2026-11-09", "dr. A", "13:00:00"]], columns=HEADERS)
        comparison = compare_frames(source, destination)
        self.assertTrue(comparison.new_rows.empty)
        self.assertEqual(len(comparison.conflicts), 1)

    def test_rejects_duplicate_keys(self):
        frame = pd.concat([source_frame(), source_frame()], ignore_index=True)
        with self.assertRaisesRegex(ImportValidationError, "Kunci ganda"):
            validate_and_serialise(frame)

    def test_sheets_time_display_format_is_not_a_conflict(self):
        for excel_time, sheets_time in [
            ("12:00:00", "12:00"),
            ("09:05:00", "9:05"),
            ("09:05:30", "9:05:30"),
        ]:
            with self.subTest(excel_time=excel_time, sheets_time=sheets_time):
                source = source_frame().astype(object)
                source.loc[0, "SENIN TASK 4"] = excel_time
                source = validate_and_serialise(source)
                destination = source.copy()
                destination.loc[0, "SENIN TASK 4"] = sheets_time
                comparison = compare_frames(source, destination)
                self.assertEqual(comparison.existing_count, 1)
                self.assertTrue(comparison.conflicts.empty)
                self.assertTrue(comparison.new_rows.empty)

    def test_different_seconds_remain_a_conflict(self):
        source = validate_and_serialise(source_frame())
        destination = source.copy()
        destination.loc[0, "SENIN TASK 4"] = "12:00:01"
        comparison = compare_frames(source, destination)
        self.assertEqual(len(comparison.conflicts), 1)
        self.assertEqual(comparison.existing_count, 0)

    def test_missing_time_marker_is_not_the_same_as_blank(self):
        source = validate_and_serialise(source_frame())
        source.loc[0, "SENIN TASK 4"] = "-"
        destination = source.copy()
        destination.loc[0, "SENIN TASK 4"] = ""
        comparison = compare_frames(source, destination)
        self.assertEqual(len(comparison.conflicts), 1)


if __name__ == "__main__":
    unittest.main()
