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


if __name__ == "__main__":
    unittest.main()
