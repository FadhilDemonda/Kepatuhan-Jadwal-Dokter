from datetime import date
from io import BytesIO
import unittest

from openpyxl import Workbook, load_workbook
import pandas as pd

from importer import ImportValidationError, SOURCE_SHEET, read_workbook, validate_and_serialise
from recap import RECAP_COLUMNS, ScheduleBook, build_recap, read_query_export, read_schedule_book, recap_excel
from recap_fixtures import DOCTOR, detail_frame, history_frame, master_bytes, query_bytes, query_frame


class RecapTest(unittest.TestCase):
    def test_reads_csv_delimiters_and_excel(self):
        for delimiter in (";", ",", "\t"):
            with self.subTest(delimiter=delimiter):
                frame = read_query_export(query_bytes(delimiter=delimiter), "export.csv")
                self.assertEqual(frame.iloc[0]["Nama Dokter"], DOCTOR)
                self.assertEqual(frame.iloc[0]["SELASA"], "")
        frame = read_query_export(query_bytes(extension="xlsx"), "export.xlsx")
        self.assertEqual(frame.iloc[0]["SENIN"], "9:05")

    def test_rejects_patient_level_export_with_wrong_schema(self):
        with self.assertRaisesRegex(ImportValidationError, "Kolom hasil query mingguan"):
            read_query_export(b'ID REGISTRASI;DOKTER\n1;contoh\n', "export.csv")

    def test_duplicate_headers_are_rejected_for_csv_and_excel(self):
        workbook = Workbook()
        workbook.active.append(["Minggu Ke", "SENIN", "SENIN"])
        workbook.active.append([45, "9:00", "10:00"])
        output = BytesIO()
        workbook.save(output)
        for filename, content in [("export.csv", b'Minggu Ke;SENIN;SENIN\n45;9:00;10:00\n'), ("export.xlsx", output.getvalue())]:
            with self.subTest(filename=filename):
                with self.assertRaisesRegex(ImportValidationError, "Header ganda"):
                    read_query_export(content, filename)

    def test_requires_both_master_tabs(self):
        with self.assertRaisesRegex(ImportValidationError, "Detail HFIS dan Riwayat Perubahan"):
            read_schedule_book(query_bytes(extension="xlsx"))

    def test_generates_template_columns_and_categories(self):
        result = build_recap(query_frame(), read_schedule_book(master_bytes()))
        self.assertEqual(list(result.data.columns), RECAP_COLUMNS)
        self.assertEqual(len(result.data.columns), 26)
        row = result.data.iloc[0]
        self.assertEqual(row["AWAL MINGGU (SENIN)"], "2026-11-09")
        self.assertEqual(row["SENIN JADWAL"], "09:00 - 12:00")
        self.assertEqual(row["SENIN TASK 4"], "09:05:00")
        self.assertEqual(row["SENIN KATEGORI"], "Sesuai")
        self.assertEqual(row["KATEGORI REAL MINGGUAN"], "Sesuai")
        self.assertEqual(row["SELASA KATEGORI"], "")

    def test_assessment_window_and_missing_task(self):
        book = read_schedule_book(master_bytes())
        for task, expected in [("08:00", "Sesuai"), ("07:59", "Tidak Sesuai"), ("09:59:59", "Sesuai"), ("10:00", "Tidak Sesuai"), ("-", "Tidak Sesuai")]:
            with self.subTest(task=task):
                result = build_recap(query_frame(SENIN=task), book)
                self.assertEqual(result.data.iloc[0]["SENIN KATEGORI"], expected)

    def test_no_schedule_does_not_overrule_a_compliant_day(self):
        result = build_recap(query_frame(SELASA="9:05"), read_schedule_book(master_bytes()))
        self.assertEqual(result.data.iloc[0]["SELASA KATEGORI"], "Tidak Ada Jadwal")
        self.assertEqual(result.data.iloc[0]["KATEGORI REAL MINGGUAN"], "Sesuai")

    def test_uses_history_before_and_after_effective_date(self):
        history = history_frame([[
            "2026-11-09", DOCTOR, DOCTOR, "SENIN", "Perubahan jadwal praktik", "Ya",
            "08:00 - 11:00", "09:00 - 12:00", "Dokumen contoh", "",
        ]])
        book = ScheduleBook(detail_frame(), history)
        self.assertEqual(book.sessions(DOCTOR, "PENYAKIT DALAM", "SENIN", date(2026, 11, 2))[0][1], "08:00 - 11:00")
        self.assertEqual(book.sessions(DOCTOR, "PENYAKIT DALAM", "SENIN", date(2026, 11, 9))[0][1], "09:00 - 12:00")

    def test_added_doctor_has_no_schedule_before_effective_date(self):
        history = history_frame([[
            "2026-11-09", DOCTOR, DOCTOR, "SENIN", "Penambahan dokter", "Ya",
            "-", "09:00 - 12:00", "Dokumen contoh", "",
        ]])
        book = ScheduleBook(detail_frame(), history)
        self.assertEqual(book.sessions(DOCTOR, "PENYAKIT DALAM", "SENIN", date(2026, 11, 2)), [])

    def test_removed_doctor_has_no_schedule_after_last_practice(self):
        history = history_frame([[
            "2026-11-10", DOCTOR, "(tidak ditemukan di HFIS)", "SENIN", "Penghapusan jadwal dokter", "Ya",
            "", "", "Dokumen contoh", "Praktik terakhir 9 November 2026",
        ]])
        book = ScheduleBook(detail_frame(), history)
        self.assertTrue(book.sessions(DOCTOR, "PENYAKIT DALAM", "SENIN", date(2026, 11, 9)))
        self.assertEqual(book.sessions(DOCTOR, "PENYAKIT DALAM", "SENIN", date(2026, 11, 16)), [])

    def test_saturday_rotation(self):
        book = ScheduleBook(detail_frame([[DOCTOR, "SABTU", "09:00 - 12:00", "MINGGU KE-1 DAN KE-3"]]), history_frame())
        self.assertTrue(book.sessions(DOCTOR, "PENYAKIT DALAM", "SABTU", date(2026, 11, 7)))
        self.assertEqual(book.sessions(DOCTOR, "PENYAKIT DALAM", "SABTU", date(2026, 11, 14)), [])

    def test_hemodialysis_and_clinic_sessions_stay_separate(self):
        name = "DR ILHAM AHMADI, Sp.PD"
        book = ScheduleBook(detail_frame([
            [name, "SENIN", "09:00 - 12:00", "JADWAL POLI"],
            [name, "SENIN", "17:00 - 20:00", "HEMODIALISA"],
        ]), history_frame())
        self.assertEqual(book.sessions(name, "PENYAKIT DALAM", "SENIN", date(2026, 11, 9))[0][1], "09:00 - 12:00")
        self.assertEqual(book.sessions(name, "HEMODIALISA", "SENIN", date(2026, 11, 9))[0][1], "17:00 - 20:00")

    def test_similar_name_requires_review_and_unknown_name_is_rejected(self):
        book = read_schedule_book(master_bytes())
        result = build_recap(query_frame(**{"Nama Dokter": "dr. Dokter Contohh, Sp.PD"}), book)
        self.assertTrue(result.mapping.iloc[0]["Perlu konfirmasi"])
        with self.assertRaisesRegex(ImportValidationError, "belum dapat dipetakan"):
            build_recap(query_frame(**{"Nama Dokter": "dr. XYZ"}), book)

    def test_exclusions_are_reported_not_silently_lost(self):
        frame = pd.concat([query_frame(), query_frame(**{"Nama Dokter": "NURUL AZIZAH"})], ignore_index=True)
        result = build_recap(frame, read_schedule_book(master_bytes()))
        self.assertEqual(len(result.data), 1)
        self.assertEqual(len(result.excluded), 1)

    def test_invalid_time_period_and_duplicate_keys_block_generation(self):
        book = read_schedule_book(master_bytes())
        for frame in [query_frame(SENIN="99:00"), query_frame(Periode="10/11/2026 s/d 16/11/2026"), pd.concat([query_frame(), query_frame()], ignore_index=True)]:
            with self.subTest(frame=frame.index.tolist()):
                with self.assertRaises(ImportValidationError):
                    build_recap(frame, book)

    def test_excel_download_roundtrips_to_ready_import(self):
        result = build_recap(query_frame(), read_schedule_book(master_bytes()))
        content = recap_excel(result)
        workbook = load_workbook(BytesIO(content), read_only=True)
        self.assertEqual(list(next(workbook[SOURCE_SHEET].values)), RECAP_COLUMNS)
        restored = validate_and_serialise(read_workbook(content))
        pd.testing.assert_frame_equal(restored, result.data)


if __name__ == "__main__":
    unittest.main()
