import unittest
from types import SimpleNamespace
from unittest.mock import patch

import pandas as pd
from streamlit.testing.v1 import AppTest


class ImportAppTest(unittest.TestCase):
    def test_success_refreshes_preview_without_repeating_append(self):
        headers = ["MINGGU KE", "AWAL MINGGU (SENIN)", "DOKTER", "SENIN TASK 4"]
        source = pd.DataFrame([[45, "2026-11-09", "dr. A", 0.5]], columns=headers)
        destination = pd.DataFrame([["44", "2026-11-02", "dr. A", "12:00:00"]], columns=headers)
        upload = SimpleNamespace(name="rekap.xlsx", getvalue=lambda: b"test-workbook")

        def read_destination(*args):
            return destination.copy()

        def append_rows(worksheet, frame):
            nonlocal destination
            destination = pd.concat([destination, frame], ignore_index=True)

        with (
            patch("streamlit.file_uploader", return_value=upload),
            patch("importer.read_workbook", return_value=source),
            patch("sheets_client.configured_client", return_value=(object(), "https://docs.google.com/spreadsheets/d/test/edit")),
            patch("sheets_client.read_destination", side_effect=read_destination),
            patch("sheets_client.append_rows", side_effect=append_rows) as append,
            patch("sheets_client.write_header") as header,
        ):
            app = AppTest.from_file("app.py").run(timeout=30)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(app.metric[0].value, "1")

            app.button[0].click().run(timeout=30)

            self.assertEqual(len(app.exception), 0)
            self.assertEqual(app.metric[0].value, "0")
            self.assertEqual(app.metric[1].value, "1")
            self.assertTrue(any("1 baris berhasil ditambahkan dan diverifikasi" in message.value for message in app.success))
            append.assert_called_once()
            header.assert_not_called()

            app.run(timeout=30)
            self.assertEqual(len(app.button), 0)
            append.assert_called_once()
            self.assertEqual(len(destination), 2)


if __name__ == "__main__":
    unittest.main()
