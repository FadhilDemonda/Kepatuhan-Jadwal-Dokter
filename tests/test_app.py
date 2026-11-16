import unittest
from contextlib import contextmanager
from types import SimpleNamespace
from unittest.mock import patch

import pandas as pd
from streamlit.testing.v1 import AppTest

from recap import RECAP_COLUMNS
from recap_fixtures import detail_frame, master_bytes, query_bytes, query_frame, DOCTOR


@contextmanager
def query_environment(frame=None, with_master=True):
    query = frame if frame is not None else query_frame()
    query_content = query_bytes(query)
    master_content = master_bytes() if with_master else None
    state = {
        "uploads": {
            "query_upload": SimpleNamespace(name="hasil-query.csv", getvalue=lambda: query_content),
            "master_upload": SimpleNamespace(name="master.xlsx", getvalue=lambda: master_content) if with_master else None,
        },
        "default_master": None,
        "destination": pd.DataFrame(columns=RECAP_COLUMNS),
    }

    def append_rows(worksheet, frame):
        state["destination"] = pd.concat([state["destination"], frame], ignore_index=True)

    with (
        patch("streamlit.file_uploader", side_effect=lambda *args, **kwargs: state["uploads"].get(kwargs.get("key"))),
        patch("recap.default_schedule", side_effect=lambda: state["default_master"]),
        patch("sheets_client.configured_client", return_value=(object(), "https://docs.google.com/spreadsheets/d/test/edit")) as client,
        patch("sheets_client.read_destination", side_effect=lambda *args: state["destination"].copy()),
        patch("sheets_client.append_rows", side_effect=append_rows) as append,
        patch("sheets_client.write_header") as header,
    ):
        yield state, client, append, header


class ImportAppTest(unittest.TestCase):
    def test_query_upload_generates_recap_before_confirming_import(self):
        with query_environment() as (state, client, append, header):
            app = AppTest.from_file("app.py").run(timeout=30)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(app.metric[0].value, "1")
            self.assertEqual(list(app.session_state.source.columns), RECAP_COLUMNS)
            self.assertTrue(app.session_state.recap_download.startswith(b"PK"))
            append.assert_not_called()
            header.assert_not_called()

            app.button[0].click().run(timeout=30)

            self.assertEqual(len(app.exception), 0)
            self.assertEqual(app.metric[0].value, "0")
            self.assertEqual(app.metric[1].value, "1")
            append.assert_called_once()
            header.assert_called_once()
            self.assertEqual(len(state["destination"]), 1)

    def test_missing_master_does_not_connect_or_write_to_sheets(self):
        with query_environment(with_master=False) as (state, client, append, header):
            app = AppTest.from_file("app.py").run(timeout=30)
            self.assertEqual(len(app.exception), 0)
            self.assertTrue(any("Lengkapi master HFIS" in message.value for message in app.warning))
            client.assert_not_called()
            append.assert_not_called()

    def test_uploaded_master_overrides_default_and_rebuilds_preview(self):
        with query_environment(with_master=False) as (state, client, append, header):
            state["default_master"] = ("default-master.xlsx", master_bytes())
            app = AppTest.from_file("app.py").run(timeout=30)
            self.assertEqual(app.session_state.source.iloc[0]["SENIN KATEGORI"], "Sesuai")
            original_hash = app.session_state.file_hash

            replacement = master_bytes(detail=detail_frame([[DOCTOR, "SENIN", "11:00 - 14:00", ""]]))
            state["uploads"]["master_upload"] = SimpleNamespace(name="pengganti.xlsx", getvalue=lambda: replacement)
            app.run(timeout=30)

            self.assertEqual(len(app.exception), 0)
            self.assertNotEqual(app.session_state.file_hash, original_hash)
            self.assertEqual(app.session_state.source.iloc[0]["SENIN KATEGORI"], "Tidak Sesuai")
            self.assertTrue(any("pengganti.xlsx" in caption.value for caption in app.caption))
            append.assert_not_called()

    def test_similar_name_requires_checkbox_before_import(self):
        query = query_frame(**{"Nama Dokter": "dr. Dokter Contohh, Sp.PD"})
        with query_environment(frame=query) as (state, client, append, header):
            app = AppTest.from_file("app.py").run(timeout=30)
            self.assertEqual(len(app.exception), 0)
            self.assertTrue(app.button[0].disabled)
            append.assert_not_called()

            app.checkbox[0].check().run(timeout=30)
            self.assertFalse(app.button[0].disabled)
            app.button[0].click().run(timeout=30)
            append.assert_called_once()

    def test_invalid_query_file_stops_before_sheets_access(self):
        with query_environment() as (state, client, append, header):
            state["uploads"]["query_upload"] = SimpleNamespace(name="rusak.csv", getvalue=lambda: b"invalid")
            app = AppTest.from_file("app.py").run(timeout=30)
            self.assertEqual(len(app.exception), 0)
            self.assertGreater(len(app.error), 0)
            client.assert_not_called()
            append.assert_not_called()

    def test_changing_master_resets_mapping_confirmation(self):
        query = query_frame(**{"Nama Dokter": "dr. Dokter Contohh, Sp.PD"})
        with query_environment(frame=query) as (state, client, append, header):
            app = AppTest.from_file("app.py").run(timeout=30)
            app.checkbox[0].check().run(timeout=30)
            self.assertFalse(app.button[0].disabled)

            replacement = master_bytes(detail=detail_frame([[DOCTOR, "SENIN", "11:00 - 14:00", ""]]))
            state["uploads"]["master_upload"] = SimpleNamespace(name="pengganti.xlsx", getvalue=lambda: replacement)
            app.run(timeout=30)

            self.assertFalse(app.checkbox[0].value)
            self.assertTrue(app.button[0].disabled)
            append.assert_not_called()

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
            patch("streamlit.radio", return_value="Rekap siap impor"),
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
