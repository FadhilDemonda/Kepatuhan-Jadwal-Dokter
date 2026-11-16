"""Google Sheets boundary. Credentials are read from Streamlit secrets only."""

from __future__ import annotations

from typing import Any

import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

from importer import ImportValidationError


SHEETS_SCOPE = ["https://www.googleapis.com/auth/spreadsheets"]


def configured_client(secrets: Any) -> tuple[gspread.Worksheet, str]:
    try:
        spreadsheet_id = secrets["google_sheets"]["spreadsheet_id"]
        worksheet_name = secrets["google_sheets"]["worksheet_name"]
        service_account_info = dict(secrets["google_service_account"])
    except (KeyError, TypeError) as exc:
        raise ImportValidationError("Konfigurasi Google Sheets belum lengkap. Isi secrets.toml terlebih dahulu.") from exc

    try:
        credentials = Credentials.from_service_account_info(service_account_info, scopes=SHEETS_SCOPE)
        client = gspread.authorize(credentials)
        spreadsheet = client.open_by_key(spreadsheet_id)
        return spreadsheet.worksheet(worksheet_name), spreadsheet.url
    except Exception as exc:
        raise ImportValidationError(
            "Google Sheets tidak dapat diakses. Periksa spreadsheet ID, nama tab, dan izin Editor untuk service account."
        ) from exc


def read_destination(worksheet: gspread.Worksheet, expected_headers: list[str]) -> pd.DataFrame:
    values = worksheet.get_all_values()
    if not values:
        return pd.DataFrame(columns=expected_headers)
    headers = values[0]
    if headers != expected_headers:
        raise ImportValidationError("Header Google Sheets tidak sama dengan header Excel. Tidak ada data yang diubah.")
    return pd.DataFrame(values[1:], columns=headers)


def append_rows(worksheet: gspread.Worksheet, frame: pd.DataFrame) -> None:
    if frame.empty:
        return
    worksheet.append_rows(frame.values.tolist(), value_input_option="RAW")


def write_header(worksheet: gspread.Worksheet, headers: list[str]) -> None:
    worksheet.append_row(headers, value_input_option="RAW")
