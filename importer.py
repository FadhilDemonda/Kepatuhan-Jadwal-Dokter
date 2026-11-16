"""Validation, normalization, and comparison rules for weekly imports."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time
from io import BytesIO
import math
import re
from typing import Any

import pandas as pd


SOURCE_SHEET = "Kepatuhan Jadwal Praktek Dokter"
KEY_COLUMNS = ("AWAL MINGGU (SENIN)", "DOKTER")


class ImportValidationError(ValueError):
    """Raised when an uploaded workbook cannot safely be imported."""


@dataclass(frozen=True)
class ImportComparison:
    new_rows: pd.DataFrame
    existing_count: int
    conflicts: pd.DataFrame


def _clean_header(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value).strip())


def _normalise_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    return re.sub(r"\s+", " ", str(value).strip()).casefold()


def _date_to_iso(value: Any) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)) or value == "":
        return ""
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return ""
    return parsed.date().isoformat()


def _time_to_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime("%H:%M:%S")
    if isinstance(value, datetime):
        return value.strftime("%H:%M:%S")
    if isinstance(value, time):
        return value.strftime("%H:%M:%S")
    if isinstance(value, (float, int)) and not isinstance(value, bool):
        if 0 <= float(value) < 1:
            seconds = round(float(value) * 24 * 60 * 60)
            seconds %= 24 * 60 * 60
            return f"{seconds // 3600:02d}:{(seconds % 3600) // 60:02d}:{seconds % 60:02d}"
    text = str(value).strip()
    clock = re.fullmatch(r"(\d{1,2}):(\d{2})(?::(\d{2}))?", text)
    if clock:
        hours, minutes, seconds = (int(part or 0) for part in clock.groups())
        if hours < 24 and minutes < 60 and seconds < 60:
            return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
    return text


def serialise_value(column: str, value: Any) -> str:
    """Return the safe, stable Sheets representation for one source value."""
    if column == KEY_COLUMNS[0]:
        return _date_to_iso(value)
    if column.endswith("TASK 4"):
        return _time_to_text(value)
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return _date_to_iso(value)
    return str(value).strip()


def read_workbook(uploaded_bytes: bytes) -> pd.DataFrame:
    """Read the expected source tab and reject any changed column layout."""
    try:
        workbook = pd.ExcelFile(BytesIO(uploaded_bytes), engine="openpyxl")
    except Exception as exc:  # pandas exposes several implementation errors here
        raise ImportValidationError("File tidak dapat dibaca sebagai Excel .xlsx.") from exc

    if SOURCE_SHEET not in workbook.sheet_names:
        raise ImportValidationError(f"Tab sumber '{SOURCE_SHEET}' tidak ditemukan.")

    frame = pd.read_excel(workbook, sheet_name=SOURCE_SHEET, dtype=object)
    frame.columns = [_clean_header(column) for column in frame.columns]
    if frame.columns.duplicated().any():
        duplicates = frame.columns[frame.columns.duplicated()].tolist()
        raise ImportValidationError(f"Header ganda ditemukan: {', '.join(duplicates)}.")
    if not set(KEY_COLUMNS).issubset(frame.columns):
        raise ImportValidationError("Header wajib 'AWAL MINGGU (SENIN)' dan 'DOKTER' tidak lengkap.")
    if frame.empty:
        raise ImportValidationError("Tab sumber tidak memiliki baris data.")

    blank_rows = frame.apply(lambda row: all(_normalise_text(value) == "" for value in row), axis=1)
    return frame.loc[~blank_rows].reset_index(drop=True)


def validate_and_serialise(frame: pd.DataFrame) -> pd.DataFrame:
    """Validate identity keys and prepare raw values for a Google Sheets append."""
    output = frame.copy()
    for column in output.columns:
        output[column] = output[column].map(lambda value: serialise_value(column, value))

    invalid_date = output[KEY_COLUMNS[0]].eq("")
    invalid_doctor = output[KEY_COLUMNS[1]].map(_normalise_text).eq("")
    if invalid_date.any() or invalid_doctor.any():
        rows = output.index[invalid_date | invalid_doctor].tolist()
        display = ", ".join(str(index + 2) for index in rows[:10])
        raise ImportValidationError(f"Tanggal awal minggu atau dokter kosong/tidak valid pada baris Excel: {display}.")

    keys = output.apply(key_for_row, axis=1)
    duplicates = keys.duplicated(keep=False)
    if duplicates.any():
        rows = output.index[duplicates].tolist()
        display = ", ".join(str(index + 2) for index in rows[:10])
        raise ImportValidationError(f"Kunci ganda (awal minggu + dokter) di Excel pada baris: {display}.")
    return output


def key_for_row(row: pd.Series) -> tuple[str, str]:
    return (_date_to_iso(row[KEY_COLUMNS[0]]), _normalise_text(row[KEY_COLUMNS[1]]))


def compare_frames(source: pd.DataFrame, destination: pd.DataFrame) -> ImportComparison:
    """Classify source records as new, identical, or conflicting with destination."""
    if list(source.columns) != list(destination.columns):
        raise ImportValidationError("Header Google Sheets tidak sama dengan header Excel. Tidak ada data yang diubah.")

    destination = destination.copy()
    destination.columns = [_clean_header(column) for column in destination.columns]
    destination_keys = destination.apply(key_for_row, axis=1)
    invalid_keys = [key for key in destination_keys if not key[0] or not key[1]]
    if invalid_keys:
        raise ImportValidationError("Google Sheets memiliki tanggal awal minggu atau dokter yang kosong. Rapikan data tujuan sebelum impor.")
    if destination_keys.duplicated(keep=False).any():
        raise ImportValidationError("Google Sheets memiliki kunci ganda (awal minggu + dokter). Rapikan data tujuan sebelum impor.")

    destination_by_key = {key: destination.iloc[index] for index, key in enumerate(destination_keys)}
    new_indexes: list[int] = []
    conflicts: list[dict[str, str]] = []
    existing_count = 0
    for index, row in source.iterrows():
        key = key_for_row(row)
        destination_row = destination_by_key.get(key)
        if destination_row is None:
            new_indexes.append(index)
        elif all(
            _normalise_text(serialise_value(column, row[column]))
            == _normalise_text(serialise_value(column, destination_row[column]))
            for column in source.columns
        ):
            existing_count += 1
        else:
            conflicts.append({
                "Awal minggu": row[KEY_COLUMNS[0]],
                "Dokter": row[KEY_COLUMNS[1]],
                "Status": "Nilai berbeda dengan data yang sudah tersimpan",
            })

    return ImportComparison(
        new_rows=source.loc[new_indexes].reset_index(drop=True),
        existing_count=existing_count,
        conflicts=pd.DataFrame(conflicts),
    )


def date_range_label(frame: pd.DataFrame) -> str:
    dates = pd.to_datetime(frame[KEY_COLUMNS[0]], errors="coerce")
    if dates.isna().any():
        return "Tanggal tidak valid"
    return f"{dates.min():%d %b %Y} s.d. {dates.max():%d %b %Y}"
