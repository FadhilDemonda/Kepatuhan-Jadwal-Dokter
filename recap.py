"""Turn weekly query exports into the 26-column compliance recap."""

from __future__ import annotations

import csv
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from difflib import SequenceMatcher
from io import BytesIO, StringIO
from pathlib import Path
import re
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from importer import ImportValidationError, SOURCE_SHEET, validate_and_serialise


DAYS = ["SENIN", "SELASA", "RABU", "KAMIS", "JUMAT", "SABTU", "MINGGU"]
QUERY_COLUMNS = ["Minggu Ke", "Periode", "Nama Dokter", "Poli / Subspesialis", *DAYS]
RECAP_COLUMNS = [
    "MINGGU KE", "AWAL MINGGU (SENIN)", "DOKTER", "SPESIALIS", "KATEGORI REAL MINGGUAN",
    *[f"{day} {field}" for day in DAYS for field in ("JADWAL", "TASK 4", "KATEGORI")],
]
DETAIL_COLUMNS = ["Nama dan Gelar", "Hari", "Jadwal", "Keterangan"]
HISTORY_COLUMNS = [
    "Tanggal Dokumen", "Nama di Dokumen", "Nama di HFIS", "Hari", "Jenis", "Berubah?",
    "Sebelumnya", "Perubahan", "Dokumen", "Keterangan",
]
NO_SCHEDULE = "Tidak Ada Jadwal"
NOT_FOUND = "(tidak ditemukan di HFIS)"
DOC_ALIASES = {
    "drg Tri Hartanto W.P, SpBM": "drg. Tri Hartanto Wahyu P, Sp.BM",
    "dr. Muhammad Imanuddin Nasution, SpPD": "dr. Muhammad Imanuddin Nasution, Sp.PD",
}
NAME_OVERRIDES = {
    "drg. Alvin, Sp.BM": "drg. Alvin Djojohandoko, Sp.BM",
    "dr. Hamidah Fadhil, Sp.KFR": "dr. Hamidah, Sp.KFR",
}
EXCLUDED_NAMES = {
    **dict.fromkeys([
        "Alfian Yunus Lukito", "Dian Mustika S", "Ismi Devi Sulistyowati", "Mohammad Sjifaa",
        "NURUL AZIZAH", "Nadila Sustanti, S.Tr.Kes", "Septia Dwi Kartika Ratri, A.Md.Kes",
    ], "Terapis, tidak ada di jadwal HFIS"),
    "dr. Abdul Latip. Mars": "Tidak ditemukan di HFIS maupun dokumen perubahan",
    "dr. Herman Fikri, SpOG": "Dokumen penghapusan tidak memuat jam praktik",
    "dr. Muhammad Imanuddin Nasution, Sp.PD": "Dokumen hanya memuat jadwal libur nasional",
    "drg. Agusnidar": "Seluruh jadwal HFIS libur dan tidak pernah berubah",
    "drg. Rosalina Catarina Haberham": "Seluruh jadwal HFIS libur dan tidak pernah berubah",
}


def default_schedule() -> tuple[str, bytes] | None:
    path = Path(__file__).parent / "data" / "JADWAL DOKTER RS AN-NISA (HFIS) update 23-10-2026.xlsx"
    return (path.name, path.read_bytes()) if path.is_file() else None


@dataclass(frozen=True)
class RecapResult:
    data: pd.DataFrame
    mapping: pd.DataFrame
    excluded: pd.DataFrame


def _clean(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return " ".join(str(value).replace("\xa0", " ").split())


def _name_key(value: str) -> str:
    return _clean(value).casefold()


def _name_core(value: str) -> str:
    value = _name_key(value).split(",")[0]
    value = re.sub(r"\b(sp\.?\s?[a-z.\-]*|mm|m\.?kes|m\.?biomed|mars)\b.*", "", value)
    value = re.sub(r"^\s*(drg|dr)\.?\s*", "", value)
    value = re.sub(r"[^a-z ]", " ", value)
    return " ".join(re.sub(r"\b(dr|drg)\b", " ", value).split())


def _checked_columns(frame: pd.DataFrame, required: list[str], label: str) -> pd.DataFrame:
    frame = frame.copy()
    frame.columns = [_clean(column) for column in frame.columns]
    if frame.columns.duplicated().any():
        raise ImportValidationError(f"Header ganda di {label}.")
    missing = [column for column in required if column not in frame]
    if missing:
        raise ImportValidationError(f"Kolom {label} belum lengkap: {', '.join(missing)}.")
    frame = frame[required].copy()
    blank = frame.apply(lambda row: all(_clean(value) == "" for value in row), axis=1)
    return frame.loc[~blank].reset_index(drop=True)


def read_query_export(content: bytes, filename: str) -> pd.DataFrame:
    """Accept the weekly pivot query, not patient-level registration exports."""
    try:
        if filename.lower().endswith(".csv"):
            try:
                text = content.decode("utf-8-sig")
            except UnicodeDecodeError:
                text = content.decode("cp1252")
            dialect = csv.Sniffer().sniff(text[:16384], delimiters=",;\t")
            header = next(csv.reader(StringIO(text), dialect))
            cleaned = [_clean(column) for column in header]
            if len(cleaned) != len(set(cleaned)):
                raise ImportValidationError("Header ganda di hasil query.")
            frame = pd.read_csv(StringIO(text), sep=dialect.delimiter, dtype=object, keep_default_na=False)
        elif filename.lower().endswith(".xlsx"):
            with pd.ExcelFile(BytesIO(content), engine="openpyxl") as workbook:
                cleaned = [_clean(cell.value) for cell in next(workbook.book.worksheets[0].iter_rows(max_row=1))]
                nonempty_headers = [column for column in cleaned if column]
                if len(nonempty_headers) != len(set(nonempty_headers)):
                    raise ImportValidationError("Header ganda di hasil query.")
                frame = pd.read_excel(workbook, sheet_name=0, dtype=object, keep_default_na=False)
        else:
            raise ImportValidationError("Hasil query harus berupa CSV atau Excel .xlsx.")
    except ImportValidationError:
        raise
    except Exception as exc:
        raise ImportValidationError("Hasil query tidak dapat dibaca. Periksa format CSV/Excel dan baris header.") from exc
    frame = _checked_columns(frame, QUERY_COLUMNS, "hasil query mingguan")
    if frame.empty:
        raise ImportValidationError("Hasil query tidak memiliki baris data.")
    return frame


def _calendar_date(value: Any) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(_clean(value), fmt).date()
        except ValueError:
            pass
    raise ImportValidationError("Tanggal pada master HFIS tidak valid. Gunakan tanggal Excel atau YYYY-MM-DD / DD/MM/YYYY.")


def _schedule_parts(value: Any) -> list[tuple[int, str]]:
    text = _clean(value)
    if text.casefold() in ("", "-", "libur", "tidak ada jadwal", "nan"):
        return []
    parts = []
    for part in text.split(","):
        match = re.fullmatch(r"(\d{1,2})[:.](\d{2})\s*-\s*(\d{1,2})[:.](\d{2})", part.strip())
        if not match:
            raise ImportValidationError(f"Format jadwal HFIS tidak dikenali: {text}.")
        hour, minute, end_hour, end_minute = map(int, match.groups())
        if hour > 23 or minute > 59 or end_hour > 24 or end_minute > 59 or (end_hour == 24 and end_minute):
            raise ImportValidationError(f"Jam jadwal HFIS tidak valid: {text}.")
        parts.append((hour * 60 + minute, f"{hour:02d}:{minute:02d} - {end_hour:02d}:{end_minute:02d}"))
    return parts


def _last_practice(note: str) -> date:
    months = {name: number for number, name in enumerate([
        "januari", "februari", "maret", "april", "mei", "juni", "juli", "agustus", "september", "oktober", "november", "desember",
    ], start=1)}
    match = re.search(r"(\d{1,2})[- ](\d{1,2}|[A-Za-z]+)[- ](\d{4})", note)
    try:
        if match:
            month = int(match[2]) if match[2].isdigit() else months[match[2].casefold()]
            return date(int(match[3]), month, int(match[1]))
    except (ValueError, KeyError):
        pass
    raise ImportValidationError("Tanggal praktik terakhir pada riwayat penghapusan tidak dapat dibaca.")


class ScheduleBook:
    def __init__(self, detail: pd.DataFrame, history: pd.DataFrame):
        detail = _checked_columns(detail, DETAIL_COLUMNS, "Detail HFIS")
        history = _checked_columns(history, HISTORY_COLUMNS, "Riwayat Perubahan")
        if detail.empty:
            raise ImportValidationError("Tab Detail HFIS kosong.")
        self.current = defaultdict(list)
        self.rotation = {}
        self.events = defaultdict(list)
        self.day_notes = {}
        self.session_notes = {}
        self.removed = {}
        names = set()
        for _, row in detail.iterrows():
            name, day, note = (_clean(row[column]) for column in ("Nama dan Gelar", "Hari", "Keterangan"))
            if not name:
                raise ImportValidationError("Ada nama dokter kosong pada Detail HFIS.")
            names.add(name)
            if day not in DAYS:
                continue
            for start, label in _schedule_parts(row["Jadwal"]):
                self.current[(name, day)].append((start, label, note))
                self.day_notes[(name, day, label)] = note
                if note.upper() in ("JADWAL POLI", "HEMODIALISA"):
                    self.session_notes[(name, label)] = note
            if note.upper().startswith("MINGGU KE"):
                if note.upper() not in ("MINGGU KE-1 DAN KE-3", "MINGGU KE 2 & 4"):
                    raise ImportValidationError(f"Aturan jadwal bergilir belum dikenali: {note}.")
                self.rotation[(name, day)] = note.upper()

        event_versions = {}
        aliases = {_name_key(key): value for key, value in DOC_ALIASES.items()}
        for _, row in history.iterrows():
            hfis_name, document_name = _clean(row["Nama di HFIS"]), _clean(row["Nama di Dokumen"])
            found = hfis_name and _name_key(hfis_name) != _name_key(NOT_FOUND)
            name = hfis_name if found else aliases.get(_name_key(document_name), document_name)
            if not name:
                raise ImportValidationError("Ada nama dokter kosong pada Riwayat Perubahan.")
            names.add(name)
            day, note, kind = (_clean(row[column]) for column in ("Hari", "Keterangan", "Jenis"))
            if kind == "Penghapusan jadwal dokter" and not found:
                self.removed[name] = _last_practice(note)
            if day not in DAYS:
                continue
            notes = [part.strip() for part in note.split(";")]
            for column in ("Sebelumnya", "Perubahan"):
                parts = _schedule_parts(row[column])
                if len(parts) == len(notes):
                    for (_, label), session_note in zip(parts, notes):
                        self.session_notes.setdefault((name, label), session_note)
            if kind not in ("Perubahan jadwal praktik", "Penambahan dokter") or _clean(row["Berubah?"]).casefold() != "ya":
                continue
            effective = _calendar_date(row["Tanggal Dokumen"])
            before, after = _schedule_parts(row["Sebelumnya"]), _schedule_parts(row["Perubahan"])
            event_key = (name, day, effective)
            version = (before, after)
            if event_key in event_versions and event_versions[event_key] != version:
                raise ImportValidationError("Ada perubahan jadwal berbeda pada dokter, hari, dan tanggal dokumen yang sama.")
            event_versions[event_key] = version
            self.events[(name, day)].append((effective, before, after))
        for events in self.events.values():
            events.sort(key=lambda event: event[0])
        self.names = sorted(names)
        self.exact = defaultdict(list)
        self.core_names = defaultdict(list)
        for name in self.names:
            self.exact[_name_key(name)].append(name)
            self.core_names[_name_core(name)].append(name)

    def resolve_name(self, name: str) -> tuple[str, str, bool]:
        override = next((target for original, target in NAME_OVERRIDES.items() if _name_key(original) == _name_key(name)), None)
        exact = self.exact.get(_name_key(override or name), [])
        if len(exact) == 1:
            return exact[0], "Pemetaan khusus" if override else "Nama sama", False
        core_matches = self.core_names.get(_name_core(name), [])
        if len(core_matches) == 1:
            return core_matches[0], "Nama inti sama", False
        if len(exact) > 1 or len(core_matches) > 1:
            raise ImportValidationError(f"Pemetaan nama ambigu untuk {name}. Samakan nama hasil query dengan master HFIS.")
        ranked = sorted(
            ((SequenceMatcher(None, _name_core(name), _name_core(candidate)).ratio(), candidate) for candidate in self.names),
            reverse=True,
        )
        if not ranked or ranked[0][0] < 0.75 or (len(ranked) > 1 and ranked[0][0] - ranked[1][0] < 0.05):
            raise ImportValidationError(f"Nama {name} belum dapat dipetakan dengan yakin. Periksa ejaan dan master HFIS.")
        return ranked[0][1], "Kemiripan nama, wajib diperiksa", True

    def sessions(self, name: str, specialty: str, day: str, visit_date: date) -> list[tuple[int, str, str]]:
        if name in self.removed and visit_date > self.removed[name]:
            return []
        events = self.events.get((name, day), [])
        past = [event for event in events if event[0] <= visit_date]
        if events:
            parts = past[-1][2] if past else events[0][1]
            sessions = [(start, label, self.day_notes.get((name, day, label), self.session_notes.get((name, label), ""))) for start, label in parts]
        else:
            sessions = list(self.current.get((name, day), []))
        if _name_core(name) == "ilham ahmadi":
            hemodialysis = "HEMODIALISA" in specialty.upper()
            sessions = [session for session in sessions if (session[2].upper() == "HEMODIALISA") == hemodialysis]
        rotation = self.rotation.get((name, day))
        nth = (visit_date.day - 1) // 7 + 1
        if (rotation == "MINGGU KE-1 DAN KE-3" and nth not in (1, 3)) or (rotation == "MINGGU KE 2 & 4" and nth not in (2, 4)):
            return []
        return sorted(sessions)


def read_schedule_book(content: bytes) -> ScheduleBook:
    try:
        workbook = pd.ExcelFile(BytesIO(content), engine="openpyxl")
        if not {"Detail HFIS", "Riwayat Perubahan"}.issubset(workbook.sheet_names):
            raise ImportValidationError("Master jadwal harus memiliki tab Detail HFIS dan Riwayat Perubahan.")
        detail = pd.read_excel(workbook, sheet_name="Detail HFIS", dtype=object, keep_default_na=False)
        history = pd.read_excel(workbook, sheet_name="Riwayat Perubahan", dtype=object, keep_default_na=False)
        return ScheduleBook(detail, history)
    except ImportValidationError:
        raise
    except Exception as exc:
        raise ImportValidationError("Master HFIS tidak dapat dibaca. Gunakan Excel .xlsx yang valid.") from exc


def _period(value: Any, row_number: int) -> date:
    text = _clean(value)
    match = re.fullmatch(r"(\d{2}/\d{2}/\d{4})(?:\s+s/d\s+(\d{2}/\d{2}/\d{4}))?", text)
    try:
        if match:
            monday = datetime.strptime(match[1], "%d/%m/%Y").date()
            if monday.weekday() == 0 and (not match[2] or datetime.strptime(match[2], "%d/%m/%Y").date() == monday + timedelta(days=6)):
                return monday
    except ValueError:
        pass
    raise ImportValidationError(f"Periode pada baris {row_number} harus dimulai Senin dan berakhir Minggu (DD/MM/YYYY s/d DD/MM/YYYY).")


def _visit_time(value: Any, row_number: int, day: str) -> time | str | None:
    if isinstance(value, (datetime, time)):
        return value.time() if isinstance(value, datetime) else value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if 0 <= value < 1:
            seconds = round(value * 86400) % 86400
            return time(seconds // 3600, seconds % 3600 // 60, seconds % 60)
    text = _clean(value).lstrip("'")
    if text in ("", "-"):
        return text or None
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(text, fmt).time()
        except ValueError:
            pass
    raise ImportValidationError(f"Jam {day} pada baris {row_number} tidak valid. Gunakan HH:MM, HH:MM:SS, '-', atau kosong.")


def build_recap(query: pd.DataFrame, schedule: ScheduleBook) -> RecapResult:
    query = _checked_columns(query, QUERY_COLUMNS, "hasil query mingguan")
    rows, mappings, excluded = [], {}, []
    excluded_by_name = {_name_key(name): reason for name, reason in EXCLUDED_NAMES.items()}
    for index, row in query.iterrows():
        row_number = index + 2
        name, specialty = _clean(row["Nama Dokter"]), _clean(row["Poli / Subspesialis"])
        if not name or not specialty:
            raise ImportValidationError(f"Nama dokter atau spesialis kosong pada baris {row_number}.")
        if _name_key(name) in excluded_by_name:
            excluded.append({"Dokter": name, "Spesialis": specialty, "Alasan": excluded_by_name[_name_key(name)]})
            continue
        if (name, specialty) not in mappings:
            mapped, method, review = schedule.resolve_name(name)
            mappings[(name, specialty)] = {"Dokter hasil query": name, "Spesialis": specialty, "Dokter HFIS": mapped, "Metode": method, "Perlu konfirmasi": review}
        mapped = mappings[(name, specialty)]["Dokter HFIS"]
        monday = _period(row["Periode"], row_number)
        try:
            week = float(row["Minggu Ke"])
            if not week.is_integer() or week < 1:
                raise ValueError
        except (ValueError, TypeError):
            raise ImportValidationError(f"Minggu Ke pada baris {row_number} harus bilangan bulat positif.") from None
        output = [int(week), monday, name, specialty, ""]
        categories = []
        for offset, day in enumerate(DAYS):
            task = _visit_time(row[day], row_number, day)
            if task is None:
                output.extend(["", "", ""])
                continue
            sessions = schedule.sessions(mapped, specialty, day, monday + timedelta(days=offset))
            if not sessions:
                category = NO_SCHEDULE
            elif task == "-":
                category = "Tidak Sesuai"
            else:
                difference = task.hour * 60 + task.minute - sessions[0][0]
                category = "Sesuai" if -60 <= difference < 60 else "Tidak Sesuai"
            categories.append(category)
            output.extend([sessions[0][1] if sessions else NO_SCHEDULE, task, category])
        if not categories:
            raise ImportValidationError(f"Seluruh hari pada baris {row_number} kosong. Periksa hasil query.")
        output[4] = "Tidak Sesuai" if "Tidak Sesuai" in categories else "Sesuai" if "Sesuai" in categories else NO_SCHEDULE
        rows.append(output)
    if not rows:
        raise ImportValidationError("Tidak ada dokter yang dapat direkap setelah pengecualian dari aturan script.")
    data = validate_and_serialise(pd.DataFrame(rows, columns=RECAP_COLUMNS))
    return RecapResult(data, pd.DataFrame(mappings.values()), pd.DataFrame(excluded))


def recap_excel(result: RecapResult) -> bytes:
    """Export the same column schema without copying historical template data."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = SOURCE_SHEET
    header_font = Font(name="Aptos Narrow", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="0F625B")
    body_font = Font(name="Aptos Narrow", size=11)
    band_fill = PatternFill("solid", fgColor="F0F5F4")
    alignment = Alignment(vertical="center", wrap_text=True)
    sheet.append(RECAP_COLUMNS)
    for record in result.data.itertuples(index=False, name=None):
        values = list(record)
        values[0] = int(values[0])
        values[1] = date.fromisoformat(values[1])
        for index in range(6, 26, 3):
            if values[index] not in ("", "-"):
                values[index] = time.fromisoformat(values[index])
        sheet.append(values)
    for row in sheet:
        for cell in row:
            if cell.row == 1:
                cell.font = header_font
                cell.fill = header_fill
            else:
                cell.font = body_font
                if cell.row % 2 == 0:
                    cell.fill = band_fill
                if cell.column == 2:
                    cell.number_format = "dd/mm/yyyy"
                elif (cell.column - 7) % 3 == 0 and cell.column >= 7:
                    cell.number_format = "hh:mm"
            cell.alignment = alignment
            if isinstance(cell.value, str):
                cell.data_type = "s"
    widths = [12, 18, 46, 26, 22] + [18, 14, 20] * 7
    for number, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(number)].width = width
    sheet.row_dimensions[1].height = 32
    sheet.freeze_panes = "F2"
    sheet.auto_filter.ref = sheet.dimensions
    for title, frame in (("Pemetaan Dokter", result.mapping), ("Tidak Direkap", result.excluded)):
        if not frame.empty:
            audit = workbook.create_sheet(title)
            audit.append(list(frame.columns))
            for row in frame.itertuples(index=False, name=None):
                audit.append(list(row))
            for row in audit:
                for cell in row:
                    if isinstance(cell.value, str):
                        cell.data_type = "s"
    output = BytesIO()
    workbook.save(output)
    return output.getvalue()
