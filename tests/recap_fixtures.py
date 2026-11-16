from io import BytesIO

import pandas as pd

from recap import DAYS, DETAIL_COLUMNS, HISTORY_COLUMNS, QUERY_COLUMNS


DOCTOR = "dr. Dokter Contoh, Sp.PD"


def query_frame(**changes):
    record = {
        "Minggu Ke": "45", "Periode": "09/11/2026 s/d 15/11/2026",
        "Nama Dokter": DOCTOR, "Poli / Subspesialis": "PENYAKIT DALAM",
        **dict.fromkeys(DAYS, ""), "SENIN": "9:05",
    }
    record.update(changes)
    return pd.DataFrame([record], columns=QUERY_COLUMNS)


def detail_frame(rows=None):
    return pd.DataFrame(rows if rows is not None else [[DOCTOR, "SENIN", "09:00 - 12:00", ""]], columns=DETAIL_COLUMNS)


def history_frame(rows=None):
    return pd.DataFrame(rows or [], columns=HISTORY_COLUMNS)


def master_bytes(detail=None, history=None):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        (detail if detail is not None else detail_frame()).to_excel(writer, sheet_name="Detail HFIS", index=False)
        (history if history is not None else history_frame()).to_excel(writer, sheet_name="Riwayat Perubahan", index=False)
    return output.getvalue()


def query_bytes(frame=None, extension="csv", delimiter=";"):
    frame = frame if frame is not None else query_frame()
    if extension == "csv":
        return frame.to_csv(index=False, sep=delimiter).encode("utf-8-sig")
    output = BytesIO()
    frame.to_excel(output, index=False, engine="openpyxl")
    return output.getvalue()
