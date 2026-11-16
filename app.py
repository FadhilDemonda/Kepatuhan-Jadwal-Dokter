from __future__ import annotations

import hashlib

import pandas as pd
import streamlit as st

from importer import ImportValidationError, compare_frames, date_range_label, read_workbook, validate_and_serialise
from sheets_client import append_rows, configured_client, read_destination, write_header


st.set_page_config(page_title="Rekap kepatuhan dokter", layout="wide")

st.markdown(
    """<style>
    .block-container { max-width: 1120px; padding-top: 3rem; padding-bottom: 3rem; }
    h1 { letter-spacing: -0.035em; font-weight: 700; }
    .stButton > button { min-height: 44px; border-radius: 8px; font-weight: 650; }
    .stButton > button[kind="primary"] { background: #0f625b; border-color: #0f625b; }
    .stButton > button[kind="primary"]:hover { background: #0b4d48; border-color: #0b4d48; }
    .stFileUploader { border-radius: 10px; }
    [data-testid="stMetric"] { border-left: 3px solid #0f625b; padding-left: 0.8rem; }
    @media (max-width: 640px) { .block-container { padding: 1.5rem 1rem; } }
    </style>""",
    unsafe_allow_html=True,
)


def reset_preview() -> None:
    for key in ("file_hash", "source", "comparison", "headers", "sheet_url"):
        st.session_state.pop(key, None)


def make_preview(uploaded_file) -> None:
    file_bytes = uploaded_file.getvalue()
    source = validate_and_serialise(read_workbook(file_bytes))
    worksheet, sheet_url = configured_client(st.secrets)
    destination = read_destination(worksheet, list(source.columns))
    comparison = compare_frames(source, destination)
    st.session_state.update(
        file_hash=hashlib.sha256(file_bytes).hexdigest(),
        source=source,
        comparison=comparison,
        headers=list(source.columns),
        sheet_url=sheet_url,
    )


st.title("Rekap kepatuhan dokter")
st.write("Unggah Excel kumulatif. Aplikasi hanya menambahkan baris yang belum ada di Google Sheets.")

uploaded_file = st.file_uploader(
    "File Excel rekap mingguan",
    type=["xlsx"],
    help="Gunakan hasil ekspor SQL dengan tab 'Kepatuhan Jadwal Praktek Dokter'.",
    on_change=reset_preview,
)

if uploaded_file is None:
    st.info("Belum ada file. Pilih atau tarik file Excel untuk melihat data baru sebelum ditambahkan.")
    st.stop()

current_hash = hashlib.sha256(uploaded_file.getvalue()).hexdigest()
if st.session_state.get("file_hash") != current_hash:
    try:
        with st.spinner("Memeriksa Excel dan membandingkan dengan Google Sheets..."):
            make_preview(uploaded_file)
    except ImportValidationError as exc:
        st.error(str(exc))
        st.stop()

source: pd.DataFrame = st.session_state.source
comparison = st.session_state.comparison
st.caption(f"{uploaded_file.name} · {len(source)} baris · {date_range_label(source)}")

metric_a, metric_b, metric_c = st.columns(3)
metric_a.metric("Akan ditambahkan", len(comparison.new_rows))
metric_b.metric("Sudah tersimpan", comparison.existing_count)
metric_c.metric("Konflik", len(comparison.conflicts))

if not comparison.conflicts.empty:
    st.warning("Baris konflik tidak akan diubah. Periksa perbedaannya di Google Sheets atau sumber Excel.")
    st.dataframe(comparison.conflicts, use_container_width=True, hide_index=True)

if comparison.new_rows.empty:
    st.success("Semua data sudah tersimpan. Tidak ada yang ditambahkan.")
    st.link_button("Buka Google Sheets", st.session_state.sheet_url)
    st.stop()

st.subheader("Preview data baru")
preview_columns = [column for column in ["MINGGU KE", "AWAL MINGGU (SENIN)", "DOKTER", "SPESIALIS", "KATEGORI REAL MINGGUAN"] if column in comparison.new_rows]
st.dataframe(comparison.new_rows[preview_columns], use_container_width=True, hide_index=True, height=360)
with st.expander("Lihat semua kolom data baru"):
    st.dataframe(comparison.new_rows, use_container_width=True, hide_index=True, height=420)

st.divider()
st.write(f"{len(comparison.new_rows)} baris baru akan ditambahkan. Data lama tidak diubah.")
if st.button("Tambahkan data", type="primary", use_container_width=True):
    try:
        with st.spinner("Memeriksa ulang Google Sheets sebelum menambahkan data..."):
            worksheet, sheet_url = configured_client(st.secrets)
            latest_destination = read_destination(worksheet, st.session_state.headers)
            latest_comparison = compare_frames(source, latest_destination)
        if len(latest_comparison.new_rows) != len(comparison.new_rows):
            st.session_state.comparison = latest_comparison
            st.warning("Data Google Sheets berubah sejak preview. Preview diperbarui, periksa lalu konfirmasi kembali.")
            st.rerun()

        with st.spinner("Menambahkan data baru ke Google Sheets..."):
            if latest_destination.empty:
                write_header(worksheet, st.session_state.headers)
            append_rows(worksheet, latest_comparison.new_rows)
            verified = compare_frames(source, read_destination(worksheet, st.session_state.headers))
        if not verified.new_rows.empty:
            st.error("Penulisan belum dapat diverifikasi sepenuhnya. Jangan klik ulang, periksa Google Sheets terlebih dahulu.")
        else:
            st.success(f"{len(latest_comparison.new_rows)} baris berhasil ditambahkan dan diverifikasi.")
            st.link_button("Buka Google Sheets", sheet_url)
            reset_preview()
    except ImportValidationError as exc:
        st.error(str(exc))
    except Exception:
        st.error("Terjadi gangguan saat menulis data. Jangan klik ulang. Periksa Google Sheets, lalu unggah ulang untuk menghitung sisa data.")
