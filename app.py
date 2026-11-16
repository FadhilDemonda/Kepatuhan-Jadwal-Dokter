from __future__ import annotations

import hashlib

import pandas as pd
import streamlit as st

from importer import ImportValidationError, compare_frames, date_range_label, read_workbook, validate_and_serialise
from recap import build_recap, default_schedule, read_query_export, read_schedule_book, recap_excel
from sheets_client import append_rows, configured_client, read_destination, write_header


st.set_page_config(page_title="Rekap kepatuhan dokter", layout="wide")

st.markdown(
    """<style>
    .block-container { max-width: 1120px; padding-top: 3rem; padding-bottom: 3rem; }
    h1 { letter-spacing: -0.035em; font-weight: 700; }
    button, [role="button"] { min-height: 44px !important; cursor: pointer !important; }
    a[href], summary, select, label[for], input[type="radio"], input[type="checkbox"] { cursor: pointer; }
    label:has(input[type="radio"]), label:has(input[type="checkbox"]) { min-height: 44px !important; }
    button:disabled, [aria-disabled="true"] { cursor: not-allowed !important; }
    button:focus-visible, a:focus-visible, summary:focus-visible { outline: 2px solid #0f766e; outline-offset: 3px; }
    .stButton > button { min-height: 44px; border-radius: 8px; font-weight: 650; }
    .stButton > button[kind="primary"] { background: #0f625b; border-color: #0f625b; }
    .stButton > button[kind="primary"]:hover { background: #0b4d48; border-color: #0b4d48; }
    .stFileUploader { border-radius: 10px; }
    @media (max-width: 640px) { .block-container { padding: 1.5rem 1rem; } }
    </style>""",
    unsafe_allow_html=True,
)


def reset_preview() -> None:
    for key in ("file_hash", "source", "comparison", "headers", "sheet_url", "recap", "recap_download"):
        st.session_state.pop(key, None)


def make_preview(uploaded_file, mode: str, master_content: bytes | None, fingerprint: str) -> None:
    file_bytes = uploaded_file.getvalue()
    recap = None
    if mode == "Hasil query (CSV / Excel)":
        recap = build_recap(read_query_export(file_bytes, uploaded_file.name), read_schedule_book(master_content))
        source = recap.data
    else:
        source = validate_and_serialise(read_workbook(file_bytes))
    worksheet, sheet_url = configured_client(st.secrets)
    destination = read_destination(worksheet, list(source.columns))
    comparison = compare_frames(source, destination)
    st.session_state.update(
        file_hash=fingerprint,
        source=source,
        comparison=comparison,
        headers=list(source.columns),
        sheet_url=sheet_url,
        recap=recap,
        recap_download=recap_excel(recap) if recap is not None else None,
    )


st.title("Rekap kepatuhan dokter")
st.write("Olah hasil query menjadi rekap, lalu tambahkan hanya baris yang belum ada di Google Sheets.")

success_message = st.session_state.pop("import_success", None)
if success_message:
    st.success(success_message)

preview_notice = st.session_state.pop("preview_notice", None)
if preview_notice:
    st.info(preview_notice)

mode = st.radio(
    "Jenis file yang diunggah",
    ["Hasil query (CSV / Excel)", "Rekap siap impor"],
    key="input_mode",
    on_change=reset_preview,
)
master_content = None
if mode == "Hasil query (CSV / Excel)":
    uploaded_file = st.file_uploader(
        "Hasil query mingguan", type=["csv", "xlsx"], key="query_upload", on_change=reset_preview,
        help="Hasil query dengan kolom Minggu Ke, Periode, Nama Dokter, Poli / Subspesialis, dan SENIN sampai MINGGU. Excel dibaca dari tab pertama.",
    )
    master_upload = st.file_uploader(
        "Master HFIS pengganti (opsional jika master default tersedia)", type=["xlsx"],
        key="master_upload", on_change=reset_preview,
        help="Harus memiliki tab Detail HFIS dan Riwayat Perubahan. Unggahan ini menggantikan master default hanya untuk sesi ini.",
    )
    if master_upload is not None:
        master_name, master_content = master_upload.name, master_upload.getvalue()
    else:
        default_master = default_schedule()
        master_name, master_content = default_master if default_master is not None else (None, None)
    if master_name:
        st.caption(f"Master aktif: {master_name}")
    else:
        st.info("Master HFIS default belum tersedia di server ini. Unggah master HFIS untuk menghitung jadwal dan kategori.")
else:
    uploaded_file = st.file_uploader(
        "File Excel rekap mingguan", type=["xlsx"], key="recap_upload", on_change=reset_preview,
        help="Rekap yang sudah memiliki 26 kolom pada tab 'Kepatuhan Jadwal Praktek Dokter'.",
    )

if uploaded_file is None:
    st.info("Belum ada file. Pilih atau tarik hasil query CSV/Excel, atau pilih mode rekap siap impor.")
    st.stop()

if mode == "Hasil query (CSV / Excel)" and master_content is None:
    st.warning("Hasil query sudah dipilih. Lengkapi master HFIS sebelum diproses; template saja tidak cukup untuk menghitung kategori.")
    st.stop()

current_hash = hashlib.sha256(
    mode.encode() + hashlib.sha256(uploaded_file.getvalue()).digest()
    + hashlib.sha256(master_content or b"").digest()
).hexdigest()
if st.session_state.get("file_hash") != current_hash:
    try:
        with st.spinner("Mengolah file dan membandingkan rekap dengan Google Sheets..."):
            make_preview(uploaded_file, mode, master_content, current_hash)
    except ImportValidationError as exc:
        st.error(str(exc))
        st.stop()

source: pd.DataFrame = st.session_state.source
comparison = st.session_state.comparison
st.caption(f"{uploaded_file.name} · {len(source)} baris · {date_range_label(source)}")

recap = st.session_state.recap
needs_review = False
if recap is not None:
    st.download_button(
        "Unduh rekap Excel", data=st.session_state.recap_download, file_name="Rekap Kepatuhan Jadwal Dokter.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", on_click="ignore",
    )
    with st.expander("Lihat hasil rekap dan pemetaan dokter"):
        st.dataframe(source, width="stretch", hide_index=True, height=360)
        st.write("Pemetaan nama hasil query ke master HFIS")
        st.dataframe(recap.mapping, width="stretch", hide_index=True)
        if not recap.excluded.empty:
            st.write("Dokter yang tidak direkap sesuai aturan script")
            st.dataframe(recap.excluded, width="stretch", hide_index=True)
    needs_review = bool(recap.mapping["Perlu konfirmasi"].any()) or not recap.excluded.empty
    if needs_review:
        st.warning("Periksa pemetaan nama berdasarkan kemiripan dan dokter yang tidak direkap. Konfirmasi diperlukan sebelum penambahan ke Sheets.")

metric_a, metric_b, metric_c = st.columns(3)
metric_a.metric("Akan ditambahkan", len(comparison.new_rows))
metric_b.metric("Sudah tersimpan", comparison.existing_count)
metric_c.metric("Konflik", len(comparison.conflicts))

if not comparison.conflicts.empty:
    st.warning("Baris konflik tidak akan diubah. Periksa perbedaannya di Google Sheets atau sumber Excel.")
    st.dataframe(comparison.conflicts, width="stretch", hide_index=True)

if comparison.new_rows.empty:
    if comparison.conflicts.empty:
        st.success("Semua data sudah tersimpan. Tidak ada yang ditambahkan.")
    else:
        st.info("Tidak ada baris baru. Baris konflik dilewati dan data Sheets tidak diubah.")
    st.link_button("Buka Google Sheets", st.session_state.sheet_url)
    st.stop()

st.subheader("Preview data baru")
preview_columns = [column for column in ["MINGGU KE", "AWAL MINGGU (SENIN)", "DOKTER", "SPESIALIS", "KATEGORI REAL MINGGUAN"] if column in comparison.new_rows]
st.dataframe(comparison.new_rows[preview_columns], width="stretch", hide_index=True, height=360)
with st.expander("Lihat semua kolom data baru"):
    st.dataframe(comparison.new_rows, width="stretch", hide_index=True, height=420)

st.divider()
st.write(f"{len(comparison.new_rows)} baris baru akan ditambahkan. Data lama tidak diubah.")
review_confirmed = not needs_review or st.checkbox(
    "Saya sudah memeriksa pemetaan nama dan dokter yang tidak direkap", key=f"review_{current_hash}",
)
if st.button("Tambahkan data", type="primary", width="stretch", disabled=not review_confirmed):
    try:
        with st.spinner("Memeriksa ulang Google Sheets sebelum menambahkan data..."):
            worksheet, sheet_url = configured_client(st.secrets)
            latest_destination = read_destination(worksheet, st.session_state.headers)
            latest_comparison = compare_frames(source, latest_destination)
        if not latest_comparison.new_rows.equals(comparison.new_rows):
            st.session_state.comparison = latest_comparison
            st.session_state.preview_notice = "Data Google Sheets berubah sejak preview. Preview diperbarui, periksa lalu konfirmasi kembali."
            st.rerun()

        with st.spinner("Menambahkan data baru ke Google Sheets..."):
            if latest_destination.empty:
                write_header(worksheet, st.session_state.headers)
            append_rows(worksheet, latest_comparison.new_rows)
            verified = compare_frames(source, read_destination(worksheet, st.session_state.headers))
        if not verified.new_rows.empty:
            st.error("Penulisan belum dapat diverifikasi sepenuhnya. Jangan klik ulang, periksa Google Sheets terlebih dahulu.")
        else:
            st.session_state.import_success = f"{len(latest_comparison.new_rows)} baris berhasil ditambahkan dan diverifikasi."
            reset_preview()
            st.rerun()
    except ImportValidationError as exc:
        st.error(str(exc))
    except Exception:
        st.error("Terjadi gangguan saat menulis data. Jangan klik ulang. Periksa Google Sheets, lalu unggah ulang untuk menghitung sisa data.")
