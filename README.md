# Rekap kepatuhan dokter

Aplikasi Streamlit untuk menambahkan baris baru dari Excel rekap mingguan ke Google Sheets.

## Siapkan Google Cloud

1. Buat service account di Google Cloud Console dan aktifkan Google Sheets API.
2. Unduh JSON kredensial service account. Jangan masukkan file atau private key ke Git.
3. Bagikan Google Sheets tujuan ke alamat `client_email` service account sebagai **Editor**.
4. Salin `.streamlit/secrets.toml.example` menjadi `.streamlit/secrets.toml`, lalu isi dengan nilai dari JSON dan ID spreadsheet.

API key biasa tidak dapat menulis ke Google Sheets privat.

## Jalankan lokal

```powershell
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Cara memakai

1. Unggah Excel hasil ekspor SQL.
2. Periksa jumlah data baru dan konflik.
3. Tekan `Tambahkan data`.
4. Aplikasi membaca ulang Sheets sebelum menulis dan memverifikasi hasil sesudahnya.

Data dibandingkan dengan kunci `AWAL MINGGU (SENIN) + DOKTER`. Data lama tidak ditimpa. Tab yang dibaca dari Excel adalah `Kepatuhan Jadwal Praktek Dokter`.
