# Rekap kepatuhan dokter

Aplikasi Streamlit untuk mengolah hasil query CSV/Excel menjadi rekap kepatuhan jadwal dokter, lalu menambahkan hanya baris baru ke Google Sheets.

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

### Hasil query CSV / Excel

1. Pilih `Hasil query (CSV / Excel)` dan unggah hasil query mingguan. Bukan kode SQL dan bukan ekspor pasien per registrasi.
2. Unggah master jadwal HFIS dengan tab `Detail HFIS` dan `Riwayat Perubahan`.
3. Aplikasi menghitung jadwal yang berlaku per tanggal serta kategori harian/mingguan, lalu menyusun 26 kolom sesuai rekap tujuan.
4. Periksa preview, pemetaan dokter, dan pengecualian. Nama yang dipetakan berdasarkan kemiripan serta pengecualian perlu dikonfirmasi sebelum penambahan.
5. Pilih `Unduh rekap Excel` untuk menyimpan hasil, atau `Tambahkan data` untuk memasukkan baris baru ke Sheets.
6. Setelah hasil tulis terverifikasi, preview otomatis diperbarui. Unggah file saja tidak menulis ke Sheets.

Kolom hasil query yang diperlukan:

```text
Minggu Ke;Periode;Nama Dokter;Poli / Subspesialis;SENIN;SELASA;RABU;KAMIS;JUMAT;SABTU;MINGGU
```

- `Periode` menggunakan `DD/MM/YYYY s/d DD/MM/YYYY`, mulai Senin sampai Minggu.
- Hari berisi jam `H:MM` / `HH:MM:SS`, `-` untuk ada kunjungan tetapi TASK 4 kosong, atau sel kosong untuk tidak ada kunjungan.
- CSV menerima pemisah titik koma, koma, atau tab; Excel `.xlsx` dibaca dari tab pertama.
- Kolom tambahan seperti `No` dan `Minggu` tidak dibawa ke rekap.
- Data dengan tanggal, jam, atau identitas ganda yang tidak valid menghentikan proses sebelum penulisan.

### Master default dan pengganti

Di lokal dan Streamlit Cloud, aplikasi memakai `data/JADWAL DOKTER RS AN-NISA (HFIS) update 23-10-2026.xlsx` sebagai master default. File HFIS ini disertakan di GitHub atas persetujuan pemilik proyek. Upload master baru menggantikan default untuk sesi tersebut; file default di repository tidak ditimpa.

Kredensial, template berisi data, dan file lain di folder `data` tetap diabaikan Git. Hanya file HFIS default tersebut yang dikecualikan. Unggahan pengganti diproses dalam memori, tidak diarsipkan ke disk, dan tidak dipush ke GitHub secara otomatis. Setelah sesi berakhir atau server restart, aplikasi kembali memakai master default. Batasi akses aplikasi hanya untuk pengguna yang berwenang.

Untuk mengganti default secara permanen, ganti isi file HFIS pada path yang sama, lalu commit dan push. Jika nama file juga diubah, sesuaikan path pada `default_schedule()` di `recap.py` dan pengecualian pada `.gitignore`. Master hasil upload tidak mengubah rekap lama di Sheets secara otomatis; impor tetap hanya menambahkan baris baru.

Master yang baru harus mempertahankan riwayat perubahan yang diperlukan untuk tanggal data yang diproses. Jadwal yang tidak ada tidak boleh ditebak dari template rekap.

### Rekap yang sudah jadi

Pilih `Rekap siap impor` untuk mengunggah Excel yang sudah memiliki tab `Kepatuhan Jadwal Praktek Dokter`. Mode ini tidak memerlukan master HFIS atau menghitung ulang kategori.

Data dibandingkan dengan kunci `AWAL MINGGU (SENIN) + DOKTER`. Data lama tidak ditimpa. Perbedaan format jam seperti `9:05` dan `09:05:00` bukan konflik; perbedaan isi tetap dilaporkan.

## Pengujian

```powershell
python -m unittest discover -s tests -v
```

Tes meliputi pembacaan CSV/Excel, jadwal historis, penambahan/penghapusan dokter, Sabtu bergilir, pemisahan poli/hemodialisa, penilaian kategori, deduplikasi, download Excel, penggantian master, konfirmasi pemetaan, dan refresh setelah impor. Pengujian UI menggunakan Sheets simulasi, bukan menulis ke spreadsheet asli.
