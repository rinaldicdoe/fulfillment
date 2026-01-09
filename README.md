# Sistem Verifikasi Data Excel

Aplikasi web untuk memverifikasi data Excel dengan membandingkan data outgoing terhadap database referensi (Everpro dan Shopee JNE Surabaya).

## Fitur Utama

- **3 Tab Upload**: Everpro, Shopee JNE Surabaya, dan File Outgoing (dengan deteksi otomatis JNE/Non-JNE)
- **Deteksi Otomatis**: Sistem otomatis mengenali file JNE atau Non-JNE berdasarkan nama file
- **Multiple Upload**: Bisa upload beberapa file outgoing sekaligus
- **Verifikasi Otomatis**: Pengecekan data berdasarkan kolom referensi dengan konversi otomatis tipe data
- **Download Hasil**: File Excel terverifikasi dengan status tambahan
- **UI Profesional**: Interface yang clean dan user-friendly

## Cara Kerja

1. **Database Referensi**:

   - File Everpro (kolom sampai BO) - referensi kolom C
   - File Shopee JNE Surabaya (kolom sampai AW) - referensi kolom E

2. **File Verifikasi**:

   - Outgoing JNE (kolom sampai L)
   - Outgoing Non JNE (kolom sampai L)

3. **Proses Verifikasi**:
   - Kolom D dari file Outgoing dibandingkan dengan kolom C (Everpro) dan kolom E (Shopee JNE Surabaya)
   - Kolom C, D, E dibaca sebagai TEXT untuk mencegah Excel precision loss
   - Semua data otomatis dinormalisasi untuk memastikan perbandingan yang konsisten
   - Jika ditemukan di salah satu database: "Terverifikasi"
   - Jika tidak ditemukan: "Tidak Terverifikasi"
   - Status ditambahkan di kolom M

## Instalasi

1. Install dependencies:

```bash
pip install -r requirements.txt
```

2. Jalankan aplikasi:

```bash
streamlit run app.py
```

3. Buka browser di http://localhost:8501

## Struktur File

```
/
├── app.py              # Aplikasi utama Streamlit
├── requirements.txt    # Dependencies Python
└── README.md          # Dokumentasi ini
```

## Cara Penggunaan

1. **Upload File Database**: Upload minimal satu file database (Everpro atau Shopee JNE Surabaya)
2. **Upload File Outgoing**:
   - Upload file yang akan diverifikasi (bisa pilih multiple files sekaligus)
   - Sistem akan otomatis mendeteksi JNE atau Non-JNE berdasarkan nama file:
     - File dengan kata "JNE" (tanpa "non") → Outgoing JNE
     - File dengan kata "non", "non-jne", atau "nonjne" → Outgoing Non-JNE
     - Jika tidak terdeteksi, Anda akan diminta memilih tipe file secara manual
3. **Proses Verifikasi**: Klik tombol verifikasi untuk memulai
4. **Download Hasil**: Download file Excel yang sudah terverifikasi

## Teknologi

- **Streamlit**: Framework web app
- **Pandas**: Manipulasi data Excel
- **OpenPyXL**: Pembaca/penulis file Excel
- **Python 3.11+**: Runtime environment

## Support

Aplikasi mendukung format file:

- `.xlsx` (Excel 2007+)
- `.xls` (Excel 97-2003)
