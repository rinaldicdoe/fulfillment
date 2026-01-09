# Changelog - Sistem Verifikasi Data Excel

## Versi 2.3 - 2026-01-09 (Apostrophe Fix) ✅

### 🔧 Perbaikan: Leading Apostrophe

**MASALAH:**

- Excel menggunakan leading apostrophe (`'`) untuk memaksa angka sebagai text
- Contoh: `'001798275859` (untuk menjaga leading zero)
- Pandas membaca dengan apostrophe: `"'001798275859"`
- Database tidak punya apostrophe: `"001798275859"`
- Hasil: TIDAK COCOK ❌

**SOLUSI:**
Strip leading apostrophe dari nilai:

```python
if val_str.startswith("'"):
    val_str = val_str[1:]  # Remove first character
```

**HASIL:**

- ✅ `'001798275859` → `001798275859`
- ✅ `'311292500061988` → `311292500061988`
- ✅ Leading zero tetap terjaga
- ✅ Semua data sekarang cocok!

---

## Versi 2.2 - 2026-01-09 (Final Fix) ✅

### 🎯 Perbaikan Utama: Excel Number Precision

**MASALAH SEBENARNYA DITEMUKAN DAN DIPERBAIKI!**

#### **Root Cause:**

- Excel memiliki **limit 15 digit untuk angka**
- Nomor AWB panjang (seperti `311292500061988`) disimpan sebagai **angka** di Excel
- Excel mengkonversi ke scientific notation atau memotong digit
- Pandas membaca sebagai `int64` atau `float64`, bukan `string`
- Meskipun sudah dinormalisasi, **digit bisa berubah** karena precision loss

#### **Solusi:**

**Force read kolom C, D, E sebagai STRING** saat load Excel:

```python
def load_excel_file(uploaded_file):
    return pd.read_excel(
        uploaded_file,
        dtype={
            2: str,  # Column C (Everpro)
            3: str,  # Column D (Outgoing)
            4: str,  # Column E (Shopee)
        }
    )
```

#### **Hasil:**

- ✅ **Kolom dibaca langsung sebagai STRING**
- ✅ **Tidak ada precision loss**
- ✅ **Leading zero tetap ada** (0123 tetap "0123")
- ✅ **Digit tidak berubah** (311292500061988 tetap "311292500061988")
- ✅ **Verifikasi berhasil**: 129/130 data terverifikasi! 🎉

#### **Perubahan:**

1. Updated `load_excel_file()` dengan parameter `dtype`
2. Removed debug logging (sudah tidak diperlukan)

---

## Versi 2.1 - 2026-01-09 (Hotfix)

### 🔧 Perbaikan Kritis

#### **Perbaikan Fungsi Normalisasi Kolom D**

- **Masalah yang Ditemukan**: Setelah implementasi v2.0, hasil verifikasi menunjukkan 0 data terverifikasi
- **Penyebab**: Konversi sederhana `str(123.0)` menghasilkan `"123.0"` yang tidak cocok dengan `"123"` di database
- **Solusi**: Implementasi fungsi `normalize_value()` yang lebih pintar:
  - Menghapus `.0` dari angka bulat (123.0 → "123")
  - Menangani float dengan desimal dengan benar (123.45 → "123.45")
  - Menghapus whitespace
  - Menangani nilai NaN/None dengan benar

**Implementasi Baru:**

```python
def normalize_value(val):
    """Normalize value to string for comparison"""
    if pd.isna(val):
        return None

    val_str = str(val).strip()

    # Remove .0 from whole numbers
    if '.' in val_str:
        try:
            float_val = float(val_str)
            if float_val.is_integer():
                val_str = str(int(float_val))
        except (ValueError, OverflowError):
            pass

    return val_str
```

**Hasil:**

- ✅ `123.0` cocok dengan `"123"`
- ✅ `123` cocok dengan `"123"`
- ✅ `"123.0"` cocok dengan `"123"`
- ✅ `123.45` cocok dengan `"123.45"`
- ✅ Semua tipe data sekarang terdeteksi dengan benar

---

## Versi 2.0 - 2026-01-09

### ✨ Fitur Baru & Perbaikan

#### 1. **Upload Tunggal dengan Deteksi Otomatis**

- **Sebelumnya**: 4 tab terpisah dengan upload terpisah untuk JNE dan Non-JNE
- **Sekarang**: 3 tab dengan **satu upload field** yang bisa menerima multiple files
- **Fitur Deteksi Otomatis**:
  - Sistem otomatis mengenali tipe file berdasarkan nama file
  - File dengan kata "JNE" (tanpa "non") → Outgoing JNE
  - File dengan kata "non", "non-jne", atau "nonjne" → Outgoing Non-JNE
  - Jika tidak terdeteksi, user diminta memilih tipe secara manual
- **Manfaat**:
  - Upload lebih cepat - bisa pilih semua file sekaligus
  - Tidak perlu navigasi antar tab atau kolom
  - Sistem pintar mengenali tipe file otomatis
  - Fallback manual jika nama file tidak standar

#### 2. **Perbaikan Masalah Tipe Data Kolom D**

- **Masalah**: Kolom D kadang terbaca sebagai angka, kadang sebagai teks, menyebabkan data tidak ditemukan saat verifikasi
- **Solusi**: Semua data otomatis dikonversi ke string (teks) sebelum perbandingan
- **Implementasi**:

  ```python
  # Konversi nilai yang dicek ke string
  value_to_check_str = str(value_to_check).strip()

  # Konversi semua nilai referensi ke string
  everpro_values = everpro_df.iloc[:, 2].astype(str).str.strip().values
  shopee_values = shopee_df.iloc[:, 4].astype(str).str.strip().values
  ```

- **Manfaat**:
  - Perbandingan konsisten antara angka dan teks
  - Tidak ada lagi data yang terlewat karena perbedaan tipe data
  - Whitespace otomatis dihapus untuk akurasi lebih baik

### 📊 Detail Perubahan Teknis

#### File yang Dimodifikasi:

1. **app.py**:

   - Fungsi `verify_data()`: Ditambahkan konversi string untuk semua perbandingan data
   - UI Tabs: Dikurangi dari 4 menjadi 3 tab
   - Upload: Single file uploader dengan `accept_multiple_files=True`
   - Auto-detection: Logic untuk mendeteksi JNE vs Non-JNE dari nama file
   - Fallback: Radio button untuk manual selection jika auto-detect gagal

2. **README.md**:
   - Updated fitur utama
   - Dokumentasi proses verifikasi dengan penjelasan konversi tipe data
   - Penjelasan cara kerja auto-detection

### 🎯 Cara Penggunaan Baru

1. **Tab 1 - Database Everpro**: Upload file referensi Everpro
2. **Tab 2 - Shopee JNE Surabaya**: Upload file referensi Shopee
3. **Tab 3 - File Outgoing**:
   - Upload satu atau lebih file outgoing sekaligus
   - Sistem otomatis mendeteksi JNE atau Non-JNE dari nama file
   - Jika tidak terdeteksi, pilih tipe secara manual

### ✅ Testing

Untuk menguji aplikasi:

```bash
streamlit run app.py
```

Kemudian buka http://localhost:8501

### 🔍 Verifikasi Perbaikan

**Test Case untuk Masalah Tipe Data**:

- Upload file dengan kolom D berisi angka (misal: 12345)
- Upload file referensi dengan nilai yang sama sebagai teks ("12345")
- Verifikasi harus menemukan kecocokan (sebelumnya tidak ditemukan)

**Test Case untuk Auto-Detection**:

- Upload file dengan nama "Outgoing JNE.xlsx" → harus terdeteksi sebagai JNE
- Upload file dengan nama "Outgoing Non-JNE.xlsx" → harus terdeteksi sebagai Non-JNE
- Upload file dengan nama "Data.xlsx" → harus muncul pilihan manual
- Upload multiple files sekaligus → semua harus terproses dengan benar
