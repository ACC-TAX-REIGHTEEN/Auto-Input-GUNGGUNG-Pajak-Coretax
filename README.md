# 🧾 Auto-Input GUNGGUNG Pajak Coretax

> **Otomasi pengolahan data penjualan Accurate → Template Coretax DJP siap upload**

Skrip Python berbasis pipeline yang membaca ekspor akuntansi dari Accurate (`ExAcc.xls`) dan data nota dari Google Sheets, melakukan validasi, pencocokan, perhitungan pajak (PPN 12%), lalu menghasilkan file Excel berformat **GUNGGUNG** yang siap diinput ke sistem Coretax DJP — semuanya cukup dengan satu klik.

---

## 📋 Daftar Isi

- [Fitur Utama](#-fitur-utama)
- [Prasyarat](#-prasyarat)
- [Struktur Folder](#-struktur-folder)
- [Cara Penggunaan](#-cara-penggunaan)
- [Alur Pipeline (Step-by-Step)](#-alur-pipeline-step-by-step)
- [Konfigurasi](#-konfigurasi)
- [Output](#-output)
- [Troubleshooting](#-troubleshooting)
- [Catatan Penting](#-catatan-penting)

---

## ✨ Fitur Utama

- **Deteksi kolom otomatis** — Tidak perlu mengatur mapping kolom manual. Skrip mendeteksi posisi kolom tanggal, nomor nota, nama pelanggan, dan nilai pajak secara dinamis berdasarkan isi data.
- **Normalisasi tanggal bilingual** — Menangani nama bulan Indonesia (Peb, Mei, Agu, Okt, Nop, Des) dan mengkonversinya ke format standar ISO `YYYY-MM-DD`.
- **Pencocokan & validasi pajak** — Mencocokkan nilai `Jumlah Pajak` antara data Accurate dan data nota Google Sheets dengan tiga mode fleksibilitas (Match / Toleransi / Losdol).
- **Kalkulasi PPN 12% otomatis** — Menghitung `Tax Base`, `OTB` (Other Tax Base), dan `VAT` sesuai formula Coretax DJP.
- **Filter data lintas masa** — Opsional menghapus baris yang "nyasar" (tanggal berbeda bulan dari mayoritas data).
- **Output siap pakai** — Menghasilkan file `GUNGGUNG-MMM-YY.xlsx` berformat template resmi yang langsung bisa diupload ke Coretax.
- **Auto-cleanup** — Semua file sementara (`*temp.xlsx`, salinan `ExAcc.xls` di folder Dapur) dibersihkan otomatis setelah proses selesai.

---

## 🔧 Prasyarat

### Python
Python **3.8+** disarankan.

### Library yang dibutuhkan
Install semua dependensi dengan:

```bash
pip install pandas openpyxl xlsxwriter xlrd xlwings
```

| Library | Kegunaan |
|---|---|
| `pandas` | Baca, transformasi, dan simpan data Excel |
| `openpyxl` | Baca/tulis `.xlsx` (engine utama) |
| `xlsxwriter` | Buat `.xlsx` baru dengan formatting |
| `xlrd` | Baca file legacy `.xls` dari Accurate |
| `xlwings` | Buka, insert baris, dan isi template Excel dengan memanfaatkan Excel/LibreOffice |

### Aplikasi tambahan
- **Microsoft Excel** atau **LibreOffice Calc** — Dibutuhkan oleh `xlwings` untuk memanipulasi template `TEMPLATE-GUNG.xlsx`.

---

## 📁 Struktur Folder

```
📦 Auto-Input-GUNGGUNG-Pajak-Coretax/
│
├── 📄 Jalankan Gunggung.py       ← File utama. Jalankan ini untuk memulai
├── 📄 ExAcc.xls                  ← [INPUT] Ekspor data Accurate bulanan (letakkan di sini)
│
└── 📁 Dapur/                     ← Folder kerja internal (jangan diubah strukturnya)
    ├── 📄 __init__.py
    ├── 📄 1_FetchFileSS.py       ← Unduh data nota dari Google Sheets
    ├── 📄 2_Acc_Cleaner.py       ← Bersihkan & normalisasi ExAcc.xls
    ├── 📄 2_Acc_CleanerDel.py    ← Filter baris berdasarkan masa pajak
    ├── 📄 3_MovDt2MtchFrFetch.py ← Normalisasi tanggal data nota
    ├── 📄 4_Merged_all.py        ← Gabungkan kedua sumber data
    ├── 📄 5_Processing.py        ← Validasi, pencocokan, dan hitung pajak
    ├── 📄 6_Premov2template.py   ← Susun kolom sesuai format Coretax
    ├── 📄 TEMPLATE-GUNG.xlsx     ← Template resmi Coretax DJP (jangan dihapus)
    ├── 📄 Hapus.conf             ← Konfigurasi mode filter masa pajak
    └── 📄 Helper.conf            ← Konfigurasi mode validasi pencocokan pajak
```

---

## 🚀 Cara Penggunaan

### Langkah 1 — Siapkan file input

1. Export data bulan berjalan dari **Accurate** ke format `.xls`.
2. Simpan file tersebut dengan nama **`ExAcc.xls`** di folder utama proyek (sejajar dengan `Jalankan Gunggung.py`).

### Langkah 2 — Isi Spreadsheet ID Google Sheets

Buka file `Dapur/1_FetchFileSS.py`, lalu ganti nilai `sheet_id` dengan ID Google Sheets yang berisi data nota terbaru:

```python
# Sebelum
sheet_id = "SPREADSHEETS ID DI SINI"

# Sesudah (contoh)
sheet_id = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms"
```

> **Cara mendapatkan ID:** Buka Google Sheets → lihat URL: `https://docs.google.com/spreadsheets/d/**[ID ADA DI SINI]**/edit`

Pastikan Google Sheets tersebut **dapat diakses publik** (minimal "Anyone with the link can view").

### Langkah 3 — (Opsional) Sesuaikan konfigurasi

Lihat bagian [Konfigurasi](#-konfigurasi) untuk mengatur mode filter dan validasi sesuai kebutuhan.

### Langkah 4 — Jalankan

```bash
python "Jalankan Gunggung.py"
```

atau klik dua kali file tersebut jika sudah ada asosiasi Python di sistem Anda.

### Langkah 5 — Pilih sheet Google Sheets

Ketika script `1_FetchFileSS.py` berjalan, program akan menampilkan daftar sheet yang tersedia dan meminta input nomor:

```
--> Ditemukan 3 sheet dalam file ini:
1. Januari 2025
2. Februari 2025
3. Maret 2025

--> Masukkan NOMOR sheet yang ingin diunduh: 
```

Ketik nomor yang sesuai lalu tekan **Enter**.

### Langkah 6 — Ambil hasil

Setelah proses selesai, file hasil akan muncul di folder utama dengan nama:

```
GUNGGUNG-MAR-25.xlsx   ← (nama menyesuaikan bulan dan tahun data)
```

---

## 🔄 Alur Pipeline (Step-by-Step)

Pipeline ini dijalankan secara berurutan oleh `Jalankan Gunggung.py`:

```
[Mulai]
   │
   ├─── initial_cleanup()
   │       Hapus sisa file *temp.xlsx dan ExAcc.xls lama dari folder Dapur
   │
   ├─── check_requirements()
   │       Verifikasi ExAcc.xls ada di folder utama
   │       Verifikasi semua file di folder Dapur lengkap
   │
   ├─── run_scripts()
   │       │
   │       ├─ Salin ExAcc.xls → Dapur/
   │       │
   │       ├─ [1] 1_FetchFileSS.py
   │       │       Unduh data nota dari Google Sheets
   │       │       Simpan sebagai: Nota_terbaru_temp.xlsx
   │       │
   │       ├─ [2] 2_Acc_Cleaner.py
   │       │       Baca ExAcc.xls (deteksi kolom otomatis)
   │       │       Normalisasi tanggal (nama bulan Indo → Eng)
   │       │       Simpan sebagai: ExAcc_clean_a_temp.xlsx
   │       │
   │       ├─ [3] 2_Acc_CleanerDel.py
   │       │       Filter baris berdasarkan Tax Date (lihat Hapus.conf)
   │       │       Simpan sebagai: ExAcc_clean_temp.xlsx
   │       │
   │       ├─ [4] 3_MovDt2MtchFrFetch.py
   │       │       Normalisasi kolom Tanggal di data nota
   │       │       Simpan sebagai: Nota_terbaru_up_temp.xlsx
   │       │
   │       ├─ [5] 4_Merged_all.py
   │       │       Gabungkan ExAcc_clean_temp + Nota_terbaru_up_temp
   │       │       Simpan sebagai: Merged_all_temp.xlsx (2 sheet)
   │       │
   │       ├─ [6] 5_Processing.py
   │       │       Cocokkan Jumlah Pajak dari kedua sumber (lihat Helper.conf)
   │       │       Jika tidak cocok → simpan laporan error & hentikan proses
   │       │       Jika cocok → hitung Tax Base, OTB, VAT
   │       │       Simpan sebagai: Pre-moving_temp.xlsx
   │       │
   │       └─ [7] 6_Premov2template.py
   │               Susun ulang kolom ke format Coretax DJP
   │               Tambahkan kolom tetap (TrxCode, BuyerName, dll.)
   │               Simpan sebagai: Pre_template_temp.xlsx
   │
   ├─── process_excel_template()
   │       Buka TEMPLATE-GUNG.xlsx via xlwings
   │       Insert baris sesuai jumlah data
   │       Tempel data mulai dari cell B5
   │       Isi bulan (B2) dan tahun (B3) otomatis
   │       Simpan sebagai: TEMPLATE-GUNG-NEW.xlsx
   │
   ├─── finalize_and_cleanup()
   │       Rename ke GUNGGUNG-MMM-YY.xlsx
   │       Pindahkan ke folder utama
   │       Hapus semua file *temp.xlsx
   │       Hapus ExAcc.xls dari folder Dapur
   │
[Selesai] ✅
```

---

## ⚙️ Konfigurasi

### `Dapur/Hapus.conf` — Mode filter masa pajak

Mengatur bagaimana skrip menangani baris yang memiliki `Tax Date` berbeda dari bulan mayoritas.

```ini
[FILTER]
mode = hapus
```

| Nilai | Perilaku |
|---|---|
| `hapus` | Baris dengan `Tax Date` berbeda bulan dari mayoritas data **dihapus** secara otomatis. Berguna jika ada transaksi "nyasar" dari bulan lain. |
| `lengkap` | Semua baris dipertahankan apa adanya. Gunakan jika Anda yakin semua data sudah benar, atau jika ingin melihat seluruh data tanpa penyaringan. |

### `Dapur/Helper.conf` — Mode validasi pencocokan pajak

Mengatur tingkat keketatan saat mencocokkan nilai `Jumlah Pajak` antara data Accurate dan data nota Google Sheets.

```ini
[SETTINGS]
mode = Toleransi
```

| Nilai | Perilaku |
|---|---|
| `Match` | Nilai pajak harus **sama persis** (selisih = 0). Paling ketat. |
| `Toleransi` | Selisih hingga **±1** masih dianggap cocok. Berguna untuk mengatasi perbedaan pembulatan antar sistem. |
| `Losdol` | **Semua data dianggap cocok** tanpa memandang selisih nilai. Program tetap menampilkan peringatan jika ada perbedaan, namun proses tidak dihentikan. |

> ⚠️ Gunakan mode `Losdol` dengan hati-hati — kesalahan data tidak akan menyebabkan proses berhenti.

---

## 📤 Output

### File utama hasil proses

```
GUNGGUNG-MAR-25.xlsx
```

Nama file dibentuk otomatis dari bulan dan tahun **mayoritas** kolom `TransactionDate` dalam data.

### Struktur kolom dalam output

| Kolom | Deskripsi |
|---|---|
| `TrxCode` | Kode transaksi (default: `Normal`) |
| `BuyerName` | Nama pembeli (default: `-` untuk transaksi B2C) |
| `BuyerIdOpt` | Opsi identitas pembeli (`NIK`) |
| `BuyerIdNumber` | Nomor identitas pembeli (`0000000000000000`) |
| `GoodServiceOpt` | Jenis barang/jasa (`A`) |
| `SerialNo` | Nomor nota / faktur |
| `TransactionDate` | Tanggal transaksi (format `YYYY-MM-DD`) |
| `TaxBaseSellingPrice` | Dasar Pengenaan Pajak (`DPP` = Nilai Faktur / 1.11) |
| `OtherTaxBaseSellingPrice` | Nilai lain DPP (`OTB` = Tax Base × 11/12) |
| `VAT` | PPN terutang (`VAT` = OTB × 12%) |
| `STLG` | Stamp Tax / Luxury Goods (default: `0`) |
| `Info` | Keterangan tambahan (default: `ok`) |

### Formula Perhitungan Pajak

```
Tax Base = CEILING(Nilai Faktur / 1.11)
OTB      = CEILING(Tax Base × (11 / 12))
VAT      = CEILING(OTB × 0.12)
```

> Semua nilai dibulatkan ke atas (`CEILING`) sesuai ketentuan perhitungan PPN 12% Coretax DJP.

---

## 🛠️ Troubleshooting

### ❌ `File tidak ditemukan: ExAcc.xls`
Pastikan file ekspor Accurate sudah diletakkan di folder utama (sejajar dengan `Jalankan Gunggung.py`) dan namanya adalah **persis** `ExAcc.xls`.

### ❌ `Kolom Tanggal tidak ditemukan`
Skrip mendeteksi kolom tanggal berdasarkan keberadaan nama bulan (Jan, Peb, Mar, dst.). Pastikan file `ExAcc.xls` merupakan ekspor langsung dari Accurate dan formatnya belum diubah secara manual.

### ❌ `Ditemukan X data tidak cocok` (proses berhenti di step 5)
Nilai `Jumlah Pajak` di Accurate tidak cocok dengan data nota di Google Sheets. Langkah penanganan:
1. Buka file `Hasil_Matching_Tidak_Cocok_temp.xlsx` di folder `Dapur/` — file ini berisi daftar baris yang bermasalah beserta kolom `Matching` dan `Ket Matching`.
2. Periksa dan koreksi data sumbernya.
3. Atau, jika perbedaan kecil (akibat pembulatan), ubah `Helper.conf` ke mode `Toleransi`.

### ❌ `xlwings Error` saat proses template
`xlwings` membutuhkan Microsoft Excel atau LibreOffice Calc yang terinstall. Pastikan salah satunya tersedia di sistem Anda. Pada Linux/Mac tanpa Excel, pastikan LibreOffice Calc sudah terinstall dan bisa dijalankan via command line.

### ❌ Error saat fetch Google Sheets
- Pastikan `sheet_id` di `1_FetchFileSS.py` sudah diisi dengan benar.
- Pastikan Google Sheets dapat diakses secara publik ("Anyone with the link can view").
- Periksa koneksi internet.

---

## 📌 Catatan Penting

- **Jangan mengubah struktur folder `Dapur/`** — semua skrip saling bergantung satu sama lain dan mengasumsikan file-file berada di lokasi yang sudah ditentukan.
- **File `TEMPLATE-GUNG.xlsx` wajib ada** dan tidak boleh dihapus — ini adalah template resmi yang diisi data secara otomatis oleh `xlwings`.
- **Jalankan dari folder utama** — bukan dari dalam folder `Dapur/`. Gunakan `Jalankan Gunggung.py` sebagai satu-satunya titik masuk.
- **Selalu cek ulang hasil output** sebelum diupload ke Coretax. Mesin bisa keliru — verifikasi manual tetap disarankan.
- File `*temp.xlsx` di folder `Dapur/` adalah file kerja sementara dan akan dihapus otomatis. Jangan simpan data penting di sana.

---

## 📜 Lisensi

Proyek ini dikembangkan untuk keperluan internal perpajakan. Silakan sesuaikan dengan kebutuhan organisasi Anda.

---

*Dikembangkan oleh [ACC-TAX-REIGHTEEN](https://github.com/ACC-TAX-REIGHTEEN)*
