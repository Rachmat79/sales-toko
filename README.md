# 📊 Sales Display Analyzer
### Executive Sales Performance Dashboard — Auto-Generate Edition

Dashboard sales berbasis Streamlit. Perbedaan utama dari versi sebelumnya:
mekanisme input data diganti dari **Upload manual** menjadi **Generate otomatis**
(aplikasi mencari sendiri file data terbaru di komputer user).

---

## 1. Cara Menjalankan

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 2. Perubahan Utama: Upload → Generate

| Sebelumnya | Sekarang |
|---|---|
| User klik "Browse files" dan pilih file manual | User klik **⚡ GENERATE DATA**, aplikasi cari file sendiri |
| File disimpan di memori browser (upload) | File dibaca langsung dari disk komputer user |
| Tidak ada info lokasi file | Menampilkan File, Lokasi, Last Modified, Format |

**Seluruh dashboard (KPI, chart, ranking, filter, tab, export Excel) tidak diubah sama sekali** —
hanya mekanisme *bagaimana data masuk ke aplikasi* yang berubah.

## 3. Cara Kerja Fitur Generate

1. **Hanya berjalan saat tombol diklik** — aplikasi TIDAK otomatis mencari file
   saat pertama dibuka. User harus klik **⚡ GENERATE DATA** di sidebar untuk
   memulai pencarian & pembacaan data.
2. **Pencarian folder** — mengecek folder berikut secara berurutan (folder pertama
   yang punya file valid akan dipakai):
   ```
   1. ~/Downloads
   2. ~/Download
   3. ~/Desktop
   4. ~/Documents
   ```
   Lokasi diambil otomatis via `Path.home()` — **tidak ada username yang di-hardcode**,
   sehingga bekerja untuk user mana pun di Windows/Mac/Linux.
3. **Deteksi file terbaru** — di dalam folder yang dipakai, sistem mencari file
   `.xlsx`, `.xls`, `.csv`, mengabaikan file temporary (`~$...`, `.tmp`, `.temp`),
   lalu mengambil yang **modified time-nya paling baru**.
4. **Loading state berurutan**: 🔍 Searching → 📂 Checking folder → 📊 Reading file →
   ⚙️ Processing → ✅ Data ready (ditampilkan via `st.status`).
5. **Info file terpilih** ditampilkan di sidebar: nama file, lokasi lengkap,
   last modified, dan format — serta caption kecil `📄 Source / 🕒 Updated` di
   atas dashboard.
6. **Tombol Refresh** — scan ulang folder & reload data terbaru kapan pun
   dibutuhkan (mis. setelah user download file baru).

## 4. Dukungan Format File

- **.xlsx** — dibaca via `pandas` + engine `openpyxl`.
- **.xls** — dibaca via `pandas` + engine `xlrd` (perlu `pip install xlrd`,
  sudah termasuk di `requirements.txt`).
- **.csv** — dibaca via `pandas`, dengan fallback encoding otomatis:
  `utf-8` → `utf-8-sig` → `latin-1`.

## 5. Error Handling

Semua kondisi berikut ditangani tanpa membuat aplikasi crash, dengan pesan
berbahasa Indonesia yang jelas:

- Tidak ada file ditemukan di folder manapun → pesan folder yang dicek.
- File corrupt / format tidak valid → pesan error + nama file.
- File sedang terkunci/digunakan aplikasi lain → pesan minta tutup file dulu.
- Semua encoding CSV gagal → pesan detail.
- Struktur kolom tidak sesuai (Sales/Category tidak ditemukan, atau
  kebetulan sama) → pesan kolom yang tersedia vs dibutuhkan.
- Data kosong setelah dibaca/difilter → pesan jelas.

## 6. Contoh Workflow Penggunaan

```
User download file sales_20260816.xlsx ke folder Downloads
                    ↓
User buka aplikasi Streamlit (dashboard masih kosong, menunggu Generate)
                    ↓
User klik "⚡ GENERATE DATA" di sidebar
                    ↓
Sistem menemukan sales_20260816.xlsx (file terbaru di Downloads)
                    ↓
Info file ditampilkan: File, Lokasi, Last Modified, Format
                    ↓
Dashboard tampil seperti biasa (KPI, chart, tab, filter)
                    ↓
Besok, user download file baru sales_20260817.xlsx
                    ↓
User klik "🔄 Refresh" di sidebar
                    ↓
Sistem otomatis memakai file terbaru & dashboard ter-update
```

## 8. Fitur Baru: Top by Category Code

Pada tab **📋 Ranking Category**, selain ranking berdasarkan Category (nama),
kini tersedia juga ranking berdasarkan **Category Code** (jika kolom tersebut
ada pada file):

- Kolom "Category Code" terdeteksi otomatis lewat mapping kolom di sidebar
  (opsional — sama seperti Store/Date, tidak wajib diisi).
- Jika terdeteksi/dipilih, tab Ranking Category menampilkan dua tabel:
  **Ranking by Category** dan **Top by Category Code**, lengkap dengan
  Rank, Sales, dan Contribution %.
- Jika tidak tersedia, ditampilkan pesan info agar user tahu cara mengaktifkannya
  (bukan error, aplikasi tetap berjalan normal).

## 9. Catatan Performa

- Export Excel di-cache (`st.cache_data`) — hanya dibangun ulang saat data/filter
  benar-benar berubah, bukan di setiap interaksi UI, sehingga dashboard tetap
  responsif meski file sumber besar.
- Pencarian folder bersifat **non-recursive** (hanya isi langsung folder, bukan
  scan seluruh hard disk) agar tetap cepat.
