# Perpustakaan Digital Terintegrasi Berbasis Python dan Tkinter

Repository ini berisi implementasi **Sistem Informasi Perpustakaan Digital** berbasis **Python** dengan antarmuka grafis **Tkinter**, yang dirancang menggunakan pendekatan **struktur data Linked List** serta penyimpanan data non-relasional berbasis file. Sistem ini dikembangkan sebagai proyek akademik untuk mendukung penelitian dan pembelajaran pada bidang **algoritma, struktur data, dan sistem informasi**.

---

## Abstrak

Pengelolaan perpustakaan secara manual memiliki berbagai keterbatasan, terutama dalam hal pencatatan data, pelacakan transaksi, serta pengendalian keterlambatan pengembalian buku. Repository ini menyajikan sebuah aplikasi perpustakaan digital berbasis desktop yang mengintegrasikan manajemen koleksi, data peminjam, transaksi peminjaman, pengembalian, serta notifikasi jatuh tempo secara otomatis. Sistem ini menerapkan struktur data **Linked List** sebagai inti pengelolaan data dan memanfaatkan file CSV sebagai media penyimpanan persisten tanpa menggunakan basis data relasional.

---

## Tujuan Pengembangan

Pengembangan sistem ini bertujuan untuk:

1. Mengimplementasikan struktur data Linked List dalam studi kasus nyata.
2. Membangun sistem informasi perpustakaan berbasis desktop yang ringan dan mandiri.
3. Menyediakan simulasi manajemen data tanpa ketergantungan pada database relasional.
4. Mendukung otomasi notifikasi pengembalian buku melalui email dan WhatsApp.
5. Menghasilkan laporan administrasi perpustakaan secara otomatis dalam format dokumen.

---

## Ruang Lingkup Sistem

Sistem yang dikembangkan mencakup ruang lingkup berikut:

- Autentikasi administrator
- Manajemen data buku
- Manajemen data mahasiswa/peminjam
- Transaksi peminjaman dan pengembalian
- Perhitungan denda keterlambatan
- Pengiriman notifikasi jatuh tempo
- Pembuatan laporan administrasi

Sistem ini merupakan **aplikasi desktop**, bukan aplikasi berbasis web.

---

## Arsitektur dan Pendekatan Teknis

### 1. Struktur Data
Sistem menggunakan struktur data **Linked List** untuk menyimpan:
- Data buku
- Data mahasiswa
- Data transaksi peminjaman

Pendekatan ini dipilih untuk menunjukkan penerapan konsep algoritmik dalam pengelolaan data dinamis.

### 2. Penyimpanan Data
Data disimpan menggunakan file:
- `data_buku.csv`
- `data_mahasiswa.csv`
- `transaksi.csv`

Model ini merepresentasikan pendekatan **non-relational data management**.

### 3. Antarmuka Pengguna
Antarmuka pengguna dikembangkan menggunakan pustaka **Tkinter** yang merupakan GUI standar pada Python.

---

## Fitur Sistem

### Autentikasi
- Login administrator berbasis kredensial statis.

### Manajemen Buku
- Tambah, ubah, dan hapus data buku
- Pencarian dan pengurutan data
- Pengelolaan stok dan ketersediaan otomatis
- Import data dari file CSV

### Manajemen Mahasiswa
- Tambah, ubah, dan hapus data mahasiswa
- Import data dari file CSV

### Transaksi Perpustakaan
- Peminjaman buku
- Pengembalian buku
- Perpanjangan masa pinjam
- Perhitungan denda otomatis berbasis hari kerja

### Notifikasi Otomatis
- Email reminder jatuh tempo
- WhatsApp reminder melalui WhatsApp Web
- Konfigurasi email dinamis melalui file JSON

### Laporan
- Laporan stok buku
- Laporan peminjaman harian
- Rekapitulasi transaksi
- Output laporan dalam format `.docx`

---

## Teknologi yang Digunakan

- Python 3.x
- Tkinter
- CSV File Handling
- `python-docx`
- `smtplib`
- `pywhatkit`
- `pyautogui`

---

## Struktur Repository

Perpustakaan-Digital-with-python-and-tkinter/
├── Perpus Digital .py
├── data_buku.csv
├── data_mahasiswa.csv
├── transaksi.csv
├── config_email.json
├── Laporan_Perpus/
│ ├── Stok_Buku/
│ ├── Transaksi_Harian/
│ └── Rekap_Lengkap/
└── README.md


---

## Cara Menjalankan Aplikasi

1. Clone repository:
```bash
git clone https://github.com/fahmi1804/Perpustakaan-Digital-with-python-and-tkinter.git
cd Perpustakaan-Digital-with-python-and-tkinter
pip install python-docx pywhatkit pyautogui
python "Perpus Digital .py"

atau juga bisa dengan mengklik 2 kali aplikasinya
