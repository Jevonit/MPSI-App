# Project Scheduling Application - CPM & Gantt Chart

Aplikasi desktop berbasis Python untuk manajemen dan penjadwalan proyek menggunakan metode CPM (Critical Path Method). Aplikasi ini membantu manajer proyek dalam menghitung jalur kritis, durasi proyek, serta memvisualisasikan jadwal dalam bentuk Network Diagram dan Gantt Chart.

## Fitur Utama

1.  **Manajemen Kegiatan Proyek**
    -   Tambah kegiatan baru dengan Nama, Durasi, dan Dependensi.
    -   Hapus kegiatan yang dipilih.
    -   Hapus semua kegiatan (Reset).

2.  **Analisis CPM (Critical Path Method)**
    -   Perhitungan otomatis untuk:
        -   ES (Earliest Start)
        -   EF (Earliest Finish)
        -   LS (Latest Start)
        -   LF (Latest Finish)
        -   Slack (Float)
    -   Identifikasi Jalur Kritis (Critical Path).
    -   Perhitungan Total Durasi Proyek.

3.  **Visualisasi Interaktif**
    -   **Network Diagram**: Menggambarkan hubungan antar kegiatan dalam bentuk graf node dan panah.
    -   **Gantt Chart**: Menampilkan jadwal kegiatan dalam timeline horizontal.
    -   Fitur Zoom (Scroll Mouse) dan Pan (Klik Kiri + Drag) pada diagram.

4.  **Import & Export Data**
    -   Import data kegiatan dari file Excel (.xlsx, .xls).
    -   Export hasil analisis dan data kegiatan ke file Excel.

## Prasyarat Sistem

Sebelum menjalankan aplikasi, pastikan komputer Anda telah terinstall:
-   **Python 3.x**: [Download Python](https://www.python.org/downloads/)
-   **Library Python**: pandas, numpy, networkx, matplotlib, openpyxl.

## Cara Instalasi

1.  **Download Source Code**
    Pastikan Anda memiliki file `Tugas.py` dan `requirements.txt` dalam satu folder.

2.  **Install Dependencies**
    Buka terminal atau command prompt (CMD/PowerShell) di folder aplikasi, lalu jalankan perintah berikut untuk menginstall library yang dibutuhkan:

    ```bash
    pip install -r requirements.txt
    ```

    Atau install secara manual:
    ```bash
    pip install pandas numpy networkx matplotlib openpyxl
    ```

## Cara Menggunakan Aplikasi

1.  **Menjalankan Aplikasi**
    Buka terminal di folder aplikasi dan jalankan perintah:
    ```bash
    python Tugas.py
    ```

2.  **Menambahkan Kegiatan**
    -   Isi **Nama Kegiatan**.
    -   Isi **Durasi** (dalam hari, angka positif).
    -   Isi **Dependensi** (opsional). Masukkan ID kegiatan prasyarat dipisahkan dengan koma (contoh: `1,2`). Jika kegiatan pertama, biarkan kosong.
    -   Klik tombol **Tambah Kegiatan**.

3.  **Melihat Hasil Analisis**
    Pindah ke tab di sebelah kanan:
    -   **CPM Analysis**: Melihat tabel detail perhitungan CPM dan jalur kritis.
    -   **Network Diagram**: Melihat visualisasi alur kerja proyek.
    -   **Gantt Chart**: Melihat jadwal pelaksanaan proyek.

4.  **Import Data dari Excel**
    -   Klik tombol **Import Excel**.
    -   Pilih file Excel yang berisi data kegiatan.
    -   Aplikasi akan mencoba mendeteksi kolom secara otomatis. Pastikan file Excel memiliki kolom yang merepresentasikan Nama, Durasi, dan Dependensi.

5.  **Export Data ke Excel**
    -   Klik tombol **Export Excel**.
    -   Pilih lokasi penyimpanan.
    -   File Excel akan berisi data kegiatan beserta hasil perhitungan CPM (ES, EF, LS, LF, dll).

## Format File Excel (Untuk Import)

Agar proses import berjalan lancar, disarankan menggunakan file Excel dengan header kolom sebagai berikut (case-insensitive):

| Nama Kolom (Variasi 1) | Nama Kolom (Variasi 2) | Deskripsi |
| :--- | :--- | :--- |
| Nama / Kegiatan | Name / Activity | Nama dari kegiatan proyek |
| Durasi / Waktu | Duration / Time | Durasi pengerjaan (angka) |
| Dependensi / Prasyarat | Dependencies / Predecessor | ID kegiatan prasyarat (dipisah koma) |

## Kredit

Dibuat untuk memenuhi Tugas UTS Mata Kuliah Manajemen Proyek Sistem Informasi (MPSI).
**Pengembang**: Jevon Ivander Thomas
