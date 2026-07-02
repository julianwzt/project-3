# Aplikasi Pendataan Mahasiswa (VB6 + PostgreSQL)

Aplikasi desktop berbasis Visual Basic 6.0 yang mengimplementasikan arsitektur Client-Server dengan koneksi langsung (direct connection) ke database PostgreSQL.

## Spesifikasi Teknis
- **Frontend & Backend:** Visual Basic 6.0 (Desktop Application)
- **Database:** PostgreSQL
- **Arsitektur:** Client-Server (Tanpa API) dengan pendekatan konsep MVC (Model-View-Controller).
- **Metode Akses:** ADODB (ActiveX Data Objects) via ODBC Driver 32-bit.
- **Logika Bisnis:** Semua query dan eksekusi berjalan di sisi aplikasi.

## Fitur Utama
- **CRUD Operasional:** Kelola data Mahasiswa dengan relasi ke tabel Jurusan.
- **Validasi Input:** Sinkronisasi otomatis antara Tanggal Lahir dan Umur (mencegah salah format).
- **Dynamic Dropdown:** Pemilihan Fakultas dan Jenjang otomatis menyesuaikan pilihan Jurusan dari database.
- **Pencarian Cepat:** Filter data realtime berdasarkan Nama atau NIM.
- **Export Data:** Menyimpan laporan data ke format Excel (`.xls`).

## Panduan Instalasi & Setup
1. Siapkan database PostgreSQL.
2. Install driver **psqlODBC 32-bit (x86)** (wajib 32-bit meskipun OS 64-bit).
3. Buka konfigurasi **ODBC Data Sources (32-bit)** di Windows.
4. Buat *System DSN* baru menggunakan driver PostgreSQL dengan nama: **KoneksiAkademik**.
5. Sesuaikan *username* dan *password* database di dalam *Module* koneksi VB6.
6. Buka `Project3.vbp` dan jalankan aplikasi.