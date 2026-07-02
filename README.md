# Aplikasi Pendataan Mahasiswa (VB6 + PostgreSQL + Minikube)

Aplikasi desktop berbasis Visual Basic 6.0 yang mengimplementasikan arsitektur *Client-Server* dengan koneksi langsung (*direct connection*) ke database PostgreSQL yang berjalan di dalam *cluster* Kubernetes (Minikube).

## Spesifikasi Teknis
- **Frontend & Backend:** Visual Basic 6.0 (Desktop Application)
- **Database:** PostgreSQL (berjalan di Minikube)
- **Arsitektur:** *Client-Server* (Tanpa API) dengan konsep MVC (*Model-View-Controller*).
- **Metode Akses:** ADODB (*ActiveX Data Objects*) via ODBC Driver 32-bit.

## Fitur Utama
- **CRUD Operasional:** Mengelola data Mahasiswa dan Jurusan.
- **Validasi Cerdas:** Sinkronisasi otomatis antara Tanggal Lahir dan Umur.
- **Dynamic Dropdown:** Data Fakultas dan Jenjang terisi otomatis mengikuti pilihan Jurusan.
- **Pencarian Realtime:** Filter data berdasarkan Nama atau NIM secara langsung.
- **Export Data:** Fitur ekspor laporan ke format Excel (`.xls`) dengan penentu lokasi penyimpanan.

## Panduan Setup & Deployment
### 1. Persiapan Database (Minikube)
1. Jalankan cluster Minikube: `minikube start`.
2. Buat *ConfigMap* dari file init.sql: 
   `kubectl create configmap db-init --from-file=init.sql`.
3. Deploy database dan service: `kubectl apply -f k8s-db.yaml`.
4. Tunggu pod pg-db berstatus **Running**: `kubectl get pods`.

### 2. Koneksi Aplikasi
1. Jalankan *tunneling* port ke database: `kubectl port-forward svc/pg-svc 5433:5432`.
2. Install **psqlODBC 32-bit (x86)**.
3. Konfigurasi **ODBC Data Sources (32-bit)**:
   - Tambahkan *System DSN* baru ("PostgreSQL Unicode").
   - Server: `localhost`, Port: `5433`, User: `postgres`, Password: `password_database_kamu`.
   - Lakukan **Test Connection** hingga berhasil.