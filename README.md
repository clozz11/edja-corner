# 🏪 Edja Corner — Minimarket Online

Aplikasi **full-stack** untuk digitalisasi toko kelontong **Edja Corner**: katalog belanja
online untuk pelanggan, kasir POS untuk staf, dan dashboard manajemen untuk pemilik —
lengkap dengan manajemen produk, stok, pesanan, dan verifikasi pembayaran.

> Versi online dari toko konvensional (seperti Alfamart/Indomaret) — dibangun sebagai
> proyek rekayasa perangkat lunak.

## ✨ Fitur

- **Katalog Online** — pelanggan menelusuri produk, menambah ke keranjang, checkout, memilih metode pembayaran, dan melacak status pesanan.
- **POS Kasir** — transaksi penjualan di tempat untuk staf kasir.
- **Dashboard Owner** — metrik penjualan, manajemen produk & kategori, stok masuk/retur, verifikasi pembayaran, dan manajemen staf.
- **Otentikasi berbasis peran** — Owner, Admin/Kasir, Gudang, Customer (JWT).

## 🧱 Arsitektur

```
edja-corner/
├── backend/     API REST — Node.js + Express + Sequelize (PostgreSQL)
└── frontend/    SPA — React + Vite
```

| Lapisan   | Teknologi                                           |
|-----------|-----------------------------------------------------|
| Frontend  | React 18, Vite 5, React Router 6                    |
| Backend   | Node.js, Express 4, Sequelize 6                      |
| Database  | PostgreSQL                                           |
| Auth      | JSON Web Token (JWT) + bcrypt                        |

## 🚀 Menjalankan secara lokal

### Prasyarat
- Node.js 18+
- PostgreSQL 14+ berjalan lokal

### 1. Database
```sql
CREATE DATABASE edja_corner;
```

### 2. Backend
```bash
cd backend
npm install
cp .env.example .env        # lalu isi DATABASE_URL / JWT_SECRET
npm start                   # http://localhost:3000
```
Tabel dibuat otomatis oleh Sequelize saat pertama kali start.

### 3. Frontend
```bash
cd frontend
npm install
cp .env.example .env        # dev: biarkan VITE_API_URL kosong (pakai proxy)
npm run dev                 # http://localhost:5173
```

## 🌐 Deployment

| Komponen  | Platform             | Catatan                                        |
|-----------|----------------------|------------------------------------------------|
| Frontend  | Vercel               | set `VITE_API_URL = https://<backend>/api`     |
| Backend   | Render (Web Service) | set `DATABASE_URL`, `JWT_SECRET`, `NODE_ENV`   |
| Database  | Render (PostgreSQL)  | pakai *Internal Database URL* sebagai `DATABASE_URL` |

## 📂 Struktur Backend

```
backend/src/
├── app.js            Entry point + koneksi DB + registrasi route
├── config/           Konfigurasi database & migrasi
├── models/           Model Sequelize (User, Product, Order, dst.)
├── controllers/      Logika bisnis per domain
├── routes/           Definisi endpoint REST
├── middleware/       Auth & otorisasi peran
└── utils/            Mapper DB ↔ frontend, helper password
```

## 👤 Akun demo (seed lokal)

| Peran    | Email                 | Password    |
|----------|-----------------------|-------------|
| Owner    | owner@edja.local      | owner123    |
| Customer | daftar via halaman Register                     |

---

Dibuat oleh **DzakyRishan** — proyek portofolio full-stack.
