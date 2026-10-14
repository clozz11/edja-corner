-- Opsional: skema utama sudah dibuat otomatis oleh sequelize.sync({ alter: true }).
-- File ini hanya referensi manual (PostgreSQL), idempotent dan aman diulang.

ALTER TABLE product
  ADD COLUMN IF NOT EXISTS harga_beli INTEGER NOT NULL DEFAULT 0;

ALTER TABLE orders
  ADD COLUMN IF NOT EXISTS payment_status VARCHAR(20) NOT NULL DEFAULT 'pending';

ALTER TABLE orders
  ADD COLUMN IF NOT EXISTS payment_method VARCHAR(50) DEFAULT 'manual';

-- Pesanan lama dianggap sudah terverifikasi agar alur tidak terblokir
UPDATE orders SET payment_status = 'verified' WHERE payment_status = 'pending' AND status NOT IN ('pending', 'cancelled');
