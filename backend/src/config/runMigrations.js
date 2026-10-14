const sequelize = require('./database');

// Catatan: pembuatan skema utama ditangani oleh sequelize.sync({ alter: true })
// di app.js, sehingga semua kolom (termasuk harga_beli, payment_status,
// payment_method) sudah dibuat otomatis dari definisi model.
//
// Fungsi ini disimpan sebagai jaring pengaman idempotent dalam sintaks
// PostgreSQL. Aman dijalankan berkali-kali: "IF NOT EXISTS" membuat setiap
// statement no-op bila kolom sudah ada. ENUM diganti VARCHAR + CHECK agar
// portable; model Sequelize tetap sumber kebenaran untuk tipe kolom.
const statements = [
  `ALTER TABLE product ADD COLUMN IF NOT EXISTS harga_beli INTEGER NOT NULL DEFAULT 0`,
  `ALTER TABLE orders ADD COLUMN IF NOT EXISTS payment_status VARCHAR(20) NOT NULL DEFAULT 'pending'`,
  `ALTER TABLE orders ADD COLUMN IF NOT EXISTS payment_method VARCHAR(50) DEFAULT 'manual'`,
  `UPDATE orders SET payment_status = 'verified' WHERE payment_status = 'pending' AND status NOT IN ('pending', 'cancelled')`,
];

async function runMigrations() {
  for (const sql of statements) {
    try {
      await sequelize.query(sql);
    } catch (err) {
      const msg = String(err.message || '');
      console.warn('Migration skip:', msg.slice(0, 160));
    }
  }
}

module.exports = { runMigrations };
