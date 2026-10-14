const path = require('path');
const { Sequelize } = require('sequelize');

require('dotenv').config({ path: path.join(__dirname, '../../.env') });

function trim(value) {
  return typeof value === 'string' ? value.trim() : value;
}

function resolveDbConfig() {
  // Render & most Postgres hosts inject a single connection URL.
  const url = trim(
    process.env.DATABASE_URL ||
      process.env.POSTGRES_URL ||
      process.env.POSTGRESQL_URL
  );

  if (url) {
    return { url, useUrl: true };
  }

  const database = trim(
    process.env.DB_NAME || process.env.PGDATABASE || process.env.DB_DATABASE
  );
  const username = trim(
    process.env.DB_USER || process.env.PGUSER || process.env.DB_USERNAME
  );
  const password =
    process.env.DB_PASSWORD ?? process.env.PGPASSWORD ?? process.env.DB_PASS ?? '';
  const host = trim(
    process.env.DB_HOST || process.env.PGHOST || process.env.DB_HOSTNAME
  );
  const port = Number(trim(process.env.DB_PORT || process.env.PGPORT || 5432));

  return { database, username, password, host, port, useUrl: false };
}

function validateConfig(cfg) {
  if (cfg.useUrl) {
    return;
  }

  const missing = [];
  if (!cfg.database) missing.push('DB_NAME / PGDATABASE');
  if (!cfg.username) missing.push('DB_USER / PGUSER');
  if (!cfg.host) missing.push('DB_HOST / PGHOST');
  if (!cfg.port || Number.isNaN(cfg.port)) missing.push('DB_PORT / PGPORT');

  if (missing.length > 0) {
    throw new Error(
      `Variabel database belum lengkap: ${missing.join(', ')}. ` +
        'Di Render, gunakan "Internal Database URL" dari service PostgreSQL sebagai DATABASE_URL.'
    );
  }
}

const cfg = resolveDbConfig();
validateConfig(cfg);

const isProduction = process.env.NODE_ENV === 'production';

// Managed Postgres (Render, Neon, Supabase, dll.) butuh SSL.
// Lokal (localhost / 127.0.0.1) tidak. Deteksi dari host; untuk URL,
// anggap butuh SSL kecuali diarahkan ke localhost.
function needsSsl() {
  if (cfg.useUrl) {
    const u = cfg.url.toLowerCase();
    if (u.includes('sslmode=disable')) return false;
    if (u.includes('localhost') || u.includes('127.0.0.1')) return false;
    return true;
  }
  const host = (cfg.host || '').toLowerCase();
  if (!host || host === 'localhost' || host === '127.0.0.1') return false;
  return true;
}

const useSsl = needsSsl();

const sequelizeOptions = {
  dialect: 'postgres',
  logging: isProduction ? false : console.log,
  dialectOptions: useSsl
    ? {
        ssl: {
          require: true,
          rejectUnauthorized: false,
        },
      }
    : {},
};

const sequelize = cfg.useUrl
  ? new Sequelize(cfg.url, sequelizeOptions)
  : new Sequelize(cfg.database, cfg.username, cfg.password, {
      ...sequelizeOptions,
      host: cfg.host,
      port: cfg.port,
    });

module.exports = sequelize;
