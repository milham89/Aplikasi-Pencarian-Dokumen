const { Pool } = require('pg');
require('dotenv').config();

const pool = new Pool({
  user: process.env.DB_USER,
  host: process.env.DB_HOST,
  database: process.env.DB_NAME,
  password: process.env.DB_PASSWORD,
  port: process.env.DB_PORT,
});

async function migrate() {
  const client = await pool.connect();
  try {
    console.log('Memulai migrasi kolom row_count pada uploaded_files...');

    // 1. Tambah kolom row_count jika belum ada
    console.log('1. Menambahkan kolom row_count...');
    await client.query(`
      ALTER TABLE uploaded_files ADD COLUMN IF NOT EXISTS row_count INTEGER DEFAULT 0;
    `);
    console.log('   Kolom row_count berhasil ditambahkan.');

    // 2. Isi data row_count untuk file yang sudah ada di database
    console.log('2. Mengisi data row_count untuk file lama (agregasi awal)...');
    const updateResult = await client.query(`
      UPDATE uploaded_files f
      SET row_count = (
        SELECT COUNT(*)::int 
        FROM document_rows r 
        WHERE r.file_id = f.id
      );
    `);
    console.log(`   Migrasi sukses. Berhasil memperbarui data untuk ${updateResult.rowCount} file.`);

    console.log('Migrasi database selesai!');
  } catch (err) {
    console.error('Gagal menjalankan migrasi:', err);
  } finally {
    client.release();
    pool.end();
  }
}

migrate();
