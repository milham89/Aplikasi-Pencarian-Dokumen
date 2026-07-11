const { Pool } = require('pg');
require('dotenv').config();

const pool = new Pool({
  user: process.env.DB_USER,
  host: process.env.DB_HOST,
  database: process.env.DB_NAME,
  password: process.env.DB_PASSWORD,
  port: process.env.DB_PORT,
});

async function optimize() {
  const client = await pool.connect();
  try {
    console.log('Memulai optimasi indeks query pratinjau database...');

    // 1. Buat Composite B-Tree Index untuk menghilangkan proses sorting berat pada 560.000+ baris
    console.log('1. Membuat composite index (file_id, sheet_name, row_number)...');
    await client.query(`
      CREATE INDEX IF NOT EXISTS idx_document_rows_file_sheet_row 
      ON document_rows (file_id, sheet_name, row_number);
    `);
    console.log('   Indeks composite berhasil dibuat.');

    // 2. Jalankan ANALYZE untuk memperbarui statistik query planner
    console.log('2. Menjalankan ANALYZE untuk memperbarui statistik tabel...');
    await client.query('ANALYZE document_rows;');
    console.log('   ANALYZE selesai.');

    console.log('Optimasi pratinjau database selesai dengan sukses!');
  } catch (err) {
    console.error('Gagal menjalankan optimasi:', err);
  } finally {
    client.release();
    pool.end();
  }
}

optimize();
