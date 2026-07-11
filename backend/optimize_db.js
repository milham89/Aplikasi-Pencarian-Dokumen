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
    console.log('Memulai optimasi database PostgreSQL...');

    // 1. Buat B-Tree Index pada file_id (SANGAT PENTING untuk Joins & File Previews)
    console.log('1. Membuat index pada file_id...');
    await client.query(`
      CREATE INDEX IF NOT EXISTS idx_document_rows_file_id ON document_rows (file_id);
    `);
    console.log('   Index file_id berhasil dibuat.');

    // 2. Aktifkan ekstensi pg_trgm untuk pencarian teks super cepat
    console.log('2. Mengaktifkan ekstensi pg_trgm...');
    try {
      await client.query(`CREATE EXTENSION IF NOT EXISTS pg_trgm;`);
      console.log('   Ekstensi pg_trgm berhasil diaktifkan.');

      // 3. Buat GIN Trigram Index pada cast teks JSONB untuk pencarian kata kunci LIKE/ILIKE cepat
      console.log('3. Membuat GIN Trigram Index pada (row_data::text)...');
      await client.query(`
        CREATE INDEX IF NOT EXISTS idx_document_rows_trgm 
        ON document_rows USING gin ((row_data::text) gin_trgm_ops);
      `);
      console.log('   GIN Trigram Index berhasil dibuat.');
    } catch (extErr) {
      console.warn('   Peringatan: Gagal mengaktifkan pg_trgm atau membuat trigram index:', extErr.message);
      console.log('   Menggunakan fallback B-Tree index standar.');
    }

    // 4. Jalankan ANALYZE untuk memperbarui statistik query planner
    console.log('4. Menjalankan ANALYZE untuk memperbarui statistik tabel...');
    await client.query('ANALYZE document_rows;');
    await client.query('ANALYZE uploaded_files;');
    console.log('   ANALYZE selesai.');

    console.log('Optimasi database selesai dengan sukses!');
  } catch (err) {
    console.error('Gagal menjalankan optimasi database:', err);
  } finally {
    client.release();
    pool.end();
  }
}

optimize();
