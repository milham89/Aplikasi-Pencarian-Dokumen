const { Pool } = require('pg');

const pool = new Pool({
  user: 'postgres',
  host: '127.0.0.1',
  database: 'excel_db',
  password: 'mysecretpassword',
  port: 5432
});

async function benchmark() {
  console.log('Testing query performance...');
  
  // 1. Measure current query time
  const start = Date.now();
  const res = await pool.query(`
    SELECT r.id, r.file_id, r.sheet_name, r.row_number
    FROM document_rows r
    WHERE r.row_data::text ILIKE '%divisi%' OR r.row_data::text ILIKE '%umum%'
    LIMIT 50;
  `);
  const elapsed = Date.now() - start;
  console.log(`Current query elapsed time: ${elapsed}ms (${res.rows.length} rows returned)`);

  // 2. Install pg_trgm extension if not present
  console.log('Installing pg_trgm extension...');
  await pool.query(`CREATE EXTENSION IF NOT EXISTS pg_trgm;`);

  // 3. Check if gin index exists
  console.log('Creating gin_trgm index on document_rows(row_data::text)...');
  const idxStart = Date.now();
  await pool.query(`
    CREATE INDEX IF NOT EXISTS idx_document_rows_trgm 
    ON document_rows 
    USING gin ((row_data::text) gin_trgm_ops);
  `);
  console.log(`Index created in ${Date.now() - idxStart}ms!`);

  // 4. Measure query time after index
  const start2 = Date.now();
  const res2 = await pool.query(`
    SELECT r.id, r.file_id, r.sheet_name, r.row_number
    FROM document_rows r
    WHERE r.row_data::text ILIKE '%divisi%' OR r.row_data::text ILIKE '%umum%'
    LIMIT 50;
  `);
  console.log(`Query elapsed time AFTER INDEX: ${Date.now() - start2}ms (${res2.rows.length} rows returned)`);

  await pool.end();
}

benchmark().catch(e => { console.error('Benchmark error:', e); pool.end(); });
