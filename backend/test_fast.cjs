const { Pool } = require('pg');

const pool = new Pool({
  user: 'postgres',
  host: '127.0.0.1',
  database: 'excel_db',
  password: 'mysecretpassword',
  port: 5432
});

async function testFast() {
  const keywords = ['divisi', 'umum', 'grup', 'manajemen', 'asset'];
  console.log('Testing Fast Tiered Search for:', keywords);

  // Step 1: Strict AND query for all keywords
  const start1 = Date.now();
  const params1 = keywords.map(k => `%${k}%`);
  const whereAnd = keywords.map((_, i) => `r.row_data::text ILIKE $${i+1}`).join(' AND ');
  const res1 = await pool.query(`
    SELECT r.id, r.file_id, r.sheet_name, r.row_number, r.row_data, f.filename
    FROM document_rows r
    JOIN uploaded_files f ON r.file_id = f.id
    WHERE ${whereAnd}
    ORDER BY r.row_number ASC
    LIMIT 60;
  `, params1);
  console.log(`Step 1 (Strict AND) elapsed: ${Date.now() - start1}ms | Rows found: ${res1.rows.length}`);

  if (res1.rows.length === 0) {
    // Step 2: Fallback to AND of top 2 longest keywords
    const sortedKws = [...keywords].sort((a,b) => b.length - a.length);
    const top2 = sortedKws.slice(0, 2);
    console.log('Step 2 fallback with top longest keywords:', top2);
    const start2 = Date.now();
    const params2 = top2.map(k => `%${k}%`);
    const whereAnd2 = top2.map((_, i) => `r.row_data::text ILIKE $${i+1}`).join(' AND ');
    const res2 = await pool.query(`
      SELECT r.id, r.file_id, r.sheet_name, r.row_number, r.row_data, f.filename
      FROM document_rows r
      JOIN uploaded_files f ON r.file_id = f.id
      WHERE ${whereAnd2}
      ORDER BY r.row_number ASC
      LIMIT 60;
    `, params2);
    console.log(`Step 2 (Longest 2 AND) elapsed: ${Date.now() - start2}ms | Rows found: ${res2.rows.length}`);
  }

  await pool.end();
}

testFast().catch(e => { console.error(e); pool.end(); });
