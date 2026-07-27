const pg = require('pg');
const Pool = pg.Pool;

const pool = new Pool({
  user: 'postgres',
  host: '127.0.0.1',
  database: 'excel_db',
  password: 'mysecretpassword',
  port: 5432
});

async function run() {
  // 1. Distribution by file/sheet for ILHAM
  const r1 = await pool.query(`
    SELECT r.file_id, f.filename, r.sheet_name, COUNT(*) as cnt
    FROM document_rows r JOIN uploaded_files f ON r.file_id=f.id
    WHERE r.row_data::text ILIKE $1
    GROUP BY r.file_id,f.filename,r.sheet_name
    ORDER BY r.file_id, r.sheet_name
  `, ['%ILHAM%']);
  console.log('=== Files/Sheets containing ILHAM ===');
  r1.rows.forEach(r => console.log(`File ${r.file_id} | ${r.filename.substring(0,40)} | Sheet: ${r.sheet_name} | Count: ${r.cnt}`));

  // 2. Rows with ILHAM + SOLEHUDIN 
  const r2 = await pool.query(`
    SELECT r.file_id, f.filename, r.sheet_name, r.row_number
    FROM document_rows r JOIN uploaded_files f ON r.file_id=f.id
    WHERE r.row_data::text ILIKE $1 AND r.row_data::text ILIKE $2
    ORDER BY r.file_id, r.sheet_name, r.row_number
    LIMIT 20
  `, ['%ILHAM%', '%SOLEHUDIN%']);
  console.log('\n=== Rows with ILHAM + SOLEHUDIN ===');
  r2.rows.forEach(r => console.log(`File ${r.file_id} | ${r.filename.substring(0,35)} | Sheet: ${r.sheet_name} | Row: ${r.row_number}`));

  // 3. Simulate chat candidates query: keywords [ILHAM, SOLEHUDIN]
  const r3 = await pool.query(`
    SELECT r.file_id, f.filename, r.sheet_name, r.row_number,
      (CASE WHEN r.row_data::text ILIKE $1 THEN 10 ELSE 0 END +
       CASE WHEN r.row_data::text ILIKE $2 THEN 10 ELSE 0 END) as score
    FROM document_rows r JOIN uploaded_files f ON r.file_id=f.id
    WHERE (r.row_data::text ILIKE $1 OR r.row_data::text ILIKE $2)
    ORDER BY score DESC, f.uploaded_at DESC, r.row_number ASC
    LIMIT 30
  `, ['%ILHAM%', '%SOLEHUDIN%']);
  console.log('\n=== Top 30 chat candidates (ILHAM OR SOLEHUDIN) ===');
  const fileCounts = {};
  r3.rows.forEach(r => {
    const key = `${r.file_id}|${r.sheet_name}`;
    if (!fileCounts[key]) fileCounts[key] = 0;
    fileCounts[key]++;
    if (fileCounts[key] <= 3) { // show max 3 per file/sheet
      console.log(`Score ${r.score} | File ${r.file_id} | ${r.filename.substring(0,30)} | Sheet: ${r.sheet_name} | Row: ${r.row_number}`);
    }
  });

  await pool.end();
}

run().catch(e => { console.error(e.message); pool.end(); });
