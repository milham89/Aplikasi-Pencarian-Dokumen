const { Pool } = require('pg');

const pool = new Pool({
  user: 'postgres',
  host: '127.0.0.1',
  database: 'excel_db',
  password: 'mysecretpassword',
  port: 5432
});

async function testQueryPerformance() {
  const keywords = ['divisi', 'umum', 'grup', 'manajemen', 'asset'];
  console.log('Testing 5-keyword search for:', keywords);

  // Strategy A: Current query (OR condition, calculate score on all matching rows)
  const startA = Date.now();
  const queryParamsA = keywords.map(k => `%${k}%`);
  const singleScores = keywords.map((_, i) => `(CASE WHEN r.row_data::text ILIKE $${i+1} THEN 10 ELSE 0 END)`).join(' + ');
  const andCond = keywords.map((_, i) => `r.row_data::text ILIKE $${i+1}`).join(' AND ');
  const scoreExprA = `(${singleScores} + CASE WHEN (${andCond}) THEN 100 ELSE 0 END)`;
  const whereClausesA = keywords.map((_, i) => `r.row_data::text ILIKE $${i+1}`).join(' OR ');
  
  const resA = await pool.query(`
    SELECT r.id, r.file_id, r.sheet_name, r.row_number, ${scoreExprA} as score
    FROM document_rows r
    WHERE ${whereClausesA}
    ORDER BY score DESC, r.row_number ASC
    LIMIT 200;
  `, queryParamsA);
  console.log(`Strategy A (current) elapsed: ${Date.now() - startA}ms | Rows returned: ${resA.rows.length}`);

  // Strategy B: Optimized 2-step (Try AND first, then OR fallback)
  const startB = Date.now();
  const whereAnd = keywords.slice(0, 3).map((_, i) => `r.row_data::text ILIKE $${i+1}`).join(' AND ');
  const resB = await pool.query(`
    SELECT r.id, r.file_id, r.sheet_name, r.row_number
    FROM document_rows r
    WHERE ${whereAnd}
    ORDER BY r.row_number ASC
    LIMIT 100;
  `, queryParamsA.slice(0, 3));
  console.log(`Strategy B (AND filter first 3 terms) elapsed: ${Date.now() - startB}ms | Rows returned: ${resB.rows.length}`);

  await pool.end();
}

testQueryPerformance().catch(e => { console.error(e); pool.end(); });
