require('dotenv').config();
const { Pool } = require('pg');

const pool = new Pool({
  user: process.env.DB_USER,
  host: process.env.DB_HOST,
  database: process.env.DB_NAME,
  password: process.env.DB_PASSWORD,
  port: process.env.DB_PORT,
});

async function run() {
  try {
    const queryText = 'P011.A001.25';
    const codeMatch = queryText.trim().match(/^[pP](\d+)[\s\.]*([a-zA-Z])[\s\.]*(\d+)(?:[\s\.]*(\d{2,4}))?$/);
    
    const unitNum = codeMatch[1];
    const boxLetter = codeMatch[2].toUpperCase();
    const folderNum = parseInt(codeMatch[3], 10);
    const yearRaw = codeMatch[4];
    
    const folder6 = String(folderNum).padStart(6, '0');
    
    let years = [];
    if (yearRaw) {
      const year = yearRaw.length === 2 ? '20' + yearRaw : yearRaw;
      years.push(year);
    } else {
      years = ['2025', '2024', '2026', '2023', '2027', '2028', '2022', '2029', '2020', '2021'];
    }
    
    const keys = ['NO.BOKS', 'NO. BOKS', 'NO BOKS', 'NO_BOKS', 'KODE BOKS', 'KODE_BOKS', 'BOKS', 'BOX'];
    const containmentClauses = [];
    const queryParams = [];
    
    years.forEach(year => {
      const boxCode = `P${unitNum}.${boxLetter}.${folder6}.${year}`;
      keys.forEach(key => {
        queryParams.push(JSON.stringify({ [key]: boxCode }));
        containmentClauses.push(`r.row_data @> $${queryParams.length}`);
      });
    });
    
    const searchQuery = `
      SELECT 
        r.id, 
        r.file_id, 
        r.sheet_name, 
        r.row_number, 
        r.row_data, 
        f.filename, 
        f.uploaded_at,
        100 as relevance_score
      FROM document_rows r
      JOIN uploaded_files f ON r.file_id = f.id
      WHERE ${containmentClauses.join(' OR ')}
      LIMIT 150;
    `;

    console.log('Query:', searchQuery);
    console.log('Params:', queryParams);

    const start = Date.now();
    const res = await pool.query(searchQuery, queryParams);
    console.log(`Executed in ${Date.now() - start}ms. Found rows: ${res.rowCount}`);
    if (res.rowCount > 0) {
      console.log('First 5 rows:');
      res.rows.slice(0, 5).forEach((r, i) => {
        console.log(`Row ${i+1}:`, r.row_data['NO.BOKS'] || r.row_data['NO. BOKS']);
      });
    }
  } catch (err) {
    console.error(err);
  } finally {
    await pool.end();
  }
}

run();
