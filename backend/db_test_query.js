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
    const searchTerms = queryText.trim().split(/[\s\.]+/).filter(Boolean);
    
    console.log('Search terms:', searchTerms);

    const whereClauses = [];
    const queryParams = [];
    
    searchTerms.forEach((term) => {
      const match = term.match(/^([a-zA-Z]+)(\d+)$/);
      if (match) {
        const letters = match[1];
        const digits = match[2];
        
        queryParams.push(`%${term}%`, `%${letters}%`, `%${digits}%`);
        const idxTerm = queryParams.length - 2;
        const idxLetters = queryParams.length - 1;
        const idxDigits = queryParams.length;
        
        whereClauses.push(`(r.row_data::text ILIKE $${idxTerm} OR (r.row_data::text ILIKE $${idxLetters} AND r.row_data::text ILIKE $${idxDigits}))`);
      } else {
        queryParams.push(`%${term}%`);
        whereClauses.push(`r.row_data::text ILIKE $${queryParams.length}`);
      }
    });

    const searchQuery = `
      SELECT 
        r.id, 
        r.file_id, 
        r.sheet_name, 
        r.row_number, 
        r.row_data, 
        f.filename, 
        f.uploaded_at
      FROM document_rows r
      JOIN uploaded_files f ON r.file_id = f.id
      WHERE ${whereClauses.join(' AND ')}
      LIMIT 200;
    `;

    console.log('Generated SQL:\n', searchQuery);
    console.log('Params:', queryParams);

    const res = await pool.query(searchQuery, queryParams);
    console.log('Results count:', res.rowCount);
    if (res.rowCount > 0) {
      console.log('First match:', JSON.stringify(res.rows[0].row_data));
    }
  } catch (err) {
    console.error(err);
  } finally {
    await pool.end();
  }
}

run();
