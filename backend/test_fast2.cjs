const { Pool } = require('pg');

const pool = new Pool({
  user: 'postgres',
  host: '127.0.0.1',
  database: 'excel_db',
  password: 'mysecretpassword',
  port: 5432
});

async function testFast2() {
  const rawQuery = "tampilkan data divisi umum grup manajemen asset";
  const searchTerms = rawQuery.split(/[\s\.]+/).filter(Boolean);
  const stopWords = new Set([
    'tampilkan', 'data', 'pada', 'yang', 'dan', 'di', 'untuk', 'ke', 'dari', 
    'adalah', 'itu', 'dengan', 'ini', 'oleh', 'seperti', 'adapun', 'atau', 
    'sebagai', 'tentang', 'yaitu', 'ia', 'kami', 'mereka', 'saya', 'anda', 
    'dia', 'kita', 'tahun', 'bulan', 'hari', 'file', 'berkas', 'dokumen', 
    'tabel', 'sheet', 'cari', 'temukan', 'info', 'informasi', 'lihat', 'berapa',
    'apakah', 'siapa', 'bagaimana', 'apa', 'tolong', 'jelaskan', 'rangkum', 'semua',
    'daftar', 'total', 'ada'
  ]);
  const keywords = searchTerms.filter(t => !stopWords.has(t.toLowerCase()) && t.length > 2);
  
  // Sort keywords by length descending (longest = most specific first)
  const sortedKeywords = [...keywords].sort((a,b) => b.length - a.length);
  // Pick top 3 longest/most specific keywords (e.g. manajemen, asset, divisi)
  const topKeywords = sortedKeywords.slice(0, 3);
  console.log('Original keywords:', keywords);
  console.log('Top specific keywords for fast query:', topKeywords);

  const start = Date.now();
  const params = topKeywords.map(k => `%${k}%`);
  const whereAnd = topKeywords.map((_, i) => `r.row_data::text ILIKE $${i+1}`).join(' AND ');
  const res = await pool.query(`
    SELECT r.id, r.file_id, r.sheet_name, r.row_number, r.row_data, f.filename
    FROM document_rows r
    JOIN uploaded_files f ON r.file_id = f.id
    WHERE ${whereAnd}
    ORDER BY r.row_number ASC
    LIMIT 60;
  `, params);
  console.log(`Top 3 specific AND query elapsed: ${Date.now() - start}ms | Rows found: ${res.rows.length}`);

  await pool.end();
}

testFast2().catch(e => { console.error(e); pool.end(); });
