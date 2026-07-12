const express = require('express');
const cors = require('cors');
const multer = require('multer');
const ExcelJS = require('exceljs');
const { Pool } = require('pg');
const fs = require('fs');
const path = require('path');
const { UAParser } = require('ua-parser-js');
require('dotenv').config();

const app = express();
const port = process.env.PORT || 5000;

// In-memory job store to track file processing progress
const uploadJobs = {};

// Ensure uploads directory exists
const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir);
}

// Helper: Parse device info from request
const getDeviceInfo = (req) => {
  const ua = req.headers['user-agent'] || '';
  const parser = new UAParser(ua);
  const result = parser.getResult();

  const ip =
    req.headers['x-forwarded-for']?.split(',')[0]?.trim() ||
    req.socket?.remoteAddress ||
    req.connection?.remoteAddress ||
    'Tidak diketahui';

  const browser = result.browser.name
    ? `${result.browser.name} ${result.browser.version || ''}`.trim()
    : 'Browser Tidak Dikenal';

  const os = result.os.name
    ? `${result.os.name} ${result.os.version || ''}`.trim()
    : 'OS Tidak Diketahui';

  const deviceType = result.device.type || 'Desktop';
  const deviceVendor = result.device.vendor || '';
  const deviceModel = result.device.model || '';
  const deviceLabel = [deviceVendor, deviceModel].filter(Boolean).join(' ') || deviceType;

  const engine = result.engine.name ? `${result.engine.name} ${result.engine.version || ''}`.trim() : '';

  return {
    ip,
    browser,
    os,
    device_type: deviceType.charAt(0).toUpperCase() + deviceType.slice(1),
    device_label: deviceLabel,
    engine,
    user_agent: ua
  };
};

// Middleware
app.use(cors());
app.use(express.json());
app.use((req, res, next) => {
  console.log(`[${new Date().toISOString()}] ${req.method} ${req.url}`);
  next();
});

// Configure Multer for disk storage (avoids Out-of-Memory on large files)
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    cb(null, `${Date.now()}-${file.originalname}`);
  }
});

const upload = multer({
  storage: storage,
  limits: { fileSize: 250 * 1024 * 1024 }, // 250MB limit
  fileFilter: (req, file, cb) => {
    if (
      file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
      file.mimetype === 'application/vnd.ms-excel' ||
      file.originalname.endsWith('.xlsx') ||
      file.originalname.endsWith('.xls')
    ) {
      cb(null, true);
    } else {
      cb(new Error('Hanya file Excel (.xlsx atau .xls) yang diperbolehkan!'));
    }
  }
});

// PostgreSQL Database Connection Pool
const pool = new Pool({
  user: process.env.DB_USER,
  host: process.env.DB_HOST,
  database: process.env.DB_NAME,
  password: process.env.DB_PASSWORD,
  port: process.env.DB_PORT,
});

// Database Migration on startup
const initDatabase = async () => {
  const client = await pool.connect();
  try {
    console.log('Memeriksa dan membuat tabel database jika belum ada...');
    
    // Create uploaded_files table
    await client.query(`
      CREATE TABLE IF NOT EXISTS uploaded_files (
        id SERIAL PRIMARY KEY,
        filename VARCHAR(255) NOT NULL,
        sheet_names TEXT[] NOT NULL,
        uploaded_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
        row_count INTEGER DEFAULT 0
      );
    `);

    // Create document_rows table with GIN Index for dynamic JSONB search
    await client.query(`
      CREATE TABLE IF NOT EXISTS document_rows (
        id SERIAL PRIMARY KEY,
        file_id INTEGER REFERENCES uploaded_files(id) ON DELETE CASCADE,
        sheet_name VARCHAR(100) NOT NULL,
        row_number INTEGER NOT NULL,
        row_data JSONB NOT NULL
      );
    `);

    // Create activity_logs table with device tracking
    await client.query(`
      CREATE TABLE IF NOT EXISTS activity_logs (
        id SERIAL PRIMARY KEY,
        activity_type VARCHAR(50) NOT NULL,
        activity_details TEXT NOT NULL,
        ip_address VARCHAR(100),
        browser VARCHAR(150),
        os VARCHAR(150),
        device_type VARCHAR(50),
        device_label VARCHAR(200),
        engine VARCHAR(100),
        user_agent TEXT,
        created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP
      );
    `);

    // Add device columns to existing table if they don't exist (migration)
    const deviceCols = ['ip_address VARCHAR(100)', 'browser VARCHAR(150)', 'os VARCHAR(150)', 'device_type VARCHAR(50)', 'device_label VARCHAR(200)', 'engine VARCHAR(100)', 'user_agent TEXT'];
    for (const col of deviceCols) {
      const colName = col.split(' ')[0];
      await client.query(`ALTER TABLE activity_logs ADD COLUMN IF NOT EXISTS ${colName} ${col.split(' ').slice(1).join(' ')};`);
    }

    // Create index on JSONB if it doesn't exist
    await client.query(`
      CREATE INDEX IF NOT EXISTS idx_document_rows_data ON document_rows USING gin (row_data);
    `);

    // Create B-Tree Index on file_id
    await client.query(`
      CREATE INDEX IF NOT EXISTS idx_document_rows_file_id ON document_rows (file_id);
    `);

    // Create Composite B-Tree Index on (file_id, sheet_name, row_number) for fast previews
    await client.query(`
      CREATE INDEX IF NOT EXISTS idx_document_rows_file_sheet_row ON document_rows (file_id, sheet_name, row_number);
    `);

    // Enable pg_trgm extension and create GIN Trigram Index
    try {
      await client.query(`CREATE EXTENSION IF NOT EXISTS pg_trgm;`);
      await client.query(`
        CREATE INDEX IF NOT EXISTS idx_document_rows_trgm 
        ON document_rows USING gin ((row_data::text) gin_trgm_ops);
      `);
    } catch (trgmErr) {
      console.warn('Gagal mengaktifkan pg_trgm saat startup:', trgmErr.message);
    }

    // Create bookmarks table
    await client.query(`
      CREATE TABLE IF NOT EXISTS bookmarks (
        id SERIAL PRIMARY KEY,
        row_id INTEGER NOT NULL REFERENCES document_rows(id) ON DELETE CASCADE,
        created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
        CONSTRAINT unique_row_bookmark UNIQUE(row_id)
      );
    `);

    console.log('Inisialisasi database PostgreSQL sukses.');
  } catch (err) {
    console.error('Gagal melakukan inisialisasi database:', err.message);
  } finally {
    client.release();
  }
};

// Call DB Init
initDatabase();

// Function to log activities to database (with optional device info)
const logActivity = async (type, details, deviceInfo = null) => {
  try {
    if (deviceInfo) {
      await pool.query(
        `INSERT INTO activity_logs 
          (activity_type, activity_details, ip_address, browser, os, device_type, device_label, engine, user_agent)
         VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9);`,
        [type, details,
          deviceInfo.ip, deviceInfo.browser, deviceInfo.os,
          deviceInfo.device_type, deviceInfo.device_label, deviceInfo.engine, deviceInfo.user_agent
        ]
      );
    } else {
      await pool.query(
        `INSERT INTO activity_logs (activity_type, activity_details) VALUES ($1, $2);`,
        [type, details]
      );
    }
  } catch (err) {
    console.error('Gagal menulis log aktivitas:', err);
  }
};

// --- BACKGROUND EXCEL PROCESSOR ---

async function processFileInBackground(jobId, filePath, originalname) {
  console.log(`[Job ${jobId}] Mulai memproses file: "${originalname}"`);
  uploadJobs[jobId] = {
    status: 'processing',
    progress: 10,
    message: 'Menginisialisasi impor...',
    processedRows: 0,
    error: null
  };

  let fileId = null;
  let client = null;
  
  try {
    // 1. Connect to DB and start transaction
    client = await pool.connect();
    await client.query('BEGIN');

    // 2. Insert record in uploaded_files
    const fileInsertRes = await client.query(
      `INSERT INTO uploaded_files (filename, sheet_names) VALUES ($1, $2) RETURNING id;`,
      [originalname, []]
    );
    fileId = fileInsertRes.rows[0].id;
    console.log(`[Job ${jobId}] ID File di database: ${fileId}`);

    uploadJobs[jobId].progress = 20;
    uploadJobs[jobId].message = 'Membaca struktur file Excel...';

    // 3. Open ExcelJS workbook reader for streaming (low memory overhead)
    const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(filePath, {
      sharedStrings: 'cache',
      hyperlinks: 'ignore',
      styles: 'ignore'
    });

    const sheetNames = [];
    let totalRowsInserted = 0;

    // 4. Iterate over worksheets
    for await (const worksheetReader of workbookReader) {
      const sheetName = worksheetReader.name;
      sheetNames.push(sheetName);
      console.log(`[Job ${jobId}] Memproses worksheet: "${sheetName}"`);
      
      uploadJobs[jobId].message = `Membaca sheet: "${sheetName}"...`;
      
      let headers = null;
      let rowsToInsert = [];
      const BATCH_SIZE = 500;

      for await (const row of worksheetReader) {
        // Build values array
        const rowValues = [];
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          let val = cell.value;
          if (val && typeof val === 'object') {
            if (val.result !== undefined) val = val.result;
            else if (val.text !== undefined) val = val.text;
            else if (val instanceof Date) {
              // keep date
            } else {
              val = JSON.stringify(val);
            }
          }
          rowValues[colNumber - 1] = val;
        });

        // Set headers from the first non-empty row
        if (!headers) {
          headers = rowValues.map((h, idx) => h !== undefined && h !== null ? h.toString().trim() : `Kolom_${idx + 1}`);
          continue;
        }

        // Map row cells to header keys
        const rowObj = {};
        let hasData = false;
        
        headers.forEach((header, colIndex) => {
          const val = rowValues[colIndex];
          rowObj[header] = val !== undefined ? val : null;
          if (val !== null && val !== undefined && val !== '') {
            hasData = true;
          }
        });

        if (hasData) {
          rowsToInsert.push({
            rowNumber: row.number,
            rowData: rowObj
          });
        }

        // Flush batch if batch size reached
        if (rowsToInsert.length >= BATCH_SIZE) {
          await insertBatch(client, fileId, sheetName, rowsToInsert);
          totalRowsInserted += rowsToInsert.length;
          rowsToInsert = [];

          // Update job status
          uploadJobs[jobId].processedRows = totalRowsInserted;
          uploadJobs[jobId].message = `Mengimpor sheet "${sheetName}": ${totalRowsInserted.toLocaleString('id-ID')} baris...`;
          
          // Yield to Event Loop
          await new Promise(resolve => setImmediate(resolve));
        }
      }

      // Insert remaining rows in sheet
      if (rowsToInsert.length > 0) {
        await insertBatch(client, fileId, sheetName, rowsToInsert);
        totalRowsInserted += rowsToInsert.length;
        uploadJobs[jobId].processedRows = totalRowsInserted;
        uploadJobs[jobId].message = `Mengimpor sheet "${sheetName}": ${totalRowsInserted.toLocaleString('id-ID')} baris...`;
        await new Promise(resolve => setImmediate(resolve));
      }
    }

    // 5. Update sheet names and row_count in uploaded_files
    await client.query(
      `UPDATE uploaded_files SET sheet_names = $1, row_count = $2 WHERE id = $3;`,
      [sheetNames, totalRowsInserted, fileId]
    );

    // 6. Log activity
    await client.query(
      `INSERT INTO activity_logs (activity_type, activity_details) VALUES ($1, $2);`,
      ['upload', `Berhasil mengunggah berkas "${originalname}" (${totalRowsInserted.toLocaleString('id-ID')} baris data)`]
    );

    // 7. Commit transaction
    await client.query('COMMIT');
    console.log(`[Job ${jobId}] Impor selesai! Total baris tersimpan: ${totalRowsInserted}`);

    // Set job success
    uploadJobs[jobId].status = 'completed';
    uploadJobs[jobId].progress = 100;
    uploadJobs[jobId].message = `Selesai! Berhasil mengimpor ${totalRowsInserted.toLocaleString('id-ID')} baris data.`;
    uploadJobs[jobId].result = {
      fileId,
      filename: originalname,
      sheetsProcessed: sheetNames.length,
      rowsInserted: totalRowsInserted
    };

  } catch (err) {
    console.error(`[Job ${jobId}] Error saat memproses file Excel:`, err);
    if (client) {
      try {
        await client.query('ROLLBACK');
      } catch (rollbackErr) {
        console.error(`[Job ${jobId}] Rollback failed:`, rollbackErr);
      }
    }
    
    // Clean up DB if record was created
    if (fileId) {
      try {
        await pool.query(`DELETE FROM uploaded_files WHERE id = $1;`, [fileId]);
      } catch (delErr) {
        console.error(`[Job ${jobId}] Clean DB record failed:`, delErr);
      }
    }

    uploadJobs[jobId].status = 'error';
    uploadJobs[jobId].progress = 100;
    uploadJobs[jobId].error = err.message;
    uploadJobs[jobId].message = 'Gagal memproses file Excel: ' + err.message;

  } finally {
    if (client) {
      client.release();
    }
    // Delete temp file from disk
    try {
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
        console.log(`[Job ${jobId}] Berkas sementara dihapus dari disk.`);
      }
    } catch (fsErr) {
      console.error(`[Job ${jobId}] Gagal menghapus berkas sementara:`, fsErr);
    }
  }
}

// Helper to insert a batch of rows
async function insertBatch(client, fileId, sheetName, batch) {
  const queryValues = [];
  const queryPlaceholders = [];
  
  batch.forEach((row, batchIndex) => {
    const paramOffset = queryValues.length;
    queryPlaceholders.push(`($${paramOffset + 1}, $${paramOffset + 2}, $${paramOffset + 3}, $${paramOffset + 4})`);
    queryValues.push(fileId, sheetName, row.rowNumber, JSON.stringify(row.rowData));
  });

  const insertRowsQuery = `
    INSERT INTO document_rows (file_id, sheet_name, row_number, row_data)
    VALUES ${queryPlaceholders.join(', ')};
  `;
  
  await client.query(insertRowsQuery, queryValues);
}

// --- API ROUTES ---

// 1. Health check & DB connection status
app.get('/api/health', async (req, res) => {
  try {
    await pool.query('SELECT 1');
    // Log device access whenever a user opens the app (health check is called on page load)
    const device = getDeviceInfo(req);
    logActivity('access', `Perangkat mengakses aplikasi dari ${device.ip}`, device);
    res.json({ status: 'ok', database: 'connected' });
  } catch (err) {
    res.status(500).json({ status: 'error', database: err.message });
  }
});

// 1b. Dashboard Statistics
app.get('/api/stats', async (req, res) => {
  try {
    // Total files and rows
    const summaryRes = await pool.query(`
      SELECT 
        COUNT(*) AS total_files,
        COALESCE(SUM(row_count), 0) AS total_rows
      FROM uploaded_files;
    `);

    // Top 10 searched keywords (last 30 days)
    const topSearchRes = await pool.query(`
      SELECT activity_details, COUNT(*) AS count
      FROM activity_logs
      WHERE activity_type = 'search'
        AND created_at >= NOW() - INTERVAL '30 days'
      GROUP BY activity_details
      ORDER BY count DESC
      LIMIT 10;
    `);

    // Docs per unit (top 15)
    const unitRes = await pool.query(`
      SELECT row_data->>'KODE UNIT' AS unit, COUNT(*) AS count
      FROM document_rows
      WHERE row_data->>'KODE UNIT' IS NOT NULL
        AND row_data->>'KODE UNIT' <> ''
      GROUP BY row_data->>'KODE UNIT'
      ORDER BY count DESC
      LIMIT 15;
    `);

    // Daily activity last 7 days
    const dailyRes = await pool.query(`
      SELECT 
        DATE(created_at AT TIME ZONE 'Asia/Jakarta') AS day,
        activity_type,
        COUNT(*) AS count
      FROM activity_logs
      WHERE created_at >= NOW() - INTERVAL '7 days'
      GROUP BY day, activity_type
      ORDER BY day ASC;
    `);

    // Total searches, uploads, deletes, access today
    const todayRes = await pool.query(`
      SELECT activity_type, COUNT(*) AS count
      FROM activity_logs
      WHERE DATE(created_at AT TIME ZONE 'Asia/Jakarta') = CURRENT_DATE
      GROUP BY activity_type;
    `);

    // Unique devices today
    const devicesRes = await pool.query(`
      SELECT DISTINCT ip_address, browser, os, device_type
      FROM activity_logs
      WHERE DATE(created_at AT TIME ZONE 'Asia/Jakarta') = CURRENT_DATE
        AND ip_address IS NOT NULL;
    `);

    res.json({
      summary: summaryRes.rows[0],
      topSearches: topSearchRes.rows,
      docsPerUnit: unitRes.rows,
      dailyActivity: dailyRes.rows,
      todayActivity: todayRes.rows,
      devicesOnline: devicesRes.rows
    });
  } catch (err) {
    console.error('Error stats:', err);
    res.status(500).json({ error: err.message });
  }
});

// 2. Upload Excel File (Triggers background processing)
app.post('/api/upload', upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'Tidak ada file yang diunggah!' });
  }

  const jobId = Date.now().toString();
  console.log(`[Upload] File upload masuk. Memulai Job ID: ${jobId}`);

  // Start background process
  processFileInBackground(jobId, req.file.path, req.file.originalname);

  // Respond immediately with jobId
  res.json({
    success: true,
    message: 'File berhasil diunggah dan sedang diproses di latar belakang.',
    jobId: jobId
  });
});

// 3. Get background upload job status
app.get('/api/upload/status/:jobId', (req, res) => {
  const job = uploadJobs[req.params.jobId];
  if (!job) {
    return res.status(404).json({ error: 'Proses impor tidak ditemukan.' });
  }
  res.json(job);
});

// 4. Search document rows by search query
app.get('/api/search', async (req, res) => {
  const queryText = req.query.q;
  const fileId = req.query.fileId;
  const filterSheet = req.query.sheet; // Optional sheet filter
  const filterUnit = req.query.unit;   // Optional unit filter

  if ((!queryText || queryText.trim() === '') && !fileId) {
    return res.json({ results: [] });
  }

  try {
    let searchQuery;
    let queryParams;

    if (fileId) {
      // Fetch all rows for a specific file (incorporating optional filters and bookmarks status)
      let filterClauses = ['r.file_id = $1'];
      queryParams = [parseInt(fileId)];

      if (filterSheet && filterSheet.trim() !== '') {
        queryParams.push(filterSheet.trim());
        filterClauses.push(`r.sheet_name = $${queryParams.length}`);
      }
      if (filterUnit && filterUnit.trim() !== '') {
        queryParams.push(`%${filterUnit.trim()}%`);
        filterClauses.push(`(r.row_data->>'KODE UNIT' ILIKE $${queryParams.length} OR r.row_data->>'KODE_UNIT' ILIKE $${queryParams.length})`);
      }

      searchQuery = `
        SELECT 
          r.id, 
          r.file_id, 
          r.sheet_name, 
          r.row_number, 
          r.row_data, 
          f.filename, 
          f.uploaded_at,
          (b.id IS NOT NULL) AS is_bookmarked
        FROM document_rows r
        JOIN uploaded_files f ON r.file_id = f.id
        LEFT JOIN bookmarks b ON b.row_id = r.id
        WHERE ${filterClauses.join(' AND ')}
        ORDER BY r.sheet_name, r.row_number
        LIMIT 250;
      `;
    } else {
      // Try to parse query as a standard box/pelaksana code (e.g. P011.A001.25 or P011.A.000001.2025)
      var codeMatch = queryText.trim().match(/^[pP](\d+)[\s\.]*([a-zA-Z])[\s\.]*(\d+)(?:[\s\.]*(\d{2,4}))?$/);
      
      if (codeMatch) {
        const unitNum = codeMatch[1];
        const boxLetter = codeMatch[2].toUpperCase();
        const folderNum = parseInt(codeMatch[3], 10);
        const yearRaw = codeMatch[4];
        
        const folder6 = String(folderNum).padStart(6, '0');
        const folder5 = String(folderNum).padStart(5, '0');
        const folder3 = String(folderNum).padStart(3, '0');
        
        let years = [];
        if (yearRaw) {
          const year4 = yearRaw.length === 2 ? '20' + yearRaw : yearRaw;
          const year2 = yearRaw.length === 4 ? yearRaw.slice(2) : yearRaw;
          years.push({ year4, year2 });
        } else {
          // Fallback check common years
          const commonYears = ['2025', '2024', '2026', '2023', '2027', '2028', '2022', '2029', '2020', '2021'];
          commonYears.forEach(y => {
            years.push({ year4: y, year2: y.slice(2) });
          });
        }
        
        const boxCodeVariations = new Set();
        years.forEach(({ year4, year2 }) => {
          // Format dots & padding
          boxCodeVariations.add(`P${unitNum}.${boxLetter}.${folder6}.${year4}`);
          boxCodeVariations.add(`P${unitNum}.${boxLetter}.${folder5}.${year4}`);
          boxCodeVariations.add(`P${unitNum}.${boxLetter}.${folder3}.${year4}`);
          
          boxCodeVariations.add(`P${unitNum}.${boxLetter}.${folder6}.${year2}`);
          boxCodeVariations.add(`P${unitNum}.${boxLetter}.${folder5}.${year2}`);
          boxCodeVariations.add(`P${unitNum}.${boxLetter}.${folder3}.${year2}`);
          
          // Format no-dots or 3-digit padding (like P011.A001.25)
          boxCodeVariations.add(`P${unitNum}.${boxLetter}${folder3}.${year2}`);
          boxCodeVariations.add(`P${unitNum}.${boxLetter}${folder3}.${year4}`);
          boxCodeVariations.add(`P${unitNum}.${boxLetter}${folder5}.${year2}`);
          boxCodeVariations.add(`P${unitNum}.${boxLetter}${folder6}.${year2}`);
          boxCodeVariations.add(`P${unitNum}.${boxLetter}${folder5}.${year4}`);
          boxCodeVariations.add(`P${unitNum}.${boxLetter}${folder6}.${year4}`);
        });
        
        const keys = ['NO.BOKS', 'NO. BOKS', 'NO BOKS', 'NO_BOKS', 'KODE BOKS', 'KODE_BOKS', 'BOKS', 'BOX'];
        const containmentClauses = [];
        queryParams = [];
        
        boxCodeVariations.forEach(boxCode => {
          keys.forEach(key => {
            queryParams.push(JSON.stringify({ [key]: boxCode }));
            containmentClauses.push(`r.row_data @> $${queryParams.length}`);
          });
        });
        
        // Add filters if any
        let filterClauses = [];
        if (filterSheet && filterSheet.trim() !== '') {
          queryParams.push(filterSheet.trim());
          filterClauses.push(`r.sheet_name = $${queryParams.length}`);
        }
        if (filterUnit && filterUnit.trim() !== '') {
          queryParams.push(`%${filterUnit.trim()}%`);
          filterClauses.push(`(r.row_data->>'KODE UNIT' ILIKE $${queryParams.length} OR r.row_data->>'KODE_UNIT' ILIKE $${queryParams.length})`);
        }

        const filterSql = filterClauses.length > 0 ? ' AND ' + filterClauses.join(' AND ') : '';

        searchQuery = `
          SELECT 
            r.id,
            r.file_id, 
            r.sheet_name, 
            r.row_number, 
            r.row_data, 
            f.filename, 
            f.uploaded_at,
            (b.id IS NOT NULL) AS is_bookmarked,
            100 as relevance_score
          FROM document_rows r
          JOIN uploaded_files f ON r.file_id = f.id
          LEFT JOIN bookmarks b ON b.row_id = r.id
          WHERE (${containmentClauses.join(' OR ')}) ${filterSql}
          ORDER BY relevance_score DESC, f.uploaded_at DESC, r.file_id, r.sheet_name, r.row_number
          LIMIT 150;
        `;
      } else {
        // Search across all rows (Google-style: match all tokens split by spaces or dots in any order)
        const searchTerms = queryText.trim().split(/[\s\.]+/).filter(Boolean);
        
        if (searchTerms.length === 0) {
          return res.json({ results: [] });
        }

        const whereClauses = [];
        const scoreExpressions = [];
        queryParams = [];
        
        searchTerms.forEach((term) => {
          const match = term.match(/^([a-zA-Z]+)(0*[1-9]\d*)$/);
          if (match) {
            const letters = match[1];
            const digits = match[2];
            
            queryParams.push(`%${term}%`, `%${letters}%`, `%${digits}%`);
            const idxTerm = queryParams.length - 2;
            const idxLetters = queryParams.length - 1;
            const idxDigits = queryParams.length;
            
            whereClauses.push(`(r.row_data::text ILIKE $${idxTerm} OR (r.row_data::text ILIKE $${idxLetters} AND r.row_data::text ILIKE $${idxDigits}))`);
            
            // Give extra weight if the mixed token matches contiguously
            scoreExpressions.push(`(CASE WHEN r.row_data::text ILIKE $${idxTerm} THEN 20 ELSE 0 END)`);
          } else {
            queryParams.push(`%${term}%`);
            const idxTerm = queryParams.length;
            whereClauses.push(`r.row_data::text ILIKE $${idxTerm}`);
            scoreExpressions.push(`(CASE WHEN r.row_data::text ILIKE $${idxTerm} THEN 10 ELSE 0 END)`);
          }
        });

        // Add high boost if the entire query matches contiguously
        queryParams.push(`%${queryText.trim()}%`);
        const idxFullQuery = queryParams.length;
        scoreExpressions.push(`(CASE WHEN r.row_data::text ILIKE $${idxFullQuery} THEN 100 ELSE 0 END)`);

        // Add filters if any
        if (filterSheet && filterSheet.trim() !== '') {
          queryParams.push(filterSheet.trim());
          whereClauses.push(`r.sheet_name = $${queryParams.length}`);
        }
        if (filterUnit && filterUnit.trim() !== '') {
          queryParams.push(`%${filterUnit.trim()}%`);
          whereClauses.push(`(r.row_data->>'KODE UNIT' ILIKE $${queryParams.length} OR r.row_data->>'KODE_UNIT' ILIKE $${queryParams.length})`);
        }

        searchQuery = `
          SELECT 
            r.id, 
            r.file_id, 
            r.sheet_name, 
            r.row_number, 
            r.row_data, 
            f.filename, 
            f.uploaded_at,
            (b.id IS NOT NULL) AS is_bookmarked,
            (${scoreExpressions.join(' + ')}) as relevance_score
          FROM document_rows r
          JOIN uploaded_files f ON r.file_id = f.id
          LEFT JOIN bookmarks b ON b.row_id = r.id
          WHERE ${whereClauses.join(' AND ')}
          ORDER BY relevance_score DESC, f.uploaded_at DESC, r.file_id, r.sheet_name, r.row_number
          LIMIT 150;
        `;
      }
    }

    let searchRes = await pool.query(searchQuery, queryParams);
    
    // FALLBACK: If fast JSONB containment search returned 0 rows for standard box query
    if (searchRes.rowCount === 0 && codeMatch) {
      console.log(`[Search] Fast JSONB containment returned 0 rows for "${queryText}". Executing fallback ILIKE query...`);
      const unitNum = codeMatch[1];
      const boxLetter = codeMatch[2];
      const folderNum = parseInt(codeMatch[3], 10);
      const yearRaw = codeMatch[4];
      const folder6 = String(folderNum).padStart(6, '0');
      const folder5 = String(folderNum).padStart(5, '0');
      
      let yearPart = '';
      if (yearRaw) {
        const year = yearRaw.length === 2 ? '20' + yearRaw : yearRaw;
        yearPart = '.' + year;
      } else {
        yearPart = '%';
      }
      
      const pattern1 = `%P${unitNum}.${boxLetter}.${folder6}${yearPart}%`;
      const pattern2_part1 = `%P${unitNum}%`;
      const pattern2_part2 = `%${boxLetter}${folder5}${yearPart}%`;
      
      const fallbackParams = [pattern1, pattern2_part1, pattern2_part2];
      
      let fallbackFilterClauses = [];
      if (filterSheet && filterSheet.trim() !== '') {
        fallbackParams.push(filterSheet.trim());
        fallbackFilterClauses.push(`r.sheet_name = $${fallbackParams.length}`);
      }
      if (filterUnit && filterUnit.trim() !== '') {
        fallbackParams.push(`%${filterUnit.trim()}%`);
        fallbackFilterClauses.push(`(r.row_data->>'KODE UNIT' ILIKE $${fallbackParams.length} OR r.row_data->>'KODE_UNIT' ILIKE $${fallbackParams.length})`);
      }

      const fallbackFilterSql = fallbackFilterClauses.length > 0 ? ' AND ' + fallbackFilterClauses.join(' AND ') : '';

      const fallbackQuery = `
        SELECT 
          r.id, 
          r.file_id, 
          r.sheet_name, 
          r.row_number, 
          r.row_data, 
          f.filename, 
          f.uploaded_at,
          (b.id IS NOT NULL) AS is_bookmarked,
          (CASE 
            WHEN r.row_data::text ILIKE $1 THEN 100 
            WHEN r.row_data::text ILIKE $2 AND r.row_data::text ILIKE $3 THEN 50 
            ELSE 0 
          END) as relevance_score
        FROM document_rows r
        JOIN uploaded_files f ON r.file_id = f.id
        LEFT JOIN bookmarks b ON b.row_id = r.id
        WHERE (r.row_data::text ILIKE $1 OR (r.row_data::text ILIKE $2 AND r.row_data::text ILIKE $3)) ${fallbackFilterSql}
        ORDER BY relevance_score DESC, f.uploaded_at DESC, r.file_id, r.sheet_name, r.row_number
        LIMIT 150;
      `;
      searchRes = await pool.query(fallbackQuery, fallbackParams);
    }
    
    // Log search activity if it's a real query (not file preview)
    if (!fileId && queryText && queryText.trim() !== '') {
      logActivity('search', `Mencari kata kunci: "${queryText.trim()}" (Menghasilkan ${searchRes.rows.length} baris data)`, getDeviceInfo(req));
    }
    
    const groupedResults = {};
    searchRes.rows.forEach(row => {
      if (!groupedResults[row.file_id]) {
        groupedResults[row.file_id] = {
          fileId: row.file_id,
          filename: row.filename,
          uploadedAt: row.uploaded_at,
          matches: []
        };
      }
      groupedResults[row.file_id].matches.push({
        id: row.id,
        sheetName: row.sheet_name,
        rowNumber: row.row_number,
        rowData: row.row_data,
        isBookmarked: !!row.is_bookmarked
      });
    });

    res.json({ results: Object.values(groupedResults) });
  } catch (err) {
    console.error('Error saat melakukan pencarian/pratinjau:', err);
    res.status(500).json({ error: 'Gagal mengambil data dari database: ' + err.message });
  }
});

// 5. Get list of all uploaded files
app.get('/api/files', async (req, res) => {
  try {
    const query = `
      SELECT 
        id, 
        filename, 
        sheet_names, 
        uploaded_at, 
        row_count
      FROM uploaded_files
      ORDER BY uploaded_at DESC;
    `;
    const filesRes = await pool.query(query);
    res.json({ files: filesRes.rows });
  } catch (err) {
    console.error('Error saat mengambil daftar file:', err);
    res.status(500).json({ error: 'Gagal mengambil daftar file: ' + err.message });
  }
});

// 5b. Get activity logs (with search + pagination)
app.get('/api/logs', async (req, res) => {
  try {
    const page  = Math.max(1, parseInt(req.query.page)  || 1);
    const limit = Math.min(500, Math.max(10, parseInt(req.query.limit) || 50));
    const q     = (req.query.q || '').trim();
    const offset = (page - 1) * limit;

    const conditions = q
      ? `WHERE activity_details ILIKE $1 OR activity_type ILIKE $1 OR ip_address ILIKE $1 OR browser ILIKE $1 OR os ILIKE $1 OR device_label ILIKE $1`
      : '';
    const params = q ? [`%${q}%`] : [];

    const countRes = await pool.query(
      `SELECT COUNT(*) FROM activity_logs ${conditions}`, params
    );
    const total = parseInt(countRes.rows[0].count);

    const dataRes = await pool.query(
      `SELECT id, activity_type, activity_details,
              ip_address, browser, os, device_type, device_label, engine, user_agent,
              created_at
       FROM activity_logs
       ${conditions}
       ORDER BY created_at DESC
       LIMIT $${params.length + 1} OFFSET $${params.length + 2}`,
      [...params, limit, offset]
    );

    res.json({
      logs: dataRes.rows,
      total,
      page,
      limit,
      totalPages: Math.ceil(total / limit)
    });
  } catch (err) {
    console.error('Error saat mengambil log aktivitas:', err);
    res.status(500).json({ error: 'Gagal mengambil log: ' + err.message });
  }
});

// 6. Delete an uploaded file
app.delete('/api/files/:id', async (req, res) => {
  const fileId = req.params.id;
  try {
    const deleteQuery = `DELETE FROM uploaded_files WHERE id = $1 RETURNING filename;`;
    const deleteRes = await pool.query(deleteQuery, [fileId]);
    
    if (deleteRes.rowCount === 0) {
      return res.status(404).json({ error: 'File tidak ditemukan!' });
    }

    const filename = deleteRes.rows[0].filename;
    await logActivity('delete', `Menghapus berkas "${filename}" beserta seluruh data terkait.`, getDeviceInfo(req));

    res.json({ 
      success: true, 
      message: `File "${filename}" beserta seluruh data terkait berhasil dihapus.` 
    });
  } catch (err) {
    console.error('Error saat menghapus file:', err);
    res.status(500).json({ error: 'Gagal menghapus file: ' + err.message });
  }
});

// 6b. Bulk delete uploaded files
app.post('/api/files/bulk-delete', async (req, res) => {
  const { ids } = req.body;
  if (!ids || !Array.isArray(ids) || ids.length === 0) {
    return res.status(400).json({ error: 'Tidak ada file ID yang diberikan!' });
  }

  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    
    // Get filenames for logging
    const fileRes = await client.query(
      `SELECT filename FROM uploaded_files WHERE id = ANY($1)`,
      [ids]
    );
    const filenames = fileRes.rows.map(r => r.filename);

    const deleteRes = await client.query(
      `DELETE FROM uploaded_files WHERE id = ANY($1)`,
      [ids]
    );

    if (deleteRes.rowCount === 0) {
      await client.query('ROLLBACK');
      return res.status(404).json({ error: 'Tidak ada file yang ditemukan untuk dihapus!' });
    }

    const device = getDeviceInfo(req);
    await logActivity('delete', `Menghapus massal ${deleteRes.rowCount} berkas: ${filenames.join(', ')}`, device);

    await client.query('COMMIT');
    res.json({ 
      success: true, 
      message: `${deleteRes.rowCount} berkas berhasil dihapus.` 
    });
  } catch (err) {
    await client.query('ROLLBACK');
    console.error('Error bulk delete:', err);
    res.status(500).json({ error: 'Gagal menghapus berkas massal: ' + err.message });
  } finally {
    client.release();
  }
});

// 7. Get bookmarks
app.get('/api/bookmarks', async (req, res) => {
  try {
    const query = `
      SELECT 
        b.id as bookmark_id,
        b.created_at as bookmarked_at,
        r.id,
        r.file_id,
        r.sheet_name,
        r.row_number,
        r.row_data,
        f.filename,
        true AS is_bookmarked
      FROM bookmarks b
      JOIN document_rows r ON b.row_id = r.id
      JOIN uploaded_files f ON r.file_id = f.id
      ORDER BY b.created_at DESC;
    `;
    const bookmarkedRes = await pool.query(query);
    res.json({ bookmarks: bookmarkedRes.rows });
  } catch (err) {
    console.error('Error bookmarks:', err);
    res.status(500).json({ error: err.message });
  }
});

// 8. Add bookmark
app.post('/api/bookmarks', async (req, res) => {
  const { rowId } = req.body;
  if (!rowId) {
    return res.status(400).json({ error: 'rowId diperlukan!' });
  }
  try {
    await pool.query(
      `INSERT INTO bookmarks (row_id) VALUES ($1) ON CONFLICT (row_id) DO NOTHING;`,
      [rowId]
    );
    res.json({ success: true, message: 'Data berhasil disimpan ke bookmark.' });
  } catch (err) {
    console.error('Error add bookmark:', err);
    res.status(500).json({ error: err.message });
  }
});

// 9. Remove bookmark
app.delete('/api/bookmarks/:rowId', async (req, res) => {
  const rowId = req.params.rowId;
  try {
    await pool.query(`DELETE FROM bookmarks WHERE row_id = $1;`, [rowId]);
    res.json({ success: true, message: 'Bookmark berhasil dihapus.' });
  } catch (err) {
    console.error('Error delete bookmark:', err);
    res.status(500).json({ error: err.message });
  }
});


// Global error handler for upload limit
app.use((err, req, res, next) => {
  if (err instanceof multer.MulterError) {
    if (err.code === 'LIMIT_FILE_SIZE') {
      return res.status(400).json({ error: 'Ukuran file terlalu besar! Maksimal 250MB.' });
    }
  }
  if (err) {
    return res.status(400).json({ error: err.message });
  }
  next();
});

// Start Express Server
app.listen(port, () => {
  console.log(`Server backend berjalan di http://localhost:${port}`);
});
