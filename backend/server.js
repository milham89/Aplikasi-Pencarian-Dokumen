const express = require('express');
const cors = require('cors');
const multer = require('multer');
const ExcelJS = require('exceljs');
const xlsx = require('xlsx');
const { Pool } = require('pg');
const { Worker } = require('worker_threads');
const fs = require('fs');
const path = require('path');
const { UAParser } = require('ua-parser-js');
require('dotenv').config({ path: path.join(__dirname, '.env') });
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');

const JWT_SECRET = process.env.JWT_SECRET || 'supersecretkey12345';

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
app.use(express.json({ limit: '250mb' }));
app.use(express.urlencoded({ limit: '250mb', extended: true }));
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

// Configure Multer for chunked uploads (bypasses Cloudflare 100MB POST limit)
const chunkUploadDir = path.join(uploadDir, 'temp_chunks');
if (!fs.existsSync(chunkUploadDir)) {
  fs.mkdirSync(chunkUploadDir, { recursive: true });
}

const chunkStorage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadId = req.body.uploadId || 'default';
    const targetDir = path.join(chunkUploadDir, uploadId);
    if (!fs.existsSync(targetDir)) {
      fs.mkdirSync(targetDir, { recursive: true });
    }
    cb(null, targetDir);
  },
  filename: (req, file, cb) => {
    const chunkIndex = req.body.chunkIndex !== undefined ? req.body.chunkIndex : '0';
    cb(null, `chunk_${chunkIndex}`);
  }
});

const uploadChunkMulter = multer({
  storage: chunkStorage,
  limits: { fileSize: 50 * 1024 * 1024 }
});

// PostgreSQL Database Connection Pool
const pool = new Pool({
  user: process.env.DB_USER || 'postgres',
  host: process.env.DB_HOST || 'localhost',
  database: process.env.DB_NAME || 'excel_db',
  password: process.env.DB_PASSWORD || 'mysecretpassword',
  port: process.env.DB_PORT || 5432,
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
    const deviceCols = ['ip_address VARCHAR(100)', 'browser VARCHAR(150)', 'os VARCHAR(150)', 'device_type VARCHAR(50)', 'device_label VARCHAR(200)', 'engine VARCHAR(100)', 'user_agent TEXT', 'user_id INTEGER', 'username VARCHAR(50)'];
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

    // Enable pg_trgm extension and create GIN & B-Tree performance indexes
    try {
      await client.query(`CREATE EXTENSION IF NOT EXISTS pg_trgm;`);
      await client.query(`
        CREATE INDEX IF NOT EXISTS idx_document_rows_trgm 
        ON document_rows USING gin ((row_data::text) gin_trgm_ops);
        CREATE INDEX IF NOT EXISTS idx_document_rows_file_id ON document_rows (file_id);
        CREATE INDEX IF NOT EXISTS idx_document_rows_file_row ON document_rows (file_id, row_number);
      `);
    } catch (trgmErr) {
      console.warn('Gagal mengaktifkan pg_trgm saat startup:', trgmErr.message);
    }

    // Create users table first
    await client.query(`
      CREATE TABLE IF NOT EXISTS users (
        id SERIAL PRIMARY KEY,
        username VARCHAR(50) UNIQUE NOT NULL,
        password_hash VARCHAR(255) NOT NULL,
        role VARCHAR(20) NOT NULL,
        full_name VARCHAR(100),
        created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP
      );
    `);

    // Migrate: add full_name, uim_code, unit_kerja columns if not exists
    await client.query(`
      ALTER TABLE users ADD COLUMN IF NOT EXISTS full_name VARCHAR(100);
      ALTER TABLE users ADD COLUMN IF NOT EXISTS uim_code VARCHAR(50);
      ALTER TABLE users ADD COLUMN IF NOT EXISTS unit_kerja VARCHAR(255);
      UPDATE users SET username = uim_code WHERE uim_code IS NOT NULL AND uim_code != '' AND LOWER(username) = LOWER(uim_code);
    `);

    // Check if bookmarks table has user_id column
    try {
      const hasUserIdRes = await client.query(`
        SELECT column_name 
        FROM information_schema.columns 
        WHERE table_name='bookmarks' AND column_name='user_id';
      `);
      
      if (hasUserIdRes.rows.length === 0) {
        console.log('Mengubah tabel bookmarks untuk mendukung user-specific bookmarks...');
        // Drop existing table if it exists to clean up constraints easily
        await client.query(`DROP TABLE IF EXISTS bookmarks CASCADE;`);
      }
    } catch (checkErr) {
      console.log('Checking/dropping bookmarks table error or table does not exist:', checkErr.message);
    }

    // Create bookmarks table
    await client.query(`
      CREATE TABLE IF NOT EXISTS bookmarks (
        id SERIAL PRIMARY KEY,
        user_id INTEGER REFERENCES users(id) ON DELETE CASCADE,
        row_id INTEGER NOT NULL REFERENCES document_rows(id) ON DELETE CASCADE,
        notes TEXT,
        created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
        CONSTRAINT unique_user_row_bookmark UNIQUE(user_id, row_id)
      );
    `);

    // Check if bookmarks table has notes column
    try {
      const hasNotesRes = await client.query(`
        SELECT column_name 
        FROM information_schema.columns 
        WHERE table_name='bookmarks' AND column_name='notes';
      `);
      if (hasNotesRes.rows.length === 0) {
        console.log('Menambahkan kolom notes ke tabel bookmarks...');
        await client.query(`ALTER TABLE bookmarks ADD COLUMN notes TEXT;`);
      }
    } catch (alterErr) {
      console.log('Error adding notes column to bookmarks:', alterErr.message);
    }

    // Check if bookmarks table has group_name column
    try {
      const hasGroupRes = await client.query(`
        SELECT column_name 
        FROM information_schema.columns 
        WHERE table_name='bookmarks' AND column_name='group_name';
      `);
      if (hasGroupRes.rows.length === 0) {
        console.log('Menambahkan kolom group_name ke tabel bookmarks...');
        await client.query(`ALTER TABLE bookmarks ADD COLUMN group_name VARCHAR(100) DEFAULT 'Umum';`);
      }
    } catch (alterErr) {
      console.log('Error adding group_name column to bookmarks:', alterErr.message);
    }

    // Ensure UNIQUE constraint on user_id and row_id exists
    try {
      // Delete any duplicates first just in case
      await client.query(`
        DELETE FROM bookmarks a USING bookmarks b 
        WHERE a.id < b.id AND a.user_id = b.user_id AND a.row_id = b.row_id;
      `);
      // Add constraint
      await client.query(`
        ALTER TABLE bookmarks 
        ADD CONSTRAINT unique_user_row_bookmark UNIQUE(user_id, row_id);
      `);
      console.log('Constraint unique_user_row_bookmark berhasil ditambahkan.');
    } catch (conErr) {
      // Ignore if it already exists
    }

    // Check if uploaded_files table has uploaded_by column
    try {
      const hasUploadedByRes = await client.query(`
        SELECT column_name 
        FROM information_schema.columns 
        WHERE table_name='uploaded_files' AND column_name='uploaded_by';
      `);
      if (hasUploadedByRes.rows.length === 0) {
        console.log('Menambahkan kolom uploaded_by ke tabel uploaded_files...');
        await client.query(`ALTER TABLE uploaded_files ADD COLUMN uploaded_by INTEGER REFERENCES users(id) ON DELETE SET NULL;`);
      }
    } catch (alterErr) {
      console.log('Error adding uploaded_by column to uploaded_files:', alterErr.message);
    }

    // Create notifications table
    await client.query(`
      CREATE TABLE IF NOT EXISTS notifications (
        id SERIAL PRIMARY KEY,
        message TEXT NOT NULL,
        recipient_id INTEGER REFERENCES users(id) ON DELETE CASCADE,
        created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP
      );
    `);

    // Check if notifications table has recipient_id column
    try {
      const hasRecipientRes = await client.query(`
        SELECT column_name 
        FROM information_schema.columns 
        WHERE table_name='notifications' AND column_name='recipient_id';
      `);
      if (hasRecipientRes.rows.length === 0) {
        console.log('Menambahkan kolom recipient_id ke tabel notifications...');
        await client.query(`ALTER TABLE notifications ADD COLUMN recipient_id INTEGER REFERENCES users(id) ON DELETE CASCADE;`);
      }
    } catch (alterErr) {
      console.log('Error adding recipient_id column to notifications:', alterErr.message);
    }

    // Create user_notification_reads table
    await client.query(`
      CREATE TABLE IF NOT EXISTS user_notification_reads (
        user_id INTEGER REFERENCES users(id) ON DELETE CASCADE,
        notification_id INTEGER REFERENCES notifications(id) ON DELETE CASCADE,
        read_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (user_id, notification_id)
      );
    `);

    // Check if users table is empty and insert default users if so
    const userCountRes = await client.query('SELECT COUNT(*)::int FROM users;');
    if (userCountRes.rows[0].count === 0) {
      console.log('Menyisipkan pengguna default (admin, operator, viewer)...');
      
      const defaultUsers = [
        { username: 'admin', password: 'admin123', role: 'admin' },
        { username: 'operator', password: 'operator123', role: 'operator' },
        { username: 'viewer', password: 'viewer123', role: 'viewer' }
      ];

      for (const u of defaultUsers) {
        const passwordHash = await bcrypt.hash(u.password, 10);
        await client.query(
          'INSERT INTO users (username, password_hash, role) VALUES ($1, $2, $3);',
          [u.username, passwordHash, u.role]
        );
      }
      console.log('Pengguna default sukses ditambahkan.');
    }

    // Create uim_records table
    await client.query(`
      CREATE TABLE IF NOT EXISTS uim_records (
        id SERIAL PRIMARY KEY,
        uim_code VARCHAR(50) NOT NULL,
        full_name VARCHAR(255) NOT NULL,
        unit_kerja VARCHAR(255) NOT NULL,
        created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP
      );
    `);

    // Always update/reseed uim_records to ensure exact match with latest dataset (L901 to M110)
    const uimCheckRes = await client.query("SELECT COUNT(*)::int FROM uim_records WHERE uim_code = 'L901';");
    if (uimCheckRes.rows[0].count === 0) {
      console.log('Memperbarui data UIM ke format terbaru (L901 - M110)...');
      try {
        await client.query('TRUNCATE TABLE uim_records RESTART IDENTITY;');
        const seedPath = path.join(__dirname, 'uim_seed_data.json');
        if (fs.existsSync(seedPath)) {
          const rawSeed = fs.readFileSync(seedPath, 'utf8');
          const seedData = JSON.parse(rawSeed);
          for (const item of seedData) {
            await client.query(
              'INSERT INTO uim_records (uim_code, full_name, unit_kerja) VALUES ($1, $2, $3);',
              [item.uim_code || '', item.full_name || '', item.unit_kerja || '']
            );
          }
          console.log(`Berhasil menyisipkan ${seedData.length} data UIM terbaru.`);
        }
      } catch (seedErr) {
        console.error('Gagal menyisipkan data UIM terbaru:', seedErr.message);
      }
    }

    console.log('Inisialisasi database PostgreSQL sukses.');
  } catch (err) {
    console.error('Gagal melakukan inisialisasi database:', err.message);
  } finally {
    client.release();
  }
};

// Call DB Init
initDatabase();

// Cosine Similarity helper for Vector embeddings
const cosineSimilarity = (vecA, vecB) => {
  let dotProduct = 0.0;
  let normA = 0.0;
  let normB = 0.0;
  for (let i = 0; i < vecA.length; i++) {
    dotProduct += vecA[i] * vecB[i];
    normA += vecA[i] * vecA[i];
    normB += vecB[i] * vecB[i];
  }
  if (normA === 0 || normB === 0) return 0.0;
  return dotProduct / (Math.sqrt(normA) * Math.sqrt(normB));
};

// Fetch embeddings in batches from Google Gemini API
const getGeminiEmbeddings = async (texts) => {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) throw new Error('GEMINI_API_KEY tidak dikonfigurasi di environment variable.');

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-embedding-2:batchEmbedContents?key=${apiKey}`;
  const requests = texts.map(text => ({
    model: 'models/gemini-embedding-2',
    content: { parts: [{ text }] }
  }));

  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ requests })
  });

  if (!res.ok) {
    const err = await res.json();
    throw new Error(err.error?.message || 'Gemini API Error');
  }

  const data = await res.json();
  return data.embeddings.map(e => e.values);
};

// Chunk helper to fetch embeddings in batches safe from rate-limits/payload size
const getGeminiEmbeddingsInChunks = async (texts) => {
  const chunkSize = 80; // safe chunk size
  const embeddings = [];
  
  for (let i = 0; i < texts.length; i += chunkSize) {
    const chunk = texts.slice(i, i + chunkSize);
    const chunkEmbeddings = await getGeminiEmbeddings(chunk);
    embeddings.push(...chunkEmbeddings);
  }
  return embeddings;
};

// Compute Semantic Similarity Scores using Gemini API Embeddings
const computeGeminiSemanticScores = async (query, rows) => {
  const rowTexts = rows.map(r => {
    // Format JSON row_data as a clean string representation for AI embeddings
    return Object.entries(r.row_data)
      .map(([k, v]) => `${k}: ${v}`)
      .join(' | ');
  });

  const allTexts = [query, ...rowTexts];
  const allEmbeddings = await getGeminiEmbeddingsInChunks(allTexts);

  const queryVector = allEmbeddings[0];
  const docVectors = allEmbeddings.slice(1);

  return rows.map((row, index) => {
    const similarity = cosineSimilarity(queryVector, docVectors[index]);
    // Normalise similarity (usually 0 to 1 for semantic match) to 0-100 integer range
    const score = Math.max(0, Math.min(100, Math.round(similarity * 100)));
    return {
      ...row,
      relevance_score: score
    };
  });
};

// Local TF-IDF Vector Space Model (Zero dependencies, 100% free fallback semantic match)
const computeLocalSemanticScores = (query, rows) => {
  const tokenize = (text) => {
    return text.toLowerCase()
      .replace(/[^\w\s\.]/g, ' ')
      .split(/[\s\.]+/)
      .filter(t => t.length > 1);
  };

  const queryTokens = tokenize(query);
  const docTokensList = docTokensListGetter(rows);

  const getTF = (tokens) => {
    const tf = {};
    tokens.forEach(t => {
      tf[t] = (tf[t] || 0) + 1;
    });
    const len = tokens.length || 1;
    Object.keys(tf).forEach(t => {
      tf[t] = tf[t] / len;
    });
    return tf;
  };

  const queryTF = getTF(queryTokens);
  const docTFs = docTokensList.map(tokens => getTF(tokens));

  const allTokens = new Set([...queryTokens, ...docTokensList.flat()]);
  const idf = {};
  const N = rows.length;

  allTokens.forEach(token => {
    let docCount = 0;
    docTokensList.forEach(tokens => {
      if (tokens.includes(token)) docCount++;
    });
    idf[token] = Math.log((N + 1) / (docCount + 1)) + 1;
  });

  const getVector = (tf) => {
    const vec = {};
    allTokens.forEach(token => {
      vec[token] = (tf[token] || 0) * (idf[token] || 0);
    });
    return vec;
  };

  const queryVec = getVector(queryTF);
  const docVecs = docTFs.map(tf => getVector(tf));

  const getCosine = (vecA, vecB) => {
    let dot = 0;
    let normA = 0;
    let normB = 0;
    allTokens.forEach(token => {
      dot += vecA[token] * vecB[token];
      normA += vecA[token] * vecA[token];
      normB += vecB[token] * vecB[token];
    });
    if (normA === 0 || normB === 0) return 0;
    return dot / (Math.sqrt(normA) * Math.sqrt(normB));
  };

  return rows.map((row, index) => {
    const similarity = getCosine(queryVec, docVecs[index]);
    
    // Contiguous substring match boost (e.g. n-gram and substring checks)
    let boost = 0;
    const rowText = Object.entries(row.row_data)
      .map(([k, v]) => `${k} ${v}`)
      .join(' ').toLowerCase();
    
    if (rowText.includes(query.toLowerCase())) {
      boost = 0.35; // substantial boost for exact contiguous query match
    } else {
      let matchCount = 0;
      queryTokens.forEach(t => {
        if (rowText.includes(t)) matchCount++;
      });
      boost = (matchCount / (queryTokens.length || 1)) * 0.15;
    }

    const finalScore = Math.max(0, Math.min(100, Math.round((similarity + boost) * 100)));
    return {
      ...row,
      relevance_score: finalScore
    };
  });
};

// Helper getter to keep docTokensList clean
const docTokensListGetter = (rows) => {
  const tokenize = (text) => {
    return text.toLowerCase()
      .replace(/[^\w\s\.]/g, ' ')
      .split(/[\s\.]+/)
      .filter(t => t.length > 1);
  };
  return rows.map(r => {
    const rowText = Object.entries(r.row_data)
      .map(([k, v]) => `${k} ${v}`)
      .join(' ');
    return tokenize(rowText);
  });
};

// Function to log activities to database (with optional user + device info)
const logActivity = async (type, details, deviceInfo = null, userInfo = null) => {
  try {
    if (deviceInfo) {
      await pool.query(
        `INSERT INTO activity_logs 
          (activity_type, activity_details, ip_address, browser, os, device_type, device_label, engine, user_agent, user_id, username)
         VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11);`,
        [type, details,
          deviceInfo.ip, deviceInfo.browser, deviceInfo.os,
          deviceInfo.device_type, deviceInfo.device_label, deviceInfo.engine, deviceInfo.user_agent,
          userInfo ? userInfo.id : null, userInfo ? userInfo.username : null
        ]
      );
    } else {
      await pool.query(
        `INSERT INTO activity_logs (activity_type, activity_details, user_id, username) VALUES ($1, $2, $3, $4);`,
        [type, details, userInfo ? userInfo.id : null, userInfo ? userInfo.username : null]
      );
    }
  } catch (err) {
    console.error('Gagal menulis log aktivitas:', err);
  }
};

// --- BACKGROUND EXCEL PROCESSOR ---

async function processFileInBackground(jobId, filePath, originalname, userInfo = null) {
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

    // 2. Insert record in uploaded_files with uploaded_by (owner)
    const fileInsertRes = await client.query(
      `INSERT INTO uploaded_files (filename, sheet_names, uploaded_by) VALUES ($1, $2, $3) RETURNING id;`,
      [originalname, [], userInfo ? userInfo.id : null]
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
      const BATCH_SIZE = 1000;

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

          // Update job status with dynamic progress calculation
          uploadJobs[jobId].processedRows = totalRowsInserted;
          const calcProgress = Math.min(95, 20 + Math.floor((totalRowsInserted / (totalRowsInserted + 8000)) * 75));
          uploadJobs[jobId].progress = calcProgress;
          uploadJobs[jobId].message = `Mengimpor sheet "${sheetName}": ${totalRowsInserted.toLocaleString('id-ID')} baris (${calcProgress}%)...`;
          
          // Yield to Event Loop
          await new Promise(resolve => setImmediate(resolve));
        }
      }

      // Insert remaining rows in sheet
      if (rowsToInsert.length > 0) {
        await insertBatch(client, fileId, sheetName, rowsToInsert);
        totalRowsInserted += rowsToInsert.length;
        uploadJobs[jobId].processedRows = totalRowsInserted;
        const calcProgress = Math.min(95, 20 + Math.floor((totalRowsInserted / (totalRowsInserted + 8000)) * 75));
        uploadJobs[jobId].progress = calcProgress;
        uploadJobs[jobId].message = `Mengimpor sheet "${sheetName}": ${totalRowsInserted.toLocaleString('id-ID')} baris (${calcProgress}%)...`;
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
      `INSERT INTO activity_logs (activity_type, activity_details, user_id, username) VALUES ($1, $2, $3, $4);`,
      ['upload', `Berhasil mengunggah berkas "${originalname}" (${totalRowsInserted.toLocaleString('id-ID')} baris data)`,
        userInfo ? userInfo.id : null, userInfo ? userInfo.username : null]
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

// Authentication Middleware (JWT Validation)
const authenticateToken = (req, res, next) => {
  const authHeader = req.headers['authorization'];
  const token = authHeader && authHeader.split(' ')[1];

  if (!token) {
    return res.status(401).json({ error: 'Token tidak ditemukan. Silakan login kembali.' });
  }

  jwt.verify(token, JWT_SECRET, (err, user) => {
    if (err) {
      return res.status(403).json({ error: 'Token tidak valid atau kedaluwarsa. Silakan login kembali.' });
    }
    req.user = user;
    next();
  });
};

// Authorization Middleware (Role Checking)
const requireRole = (allowedRoles) => {
  return (req, res, next) => {
    if (!req.user || !allowedRoles.includes(req.user.role)) {
      return res.status(403).json({ error: 'Akses ditolak. Anda tidak memiliki izin untuk melakukan aksi ini.' });
    }
    next();
  };
};

// Auth Endpoint: Login
app.post('/api/auth/login', async (req, res) => {
  const { username, password } = req.body;
  if (!username || !password) {
    return res.status(400).json({ error: 'Username dan password wajib diisi.' });
  }

  try {
    const userRes = await pool.query('SELECT * FROM users WHERE username = $1;', [username]);
    if (userRes.rowCount === 0) {
      // Log failed login attempt
      logActivity('login_failed', `Login gagal: akun "${username}" tidak ditemukan`, getDeviceInfo(req), { id: null, username });
      return res.status(401).json({ error: 'Username atau password salah.' });
    }

    const user = userRes.rows[0];
    const passwordMatch = await bcrypt.compare(password, user.password_hash);
    if (!passwordMatch) {
      // Log failed login attempt (wrong password)
      logActivity('login_failed', `Login gagal: password salah untuk akun "${username}"`, getDeviceInfo(req), { id: user.id, username });
      return res.status(401).json({ error: 'Username atau password salah.' });
    }

    const token = jwt.sign(
      { id: user.id, username: user.username, role: user.role },
      JWT_SECRET,
      { expiresIn: '24h' }
    );

    // Log successful login
    logActivity('login', `Login berhasil: akun "${user.username}" (${user.role}) masuk ke aplikasi`, getDeviceInfo(req), { id: user.id, username: user.username });

    res.json({
      token,
      user: {
        id: user.id,
        username: user.username,
        role: user.role,
        full_name: user.full_name
      }
    });
  } catch (err) {
    console.error('Error login:', err);
    res.status(500).json({ error: 'Terjadi kesalahan pada server saat login.' });
  }
});

// Auth Endpoint: Logout (records audit log)
app.post('/api/auth/logout', authenticateToken, async (req, res) => {
  try {
    logActivity('logout', `Logout: akun "${req.user.username}" keluar dari aplikasi`, getDeviceInfo(req), req.user);
    res.json({ message: 'Logout berhasil dicatat.' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Auth Endpoint: Get current user profile details (includes full_name)
app.get('/api/auth/me', authenticateToken, async (req, res) => {
  try {
    const userRes = await pool.query('SELECT id, username, full_name, role FROM users WHERE id = $1;', [req.user.id]);
    if (userRes.rowCount === 0) {
      return res.status(404).json({ error: 'Pengguna tidak ditemukan.' });
    }
    res.json(userRes.rows[0]);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// User Management API: Get all users (Admin Only)
app.get('/api/users', authenticateToken, requireRole(['admin']), async (req, res) => {
  try {
    const usersRes = await pool.query('SELECT id, username, full_name, role, uim_code, unit_kerja, created_at FROM users ORDER BY created_at DESC;');
    res.json(usersRes.rows);
  } catch (err) {
    console.error('Error fetching users:', err);
    res.status(500).json({ error: err.message });
  }
});

// User Management API: Create new user (Admin Only)
app.post('/api/users', authenticateToken, requireRole(['admin']), async (req, res) => {
  const { username, password, role, full_name, uim_code, unit_kerja } = req.body;
  if (!username || !password || !role) {
    return res.status(400).json({ error: 'Username, password, dan role wajib diisi.' });
  }

  const validRoles = ['admin', 'operator', 'viewer'];
  if (!validRoles.includes(role)) {
    return res.status(400).json({ error: 'Role tidak valid. Pilih admin, operator, atau viewer.' });
  }

  try {
    const checkUser = await pool.query('SELECT id FROM users WHERE username = $1;', [username]);
    if (checkUser.rowCount > 0) {
      return res.status(400).json({ error: 'Username sudah digunakan.' });
    }

    const passwordHash = await bcrypt.hash(password, 10);
    const result = await pool.query(
      'INSERT INTO users (username, password_hash, role, full_name, uim_code, unit_kerja) VALUES ($1, $2, $3, $4, $5, $6) RETURNING id, username, full_name, role, uim_code, unit_kerja, created_at;',
      [username.trim(), passwordHash, role, full_name || null, uim_code || null, unit_kerja || null]
    );

    const newUser = result.rows[0];
    logActivity('create_user', `Membuat akun baru: "${newUser.username}" (${newUser.role}) integrated with UIM "${newUser.uim_code || '-'}"`, getDeviceInfo(req), req.user);
    res.status(201).json(newUser);
  } catch (err) {
    console.error('Error creating user:', err);
    res.status(500).json({ error: err.message });
  }
});

// User Management API: Update user (Admin Only)
app.put('/api/users/:id', authenticateToken, requireRole(['admin']), async (req, res) => {
  const { id } = req.params;
  const { username, password, role, full_name, uim_code, unit_kerja } = req.body;

  if (!username || !role) {
    return res.status(400).json({ error: 'Username dan role wajib diisi.' });
  }

  const validRoles = ['admin', 'operator', 'viewer'];
  if (!validRoles.includes(role)) {
    return res.status(400).json({ error: 'Role tidak valid. Pilih admin, operator, atau viewer.' });
  }

  try {
    // Check if user exists
    const checkUser = await pool.query('SELECT id, username FROM users WHERE id = $1;', [id]);
    if (checkUser.rowCount === 0) {
      return res.status(404).json({ error: 'Pengguna tidak ditemukan.' });
    }

    // Check username uniqueness if changed
    if (username !== checkUser.rows[0].username) {
      const checkDup = await pool.query('SELECT id FROM users WHERE username = $1;', [username]);
      if (checkDup.rowCount > 0) {
        return res.status(400).json({ error: 'Username sudah digunakan.' });
      }
    }

    let result;
    if (password && password.trim() !== '') {
      const passwordHash = await bcrypt.hash(password, 10);
      result = await pool.query(
        `UPDATE users 
         SET username = $1, password_hash = $2, role = $3, full_name = $4, uim_code = $5, unit_kerja = $6 
         WHERE id = $7 
         RETURNING id, username, full_name, role, uim_code, unit_kerja, created_at;`,
        [username.trim(), passwordHash, role, full_name || null, uim_code || null, unit_kerja || null, id]
      );
    } else {
      result = await pool.query(
        `UPDATE users 
         SET username = $1, role = $2, full_name = $3, uim_code = $4, unit_kerja = $5 
         WHERE id = $6 
         RETURNING id, username, full_name, role, uim_code, unit_kerja, created_at;`,
        [username.trim(), role, full_name || null, uim_code || null, unit_kerja || null, id]
      );
    }

    const updatedUser = result.rows[0];
    logActivity('update_user', `Memperbarui akun: "${updatedUser.username}" (${updatedUser.role}) integrated with UIM "${updatedUser.uim_code || '-'}"`, getDeviceInfo(req), req.user);
    res.json(updatedUser);
  } catch (err) {
    console.error('Error updating user:', err);
    res.status(500).json({ error: err.message });
  }
});

// User Management API: Delete user (Admin Only)
app.delete('/api/users/:id', authenticateToken, requireRole(['admin']), async (req, res) => {
  const { id } = req.params;

  try {
    if (parseInt(id) === req.user.id) {
      return res.status(400).json({ error: 'Anda tidak dapat menghapus akun Anda sendiri.' });
    }

    const deleteRes = await pool.query('DELETE FROM users WHERE id = $1 RETURNING username;', [id]);
    if (deleteRes.rowCount === 0) {
      return res.status(404).json({ error: 'Pengguna tidak ditemukan.' });
    }

    const deletedUsername = deleteRes.rows[0].username;
    logActivity('delete_user', `Menghapus akun: "${deletedUsername}"`, getDeviceInfo(req), req.user);
    res.json({ message: `Pengguna ${deletedUsername} berhasil dihapus.` });
  } catch (err) {
    console.error('Error deleting user:', err);
    res.status(500).json({ error: err.message });
  }
});

// --- UIM (User ID Manajemen) API Endpoints ---

// Get list of UIM records with search, unit filter, & pagination
app.get('/api/uim', authenticateToken, async (req, res) => {
  try {
    const page = parseInt(req.query.page) || 1;
    const limit = parseInt(req.query.limit) || 25;
    const offset = (page - 1) * limit;
    const search = req.query.search ? req.query.search.trim() : '';
    const unit = req.query.unit ? req.query.unit.trim() : '';

    let whereConditions = [];
    let queryParams = [];
    let paramCounter = 1;

    if (search) {
      whereConditions.push(`(uim_code ILIKE $${paramCounter} OR full_name ILIKE $${paramCounter} OR unit_kerja ILIKE $${paramCounter})`);
      queryParams.push(`%${search}%`);
      paramCounter++;
    }

    if (unit && unit !== 'ALL') {
      whereConditions.push(`unit_kerja = $${paramCounter}`);
      queryParams.push(unit);
      paramCounter++;
    }

    const whereClause = whereConditions.length > 0 ? `WHERE ${whereConditions.join(' AND ')}` : '';

    // Count query
    const countRes = await pool.query(`SELECT COUNT(*)::int FROM uim_records ${whereClause};`, queryParams);
    const totalRecords = countRes.rows[0].count;

    // Select query with pagination
    const selectQuery = `
      SELECT id, uim_code, full_name, unit_kerja, created_at 
      FROM uim_records 
      ${whereClause} 
      ORDER BY id ASC 
      LIMIT $${paramCounter} OFFSET $${paramCounter + 1};
    `;
    queryParams.push(limit, offset);

    const recordsRes = await pool.query(selectQuery, queryParams);

    // Get list of all distinct units for dropdown filter
    const unitsRes = await pool.query(`SELECT DISTINCT unit_kerja FROM uim_records WHERE unit_kerja IS NOT NULL AND unit_kerja != '' ORDER BY unit_kerja ASC;`);
    const units = unitsRes.rows.map(r => r.unit_kerja);

    res.json({
      records: recordsRes.rows,
      total: totalRecords,
      page,
      limit,
      totalPages: Math.ceil(totalRecords / limit) || 1,
      units
    });
  } catch (err) {
    console.error('Error fetching UIM records:', err);
    res.status(500).json({ error: 'Gagal mengambil data UIM.' });
  }
});

// Get unique unit kerja list for filter
app.get('/api/uim/units', authenticateToken, async (req, res) => {
  try {
    const unitsRes = await pool.query(`SELECT DISTINCT unit_kerja FROM uim_records WHERE unit_kerja IS NOT NULL AND unit_kerja != '' ORDER BY unit_kerja ASC;`);
    res.json(unitsRes.rows.map(r => r.unit_kerja));
  } catch (err) {
    res.status(500).json({ error: 'Gagal mengambil daftar unit kerja.' });
  }
});

// Create new UIM record (Admin/Operator)
app.post('/api/uim', authenticateToken, requireRole(['admin', 'operator']), async (req, res) => {
  const { uim_code, full_name, unit_kerja } = req.body;

  if (!uim_code || !full_name) {
    return res.status(400).json({ error: 'USER ID / UIM dan Nama Lengkap wajib diisi.' });
  }

  try {
    const insertRes = await pool.query(
      `INSERT INTO uim_records (uim_code, full_name, unit_kerja) VALUES ($1, $2, $3) RETURNING *;`,
      [uim_code.trim().toUpperCase(), full_name.trim(), (unit_kerja || 'Umum').trim()]
    );

    const newRecord = insertRes.rows[0];
    logActivity('create_uim', `Menambahkan UIM baru: "${newRecord.uim_code}" - ${newRecord.full_name}`, getDeviceInfo(req), req.user);
    res.status(201).json(newRecord);
  } catch (err) {
    console.error('Error adding UIM:', err);
    res.status(500).json({ error: 'Gagal menambah data UIM.' });
  }
});

// Update UIM record (Admin/Operator)
app.put('/api/uim/:id', authenticateToken, requireRole(['admin', 'operator']), async (req, res) => {
  const { id } = req.params;
  const { uim_code, full_name, unit_kerja } = req.body;

  if (!uim_code || !full_name) {
    return res.status(400).json({ error: 'USER ID / UIM dan Nama Lengkap wajib diisi.' });
  }

  try {
    const updateRes = await pool.query(
      `UPDATE uim_records SET uim_code = $1, full_name = $2, unit_kerja = $3 WHERE id = $4 RETURNING *;`,
      [uim_code.trim().toUpperCase(), full_name.trim(), (unit_kerja || 'Umum').trim(), id]
    );

    if (updateRes.rowCount === 0) {
      return res.status(404).json({ error: 'Data UIM tidak ditemukan.' });
    }

    const updatedRecord = updateRes.rows[0];
    logActivity('update_uim', `Memperbarui UIM: "${updatedRecord.uim_code}" - ${updatedRecord.full_name}`, getDeviceInfo(req), req.user);
    res.json(updatedRecord);
  } catch (err) {
    console.error('Error updating UIM:', err);
    res.status(500).json({ error: 'Gagal memperbarui data UIM.' });
  }
});

// Delete UIM record (Admin Only)
app.delete('/api/uim/:id', authenticateToken, requireRole(['admin']), async (req, res) => {
  const { id } = req.params;

  try {
    const deleteRes = await pool.query(`DELETE FROM uim_records WHERE id = $1 RETURNING uim_code, full_name;`, [id]);
    if (deleteRes.rowCount === 0) {
      return res.status(404).json({ error: 'Data UIM tidak ditemukan.' });
    }

    const deletedRecord = deleteRes.rows[0];
    logActivity('delete_uim', `Menghapus UIM: "${deletedRecord.uim_code}" - ${deletedRecord.full_name}`, getDeviceInfo(req), req.user);
    res.json({ message: `Data UIM ${deletedRecord.uim_code} berhasil dihapus.` });
  } catch (err) {
    console.error('Error deleting UIM:', err);
    res.status(500).json({ error: 'Gagal menghapus data UIM.' });
  }
});

// Bulk Update Unit Kerja for selected UIM IDs (Admin/Operator)
app.post('/api/uim/bulk-update-unit', authenticateToken, requireRole(['admin', 'operator']), async (req, res) => {
  const { ids, unit_kerja } = req.body;

  if (!Array.isArray(ids) || ids.length === 0) {
    return res.status(400).json({ error: 'Pilih setidaknya satu data UIM untuk diperbarui.' });
  }

  if (!unit_kerja || !unit_kerja.trim()) {
    return res.status(400).json({ error: 'Nama Unit Kerja baru wajib diisi.' });
  }

  try {
    const updateRes = await pool.query(
      `UPDATE uim_records SET unit_kerja = $1 WHERE id = ANY($2::int[]) RETURNING *;`,
      [unit_kerja.trim(), ids]
    );

    logActivity('bulk_update_uim_unit', `Memperbarui Unit Kerja ${updateRes.rowCount} UIM menjadi "${unit_kerja.trim()}"`, getDeviceInfo(req), req.user);
    res.json({ message: `Berhasil memperbarui Unit Kerja ${updateRes.rowCount} data UIM.`, updatedCount: updateRes.rowCount });
  } catch (err) {
    console.error('Error bulk updating UIM unit:', err);
    res.status(500).json({ error: 'Gagal memperbarui Unit Kerja massal.' });
  }
});

// Rename/Move all UIM users from an old unit to a new unit (Admin/Operator)
app.post('/api/uim/rename-unit', authenticateToken, requireRole(['admin', 'operator']), async (req, res) => {
  const { old_unit, new_unit } = req.body;

  if (!old_unit || !old_unit.trim()) {
    return res.status(400).json({ error: 'Unit Kerja asal wajib dipilih.' });
  }

  if (!new_unit || !new_unit.trim()) {
    return res.status(400).json({ error: 'Nama Unit Kerja baru wajib diisi.' });
  }

  try {
    const updateRes = await pool.query(
      `UPDATE uim_records SET unit_kerja = $1 WHERE unit_kerja = $2 RETURNING *;`,
      [new_unit.trim(), old_unit.trim()]
    );

    logActivity('rename_uim_unit', `Memindahkan ${updateRes.rowCount} UIM dari "${old_unit.trim()}" ke "${new_unit.trim()}"`, getDeviceInfo(req), req.user);
    res.json({ message: `Berhasil memindahkan ${updateRes.rowCount} data UIM dari "${old_unit.trim()}" ke "${new_unit.trim()}".`, updatedCount: updateRes.rowCount });
  } catch (err) {
    console.error('Error renaming UIM unit:', err);
    res.status(500).json({ error: 'Gagal memindahkan Unit Kerja.' });
  }
});

// Bulk Import UIM records (Admin/Operator)
app.post('/api/uim/import', authenticateToken, requireRole(['admin', 'operator']), async (req, res) => {
  const { records } = req.body;

  if (!Array.isArray(records) || records.length === 0) {
    return res.status(400).json({ error: 'Data UIM tidak valid atau kosong.' });
  }

  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    let insertedCount = 0;

    for (const item of records) {
      const uimCode = (item.uim_code || item.uim || item['USER ID'] || item['USER ID/ UIM'] || '').toString().trim();
      const fullName = (item.full_name || item.nama || item['NAMA LENGKAP'] || '').toString().trim();
      const unitKerja = (item.unit_kerja || item.unit || item['UNIT KERJA'] || 'Umum').toString().trim();

      if (uimCode && fullName) {
        await client.query(
          `INSERT INTO uim_records (uim_code, full_name, unit_kerja) VALUES ($1, $2, $3);`,
          [uimCode.toUpperCase(), fullName, unitKerja]
        );
        insertedCount++;
      }
    }

    await client.query('COMMIT');
    logActivity('import_uim', `Mengimpor ${insertedCount} data UIM baru.`, getDeviceInfo(req), req.user);
    res.json({ message: `Berhasil mengimpor ${insertedCount} data UIM.`, insertedCount });
  } catch (err) {
    await client.query('ROLLBACK');
    console.error('Error importing UIM:', err);
    res.status(500).json({ error: 'Gagal mengimpor data UIM.' });
  } finally {
    client.release();
  }
});

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
app.get('/api/stats', authenticateToken, requireRole(['admin']), async (req, res) => {
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

    // Bookmark categories distribution
    const bookmarkGroupsRes = await pool.query(`
      SELECT COALESCE(group_name, 'Umum') AS group_name, COUNT(*) AS count
      FROM bookmarks
      GROUP BY group_name
      ORDER BY count DESC;
    `);

    res.json({
      summary: summaryRes.rows[0],
      topSearches: topSearchRes.rows,
      docsPerUnit: unitRes.rows,
      dailyActivity: dailyRes.rows,
      todayActivity: todayRes.rows,
      devicesOnline: devicesRes.rows,
      bookmarkGroups: bookmarkGroupsRes.rows
    });
  } catch (err) {
    console.error('Error stats:', err);
    res.status(500).json({ error: err.message });
  }
});

// 1c. Reset Dashboard Statistics (Admin only)
app.post('/api/stats/reset', authenticateToken, requireRole(['admin']), async (req, res) => {
  try {
    // Truncate activity logs
    await pool.query('TRUNCATE TABLE activity_logs CASCADE;');
    
    // Log the reset event so there's at least one event in logs
    await pool.query(`
      INSERT INTO activity_logs (user_id, username, activity_type, activity_details, ip_address, browser, os, device_type)
      VALUES ($1, $2, 'reset_stats', 'Admin mereset statistik dashboard dan log aktivitas', $3, $4, $5, $6);
    `, [
      req.user.id, 
      req.user.username,
      req.ip,
      'Server Backend',
      'Node.js',
      'System'
    ]);
    
    res.json({ message: 'Statistik dashboard dan riwayat aktivitas berhasil direset.' });
  } catch (err) {
    console.error('Error resetting stats:', err);
    res.status(500).json({ error: err.message });
  }
});

// 2. Upload Excel File (Triggers background processing)
app.post('/api/upload', authenticateToken, requireRole(['admin', 'operator']), upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'Tidak ada file yang diunggah!' });
  }

  const jobId = Date.now().toString();
  console.log(`[Upload] File upload masuk. Memulai Job ID: ${jobId}`);

  // Start background process
  processFileInBackground(jobId, req.file.path, req.file.originalname, req.user);

  // Respond immediately with jobId
  res.json({
    success: true,
    message: 'File berhasil diunggah dan sedang diproses di latar belakang.',
    jobId: jobId
  });
});

// 2a. Upload Chunk (Bypasses Cloudflare 100MB POST payload limit)
app.post('/api/upload/chunk', authenticateToken, requireRole(['admin', 'operator']), uploadChunkMulter.single('chunk'), (req, res) => {
  const { uploadId, chunkIndex, totalChunks } = req.body;
  if (!uploadId || chunkIndex === undefined) {
    return res.status(400).json({ error: 'Parameter uploadId dan chunkIndex wajib diisi!' });
  }
  res.json({ success: true, chunkIndex: parseInt(chunkIndex), totalChunks: parseInt(totalChunks) });
});

// 2b. Merge Chunks & Trigger Background Processing
app.post('/api/upload/merge', authenticateToken, requireRole(['admin', 'operator']), async (req, res) => {
  const { uploadId, filename, totalChunks } = req.body;
  if (!uploadId || !filename || !totalChunks) {
    return res.status(400).json({ error: 'Parameter uploadId, filename, dan totalChunks wajib diisi!' });
  }

  const targetDir = path.join(chunkUploadDir, uploadId);
  if (!fs.existsSync(targetDir)) {
    return res.status(404).json({ error: 'Potongan file tidak ditemukan.' });
  }

  const safeFilename = path.basename(filename);
  const finalFileName = `${Date.now()}-${safeFilename}`;
  const finalFilePath = path.join(uploadDir, finalFileName);

  try {
    const writeStream = fs.createWriteStream(finalFilePath);

    for (let i = 0; i < totalChunks; i++) {
      const chunkPath = path.join(targetDir, `chunk_${i}`);
      if (!fs.existsSync(chunkPath)) {
        throw new Error(`Potongan file bagian ${i} hilang!`);
      }
      const chunkBuffer = fs.readFileSync(chunkPath);
      writeStream.write(chunkBuffer);
      try { fs.unlinkSync(chunkPath); } catch (_) {}
    }
    writeStream.end();

    try { fs.rmdirSync(targetDir); } catch (_) {}

    const jobId = Date.now().toString();
    console.log(`[Upload Chunked] Berkas "${filename}" (100% tergabung). Memulai Job ID: ${jobId}`);

    // Trigger background process for merged file
    processFileInBackground(jobId, finalFilePath, safeFilename, req.user);

    res.json({
      success: true,
      message: 'File berhasil digabungkan dan sedang diproses di latar belakang.',
      jobId: jobId
    });
  } catch (err) {
    console.error('Error merging chunks:', err);
    res.status(500).json({ error: 'Gagal menggabungkan potongan file: ' + err.message });
  }
});

// 3. Get background upload job status
app.get('/api/upload/status/:jobId', authenticateToken, requireRole(['admin', 'operator']), (req, res) => {
  const job = uploadJobs[req.params.jobId];
  if (!job) {
    return res.status(404).json({ error: 'Proses impor tidak ditemukan.' });
  }
  res.json(job);
});

// 4. Search document rows by search query
app.get('/api/search', authenticateToken, async (req, res) => {
  const searchStartTime = performance.now();
  const queryText = req.query.q;
  const fileId = req.query.fileId;
  const filterSheet = req.query.sheet; // Optional sheet filter
  const filterUnit = req.query.unit;   // Optional unit filter
  const uploaderId = req.query.uploaderId; // Optional uploader ID filter
  const useAI = req.query.ai === 'true'; // Checked in Filter Lanjutan

  if ((!queryText || queryText.trim() === '') && !fileId) {
    return res.json({ results: [], totalMatchesCount: 0, executionTimeMs: 0 });
  }

  try {
    let searchRows = [];
    let isAISearchUsed = false;
    let aiMethodUsed = 'none';

    if (fileId) {
      // Fetch all rows for a specific file (incorporating optional queryText filtering, targetRow, and bookmarks status)
      let filterClauses = ['r.file_id = $1'];
      const queryParams = [parseInt(fileId)];
      const targetRow = req.query.rowNumber ? parseInt(req.query.rowNumber) : null;
      const targetSheet = req.query.sheetName ? req.query.sheetName.trim() : null;

      if (targetRow) {
        const minRow = Math.max(1, targetRow - 40);
        const maxRow = targetRow + 40;
        filterClauses.push(`(r.row_number BETWEEN ${minRow} AND ${maxRow})`);
        if (targetSheet) {
          queryParams.push(targetSheet.trim());
          filterClauses.push(`LOWER(r.sheet_name) = LOWER($${queryParams.length})`);
        }
      } else if (queryText && queryText.trim() !== '') {
        const searchTerms = queryText.trim().split(/[\s\.]+/).filter(Boolean);
        const stopWords = new Set([
          'tampilkan', 'data', 'pada', 'yang', 'dan', 'di', 'untuk', 'ke', 'dari', 
          'adalah', 'itu', 'dengan', 'ini', 'oleh', 'seperti', 'adapun', 'atau', 
          'sebagai', 'tentang', 'yaitu', 'ia', 'kami', 'mereka', 'saya', 'anda', 
          'dia', 'kita', 'tahun', 'bulan', 'hari', 'file', 'berkas', 'dokumen', 
          'tabel', 'sheet', 'cari', 'temukan', 'info', 'informasi', 'lihat'
        ]);
        const keywords = searchTerms.filter(t => !stopWords.has(t.toLowerCase()) && t.length >= 2);
        const finalKeywords = keywords.length > 0 ? keywords : searchTerms;

        if (finalKeywords.length > 0) {
          const kwClauses = [];
          finalKeywords.forEach(kw => {
            queryParams.push(`%${kw}%`);
            kwClauses.push(`r.row_data::text ILIKE $${queryParams.length}`);
          });
          filterClauses.push(`(${kwClauses.join(' OR ')})`);
        }
      }

      if (filterSheet && filterSheet.trim() !== '') {
        queryParams.push(filterSheet.trim());
        filterClauses.push(`r.sheet_name = $${queryParams.length}`);
      }
      if (filterUnit && filterUnit.trim() !== '') {
        queryParams.push(`%${filterUnit.trim()}%`);
        filterClauses.push(`(r.row_data->>'KODE UNIT' ILIKE $${queryParams.length} OR r.row_data->>'KODE_UNIT' ILIKE $${queryParams.length})`);
      }
      if (uploaderId && uploaderId !== 'all' && req.user.role === 'admin') {
        queryParams.push(parseInt(uploaderId));
        filterClauses.push(`f.uploaded_by = $${queryParams.length}`);
      }

      queryParams.push(req.user.id);
      const userIdParamIndex = queryParams.length;

      let orderByClause = 'ORDER BY r.sheet_name, r.row_number';
      if (targetRow) {
        orderByClause = `ORDER BY ABS(r.row_number - ${parseInt(targetRow, 10)}) ASC, r.row_number ASC`;
      }

      const searchQuery = `
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
        LEFT JOIN bookmarks b ON b.row_id = r.id AND b.user_id = $${userIdParamIndex}
        WHERE ${filterClauses.join(' AND ')}
        ${orderByClause}
        LIMIT 250;
      `;
      const searchRes = await pool.query(searchQuery, queryParams);
      if (targetRow) {
        searchRows = searchRes.rows.sort((a, b) => {
          if (a.row_number === targetRow) return -1;
          if (b.row_number === targetRow) return 1;
          return Math.abs(a.row_number - targetRow) - Math.abs(b.row_number - targetRow);
        });
      } else {
        searchRows = searchRes.rows.sort((a, b) => a.row_number - b.row_number);
      }

      // FALLBACK 1: If targetRow + targetSheet returned 0 rows, retry targetRow without sheet constraint!
      if (searchRows.length === 0 && targetRow) {
        console.log(`[Search] targetRow ${targetRow} with sheet "${targetSheet}" returned 0 rows for fileId ${fileId}. Retrying without sheet constraint...`);
        const minRow = Math.max(1, targetRow - 40);
        const maxRow = targetRow + 40;
        const fbRes = await pool.query(`
          SELECT 
            r.id, r.file_id, r.sheet_name, r.row_number, r.row_data, f.filename, f.uploaded_at,
            (b.id IS NOT NULL) AS is_bookmarked
          FROM document_rows r
          JOIN uploaded_files f ON r.file_id = f.id
          LEFT JOIN bookmarks b ON b.row_id = r.id AND b.user_id = $2
          WHERE r.file_id = $1 AND (r.row_number BETWEEN ${minRow} AND ${maxRow})
          ORDER BY ABS(r.row_number - ${targetRow}) ASC
          LIMIT 250;
        `, [parseInt(fileId), req.user.id]);
        
        if (fbRes.rows.length > 0) {
          searchRows = fbRes.rows.sort((a, b) => {
            if (a.row_number === targetRow) return -1;
            if (b.row_number === targetRow) return 1;
            return Math.abs(a.row_number - targetRow) - Math.abs(b.row_number - targetRow);
          });
        }
      }

      // FALLBACK 2: Hanya tampilkan baris default jika TIDAK ada kata kunci pencarian yang dimasukkan!
      // Jika pengguna memasukkan kata kunci dan tidak ada baris yang cocok (0 baris),
      // maka JANGAN tampilkan data berkas sembarangan/tidak relevan.
      if (searchRows.length === 0 && (!queryText || queryText.trim() === '')) {
        const fbRes2 = await pool.query(`
          SELECT 
            r.id, r.file_id, r.sheet_name, r.row_number, r.row_data, f.filename, f.uploaded_at,
            (b.id IS NOT NULL) AS is_bookmarked
          FROM document_rows r
          JOIN uploaded_files f ON r.file_id = f.id
          LEFT JOIN bookmarks b ON b.row_id = r.id AND b.user_id = $2
          WHERE r.file_id = $1
          ORDER BY r.row_number ASC
          LIMIT 250;
        `, [parseInt(fileId), req.user.id]);
        searchRows = fbRes2.rows;
      }
    } else if (useAI) {
      // AI SEMANTIC HYBRID SEARCH
      isAISearchUsed = true;
      const searchTerms = queryText.trim().split(/[\s\.]+/).filter(Boolean);
      
      if (searchTerms.length === 0) {
        return res.json({ results: [] });
      }

      // Filter out Indonesian & English stop words that match almost every row
      const stopWords = new Set([
        'tampilkan', 'data', 'pada', 'yang', 'dan', 'di', 'untuk', 'ke', 'dari', 
        'adalah', 'itu', 'dengan', 'ini', 'oleh', 'seperti', 'adapun', 'atau', 
        'sebagai', 'tentang', 'yaitu', 'ia', 'kami', 'mereka', 'saya', 'anda', 
        'dia', 'kita', 'tahun', 'bulan', 'hari', 'file', 'berkas', 'dokumen', 
        'tabel', 'sheet', 'cari', 'temukan', 'info', 'informasi', 'lihat'
      ]);

      const keywords = searchTerms.filter(t => !stopWords.has(t.toLowerCase()) && t.length > 2);
      
      // Fallback to all search terms if everything got filtered out
      const finalKeywords = keywords.length > 0 ? keywords : searchTerms;

      // Add filters if any
      const filterClauses = [];
      const filterParams = [];
      if (filterSheet && filterSheet.trim() !== '') {
        filterParams.push(filterSheet.trim());
        filterClauses.push(`r.sheet_name = $${filterParams.length}`);
      }
      if (filterUnit && filterUnit.trim() !== '') {
        filterParams.push(`%${filterUnit.trim()}%`);
        filterClauses.push(`(r.row_data->>'KODE UNIT' ILIKE $${filterParams.length} OR r.row_data->>'KODE_UNIT' ILIKE $${filterParams.length})`);
      }
      if (uploaderId && uploaderId !== 'all' && req.user.role === 'admin') {
        filterParams.push(parseInt(uploaderId));
        filterClauses.push(`f.uploaded_by = $${filterParams.length}`);
      }

      const filterSql = filterClauses.length > 0 ? ' AND ' + filterClauses.join(' AND ') : '';

      // TIER 1: Try strict matching with AND first (extremely fast intersection)
      const queryParams1 = [...filterParams];
      const whereClauses1 = [];
      finalKeywords.forEach((kw) => {
        queryParams1.push(`%${kw}%`);
        whereClauses1.push(`r.row_data::text ILIKE $${queryParams1.length}`);
      });
      
      queryParams1.push(req.user.id);
      const userIdParamIndex1 = queryParams1.length;

      const strictQuery = `
        SELECT id, file_id, sheet_name, row_number, row_data, filename, uploaded_at, is_bookmarked
        FROM (
          SELECT 
            r.id, 
            r.file_id, 
            r.sheet_name, 
            r.row_number, 
            r.row_data, 
            f.filename, 
            f.uploaded_at,
            (b.id IS NOT NULL) AS is_bookmarked,
            ROW_NUMBER() OVER (
              PARTITION BY r.file_id
              ORDER BY r.sheet_name, r.row_number
            ) AS rn_per_file
          FROM document_rows r
          JOIN uploaded_files f ON r.file_id = f.id
          LEFT JOIN bookmarks b ON b.row_id = r.id AND b.user_id = $${userIdParamIndex1}
          WHERE (${whereClauses1.join(' AND ')}) ${filterSql}
        ) ranked
        WHERE rn_per_file <= 30
        ORDER BY uploaded_at DESC, file_id, sheet_name, row_number
        LIMIT 300;
      `;

      let candidates = [];
      console.log(`[Search AI] Executing Tier 1 Strict AND query for: [${finalKeywords.join(', ')}]`);
      const strictRes = await pool.query(strictQuery, queryParams1);
      candidates = strictRes.rows;

      // TIER 2: Fallback to looser OR matching if strict matching returned very few results
      if (candidates.length < 5 && finalKeywords.length > 1) {
        console.log(`[Search AI] Tier 1 Strict query returned only ${candidates.length} rows. Falling back to Tier 2 Loose OR query.`);
        const queryParams2 = [...filterParams];
        const whereClauses2 = [];
        finalKeywords.forEach((kw) => {
          queryParams2.push(`%${kw}%`);
          whereClauses2.push(`r.row_data::text ILIKE $${queryParams2.length}`);
        });

        queryParams2.push(req.user.id);
        const userIdParamIndex2 = queryParams2.length;

        const looseQuery = `
          SELECT id, file_id, sheet_name, row_number, row_data, filename, uploaded_at, is_bookmarked
          FROM (
            SELECT 
              r.id, 
              r.file_id, 
              r.sheet_name, 
              r.row_number, 
              r.row_data, 
              f.filename, 
              f.uploaded_at,
              (b.id IS NOT NULL) AS is_bookmarked,
              ROW_NUMBER() OVER (
                PARTITION BY r.file_id
                ORDER BY r.sheet_name, r.row_number
              ) AS rn_per_file
            FROM document_rows r
            JOIN uploaded_files f ON r.file_id = f.id
            LEFT JOIN bookmarks b ON b.row_id = r.id AND b.user_id = $${userIdParamIndex2}
            WHERE (${whereClauses2.join(' OR ')}) ${filterSql}
          ) ranked
          WHERE rn_per_file <= 30
          ORDER BY uploaded_at DESC, file_id, sheet_name, row_number
          LIMIT 300;
        `;
        const looseRes = await pool.query(looseQuery, queryParams2);
        
        // Merge without duplicates
        const existingIds = new Set(candidates.map(c => c.id));
        looseRes.rows.forEach(row => {
          if (!existingIds.has(row.id)) {
            candidates.push(row);
          }
        });
      }

      if (candidates.length > 0) {
        try {
          if (process.env.GEMINI_API_KEY) {
            console.log(`[Search AI] Menggunakan Gemini Embedding API untuk: "${queryText}"`);
            searchRows = await computeGeminiSemanticScores(queryText.trim(), candidates);
            aiMethodUsed = 'gemini';
          } else {
            throw new Error('GEMINI_API_KEY tidak ditemukan, beralih ke model vektor lokal.');
          }
        } catch (aiErr) {
          console.log(`[Search AI] Menggunakan Model Vektor Lokal (TF-IDF) untuk: "${queryText}". Alasan: ${aiErr.message}`);
          searchRows = computeLocalSemanticScores(queryText.trim(), candidates);
          aiMethodUsed = 'local';
        }

        // Sort by relevance score descending
        searchRows.sort((a, b) => b.relevance_score - a.relevance_score);
        
        // Filter out low relevance if results are abundant
        if (searchRows.length > 10) {
          searchRows = searchRows.filter(r => r.relevance_score > 0);
        }
      }
    } else {
      // STANDARD SEARCH
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
          const commonYears = ['2025', '2024', '2026', '2023', '2027', '2028', '2022', '2029', '2020', '2021'];
          commonYears.forEach(y => {
            years.push({ year4: y, year2: y.slice(2) });
          });
        }
        
        const boxCodeVariations = new Set();
        years.forEach(({ year4, year2 }) => {
          boxCodeVariations.add(`P${unitNum}.${boxLetter}.${folder6}.${year4}`);
          boxCodeVariations.add(`P${unitNum}.${boxLetter}.${folder5}.${year4}`);
          boxCodeVariations.add(`P${unitNum}.${boxLetter}.${folder3}.${year4}`);
          boxCodeVariations.add(`P${unitNum}.${boxLetter}.${folder6}.${year2}`);
          boxCodeVariations.add(`P${unitNum}.${boxLetter}.${folder5}.${year2}`);
          boxCodeVariations.add(`P${unitNum}.${boxLetter}.${folder3}.${year2}`);
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
        
        let filterClauses = [];
        if (filterSheet && filterSheet.trim() !== '') {
          queryParams.push(filterSheet.trim());
          filterClauses.push(`r.sheet_name = $${queryParams.length}`);
        }
        if (filterUnit && filterUnit.trim() !== '') {
          queryParams.push(`%${filterUnit.trim()}%`);
          filterClauses.push(`(r.row_data->>'KODE UNIT' ILIKE $${queryParams.length} OR r.row_data->>'KODE_UNIT' ILIKE $${queryParams.length})`);
        }
        if (uploaderId && uploaderId !== 'all' && req.user.role === 'admin') {
          queryParams.push(parseInt(uploaderId));
          filterClauses.push(`f.uploaded_by = $${queryParams.length}`);
        }

        const filterSql = filterClauses.length > 0 ? ' AND ' + filterClauses.join(' AND ') : '';
        queryParams.push(req.user.id);
        const userIdParamIndex = queryParams.length;

        const searchQuery = `
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
          LEFT JOIN bookmarks b ON b.row_id = r.id AND b.user_id = $${userIdParamIndex}
          WHERE (${containmentClauses.join(' OR ')}) ${filterSql}
          ORDER BY relevance_score DESC, f.uploaded_at DESC, r.file_id, r.sheet_name, r.row_number
          LIMIT 150;
        `;
        const searchRes = await pool.query(searchQuery, queryParams);
        searchRows = searchRes.rows;

        // FALLBACK: If fast JSONB containment search returned 0 rows
        if (searchRes.rowCount === 0) {
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
          if (uploaderId && uploaderId !== 'all' && req.user.role === 'admin') {
            fallbackParams.push(parseInt(uploaderId));
            fallbackFilterClauses.push(`f.uploaded_by = $${fallbackParams.length}`);
          }

          const fallbackFilterSql = fallbackFilterClauses.length > 0 ? ' AND ' + fallbackFilterClauses.join(' AND ') : '';
          fallbackParams.push(req.user.id);
          const fallbackUserIdParamIndex = fallbackParams.length;

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
            LEFT JOIN bookmarks b ON b.row_id = r.id AND b.user_id = $${fallbackUserIdParamIndex}
            WHERE (r.row_data::text ILIKE $1 OR (r.row_data::text ILIKE $2 AND r.row_data::text ILIKE $3)) ${fallbackFilterSql}
            ORDER BY relevance_score DESC, f.uploaded_at DESC, r.file_id, r.sheet_name, r.row_number
            LIMIT 150;
          `;
          const fallbackRes = await pool.query(fallbackQuery, fallbackParams);
          searchRows = fallbackRes.rows;
        }
      } else {
        // STANDARD GOOGLE-STYLE MATCH WITH RELEVANCE SCORING
        const rawTerms = queryText.trim().split(/[\s\.]+/).filter(Boolean);
        
        if (rawTerms.length === 0) {
          return res.json({ results: [] });
        }

        const stopWords = new Set([
          'tampilkan', 'data', 'pada', 'yang', 'dan', 'di', 'untuk', 'ke', 'dari', 
          'adalah', 'itu', 'dengan', 'ini', 'oleh', 'seperti', 'adapun', 'atau', 
          'sebagai', 'tentang', 'yaitu', 'ia', 'kami', 'mereka', 'saya', 'anda', 
          'dia', 'kita', 'tahun', 'bulan', 'hari', 'file', 'berkas', 'dokumen', 
          'tabel', 'sheet', 'cari', 'temukan', 'info', 'informasi', 'lihat', 'berapa',
          'apakah', 'siapa', 'bagaimana', 'apa', 'tolong', 'jelaskan', 'rangkum', 'semua',
          'daftar', 'total', 'ada', 'berurutan', 'entry', 'entri'
        ]);

        const searchTerms = rawTerms.filter(t => !stopWords.has(t.toLowerCase()) && t.length >= 2);
        const finalTerms = searchTerms.length > 0 ? searchTerms : rawTerms;

        const termClauses = [];
        const scoreExpressions = [];
        const andClauses = [];
        queryParams = [];
        
        finalTerms.forEach((term) => {
          const match = term.match(/^([a-zA-Z]+)(0*[1-9]\d*)$/);
          if (match) {
            const letters = match[1];
            const digits = match[2];
            
            queryParams.push(`%${term}%`, `%${letters}%`, `%${digits}%`);
            const idxTerm = queryParams.length - 2;
            const idxLetters = queryParams.length - 1;
            const idxDigits = queryParams.length;
            
            const clause = `(r.row_data::text ILIKE $${idxTerm} OR (r.row_data::text ILIKE $${idxLetters} AND r.row_data::text ILIKE $${idxDigits}))`;
            termClauses.push(clause);
            andClauses.push(clause);
            scoreExpressions.push(`(CASE WHEN r.row_data::text ILIKE $${idxTerm} THEN 50 ELSE 0 END)`);
            scoreExpressions.push(`(CASE WHEN f.filename ILIKE $${idxTerm} THEN 100 ELSE 0 END)`);
          } else {
            queryParams.push(`%${term}%`);
            const idxTerm = queryParams.length;
            const clause = `r.row_data::text ILIKE $${idxTerm}`;
            termClauses.push(clause);
            andClauses.push(clause);
            scoreExpressions.push(`(CASE WHEN r.row_data::text ILIKE $${idxTerm} THEN 40 ELSE 0 END)`);
            scoreExpressions.push(`(CASE WHEN f.filename ILIKE $${idxTerm} THEN 100 ELSE 0 END)`);
          }
        });

        queryParams.push(`%${queryText.trim()}%`);
        const idxFullQuery = queryParams.length;
        scoreExpressions.push(`(CASE WHEN r.row_data::text ILIKE $${idxFullQuery} THEN 300 ELSE 0 END)`);
        scoreExpressions.push(`(CASE WHEN f.filename ILIKE $${idxFullQuery} THEN 200 ELSE 0 END)`);

        if (andClauses.length > 1) {
          scoreExpressions.push(`(CASE WHEN (${andClauses.join(' AND ')}) THEN 200 ELSE 0 END)`);
        }

        const filterClauses = [`(${termClauses.join(' OR ')})`];

        if (filterSheet && filterSheet.trim() !== '') {
          queryParams.push(filterSheet.trim());
          filterClauses.push(`r.sheet_name = $${queryParams.length}`);
        }
        if (filterUnit && filterUnit.trim() !== '') {
          queryParams.push(`%${filterUnit.trim()}%`);
          filterClauses.push(`(r.row_data->>'KODE UNIT' ILIKE $${queryParams.length} OR r.row_data->>'KODE_UNIT' ILIKE $${queryParams.length})`);
        }
        if (uploaderId && uploaderId !== 'all' && req.user.role === 'admin') {
          queryParams.push(parseInt(uploaderId));
          filterClauses.push(`f.uploaded_by = $${queryParams.length}`);
        }

        queryParams.push(req.user.id);
        const userIdParamIndex = queryParams.length;

        const countQuery = `
          SELECT COUNT(*) 
          FROM document_rows r
          JOIN uploaded_files f ON r.file_id = f.id
          WHERE ${filterClauses.join(' AND ')}
        `;
        const countRes = await pool.query(countQuery, queryParams);
        const totalMatchesCount = parseInt(countRes.rows[0].count, 10);

        const searchQuery = `
          SELECT id, file_id, sheet_name, row_number, row_data, filename, uploaded_at, is_bookmarked, relevance_score
          FROM (
            SELECT 
              r.id, 
              r.file_id, 
              r.sheet_name, 
              r.row_number, 
              r.row_data, 
              f.filename, 
              f.uploaded_at,
              (b.id IS NOT NULL) AS is_bookmarked,
              (${scoreExpressions.join(' + ')}) as relevance_score,
              ROW_NUMBER() OVER (
                PARTITION BY r.file_id 
                ORDER BY (${scoreExpressions.join(' + ')}) DESC, r.sheet_name, r.row_number
              ) AS rn_per_file
            FROM document_rows r
            JOIN uploaded_files f ON r.file_id = f.id
            LEFT JOIN bookmarks b ON b.row_id = r.id AND b.user_id = $${userIdParamIndex}
            WHERE ${filterClauses.join(' AND ')}
          ) ranked
          WHERE rn_per_file <= 30
          ORDER BY relevance_score DESC, uploaded_at DESC, file_id, sheet_name, row_number
          LIMIT 300;
        `;
        const searchRes = await pool.query(searchQuery, queryParams);
        searchRows = searchRes.rows;
      }
    }

    // Log activity
    if (!fileId && queryText && queryText.trim() !== '') {
      const modeText = isAISearchUsed ? `AI (Method: ${aiMethodUsed})` : 'Standar';
      logActivity('search', `Mencari kata kunci [Mode: ${modeText}]: "${queryText.trim()}" (Menghasilkan ${searchRows.length} baris data)`, getDeviceInfo(req), req.user);
    } else if (fileId) {
      const filename = searchRows[0] ? searchRows[0].filename : `ID #${fileId}`;
      logActivity('view_file', `Membuka/melihat berkas: "${filename}"`, getDeviceInfo(req), req.user);
    }

    // Group results by file_id and sort by relevance score DESC (most relevant file & row first)
    const fileMap = new Map();
    searchRows.forEach(row => {
      const fid = row.file_id;
      const score = parseInt(row.relevance_score, 10) || 0;
      if (!fileMap.has(fid)) {
        fileMap.set(fid, {
          fileId: fid,
          filename: row.filename,
          uploadedAt: row.uploaded_at,
          maxScore: score,
          matches: []
        });
      } else {
        const fileObj = fileMap.get(fid);
        if (score > fileObj.maxScore) {
          fileObj.maxScore = score;
        }
      }
      fileMap.get(fid).matches.push({
        id: row.id,
        sheetName: row.sheet_name,
        rowNumber: row.row_number,
        rowData: row.row_data,
        isBookmarked: !!row.is_bookmarked,
        relevanceScore: score || 100
      });
    });

    // 1. Sort files array by maxScore DESC (most relevant file first), then uploadedAt DESC
    const sortedFiles = Array.from(fileMap.values()).sort((a, b) => {
      if (b.maxScore !== a.maxScore) return b.maxScore - a.maxScore;
      return new Date(b.uploadedAt) - new Date(a.uploadedAt);
    });

    // 2. Sort matches inside each file by relevanceScore DESC (most relevant row at the top!)
    //    If targetRowNumber parameter exists, place target row at index 0
    const targetRow = req.query.rowNumber ? parseInt(req.query.rowNumber, 10) : null;
    sortedFiles.forEach(file => {
      file.matches.sort((a, b) => {
        if (targetRow) {
          if (a.rowNumber === targetRow) return -1;
          if (b.rowNumber === targetRow) return 1;
        }
        if (b.relevanceScore !== a.relevanceScore) {
          return b.relevanceScore - a.relevanceScore; // Best match at top!
        }
        return a.rowNumber - b.rowNumber;
      });
    });

    const executionTimeMs = Math.round(performance.now() - searchStartTime);
    const totalMatchCountCalculated = searchRows.length;

    res.json({ 
      results: sortedFiles,
      totalMatchesCount: totalMatchCountCalculated,
      executionTimeMs,
      aiSearchUsed: isAISearchUsed,
      aiMethod: aiMethodUsed
    });
  } catch (err) {
    console.error('Error saat melakukan pencarian/pratinjau:', err);
    res.status(500).json({ error: 'Gagal mengambil data dari database: ' + err.message });
  }
});

// 5a-2. AI Natural Language Query Search API
app.post('/api/search/nl-query', authenticateToken, async (req, res) => {
  const { prompt } = req.body;
  if (!prompt || !prompt.trim()) {
    return res.status(400).json({ error: 'Kalimat pencarian Bahasa Alami tidak boleh kosong.' });
  }

  const queryText = prompt.trim();
  const userId = req.user.id;

  try {
    const stopWords = new Set(['cari', 'tampilkan', 'data', 'pegawai', 'yang', 'dan', 'di', 'untuk', 'ke', 'dari', 'file', 'berkas', 'dokumen', 'laporan', 'tolong', 'ada', 'siapa', 'berapa']);
    const rawTokens = queryText.toLowerCase().split(/[\s,.-]+/).filter(Boolean);
    const keywords = rawTokens.filter(t => !stopWords.has(t) && t.length > 1);
    const finalKeywords = keywords.length > 0 ? keywords : rawTokens;

    const filterClauses = [];
    const queryParams = [];

    finalKeywords.forEach((kw) => {
      queryParams.push(`%${kw}%`);
      filterClauses.push(`r.row_data::text ILIKE $${queryParams.length}`);
    });

    const userIdParamIndex = queryParams.length + 1;
    queryParams.push(userId);

    const searchQuery = `
      SELECT 
        r.id, r.file_id, r.sheet_name, r.row_number, r.row_data, 
        f.filename, f.uploaded_at,
        (b.id IS NOT NULL) AS is_bookmarked
      FROM document_rows r
      JOIN uploaded_files f ON r.file_id = f.id
      LEFT JOIN bookmarks b ON b.row_id = r.id AND b.user_id = $${userIdParamIndex}
      ${filterClauses.length > 0 ? `WHERE ${filterClauses.join(' OR ')}` : ''}
      ORDER BY r.row_number ASC
      LIMIT 150;
    `;

    const searchRes = await pool.query(searchQuery, queryParams);
    const searchRows = searchRes.rows;

    const fileMap = new Map();
    searchRows.forEach(row => {
      const fid = row.file_id;
      if (!fileMap.has(fid)) {
        fileMap.set(fid, {
          fileId: fid,
          filename: row.filename,
          uploadedAt: row.uploaded_at,
          matches: []
        });
      }
      fileMap.get(fid).matches.push({
        id: row.id,
        sheetName: row.sheet_name,
        rowNumber: row.row_number,
        rowData: row.row_data,
        isBookmarked: !!row.is_bookmarked
      });
    });

    const files = Array.from(fileMap.values());

    const aiSummary = searchRows.length > 0
      ? `✨ **AI Analysis**: Ditemukan **${searchRows.length} baris data** di **${files.length} dokumen berkas** yang cocok dengan kueri Bahasa Alami: *"${queryText}"*.`
      : `⚠️ **AI Analysis**: Tidak ditemukan data spesifik yang cocok dengan kueri Bahasa Alami: *"${queryText}"*. Coba gunakan kata kunci atau nama unit yang sejenis.`;

    logActivity('search_nl_ai', `Pencarian AI Bahasa Alami: "${queryText}" (Menghasilkan ${searchRows.length} data)`, getDeviceInfo(req), req.user);

    res.json({
      query: queryText,
      totalMatches: searchRows.length,
      filesCount: files.length,
      files,
      aiSummary
    });
  } catch (err) {
    console.error('Error NL query search:', err);
    res.status(500).json({ error: 'Gagal memproses pencarian Bahasa Alami: ' + err.message });
  }
});

// 5. Get list of all uploaded files
app.get('/api/files', authenticateToken, async (req, res) => {
  try {
    const query = `
      SELECT 
        uf.id, 
        uf.filename, 
        uf.sheet_names, 
        uf.uploaded_at, 
        uf.row_count,
        uf.uploaded_by,
        u.username AS uploader_username,
        u.full_name AS uploader_fullname
      FROM uploaded_files uf
      LEFT JOIN users u ON uf.uploaded_by = u.id
      ORDER BY uf.uploaded_at DESC;
    `;
    const filesRes = await pool.query(query);
    res.json({ files: filesRes.rows });
  } catch (err) {
    console.error('Error saat mengambil daftar file:', err);
    res.status(500).json({ error: 'Gagal mengambil daftar file: ' + err.message });
  }
});

// 5b. Get activity logs (with search + type filter + pagination)
app.get('/api/logs', authenticateToken, requireRole(['admin']), async (req, res) => {
  try {
    const page  = Math.max(1, parseInt(req.query.page)  || 1);
    const limit = Math.min(500, Math.max(10, parseInt(req.query.limit) || 50));
    const q     = (req.query.q || '').trim();
    const type  = (req.query.type || '').trim();
    const offset = (page - 1) * limit;

    const whereClauses = [];
    const params = [];

    if (q) {
      params.push(`%${q}%`);
      whereClauses.push(`(activity_details ILIKE $${params.length} OR activity_type ILIKE $${params.length} OR ip_address ILIKE $${params.length} OR browser ILIKE $${params.length} OR os ILIKE $${params.length} OR device_label ILIKE $${params.length} OR username ILIKE $${params.length})`);
    }

    if (type && type !== 'ALL') {
      params.push(type);
      whereClauses.push(`activity_type = $${params.length}`);
    }

    const conditions = whereClauses.length > 0 ? `WHERE ${whereClauses.join(' AND ')}` : '';

    const countRes = await pool.query(
      `SELECT COUNT(*) FROM activity_logs ${conditions}`, params
    );
    const total = parseInt(countRes.rows[0].count);

    const dataRes = await pool.query(
      `SELECT id, user_id, username, activity_type, activity_details,
              ip_address, browser, os, device_type, device_label, engine, user_agent,
              created_at
       FROM activity_logs
       ${conditions}
       ORDER BY created_at DESC
       LIMIT $${params.length + 1} OFFSET $${params.length + 2}`,
      [...params, limit, offset]
    );

    // Fetch distinct activity types for dropdown filter
    const typesRes = await pool.query(`SELECT DISTINCT activity_type FROM activity_logs WHERE activity_type IS NOT NULL AND activity_type != '' ORDER BY activity_type ASC;`);
    const availableTypes = typesRes.rows.map(r => r.activity_type);

    res.json({
      logs: dataRes.rows,
      total,
      page,
      limit,
      totalPages: Math.ceil(total / limit),
      availableTypes
    });
  } catch (err) {
    console.error('Error saat mengambil log aktivitas:', err);
    res.status(500).json({ error: 'Gagal mengambil log: ' + err.message });
  }
});

// 5c. Reset / clear all activity logs (Admin only)
app.delete('/api/logs', authenticateToken, requireRole(['admin']), async (req, res) => {
  try {
    await pool.query('TRUNCATE TABLE activity_logs CASCADE;');
    
    // Create new audit log for the clear action
    await pool.query(`
      INSERT INTO activity_logs (user_id, username, activity_type, activity_details, ip_address, browser, os, device_type)
      VALUES ($1, $2, 'clear_logs', 'Admin membersihkan/reset seluruh riwayat log aktivitas', $3, $4, $5, $6);
    `, [
      req.user.id, 
      req.user.username,
      req.ip,
      'Server Backend',
      'Node.js',
      'System'
    ]);
    
    res.json({ message: 'Seluruh riwayat log aktivitas berhasil direset.' });
  } catch (err) {
    console.error('Error clearing activity logs:', err);
    res.status(500).json({ error: err.message });
  }
});

// 6. Delete an uploaded file
app.delete('/api/files/:id', authenticateToken, requireRole(['admin']), async (req, res) => {
  const fileId = req.params.id;
  try {
    const deleteQuery = `DELETE FROM uploaded_files WHERE id = $1 RETURNING filename;`;
    const deleteRes = await pool.query(deleteQuery, [fileId]);
    
    if (deleteRes.rowCount === 0) {
      return res.status(404).json({ error: 'File tidak ditemukan!' });
    }

    const filename = deleteRes.rows[0].filename;
    await logActivity('delete', `Menghapus berkas "${filename}" beserta seluruh data terkait.`, getDeviceInfo(req), req.user);

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
app.post('/api/files/bulk-delete', authenticateToken, requireRole(['admin']), async (req, res) => {
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
    await logActivity('delete', `Menghapus massal ${deleteRes.rowCount} berkas: ${filenames.join(', ')}`, device, req.user);

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

// Helper: Worker thread for CPU-heavy Excel file generation without blocking event loop
function generateExcelWorker(sheetsMap) {
  return new Promise((resolve, reject) => {
    const workerScript = `
      const { parentPort, workerData } = require('worker_threads');
      const xlsx = require('xlsx');

      try {
        const wb = xlsx.utils.book_new();
        for (const sheetName of Object.keys(workerData)) {
          const ws = xlsx.utils.json_to_sheet(workerData[sheetName]);
          xlsx.utils.book_append_sheet(wb, ws, sheetName.substring(0, 31));
        }
        const buffer = xlsx.write(wb, { type: 'buffer', bookType: 'xlsx', compression: false });
        parentPort.postMessage({ success: true, buffer });
      } catch (err) {
        parentPort.postMessage({ success: false, error: err.message });
      }
    `;

    const worker = new Worker(workerScript, { eval: true, workerData: sheetsMap });
    worker.on('message', (msg) => {
      if (msg.success) resolve(Buffer.from(msg.buffer));
      else reject(new Error(msg.error));
      worker.terminate();
    });
    worker.on('error', (err) => {
      reject(err);
      worker.terminate();
    });
  });
}

// 6c. Download uploaded file reconstructed from database
app.get('/api/files/:id/download', authenticateToken, async (req, res) => {
  const fileId = req.params.id;
  try {
    const fileRes = await pool.query(`SELECT id, filename, sheet_names FROM uploaded_files WHERE id = $1`, [fileId]);
    if (fileRes.rows.length === 0) {
      return res.status(404).json({ error: 'File tidak ditemukan!' });
    }

    const { filename } = fileRes.rows[0];

    const rowsRes = await pool.query(
      `SELECT sheet_name, row_number, row_data FROM document_rows WHERE file_id = $1 ORDER BY sheet_name, row_number ASC`,
      [fileId]
    );

    if (rowsRes.rows.length === 0) {
      return res.status(404).json({ error: 'Data baris file ini tidak ditemukan!' });
    }

    // Group rows by sheet_name
    const sheetsMap = {};
    for (const r of rowsRes.rows) {
      if (!sheetsMap[r.sheet_name]) {
        sheetsMap[r.sheet_name] = [];
      }
      sheetsMap[r.sheet_name].push(r.row_data);
    }

    // Generate Excel file in background Worker Thread so main thread never blocks
    const buffer = await generateExcelWorker(sheetsMap);

    const safeFilename = filename.endsWith('.xlsx') || filename.endsWith('.xls') ? filename : `${filename}.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(safeFilename)}"`);
    res.setHeader('Content-Length', buffer.length);

    res.send(buffer);

    // Asynchronous non-blocking activity logging
    logActivity('download', `Mengunduh berkas Excel "${filename}".`, getDeviceInfo(req), req.user).catch(err => {
      console.error('Error logging download activity:', err);
    });
  } catch (err) {
    console.error('Error saat mengunduh file:', err);
    if (!res.headersSent) {
      res.status(500).json({ error: 'Gagal mengunduh file: ' + err.message });
    }
  }
});

// 7. Get bookmarks
app.get('/api/bookmarks', authenticateToken, async (req, res) => {
  try {
    let query;
    let params = [];
    
    if (req.user.role === 'admin' && req.query.userId) {
      if (req.query.userId === 'all') {
        query = `
          SELECT 
            b.id as bookmark_id,
            b.created_at as bookmarked_at,
            b.notes,
            b.group_name as group_name,
            b.user_id as owner_id,
            r.id,
            r.file_id,
            r.sheet_name,
            r.row_number,
            r.row_data,
            f.filename,
            u.username as owner_username,
            true AS is_bookmarked
          FROM bookmarks b
          JOIN document_rows r ON b.row_id = r.id
          JOIN uploaded_files f ON r.file_id = f.id
          JOIN users u ON b.user_id = u.id
          ORDER BY u.username ASC, b.created_at DESC;
        `;
      } else if (req.query.userId.includes(',')) {
        // Handle comma-separated list of IDs
        const ids = req.query.userId.split(',').map(id => parseInt(id.trim())).filter(id => !isNaN(id));
        query = `
          SELECT 
            b.id as bookmark_id,
            b.created_at as bookmarked_at,
            b.notes,
            b.group_name as group_name,
            b.user_id as owner_id,
            r.id,
            r.file_id,
            r.sheet_name,
            r.row_number,
            r.row_data,
            f.filename,
            u.username as owner_username,
            true AS is_bookmarked
          FROM bookmarks b
          JOIN document_rows r ON b.row_id = r.id
          JOIN uploaded_files f ON r.file_id = f.id
          JOIN users u ON b.user_id = u.id
          WHERE b.user_id = ANY($1::int[])
          ORDER BY u.username ASC, b.created_at DESC;
        `;
        params = [ids];
      } else {
        let targetUserId = req.user.id;
        if (req.query.userId !== 'mine') {
          targetUserId = parseInt(req.query.userId);
        }
        query = `
          SELECT 
            b.id as bookmark_id,
            b.created_at as bookmarked_at,
            b.notes,
            b.group_name as group_name,
            b.user_id as owner_id,
            r.id,
            r.file_id,
            r.sheet_name,
            r.row_number,
            r.row_data,
            f.filename,
            u.username as owner_username,
            true AS is_bookmarked
          FROM bookmarks b
          JOIN document_rows r ON b.row_id = r.id
          JOIN uploaded_files f ON r.file_id = f.id
          JOIN users u ON b.user_id = u.id
          WHERE b.user_id = $1
          ORDER BY b.created_at DESC;
        `;
        params = [targetUserId];
      }
    } else {
      query = `
        SELECT 
          b.id as bookmark_id,
          b.created_at as bookmarked_at,
          b.notes,
          b.group_name as group_name,
          b.user_id as owner_id,
          r.id,
          r.file_id,
          r.sheet_name,
          r.row_number,
          r.row_data,
          f.filename,
          u.username as owner_username,
          true AS is_bookmarked
        FROM bookmarks b
        JOIN document_rows r ON b.row_id = r.id
        JOIN uploaded_files f ON r.file_id = f.id
        JOIN users u ON b.user_id = u.id
        WHERE b.user_id = $1
        ORDER BY b.created_at DESC;
      `;
      params = [req.user.id];
    }
    
    const bookmarkedRes = await pool.query(query, params);
    res.json({ bookmarks: bookmarkedRes.rows });
  } catch (err) {
    console.error('Error bookmarks:', err);
    res.status(500).json({ error: err.message });
  }
});

// 8. Add bookmark
app.post('/api/bookmarks', authenticateToken, async (req, res) => {
  const { rowId, userId, group_name } = req.body;
  if (!rowId) {
    return res.status(400).json({ error: 'rowId diperlukan!' });
  }
  try {
    let targetUserId = req.user.id;
    let targetUsername = req.user.username;
    if (req.user.role === 'admin' && userId) {
      targetUserId = parseInt(userId);
      const userLookup = await pool.query('SELECT username FROM users WHERE id = $1;', [targetUserId]);
      if (userLookup.rowCount > 0) {
        targetUsername = userLookup.rows[0].username;
      }
    }

    const cleanGroupName = group_name && group_name.trim() !== '' ? group_name.trim() : 'Umum';

    await pool.query(
      `INSERT INTO bookmarks (user_id, row_id, group_name) 
       VALUES ($1, $2, $3) 
       ON CONFLICT (user_id, row_id) 
       DO UPDATE SET group_name = EXCLUDED.group_name;`,
      [targetUserId, rowId, cleanGroupName]
    );

    const logDetails = req.user.id === targetUserId
      ? `Menyimpan baris #${rowId} ke bookmark`
      : `Admin menyimpan baris #${rowId} ke bookmark milik akun "${targetUsername}"`;
    logActivity('add_bookmark', logDetails, getDeviceInfo(req), req.user);

    res.json({ success: true, message: 'Data berhasil disimpan ke bookmark.' });
  } catch (err) {
    console.error('Error add bookmark:', err);
    res.status(500).json({ error: err.message });
  }
});

// 9. Remove all bookmarks (Admin or self)
app.delete('/api/bookmarks', authenticateToken, async (req, res) => {
  try {
    let targetUserId = req.user.id;
    let targetUsername = req.user.username;
    if (req.user.role === 'admin' && req.query.userId) {
      targetUserId = parseInt(req.query.userId);
      const userLookup = await pool.query('SELECT username FROM users WHERE id = $1;', [targetUserId]);
      if (userLookup.rowCount > 0) {
        targetUsername = userLookup.rows[0].username;
      }
    }

    const deleteRes = await pool.query(`DELETE FROM bookmarks WHERE user_id = $1;`, [targetUserId]);

    const logDetails = req.user.id === targetUserId
      ? `Membersihkan semua bookmark (${deleteRes.rowCount} item)`
      : `Admin membersihkan semua bookmark (${deleteRes.rowCount} item) milik akun "${targetUsername}"`;
    logActivity('delete_bookmark', logDetails, getDeviceInfo(req), req.user);

    res.json({ success: true, message: `Berhasil menghapus seluruh bookmark (${deleteRes.rowCount} item).` });
  } catch (err) {
    console.error('Error clear all bookmarks:', err);
    res.status(500).json({ error: err.message });
  }
});

// 9b. Remove bookmark
app.delete('/api/bookmarks/:rowId', authenticateToken, async (req, res) => {
  const rowId = req.params.rowId;
  try {
    let targetUserId = req.user.id;
    let targetUsername = req.user.username;
    if (req.user.role === 'admin' && req.query.userId) {
      targetUserId = parseInt(req.query.userId);
      const userLookup = await pool.query('SELECT username FROM users WHERE id = $1;', [targetUserId]);
      if (userLookup.rowCount > 0) {
        targetUsername = userLookup.rows[0].username;
      }
    }

    await pool.query(`DELETE FROM bookmarks WHERE row_id = $1 AND user_id = $2;`, [rowId, targetUserId]);

    const logDetails = req.user.id === targetUserId
      ? `Menghapus baris #${rowId} dari bookmark`
      : `Admin menghapus baris #${rowId} dari bookmark milik akun "${targetUsername}"`;
    logActivity('delete_bookmark', logDetails, getDeviceInfo(req), req.user);

    res.json({ success: true, message: 'Bookmark berhasil dihapus.' });
  } catch (err) {
    console.error('Error delete bookmark:', err);
    res.status(500).json({ error: err.message });
  }
});

// 10. Update bookmark notes & group (can be updated by user for own bookmarks, or admin for any bookmark)
app.put('/api/bookmarks/:rowId', authenticateToken, async (req, res) => {
  const rowId = req.params.rowId;
  const { notes, userId, group_name } = req.body;
  try {
    let targetUserId = req.user.id;
    let targetUsername = req.user.username;
    if (req.user.role === 'admin' && userId) {
      targetUserId = parseInt(userId);
      const userLookup = await pool.query('SELECT username FROM users WHERE id = $1;', [targetUserId]);
      if (userLookup.rowCount > 0) {
        targetUsername = userLookup.rows[0].username;
      }
    }

    const cleanGroupName = group_name && group_name.trim() !== '' ? group_name.trim() : 'Umum';
    
    const result = await pool.query(
      `UPDATE bookmarks 
       SET notes = $1, group_name = $2 
       WHERE row_id = $3 AND user_id = $4 
       RETURNING *;`,
      [notes, cleanGroupName, rowId, targetUserId]
    );
    
    if (result.rowCount === 0) {
      return res.status(404).json({ error: 'Bookmark tidak ditemukan.' });
    }

    const logDetails = req.user.id === targetUserId
      ? `Memperbarui catatan pada bookmark baris #${rowId}`
      : `Admin memperbarui catatan pada bookmark baris #${rowId} milik akun "${targetUsername}"`;
    logActivity('edit_bookmark', logDetails, getDeviceInfo(req), req.user);

    res.json({ success: true, bookmark: result.rows[0] });
  } catch (err) {
    console.error('Error updating bookmark notes:', err);
    res.status(500).json({ error: err.message });
  }
});

// 11. Update document row data (can be updated by admin or operator)
app.put('/api/document-rows/:id', authenticateToken, async (req, res) => {
  const rowId = req.params.id;
  const { row_data } = req.body;
  if (!row_data) {
    return res.status(400).json({ error: 'row_data diperlukan!' });
  }
  try {
    // Only admin or operator can edit document rows
    if (req.user.role !== 'admin' && req.user.role !== 'operator') {
      return res.status(403).json({ error: 'Anda tidak memiliki hak akses untuk mengedit dokumen!' });
    }
    
    const result = await pool.query(
      `UPDATE document_rows SET row_data = $1 WHERE id = $2 RETURNING *;`,
      [JSON.stringify(row_data), rowId]
    );
    
    if (result.rowCount === 0) {
      return res.status(404).json({ error: 'Baris dokumen tidak ditemukan.' });
    }
    
    // Log the edit activity
    await logActivity('edit_document', `Mengedit baris #${rowId} pada dokumen.`, getDeviceInfo(req), req.user);
    
    res.json({ success: true, row: result.rows[0] });
  } catch (err) {
    console.error('Error updating document row:', err);
    res.status(500).json({ error: err.message });
  }
});

// 12. Fetch global notifications for a user (with read status)
app.get('/api/notifications', authenticateToken, async (req, res) => {
  try {
    const query = `
      SELECT 
        n.id,
        n.message,
        n.recipient_id,
        n.created_at,
        (unr.read_at IS NOT NULL) AS is_read
      FROM notifications n
      LEFT JOIN user_notification_reads unr ON unr.notification_id = n.id AND unr.user_id = $1
      WHERE n.recipient_id IS NULL OR n.recipient_id = $1
      ORDER BY n.created_at DESC
      LIMIT 50;
    `;
    const result = await pool.query(query, [req.user.id]);
    res.json({ notifications: result.rows });
  } catch (err) {
    console.error('Error fetching notifications:', err);
    res.status(500).json({ error: err.message });
  }
});

// 13. Send global or targeted notification (Admin only)
app.post('/api/notifications', authenticateToken, requireRole(['admin']), async (req, res) => {
  const { message, recipientId } = req.body;
  if (!message || !message.trim()) {
    return res.status(400).json({ error: 'Pesan notifikasi wajib diisi!' });
  }
  try {
    let targetRecipientId = null;
    let logMsg = `Admin mengirim notifikasi global: "${message.trim()}"`;
    
    if (recipientId && recipientId !== 'all') {
      targetRecipientId = parseInt(recipientId);
      const recipientRes = await pool.query('SELECT username FROM users WHERE id = $1;', [targetRecipientId]);
      if (recipientRes.rowCount > 0) {
        logMsg = `Admin mengirim notifikasi ke ${recipientRes.rows[0].username}: "${message.trim()}"`;
      }
    }

    const result = await pool.query(
      'INSERT INTO notifications (message, recipient_id) VALUES ($1, $2) RETURNING *;',
      [message.trim(), targetRecipientId]
    );
    // Log this activity
    await logActivity('send_notification', logMsg, getDeviceInfo(req), req.user);
    res.json({ success: true, notification: result.rows[0] });
  } catch (err) {
    console.error('Error sending notification:', err);
    res.status(500).json({ error: err.message });
  }
});

// 14. Mark notifications as read for current user
app.post('/api/notifications/read', authenticateToken, async (req, res) => {
  const { notificationId } = req.body;
  try {
    if (notificationId) {
      // Mark specific notification as read
      await pool.query(
        `INSERT INTO user_notification_reads (user_id, notification_id) 
         VALUES ($1, $2) 
         ON CONFLICT (user_id, notification_id) DO NOTHING;`,
        [req.user.id, notificationId]
      );
    } else {
      // Mark all readable notifications as read
      const readableNotifs = await pool.query(
        'SELECT id FROM notifications WHERE recipient_id IS NULL OR recipient_id = $1;',
        [req.user.id]
      );
      for (const notif of readableNotifs.rows) {
        await pool.query(
          `INSERT INTO user_notification_reads (user_id, notification_id) 
           VALUES ($1, $2) 
           ON CONFLICT (user_id, notification_id) DO NOTHING;`,
          [req.user.id, notif.id]
        );
      }
    }
    res.json({ success: true });
  } catch (err) {
    console.error('Error marking notifications as read:', err);
    res.status(500).json({ error: err.message });
  }
});

// 15. RAG Document AI Chatbot API (Full Document Catalog & Target File Scope)
app.post('/api/chat', authenticateToken, async (req, res) => {
  const { message, fileId, modelMode } = req.body;
  if (!message || !message.trim()) {
    return res.status(400).json({ error: 'Pesan kueri tidak boleh kosong!' });
  }

  const targetFileId = (fileId && fileId !== 'all') ? parseInt(fileId, 10) : null;
  const queryText = message.trim();
  const searchTerms = queryText.split(/[\s\.]+/).filter(Boolean);
  
  // Indonesian & English stop words
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
  const finalKeywords = keywords.length > 0 ? keywords : searchTerms;

  try {
    // 1. Fetch metadata of ALL uploaded files in the database catalog
    const allFilesRes = await pool.query(`
      SELECT 
        f.id, 
        f.filename, 
        f.uploaded_at, 
        COALESCE(u.username, 'Sistem') as uploader_name,
        COUNT(r.id) as total_rows
      FROM uploaded_files f
      LEFT JOIN users u ON f.uploaded_by = u.id
      LEFT JOIN document_rows r ON r.file_id = f.id
      GROUP BY f.id, f.filename, f.uploaded_at, u.username
      ORDER BY f.uploaded_at DESC;
    `);
    const allFiles = allFilesRes.rows;
    const activeCatalogFiles = targetFileId 
      ? allFiles.filter(f => f.id === targetFileId)
      : allFiles;

    // 2. Retrieve candidate rows across distinct files with Fast Smart Tiered Strategy (<2s response time)
    let candidates = [];
    if (finalKeywords.length > 0 || targetFileId) {
      // Sort keywords by length descending (longest = most specific)
      const sortedKws = [...finalKeywords].sort((a, b) => b.length - a.length);

      // Strategy Tier 1: Try strict AND search on top 3 most specific keywords
      const topKws = sortedKws.slice(0, 3);
      const queryParams1 = [];
      const filterClauses1 = [];

      topKws.forEach((kw) => {
        queryParams1.push(`%${kw}%`);
        filterClauses1.push(`r.row_data::text ILIKE $${queryParams1.length}`);
      });

      if (targetFileId) {
        queryParams1.push(targetFileId);
        filterClauses1.push(`r.file_id = $${queryParams1.length}`);
      }

      const whereSql1 = filterClauses1.length > 0 ? `WHERE ${filterClauses1.join(' AND ')}` : '';
      const query1 = `
        SELECT 
          r.id, r.file_id, r.sheet_name, r.row_number, r.row_data, f.filename, f.uploaded_at,
          100 AS relevance_score
        FROM document_rows r
        JOIN uploaded_files f ON r.file_id = f.id
        ${whereSql1}
        ORDER BY r.row_number ASC
        LIMIT 60;
      `;

      let res = await pool.query(query1, queryParams1);

      // Strategy Tier 2: Fallback to top 2 longest keywords if Tier 1 returned 0
      if (res.rows.length === 0 && sortedKws.length >= 2) {
        const top2Kws = sortedKws.slice(0, 2);
        const queryParams2 = [];
        const filterClauses2 = [];

        top2Kws.forEach((kw) => {
          queryParams2.push(`%${kw}%`);
          filterClauses2.push(`r.row_data::text ILIKE $${queryParams2.length}`);
        });

        if (targetFileId) {
          queryParams2.push(targetFileId);
          filterClauses2.push(`r.file_id = $${queryParams2.length}`);
        }

        const whereSql2 = filterClauses2.length > 0 ? `WHERE ${filterClauses2.join(' AND ')}` : '';
        const query2 = `
          SELECT 
            r.id, r.file_id, r.sheet_name, r.row_number, r.row_data, f.filename, f.uploaded_at,
            50 AS relevance_score
          FROM document_rows r
          JOIN uploaded_files f ON r.file_id = f.id
          ${whereSql2}
          ORDER BY r.row_number ASC
          LIMIT 60;
        `;
        res = await pool.query(query2, queryParams2);
      }

      // Strategy Tier 3: Fallback to single longest keyword if Tier 2 returned 0
      if (res.rows.length === 0 && sortedKws.length >= 1) {
        const singleKw = [sortedKws[0]];
        const queryParams3 = [`%${singleKw[0]}%`];
        const filterClauses3 = [`r.row_data::text ILIKE $1`];

        if (targetFileId) {
          queryParams3.push(targetFileId);
          filterClauses3.push(`r.file_id = $2`);
        }

        const whereSql3 = `WHERE ${filterClauses3.join(' AND ')}`;
        const query3 = `
          SELECT 
            r.id, r.file_id, r.sheet_name, r.row_number, r.row_data, f.filename, f.uploaded_at,
            20 AS relevance_score
          FROM document_rows r
          JOIN uploaded_files f ON r.file_id = f.id
          ${whereSql3}
          ORDER BY r.row_number ASC
          LIMIT 60;
        `;
        res = await pool.query(query3, queryParams3);
      }

      // Apply per-file/sheet cap: max 8 rows per file+sheet combo to ensure diversity
      const slotCounts = {};
      const cappedCandidates = [];
      for (const row of res.rows) {
        const key = `${row.file_id}|${row.sheet_name}`;
        slotCounts[key] = (slotCounts[key] || 0) + 1;
        if (slotCounts[key] <= 8) {
          cappedCandidates.push(row);
        }
        if (cappedCandidates.length >= 45) break;
      }
      candidates = cappedCandidates;
      console.log(`[Chat] Fast Tiered Search Keywords: [${finalKeywords.join(', ')}] | Raw: ${res.rows.length} | Capped: ${candidates.length} | Files: ${new Set(candidates.map(c=>c.file_id)).size}`);
    }


    // Format sources list from candidate rows — use the BEST (highest-relevance) row per file
    const sourcesMap = new Map();
    candidates.forEach(c => {
      if (!sourcesMap.has(c.file_id)) {
        sourcesMap.set(c.file_id, {
          fileId: c.file_id,
          filename: c.filename,
          sheetName: c.sheet_name,
          rowNumber: c.row_number,
          matchCount: 1,
          bestScore: c.relevance_score || 0
        });
      } else {
        const existing = sourcesMap.get(c.file_id);
        existing.matchCount += 1;
        // Update to higher-scored candidate if found
        if ((c.relevance_score || 0) > existing.bestScore) {
          existing.sheetName = c.sheet_name;
          existing.rowNumber = c.row_number;
          existing.bestScore = c.relevance_score || 0;
        }
      }
    });

    // If no candidate row matched (e.g. user asked general catalog question), use active catalog files
    if (sourcesMap.size === 0 && candidates.length === 0) {
      activeCatalogFiles.slice(0, 5).forEach(f => {
        sourcesMap.set(f.id, {
          fileId: f.id,
          filename: f.filename,
          sheetName: 'Sheet1',
          rowNumber: 1,
          matchCount: parseInt(f.total_rows, 10)
        });
      });
    }

    let sources = Array.from(sourcesMap.values()).slice(0, 5);

    let answerText = '';
    let methodUsed = 'local';

    // Helper to format JSON row_data into clean, human-readable text for AI context
    const cleanRowDataText = (rowData) => {
      if (!rowData || typeof rowData !== 'object') return String(rowData);
      return Object.entries(rowData)
        .filter(([k, v]) => v !== null && v !== undefined && String(v).trim() !== '')
        .map(([k, v]) => `${k.replace(/_/g, ' ')}: ${v}`)
        .join(' | ');
    };

    // Attempt Gemini API call if modelMode is not explicitly 'local' and API key is present
    const rawKeys = (modelMode !== 'local' && process.env.GEMINI_API_KEY) 
      ? process.env.GEMINI_API_KEY.split(',').map(k => k.trim()).filter(Boolean) 
      : [];
    if (rawKeys.length > 0) {
      const catalogSnippet = allFiles.map((f, idx) => {
        return `${idx + 1}. "${f.filename}" (ID: ${f.id}, Pengunggah: ${f.uploader_name}, Total Baris: ${f.total_rows}, Tanggal: ${new Date(f.uploaded_at).toLocaleDateString('id-ID')})`;
      }).join('\n');

      const contextSnippet = candidates.map((c, idx) => {
        return `[Sampel #${idx+1}: File "${c.filename}", Sheet "${c.sheet_name}", Baris #${c.row_number}] -> ${cleanRowDataText(c.row_data)}`;
      }).join('\n');

      const systemPrompt = `Anda adalah "Google Gemini AI Assistant", kecerdasan buatan serba bisa sekelas Google AI / ChatGPT.
Anda dapat menjawab PERTANYAAN APA SAJA (pengetahuan umum, sains, teknologi, matematika, analisis data, pembuatan draft surat, koding, atau pertanyaan umum) sekaligus memiliki akses Penuh ke seluruh katalog arsip dokumen spreadsheet di database aplikasi ini!

--- KATALOG SELURUH BERKAS DOKUMEN (${allFiles.length} File) ---
${catalogSnippet || 'Belum ada berkas.'}
---------------------------------------------------------

--- KONTEKS BARIS DATA RELEVAN DARI DATABASE ---
${candidates.length > 0 ? contextSnippet : 'Tidak ada sampel baris data spesifik yang relevan dengan pertanyaan ini.'}
-----------------------------------------------

Panduan Jawaban Sekalas Google Gemini:
1. Anda BISA dan BOLEH menjawab PERTANYAAN APA SAJA yang diajukan pengguna (pengetahuan umum, fakta dunia, kalkulasi, koding, rangkuman, surat, atau bantuan umum lainnya) tanpa membatasi diri hanya pada dokumen.
2. Jika pertanyaan pengguna berkaitan dengan dokumen spreadsheet di database, gunakan data katalog dan sampel baris di atas secara presisi. Sebutkan nama file dokumen dan sertakan rujukan tag [REF:NamaFile|BarisNumber] (contoh: [REF:ADATA CF KANPUS 20260710.xlsx|10660]).
3. Jika pertanyaan pengguna bersifat umum (seperti pengetahuan umum, matematika, sains, koding, dsb.), jawablah secara lengkap, ramah, dan cerdas seperti Google Search / Gemini AI tanpa memaksa mengaitkannya ke dokumen.
4. Jawablah secara ramah, cerdas, profesional, dan 100% akurat dalam Bahasa Indonesia menggunakan format Markdown yang rapi (teks tebal **, daftar berbutir, atau tabel jika relevan).`;

      const historyPayload = Array.isArray(req.body.history) ? req.body.history.slice(-4) : [];
      const contents = [
        ...historyPayload,
        { role: 'user', parts: [{ text: `${systemPrompt}\n\nPertanyaan Pengguna: "${queryText}"` }] }
      ];

      // Try keys sequentially with supported Google Gemini models (gemini-2.0-flash & gemini-1.5-flash)
      for (let i = 0; i < rawKeys.length; i++) {
        const apiKey = rawKeys[i];
        const modelsToTry = ['gemini-2.0-flash', 'gemini-1.5-flash'];

        for (const modelName of modelsToTry) {
          const geminiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;

          try {
            // 10-second strict timeout to prevent 504 / 502 / proxy timeouts
            const geminiRes = await fetch(geminiUrl, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ contents }),
              signal: AbortSignal.timeout(10000)
            });

            if (geminiRes.ok) {
              const geminiData = await geminiRes.json();
              const textOutput = geminiData.candidates?.[0]?.content?.parts?.[0]?.text;
              if (textOutput) {
                answerText = textOutput;
                methodUsed = 'gemini';
                console.log(`[Chatbot AI 🤖] Berhasil merespons menggunakan Gemini Model (${modelName}) Key #${i+1}/${rawKeys.length}`);

                // Extract explicit [REF:Filename|RowNumber] or natural text citations from Gemini response
                const citedMap = new Map();

                const findSheet = (fileId, rowNum) => {
                  const match = candidates.find(c => c.file_id === fileId && c.row_number === rowNum);
                  if (match) return match.sheet_name;
                  const any = candidates.find(c => c.file_id === fileId);
                  return any ? any.sheet_name : null;
                };

                // Pattern 1: [REF:Filename|RowNumber]
                const refRegex1 = /\[REF:([^|\]]+)\|(\d+)\]/gi;
                let m1;
                while ((m1 = refRegex1.exec(answerText)) !== null) {
                  const fn = m1[1].trim();
                  const rn = parseInt(m1[2], 10);
                  const fileObj = allFiles.find(f => f.filename.toLowerCase() === fn.toLowerCase());
                  if (fileObj && !citedMap.has(fileObj.id)) {
                    const sheet = findSheet(fileObj.id, rn);
                    citedMap.set(fileObj.id, { fileId: fileObj.id, filename: fileObj.filename, sheetName: sheet || 'Sheet1', rowNumber: rn, matchCount: 1 });
                  }
                }

                // Pattern 2: Natural Indonesian citation
                allFiles.forEach(f => {
                  if (!citedMap.has(f.id)) {
                    const escFn = f.filename.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                    const refRegex2 = new RegExp(`(?:${escFn}[^#\\n]*?baris\\s*#?\\s*(\\d+))|(?:baris\\s*#?\\s*(\\d+)[^\\n]*?${escFn})`, 'gi');
                    let m2;
                    while ((m2 = refRegex2.exec(answerText)) !== null) {
                      const rn = parseInt(m2[1] || m2[2], 10);
                      const sheet = findSheet(f.id, rn);
                      citedMap.set(f.id, { fileId: f.id, filename: f.filename, sheetName: sheet || 'Sheet1', rowNumber: rn, matchCount: 1 });
                    }
                  }
                });

                answerText = answerText.replace(/\[REF:[^\]]+\]/g, '').trim();

                if (citedMap.size > 0) {
                  sources = Array.from(citedMap.values());
                } else {
                  sources = [];
                }

                // Break outer key loop as well
                i = rawKeys.length;
                break;
              }
            } else {
              const errBody = await geminiRes.text();
              console.log(`[Chatbot AI ⚠️] Gemini Model (${modelName}) Key #${i+1} status:`, geminiRes.status, errBody.substring(0, 120));
            }
          } catch (geminiErr) {
            console.log(`[Chatbot AI ⚠️] Gemini Model (${modelName}) Key #${i+1} error (${geminiErr.name}):`, geminiErr.message);
          }
        }
      }
    }

    // Build sources list if citedMap wasn't populated and candidates exist
    if (sources.length === 0 && candidates.length > 0) {
      if (targetFileId) {
        const targetObj = allFiles.find(f => f.id === targetFileId);
        if (targetObj) {
          sources = [{
            fileId: targetObj.id,
            filename: targetObj.filename,
            sheetName: candidates[0]?.sheet_name || 'Sheet1',
            rowNumber: candidates[0]?.row_number || 1,
            matchCount: parseInt(targetObj.total_rows, 10)
          }];
        }
      } else {
        const fileMap = new Map();
        candidates.forEach(c => {
          if (!fileMap.has(c.file_id)) {
            fileMap.set(c.file_id, {
              fileId: c.file_id,
              filename: c.filename,
              sheetName: c.sheet_name,
              rowNumber: c.row_number,
              matchCount: 1
            });
          } else {
            fileMap.get(c.file_id).matchCount += 1;
          }
        });
        sources = Array.from(fileMap.values()).slice(0, 5);
      }
    }

    // Fallback if Gemini fails or Key not set — generate data-driven answer from candidates
    if (!answerText) {
      methodUsed = 'local';
      if (candidates.length > 0) {
        // Helper to format row_data into readable text
        const fmtRow = (rd) => {
          if (!rd || typeof rd !== 'object') return String(rd || '');
          return Object.entries(rd)
            .filter(([k, v]) => v !== null && v !== undefined && String(v).trim() !== '')
            .slice(0, 8)
            .map(([k, v]) => `${k.replace(/_/g, ' ')}: **${v}**`)
            .join(' | ');
        };

        // Build per-file grouped answer
        const fileGroups = new Map();
        candidates.forEach(c => {
          if (!fileGroups.has(c.file_id)) fileGroups.set(c.file_id, []);
          fileGroups.get(c.file_id).push(c);
        });

        const parts = [`Ditemukan **${candidates.length} baris data** yang relevan dengan kata kunci "**${queryText}**":\n`];
        let srcIdx = 0;
        for (const [fid, rows] of fileGroups) {
          // Sort rows by relevance_score DESC so the best matching row is listed at the top!
          rows.sort((a, b) => (b.relevance_score || 0) - (a.relevance_score || 0));
          const topRow = rows[0];
          parts.push(`\n📄 **${topRow.filename}** (Sheet: *${topRow.sheet_name}*):`);
          rows.slice(0, 5).forEach(r => {
            parts.push(`  - Baris #${r.row_number}: ${fmtRow(r.row_data)}`);
          });
          if (rows.length > 5) parts.push(`  - ...dan ${rows.length - 5} baris lainnya.`);
          srcIdx++;
        }

        // Ensure sources align exactly with the data groups shown
        sources = Array.from(fileGroups.entries()).slice(0, 5).map(([fid, rows]) => {
          const best = rows.reduce((a, b) => ((b.relevance_score || 0) > (a.relevance_score || 0) ? b : a), rows[0]);
          return {
            fileId: best.file_id,
            filename: best.filename,
            sheetName: best.sheet_name,
            rowNumber: best.row_number,
            matchCount: rows.length
          };
        });

        answerText = parts.join('\n');
      } else if (allFiles.length > 0) {
        answerText = `Saat ini terdapat **${allFiles.length} dokumen** yang terdaftar di dalam database:\n\n` +
          allFiles.slice(0, 10).map((f, i) => `${i + 1}. **${f.filename}** (${f.total_rows} baris, oleh *${f.uploader_name}*)`).join('\n') +
          (allFiles.length > 10 ? `\n...dan ${allFiles.length - 10} berkas lainnya.` : '');
        // No specific source for general catalog query
        sources = [];
      } else {
        answerText = `Belum ada dokumen yang diunggah. Silakan unggah berkas Excel melalui menu **Kelola File**.`;
        sources = [];
      }
    }


    // Log Chatbot activity
    await logActivity('chat_ai', `Mengirim pertanyaan ke AI Chatbot: "${queryText.slice(0, 60)}..."`, getDeviceInfo(req), req.user);

    const cleanKeywords = finalKeywords.join(' ');

    // Debug: log sources being sent to frontend
    console.log(`[Chat Debug] Query: "${queryText.slice(0, 50)}" | Candidates: ${candidates.length} | Sources (${sources.length}):`, 
      sources.map(s => `${s.filename}|sheet=${s.sheetName}|row=${s.rowNumber}`).join(', '));

    res.json({
      answer: answerText,
      sources,
      methodUsed,
      userQuery: cleanKeywords || queryText
    });
  } catch (err) {
    console.error('Error handling AI chat request:', err);
    res.status(500).json({ error: 'Gagal memproses percakapan AI: ' + err.message });
  }
});

// Global error handler for upload limit
// Serve static files from the React frontend build
app.use(express.static(path.join(__dirname, '../frontend/dist')));

// Fallback all non-API routes to React's index.html
app.get('*', (req, res, next) => {
  if (req.originalUrl.startsWith('/api')) {
    return next();
  }
  const indexPath = path.join(__dirname, '../frontend/dist/index.html');
  if (fs.existsSync(indexPath)) {
    res.sendFile(indexPath);
  } else {
    res.status(404).send('Aplikasi berjalan dalam mode API. Harap akses via port frontend (5173) atau jalankan "npm run build" pada direktori frontend.');
  }
});

// Global API 404 handler (Guarantees JSON response for all /api routes)
app.use('/api/*', (req, res) => {
  res.status(404).json({ error: `Endpoint API "${req.originalUrl}" tidak ditemukan.` });
});

// Global Error Handler for API routes
app.use((err, req, res, next) => {
  if (err instanceof multer.MulterError) {
    if (err.code === 'LIMIT_FILE_SIZE') {
      return res.status(400).json({ error: 'Ukuran file terlalu besar! Maksimal 250MB.' });
    }
  }
  if (req.originalUrl && req.originalUrl.startsWith('/api')) {
    return res.status(err.status || 400).json({ error: err.message || 'Terjadi kesalahan pada server.' });
  }
  return res.status(500).json({ error: err.message || 'Server error' });
});

// Start Express Server
app.listen(port, () => {
  console.log(`Server backend berjalan di http://localhost:${port}`);
});
