const express = require('express');
const cors = require('cors');
const multer = require('multer');
const ExcelJS = require('exceljs');
const { Pool } = require('pg');
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
  user: process.env.DB_USER || 'postgres',
  host: process.env.DB_HOST || '127.0.0.1',
  database: process.env.DB_NAME || 'excel_db',
  password: process.env.DB_PASSWORD || 'mysecretpassword',
  port: process.env.DB_PORT || '5432',
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

    // Migrate: add full_name column if not exists
    await client.query(`
      ALTER TABLE users ADD COLUMN IF NOT EXISTS full_name VARCHAR(100);
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

    // Set default full names for seeded users if null
    await client.query(`UPDATE users SET full_name = 'Administrator System' WHERE username = 'admin' AND full_name IS NULL;`);
    await client.query(`UPDATE users SET full_name = 'Operator Staff' WHERE username = 'operator' AND full_name IS NULL;`);
    await client.query(`UPDATE users SET full_name = 'Viewer Guest' WHERE username = 'viewer' AND full_name IS NULL;`);

    console.log('Inisialisasi database PostgreSQL sukses.');
  } catch (err) {
    console.error('Gagal melakukan inisialisasi database:', err.message);
  } finally {
    client.release();
  }
};

// Call DB Init
initDatabase();

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
    const usersRes = await pool.query('SELECT id, username, full_name, role, created_at FROM users ORDER BY created_at DESC;');
    res.json(usersRes.rows);
  } catch (err) {
    console.error('Error fetching users:', err);
    res.status(500).json({ error: err.message });
  }
});

// User Management API: Create new user (Admin Only)
app.post('/api/users', authenticateToken, requireRole(['admin']), async (req, res) => {
  const { username, password, role, full_name } = req.body;
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
      'INSERT INTO users (username, password_hash, role, full_name) VALUES ($1, $2, $3, $4) RETURNING id, username, full_name, role, created_at;',
      [username, passwordHash, role, full_name || null]
    );

    const newUser = result.rows[0];
    logActivity('create_user', `Membuat akun baru: "${newUser.username}" dengan hak akses ${newUser.role}`, getDeviceInfo(req), req.user);
    res.status(201).json(newUser);
  } catch (err) {
    console.error('Error creating user:', err);
    res.status(500).json({ error: err.message });
  }
});

// User Management API: Update user (Admin Only)
app.put('/api/users/:id', authenticateToken, requireRole(['admin']), async (req, res) => {
  const { id } = req.params;
  const { username, password, role, full_name } = req.body;

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
         SET username = $1, password_hash = $2, role = $3, full_name = $4 
         WHERE id = $5 
         RETURNING id, username, full_name, role, created_at;`,
        [username.trim(), passwordHash, role, full_name || null, id]
      );
    } else {
      result = await pool.query(
        `UPDATE users 
         SET username = $1, role = $2, full_name = $3 
         WHERE id = $4 
         RETURNING id, username, full_name, role, created_at;`,
        [username.trim(), role, full_name || null, id]
      );
    }

    const updatedUser = result.rows[0];
    logActivity('update_user', `Memperbarui akun: "${updatedUser.username}" (${updatedUser.role})`, getDeviceInfo(req), req.user);
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
  const queryText = req.query.q;
  const fileId = req.query.fileId;
  const filterSheet = req.query.sheet; // Optional sheet filter
  const filterUnit = req.query.unit;   // Optional unit filter
  const uploaderId = req.query.uploaderId; // Optional uploader ID filter

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
      if (uploaderId && uploaderId !== 'all' && req.user.role === 'admin') {
        queryParams.push(parseInt(uploaderId));
        filterClauses.push(`f.uploaded_by = $${queryParams.length}`);
      }

      queryParams.push(req.user.id);
      const userIdParamIndex = queryParams.length;

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
        LEFT JOIN bookmarks b ON b.row_id = r.id AND b.user_id = $${userIdParamIndex}
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
        if (uploaderId && uploaderId !== 'all' && req.user.role === 'admin') {
          queryParams.push(parseInt(uploaderId));
          filterClauses.push(`f.uploaded_by = $${queryParams.length}`);
        }

        const filterSql = filterClauses.length > 0 ? ' AND ' + filterClauses.join(' AND ') : '';

        queryParams.push(req.user.id);
        const userIdParamIndex = queryParams.length;

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
          LEFT JOIN bookmarks b ON b.row_id = r.id AND b.user_id = $${userIdParamIndex}
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
        if (uploaderId && uploaderId !== 'all' && req.user.role === 'admin') {
          queryParams.push(parseInt(uploaderId));
          whereClauses.push(`f.uploaded_by = $${queryParams.length}`);
        }

        queryParams.push(req.user.id);
        const userIdParamIndex = queryParams.length;

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
          LEFT JOIN bookmarks b ON b.row_id = r.id AND b.user_id = $${userIdParamIndex}
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
      searchRes = await pool.query(fallbackQuery, fallbackParams);
    }
    
    // Log search activity or file preview
    if (!fileId && queryText && queryText.trim() !== '') {
      logActivity('search', `Mencari kata kunci: "${queryText.trim()}" (Menghasilkan ${searchRes.rows.length} baris data)`, getDeviceInfo(req), req.user);
    } else if (fileId) {
      const filename = searchRes.rows[0] ? searchRes.rows[0].filename : `ID #${fileId}`;
      logActivity('view_file', `Membuka/melihat berkas: "${filename}"`, getDeviceInfo(req), req.user);
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

// 5b. Get activity logs (with search + pagination)
app.get('/api/logs', authenticateToken, requireRole(['admin']), async (req, res) => {
  try {
    const page  = Math.max(1, parseInt(req.query.page)  || 1);
    const limit = Math.min(500, Math.max(10, parseInt(req.query.limit) || 50));
    const q     = (req.query.q || '').trim();
    const offset = (page - 1) * limit;

    const conditions = q
      ? `WHERE (activity_details ILIKE $1 OR activity_type ILIKE $1 OR ip_address ILIKE $1 OR browser ILIKE $1 OR os ILIKE $1 OR device_label ILIKE $1 OR username ILIKE $1)`
      : '';
    const params = q ? [`%${q}%`] : [];

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

// Start Express Server
app.listen(port, () => {
  console.log(`Server backend berjalan di http://localhost:${port}`);
});
