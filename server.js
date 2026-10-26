require('dotenv').config();

const express = require('express');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const cors = require('cors');
const { Pool } = require('pg');

const app = express();
const PORT = Number(process.env.PORT || 4000);


if (!process.env.DATABASE_URL) {
  console.error('ERROR: DATABASE_URL environment variable is missing.');
  process.exit(1);
}

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: { rejectUnauthorized: false },
  connectionTimeoutMillis: 15000,
  idleTimeoutMillis: 30000,
  max: 5
});

app.use(cors({
  origin: '*',
  methods: ['GET', 'POST', 'PUT', 'DELETE'],
  allowedHeaders: ['Content-Type', 'Authorization']
}));

app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

function getPeriod(period) {
  const value = String(period || 'SEP_2026')
    .trim()
    .replace(/\s+/g, '_')
    .toUpperCase();

  if (!/^[A-Z0-9_]{3,40}$/.test(value)) {
    throw new Error('Invalid period');
  }

  return value;
}

function mrKey(period) {
  return `mr:${getPeriod(period)}`;
}

const EMPTY_MR_DATA = {
  mrSheets: [],
  lastIndex: 0,
  do2Groups: []
};

const EMPTY_APPROVALS = {
  pendingUsers: [],
  approvedUsers: [],
  requests: []
};

async function initializeDatabase() {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS public.cpl_app_data (
      data_key TEXT PRIMARY KEY,
      data_value JSONB NOT NULL,
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    )
  `);

  await pool.query(`
    ALTER TABLE public.cpl_app_data
    ENABLE ROW LEVEL SECURITY
  `);

  await pool.query('SELECT 1');
  console.log('Supabase PostgreSQL connection successful.');
}

async function readData(key, fallback) {
  const result = await pool.query(
    `SELECT data_value
       FROM public.cpl_app_data
      WHERE data_key = $1`,
    [key]
  );

  return result.rows.length
    ? result.rows[0].data_value
    : fallback;
}

async function saveData(key, value) {
  await pool.query(
    `INSERT INTO public.cpl_app_data
       (data_key, data_value, updated_at)
     VALUES ($1, $2::jsonb, NOW())
     ON CONFLICT (data_key)
     DO UPDATE SET
       data_value = EXCLUDED.data_value,
       updated_at = NOW()`,
    [key, JSON.stringify(value)]
  );
}

// Google Drive JSON backup.
// Backup failure will not undo a successful database save.
async function backupToGoogleDrive(period, data) {
  const url = process.env.GOOGLE_BACKUP_URL;
  const secret = process.env.GOOGLE_BACKUP_SECRET;

  if (!url || !secret) {
    console.error('Google Drive backup is not configured.');
    return false;
  }

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        secret,
        period,
        data
      }),
      signal: AbortSignal.timeout(30000)
    });

    const result = await response.json();

    if (!response.ok || !result.success) {
      console.error(
        'Google Drive backup failed:',
        result.error || response.status
      );
      return false;
    }

    console.log(
      'Google Drive backup successful:',
      result.filename
    );

    return true;
  } catch (error) {
    console.error(
      'Google Drive backup error:',
      error.message
    );
    return false;
  }
}

function parseAnyDateToDDMonYYYY(value) {
  if (value === null || value === undefined || value === '') {
    return '';
  }

  const months = [
    'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
    'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'
  ];

  let dateValue;

  if (value instanceof Date) {
    dateValue = new Date(value.getTime() + 43200000);
  } else {
    const textValue = String(value).trim();

    if (
      textValue.includes('GMT') ||
      textValue.includes('Time') ||
      (isNaN(textValue) && !isNaN(Date.parse(textValue)))
    ) {
      dateValue = new Date(
        new Date(textValue).getTime() + 43200000
      );
    } else {
      const serial = parseFloat(textValue);

      if (!isNaN(serial) && serial > 10000 && serial < 80000) {
        dateValue = new Date(
          (Math.floor(serial - 25569) * 86400 * 1000) +
          43200000
        );
      }
    }

    if (!dateValue) return textValue;
  }

  if (isNaN(dateValue.getTime())) return '';

  return `${String(dateValue.getUTCDate()).padStart(2, '0')}-${months[dateValue.getUTCMonth()]}-${dateValue.getUTCFullYear()}`;
}

// Database health check.
app.get('/api/db-health', async (req, res) => {
  try {
    await pool.query('SELECT 1');

    res.json({
      success: true,
      database: 'connected'
    });
  } catch (error) {
    console.error(
      'Database health check failed:',
      error.message
    );

    res.status(500).json({
      success: false,
      database: 'error'
    });
  }
});

// Excel master list API.
app.get('/api/master-data', (req, res) => {
  try {
    const files = fs.readdirSync(__dirname).filter(
      file =>
        file.toLowerCase().endsWith('.xlsx') &&
        !file.startsWith('~$')
    );

    if (!files.length) {
      return res.status(404).json({
        error: 'Koi bhi .xlsx file project folder me nahi mili!'
      });
    }

    const fileName = files[0];

    const workbook = XLSX.readFile(
      path.join(__dirname, fileName),
      { cellDates: true }
    );

    const sheet = workbook.Sheets.MASTERLIST;

    if (!sheet) {
      return res.status(404).json({
        error: `Sheet MASTERLIST not found in ${fileName}!`
      });
    }

    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: ''
    });

    if (rows.length < 2) {
      return res.json({
        headers: [],
        data: []
      });
    }

    const headers = rows[1].slice(0, 24).map(
      value => String(value || '').trim()
    );

    const data = [];

    for (let i = 2; i < rows.length; i++) {
      const row = rows[i];
      const code = String(row[1] || '').trim();
      const name = String(row[3] || '').trim();

      if (!code && !name) continue;

      const rowData = Array.from({ length: 24 }, (_, col) => {
        const value = row[col] === undefined ? '' : row[col];

        return String(
          [4, 22, 23].includes(col)
            ? parseAnyDateToDDMonYYYY(value)
            : (value || '')
        ).trim();
      });

      data.push({
        code,
        name,
        rowData
      });
    }

    res.json({
      headers,
      data
    });
  } catch (error) {
    console.error('Master data error:', error);

    res.status(500).json({
      error: 'Master data read failed'
    });
  }
});

// Read period-wise MR sheets and DO-II data.
app.get('/api/mr-sheets', async (req, res) => {
  try {
    const key = mrKey(req.query.period);
    const data = await readData(key, EMPTY_MR_DATA);

    res.json(data);
  } catch (error) {
    console.error('MR read failed:', error.message);

    res.status(500).json({
      error: 'MR data read failed'
    });
  }
});

// Serialize writes for each period.
const writeQueues = new Map();

function queuePeriodWrite(period, callback) {
  const previous =
    writeQueues.get(period) || Promise.resolve();

  const current = previous.catch(() => {}).then(callback);

  writeQueues.set(period, current);

  current.finally(() => {
    if (writeQueues.get(period) === current) {
      writeQueues.delete(period);
    }
  }).catch(() => {});

  return current;
}

// Save MR sheets and DO-II data.
app.post('/api/mr-sheets', async (req, res) => {
  let period;

  try {
    period = getPeriod(req.body.period);
  } catch (error) {
    return res.status(400).json({
      error: error.message
    });
  }

  try {
    const result = await queuePeriodWrite(
      period,
      async () => {
        const key = mrKey(period);

        const existing = await readData(
          key,
          EMPTY_MR_DATA
        );

        const incoming = req.body.mrSheets;

        if (
          incoming !== undefined &&
          !Array.isArray(incoming)
        ) {
          throw new Error('mrSheets must be an array');
        }

        const oldSheets = Array.isArray(existing.mrSheets)
          ? existing.mrSheets
          : [];

        const nextSheets =
          incoming === undefined ? oldSheets : incoming;

        // Explicitly-authorized workflow actions. Normal saves cannot silently
        // delete submitted MRs or reopen them; the frontend must name the MR.
        const normalizeLabels = value => Array.isArray(value)
          ? new Set(value.map(v => String(v || '').trim().toUpperCase()).filter(Boolean))
          : new Set();
        const deleteLabels = normalizeLabels(req.body.deleteLabels);
        const approvedEditLabels = normalizeLabels(req.body.approvedEditLabels);

        const nextByLabel = new Map(
          nextSheets.map(sheet => [
            String(sheet.label || '').trim().toUpperCase(),
            sheet
          ])
        );

        if (oldSheets.length && incoming !== undefined && !nextSheets.length) {
          const allOldLabelsExplicitlyDeleted = oldSheets.every(sheet =>
            deleteLabels.has(String(sheet.label || '').trim().toUpperCase())
          );
          if (!allOldLabelsExplicitlyDeleted) {
            throw new Error('Refusing to replace existing MR records with an empty array');
          }
        }

        for (const oldSheet of oldSheets) {
          const label = String(oldSheet.label || '').trim().toUpperCase();
          const newSheet = nextByLabel.get(label);

          if (!newSheet) {
            if (deleteLabels.has(label)) continue;
            throw new Error(`Refusing to remove existing MR: ${label}`);
          }

          // Keep edit approval on the server until that MR is submitted again.
          // The approval request itself unblocks the sheet; later full-list saves
          // (including the final resubmission) may then replace its row count.
          const editApproved =
            oldSheet.editApproved === true ||
            approvedEditLabels.has(label) ||
            oldSheet.isSubmitted !== true;

          // Reducing rows or reopening a submitted MR is allowed only for the
          // exact MR approved for editing by the admin.
          if (
            (oldSheet.rows || []).length > (newSheet.rows || []).length &&
            !editApproved
          ) {
            throw new Error(`Refusing to reduce rows for ${label}`);
          }

          if (oldSheet.isSubmitted && !newSheet.isSubmitted && !editApproved) {
            throw new Error(`Refusing to unsubmit ${label}`);
          }

          // Only server state can carry this permission forward. Clear it once
          // the edited MR is submitted again, so future row reductions require
          // a fresh approval.
          newSheet.editApproved = editApproved && !newSheet.isSubmitted;
        }

        const payload = {
          mrSheets: nextSheets,
          lastIndex: req.body.lastIndex !== undefined
            ? req.body.lastIndex
            : (existing.lastIndex || 0),
          do2Groups: req.body.do2Groups !== undefined
            ? req.body.do2Groups
            : (existing.do2Groups || [])
        };

        if (!Array.isArray(payload.do2Groups)) {
          throw new Error('do2Groups must be an array');
        }

        await saveData(key, payload);

        // Verify the database save before making a backup.
        const saved = await readData(key, null);

        if (
          !saved ||
          !Array.isArray(saved.mrSheets)
        ) {
          throw new Error(
            'Database save verification failed'
          );
        }

        // Start Drive backup after the verified database save, but do not keep
        // the user's submit request waiting for the external backup service.
        // Backup failures are logged by backupToGoogleDrive and do not undo DB.
        backupToGoogleDrive(period, saved).catch(error => {
          console.error('Google Drive backup scheduling error:', error.message);
        });

        return {
          success: true,
          period,
          mrCount: saved.mrSheets.length,
          backupQueued: true
        };
      }
    );

    res.json(result);
  } catch (error) {
    console.error('MR save failed:', error.message);

    res.status(409).json({
      success: false,
      error: error.message
    });
  }
});

// Read approval data.
app.get('/api/approvals', async (req, res) => {
  try {
    const data = await readData(
      'approvals',
      EMPTY_APPROVALS
    );

    res.json(data);
  } catch (error) {
    console.error('Approval read failed:', error.message);

    res.status(500).json({
      error: 'Approval read failed'
    });
  }
});

// Approve an MR edit request in one database transaction so the sheet is
// reopened and its request is removed together, without multiple browser trips.
app.post('/api/approvals/approve-edit', async (req, res) => {
  let period;

  try {
    period = getPeriod(req.body.period);
  } catch (error) {
    return res.status(400).json({ error: error.message });
  }

  const label = String(req.body.mrLabel || '').trim();
  const user = String(req.body.user || '').trim().toUpperCase();
  if (!label || !user || label.toUpperCase() === 'LOGIN APPROVAL') {
    return res.status(400).json({ error: 'MR label and requesting user are required' });
  }

  let client;
  let transactionOpen = false;

  try {
    client = await pool.connect();
    await client.query('BEGIN');
    transactionOpen = true;

    const mrDataKey = mrKey(period);
    const records = await client.query(
      `SELECT data_key, data_value
         FROM public.cpl_app_data
        WHERE data_key = ANY($1::text[])
        FOR UPDATE`,
      [[mrDataKey, 'approvals']]
    );
    const recordByKey = new Map(records.rows.map(row => [row.data_key, row.data_value]));
    const mrData = recordByKey.get(mrDataKey) || EMPTY_MR_DATA;
    const approvals = recordByKey.get('approvals') || EMPTY_APPROVALS;
    const mrSheets = Array.isArray(mrData.mrSheets) ? mrData.mrSheets : [];
    const requests = Array.isArray(approvals.requests) ? approvals.requests : [];

    const requestIndex = requests.findIndex(request =>
      String(request.mrLabel || '').trim().toUpperCase() === label.toUpperCase() &&
      String(request.user || '').trim().toUpperCase() === user
    );
    if (requestIndex === -1) {
      await client.query('ROLLBACK');
      transactionOpen = false;
      return res.status(409).json({ error: 'Edit request is no longer pending; refresh the request list.' });
    }

    const sheetIndex = mrSheets.findIndex(sheet =>
      String(sheet.label || '').trim().toUpperCase() === label.toUpperCase()
    );
    if (sheetIndex === -1) {
      await client.query('ROLLBACK');
      transactionOpen = false;
      return res.status(404).json({ error: `Sheet not found: ${label}` });
    }

    const updatedSheets = mrSheets.map(sheet => ({ ...sheet }));
    updatedSheets[sheetIndex] = {
      ...updatedSheets[sheetIndex],
      isSubmitted: false,
      assignedUser: user,
      editApproved: true
    };
    const updatedMrData = { ...mrData, mrSheets: updatedSheets };
    const updatedApprovals = {
      pendingUsers: Array.isArray(approvals.pendingUsers) ? approvals.pendingUsers : [],
      approvedUsers: Array.isArray(approvals.approvedUsers) ? approvals.approvedUsers : [],
      requests: requests.filter((_, index) => index !== requestIndex)
    };

    await client.query(
      `INSERT INTO public.cpl_app_data (data_key, data_value, updated_at)
       VALUES ($1, $2::jsonb, NOW())
       ON CONFLICT (data_key) DO UPDATE SET
         data_value = EXCLUDED.data_value,
         updated_at = EXCLUDED.updated_at`,
      [mrDataKey, JSON.stringify(updatedMrData)]
    );
    await client.query(
      `INSERT INTO public.cpl_app_data (data_key, data_value, updated_at)
       VALUES ('approvals', $1::jsonb, NOW())
       ON CONFLICT (data_key) DO UPDATE SET
         data_value = EXCLUDED.data_value,
         updated_at = EXCLUDED.updated_at`,
      [JSON.stringify(updatedApprovals)]
    );

    await client.query('COMMIT');
    transactionOpen = false;

    // The database change is committed; Drive backup runs in the background.
    backupToGoogleDrive(period, updatedMrData).catch(error => {
      console.error('Google Drive backup scheduling error:', error.message);
    });

    res.json({
      success: true,
      mrLabel: updatedSheets[sheetIndex].label,
      assignedUser: user,
      sheet: updatedSheets[sheetIndex],
      requests: updatedApprovals.requests
    });
  } catch (error) {
    if (transactionOpen) {
      await client.query('ROLLBACK').catch(() => {});
    }
    console.error('Edit request approval failed:', error.message);
    res.status(500).json({ error: 'Edit request approval failed; no partial update was kept.' });
  } finally {
    if (client) client.release();
  }
});

// Save approval data.
app.post('/api/approvals', async (req, res) => {
  try {
    const {
      pendingUsers,
      approvedUsers,
      requests
    } = req.body;

    if (
      !Array.isArray(pendingUsers) ||
      !Array.isArray(approvedUsers) ||
      !Array.isArray(requests)
    ) {
      return res.status(400).json({
        error: 'Invalid approvals payload'
      });
    }

    await saveData('approvals', {
      pendingUsers,
      approvedUsers,
      requests
    });

    res.json({
      success: true
    });
  } catch (error) {
    console.error('Approval save failed:', error.message);

    res.status(500).json({
      error: 'Approval save failed'
    });
  }
});

// Start only after confirming database availability.
initializeDatabase()
  .then(() => {
    app.listen(PORT, '0.0.0.0', () => {
      console.log(
        `CPL server listening on port ${PORT}`
      );

      console.log(
        'MR sheets, DO-II and approvals use Supabase PostgreSQL.'
      );
    });
  })
  .catch(error => {
    console.error(
      'Supabase initialization failed:',
      error.message
    );

    process.exit(1);
  });

process.on('SIGTERM', async () => {
  await pool.end();
  process.exit(0);
});

