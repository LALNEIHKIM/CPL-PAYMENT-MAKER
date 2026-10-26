const express = require('express');

const path = require('path');

const fs = require('fs');

const xlsx = require('xlsx');

const cors = require('cors');



const app = express();

app.use(cors());

app.use(express.json({ limit: '25mb' }));

app.use(express.static(path.join(__dirname, 'public')));



const storeDir = path.join(__dirname, 'store-data');

const usersFile = path.join(storeDir, 'users.json');

const submissionsFile = path.join(storeDir, 'submissions.json');

const requestsFile = path.join(storeDir, 'edit-requests.json');

const approvedEditsFile = path.join(storeDir, 'approved-edits.json');

fs.mkdirSync(storeDir, { recursive: true });

const databaseUrl = String(process.env.DATABASE_URL || '').trim();
let databasePool = null;
let databaseReady = false;
const databaseStoreCache = new Map();
const pendingDatabaseWrites = new Map();

function storageKey(file) {
  return path.relative(storeDir, file).split(path.sep).join('/');
}

function readStore(file, fallback = []) {
  if (databaseReady) return databaseStoreCache.has(storageKey(file)) ? databaseStoreCache.get(storageKey(file)) : fallback;
  try { return JSON.parse(fs.readFileSync(file, 'utf8')); } catch { return fallback; }
}

function writeStore(file, value) {
  if (!databaseReady) {
    fs.writeFileSync(file, JSON.stringify(value, null, 2));
    return;
  }
  const key = storageKey(file);
  const snapshot = JSON.parse(JSON.stringify(value));
  databaseStoreCache.set(key, snapshot);
  const previous = pendingDatabaseWrites.get(key) || Promise.resolve();
  const task = previous.catch(() => {}).then(() => databasePool.query(
    'INSERT INTO app_private.app_store (key, value, updated_at) VALUES ($1, $2::jsonb, NOW()) ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value, updated_at = NOW()',
    [key, JSON.stringify(snapshot)]
  ));
  pendingDatabaseWrites.set(key, task);
  task.finally(() => { if (pendingDatabaseWrites.get(key) === task) pendingDatabaseWrites.delete(key); }).catch(() => {});
}

async function flushDatabaseWrites() {
  while (pendingDatabaseWrites.size) await Promise.all([...pendingDatabaseWrites.values()]);
}

async function initializeDatabaseStore() {
  if (!databaseUrl) return;
  const { Pool } = require('pg');
  databasePool = new Pool({ connectionString: databaseUrl, max: 5 });
  await databasePool.query('SELECT 1');
  await databasePool.query('CREATE SCHEMA IF NOT EXISTS app_private');
  await databasePool.query('CREATE TABLE IF NOT EXISTS app_private.app_store (key TEXT PRIMARY KEY, value JSONB NOT NULL, updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW())');
  let result = await databasePool.query('SELECT key, value FROM app_private.app_store');
  if (result.rowCount === 0) {
    const legacyFiles = [usersFile, submissionsFile, requestsFile, approvedEditsFile];
    if (fs.existsSync(mrJobsDir)) {
      for (const name of fs.readdirSync(mrJobsDir)) if (name.toLowerCase().endsWith('.json')) legacyFiles.push(path.join(mrJobsDir, name));
    }
    for (const file of legacyFiles) {
      if (!fs.existsSync(file)) continue;
      try {
        const value = JSON.parse(fs.readFileSync(file, 'utf8'));
        await databasePool.query('INSERT INTO app_private.app_store (key, value) VALUES ($1, $2::jsonb) ON CONFLICT (key) DO NOTHING', [storageKey(file), JSON.stringify(value)]);
      } catch (error) { console.warn('Legacy JSON import skipped:', path.basename(file), error.message); }
    }
    result = await databasePool.query('SELECT key, value FROM app_private.app_store');
  }
  for (const row of result.rows) databaseStoreCache.set(row.key, row.value);
  databaseReady = true;
  console.log(`Supabase storage ready (${databaseStoreCache.size} records).`);
}

// Successful JSON responses wait until all changes made by the request are durable in Postgres.
app.use((req, res, next) => {
  const sendJson = res.json.bind(res);
  res.json = function (payload) {
    if (!databaseReady || pendingDatabaseWrites.size === 0) return sendJson(payload);
    flushDatabaseWrites().then(() => sendJson(payload)).catch(error => {
      console.error('Supabase write failed; request not acknowledged:', error.message);
      res.statusCode = 503;
      res.setHeader('Content-Type', 'application/json; charset=utf-8');
      sendJson({ error: 'Database save failed. Please retry; local draft remains on your browser.' });
    });
    return res;
  };
  next();
});

function isAdmin(req) { return req.get('x-cpl-admin') === 'admin'; }

function requestUser(req) { return String(req.get('x-cpl-user') || '').trim(); }

function isMRJobsAdmin(req) { return String(req.get('x-cpl-admin-password') || '') === String(process.env.CPL_ADMIN_PASSWORD || 'admin'); }

app.post('/api/admin/verify', (req, res) => {

  const password = String(req.body.password || '');

  if (password !== String(process.env.CPL_ADMIN_PASSWORD || 'admin')) return res.status(401).json({ error: 'Incorrect admin password' });

  res.json({ ok: true });

});





// MR / Job master list: stored on server separately for each wages period.

const mrJobsDir = path.join(storeDir, 'mr-jobs');

fs.mkdirSync(mrJobsDir, { recursive: true });

const defaultMRJobs = [{"mrNo": "MR NO. 374", "jobNo": "ADG (B)/BFR Cell/A/A & E/S/2024/832 FMN"}, {"mrNo": "MR NO. 375", "jobNo": "ADG (B)/BFR Cell/A/A & E/S/2024/831 FMN"}, {"mrNo": "MR NO. 9", "jobNo": "Maint CHK Road (2026-27) Km 0.00 to 10.00"}, {"mrNo": "MR NO. 380", "jobNo": "80/515 (PSK)"}, {"mrNo": "MR NO. 376", "jobNo": "Maint CHK Road (2026-27) Km 10.00 to 75.50"}, {"mrNo": "MR NO. 41", "jobNo": "710-A/50 Rehab CHK road (PMT)"}, {"mrNo": "MR NO. 377", "jobNo": "710-A/52 (24 TF)"}, {"mrNo": "MR NO. 378", "jobNo": "Maint K-V-K Road (2026-27)"}, {"mrNo": "MR NO. 8", "jobNo": "Maint DVF (2026-27) KM 0.00 to 35.20"}, {"mrNo": "MR NO. 379", "jobNo": "Maint FT KM 0.00 to 10.775 (2026-27)"}, {"mrNo": "MR NO. 381", "jobNo": "Maint NVT Road (2026-27) KM 0-45.358"}, {"mrNo": "MR NO. 27", "jobNo": "L-3307/07 TBSPS road (Pmt)"}, {"mrNo": "MR NO. 28", "jobNo": "L-3307/07 TBSPS road (Surf)"}, {"mrNo": "MR NO. 29", "jobNo": "L-3307/10 TBSPS road (Fmn)"}, {"mrNo": "MR NO. 30", "jobNo": "L-3307/10 TBSPS road (PMT)"}, {"mrNo": "MR NO. 31", "jobNo": "L-3307/12 (DPR)"}, {"mrNo": "MR NO. 32", "jobNo": "Maint S-T Road (2026-27) KM 82.00-164.53"}, {"mrNo": "MR NO. 33", "jobNo": "709/157 (Fmn)"}, {"mrNo": "MR NO. 34", "jobNo": "709/157 (PMT)"}, {"mrNo": "MR NO. 35", "jobNo": "Maint K-D-A Road (2026-27) KM 31.12-42.00"}, {"mrNo": "MR NO. 36", "jobNo": "Maint S-T Road (2026-27)"}, {"mrNo": "MR NO. 37", "jobNo": "ADG(B)/BFR Cell/A/A & E/S/2024/832 FMN"}, {"mrNo": "MR NO. 38", "jobNo": "ADG(B)/BFR Cell/A/A & E/S/2024/832 SURF"}, {"mrNo": "MR NO. 39", "jobNo": "ADG(B)/BFR Cell/A/A & E/S/2024/831 FMN"}, {"mrNo": "MR NO. 40", "jobNo": "ADG(B)/BFR Cell/A/A & E/S/2024/831 SURF"}];

function mrJobsFile(period) {

  const safe = String(period || '').replace(/[^a-zA-Z0-9\_-]/g, '_').slice(0,80);

  return path.join(mrJobsDir, (safe || 'default') + '.json');

}

function readMRJobs(period) {

  const file = mrJobsFile(period);

  if (!databaseReady && !fs.existsSync(file)) return [];

  return readStore(file, []);

}

app.get('/api/mr-jobs/:period', (req, res) => {

  const period = decodeURIComponent(req.params.period || '').trim();

  if (!period) return res.status(400).json({ error: 'Period is required' });

  res.json({ period, entries: readMRJobs(period) });

});

app.post('/api/mr-jobs/:period', (req, res) => {

  if (!isMRJobsAdmin(req)) return res.status(403).json({ error: 'Admin password verification required' });

  const period = decodeURIComponent(req.params.period || '').trim();

  const mrNo = String(req.body.mrNo || '').trim().toUpperCase();

  const jobNo = String(req.body.jobNo || '').trim();

  if (!period || !mrNo || !jobNo) return res.status(400).json({ error: 'Period, MR number and Job number/description are required' });

  const entries = readMRJobs(period);

  const idx = entries.findIndex(x => String(x.mrNo).toLowerCase() === mrNo.toLowerCase());

  const item = { mrNo, jobNo };

  if (idx >= 0) entries[idx] = item; else entries.push(item);

  writeStore(mrJobsFile(period), entries);

  res.json({ ok: true, period, entries });

});



// Permanently delete one MR / Job entry for the selected period.

// Requires the same admin password header used by the save endpoint.

app.delete('/api/mr-jobs/:period/:mrNo', (req, res) => {

  if (!isMRJobsAdmin(req)) {

    return res.status(403).json({ error: 'Admin password verification required' });

  }



  const period = String(req.params.period || '').trim();

  const mrNo = decodeURIComponent(String(req.params.mrNo || '')).trim();



  if (!period || !mrNo) {

    return res.status(400).json({ error: 'Period and MR number are required' });

  }



  try {

    const entries = readMRJobs(period);

    const remaining = entries.filter(

      entry => String(entry.mrNo || '').trim().toLowerCase() !== mrNo.toLowerCase()

    );



    if (remaining.length === entries.length) {

      return res.status(404).json({ error: 'MR entry not found for this period' });

    }



    writeStore(mrJobsFile(period), remaining);

    return res.json({ ok: true, period, deletedMR: mrNo, entries: remaining });

  } catch (error) {

    console.error('MR / Job delete failed:', error);

    return res.status(500).json({ error: 'Failed to delete MR entry' });

  }

});



app.post('/api/register', (req, res) => {

  const name = String(req.body.name || '').trim();

  if (!name) return res.status(400).json({ error: 'Name is required' });

  const users = readStore(usersFile);

  if (users.some(u => u.name.toLowerCase() === name.toLowerCase())) return res.status(409).json({ error: 'Name already submitted. Please contact admin for approval.' });

  users.push({ name, approved: false, createdAt: new Date().toISOString() });

  writeStore(usersFile, users);

  res.json({ message: 'Please contact admin for approval.' });

});

app.get('/api/users', (req, res) => isAdmin(req) ? res.json(readStore(usersFile)) : res.status(403).json({ error: 'Admin only' }));

app.post('/api/users/:name/approve', (req, res) => {

  if (!isAdmin(req)) return res.status(403).json({ error: 'Admin only' });

  const users = readStore(usersFile); const user = users.find(u => u.name.toLowerCase() === decodeURIComponent(req.params.name).toLowerCase());

  if (!user) return res.status(404).json({ error: 'User not found' }); user.approved = true; writeStore(usersFile, users); res.json({ ok: true });

});

app.post('/api/login', (req, res) => {

  const name = String(req.body.name || '').trim();

  const user = readStore(usersFile).find(u => u.name.toLowerCase() === name.toLowerCase() && u.approved);

  if (!user) return res.status(403).json({ error: 'Please contact admin for approval.' });

  res.json({ ok: true, name: user.name });

});

app.post('/api/submit', (req, res) => {

  const name = String(req.body.name || '').trim();

  const ownerKey = (name || 'ADMIN').toLowerCase();

  const user = readStore(usersFile).find(u => u.name.toLowerCase() === name.toLowerCase() && u.approved);

  if (!user && !isAdmin(req)) return res.status(403).json({ error: 'Please contact admin for approval.' });

  const { period, data } = req.body;

  if (!period || !data || !Array.isArray(data.mrSheets)) return res.status(400).json({ error: 'Invalid store data' });

  const submissions = readStore(submissionsFile);

  const norm = value => String(value || '').trim().toLowerCase();

  const requestedLabels = data.mrSheets.map(sheet => norm(sheet.label)).filter(Boolean);

  const index = submissions.findIndex(sub =>
    norm(sub.period) === norm(period) && norm(sub.owner || 'ADMIN') === ownerKey
  );

  // A user can re-submit their own approved/edited MR to update the old stored copy.
  // Only another owner's MR with the same period and label is a true duplicate.
  const duplicate = submissions.find((sub, subIndex) =>
    subIndex !== index &&
    norm(sub.period) === norm(period) &&
    norm(sub.owner || 'ADMIN') !== ownerKey &&
    (sub.data?.mrSheets || []).some(sheet => requestedLabels.includes(norm(sheet.label)))
  );

  if (duplicate) {
    const duplicateLabel = (duplicate.data?.mrSheets || []).find(sheet => requestedLabels.includes(norm(sheet.label)))?.label;
    return res.status(409).json({ error: `Duplicate submission blocked: ${duplicateLabel} is already stored for ${period} by another account.` });
  }

  let mergedSheets = data.mrSheets;

  if (index >= 0) {

    const previousSheets = submissions[index].data?.mrSheets || [];

    const byLabel = new Map(previousSheets.map(sheet => [norm(sheet.label), sheet]));

    for (const sheet of data.mrSheets) byLabel.set(norm(sheet.label), sheet);

    mergedSheets = Array.from(byLabel.values());

  }

  const record = { period, owner: name || 'ADMIN', data: { ...data, mrSheets: mergedSheets }, submittedAt: new Date().toISOString() };

  if (index >= 0) submissions[index] = record; else submissions.push(record);

  // Once the edited MR is submitted, remove its returned-for-edit copy so it
  // becomes visible in Store Data again and is not re-added to Payment later.
  const requestedSet = new Set(requestedLabels);
  const approvedEdits = readStore(approvedEditsFile).map(item => {
    if (norm(item.period) !== norm(period) || norm(item.owner || 'ADMIN') !== ownerKey) return item;
    return { ...item, mrSheets: (item.mrSheets || []).filter(sheet => !requestedSet.has(norm(sheet.label))) };
  }).filter(item => (item.mrSheets || []).length > 0);

  writeStore(submissionsFile, submissions);
  writeStore(approvedEditsFile, approvedEdits);
  res.json({ ok: true });

});

// List all submitted periods available to this account for the ADDING SYSTEM picker.

app.get('/api/submitted-periods', (req, res) => {

  const user = requestUser(req);

  let submissions = readStore(submissionsFile);

  if (!isAdmin(req)) submissions = submissions.filter(s => user && String(s.owner || '').toLowerCase() === user.toLowerCase());

  const byPeriod = new Map();

  for (const sub of submissions) {

    const period = String(sub.period || '').trim();

    if (!period) continue;

    const mrCount = (sub.data?.mrSheets || []).length;

    if (!mrCount) continue;

    const existing = byPeriod.get(period) || { period, mrCount: 0 };

    existing.mrCount += mrCount;

    byPeriod.set(period, existing);

  }

  const months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];

  const periods = Array.from(byPeriod.values()).sort((a,b) => {

    const parse = p => { const m = String(p.period).match(/^([A-Z]{3})\s+(\d{4})$/i); return m ? Number(m[2])*12 + months.indexOf(m[1].toUpperCase()) : -1; };

    return parse(b) - parse(a);

  });

  res.json({ periods });

});



app.get('/api/store-data/:period', (req, res) => {
  const period = decodeURIComponent(req.params.period || '').trim();
  const norm = value => String(value || '').trim().toLowerCase();

  try {
    const allSubmissions = readStore(submissionsFile);
    const periodSubmissions = allSubmissions.filter(
      s => norm(s.period) === norm(period)
    );
    const user = requestUser(req);
    const approvedEdits = readStore(approvedEditsFile);

    // Hide an approved MR only if it was actually moved into approved-edits.json.
    // An approved request alone must never make a still-stored MR disappear.
    const actuallyMovedKeys = new Set();
    for (const approved of approvedEdits) {
      if (norm(approved.period) !== norm(period)) continue;
      for (const mr of (approved.mrSheets || [])) {
        actuallyMovedKeys.add(`${norm(approved.period)}::${norm(mr.label)}`);
      }
    }

    let visible = periodSubmissions.map(sub => ({
      ...sub,
      data: {
        ...(sub.data || {}),
        mrSheets: (sub.data?.mrSheets || []).filter(mr =>
          !actuallyMovedKeys.has(`${norm(period)}::${norm(mr.label)}`)
        )
      }
    })).filter(sub => (sub.data.mrSheets || []).length > 0);

    // Admin sees all remaining MR; regular users see only their own.
    if (!isAdmin(req)) {
      visible = visible.filter(
        sub => user && norm(sub.owner) === norm(user)
      );
    }

    const data = visible.flatMap(sub => (sub.data.mrSheets || []).map(mr => ({
      ...mr,
      submittedBy: String(sub.owner || 'ADMIN')
    })));
    console.log('[STORE DATA DEBUG]', {
      requestedPeriod: period,
      submissionsFile,
      totalSubmissions: allSubmissions.length,
      matchingPeriodSubmissions: periodSubmissions.length,
      matchingOwners: periodSubmissions.map(sub => sub.owner || 'ADMIN'),
      matchingMRLabels: periodSubmissions.flatMap(sub =>
        (sub.data?.mrSheets || []).map(mr => mr.label)
      ),
      admin: isAdmin(req),
      user: user || null,
      approvedEditRecords: approvedEdits.length,
      returnedMRCount: data.length
    });

    res.json({ period, data, found: data.length > 0 });
  } catch (error) {
    console.error('[STORE DATA ERROR]', error);
    res.status(500).json({ error: 'Failed to load Store Data' });
  }
});

// Approved edit grants are visible only to the matching requester.

// The payment screen can use this endpoint to move the approved MR into Screen 3.

app.get('/api/approved-edits/:period', (req, res) => {

  const period = decodeURIComponent(req.params.period || '');

  const user = requestUser(req);

  if (!user && !isAdmin(req)) return res.status(400).json({ error: 'User identity required' });

  const all = readStore(approvedEditsFile);

  const matched = all.filter(x => String(x.period) === period &&

    String(x.owner || '').toLowerCase() === String(user || 'ADMIN').toLowerCase());

  res.json({ period, data: matched.flatMap(x => x.mrSheets || []) });

});



app.get('/api/my-edit-approvals', (req, res) => {

  const user = requestUser(req);

  if (!user) return res.status(400).json({ error: 'User identity required' });

  const approvals = readStore(requestsFile)

    .filter(r => String(r.name || '').toLowerCase() === user.toLowerCase() && r.status === 'approved')

    .map(r => ({ id: r.id, name: r.name, period: r.period, mrLabel: r.mrLabel, approvedAt: r.approvedAt }));

  res.json({ approvals });

});

app.post('/api/edit-requests', (req, res) => {

  const name = String(req.body.name || '').trim(); const { period, mrLabel, message } = req.body;

  if (!name || !message) return res.status(400).json({ error: 'Name and message are required' });

  const requests = readStore(requestsFile); requests.push({ id: Date.now().toString(), name, period, mrLabel, message, status: 'pending', createdAt: new Date().toISOString() });

  writeStore(requestsFile, requests); res.json({ ok: true });

});

app.get('/api/edit-requests', (req, res) => isAdmin(req) ? res.json(readStore(requestsFile)) : res.status(403).json({ error: 'Admin only' }));

app.post('/api/edit-requests/:id/deny', (req, res) => {
  if (!isAdmin(req)) return res.status(403).json({ error: 'Admin only' });
  const requests = readStore(requestsFile);
  const item = requests.find(r => String(r.id) === String(req.params.id));
  if (!item) return res.status(404).json({ error: 'Request not found' });
  if (item.status !== 'pending') return res.status(409).json({ error: 'Only pending requests can be denied.' });
  item.status = 'denied';
  item.deniedAt = new Date().toISOString();
  item.deniedBy = req.user?.name || req.user?.username || 'ADMIN';
  writeStore(requestsFile, requests);
  res.json({ ok: true, denied: true, id: item.id, mrLabel: item.mrLabel });
});

app.post('/api/edit-requests/:id/approve', (req, res) => {

  if (!isAdmin(req)) return res.status(403).json({ error: 'Admin only' });

  const requests = readStore(requestsFile);

  const item = requests.find(r => String(r.id) === String(req.params.id));

  if (!item) return res.status(404).json({ error: 'Request not found' });

  if (item.status === 'approved') return res.json({ ok: true, alreadyApproved: true });



  const submissions = readStore(submissionsFile);

  const approvedEdits = readStore(approvedEditsFile);

  let movedSheets = [];

  const remaining = submissions.map(sub => {

    if (String(sub.period) !== String(item.period) ||

        String(sub.owner || '').toLowerCase() !== String(item.name || '').toLowerCase()) return sub;

    const keep = [];

    for (const mr of (sub.data?.mrSheets || [])) {

      if (String(mr.label || '').trim().toLowerCase() === String(item.mrLabel || '').trim().toLowerCase()) {

        movedSheets.push(mr);

      } else keep.push(mr);

    }

    return { ...sub, data: { ...sub.data, mrSheets: keep } };

  }).filter(sub => (sub.data?.mrSheets || []).length > 0);



  // Never mark a request approved unless its exact MR was found and moved.

  // This prevents a false approval from hiding an MR while leaving it nowhere to edit.

  if (!movedSheets.length) {

    return res.status(404).json({

      error: 'Matching MR was not found in Store Data. No data was changed.',

      movedToPayment: false,

      requester: item.name,

      period: item.period,

      mrLabel: item.mrLabel

    });

  }



  // Make approval idempotent: do not add the same request twice.

  const existingApproved = approvedEdits.some(x => String(x.id) === String(item.id));

  if (!existingApproved) {

    approvedEdits.push({

      id: String(item.id),

      owner: item.name,

      period: item.period,

      mrSheets: movedSheets,

      approvedAt: new Date().toISOString(),

      claimed: false

    });

  }



  // Persist the moved MR, then update the request status.

  writeStore(approvedEditsFile, approvedEdits);

  writeStore(submissionsFile, remaining);

  item.status = 'approved';

  item.approvedAt = new Date().toISOString();

  item.movedToPayment = true;

  writeStore(requestsFile, requests);

  res.json({ ok: true, movedToPayment: true, count: movedSheets.length, requester: item.name, period: item.period, mrLabel: item.mrLabel });

});



app.get('/', (req, res) => {

  res.sendFile(path.join(__dirname, 'public', 'login.html'));

});



// Master Excel file ka naam kuch bhi ho sakta hai; MASTERLIST sheet ko header row se detect karo.

app.get('/api/master-data', (req, res) => {

  try {

    const excelExtensions = new Set(['.xlsx', '.xlsm', '.xls', '.xlsb']);

    const files = fs.readdirSync(__dirname)

      .filter(name => !name.startsWith('~$') && excelExtensions.has(path.extname(name).toLowerCase()))

      .sort((a, b) => {

        const aPreferred = /^SEP\s*2026.*\.xlsx$/i.test(a) ? 0 : 1;

        const bPreferred = /^SEP\s*2026.*\.xlsx$/i.test(b) ? 0 : 1;

        return aPreferred - bPreferred || a.localeCompare(b);

      });

    let best = null;

    for (const file of files) {

      try {

        const workbook = xlsx.readFile(path.join(__dirname, file));

        const sheetName = workbook.SheetNames.find(name => String(name).trim().toUpperCase() === 'MASTERLIST');

        if (!sheetName) continue;

        const worksheet = workbook.Sheets[sheetName];

        const rows = xlsx.utils.sheet_to_json(worksheet, { header: 1, defval: '', raw: false });

        // MASTERLIST may have a merged MR/job title in row 1 and headers in row 2.

        const headerIndex = rows.findIndex(row => {

          const cells = (row || []).map(v => String(v || '').replace(/\s+/g, ' ').trim().toUpperCase());

          return cells.some(v => v === 'CODE' || v === 'CODE NO.' || v === 'CODE NO') &&

                 cells.some(v => v.includes('NAME'));

        });

        if (headerIndex < 0) continue;

        const dataRows = [];

        for (let i = headerIndex + 1; i < rows.length; i++) {

          const r = rows[i] || [];

          const code = String(r[1] ?? '').trim();

          const name = String(r[3] ?? '').trim();

          if (!code && !name) continue;

          const designation = String(r[2] || 'MAZ').trim();

          const fatherName = String(r[5] || '').trim();

          const gender = String(r[6] || 'MALE').trim();

          const accountNo = String(r[15] || '').trim();

          const ifsCode = String(r[16] || '').trim();

          const enrolmentDate = String(r[22] || '').trim();

          const dischargeDate = String(r[23] || '').trim();

          // Map MASTERLIST source columns into the MR MAKER's 44-column row layout.

          // Source Excel: B=Code, C=Trade, D=Name, F=Father, P=Bank A/C,

          // Q=IFSC, R=Working Days, S=Rate, T=Net Amount, W=Enrolment, X=Discharge.

          const workingDays = String(r[17] ?? '').trim();

          const rate = String(r[18] ?? '').trim();

          const netAmount = String(r[19] ?? '').trim();

          const rowArr = new Array(44).fill('');

          rowArr[1] = code;

          rowArr[2] = designation;

          rowArr[3] = name;

          rowArr[4] = fatherName;

          rowArr[5] = gender || 'MALE';

          rowArr[37] = workingDays;

          rowArr[38] = rate;

          rowArr[39] = netAmount;

          rowArr[40] = accountNo;

          rowArr[41] = ifsCode;

          rowArr[42] = enrolmentDate;

          rowArr[43] = dischargeDate;

          dataRows.push({

            code, name, designation, fatherName, gender: gender || 'MALE',

            workingDays, rate, netAmount, accountNo, ifsCode, enrolmentDate, dischargeDate,

            rowData: rowArr

          });

        }

        if (!best || dataRows.length > best.data.length) best = { file, sheetName, headers: rows[headerIndex] || [], data: dataRows };

      } catch (readError) {

        console.warn(`Master Excel read skipped (${file}):`, readError.message);

      }

    }

    if (!best) return res.json({ headers: [], data: [], sourceFile: '', sheetName: '', message: 'Project folder mein Excel file ke andar MASTERLIST sheet aur CODE/NAME header nahi mila.' });

    res.json({ headers: best.headers, data: best.data, sourceFile: best.file, sheetName: best.sheetName, count: best.data.length });

  } catch (error) {

    console.error('MASTERLIST Excel read error:', error);

    res.status(500).json({ error: 'MASTERLIST Excel read failed: ' + error.message });

  }

});



const PORT = Number(process.env.PORT || 3000);

async function startServer() {
  await initializeDatabaseStore();
  app.listen(PORT, '0.0.0.0', () => console.log(`Server is running on port ${PORT}`));
}

startServer().catch(error => {
  console.error('Server startup failed:', error.message);
  process.exit(1);
});