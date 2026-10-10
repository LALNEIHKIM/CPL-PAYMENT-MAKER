const express = require('express');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const cors = require('cors'); // Sirf ek baar import karna hai

const app = express();
const PORT = 4000;

// CORS Enable kiya taaki kisi bhi PC ya IP se requests accept ho sakein
app.use(cors({
  origin: '*',
  methods: ['GET', 'POST', 'PUT', 'DELETE'],
  allowedHeaders: ['Content-Type', 'Authorization']
}));

// BADHAYI GAYI LIMIT: Large MR sheets aur data ke liye 50mb limit set ki gayi hai
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));

app.use(express.static(path.join(__dirname, 'public')));

// DEFAULT ROUTE: Open login.html directly when visiting http://localhost:4000/
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

// Helper function to get filename based on period
function getDataFilePath(period) {
  const cleanPeriod = period ? String(period).replace(/[^a-zA-Z0-9_]/g, '_').toUpperCase() : 'SEP_2026';
  return `./mr_data_${cleanPeriod}.json`;
}

// UNIVERSAL DATE FORMATTER (HANDLES SERIAL, DATE OBJECT & RAW STRING)
function parseAnyDateToDDMonYYYY(val) {
  if (val === null || val === undefined || val === '') return '';

  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

  // 1. Agar Date Object hai
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    const adjustedDate = new Date(val.getTime() + (12 * 60 * 60 * 1000));
    const day = String(adjustedDate.getUTCDate()).padStart(2, '0');
    const mon = months[adjustedDate.getUTCMonth()];
    const year = adjustedDate.getUTCFullYear();
    return `${day}-${mon}-${year}`;
  }

  const strVal = String(val).trim();

  // 2. Agar raw Date string hai
  if (strVal.includes('GMT') || strVal.includes('Time') || (isNaN(strVal) && !isNaN(Date.parse(strVal)))) {
    const d = new Date(strVal);
    if (!isNaN(d.getTime())) {
      const adjustedDate = new Date(d.getTime() + (12 * 60 * 60 * 1000));
      const day = String(adjustedDate.getUTCDate()).padStart(2, '0');
      const mon = months[adjustedDate.getUTCMonth()];
      const year = adjustedDate.getUTCFullYear();
      return `${day}-${mon}-${year}`;
    }
  }

  // 3. Agar Excel Serial Number hai
  const serial = parseFloat(strVal);
  if (!isNaN(serial) && serial > 10000 && serial < 80000) {
    const utcDays = Math.floor(serial - 25569);
    const utcValue = utcDays * 86400;
    const dateInfo = new Date((utcValue * 1000) + (12 * 60 * 60 * 1000));

    const day = String(dateInfo.getUTCDate()).padStart(2, '0');
    const mon = months[dateInfo.getUTCMonth()];
    const year = dateInfo.getUTCFullYear();
    return `${day}-${mon}-${year}`;
  }

  return strVal;
}

// AUTO-PICK CURRENT MONTH'S EXCEL FILE
app.get('/api/master-data', (req, res) => {
  try {
    const files = fs.readdirSync(__dirname).filter(f => f.endsWith('.xlsx') && !f.startsWith('~$'));

    if (files.length === 0) {
      return res.status(404).json({ error: "Koi bhi .xlsx file project folder me nahi mili!" });
    }

    const activeExcelFile = files[0];
    const filePath = path.join(__dirname, activeExcelFile);
    console.log(`[EXCEL LOADED]: Reading ${activeExcelFile} -> Sheet: MASTERLIST`);

    const workbook = XLSX.readFile(filePath, { cellDates: true });
    const sheet = workbook.Sheets['MASTERLIST'];

    if (!sheet) {
      return res.status(404).json({ error: `Sheet 'MASTERLIST' not found in ${activeExcelFile}!` });
    }

    const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    if (!rawRows || rawRows.length < 2) {
      return res.json({ headers: [], data: [] });
    }

    const headers = rawRows[1].slice(0, 24).map(h => String(h || '').trim());

    const data = [];
    for (let r = 2; r < rawRows.length; r++) {
      const row = rawRows[r];
      const code = String(row[1] || '').trim();
      const name = String(row[3] || '').trim();

      if (code || name) {
        const rowData = [];
        for (let c = 0; c < 24; c++) {
          let cellValue = row[c] !== undefined ? row[c] : '';

          if (c === 4 || c === 22 || c === 23) {
            cellValue = parseAnyDateToDDMonYYYY(cellValue);
          }

          rowData.push(String(cellValue || '').trim());
        }
        data.push({ code, name, rowData });
      }
    }

    res.json({ headers, data });
  } catch (err) {
    console.error("Error loading master data:", err.message);
    res.status(500).json({ error: err.message });
  }
});

// Get MR Sheets & DO-II Data API (Period-wise support)
app.get('/api/mr-sheets', (req, res) => {
  const period = req.query.period || 'SEP_2026';
  const dataFile = getDataFilePath(period);

  if (fs.existsSync(dataFile)) {
    fs.readFile(dataFile, 'utf8', (err, data) => {
      if (err) return res.status(500).json({ error: "Failed to read data" });
      try {
        res.json(JSON.parse(data));
      } catch (e) {
        res.json({ mrSheets: [], lastIndex: 0, do2Groups: [] });
      }
    });
  } else {
    const legacyFile = './mr_data.json';
    if (period === 'SEP_2026' && fs.existsSync(legacyFile)) {
      fs.readFile(legacyFile, 'utf8', (err, data) => {
        if (!err) {
          try { return res.json(JSON.parse(data)); } catch (e) {}
        }
        res.json({ mrSheets: [], lastIndex: 0, do2Groups: [] });
      });
    } else {
      res.json({ mrSheets: [], lastIndex: 0, do2Groups: [] });
    }
  }
});

// Save MR Sheets & DO-II Data API (Period-wise saving)
app.post('/api/mr-sheets', (req, res) => {
  const { mrSheets, lastIndex, do2Groups, period } = req.body;
  const activePeriod = period || 'SEP_2026';
  const dataFile = getDataFilePath(activePeriod);

  let existingData = {};
  if (fs.existsSync(dataFile)) {
    try {
      existingData = JSON.parse(fs.readFileSync(dataFile, 'utf8'));
    } catch (e) {}
  }

  const payload = {
    mrSheets: mrSheets !== undefined ? mrSheets : (existingData.mrSheets || []),
    lastIndex: lastIndex !== undefined ? lastIndex : (existingData.lastIndex || 0),
    do2Groups: do2Groups !== undefined ? do2Groups : (existingData.do2Groups || [])
  };

  fs.writeFile(dataFile, JSON.stringify(payload, null, 2), (err) => {
    if (err) return res.status(500).json({ error: "Failed to save data" });
    res.json({ success: true });
  });
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`=========================================`);
  console.log(`Server running at http://localhost:${PORT}`);
  console.log(`CORS Enabled for multi-PC connection`);
  console.log(`Period-Wise Data Isolation Enabled`);
  console.log(`=========================================`);
});