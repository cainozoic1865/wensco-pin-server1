const express = require('express');
const { google } = require('googleapis');
const axios = require('axios');
const https = require('https');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

const SHEET_ID = process.env.SHEET_ID;
const SHEET_NAME = process.env.SHEET_NAME;
const DEVICE_ID = process.env.DEVICE_ID;
const BRIDGE_ID = process.env.BRIDGE_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const USER_EMAIL = process.env.USER_EMAIL;
const GOOGLE_SERVICE_ACCOUNT_EMAIL = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
const GOOGLE_PRIVATE_KEY = process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n');

const auth = new google.auth.JWT(
  GOOGLE_SERVICE_ACCOUNT_EMAIL,
  null,
  GOOGLE_PRIVATE_KEY,
  ['https://www.googleapis.com/auth/spreadsheets']
);
const sheets = google.sheets({ version: 'v4', auth });

async function getAccessToken() {
  const agent = new https.Agent({ rejectUnauthorized: false });

  const res = await axios.post('https://api.igloohome.co/v2/token', {
    grant_type: 'client_credentials',
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    user_email: USER_EMAIL
  }, {
    headers: { 'Content-Type': 'application/json' },
    httpsAgent: agent
  });

  return res.data.access_token;
}

async function createIgloohomePin(token, start, end) {
  const agent = new https.Agent({ rejectUnauthorized: false });

  const url = `https://api.igloohome.co/v2/devices/${DEVICE_ID}/pins/duration/hourly`;
  const res = await axios.post(url, {
    start,
    end,
    timezone: 'Asia/Taipei',
    bridge_id: BRIDGE_ID
  }, {
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    httpsAgent: agent
  });

  return res.data.pin;
}

async function writePinToSheet(rowIndex, pin) {
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!H${rowIndex + 1}`,
    valueInputOption: 'RAW',
    requestBody: {
      values: [[pin]]
    }
  });
}

async function processSheet() {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: SHEET_NAME
  });

  const rows = res.data.values;
  if (!rows || rows.length === 0) {
    console.log('❗ 表單資料為空');
    return 'No data';
  }

  const token = await getAccessToken();

  let resultLog = '';

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const email = row[3];
    const date = row[4];
    const startTime = row[5];
    const endTime = row[6];
    const pinExists = row[7];

    if (!email || !date || !startTime || !endTime || pinExists) continue;

    const dateStr = new Date(date).toISOString().split('T')[0];
    const start = new Date(`${dateStr}T${formatTime(startTime)}:00+08:00`);
    const end = new Date(`${dateStr}T${formatTime(endTime)}:00+08:00`);

    const pin = await createIgloohomePin(token, start.toISOString(), end.toISOString());
    await writePinToSheet(i, pin);

    const logLine = `✅ 第 ${i + 1} 列建立 PIN：${pin}`;
    console.log(logLine);
    resultLog += logLine + '\n';
  }

  return resultLog || 'No new pins to process';
}

function formatTime(value) {
  if (typeof value === 'string') {
    const date = new Date(`2000-01-01 ${value}`);
    return date.toTimeString().substring(0, 5);
  } else {
    return new Date(value).toTimeString().substring(0, 5);
  }
}

// === Express API ===

app.get('/', (req, res) => {
  res.send('✅ Wensco PIN Server is running.');
});

app.get('/run', async (req, res) => {
  try {
    const result = await processSheet();
    res.send(result);
  } catch (err) {
    console.error('❌ 執行失敗：', err.message);
    res.status(500).send(`Error: ${err.message}`);
  }
});

app.listen(PORT, () => {
  console.log(`🚀 Server is running on port ${PORT}`);
});
