const { google } = require('googleapis');
const axios = require('axios');
require('dotenv').config();

const SHEET_ID = process.env.SHEET_ID;
const SHEET_NAME = process.env.SHEET_NAME;
const DEVICE_ID = process.env.DEVICE_ID;
const BRIDGE_ID = process.env.BRIDGE_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const USER_EMAIL = process.env.USER_EMAIL;

const GOOGLE_SERVICE_ACCOUNT_EMAIL = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
const GOOGLE_PRIVATE_KEY = process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n');

// 建立 Google Sheets API 客戶端
const auth = new google.auth.JWT(
  GOOGLE_SERVICE_ACCOUNT_EMAIL,
  null,
  GOOGLE_PRIVATE_KEY,
  ['https://www.googleapis.com/auth/spreadsheets']
);
const sheets = google.sheets({ version: 'v4', auth });

// 建立 Igloohome Access Token
async function getAccessToken() {
  const res = await axios.post('https://api.igloohome.io/v2/token', {
    grant_type: 'client_credentials',
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    user_email: USER_EMAIL
  }, {
    headers: { 'Content-Type': 'application/json' }
  });

  return res.data.access_token;
}

// 建立 PIN
async function createIgloohomePin(token, start, end) {
  const url = `https://api.igloohome.io/v2/devices/${DEVICE_ID}/pins/duration/hourly`;
  const res = await axios.post(url, {
    start,
    end,
    timezone: 'Asia/Taipei',
    bridge_id: BRIDGE_ID
  }, {
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    }
  });

  return res.data.pin;
}

// 將 PIN 寫回試算表
async function writePinToSheet(rowIndex, pin) {
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!H${rowIndex + 1}`, // H 欄是第 8 欄
    valueInputOption: 'RAW',
    requestBody: {
      values: [[pin]]
    }
  });
}

// 主要處理流程
async function processSheet() {
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: SHEET_NAME
    });

    const rows = res.data.values;
    if (!rows || rows.length === 0) {
      console.log('❗ 表單資料為空');
      return;
    }

    const token = await getAccessToken();

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
      console.log(`✅ 第 ${i + 1} 列建立 PIN：${pin}`);
      await writePinToSheet(i, pin);
    }
  } catch (err) {
    console.error('❌ 錯誤：', err.response?.data || err.message);
  }
}

// 將「上午 8:00:00」或 Date 格式轉換為 24 小時制字串
function formatTime(value) {
  if (typeof value === 'string') {
    const date = new Date(`2000-01-01 ${value}`);
    return date.toTimeString().substring(0, 5);
  } else {
    return new Date(value).toTimeString().substring(0, 5);
  }
}

// 啟動主程序
processSheet();
