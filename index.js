// ✅ Wensco PIN Server - index.js with improved error handling

import express from 'express';
import axios from 'axios';
import dotenv from 'dotenv';
import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';

dotenv.config();

const app = express();
const port = process.env.PORT || 10000;

app.get('/', (req, res) => {
  res.send('✅ Wensco PIN Server is running.');
});

app.get('/run', async (req, res) => {
  try {
    const result = await processSheet();
    res.send(result);
  } catch (err) {
    console.error('❌ 執行失敗：', err.message);
    console.error('🔥 詳細錯誤：', err);
    res.status(500).send(`Error: ${err?.response?.data || err?.message || 'Unknown error'}`);
  }
});

async function getAccessToken() {
  try {
    const response = await axios.post('https://api.igloohome.co/v2/token', {
      grant_type: 'client_credentials',
      client_id: process.env.CLIENT_ID,
      client_secret: process.env.CLIENT_SECRET,
      user_email: process.env.USER_EMAIL
    });
    return response.data.access_token;
  } catch (err) {
    console.error('🔒 Token 取得失敗：', err.message);
    console.error('📦 API Response:', err.response?.data);
    return null;
  }
}

async function createAlgoPin(start, end) {
  const token = await getAccessToken();
  if (!token) {
    throw new Error('無法取得 access token，停止建立 PIN');
  }

  try {
    const response = await axios.post(
      `https://api.igloohome.co/v2/devices/${process.env.DEVICE_ID}/pins/duration/hourly`,
      {
        start,
        end,
        timezone: 'Asia/Taipei',
        bridge_id: process.env.BRIDGE_ID
      },
      {
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json'
        }
      }
    );

    return response.data.pin;
  } catch (err) {
    console.error('🔐 建立 PIN 失敗：', err.message);
    console.error('📦 API Response:', err.response?.data);
    return null;
  }
}

async function processSheet() {
  try {
    const auth = new JWT({
      email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
      scopes: ['https://www.googleapis.com/auth/spreadsheets']
    });

    const doc = new GoogleSpreadsheet(process.env.SHEET_ID, auth);
    await doc.loadInfo();
    const sheet = doc.sheetsByTitle[process.env.SHEET_NAME];
    const rows = await sheet.getRows();

    let updates = [];

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (!row.PIN && row['預約日期'] && row['開始時間'] && row['結束時間']) {
        const startDate = parseDateTime(row['預約日期'], row['開始時間']);
        const endDate = parseDateTime(row['預約日期'], row['結束時間']);

        const pin = await createAlgoPin(startDate.toISOString(), endDate.toISOString());
        if (!pin) {
          updates.push(`⚠️ 第 ${i + 2} 列建立 PIN 失敗`);
          continue;
        }

        row.PIN = pin;
        await row.save();
        updates.push(`✅ 第 ${i + 2} 列建立 PIN：${pin}`);
      }
    }

    return updates.length > 0 ? updates.join('\n') : 'No new pins to process';
  } catch (err) {
    console.error('🧾 Google Sheet 處理失敗：', err.message);
    throw err;
  }
}

function parseDateTime(dateStr, timeStr) {
  const date = new Date(dateStr);
  const [h, m] = timeStr.replace('上午', '').replace('下午', '').trim().split(':');
  const hour = timeStr.includes('下午') ? parseInt(h) + 12 : parseInt(h);
  date.setHours(hour);
  date.setMinutes(parseInt(m));
  date.setSeconds(0);
  return date;
}

app.listen(port, () => {
  console.log(`🚀 Server is running on port ${port}`);
});
