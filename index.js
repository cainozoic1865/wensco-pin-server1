// index.js (修正版 - 支援中文時間格式解析)
import express from 'express';
import axios from 'axios';
import https from 'https';
import { GoogleSpreadsheet } from 'google-spreadsheet';
import 'dotenv/config';

const app = express();
const port = process.env.PORT || 10000;

const {
  CLIENT_ID,
  CLIENT_SECRET,
  USER_EMAIL,
  BRIDGE_ID,
  DEVICE_ID,
  GOOGLE_PRIVATE_KEY,
  GOOGLE_SERVICE_ACCOUNT_EMAIL,
  SHEET_ID,
  SHEET_NAME
} = process.env;

const httpsAgent = new https.Agent({ rejectUnauthorized: false });

const getAccessToken = async () => {
  try {
    const res = await axios.post(
      'https://api.igloohome.co/v2/token',
      {
        grant_type: 'client_credentials',
        client_id: CLIENT_ID,
        client_secret: CLIENT_SECRET,
        user_email: USER_EMAIL
      },
      { httpsAgent }
    );
    return res.data.access_token;
  } catch (err) {
    console.error('❌ 取得 Access Token 失敗：', err.message);
    throw err;
  }
};

const createAlgoPin = async (accessToken, startTime, endTime) => {
  try {
    const res = await axios.post(
      `https://api.igloohome.co/v2/devices/${DEVICE_ID}/algo-pins/duration-hourly`,
      {
        type: 'duration_hourly',
        start_date: startTime,
        end_date: endTime,
        name: 'Wensco auto PIN'
      },
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        httpsAgent
      }
    );
    return res.data.algo_pin;
  } catch (err) {
    console.error('❌ 建立 ALGO PIN 失敗：', err.message);
    throw err;
  }
};

function parseTaiwanTime(dateStr, timeStr) {
  const date = new Date(dateStr);
  const isAM = timeStr.includes('上午');
  const isPM = timeStr.includes('下午');
  const timeCleaned = timeStr.replace('上午', '').replace('下午', '').trim();
  const [hourStr, minStr] = timeCleaned.split(':');
  let hour = parseInt(hourStr);
  if (isPM && hour < 12) hour += 12;
  if (isAM && hour === 12) hour = 0;
  date.setHours(hour);
  date.setMinutes(parseInt(minStr));
  date.setSeconds(0);
  return date.toISOString();
}

const processSheet = async () => {
  try {
    const doc = new GoogleSpreadsheet(SHEET_ID);
    await doc.useServiceAccountAuth({
      client_email: GOOGLE_SERVICE_ACCOUNT_EMAIL,
      private_key: GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n')
    });

    await doc.loadInfo();
    const sheet = doc.sheetsByTitle[SHEET_NAME];
    const rows = await sheet.getRows();

    const now = new Date();
    for (const row of rows) {
      if (!row['PIN碼']) {
        try {
          const start = parseTaiwanTime(row['預約日期'], row['開始時間']);
          const end = parseTaiwanTime(row['預約日期'], row['結束時間']);
          const startDate = new Date(start);
          if (startDate < now) {
            row['狀態'] = '已過期';
            await row.save();
            continue;
          }

          const accessToken = await getAccessToken();
          const pin = await createAlgoPin(accessToken, start, end);

          row['PIN碼'] = pin;
          row['狀態'] = 'Success';
          await row.save();
          console.log(`✅ 為 ${row['Email（接收PIN碼）']} 產生 PIN：${pin}`);
        } catch (innerErr) {
          row['狀態'] = '❌ 失敗';
          row['PIN碼'] = `錯誤: ${innerErr.message}`;
          await row.save();
          console.error('❌ 單筆處理失敗：', innerErr.message);
        }
      }
    }
    return '✅ Google Sheet 處理完成';
  } catch (err) {
    console.error('🧾 Google Sheet 處理失敗：', err);
    throw err;
  }
};

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

app.listen(port, () => {
  console.log(`🚀 Server is running on port ${port}`);
});
