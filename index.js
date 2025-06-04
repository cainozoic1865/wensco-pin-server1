import express from 'express';
import { GoogleSpreadsheet } from 'google-spreadsheet';
import axios from 'axios';
import dotenv from 'dotenv';

dotenv.config();

const app = express();
const PORT = process.env.PORT || 10000;

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

async function processSheet() {
  const doc = new GoogleSpreadsheet(process.env.SHEET_ID);

  await doc.useServiceAccountAuth({
    client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
  });

  await doc.loadInfo();
  const sheet = doc.sheetsByTitle[process.env.SHEET_NAME];
  const rows = await sheet.getRows();
  const lastRow = rows[rows.length - 1];

  const email = lastRow['Email（接收PIN碼）'];
  const name = lastRow['聯絡人姓名'];
  const company = lastRow['租戶公司名稱'];
  const date = lastRow['預約日期'];
  const start = convertTime(date, lastRow['開始時間']);
  const end = convertTime(date, lastRow['結束時間']);

  const accessToken = await getAccessToken();

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
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      }
    }
  );

  const pin = response.data.pin;
  lastRow['PIN'] = pin;
  await lastRow.save();

  return `✅ 已產生 PIN：${pin} 給 ${email}`;
}

function convertTime(date, time) {
  const d = new Date(date);
  const [hour, minute] = time.replace('上午', '').replace('下午', '').split(':').map(Number);
  const isPM = time.includes('下午');
  d.setHours(isPM ? hour + 12 : hour);
  d.setMinutes(minute);
  return d.toISOString();
}

async function getAccessToken() {
  const response = await axios.post('https://api.igloohome.co/v2/token', {
    grant_type: 'client_credentials',
    client_id: process.env.CLIENT_ID,
    client_secret: process.env.CLIENT_SECRET,
    user_email: process.env.USER_EMAIL
  });
  return response.data.access_token;
}

app.listen(PORT, () => {
  console.log(`🚀 Server is running on port ${PORT}`);
});
