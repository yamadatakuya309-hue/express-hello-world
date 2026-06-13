const express = require('express');
const line = require('@line/bot-sdk');
const { google } = require('googleapis');

const app = express();

const config = {
  channelAccessToken: 'a9Gm8O+G17f9rIcBAiZ1cD5O1AR2jyPrDJYZfxlYbn0S9+igtU4GIyWyB2hcwSwXkflxCKifxFZGd0YuhTfqfHqxnZWgyG91K6u0X1qhPebtKfsuIrfRoL058XvAI+FLA1PwrNUwjCCm8WPmN67hwAdB04t89/1O/w1cDnyilFU=',
  channelSecret: '14bd181a88c13f4e28d0b28f6626c682'
};

const client = new line.Client(config);
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

async function appendToSheet(values) {
  const credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  const sheets = google.sheets({ version: 'v4', auth });
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: 'Sheet1!A:D',
    valueInputOption: 'USER_ENTERED',
    resource: { values: [values] },
  });
}

app.post('/webhook', line.middleware(config), (req, res) => {
  Promise.all(req.body.events.map(handleEvent))
    .then((result) => res.json(result))
    .catch((err) => {
      console.error(err);
      res.status(500).end();
    });
});

async function handleEvent(event) {
  if (event.type === 'join') {
    return client.replyMessage(event.replyToken, {
      type: 'text',
      text: '植栽メンテナンス報告ボットです！よろしくお願いします。'
    });
  }

  if (event.type !== 'message' || event.message.type !== 'text') {
    return Promise.resolve(null);
  }

  const userMessage = event.message.text;
  const now = new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' });
  const userId = event.source.userId || '';
  const groupId = event.source.groupId || '個人トーク';

  await appendToSheet([now, userId, groupId, userMessage]);

  return client.replyMessage(event.replyToken, {
    type: 'text',
    text: '受け取りました：' + userMessage
  });
}

const port = process.env.PORT || 10000;
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
