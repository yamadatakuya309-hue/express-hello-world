const express = require('express');
const line = require('@line/bot-sdk');
const { google } = require('googleapis');

const app = express();

// LINE設定
const lineConfig = {
  channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN,
  channelSecret: process.env.LINE_CHANNEL_SECRET,
};
const client = new line.Client(lineConfig);

// Google Sheets認証
function getAuthClient() {
  const credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS);
  return new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
}

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

// ============================
// メッセージをシート1に記録
// ============================
async function appendToSheet(datetime, senderId, groupId, message) {
  const auth = await getAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: 'シート1!A:D',
    valueInputOption: 'USER_ENTERED',
    resource: {
      values: [[datetime, senderId, groupId, message]],
    },
  });
}

// ============================
// エラーログをスプレッドシートに記録
// ============================
async function appendToErrorLog(datetime, senderId, groupId, message, errorDetail) {
  try {
    const auth = await getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: 'エラーログ!A:E',
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: [[datetime, senderId, groupId, message, errorDetail]],
      },
    });
  } catch (err) {
    console.error('エラーログ記録失敗:', err);
  }
}

// ============================
// 全シートを横断して物件名を検索し、指定列を更新
// ============================
async function updateCellByPropertyName(propertyName, columnLetter, newValue) {
  const auth = await getAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  // シート一覧を取得
  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const sheetNames = spreadsheet.data.sheets.map(s => s.properties.title);

  for (const sheetName of sheetNames) {
    // 記録・管理用シートはスキップ
    if (sheetName === 'シート1' || sheetName === 'エラーログ') continue;

    // B列（物件名）を取得
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!B:B`,
    });

    const rows = res.data.values || [];
    for (let i = 0; i < rows.length; i++) {
      const cellValue = (rows[i][0] || '').trim();
      if (cellValue === propertyName.trim()) {
        const rowNumber = i + 1;
        const updateRange = `${sheetName}!${columnLetter}${rowNumber}`;
        await sheets.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range: updateRange,
          valueInputOption: 'USER_ENTERED',
          resource: { values: [[newValue]] },
        });
        return { success: true, sheetName, rowNumber, updateRange };
      }
    }
  }

  return { success: false };
}

// ============================
// Webhookエンドポイント
// ============================
app.post('/webhook', line.middleware(lineConfig), async (req, res) => {
  res.sendStatus(200);

  const events = req.body.events;
  for (const event of events) {

    // グループ参加時の挨拶
    if (event.type === 'join') {
      await client.replyMessage(event.replyToken, {
        type: 'text',
        text: '植栽メンテナンス報告システムへようこそ！\n\nメッセージを送ると自動で記録されます。\n\n【セル更新コマンド】\n更新 物件名 列 内容\n例：更新 山田様邸　請負工事 J 2026.06.13山田施工',
      });
      continue;
    }

    if (event.type !== 'message' || event.message.type !== 'text') continue;

    const text = event.message.text.trim();
    const senderId = event.source.userId || '';
    const groupId = event.source.groupId || event.source.roomId || '';
    const datetime = new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' });

    // ============================
    // 「更新」コマンドの処理
    // ============================
    if (text.startsWith('更新 ')) {
      const parts = text.split(' ');

      // コマンド形式チェック
      if (parts.length < 4) {
        const errorDetail = 'コマンド形式エラー：要素が不足しています';
        await appendToErrorLog(datetime, senderId, groupId, text, errorDetail);
        await client.replyMessage(event.replyToken, {
          type: 'text',
          text: '⚠️ コマンドの形式が正しくありません。\n\n正しい形式：\n更新 物件名 列 内容\n\n例：\n更新 山田様邸　請負工事 J 2026.06.13山田施工\n\n※内容は管理者が確認します。',
        });
        continue;
      }

      const propertyName = parts[1];
      const columnLetter = parts[2].toUpperCase();
      const newValue = parts.slice(3).join(' ');

      // 列名バリデーション
      if (!/^[A-Z]+$/.test(columnLetter)) {
        const errorDetail = `列名エラー：「${columnLetter}」は無効な列名です`;
        await appendToErrorLog(datetime, senderId, groupId, text, errorDetail);
        await client.replyMessage(event.replyToken, {
          type: 'text',
          text: `⚠️ 列名「${columnLetter}」が正しくありません。\nA〜Zのアルファベットで指定してください。\n\n※内容は管理者が確認します。`,
        });
        continue;
      }

      try {
        const result = await updateCellByPropertyName(propertyName, columnLetter, newValue);
        if (result.success) {
          await client.replyMessage(event.replyToken, {
            type: 'text',
            text: `✅ 更新しました！\n\n物件名：${propertyName}\nシート：${result.sheetName}\n列：${columnLetter}列（${result.rowNumber}行目）\n内容：${newValue}`,
          });
        } else {
          // 物件名が見つからない → エラーログに記録
          const errorDetail = `物件名「${propertyName}」がスプレッドシートに見つかりませんでした`;
          await appendToErrorLog(datetime, senderId, groupId, text, errorDetail);
          await client.replyMessage(event.replyToken, {
            type: 'text',
            text: `⚠️ 「${propertyName}」が見つかりませんでした。\n\n内容は記録しましたので管理者が確認します。`,
          });
        }
      } catch (err) {
        // システムエラー → エラーログに記録
        const errorDetail = `システムエラー：${err.message}`;
        await appendToErrorLog(datetime, senderId, groupId, text, errorDetail);
        console.error('更新エラー:', err);
        await client.replyMessage(event.replyToken, {
          type: 'text',
          text: '⚠️ 更新中にエラーが発生しました。\n\n内容は記録しましたので管理者が確認します。',
        });
      }
      continue;
    }

    // ============================
    // 通常メッセージの記録
    // ============================
    try {
      await appendToSheet(datetime, senderId, groupId, text);
    } catch (err) {
      console.error('記録エラー:', err);
    }
  }
});

// ヘルスチェック（UptimeRobot用）
app.get('/', (req, res) => {
  res.send('植栽メンテナンス報告システム 稼働中');
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
