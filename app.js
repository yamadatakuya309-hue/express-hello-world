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

// 列番号をアルファベットに変換（例：10 → J）
function colNumberToLetter(n) {
  let result = '';
  while (n > 0) {
    n--;
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result;
}

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
    resource: { values: [[datetime, senderId, groupId, message]] },
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
      resource: { values: [[datetime, senderId, groupId, message, errorDetail]] },
    });
  } catch (err) {
    console.error('エラーログ記録失敗:', err);
  }
}

// ============================
// 全シートを横断してE列の識別コードを検索し
// J列以降で「予定」のみのセルを探して上書き
// ============================
async function completeByCode(identCode, completionValue) {
  const auth = await getAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const sheetNames = spreadsheet.data.sheets.map(s => s.properties.title);

  for (const sheetName of sheetNames) {
    if (sheetName === 'シート1' || sheetName === 'エラーログ') continue;

    // E列（識別コード）を取得
    const resE = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!E:E`,
    });

    const eRows = resE.data.values || [];
    for (let i = 0; i < eRows.length; i++) {
      const cellValue = (eRows[i][0] || '').trim();
      if (cellValue === identCode.trim()) {
        const rowNumber = i + 1;

        // J列(10)以降を右に向かって無限に探す
        // まず行全体を取得
        const resRow = await sheets.spreadsheets.values.get({
          spreadsheetId: SPREADSHEET_ID,
          range: `${sheetName}!${rowNumber}:${rowNumber}`,
        });

        const rowData = (resRow.data.values || [[]])[0] || [];

        // J列はindex9（A=0,B=1,...,J=9）
        for (let colIndex = 9; colIndex < rowData.length + 20; colIndex++) {
          const colValue = (rowData[colIndex] || '').trim();

          // 「予定」のみ（完全一致）
          if (colValue === '予定') {
            const colLetter = colNumberToLetter(colIndex + 1);
            const updateRange = `${sheetName}!${colLetter}${rowNumber}`;
            await sheets.spreadsheets.values.update({
              spreadsheetId: SPREADSHEET_ID,
              range: updateRange,
              valueInputOption: 'USER_ENTERED',
              resource: { values: [[completionValue]] },
            });
            return { success: true, sheetName, rowNumber, column: colLetter };
          }

          // 空欄が3連続したら終了
          if (
            (rowData[colIndex] || '').trim() === '' &&
            (rowData[colIndex + 1] || '').trim() === '' &&
            (rowData[colIndex + 2] || '').trim() === ''
          ) break;
        }

        // 物件は見つかったが「予定」セルがない
        return { success: false, reason: 'no_yotei', sheetName };
      }
    }
  }
  return { success: false, reason: 'not_found' };
}

// ============================
// 全シートを横断してB列の物件名を検索し、指定列を更新
// ============================
async function updateCellByPropertyName(propertyName, columnLetter, newValue) {
  const auth = await getAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const sheetNames = spreadsheet.data.sheets.map(s => s.properties.title);

  for (const sheetName of sheetNames) {
    if (sheetName === 'シート1' || sheetName === 'エラーログ') continue;

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
        text: '植栽メンテナンス報告システムへようこそ！\n\n【完了報告】\n識別コード、担当者名\n例：伊勢崎市東町3-4、荒岡\n\n【セル直接更新】\n更新 物件名 列 内容\n例：更新 山田様邸　請負工事 J 2026.06.13山田施工',
      });
      continue;
    }

    if (event.type !== 'message' || event.message.type !== 'text') continue;

    const text = event.message.text.trim();
    const senderId = event.source.userId || '';
    const groupId = event.source.groupId || event.source.roomId || '';
    const datetime = new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' });

    // ============================
    // 完了報告：識別コード、担当者名
    // 読点（、）が含まれていて「更新 」で始まらない場合
    // ============================
    if (text.includes('、') && !text.startsWith('更新 ')) {
      const parts = text.split('、');
      if (parts.length < 2) {
        const errorDetail = '完了報告の形式エラー：読点の前後に内容がありません';
        await appendToErrorLog(datetime, senderId, groupId, text, errorDetail);
        await client.replyMessage(event.replyToken, {
          type: 'text',
          text: '⚠️ 形式が正しくありません。\n\n正しい形式：\n識別コード、担当者名\n\n例：\n伊勢崎市東町3-4、荒岡\n\n※内容は管理者が確認します。',
        });
        continue;
      }

      const identCode = parts[0].trim();
      const worker = parts[1].trim();

      // 今日の日付を自動取得
      const today = new Date();
      const dateStr = `${today.getFullYear()}.${String(today.getMonth() + 1).padStart(2, '0')}.${String(today.getDate()).padStart(2, '0')}`;
      const completionValue = `${dateStr} ${worker}`;

      try {
        const result = await completeByCode(identCode, completionValue);

        if (result.success) {
          await client.replyMessage(event.replyToken, {
            type: 'text',
            text: `✅ 完了報告を記録しました！\n\n現場：${identCode}\nシート：${result.sheetName}\n列：${result.column}列\n記録内容：${completionValue}`,
          });
        } else if (result.reason === 'no_yotei') {
          const errorDetail = `「予定」セルなし：識別コード「${identCode}」（${result.sheetName}）のJ列以降に「予定」が見つかりませんでした`;
          await appendToErrorLog(datetime, senderId, groupId, text, errorDetail);
          await client.replyMessage(event.replyToken, {
            type: 'text',
            text: `⚠️ 「${identCode}」の予定セルが見つかりませんでした。\n\n内容は記録しましたので管理者が確認します。`,
          });
        } else {
          const errorDetail = `識別コード「${identCode}」がスプレッドシートに見つかりませんでした`;
          await appendToErrorLog(datetime, senderId, groupId, text, errorDetail);
          await client.replyMessage(event.replyToken, {
            type: 'text',
            text: `⚠️ 「${identCode}」が見つかりませんでした。\n\n内容は記録しましたので管理者が確認します。`,
          });
        }
      } catch (err) {
        const errorDetail = `システムエラー：${err.message}`;
        await appendToErrorLog(datetime, senderId, groupId, text, errorDetail);
        console.error('完了報告エラー:', err);
        await client.replyMessage(event.replyToken, {
          type: 'text',
          text: '⚠️ エラーが発生しました。\n\n内容は記録しましたので管理者が確認します。',
        });
      }
      continue;
    }

    // ============================
    // 「更新」コマンドの処理
    // ============================
    if (text.startsWith('更新 ')) {
      const parts = text.split(' ');
      if (parts.length < 4) {
        const errorDetail = 'コマンド形式エラー：更新コマンドの要素が不足しています';
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
          const errorDetail = `物件名「${propertyName}」がスプレッドシートに見つかりませんでした`;
          await appendToErrorLog(datetime, senderId, groupId, text, errorDetail);
          await client.replyMessage(event.replyToken, {
            type: 'text',
            text: `⚠️ 「${propertyName}」が見つかりませんでした。\n\n内容は記録しましたので管理者が確認します。`,
          });
        }
      } catch (err) {
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
