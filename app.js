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
// 【新機能】複数の識別コードをまとめたメッセージを解析する
// 例：
//   伊勢崎市西久保1-3、\n伊勢崎市西久保1-5、\n完了です
//   → ['伊勢崎市西久保1-3', '伊勢崎市西久保1-5'] と worker='完了です'
//
// ルール：
//   - 改行で各行に分割
//   - 末尾の「、」「，」「,」を除去
//   - 「完了」「終了」「以上」「done」「ok」で始まる行 → worker（担当者名）として扱う
//   - 空行は除外
//   - 担当者名が見つからなかった場合は最後の行を担当者名として扱う
// ============================
function parseCompletionMessage(text) {
  // 改行で分割し、前後の空白と末尾の読点を除去
  const lines = text
    .split(/\n/)
    .map(line => line.trim().replace(/[、，,]$/, ''))
    .filter(line => line.length > 0);

  // 「完了」「終了」「以上」「done」「ok」で始まる行を担当者/ステータス行として検出
  const statusPattern = /^(完了|終了|以上|done|ok)/i;

  const identCodes = [];
  let worker = null;

  for (const line of lines) {
    if (statusPattern.test(line)) {
      // ステータス行：「完了です」→ worker = '完了'、「完了 荒岡」→ worker = '荒岡'
      const workerMatch = line.replace(statusPattern, '').trim();
      worker = workerMatch.length > 0 ? workerMatch : line;
    } else {
      identCodes.push(line);
    }
  }

  // 担当者名が見つからなかった場合は最後の行を担当者として扱う
  if (!worker && identCodes.length > 0) {
    worker = identCodes.pop();
  }

  return { identCodes, worker };
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
        text: '植栽メンテナンス報告システムへようこそ！\n\n【完了報告】\n識別コード、担当者名\n例：伊勢崎市東町3-4、荒岡\n\n【複数まとめて完了報告】\n識別コードを改行して最後に担当者名\n例：\n伊勢崎市西久保1-3、\n伊勢崎市西久保1-5、\n荒岡\n\n【セル直接更新】\n更新 物件名 列 内容\n例：更新 山田様邸　請負工事 J 2026.06.13山田施工',
      });
      continue;
    }

    if (event.type !== 'message' || event.message.type !== 'text') continue;

    const text = event.message.text.trim();
    const senderId = event.source.userId || '';
    const groupId = event.source.groupId || event.source.roomId || '';
    const datetime = new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' });

    // ============================
    // 完了報告：識別コード（複数可）、担当者名
    // 読点（、）が含まれていて「更新 」で始まらない場合
    // ============================
    if (text.includes('、') && !text.startsWith('更新 ')) {

      const { identCodes, worker } = parseCompletionMessage(text);

      if (identCodes.length === 0 || !worker) {
        const errorDetail = '完了報告の形式エラー：識別コードまたは担当者名が取得できませんでした';
        await appendToErrorLog(datetime, senderId, groupId, text, errorDetail);
        await client.replyMessage(event.replyToken, {
          type: 'text',
          text: '⚠️ 形式が正しくありません。\n\n【1件の場合】\n識別コード、担当者名\n例：伊勢崎市東町3-4、荒岡\n\n【複数の場合】\n識別コードを改行して最後に担当者名\n例：\n伊勢崎市西久保1-3、\n伊勢崎市西久保1-5、\n荒岡\n\n※内容は管理者が確認します。',
        });
        continue;
      }

      // 今日の日付を自動取得
      const today = new Date();
      const dateStr = `${today.getFullYear()}.${String(today.getMonth() + 1).padStart(2, '0')}.${String(today.getDate()).padStart(2, '0')}`;
      const completionValue = `${dateStr} ${worker}`;

      // 複数の識別コードを順番に処理
      const results = [];
      for (const identCode of identCodes) {
        try {
          const result = await completeByCode(identCode, completionValue);
          results.push({ identCode, ...result });
        } catch (err) {
          results.push({ identCode, success: false, reason: 'error', error: err.message });
        }
      }

      // 結果のサマリーを作成
      const successList = results.filter(r => r.success);
      const failList = results.filter(r => !r.success);

      let replyText = '';

      if (successList.length > 0) {
        replyText += `✅ ${successList.length}件の完了報告を記録しました！\n`;
        replyText += `記録内容：${completionValue}\n\n`;
        for (const r of successList) {
          replyText += `・${r.identCode}（${r.sheetName} ${r.column}列）\n`;
        }
      }

      if (failList.length > 0) {
        replyText += `\n⚠️ ${failList.length}件が記録できませんでした：\n`;
        for (const r of failList) {
          const reason =
            r.reason === 'no_yotei' ? '「予定」セルが見つかりません' :
            r.reason === 'not_found' ? '識別コードが見つかりません' :
            'エラーが発生しました';
          replyText += `・${r.identCode}：${reason}\n`;

          // エラーログに記録
          await appendToErrorLog(datetime, senderId, groupId, text,
            `識別コード「${r.identCode}」：${reason}`);
        }
        replyText += '\n内容は管理者が確認します。';
      }

      await client.replyMessage(event.replyToken, {
        type: 'text',
        text: replyText.trim(),
      });

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
