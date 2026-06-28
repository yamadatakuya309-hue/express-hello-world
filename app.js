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

    const resE = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!E:E`,
    });

    const eRows = resE.data.values || [];
    for (let i = 0; i < eRows.length; i++) {
      const cellValue = (eRows[i][0] || '').trim();
      if (cellValue === identCode.trim()) {
        const rowNumber = i + 1;

        const resRow = await sheets.spreadsheets.values.get({
          spreadsheetId: SPREADSHEET_ID,
          range: `${sheetName}!${rowNumber}:${rowNumber}`,
        });

        const rowData = (resRow.data.values || [[]])[0] || [];

        for (let colIndex = 9; colIndex < rowData.length + 20; colIndex++) {
          const colValue = (rowData[colIndex] || '').trim();

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

          if (
            (rowData[colIndex] || '').trim() === '' &&
            (rowData[colIndex + 1] || '').trim() === '' &&
            (rowData[colIndex + 2] || '').trim() === ''
          ) break;
        }

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
// メッセージを解析して「識別コードと担当者名」のペア一覧を返す
//
// 対応フォーマット：
//
// 【パターンA】各行に担当者名あり（今回の問題ケース）
//   太田市南矢島1-1、荒岡
//   太田市南矢島1-2、荒岡
//   太田市南矢島1-3、荒岡
//   → [{identCode:'太田市南矢島1-1', worker:'荒岡'}, ...]
//
// 【パターンB】識別コードのみ複数行、最後に担当者名
//   伊勢崎市西久保1-3、
//   伊勢崎市西久保1-5、
//   荒岡
//   → [{identCode:'伊勢崎市西久保1-3', worker:'荒岡'}, ...]
//
// 【パターンC】1件のみ
//   伊勢崎市東町3-4、荒岡
//   → [{identCode:'伊勢崎市東町3-4', worker:'荒岡'}]
// ============================
function parseCompletionMessage(text) {
  // 改行で分割し、前後の空白を除去・空行を除外
  const lines = text
    .split(/\n/)
    .map(line => line.trim())
    .filter(line => line.length > 0);

  const entries = [];

  for (const line of lines) {
    // 「、」が含まれる行 → 「識別コード、担当者名」として解析
    if (line.includes('、')) {
      const separatorIndex = line.indexOf('、');
      const identCode = line.substring(0, separatorIndex).trim();
      const worker = line.substring(separatorIndex + 1).trim();

      if (identCode) {
        entries.push({
          identCode,
          // 担当者名が空（例：「伊勢崎市西久保1-3、」）の場合はnullにして後で補完
          worker: worker.length > 0 ? worker : null,
        });
      }
    }
    // 「、」が含まれない行 → 担当者名のみの行（パターンBの最終行）
    else {
      const workerOnly = line.trim();
      if (workerOnly.length > 0) {
        // 直前までのworkerがnullのエントリに担当者名を補完
        for (const entry of entries) {
          if (entry.worker === null) {
            entry.worker = workerOnly;
          }
        }
      }
    }
  }

  return entries;
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
        text: '植栽メンテナンス報告システムへようこそ！\n\n【完了報告（1件）】\n識別コード、担当者名\n例：伊勢崎市東町3-4、荒岡\n\n【完了報告（複数・担当者名を各行に）】\n太田市南矢島1-1、荒岡\n太田市南矢島1-2、荒岡\n\n【完了報告（複数・担当者名を最後に）】\n伊勢崎市西久保1-3、\n伊勢崎市西久保1-5、\n荒岡\n\n【セル直接更新】\n更新 物件名 列 内容\n例：更新 山田様邸　請負工事 J 2026.06.13山田施工',
      });
      continue;
    }

    if (event.type !== 'message' || event.message.type !== 'text') continue;

    const text = event.message.text.trim();
    const senderId = event.source.userId || '';
    const groupId = event.source.groupId || event.source.roomId || '';
    const datetime = new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' });

    // ============================
    // 完了報告：「、」が含まれていて「更新 」で始まらない場合
    // ============================
    if (text.includes('、') && !text.startsWith('更新 ')) {

      const entries = parseCompletionMessage(text);

      // 解析失敗
      if (entries.length === 0) {
        await appendToErrorLog(datetime, senderId, groupId, text, '完了報告の解析失敗：エントリが0件');
        await client.replyMessage(event.replyToken, {
          type: 'text',
          text: '⚠️ 形式が正しくありません。\n\n例：\n伊勢崎市東町3-4、荒岡\n\n※内容は管理者が確認します。',
        });
        continue;
      }

      // 担当者名が取得できなかったエントリがある
      const missingWorker = entries.filter(e => !e.worker);
      if (missingWorker.length > 0) {
        await appendToErrorLog(datetime, senderId, groupId, text,
          `担当者名が取得できませんでした：${missingWorker.map(e => e.identCode).join('、')}`);
        await client.replyMessage(event.replyToken, {
          type: 'text',
          text: '⚠️ 担当者名が読み取れませんでした。\n\n【各行に担当者名をつける場合】\n太田市南矢島1-1、荒岡\n太田市南矢島1-2、荒岡\n\n【最後にまとめて担当者名をつける場合】\n伊勢崎市西久保1-3、\n伊勢崎市西久保1-5、\n荒岡\n\n※内容は管理者が確認します。',
        });
        continue;
      }

      // 今日の日付を自動取得
      const today = new Date();
      const dateStr = `${today.getFullYear()}.${String(today.getMonth() + 1).padStart(2, '0')}.${String(today.getDate()).padStart(2, '0')}`;

      // 各エントリを順番に処理
      const results = [];
      for (const entry of entries) {
        const completionValue = `${dateStr} ${entry.worker}`;
        try {
          const result = await completeByCode(entry.identCode, completionValue);
          results.push({ ...entry, completionValue, ...result });
        } catch (err) {
          results.push({ ...entry, completionValue, success: false, reason: 'error', error: err.message });
        }
      }

      // 結果サマリーを作成
      const successList = results.filter(r => r.success);
      const failList = results.filter(r => !r.success);

      let replyText = '';

      if (successList.length > 0) {
        replyText += `✅ ${successList.length}件の完了報告を記録しました！\n\n`;
        for (const r of successList) {
          replyText += `・${r.identCode}（${r.sheetName} ${r.column}列）\n  ${r.completionValue}\n`;
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
