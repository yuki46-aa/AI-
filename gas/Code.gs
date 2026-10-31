/**
 * yado. お問い合わせ受付
 *
 * 設定手順:
 * 1. Googleスプレッドシートを新規作成する
 * 2. 拡張機能 → Apps Script を開き、このファイルの内容を貼り付ける
 * 3. 下の SPREADSHEET_ID にスプレッドシートIDを入れる
 *    （URLの /d/ と /edit の間）
 * 4. デプロイ → 新しいデプロイ → 種類: ウェブアプリ
 *    - 次のユーザーとして実行: 自分
 *    - アクセスできるユーザー: 全員
 * 5. 発行されたウェブアプリURLを index.html の GAS_ENDPOINT に貼る
 */

const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID';
const SHEET_NAME = 'inquiries';
const HEADERS = ['受信日時', 'お名前', 'メール', '宿名', '宿泊日', '人数', '希望条件', '流入元'];

function doPost(e) {
  try {
    const payload = parsePayload_(e);
    const error = validate_(payload);
    if (error) return json_({ ok: false, error: error }, 400);

    const sheet = getSheet_();
    sheet.appendRow([
      new Date(),
      payload.name,
      payload.email,
      payload.hotel,
      payload.date,
      payload.guests,
      payload.message,
      payload.source
    ]);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) }, 500);
  }
}

function doGet() {
  return json_({ ok: true, service: 'yado-inquiry' });
}

function parsePayload_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('送信データが空です');
  }
  const raw = e.postData.contents;
  const type = String(e.postData.type || '');
  if (type.indexOf('application/x-www-form-urlencoded') !== -1) {
    return {
      name: param_(e, 'name'),
      email: param_(e, 'email'),
      hotel: param_(e, 'hotel'),
      date: param_(e, 'date'),
      guests: param_(e, 'guests'),
      message: param_(e, 'message'),
      source: param_(e, 'source') || 'lp'
    };
  }
  return JSON.parse(raw);
}

function param_(e, key) {
  return (e.parameter && e.parameter[key]) ? String(e.parameter[key]).trim() : '';
}

function validate_(payload) {
  const name = String(payload.name || '').trim();
  const email = String(payload.email || '').trim();
  const hotel = String(payload.hotel || '').trim();
  const date = String(payload.date || '').trim();
  if (!name) return 'お名前を入力してください';
  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return 'メールアドレスを正しく入力してください';
  if (!hotel) return '宿の名前を入力してください';
  if (!date) return '宿泊日を入力してください';
  payload.name = name;
  payload.email = email;
  payload.hotel = hotel;
  payload.date = date;
  payload.guests = String(payload.guests || '').trim();
  payload.message = String(payload.message || '').trim();
  payload.source = String(payload.source || 'lp').trim();
  return '';
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function json_(body) {
  return ContentService
    .createTextOutput(JSON.stringify(body))
    .setMimeType(ContentService.MimeType.JSON);
}
