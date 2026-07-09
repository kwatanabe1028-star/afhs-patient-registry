/**
 * AFHS 患者集計 — Google Apps Script
 *
 * 【セットアップ】
 * 1. Google スプレッドシートを新規作成
 * 2. 拡張機能 → Apps Script にこのコードを貼り付けて保存
 * 3. プロジェクトの設定 → スクリプト プロパティ → `API_TOKEN` に
 *    十分ランダムな文字列を設定（このファイルは公開リポジトリに
 *    コミットされるため、トークンをコードに直書きしないこと）
 * 4. デプロイ → 新しいデプロイ → 種類:「ウェブアプリ」
 *    - 実行ユーザー: 自分
 *    - アクセスできるユーザー: 全員
 * 5. デプロイ URL とスクリプト プロパティに設定したトークンを
 *    GitHub Pages アプリの設定画面に入力
 *
 * 【シート】
 * - AFHS実施台帳 … 送信データ（自動作成）
 */

const SHEET_NAME = 'AFHS実施台帳';

const HEADERS = [
  '実施日',         // A
  '患者ID',         // B
  'HCU入室日数',    // C
  '診療科',         // D
  '疾患',           // E
  '実施区分',       // F
  '回数',           // G
  'メモ',           // H
];

// ── POST 受信 ──────────────────────────────────────
function doPost(e) {
  const rawBody = e && e.postData && e.postData.contents;
  try {
    const data = JSON.parse(rawBody);
    if (!isAuthorized_(data.token)) {
      return jsonResponse({ status: 'error', message: 'unauthorized' });
    }
    // 歯抜けの行が黙って作られるのを防ぐ。原因（同時送信の衝突や通信経路での
    // 欠落等）が何であれ、必須項目が欠けた状態では書き込まずエラーを返す。
    const missing = missingFields_(data);
    if (missing.length > 0) {
      console.log('doPost rejected: missing=' + missing.join(',') + ' body=' + rawBody);
      return jsonResponse({ status: 'error', message: '必須項目が届いていません: ' + missing.join('、') });
    }
    const sheet = getOrCreateSheet();
    const row = [
      toSlashDate_(data.date),
      data.patientId      || '',
      Number(data.hcuDay) || '',
      data.department     || '',
      data.diagnosis      || '',
      data.sessionType    || '',
      Number(data.sessionNumber) || '',
      data.memo           || '',
    ];
    const targetRow = getNextDataRow_(sheet);
    sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
    return jsonResponse({ status: 'ok' });
  } catch (err) {
    console.log('doPost error: ' + err.message + ' body=' + rawBody);
    return jsonResponse({ status: 'error', message: err.message });
  }
}

// 必須項目（メモ以外）が欠けていないか確認する。フロント側の
// validatePayload と同じ項目セットを、サーバー側でも独立にチェックする。
function missingFields_(data) {
  const required = ['date', 'patientId', 'hcuDay', 'department', 'diagnosis', 'sessionType', 'sessionNumber'];
  return required.filter(key => {
    const v = data[key];
    return v === undefined || v === null || String(v).trim() === '';
  });
}

// ── GET ────────────────────────────────────────────
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  const patientId = e && e.parameter && e.parameter.patientId;

  if (action === 'getPatient' && patientId) {
    if (!isAuthorized_(e.parameter.token)) {
      return jsonResponse({ status: 'error', message: 'unauthorized' });
    }
    try {
      return jsonResponse({ status: 'ok', patient: findPatientHistory(patientId) });
    } catch (err) {
      return jsonResponse({ status: 'error', message: err.message });
    }
  }

  if (action === 'getStats') {
    if (!isAuthorized_(e.parameter.token)) {
      return jsonResponse({ status: 'error', message: 'unauthorized' });
    }
    try {
      return jsonResponse({ status: 'ok', stats: getMonthlyStats() });
    } catch (err) {
      return jsonResponse({ status: 'error', message: err.message });
    }
  }

  return ContentService
    .createTextOutput('AFHS 患者集計 API は正常に動作しています。')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ── 患者履歴（プリフィル用） ────────────────────────
function findPatientHistory(patientId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    return { found: false, maxSession: 0, department: '', diagnosis: '', records: [] };
  }

  const rows = sheet.getDataRange().getValues();
  const id = String(patientId).trim();
  let maxSession = 0;
  let department = '';
  let diagnosis = '';
  const records = [];

  for (let r = 1; r < rows.length; r++) {
    if (String(rows[r][1]).trim() !== id) continue;
    const sessionNum = Number(rows[r][6]) || 0;
    if (sessionNum > maxSession) maxSession = sessionNum;
    department = rows[r][3] || department;
    diagnosis  = rows[r][4] || diagnosis;
    records.push({
      date: formatDate(rows[r][0]),
      hcuDay: rows[r][2],
      sessionType: rows[r][5],
      sessionNumber: sessionNum,
    });
  }

  return {
    found: records.length > 0,
    maxSession,
    suggestedSession: maxSession + 1,
    department,
    diagnosis,
    records: records.slice(-5),
  };
}

// ── 月次サマリー（簡易） ────────────────────────────
function getMonthlyStats() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    return { total: 0, initial: 0, repeat: 0 };
  }

  const now = new Date();
  const y = now.getFullYear();
  const m = now.getMonth();
  const monthStart = new Date(y, m, 1);
  const monthEnd   = new Date(y, m + 1, 1);

  const rows = sheet.getDataRange().getValues();
  let total = 0;
  let initial = 0;
  let repeat = 0;

  for (let r = 1; r < rows.length; r++) {
    const d = new Date(rows[r][0]);
    if (isNaN(d) || d < monthStart || d >= monthEnd) continue;
    total++;
    const sessionNum = Number(rows[r][6]) || 0;
    const sessionType = String(rows[r][5]);
    // 実施区分を優先し、初回・再カンファを排他的に分類する（二重カウント防止）
    if (sessionType === '再カンファ') {
      repeat++;
    } else if (sessionType === '初回カンファ') {
      initial++;
    } else if (sessionNum > 1) {
      repeat++;
    } else {
      initial++;
    }
  }

  return { total, initial, repeat, month: Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM') };
}

// ── ヘルパー ───────────────────────────────────────
// 実施日（A列）だけを見て「本当の最終行」を判定する。
// I列などに数式（ARRAYFORMULA等）があると、空文字の見かけ上の内容に
// appendRow/getLastRowがだまされて、遠く離れた行に書き込んでしまうため。
function getNextDataRow_(sheet) {
  const colA = sheet.getRange(1, 1, sheet.getMaxRows(), 1).getValues();
  for (let r = colA.length - 1; r >= 1; r--) {
    if (colA[r][0] !== '') return r + 2;
  }
  return 2;
}

function toSlashDate_(ymd) {
  return ymd ? String(ymd).replace(/-/g, '/') : '';
}

function isAuthorized_(token) {
  const expected = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
  return !!expected && token === expected;
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setBackground('#286858');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 100);
    sheet.setColumnWidth(8, 240);
  }
  return sheet;
}

function formatDate(val) {
  if (!val) return '';
  const d = new Date(val);
  if (isNaN(d)) return String(val);
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
