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
  try {
    const data = JSON.parse(e.postData.contents);
    if (!isAuthorized_(data.token)) {
      return jsonResponse({ status: 'error', message: 'unauthorized' });
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
    sheet.appendRow(row);
    return jsonResponse({ status: 'ok' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
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
    return { total: 0, initial: 0, repeat: 0, outOfWindow: 0 };
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
  let outOfWindow = 0;

  for (let r = 1; r < rows.length; r++) {
    const d = new Date(rows[r][0]);
    if (isNaN(d) || d < monthStart || d >= monthEnd) continue;
    total++;
    const sessionNum = Number(rows[r][6]) || 0;
    const hcuDay = Number(rows[r][2]) || 0;
    const sessionType = String(rows[r][5]);
    if (sessionNum === 1 || sessionType === '初回カンファ') initial++;
    if (sessionNum > 1 || sessionType === '再カンファ') repeat++;
    if ((sessionNum === 1 || sessionType === '初回カンファ') && (hcuDay < 2 || hcuDay > 4)) {
      outOfWindow++;
    }
  }

  return { total, initial, repeat, outOfWindow, month: Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM') };
}

// ── ヘルパー ───────────────────────────────────────
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
