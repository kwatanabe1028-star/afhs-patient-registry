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
 * - 現場メモ … ポータルのメモ投稿（自動作成）
 */

const SHEET_NAME = 'AFHS実施台帳';
const NOTES_SHEET = '現場メモ';

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

const NOTE_HEADERS = [
  '日時',   // A
  '職種',   // B
  '種別',   // C
  '本文',   // D
  'ピン',   // E
];

const NOTE_TYPES = ['引き継ぎ', '変更希望', '問題点', 'イベント', '一言', '重要連絡'];
const NOTE_ROLES = ['医師', 'Ns', '薬剤', '栄養', 'リハ'];

// ── POST 受信 ──────────────────────────────────────
function doPost(e) {
  const rawBody = e && e.postData && e.postData.contents;
  try {
    const data = JSON.parse(rawBody);
    if (!isAuthorized_(data.token)) {
      return jsonResponse({ status: 'error', message: 'unauthorized' });
    }

    if (data.action === 'postNote') {
      return handlePostNote_(data);
    }

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

function missingFields_(data) {
  const required = ['date', 'patientId', 'hcuDay', 'department', 'diagnosis', 'sessionType', 'sessionNumber'];
  return required.filter(key => {
    const v = data[key];
    return v === undefined || v === null || String(v).trim() === '';
  });
}

function handlePostNote_(data) {
  const role = String(data.role || '').trim();
  const type = String(data.type || '').trim();
  const body = String(data.body || '').trim();
  if (!NOTE_ROLES.includes(role)) {
    return jsonResponse({ status: 'error', message: '職種が不正です' });
  }
  if (!NOTE_TYPES.includes(type)) {
    return jsonResponse({ status: 'error', message: '種別が不正です' });
  }
  if (!body) {
    return jsonResponse({ status: 'error', message: '本文が空です' });
  }
  if (body.length > 500) {
    return jsonResponse({ status: 'error', message: '本文が長すぎます（500字以内）' });
  }
  const pinned = type === '重要連絡';
  const sheet = getOrCreateNotesSheet_();
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
  const targetRow = getNextDataRow_(sheet);
  sheet.getRange(targetRow, 1, 1, 5).setValues([[now, role, type, body, pinned ? 'true' : '']]);
  return jsonResponse({ status: 'ok' });
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
      return jsonResponse({ status: 'ok', stats: getStats_() });
    } catch (err) {
      return jsonResponse({ status: 'error', message: err.message });
    }
  }

  if (action === 'listNotes') {
    if (!isAuthorized_(e.parameter.token)) {
      return jsonResponse({ status: 'error', message: 'unauthorized' });
    }
    try {
      const role = e.parameter.role || '';
      const type = e.parameter.type || '';
      const limit = Math.min(Number(e.parameter.limit) || 50, 100);
      return jsonResponse({ status: 'ok', notes: listNotes_(role, type, limit) });
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

// ── 集計（今月＋累計） ──────────────────────────────
function getStats_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const now = new Date();
  const monthLabel = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM');
  if (!sheet) {
    const empty = { total: 0, initial: 0, repeat: 0 };
    return {
      month: Object.assign({ label: monthLabel }, empty),
      allTime: empty,
      total: 0,
      initial: 0,
      repeat: 0,
      monthLabel: monthLabel,
    };
  }

  const y = now.getFullYear();
  const m = now.getMonth();
  const monthStart = new Date(y, m, 1);
  const monthEnd   = new Date(y, m + 1, 1);

  const rows = sheet.getDataRange().getValues();
  const month = { total: 0, initial: 0, repeat: 0, label: monthLabel };
  const allTime = { total: 0, initial: 0, repeat: 0 };

  for (let r = 1; r < rows.length; r++) {
    const d = new Date(rows[r][0]);
    if (isNaN(d)) continue;
    const classified = classifySession_(rows[r][5], rows[r][6]);
    allTime.total++;
    if (classified === 'repeat') allTime.repeat++;
    else allTime.initial++;

    if (d >= monthStart && d < monthEnd) {
      month.total++;
      if (classified === 'repeat') month.repeat++;
      else month.initial++;
    }
  }

  // 後方互換: 旧 flat 形（今月）も返す
  return {
    month: month,
    allTime: allTime,
    total: month.total,
    initial: month.initial,
    repeat: month.repeat,
    monthLabel: monthLabel,
  };
}

function classifySession_(sessionType, sessionNum) {
  const type = String(sessionType);
  const num = Number(sessionNum) || 0;
  if (type === '再カンファ') return 'repeat';
  if (type === '初回カンファ') return 'initial';
  if (num > 1) return 'repeat';
  return 'initial';
}

// 旧名互換
function getMonthlyStats() {
  const s = getStats_();
  return { total: s.total, initial: s.initial, repeat: s.repeat, month: s.monthLabel };
}

// ── 現場メモ ───────────────────────────────────────
function listNotes_(roleFilter, typeFilter, limit) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOTES_SHEET);
  if (!sheet) return [];

  const rows = sheet.getDataRange().getValues();
  const notes = [];
  for (let r = rows.length - 1; r >= 1; r--) {
    const role = String(rows[r][1] || '');
    const type = String(rows[r][2] || '');
    const body = String(rows[r][3] || '');
    if (!body) continue;
    if (roleFilter && roleFilter !== 'すべて' && role !== roleFilter) continue;
    if (typeFilter && typeFilter !== 'すべて' && type !== typeFilter) continue;
    notes.push({
      at: formatNoteAt_(rows[r][0]),
      role: role,
      type: type,
      body: body,
      pinned: String(rows[r][4]) === 'true' || type === '重要連絡',
    });
    if (notes.length >= limit) break;
  }

  notes.sort(function (a, b) {
    if (a.pinned !== b.pinned) return a.pinned ? -1 : 1;
    return String(b.at).localeCompare(String(a.at));
  });
  return notes;
}

function formatNoteAt_(val) {
  if (!val) return '';
  if (val instanceof Date && !isNaN(val)) {
    return Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
  }
  return String(val);
}

function getOrCreateNotesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(NOTES_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(NOTES_SHEET);
    sheet.appendRow(NOTE_HEADERS);
    const headerRange = sheet.getRange(1, 1, 1, NOTE_HEADERS.length);
    headerRange.setBackground('#286858');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 140);
    sheet.setColumnWidth(4, 360);
  }
  return sheet;
}

// ── ヘルパー ───────────────────────────────────────
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
