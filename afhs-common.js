/**
 * AFHS 共通 — localStorage / GAS fetch / 定数
 * ポータル・registry・role スタブから読み込む
 *
 * GAS URL はコードに固定。変更時はここを直して push する。
 */
(function (global) {
  const LS = {
    testMode: 'afhs_testMode',
    preferredRole: 'afhs_preferredRole',
  };

  const CONFIG = {
    gasUrl: 'https://script.google.com/macros/s/AKfycbzJYpa8OyI4hNQi-ag16bIGKl4whbUhG071TJ0npKtsTkZMcUcgjyt8M9bHnsNWWog2fA/exec',
    notionUrl: 'https://www.notion.so/00_AFHS-3ae7d3417e898032b4c1e916af530fd8',
    sheetUrl: 'https://docs.google.com/spreadsheets/d/1b9Olf6HU8kou32DEvaU5bIikI7h2fxASy0IVsSnQbds/edit?usp=sharing',
    roles: ['医師', 'Ns', '薬剤', '栄養', 'リハ'],
    noteTypes: ['引き継ぎ', '変更希望', '問題点', 'イベント', '一言', '重要連絡'],
    roleFiles: {
      医師: 'role-doctor.html',
      Ns: 'role-ns.html',
      薬剤: 'role-pharm.html',
      栄養: 'role-nutrition.html',
      リハ: 'role-rehab.html',
    },
  };

  function ls(key, val) {
    const k = LS[key] || key;
    if (val === undefined) return localStorage.getItem(k) || '';
    localStorage.setItem(k, val);
  }

  function gasUrl() {
    return CONFIG.gasUrl;
  }

  function isTestMode() {
    const v = ls('testMode');
    return v === '1' || v === 'true';
  }

  function hasGas() {
    return !!gasUrl();
  }

  async function gasFetch(url) {
    const res = await fetch(url);
    return res.json();
  }

  async function gasPost(url, payload) {
    const res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload),
      redirect: 'follow',
    });
    return res.json();
  }

  function friendlyErrorMessage(msg) {
    if (!msg) return 'エラーが発生しました';
    return String(msg);
  }

  function normalizeStats(raw) {
    if (!raw) return null;
    const month = raw.month && typeof raw.month === 'object'
      ? raw.month
      : { total: raw.total || 0, initial: raw.initial || 0, repeat: raw.repeat || 0, label: raw.monthLabel || raw.month || '' };
    const allTime = raw.allTime || { total: 0, initial: 0, repeat: 0 };
    return { month, allTime };
  }

  async function loadStats() {
    const url = gasUrl();
    if (!url || isTestMode()) return { ok: false, reason: 'unset' };
    const json = await gasFetch(`${url}?action=getStats`);
    if (json.status !== 'ok') return { ok: false, reason: json.message || 'error' };
    return { ok: true, stats: normalizeStats(json.stats) };
  }

  async function listNotes(opts) {
    opts = opts || {};
    const url = gasUrl();
    if (!url || isTestMode()) return { ok: false, reason: 'unset', notes: [] };
    const q = new URLSearchParams({
      action: 'listNotes',
      role: opts.role || 'すべて',
      type: opts.type || 'すべて',
      limit: String(opts.limit || 40),
    });
    const json = await gasFetch(`${url}?${q.toString()}`);
    if (json.status !== 'ok') return { ok: false, reason: json.message || 'error', notes: [] };
    return { ok: true, notes: json.notes || [] };
  }

  async function postNote(note) {
    const url = gasUrl();
    if (!url) throw new Error('GAS URL 未設定');
    if (isTestMode()) return { status: 'ok', test: true };
    return gasPost(url, {
      action: 'postNote',
      role: note.role,
      type: note.type,
      body: note.body,
    });
  }

  function openSheet() {
    window.open(CONFIG.sheetUrl, '_blank', 'noopener,noreferrer');
    return true;
  }

  /** @deprecated トークンゲート廃止。openSheet と同じ */
  function openSheetWithTokenGate() {
    return openSheet();
  }

  function applyUrlParams() {
    // 旧 ?gas= / ?token= は無視（URL は CONFIG 固定）
  }

  global.AFHS = {
    LS,
    CONFIG,
    ls,
    gasUrl,
    isTestMode,
    hasGas,
    gasFetch,
    gasPost,
    friendlyErrorMessage,
    normalizeStats,
    loadStats,
    listNotes,
    postNote,
    openSheet,
    openSheetWithTokenGate,
    applyUrlParams,
  };
})(window);
