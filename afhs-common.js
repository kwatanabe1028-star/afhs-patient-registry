/**
 * AFHS 共通 — localStorage / GAS fetch / 定数
 * ポータル・registry・role スタブから読み込む
 */
(function (global) {
  const LS = {
    gasUrl: 'afhs_gasUrl',
    apiToken: 'afhs_apiToken',
    testMode: 'afhs_testMode',
    preferredRole: 'afhs_preferredRole',
  };

  const CONFIG = {
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

  function isTestMode() {
    return ls('testMode') === '1';
  }

  function hasToken() {
    return !!(ls('apiToken') || '').trim();
  }

  function hasGas() {
    return !!(ls('gasUrl') || '').trim();
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
    if (msg === 'unauthorized') return 'アクセストークンが違います。設定を確認してください';
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
    const url = ls('gasUrl');
    const token = ls('apiToken') || '';
    if (!url || isTestMode()) return { ok: false, reason: 'unset' };
    const json = await gasFetch(`${url}?action=getStats&token=${encodeURIComponent(token)}`);
    if (json.status !== 'ok') return { ok: false, reason: json.message || 'error' };
    return { ok: true, stats: normalizeStats(json.stats) };
  }

  async function listNotes(opts) {
    opts = opts || {};
    const url = ls('gasUrl');
    const token = ls('apiToken') || '';
    if (!url || isTestMode()) return { ok: false, reason: 'unset', notes: [] };
    const q = new URLSearchParams({
      action: 'listNotes',
      token: token,
      role: opts.role || 'すべて',
      type: opts.type || 'すべて',
      limit: String(opts.limit || 40),
    });
    const json = await gasFetch(`${url}?${q.toString()}`);
    if (json.status !== 'ok') return { ok: false, reason: json.message || 'error', notes: [] };
    return { ok: true, notes: json.notes || [] };
  }

  async function postNote(note) {
    const url = ls('gasUrl');
    const token = ls('apiToken') || '';
    if (!url) throw new Error('GAS URL 未設定');
    if (isTestMode()) return { status: 'ok', test: true };
    return gasPost(url, {
      action: 'postNote',
      token: token,
      role: note.role,
      type: note.type,
      body: note.body,
    });
  }

  function openSheetWithTokenGate() {
    if (!hasToken()) {
      alert('実施台帳を開くには、設定でアクセストークンを入力してください。');
      return false;
    }
    window.open(CONFIG.sheetUrl, '_blank', 'noopener,noreferrer');
    return true;
  }

  function applyUrlParams() {
    const p = new URLSearchParams(location.search);
    if (p.get('gas')) ls('gasUrl', p.get('gas'));
    if (p.get('token')) ls('apiToken', p.get('token'));
  }

  global.AFHS = {
    LS,
    CONFIG,
    ls,
    isTestMode,
    hasToken,
    hasGas,
    gasFetch,
    gasPost,
    friendlyErrorMessage,
    normalizeStats,
    loadStats,
    listNotes,
    postNote,
    openSheetWithTokenGate,
    applyUrlParams,
  };
})(window);
