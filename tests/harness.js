'use strict';
/* Test harness: loads bingAutoSearch.js into a vm sandbox with stubbed
 * browser APIs, exposes its internals for unit tests.
 * Run: node tests/test.js   (from repo root)
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const SRC = fs.readFileSync(path.join(__dirname, '..', 'bingAutoSearch.js'), 'utf8');

const EXPORT_HOOK = `
;globalThis.__BAS__ = {
  getConfig, saveConfig, getStorageData, getToday, canRunSearch,
  claimTask, releaseTask, heartbeatTask, checkAndResetDay,
  checkScheduledExecution, startSearch, startSearchLoop, stopTimer,
  performSearch, executeSearch, onTaskCompleted, toggleScript,
  getRandomKeyword, getRandomKeywordFromPool, getEnWordKeyword,
  getUniqueKeywordFromPool, getComboKey, isComboUsed, markComboUsed, escapeHtml,
  removeDuplicateWords, filterDuplicateFixes,
  mergeAndDeduplicateKeywords, mergeAndDeduplicateFixes,
  resetComboTracking, clearUsedKeywords, addUsedFullKeyword, isFullKeywordUsed,
  getSearchHistory, addSearchHistory, updateSearchHistoryUI,
  updateStatusAfterInit, updateStatus, updateCountdownUI, updateUI, updateStatusBadge,
  getBingPageType, isMobile, getRandomInterval,
  cleanCount, doAutoScroll, startScrollLoop, haltTask, setBtn, showComplete,
  loadExternalKeywords, loadPanelKeywords, init,
  getTabTaskStatus, setTabTaskStatus, isTaskRunning,
  STATUS_PAUSED, STATUS_RUNNING, STATUS_RESTING, CONFIG,
  _state: () => ({ taskStatus, timerActive, timerStart, timerInterval, nextExecuteTime,
    lastSeenDate, tabId, keywordsPool: [...keywordsPool], keywordFixPool: [...keywordFixPool],
    enWordFixPool: [...enWordFixPool], bingNewsKeywords: [...bingNewsKeywords],
    usedKeywordsToday: [...usedKeywordsToday], usedFullKeywords: [...usedFullKeywords] }),
  _set: (p) => {
    if ('taskStatus' in p) setTabTaskStatus(p.taskStatus);
    if ('lastSeenDate' in p) { lastSeenDate = p.lastSeenDate; try { sessionStorage.setItem('bing_last_seen', p.lastSeenDate); } catch (e) {} }
    if ('tabId' in p) { tabId = p.tabId; try { sessionStorage.setItem('bing_tab_id', p.tabId); } catch (e) {} }
    if ('bingNewsKeywords' in p) bingNewsKeywords = p.bingNewsKeywords;
    if ('keywordsPool' in p) keywordsPool = p.keywordsPool;
    if ('keywordFixPool' in p) keywordFixPool = p.keywordFixPool;
    if ('enWordFixPool' in p) enWordFixPool = p.enWordFixPool;
  }
};
`;

function buildSource() {
  const marker = 'if (/(^|\\.)bing\\.com$/i.test(window.location.hostname)) {';
  const idx = SRC.lastIndexOf(marker);
  if (idx < 0) throw new Error('footer marker not found - userscript structure changed');
  return SRC.slice(0, idx) + EXPORT_HOOK + '\n})();\n';
}

function mulberry32(seed) {
  let a = seed >>> 0;
  return function () {
    a |= 0; a = (a + 0x6D2B79F5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function escHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function createEnv(opts = {}) {
  const gm = opts.sharedGM || new Map();
  const ls = opts.sharedLS || new Map();
  const ss = new Map();
  if (opts.seedSS) for (const [k, v] of Object.entries(opts.seedSS)) ss.set(k, v);
  const store = (m) => ({
    getItem: (k) => (m.has(String(k)) ? m.get(String(k)) : null),
    setItem: (k, v) => { m.set(String(k), String(v)); },
    removeItem: (k) => { m.delete(String(k)); },
    clear: () => { m.clear(); },
  });

  let mockNow = opts.now ?? Date.UTC(2026, 8, 17, 16, 0, 0); // 2026-09-18 00:00 Asia/Taipei
  const RealDate = Date;
  class FakeDate extends RealDate {
    constructor(...a) { super(...(a.length ? a : [mockNow])); }
    static now() { return mockNow; }
  }
  const setMockNow = (ms) => { mockNow = ms; };
  const getMockNow = () => mockNow;

  // ---- fake timers (deterministic) ----
  let seq = 1;
  const timers = new Map();
  const setTimeoutFn = (cb, ms, ...a) => { const id = seq++; timers.set(id, { cb, at: mockNow + (ms || 0), ms: ms || 0, args: a, repeat: false }); return id; };
  const clearTimeoutFn = (id) => { timers.delete(id); };
  const setIntervalFn = (cb, ms, ...a) => { const id = seq++; timers.set(id, { cb, at: mockNow + (ms || 0), ms: ms || 0, args: a, repeat: true }); return id; };
  const clearIntervalFn = (id) => { timers.delete(id); };
  async function advance(ms) {
    const end = mockNow + ms;
    let guard = 0;
    for (;;) {
      let nextId = -1, nextAt = Infinity;
      for (const [id, t] of timers) if (t.at <= end && t.at < nextAt) { nextAt = t.at; nextId = id; }
      if (nextId < 0 || ++guard > 20000) break;
      const t = timers.get(nextId);
      mockNow = t.at;
      if (!t.repeat) timers.delete(nextId); else t.at = mockNow + t.ms;
      t.cb(...t.args);
      // mirror the real event loop: drain microtasks after each macrotask
      for (let i = 0; i < 20; i++) await Promise.resolve();
    }
    mockNow = end;
  }
  const pendingTimers = () => [...timers.values()].map((t) => t.ms);

  // ---- location ----
  const hrefWrites = [];
  let curHref = opts.url || 'https://www.bing.com/';
  const locationObj = {};
  Object.defineProperty(locationObj, 'href', {
    enumerable: true, configurable: true,
    get: () => curHref,
    set: (v) => { hrefWrites.push(String(v)); curHref = String(v); },
  });
  for (const k of ['origin', 'pathname', 'search', 'hostname']) {
    Object.defineProperty(locationObj, k, { enumerable: true, configurable: true, get: () => new URL(curHref)[k] });
  }
  const setUrl = (v) => { curHref = String(v); hrefWrites.length = 0; };

  // ---- DOM stubs ----
  const nullIds = new Set();
  const elements = new Map();
  const qsMap = new Map();
  const inserted = [];
  const hooks = {
    keywordsJson: null, newsHtml: '<html></html>', parserTitles: [],
    formSubmit: null, elClick: null, confirm: true,
  };
  function makeEl(id) {
    const e = {
      id, value: '', className: '', style: {}, title: '',
      _text: '', _html: null, _attrs: {}, _cls: new Set(), _children: new Map(), _listeners: {},
      onclick: null, onmousedown: null, onkeydown: null,
      focus() {}, blur() {}, dispatchEvent() { return true; },
      addEventListener(t, fn) { (this._listeners[t] = this._listeners[t] || []).push(fn); },
      click() { if (hooks.elClick) hooks.elClick(id); },
      submit() {
        if (hooks.formSubmit) { hooks.formSubmit(id); return; }
        const inp = elements.get('sb_form_q');
        const loc = new URL(curHref);
        hrefWrites.push(loc.origin + '/search?q=' + encodeURIComponent(inp ? inp.value : ''));
        curHref = hrefWrites[hrefWrites.length - 1];
      },
      querySelector(sel) { if (!e._children.has(sel)) e._children.set(sel, makeEl(sel)); return e._children.get(sel); },
      querySelectorAll() { return []; },
      setAttribute(k, v) { this._attrs[k] = String(v); },
      getAttribute(k) { return k in this._attrs ? this._attrs[k] : null; },
    };
    Object.defineProperty(e, 'textContent', {
      enumerable: true, configurable: true,
      get() { return this._text; }, set(v) { this._text = String(v); },
    });
    Object.defineProperty(e, 'innerHTML', {
      enumerable: true, configurable: true,
      get() { return this._html !== null ? this._html : escHtml(this._text); },
      set(v) { this._html = String(v); },
    });
    e.classList = {
      add: (...c) => c.forEach((x) => e._cls.add(x)),
      remove: (...c) => c.forEach((x) => e._cls.delete(x)),
      toggle: (c) => { if (e._cls.has(c)) e._cls.delete(c); else e._cls.add(c); return e._cls.has(c); },
      contains: (c) => e._cls.has(c),
    };
    return e;
  }
  const listeners = {};
  const documentStub = {
    readyState: 'complete', hidden: false,
    body: { insertAdjacentHTML: (pos, html) => { inserted.push(html); }, scrollHeight: 2500 },
    documentElement: { scrollHeight: 2500 },
    getElementById: (id) => {
      if (nullIds.has(id)) return null;
      if (!elements.has(id)) elements.set(id, makeEl(id));
      return elements.get(id);
    },
    querySelector: (sel) => (qsMap.has(sel) ? qsMap.get(sel) : null),
    querySelectorAll: () => [],
    createElement: () => makeEl('x'),
    addEventListener: (t, fn) => { (listeners[t] = listeners[t] || []).push(fn); },
  };
  const windowStub = {
    location: locationObj, innerWidth: 1280, innerHeight: 800,
    addEventListener: (t, fn) => { (listeners['win:' + t] = listeners['win:' + t] || []).push(fn); },
    history: { pushState() {}, replaceState() {} },
    scrollTo() {},
  };

  function setLoggedIn() {} // login gate removed: searches run regardless of login state

  // ---- network stubs ----
  const fetchCalls = [];
  const fetchStub = async (url) => {
    fetchCalls.push(String(url));
    if (String(url).includes('raw.githubusercontent')) {
      return { ok: true, status: 200, json: async () => hooks.keywordsJson || { keywords: ['外部關鍵字一', '外部關鍵字二'], keywordFix: ['外修'], enWordFix: ['外英'] } };
    }
    return { ok: true, status: 200, text: async () => hooks.newsHtml };
  };
  class DOMParserStub {
    parseFromString() {
      return { querySelectorAll: () => hooks.parserTitles.map((t) => ({ getAttribute: () => t, textContent: t })) };
    }
  }
  class EventStub {
    constructor(type, init) { this.type = type; this.bubbles = !!(init && init.bubbles); }
  }
  const AbortCtrl = typeof AbortController !== 'undefined' ? AbortController : class { constructor() { this.signal = {}; } abort() {} };

  const sandbox = {
    console,
    localStorage: store(ls), sessionStorage: store(ss),
    document: documentStub, window: windowStub,
    navigator: { userAgent: opts.ua || 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/126' },
    GM_addStyle: (css) => { inserted.push('STYLE:' + css.length); },
    GM_getValue: (k) => (gm.has(k) ? gm.get(k) : undefined),
    GM_setValue: (k, v) => { gm.set(k, v); },
    fetch: fetchStub, AbortController: AbortCtrl, DOMParser: DOMParserStub,
    Event: EventStub, confirm: () => hooks.confirm,
    setTimeout: setTimeoutFn, clearTimeout: clearTimeoutFn,
    setInterval: setIntervalFn, clearInterval: clearIntervalFn,
    URL, Intl, JSON, Date: FakeDate,
  };
  const ctx = vm.createContext(sandbox);
  const seedRandom = (s) => { sandbox.__seedRandom = mulberry32(s); vm.runInContext('Math.random = globalThis.__seedRandom;', ctx); };
  seedRandom(opts.seed ?? 12345);
  vm.runInContext(buildSource(), ctx);
  const api = ctx.__BAS__;

  const flush = async (n = 50) => { for (let i = 0; i < n; i++) await Promise.resolve(); };
  const gmJson = (k) => { const v = gm.get(k); return v === undefined ? null : JSON.parse(v); };

  return {
    api, advance, flush, pendingTimers, setMockNow, getMockNow,
    setUrl, hrefWrites, gm, gmJson, ls, ss, nullIds, qsMap, elements, inserted,
    hooks, fetchCalls, listeners, setLoggedIn,
    reseed: (s) => seedRandom(s),
  };
}

module.exports = { createEnv };
