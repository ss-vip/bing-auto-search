// ==UserScript==
// @name         Bing Auto Search
// @version      2026081901
// @description  無人值守 Bing 自動隨機搜尋
// @author       Hank
// @match        https://*.bing.com/*
// @icon         https://www.google.com/s2/favicons?sz=64&domain=bing.com
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_addStyle
// @run-at       document-end
// @license      GPL-3.0
// @namespace    https://greasyfork.org/zh-TW/users/933219-tw1720
// @supportURL   https://github.com/ss-vip/bing-auto-search
// @updateURL    https://update.greasyfork.org/scripts/572057/Bing%20Auto%20Search.user.js
// @downloadURL  https://update.greasyfork.org/scripts/572057/Bing%20Auto%20Search.user.js
// ==/UserScript==
(function () {
  'use strict';
  const CONFIG = {
    max_pc: 45,
    max_ph: 35,
    min_interval: 50,
    max_interval: 120,
    timezone: 'Asia/Taipei',
    keywordsUrl: 'https://raw.githubusercontent.com/ss-vip/bing-auto-search/refs/heads/main/example.json',
    bingNewsUrl: 'https://www.bing.com/news/search?q=%e7%86%b1%e9%96%80%e5%a0%b1%e5%b0%8e&nvaug=%5bNewsVertical+Category%3d%22rt_MaxClass%22%5d',
    fixWeight: {
      none: 60,
      prefix: 20,
      suffix: 20
    },
    enFixWeight: {
      none: 60,
      prefix: 10,
      suffix: 20,
      both: 10
    },
    sourceWeight: {
      chinese: 55,
      english: 45
    },
    defaultKeywordsPool: [
    'Python 教學', 'Java 環境變數', 'Linux 常用指令', 'Docker 部署', 'React vs Vue', 'ChatGPT API 教學', 'GitHub Copilot 評測',
    'SQL 優化 技巧', '正則表達式 教學', 'C++ 指標 教學', 'Rust 入門 教學', 'Unity 遊戲開發', 'VS Code',
    'Python 爬蟲 教學', 'MacBook Pro', '必玩Steam遊戲',
    '機械鍵盤', '降噪耳機', '智慧手錶', '感冒吃什麼改善', '番茄炒蛋做法', '避免近視眼', '減肥食譜', '影集', '超商便宜攻略',
    '小資旅游攻略', '遊樂園門票優惠', '自駕旅遊', '今日金價', '美元匯率', '油價走勢'
    ],
    defaultKeywordFixPool: ['最新', '資訊', '近期', '說明', '是啥', '常見', '有啥', '最好', '最推', '超便', '很優', '推薦'],
    defaultEnWordFixPool: ['英文', '中文', '翻譯', '解釋', '意思', '造句', '定義', '用法', '例句', '解說', '範例', '簡述']
  };
  (function validateConfig() {
    if (CONFIG.max_pc < 1) CONFIG.max_pc = 1;
    if (CONFIG.max_ph < 1) CONFIG.max_ph = 1;
    const minInt = Math.max(1, parseInt(CONFIG.min_interval) || 1);
    const maxInt = Math.max(minInt, parseInt(CONFIG.max_interval) || minInt);
    CONFIG.min_interval = minInt;
    CONFIG.max_interval = maxInt;
    const totalWeight = (cfg) => (cfg.none || 0) + (cfg.prefix || 0) + (cfg.suffix || 0) + (cfg.both || 0);
    const hasInvalidWeight = (cfg) => Object.values(cfg).some(v => v < 0);
    if (totalWeight(CONFIG.fixWeight) !== 100 || hasInvalidWeight(CONFIG.fixWeight)) {
      const w = CONFIG.fixWeight;
      const none = Math.max(0, Math.min(100, w.none || 20));
      const prefix = Math.max(0, Math.min(100, w.prefix || 40));
      const suffix = 100 - none - prefix;
      CONFIG.fixWeight = { none, prefix, suffix: Math.max(0, suffix) };
    }
    if (totalWeight(CONFIG.enFixWeight) !== 100 || hasInvalidWeight(CONFIG.enFixWeight)) {
      const w = CONFIG.enFixWeight;
      const none = Math.max(0, Math.min(100, w.none || 25));
      const prefix = Math.max(0, Math.min(100, w.prefix || 25));
      const suffix = Math.max(0, Math.min(100, w.suffix || 25));
      const both = 100 - none - prefix - suffix;
      CONFIG.enFixWeight = { none, prefix, suffix, both: Math.max(0, both) };
    }
    const srcWeight = CONFIG.sourceWeight;
    const srcTotal = (srcWeight.chinese || 0) + (srcWeight.english || 0);
    if (srcTotal !== 100) {
      const chinese = Math.max(0, Math.min(100, srcWeight.chinese || 50));
      CONFIG.sourceWeight = { chinese, english: 100 - chinese };
    }
  })();
  let keywordsPool = CONFIG.defaultKeywordsPool;
  let keywordFixPool = CONFIG.defaultKeywordFixPool;
  let enWordFixPool = CONFIG.defaultEnWordFixPool;
  let bingNewsKeywords = [];
  let usedPrefixSuffixCombos = new Set();
  let usedKeywordsToday = new Set();
  let usedFullKeywords = new Set();
  const MAX_RECENT_HISTORY = 50;
  const MAX_FULL_KEYWORDS = 200;
  function mergeAndDeduplicateKeywords(externalKeywords, defaultKeywords) {
    const combined = [...new Set([...defaultKeywords, ...externalKeywords])];
    return combined.filter(k => k && k.trim().length > 0);
  }
  function mergeAndDeduplicateFixes(externalFixes, defaultFixes) {
    return [...new Set([...defaultFixes, ...externalFixes])].filter(f => f && f.trim().length > 0);
  }
  function resetComboTracking() {
    usedPrefixSuffixCombos.clear();
  }
  function getComboKey(prefix, suffix, baseKeyword) {
    return `${prefix || 'none'}_${suffix || 'none'}_${baseKeyword}`;
  }
  function isComboUsed(prefix, suffix, baseKeyword) {
    return usedPrefixSuffixCombos.has(getComboKey(prefix, suffix, baseKeyword));
  }
  function markComboUsed(prefix, suffix, baseKeyword) {
    usedPrefixSuffixCombos.add(getComboKey(prefix, suffix, baseKeyword));
  }
  function getUniqueKeywordFromPool() {
    const available = keywordsPool.filter(k => !usedKeywordsToday.has(k));
    let keyword;
    if (available.length > 0) {
      keyword = available[Math.floor(Math.random() * available.length)];
    } else {
      usedKeywordsToday.clear();
      keyword = keywordsPool[Math.floor(Math.random() * keywordsPool.length)];
    }
    usedKeywordsToday.add(keyword);
    if (usedKeywordsToday.size > MAX_RECENT_HISTORY) {
      const first = usedKeywordsToday.values().next().value;
      usedKeywordsToday.delete(first);
    }
    return keyword;
  }
  function clearUsedKeywords() {
    usedKeywordsToday.clear();
    usedFullKeywords.clear();
  }
  function addUsedFullKeyword(keyword) {
    usedFullKeywords.add(keyword);
    if (usedFullKeywords.size > MAX_FULL_KEYWORDS) {
      const arr = Array.from(usedFullKeywords);
      arr.shift();
      usedFullKeywords = new Set(arr);
    }
  }
  function isFullKeywordUsed(keyword) {
    return usedFullKeywords.has(keyword);
  }
  function removeDuplicateWords(keyword) {
    const words = keyword.split(/\s+/);
    const seen = new Set();
    const uniqueWords = words.filter(word => {
      if (seen.has(word)) return false;
      seen.add(word);
      return true;
    });
    return uniqueWords.join(' ');
  }
  function filterDuplicateFixes(baseKeyword, fixes) {
    const baseWords = new Set(baseKeyword.split(/\s+/));
    return fixes.filter(fix => !baseWords.has(fix));
  }
  const STORAGE_KEY = 'bingAutoSearch';
  const JOKE_API_URL = 'https://v2.jokeapi.dev/joke/Any?blacklistFlags=nsfw,religious,political,racist,sexist,explicit&type=single&amount=10';
  const KEYWORDS_CACHE_KEY = 'bing_keywords_cache';
  const TASK_STATUS_KEY = 'bing_task_status';
  const SEARCH_HISTORY_KEY = 'bing_search_history';
  const MAX_HISTORY_RECORDS = 5;
  const WAKEUP_TRIGGER_KEY = 'bing_auto_wakeup';
  const CROSSDAY_CHECK_KEY = 'bing_crossday_check';
const TASK_OWNER_KEY = 'bing_task_owner';
  const STATUS_PAUSED = 'paused';
  const STATUS_RUNNING = 'running';
  const STATUS_RESTING = 'resting';
  function getSearchHistory() {
    try {
      const data = localStorage.getItem(SEARCH_HISTORY_KEY);
      if (data) return JSON.parse(data);
    } catch (e) { }
    return [];
  }
  function addSearchHistory(keyword) {
    const history = getSearchHistory();
    const now = new Date();
    const record = {
      keyword: keyword,
      time: now.toLocaleString('zh-TW', { timeZone: CONFIG.timezone || undefined, month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' })
    };
    history.unshift(record);
    if (history.length > MAX_HISTORY_RECORDS) {
      history.pop();
    }
    try {
      localStorage.setItem(SEARCH_HISTORY_KEY, JSON.stringify(history));
    } catch (e) { }
    updateSearchHistoryUI();
  }
  function updateSearchHistoryUI() {
    const historyContainer = document.getElementById('br_history_content');
    if (!historyContainer) return;
    const history = getSearchHistory();
    if (history.length === 0) {
      historyContainer.innerHTML = '<div style="color: #999; font-size: 12px; text-align: center; padding: 8px;">尚無搜尋記錄</div>';
      return;
    }
    historyContainer.innerHTML = history.map(record => `
<div style="padding: 6px 0; border-bottom: 1px solid #f0f0f0; font-size: 12px;">
<div style="color: #333; word-break: break-all;">${escapeHtml(record.keyword)}</div>
<div style="color: #888; font-size: 11px; margin-top: 2px;">${record.time}</div>
</div>
`).join('');
  }
  function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }
  let taskStatus = STATUS_PAUSED;
  let timerStart = 0;
  let timerInterval = 0;
  let timerActive = false;
  let timerHandle = null;
  let isDragging = false;
  let dragX = 0, dragY = 0;
  let checkInterval = null;
  let nextExecuteTime = 0;
  let scrollInterval = null;
  let scrollTimeout = null;
  let tabId = sessionStorage.getItem('bing_tab_id') || Math.random().toString(36).slice(2);
  try { sessionStorage.setItem('bing_tab_id', tabId); } catch (e) { }
  function getTabTaskStatus() {
    try {
      const stored = sessionStorage.getItem(TASK_STATUS_KEY);
      if (stored && [STATUS_PAUSED, STATUS_RUNNING, STATUS_RESTING].includes(stored)) {
        return stored;
      }
    } catch (e) { }
    return null;
  }
  function setTabTaskStatus(status) {
    taskStatus = status;
    try {
      sessionStorage.setItem(TASK_STATUS_KEY, status);
    } catch (e) { }
  }
  async function loadExternalKeywords() {
    const today = getToday();
    let cacheData = null;
    try {
      const cached = localStorage.getItem(KEYWORDS_CACHE_KEY);
      if (cached) cacheData = JSON.parse(cached);
    } catch (e) {
      localStorage.removeItem(KEYWORDS_CACHE_KEY);
    }
    if (cacheData && cacheData.date === today) {
      keywordsPool = mergeAndDeduplicateKeywords(cacheData.keywords || [], CONFIG.defaultKeywordsPool);
      keywordFixPool = mergeAndDeduplicateFixes(cacheData.keywordFix || [], CONFIG.defaultKeywordFixPool);
      enWordFixPool = mergeAndDeduplicateFixes(cacheData.enWordFix || [], CONFIG.defaultEnWordFixPool);
      return true;
    }
    if (CONFIG.keywordsUrl) {
      try {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 5000);
        const response = await fetch(CONFIG.keywordsUrl, { signal: controller.signal });
        clearTimeout(timeoutId);
        if (!response.ok) throw new Error('HTTP ' + response.status);
        const data = await response.json();
        const externalKeywords = Array.isArray(data.keywords) ? data.keywords.filter(k => k && k.trim()) : [];
        const externalKeywordFix = Array.isArray(data.keywordFix) ? data.keywordFix.filter(f => f && f.trim()) : [];
        const externalEnWordFix = Array.isArray(data.enWordFix) ? data.enWordFix.filter(f => f && f.trim()) : [];
        keywordsPool = mergeAndDeduplicateKeywords(externalKeywords, CONFIG.defaultKeywordsPool);
        keywordFixPool = mergeAndDeduplicateFixes(externalKeywordFix, CONFIG.defaultKeywordFixPool);
        enWordFixPool = mergeAndDeduplicateFixes(externalEnWordFix, CONFIG.defaultEnWordFixPool);
        localStorage.setItem(KEYWORDS_CACHE_KEY, JSON.stringify({
          date: today,
          version: data.version || null,
          keywords: externalKeywords,
          keywordFix: externalKeywordFix,
          enWordFix: externalEnWordFix,
          lastFetch: Date.now()
        }));
        console.log(`[BAS] 詞彙庫已更新: ${keywordsPool.length} 組`);
        return true;
      } catch (e) {
        console.log('[BAS] 外部詞彙載入失敗，使用快取或預設');
      }
    }
    if (cacheData && (cacheData.keywords || cacheData.keywordFix || cacheData.enWordFix)) {
      keywordsPool = mergeAndDeduplicateKeywords(cacheData.keywords || [], CONFIG.defaultKeywordsPool);
      keywordFixPool = mergeAndDeduplicateFixes(cacheData.keywordFix || [], CONFIG.defaultKeywordFixPool);
    enWordFixPool = mergeAndDeduplicateFixes(cacheData.enWordFix || [], CONFIG.defaultEnWordFixPool);
    console.log(`[BAS] 使用本地快取: ${keywordsPool.length} 組`);
    return true;
  }
  keywordsPool = CONFIG.defaultKeywordsPool;
  keywordFixPool = CONFIG.defaultKeywordFixPool;
  enWordFixPool = CONFIG.defaultEnWordFixPool;
  return false;
}
  async function loadPanelKeywords() {
    if (!CONFIG.bingNewsUrl) return;
    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 5000);
      const response = await fetch(CONFIG.bingNewsUrl, {
        signal: controller.signal
      });
      clearTimeout(timeoutId);
      if (!response.ok) throw new Error('HTTP ' + response.status);
      const htmlText = await response.text();
      const parser = new DOMParser();
      const doc = parser.parseFromString(htmlText, 'text/html');
      const titles = [];
      let items = doc.querySelectorAll('.na_t.news_title');
      items.forEach(item => {
        const title = item.getAttribute('title')?.trim();
        if (title && title.length > 0) {
          titles.push(title);
        }
      });
      if (titles.length === 0) {
        const bAlgoItems = doc.querySelectorAll('.b_algo a, .b_ans a');
        bAlgoItems.forEach(item => {
          const text = item.textContent?.trim();
          if (text && text.length > 4 && text.length < 80 &&
          !text.includes('·') &&
          !text.match(/[A-Z][a-z]+\s+[A-Z][a-z]+/)) {
            titles.push(text);
          }
        });
      }
      const uniqueTitles = [...new Set(titles)].slice(0, 30);
      if (uniqueTitles.length > 0) {
        bingNewsKeywords = uniqueTitles;
      }
    } catch (e) {
      console.log('[BAS] Bing News 關鍵字載入失敗');
    }
  }
  function init() {
    resetComboTracking();
    loadExternalKeywords();
    loadPanelKeywords();
    checkAndResetDay();
    const savedStatus = getTabTaskStatus();
    if (savedStatus && savedStatus !== STATUS_PAUSED) {
      setTabTaskStatus(savedStatus);
      if (savedStatus === STATUS_RUNNING) {
        setTimeout(() => startSearch(), 1500);
      }
    } else if (!savedStatus && getConfig().autoStart === true) {
      setTimeout(() => startSearch(), 1500);
    } else {
      setTabTaskStatus(STATUS_PAUSED);
    }
    initStyles();
    initUI();
    if (taskStatus === STATUS_RUNNING && window.location.pathname.includes('/search')) {
      doAutoScroll();
    }
    startKeepAlive();
    setupCrossDayListener();
    if (document.readyState === 'complete') {
      setTimeout(() => {
        if (isTaskRunning() && window.location.pathname.includes('/search')) {
          doAutoScroll();
        }
      }, 3000);
    } else {
      window.addEventListener('load', () => {
        setTimeout(() => {
          if (isTaskRunning() && window.location.pathname.includes('/search')) {
            doAutoScroll();
          }
        }, 3000);
      });
    }
    let lastTaskStatus = taskStatus;
    setInterval(() => {
      if (lastTaskStatus !== taskStatus) {
        lastTaskStatus = taskStatus;
        updateStatusBadge(taskStatus);
      }
    }, 500);
  }
  function startKeepAlive() {
    document.addEventListener('visibilitychange', handleVisibilityChange);
    checkScheduledExecution();
    if (checkInterval) clearInterval(checkInterval);
    checkInterval = setInterval(() => {
      checkAndResetDay();
      checkScheduledExecution();
      heartbeatTask();
      updateUI();
    }, 10000);
  }
  function handleVisibilityChange() {
    if (!document.hidden) {
      checkAndResetDay();
      checkScheduledExecution();
      if (isTaskRunning() && window.location.pathname.includes('/search')) {
        doAutoScroll();
      }
    }
  }
  function checkScheduledExecution() {
    const now = Date.now();
    let scheduledTime = nextExecuteTime;
    if (scheduledTime === 0) {
      try {
        const saved = localStorage.getItem('bing_auto_schedule');
        if (saved) {
          const data = JSON.parse(saved);
          if (data.time > 0 && (now - data.timestamp) < 3600000) {
            scheduledTime = data.time;
          }
        }
      } catch (e) { }
    }
    if (!isTaskRunning() || taskStatus === STATUS_RESTING) return;
    if (scheduledTime > 0 && now >= scheduledTime) {
      nextExecuteTime = 0;
      localStorage.removeItem('bing_auto_schedule');
      performSearch();
      return;
    }
    if (isTaskRunning() && scheduledTime === 0) {
      const elapsed = Date.now() - timerStart;
      const remaining = timerInterval - elapsed;
      if (remaining > 0) {
        nextExecuteTime = now + remaining;
        saveScheduleTime(nextExecuteTime);
      }
    }
  }
  function checkAndResetDay() {
    const stored = getStorageData();
    const today = getToday();
    if (stored && stored.lastDate !== today) {
      const crossdayMark = localStorage.getItem(CROSSDAY_CHECK_KEY);
      if (crossdayMark !== today) {
        console.log('[BAS] 檢測到跨天，執行重置...');
        const newConfig = {
          date: today,
          lastDate: today,
          pc_count: 0,
          ph_count: 0,
          autoStart: true
        };
        saveConfig(newConfig);
        resetComboTracking();
        clearUsedKeywords();
        localStorage.setItem(CROSSDAY_CHECK_KEY, today);
        broadcastWakeup();
        updateStatus("跨天重置成功! 任務進行中...", "#e67e22");
        console.log('[BAS] 跨天重置完成');
      }
      if (taskStatus === STATUS_RESTING && claimTask()) {
        setTabTaskStatus(STATUS_RUNNING);
        updateStatus("腳本運行中...", "#e67e22");
        updateStatusBadge(STATUS_RUNNING);
        const btn = document.getElementById('br_toggle_btn');
        if (btn) { btn.textContent = "⏸ 暫停搜尋"; btn.className = "br_btn br_btn_stop"; }
        startSearchLoop();
        doAutoScroll();
      }
      updateUI();
    } else if (!stored) {
      localStorage.setItem(CROSSDAY_CHECK_KEY, today);
    }
  }
  function broadcastWakeup() {
    try {
      localStorage.setItem(WAKEUP_TRIGGER_KEY, JSON.stringify({
        action: 'WAKEUP',
        timestamp: Date.now()
      }));
      localStorage.removeItem(WAKEUP_TRIGGER_KEY);
    } catch (e) { }
  }
  function setupCrossDayListener() {
    window.addEventListener('storage', (e) => {
      if (e.key === WAKEUP_TRIGGER_KEY && e.newValue) {
        try {
      const data = JSON.parse(e.newValue);
      if (data.action === 'WAKEUP') {
        checkAndResetDay();
      }
        } catch (err) {}
      }
    });
  }
  function canRunSearch(config) {
    const currentPageType = getBingPageType();
    return (currentPageType === 'pc' && config.pc_count < CONFIG.max_pc) || (currentPageType === 'ph' && config.ph_count < CONFIG.max_ph);
  }
  function initStyles() {
    GM_addStyle(`#br_reward_tool{position:fixed;right:30px;bottom:30px;left:auto;top:auto;background:#fff;padding:0;border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,0.15);width:260px;z-index:9999999;cursor:default;user-select:none;border:1px solid #dcdcdc;box-sizing:border-box;text-align:left;line-height:1.5;color:#333}#br_reward_tool *{box-sizing:border-box}.br_header{position:relative;height:40px;border-top-left-radius:8px;border-top-right-radius:8px;background:#f5f5f5;border-bottom:1px solid #e0e0e0;display:flex;align-items:center;justify-content:space-between;padding:0 12px;cursor:move;width:100%}.br_title{font-size:14px;font-weight:600;color:#444}.br_date{font-size:11px;color:#888;margin-left:8px;font-weight:normal}.br_minimize-btn{border:none;background:none;cursor:pointer;font-size:20px;color:#666;padding:0;width:24px;height:24px;display:flex;align-items:center;justify-content:center}.br_minimize-btn:hover{color:#0078d4;background:#e0e0e0;border-radius:4px}.br_panel-content{padding:15px;background:#fff;border-bottom-left-radius:8px;border-bottom-right-radius:8px}.br_btn{display:block;width:100%;margin:8px 0;padding:8px 0;color:#fff;border-radius:4px;text-align:center;font-weight:600;text-decoration:none;font-size:14px;cursor:pointer;border:none;outline:none}.br_btn_start{background:#0078d4}.br_btn_start:hover{background:#005bb5}.br_btn_stop{background:#d63031}.br_btn_stop:hover{background:#c0392b}.br_btn_reset{background:#f0f0f0;color:#333 !important;border:1px solid #ccc !important;font-weight:normal !important;margin-top:10px}.br_btn_reset:hover{background:#e0e0e0}#br_reward_tool p{margin:8px 0;color:#444;font-size:13px;display:flex;justify-content:space-between;align-items:center}.br_count{font-weight:bold;color:#0078d4;font-size:14px}#br_status_text{color:#666;font-size:12px;margin-top:12px;text-align:center;display:block;background:#f9f9f9;padding:4px;border-radius:4px}#br_countdown{color:#e67e22;font-weight:bold}#br_reward_tool.br_minimized{width:50px !important;height:50px !important;padding:0 !important;background:transparent !important;box-shadow:none !important;border:none !important;right:30px !important;bottom:50px !important}#br_reward_tool.br_minimized .br_header,#br_reward_tool.br_minimized .br_panel-content{display:none !important}.br_mini-icon{width:50px;height:50px;border-radius:50%;background:#0078d4;display:flex;align-items:center;justify-content:center;color:#fff;font-size:12px;cursor:pointer;box-shadow:0 4px 12px rgba(0,0,0,0.3);font-weight:bold;border:2px solid #fff;text-align:center;line-height:1.2}.br_mini-icon:hover{background:#005bb5}#br_reward_tool:not(.br_minimized) .br_mini-icon{display:none}.br_mini-icon.running{background:#d63031}.br_live-indicator{display:inline-block;width:8px;height:8px;border-radius:50%;background:#27ae60;margin-right:6px}.br_mini-icon.paused{background:#0078d4}.br_mini-icon.resting{background:#27ae60}.br_status-badge{display:inline-block;font-size:10px;padding:2px 6px;border-radius:3px;margin-left:6px;vertical-align:middle}.br_status-badge.paused{background:#666;color:#fff}.br_status-badge.running{background:#e67e22;color:#fff}.br_status-badge.resting{background:#27ae60;color:#fff}.br_history-accordion{margin-top:12px;border:1px solid #e0e0e0;border-radius:6px;overflow:hidden}.br_history-header{display:flex;justify-content:space-between;align-items:center;padding:8px 12px;background:#f9f9f9;cursor:pointer;font-size:13px;font-weight:500;color:#444;user-select:none}.br_history-header:hover{background:#f0f0f0}.br_divider{border-top:1px solid #eee;margin:10px 0}.br_history-arrow{font-size:10px;color:#888}.br_history-header.expanded .br_history-arrow{transform:rotate(180deg)}.br_history-content{display:none;max-height:200px;overflow-y:auto;background:#fff;padding:8px 12px}.br_history-content.show{display:block}`);

  }
  function initUI() {
    const countInfo = getConfig();
    const today = getToday();
    const toolHtml = `
<div id="br_reward_tool" class="br_minimized">
<div class="br_header">
<span class="br_title"><span class="br_live-indicator"></span>隨機搜尋 <span class="br_status-badge" id="br_status_badge">暫停</span></span>
<span class="br_date">${today}</span>
<button class="br_minimize-btn">–</button>
</div>
<div class="br_panel-content">
<button id="br_toggle_btn" class="br_btn br_btn_start">▶ 開始搜尋</button>
<div class="br_divider"></div>
<p>桌面版搜尋: <span><span class="br_count" id="pc_count">${countInfo.pc_count}</span> / ${CONFIG.max_pc}</span></p>
<p>行動版搜尋: <span><span class="br_count" id="ph_count">${countInfo.ph_count}</span> / ${CONFIG.max_ph}</span></p>
<p>下一次搜尋: <span id="br_countdown">--</span></p>
<span id="br_status_text">等待開始...</span>
<button id="br_reset_btn" class="br_btn br_btn_reset">↺ 重置今日計數</button>
<div class="br_history-accordion">
<div class="br_history-header" id="br_history_header">
<span>📜 最近搜尋記錄</span>
<span class="br_history-arrow">▼</span>
</div>
<div class="br_history-content" id="br_history_content">
<div style="color: #999; font-size: 12px; text-align: center; padding: 8px;">尚無搜尋記錄</div>
</div>
</div>
</div>
<div class="br_mini-icon">Bing</div>
</div>
`;
    if (document.body) {
      document.body.insertAdjacentHTML('beforeend', toolHtml);
    } else {
      window.addEventListener('load', function () { document.body.insertAdjacentHTML('beforeend', toolHtml); }, { once: true });
    }
    setTimeout(() => {
      const toolBox = document.getElementById('br_reward_tool');
      const toggleBtn = document.getElementById('br_toggle_btn');
      const resetBtn = document.getElementById('br_reset_btn');
      if (!toolBox) return;
      toggleBtn.onclick = () => { toggleScript(); };
      resetBtn.onclick = () => { cleanCount(toolBox); };
      const minBtn = toolBox.querySelector('.br_minimize-btn');
      const miniIcon = toolBox.querySelector('.br_mini-icon');
      const header = toolBox.querySelector('.br_header');
      minBtn.onclick = (e) => {
        e.stopPropagation();
        toolBox.classList.add('br_minimized');
        toolBox.style.right = '30px'; toolBox.style.bottom = '30px'; toolBox.style.left = 'auto'; toolBox.style.top = 'auto';
      };
      miniIcon.onclick = (e) => {
        e.stopPropagation();
        toolBox.classList.remove('br_minimized');
        toolBox.style.right = '30px'; toolBox.style.bottom = '30px'; toolBox.style.left = 'auto'; toolBox.style.top = 'auto';
      };
      header.onmousedown = (e) => {
        isDragging = true;
        dragX = e.clientX - toolBox.offsetLeft;
        dragY = e.clientY - toolBox.offsetTop;
      };
      document.addEventListener('mousemove', (e) => {
        if (!isDragging) return;
        e.preventDefault();
        let l = e.clientX - dragX;
        let t = e.clientY - dragY;
        l = Math.max(0, Math.min(window.innerWidth - toolBox.offsetWidth, l));
        t = Math.max(0, Math.min(window.innerHeight - toolBox.offsetHeight, t));
        toolBox.style.left = l + 'px';
        toolBox.style.top = t + 'px';
        toolBox.style.right = 'auto';
        toolBox.style.bottom = 'auto';
      });
      document.addEventListener('mouseup', () => { isDragging = false; });
      const historyHeader = document.getElementById('br_history_header');
      const historyContent = document.getElementById('br_history_content');
      if (historyHeader && historyContent) {
        historyHeader.onclick = () => {
          historyHeader.classList.toggle('expanded');
          historyContent.classList.toggle('show');
        };
        updateSearchHistoryUI();
      }
      updateStatusAfterInit();
    }, 500);
  }
  function updateStatusAfterInit() {
    const config = getConfig();
    const canRun = canRunSearch(config);
    if (taskStatus === STATUS_PAUSED) {
      updateStatus("等待開始...", "#666");
      updateStatusBadge(STATUS_PAUSED);
    } else if (taskStatus === STATUS_RUNNING && canRun) {
      updateStatus("腳本運行中...", "#e67e22");
      updateStatusBadge(STATUS_RUNNING);
    } else if (taskStatus === STATUS_RUNNING && !canRun) {
      onTaskCompleted();
    } else if (taskStatus === STATUS_RESTING) {
      updateStatus("任務已完成! 等待明日...", "#27ae60");
      updateCountdownUI("完成");
      updateStatusBadge(STATUS_RESTING);
    }
  }
  function getTaskOwnerKey() {
    return TASK_OWNER_KEY + '_' + getBingPageType();
  }
  function claimTask(force) {
    try {
      const key = getTaskOwnerKey();
      const raw = localStorage.getItem(key);
      if (!force && raw) {
        const o = JSON.parse(raw);
        if (o.id !== tabId && Date.now() - o.ts < 60000) return false;
      }
      localStorage.setItem(key, JSON.stringify({ id: tabId, ts: Date.now() }));
      return true;
    } catch (e) { return true; }
  }
  function releaseTask() {
    try {
      const key = getTaskOwnerKey();
      const raw = localStorage.getItem(key);
      if (raw) {
        const o = JSON.parse(raw);
        if (o.id === tabId) localStorage.removeItem(key);
      }
    } catch (e) { }
  }
  function heartbeatTask() {
    if (!isTaskRunning()) return;
    try {
      localStorage.setItem(getTaskOwnerKey(), JSON.stringify({ id: tabId, ts: Date.now() }));
    } catch (e) { }
  }
  function toggleScript() {
    const btn = document.getElementById('br_toggle_btn');
    if (isTaskRunning()) {
      setTabTaskStatus(STATUS_PAUSED);
      releaseTask();
      stopAutoScroll();
      stopTimer();
      btn.textContent = "▶ 繼續搜尋";
      btn.className = "br_btn br_btn_start";
      updateStatus("已暫停", "#666");
      updateCountdownUI("--");
      updateStatusBadge(STATUS_PAUSED);
    } else {
      startSearch(true);
    }
  }
  function startSearch(force) {
    checkLoginStatus();
    const config = getConfig();
    if (!claimTask(!!force)) {
      setTabTaskStatus(STATUS_PAUSED);
      updateStatus("其他分頁正在執行任務", "#e67e22");
      updateStatusBadge(STATUS_PAUSED);
      return;
    }
    const currentPageType = getBingPageType();
    if (currentPageType === 'pc' && config.pc_count >= CONFIG.max_pc) {
      releaseTask();
      setTabTaskStatus(STATUS_RESTING);
      stopAutoScroll();
      stopTimer();
      updateStatus("桌面版任務已達標", "#27ae60");
      updateCountdownUI("完成");
      updateStatusBadge(STATUS_RESTING);
      return;
    }
    if (currentPageType === 'ph' && config.ph_count >= CONFIG.max_ph) {
      releaseTask();
      setTabTaskStatus(STATUS_RESTING);
      stopAutoScroll();
      stopTimer();
      updateStatus("行動版任務已達標", "#27ae60");
      updateCountdownUI("完成");
      updateStatusBadge(STATUS_RESTING);
      return;
    }
    setTabTaskStatus(STATUS_RUNNING);
    const btn = document.getElementById('br_toggle_btn');
    if (btn) { btn.textContent = "⏸ 暫停搜尋"; btn.className = "br_btn br_btn_stop"; }
    updateStatus("腳本運行中...", "#e67e22");
    startSearchLoop();
    updateStatusBadge(STATUS_RUNNING);
  }
  function startSearchLoop() {
    if (!isTaskRunning()) return;
    if (timerActive) return;
    const config = getConfig();
    const currentPageType = getBingPageType();
    if (currentPageType === 'pc' && config.pc_count >= CONFIG.max_pc) { onTaskCompleted(); return; }
    if (currentPageType === 'ph' && config.ph_count >= CONFIG.max_ph) { onTaskCompleted(); return; }
    timerStart = Date.now();
    timerInterval = getRandomInterval();
    nextExecuteTime = Date.now() + timerInterval;
    saveScheduleTime(nextExecuteTime);
    updateCountdownUI(Math.ceil(timerInterval / 1000));
    timerActive = true;
    timerLoop();
  }
  function stopTimer() {
    timerActive = false;
    lastSecondUpdate = 0;
    if (timerHandle) {
      clearTimeout(timerHandle);
      timerHandle = null;
    }
  }
  let lastSecondUpdate = 0;
  function timerLoop() {
    if (!isTaskRunning()) {
      stopTimer();
      return;
    }
    const elapsed = Date.now() - timerStart;
    const remaining = Math.max(0, Math.ceil((timerInterval - elapsed) / 1000));
    if (remaining !== lastSecondUpdate) {
      lastSecondUpdate = remaining;
      updateCountdownUI(remaining);
      if (remaining > 0) {
        nextExecuteTime = Date.now() + (remaining * 1000);
        saveScheduleTime(nextExecuteTime);
      }
    }
    if (elapsed >= timerInterval) {
      stopTimer();
      updateCountdownUI("正在跳轉...");
      lastSecondUpdate = 0;
      nextExecuteTime = 0;
      saveScheduleTime(0);
      performSearch();
      return;
    }
    timerHandle = setTimeout(timerLoop, 250);
  }
  function saveScheduleTime(time) {
    try {
      localStorage.setItem('bing_auto_schedule', JSON.stringify({
        time: time,
        timestamp: Date.now()
      }));
    } catch (e) { }
  }
  function performSearch() {
    checkLoginStatus();
    if (!isTaskRunning()) return;
    const LOCK_KEY = 'bing_count_lock_' + getBingPageType();
    try {
      const held = localStorage.getItem(LOCK_KEY);
      if (held && Number(held) > Date.now() - 5000) return;
      localStorage.setItem(LOCK_KEY, String(Date.now()));
    } catch (e) { /* 忽略錯誤，單分頁場景直接執行 */ }
    const config = getConfig();
    const currentPageType = getBingPageType();
    const releaseLock = () => { try { localStorage.removeItem(LOCK_KEY); } catch (e) { } };
    if (currentPageType === 'pc' && config.pc_count >= CONFIG.max_pc) { releaseLock(); onTaskCompleted(); return; }
    if (currentPageType === 'ph' && config.ph_count >= CONFIG.max_ph) { releaseLock(); onTaskCompleted(); return; }
    let newConfig = { ...config };
    if (currentPageType === 'pc') newConfig.pc_count++;
    else newConfig.ph_count++;
    saveConfig(newConfig);
    releaseLock();
    updateUI();
    if ((currentPageType === 'pc' && newConfig.pc_count >= CONFIG.max_pc) || (currentPageType === 'ph' && newConfig.ph_count >= CONFIG.max_ph)) {
      onTaskCompleted();
      return;
    }
    getRandomKeyword().then(keyword => {
      let attempts = 0;
      const tryNext = (kw) => {
        if (isFullKeywordUsed(kw) && attempts < 5) {
          attempts++;
          getRandomKeyword().then(tryNext);
          return;
        }
        addUsedFullKeyword(kw);
        executeSearch(kw);
      };
      tryNext(keyword);
    });
  }
  function executeSearch(keyword) {
    try {
      let input = document.getElementById("sb_form_q");
      let btn = document.getElementById("sb_form_go");
      let form = document.getElementById("sb_form");
      if (!input) input = document.querySelector("input.b_searchbox") || document.querySelector("input[name='q']");
      if (!form) form = document.querySelector("form.b_searchbox");
      if (input) {
        input.focus();
        input.value = keyword;
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
        input.blur();
      }
      updateStatus(`正在搜尋: ${keyword}`, "#0078d4");
      addSearchHistory(keyword);
      setTimeout(() => {
        try {
          if (form) {
            form.submit();
          } else {
            if (!btn) btn = document.querySelector("button.b_searchboxSubmit") || document.querySelector("a[title='Search']") || document.querySelector(".search_icon");
            if (btn) btn.click();
          }
        } catch (e) { }
      }, 300);
      setTimeout(() => {
        const loc = new URL(window.location.href);
        if (isTaskRunning() && !(loc.pathname.startsWith('/search') && loc.search.startsWith('?'))) {
          window.location.href = 'https://www.bing.com/search?q=' + encodeURIComponent(keyword);
        }
      }, 4000);
    } catch (e) { }
  }
  function onTaskCompleted() {
    releaseTask();
    setTabTaskStatus(STATUS_RESTING);
    stopAutoScroll();
    stopTimer();
    const btn = document.getElementById('br_toggle_btn');
    if (btn) { btn.textContent = "▶ 開始搜尋"; btn.className = "br_btn br_btn_start"; }
    updateStatus("任務已完成! 等待明日自動重啟...", "#27ae60");
    updateCountdownUI("完成");
    updateStatusBadge(STATUS_RESTING);
  }
  function isTaskRunning() {
    return taskStatus === STATUS_RUNNING;
  }
  function getToday() {
    const d = new Date();
    if (CONFIG.timezone) {
      return new Intl.DateTimeFormat('en-CA', { timeZone: CONFIG.timezone, year: 'numeric', month: '2-digit', day: '2-digit' }).format(d);
    }
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  }
  function getStorageData() {
    try { const data = GM_getValue(STORAGE_KEY); if (data) return JSON.parse(data); } catch (e) { }
    return null;
  }
  function saveConfig(config) {
    try { GM_setValue(STORAGE_KEY, JSON.stringify(config)); } catch (e) { }
  }
  function getConfig() {
    const today = getToday();
    const stored = getStorageData();
    if (!stored || stored.lastDate !== today) {
      return { date: today, lastDate: today, pc_count: 0, ph_count: 0, autoStart: true };
    }
    return stored;
  }
  function getBingPageType() {
    const url = new URL(window.location.href);
    const form = url.searchParams.get('FORM');
    if (/(^|\.)m\.bing\.com$/i.test(url.hostname) || form === 'MH2MBB' || form === 'MBLAD') return 'ph';
    if (form === 'MH16PS' || form === 'HDRS2') return 'pc';
    return isMobile() ? 'ph' : 'pc';
  }
  function isMobile() {
    return /mobile|android|iphone|ipad|touch/i.test(navigator.userAgent.toLowerCase());
  }
  function getRandomInterval() {
    return Math.floor(Math.random() * ((CONFIG.max_interval - CONFIG.min_interval) * 1000 + 1)) + CONFIG.min_interval * 1000;
  }
  function updateStatus(text, color) {
    const el = document.getElementById("br_status_text");
    if (el) { el.textContent = text; el.style.color = color || "#333"; }
  }
  function updateCountdownUI(content) {
    const el = document.getElementById("br_countdown");
    if (!el) return;
    if (typeof content === 'string') el.textContent = content;
    else el.textContent = content > 0 ? `${Math.floor(content)}秒` : '跳轉中...';
  }
  function updateUI() {
    const data = getConfig();
    const pcEl = document.getElementById('pc_count');
    const phEl = document.getElementById('ph_count');
    if (pcEl) pcEl.textContent = String(data.pc_count);
    if (phEl) phEl.textContent = String(data.ph_count);
  }
  function updateStatusBadge(status) {
    const badge = document.getElementById('br_status_badge');
    const miniIcon = document.querySelector('.br_mini-icon');
    if (badge) {
      badge.className = 'br_status-badge';
      badge.classList.add(status);
      switch (status) {
      case STATUS_PAUSED:
        badge.textContent = '暫停';
        break;
      case STATUS_RUNNING:
        badge.textContent = '進行中';
        break;
      case STATUS_RESTING:
        badge.textContent = '休息中';
        break;
      }
    }
    if (miniIcon) {
      miniIcon.classList.remove('paused', 'running', 'resting');
      miniIcon.classList.add(status);
    }
  }
  function cleanCount(toolBox) {
    if (confirm("確定要重置今日的搜尋計數嗎？")) {
      const today = getToday();
      saveConfig({ date: today, lastDate: today, pc_count: 0, ph_count: 0, autoStart: false });
      setTabTaskStatus(STATUS_PAUSED);
      releaseTask();
      stopAutoScroll();
      stopTimer();
      updateUI();
      updateStatusBadge(STATUS_PAUSED);
      updateStatus("等待開始...", "#666");
      const btn = document.getElementById('br_toggle_btn');
      if (btn) { btn.textContent = "▶ 開始搜尋"; btn.className = "br_btn br_btn_start"; }
    }
  }
  function doAutoScroll() {
    if (!window.location.pathname.includes('/search')) {
      return;
    }
    if (!isTaskRunning()) {
      return;
    }
    stopAutoScroll();
    window.scrollTo({ top: 0, behavior: 'smooth' });
    scrollTimeout = setTimeout(() => {
      startScrollLoop();
    }, 3000);
  }
  function startScrollLoop() {
    if (!isTaskRunning() || !window.location.pathname.includes('/search')) {
      stopAutoScroll();
      return;
    }
    const scrollHeight = Math.max(document.body.scrollHeight, document.documentElement.scrollHeight);
    window.scrollTo({ top: scrollHeight, behavior: 'smooth' });
    scrollTimeout = setTimeout(() => {
      if (isTaskRunning() && window.location.pathname.includes('/search')) {
        window.scrollTo({ top: 0, behavior: 'smooth' });
        scrollInterval = setInterval(() => {
          if (!isTaskRunning() || !window.location.pathname.includes('/search')) {
            stopAutoScroll();
            return;
          }
          const sh = Math.max(document.body.scrollHeight, document.documentElement.scrollHeight);
          window.scrollTo({ top: sh, behavior: 'smooth' });
          scrollTimeout = setTimeout(() => {
            if (isTaskRunning() && window.location.pathname.includes('/search')) {
              window.scrollTo({ top: 0, behavior: 'smooth' });
            }
          }, 2000);
        }, 10000);
      }
    }, 2000);
  }
  function stopAutoScroll() {
    if (scrollInterval) {
      clearInterval(scrollInterval);
      scrollInterval = null;
    }
    if (scrollTimeout) {
      clearTimeout(scrollTimeout);
      scrollTimeout = null;
    }
  }
  function checkLoginStatus() {
    const idP = document.querySelector('#id_p');
    if (idP && idP.src && !/^data:/.test(idP.src)) return true;
    const idN = document.querySelector('#id_n');
    if (idN && idN.textContent.trim()) return true;
    if (!document.querySelector('#id_a')) return true;
    console.log('[BAS] 請登入後領取獎勵');
    updateStatus('請登入後領取獎勵', '#d63031');
    return false;
  }
  async function getRandomKeyword() {
    if (bingNewsKeywords.length > 0) {
      const available = bingNewsKeywords.filter(k => !usedKeywordsToday.has(k));
      if (available.length > 0) {
        const keyword = available[Math.floor(Math.random() * available.length)];
        usedKeywordsToday.add(keyword);
        if (usedKeywordsToday.size > MAX_RECENT_HISTORY) {
          const first = usedKeywordsToday.values().next().value;
          usedKeywordsToday.delete(first);
        }
        return keyword;
      } else {
        usedKeywordsToday.clear();
        const keyword = bingNewsKeywords[Math.floor(Math.random() * bingNewsKeywords.length)];
        usedKeywordsToday.add(keyword);
        return keyword;
      }
    }
    const { chinese, english } = CONFIG.sourceWeight;
    const roll = Math.random() * 100;
    if (roll < chinese) {
      return getRandomKeywordFromPool();
    } else {
      return await getEnWordKeyword();
    }
  }
  function getRandomKeywordFromPool() {
    const { none, prefix, suffix } = CONFIG.fixWeight;
    for (let attempt = 0; attempt < 15; attempt++) {
      const baseKeyword = getUniqueKeywordFromPool();
      const positionRoll = Math.random() * 100;
      let selectedFix = null;
      const fixType = positionRoll < prefix ? 'prefix' : positionRoll < prefix + suffix ? 'suffix' : 'none';
      if (fixType !== 'none') {
        const availableFixes = filterDuplicateFixes(baseKeyword, keywordFixPool);
        if (availableFixes.length === 0) continue;
        selectedFix = availableFixes[Math.floor(Math.random() * availableFixes.length)];
      }
      const keyPrefix = fixType === 'prefix' ? selectedFix : null;
      const keySuffix = fixType === 'suffix' ? selectedFix : null;
      if (!isComboUsed(keyPrefix, keySuffix, baseKeyword)) {
        markComboUsed(keyPrefix, keySuffix, baseKeyword);
        const result = fixType === 'prefix' ? `${selectedFix} ${baseKeyword}` : fixType === 'suffix' ? `${baseKeyword} ${selectedFix}` : baseKeyword;
        return removeDuplicateWords(result);
      }
    }
    resetComboTracking();
    const baseKeyword = getUniqueKeywordFromPool();
    const roll = Math.random() * 100;
    let result;
    if (roll < prefix) {
      const availableFixes = filterDuplicateFixes(baseKeyword, keywordFixPool);
      const fix = availableFixes.length > 0
      ? availableFixes[Math.floor(Math.random() * availableFixes.length)]
      : '';
      result = fix ? `${fix} ${baseKeyword}` : baseKeyword;
    } else if (roll < prefix + suffix) {
      const availableFixes = filterDuplicateFixes(baseKeyword, keywordFixPool);
      const fix = availableFixes.length > 0
      ? availableFixes[Math.floor(Math.random() * availableFixes.length)]
      : '';
      result = fix ? `${baseKeyword} ${fix}` : baseKeyword;
    } else {
      result = baseKeyword;
    }
    result = removeDuplicateWords(result);
    return result;
  }
  async function getEnWordKeyword() {
    const { none, prefix, suffix, both } = CONFIG.enFixWeight;
    let joke = null;
    try { const c = JSON.parse(localStorage.getItem('bing_joke_cache') || '[]'); if (c.length > 0) { joke = c.shift(); localStorage.setItem('bing_joke_cache', JSON.stringify(c)); } } catch (e) { }
    if (joke == null) {
      try {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 5000);
        const response = await fetch(JOKE_API_URL, { signal: controller.signal });
        clearTimeout(timeoutId);
        if (!response.ok) {
          throw new Error('API 請求失敗: ' + response.status);
        }
        const data = await response.json();
        const jokes = Array.isArray(data.jokes) ? data.jokes.map(j => j.joke).filter(j => j && j.trim()) : (data.joke ? [data.joke] : []);
        if (jokes.length > 0) { joke = jokes.shift(); localStorage.setItem('bing_joke_cache', JSON.stringify(jokes)); }
      } catch (e) { }
    }
    if (joke) {
      const av = (w) => enWordFixPool.filter(f => !w.toLowerCase().includes(f.toLowerCase()));
      const validWords = joke.split(/\s+/)
        .map(word => word.replace(/[^a-zA-Z]/g, ''))
        .filter(cleanWord => cleanWord.length >= 5);
        if (validWords.length > 0) {
          const wordCount = Math.floor(Math.random() * 3) + 1;
          const selectedWords = [];
          const pool = [...validWords];
          for (let i = 0; i < wordCount && pool.length > 0; i++) {
            selectedWords.push(pool.splice(Math.floor(Math.random() * pool.length), 1)[0]);
          }
          let enWord = selectedWords.join(' ');
          for (let attempt = 0; attempt < 10; attempt++) {
            const baseWord = selectedWords.join(' ');
            const positionRoll = Math.random() * 100;
            let tempEnWord = baseWord;
            if (positionRoll >= none && positionRoll < none + prefix) {
              const availableFixes = av(baseWord);
              if (availableFixes.length === 0) continue;
              const p = availableFixes[Math.floor(Math.random() * availableFixes.length)];
              tempEnWord = `${p} ${baseWord}`;
            } else if (positionRoll >= none + prefix && positionRoll < none + prefix + suffix) {
              const availableFixes = av(baseWord);
              if (availableFixes.length === 0) continue;
              const f = availableFixes[Math.floor(Math.random() * availableFixes.length)];
              tempEnWord = `${baseWord} ${f}`;
            } else if (positionRoll >= none + prefix + suffix) {
              const prefixPool = av(baseWord);
              if (prefixPool.length < 2) continue;
              const pi = Math.floor(Math.random() * prefixPool.length);
              const p = prefixPool[pi];
              const f = prefixPool[(pi + 1) % prefixPool.length];
              tempEnWord = `${p} ${baseWord} ${f}`;
            }
            enWord = removeDuplicateWords(tempEnWord);
            return enWord;
          }
          const positionRoll = Math.random() * 100;
          if (positionRoll < none) {
            return removeDuplicateWords(enWord);
          } else if (positionRoll < none + prefix) {
            const availableFixes = av(enWord);
            if (availableFixes.length > 0) {
              const p = availableFixes[Math.floor(Math.random() * availableFixes.length)];
              return removeDuplicateWords(`${p} ${enWord}`);
            }
          } else if (positionRoll < none + prefix + suffix) {
            const availableFixes = av(enWord);
            if (availableFixes.length > 0) {
              const f = availableFixes[Math.floor(Math.random() * availableFixes.length)];
              return removeDuplicateWords(`${enWord} ${f}`);
            }
          }
          return removeDuplicateWords(enWord);
        }
      }
    return getRandomKeywordFromPool();
  }
  if (/(^|\.)bing\.com$/i.test(window.location.hostname)) {
    init();
    let lastUrl = window.location.href;
    const urlObserver = new MutationObserver(() => {
      if (window.location.href !== lastUrl) {
        lastUrl = window.location.href;
        setTimeout(() => {
          if (isTaskRunning() && window.location.pathname.includes('/search')) {
            doAutoScroll();
            startSearchLoop();
          }
        }, 3000);
      }
    });
    urlObserver.observe(document.body, { childList: true, subtree: true });
  }
})();
