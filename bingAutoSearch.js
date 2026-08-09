// ==UserScript==
// @name         Bing Auto Search
// @version      2026041301
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
// ==/UserScript==

(function () {
  'use strict';

  const CONFIG = {
    max_pc: 45, // 桌面版搜尋次數上限
    max_ph: 35, // 行動版搜尋次數上限
    min_interval: 50, // 最小隨機秒數
    max_interval: 120, // 最大隨機秒數
    keywordsUrl: 'https://raw.githubusercontent.com/ss-vip/bing-auto-search/refs/heads/main/example.json', // 外部詞彙池 URL（JSON 格式）
    bingNewsUrl: 'https://www.bing.com/news/search?q=%e7%86%b1%e9%96%80%e5%a0%b1%e5%b0%8e&nvaug=%5bNewsVertical+Category%3d%22rt_MaxClass%22%5d', // Bing 熱門新聞

    // 詞綴權重配置（數值為百分比，0-100）
    fixWeight: {
      none: 60,   // 不加詞綴的機率
      prefix: 20, // 加前綴的機率
      suffix: 20  // 加後綴的機率
    },

    // 英文詞綴權重配置（數值為百分比，0-100）
    enFixWeight: {
      none: 60,   // 不加詞綴的機率
      prefix: 10, // 加前綴的機率
      suffix: 20, // 加後綴的機率
      both: 10    // 兩邊都加的機率
    },

    // 中文詞庫與英文笑話 API 關鍵字使用權重（數值為百分比，0-100）
    sourceWeight: {
      chinese: 55, // 中文機率
      english: 45  // 英文機率
    },

    defaultKeywordsPool: [
      'Python 教學', 'Java 環境變數', 'Linux 常用指令', 'Docker 部署', 'React vs Vue', 'ChatGPT API 教學', 'GitHub Copilot 評測',
      'SQL 優化 技巧', '正則表達式 教學', 'C++ 指標 教學', 'Rust 入門 教學', 'Unity 遊戲開發', 'VS Code',
      'Python 爬蟲 教學', 'iPhone 16', 'RTX 5090', 'MacBook Pro', 'PS5 Pro', 'Switch 2', '必玩Steam遊戲',
      '機械鍵盤', '降噪耳機', '智慧手錶', '感冒吃什麼改善', '番茄炒蛋做法', '避免近視眼', '減肥食譜', '影集', '超商便宜攻略',
      '小資旅游攻略', '遊樂園門票優惠', '自駕旅遊', '今日金價', '美元匯率', '油價走勢'
    ],
    defaultKeywordFixPool: ['最新', '資訊', '近期', '說明', '是啥', '常見', '有啥', '最好', '最推', '超便', '很優', '推薦'],
    defaultEnWordFixPool: ['英文', '中文', '翻譯', '解釋', '意思', '造句', '定義', '用法', '例句', '解說', '範例', '簡述']
  };

  // 防呆：驗證並修正 CONFIG 數值
  (function validateConfig() {
    // 驗證搜尋次數上限
    if (CONFIG.max_pc < 1) CONFIG.max_pc = 1;
    if (CONFIG.max_ph < 1) CONFIG.max_ph = 1;

    // 驗證隨機秒數（防呆：最小不能大於最大）
    const minInt = Math.max(1, parseInt(CONFIG.min_interval) || 1);
    const maxInt = Math.max(minInt, parseInt(CONFIG.max_interval) || minInt);
    CONFIG.min_interval = minInt;
    CONFIG.max_interval = maxInt;

    // 驗證詞綴權重（確保總和為 100%）
    const totalWeight = (cfg) => (cfg.none || 0) + (cfg.prefix || 0) + (cfg.suffix || 0) + (cfg.both || 0);

    if (totalWeight(CONFIG.fixWeight) !== 100) {
      const w = CONFIG.fixWeight;
      const none = Math.max(0, Math.min(100, w.none || 20));
      const prefix = Math.max(0, Math.min(100, w.prefix || 40));
      const suffix = 100 - none - prefix;
      CONFIG.fixWeight = { none, prefix, suffix: Math.max(0, suffix) };
    }

    if (totalWeight(CONFIG.enFixWeight) !== 100) {
      const w = CONFIG.enFixWeight;
      const none = Math.max(0, Math.min(100, w.none || 25));
      const prefix = Math.max(0, Math.min(100, w.prefix || 25));
      const suffix = Math.max(0, Math.min(100, w.suffix || 25));
      const both = 100 - none - prefix - suffix;
      CONFIG.enFixWeight = { none, prefix, suffix, both: Math.max(0, both) };
    }

    // 驗證來源權重（確保總和為 100%）
    const srcWeight = CONFIG.sourceWeight;
    const srcTotal = (srcWeight.chinese || 0) + (srcWeight.english || 0);
    if (srcTotal !== 100) {
      const chinese = Math.max(0, Math.min(100, srcWeight.chinese || 50));
      CONFIG.sourceWeight = { chinese, english: 100 - chinese };
    }
  })();

  // 詞彙池（含來源標記）
  let keywordsPool = CONFIG.defaultKeywordsPool;
  let keywordFixPool = CONFIG.defaultKeywordFixPool;
  let enWordFixPool = CONFIG.defaultEnWordFixPool;
  let bingNewsKeywords = []; // 從 Bing News 熱門新聞取得的關鍵字（優先使用）

  // 詞綴組合記錄（用於去重）
  let usedPrefixSuffixCombos = new Set();

  // 記錄已使用的關鍵詞（避免短期內重複）
  let usedKeywordsToday = new Set();
  let usedFullKeywords = new Set(); // 記錄已送出的完整搜尋關鍵字
  const MAX_RECENT_HISTORY = 50;  // 保留最近 50 個
  const MAX_FULL_KEYWORDS = 200; // 保留最近 200 個完整關鍵字

  // 合併並去重詞彙池
  function mergeAndDeduplicateKeywords(externalKeywords, defaultKeywords) {
    const combined = [...new Set([...defaultKeywords, ...externalKeywords])];
    return combined.filter(k => k && k.trim().length > 0);
  }

  // 合併並去重詞綴池
  function mergeAndDeduplicateFixes(externalFixes, defaultFixes) {
    return [...new Set([...defaultFixes, ...externalFixes])].filter(f => f && f.trim().length > 0);
  }

  // 重置詞綴組合記錄（跨天或全部用完時）
  function resetComboTracking() {
    usedPrefixSuffixCombos.clear();
  }

  // 生成唯一詞綴組合鍵值
  function getComboKey(prefix, suffix, baseKeyword) {
    return `${prefix || 'none'}_${suffix || 'none'}_${baseKeyword}`;
  }

  // 檢查詞綴組合是否已使用
  function isComboUsed(prefix, suffix, baseKeyword) {
    return usedPrefixSuffixCombos.has(getComboKey(prefix, suffix, baseKeyword));
  }

  // 標記詞綴組合已使用
  function markComboUsed(prefix, suffix, baseKeyword) {
    usedPrefixSuffixCombos.add(getComboKey(prefix, suffix, baseKeyword));
  }

  // 從詞庫中取得未重複的關鍵詞
  function getUniqueKeywordFromPool() {
    // 嘗試找到未使用過的詞
    const available = keywordsPool.filter(k => !usedKeywordsToday.has(k));

    let keyword;
    if (available.length > 0) {
      keyword = available[Math.floor(Math.random() * available.length)];
    } else {
      // 詞庫用完了，從頭重置
      usedKeywordsToday.clear();
      keyword = keywordsPool[Math.floor(Math.random() * keywordsPool.length)];
    }

    // 記錄並維護歷史數量
    usedKeywordsToday.add(keyword);
    if (usedKeywordsToday.size > MAX_RECENT_HISTORY) {
      // 移除最早的記錄
      const first = usedKeywordsToday.values().next().value;
      usedKeywordsToday.delete(first);
    }

    return keyword;
  }

  // 清除當日關鍵詞記錄（跨天時呼叫）
  function clearUsedKeywords() {
    usedKeywordsToday.clear();
    usedFullKeywords.clear();
  }

  // 記錄已送出的完整搜尋關鍵字
  function addUsedFullKeyword(keyword) {
    usedFullKeywords.add(keyword);
    // 限制記錄數量，防止記憶體過度使用
    if (usedFullKeywords.size > MAX_FULL_KEYWORDS) {
      // 移除最早的記錄（Set 無序，使用 Array 轉換）
      const arr = Array.from(usedFullKeywords);
      arr.shift();
      usedFullKeywords = new Set(arr);
    }
  }

  // 檢查完整關鍵字是否已送出過
  function isFullKeywordUsed(keyword) {
    return usedFullKeywords.has(keyword);
  }

  // 移除重複詞彙（非連續重複也要移除，如 "最新 油價走勢 最新" -> "最新 油價走勢"）
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

  // 檢查並移除與基礎關鍵詞重複的詞綴
  function filterDuplicateFixes(baseKeyword, fixes) {
    const baseWords = new Set(baseKeyword.split(/\s+/));
    // 過濾掉會與基礎關鍵詞重複的詞綴
    return fixes.filter(fix => !baseWords.has(fix));
  }

  // 檢查最終關鍵詞是否包含非連續重複
  function hasNonContiguousDuplicates(keyword) {
    const words = keyword.split(/\s+/);
    const seen = new Set();
    for (const word of words) {
      if (seen.has(word)) return true;
      seen.add(word);
    }
    return false;
  }

  const STORAGE_KEY = 'bingAutoSearch';
  const JOKE_API_URL = 'https://v2.jokeapi.dev/joke/Any?blacklistFlags=nsfw,religious,political,racist,sexist,explicit&type=single';
  const KEYWORDS_CACHE_KEY = 'bing_keywords_cache';
  const TASK_STATUS_KEY = 'bing_task_status';  // sessionStorage key for per-tab status
  const SEARCH_HISTORY_KEY = 'bing_search_history';  // 歷史搜尋記錄 key
  const MAX_HISTORY_RECORDS = 5;  // 歷史記錄上限
  const WAKEUP_TRIGGER_KEY = 'bing_auto_wakeup';  // 喚醒觸發器 key
  const CROSSDAY_CHECK_KEY = 'bing_crossday_check';  // 跨天檢查標記

  // 任務狀態常數
  const STATUS_PAUSED = 'paused';    // 已暫停：手動暫停後，不會進行任務
  const STATUS_RUNNING = 'running';  // 進行中：自動搜尋+滾動
  const STATUS_RESTING = 'resting'; // 休息中：已達每日上限，等待跨天自動重置

  // ============================================
  // 歷史搜尋記錄管理
  // ============================================
  function getSearchHistory() {
    try {
      const data = localStorage.getItem(SEARCH_HISTORY_KEY);
      if (data) return JSON.parse(data);
    } catch (e) { /* 忽略錯誤 */ }
    return [];
  }

  function addSearchHistory(keyword) {
    const history = getSearchHistory();
    const now = new Date();
    const record = {
      keyword: keyword,
      time: now.toLocaleString('zh-TW', { year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' })
    };

    // 加入到最前面
    history.unshift(record);

    // 保持上限為 5 筆
    if (history.length > MAX_HISTORY_RECORDS) {
      history.pop();
    }

    try {
      localStorage.setItem(SEARCH_HISTORY_KEY, JSON.stringify(history));
    } catch (e) { /* 忽略錯誤 */ }

    // 更新 UI（如果麵板已展開）
    updateSearchHistoryUI();
  }

  function updateSearchHistoryUI() {
    const historyContainer = document.getElementById('br_search_history_list');
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

  let taskStatus = STATUS_PAUSED;  // 當前任務狀態（每個分頁獨立）
  let timerStart = 0;
  let timerInterval = 0;
  let timerActive = false;  // 是否有活動的計時鏈（防止 startSearchLoop 重複啟動）
  let timerHandle = null;  // 計時 setTimeout handle
  let lastSearchTime = 0;  // 上次執行搜尋時間戳（防止雙重觸發）
  let isDragging = false;
  let dragX = 0, dragY = 0;
  let checkInterval = null;
  let currentKeyword = '';
  let nextExecuteTime = 0;  // 下次執行時間戳
  let isBackground = false;
  let scrollInterval = null;  // 滾動間隔計時器
  let scrollTimeout = null;  // 滾動超時計時器

  // 獲取當前分頁的任務狀態
  function getTabTaskStatus() {
    try {
      const stored = sessionStorage.getItem(TASK_STATUS_KEY);
      if (stored && [STATUS_PAUSED, STATUS_RUNNING, STATUS_RESTING].includes(stored)) {
        return stored;
      }
    } catch (e) { /* 忽略錯誤 */ }
    return null;  // 返回 null 表示沒有保存的狀態
  }

  // 設置當前分頁的任務狀態
  function setTabTaskStatus(status) {
    taskStatus = status;
    try {
      sessionStorage.setItem(TASK_STATUS_KEY, status);
    } catch (e) { /* 忽略錯誤 */ }
  }

  // 外部詞彙池載入（含 localStorage 緩存，每日檢查更新）
  async function loadExternalKeywords() {
    // 先檢查登入狀態
    if (!checkLoginStatus()) {
      console.log('[Bing Auto Search] 未登入，跳過外部詞彙載入');
      return false;
    }

    const today = getToday();
    let cacheData = null;
    try {
      const cached = localStorage.getItem(KEYWORDS_CACHE_KEY);
      if (cached) cacheData = JSON.parse(cached);
    } catch (e) {
      localStorage.removeItem(KEYWORDS_CACHE_KEY);
    }

    // 當天已載入過，直接使用快取（避免每次開分頁都重新 fetch）
    if (cacheData && cacheData.date === today) {
      keywordsPool = mergeAndDeduplicateKeywords(cacheData.keywords || [], CONFIG.defaultKeywordsPool);
      keywordFixPool = mergeAndDeduplicateFixes(cacheData.keywordFix || [], CONFIG.defaultKeywordFixPool);
      enWordFixPool = mergeAndDeduplicateFixes(cacheData.enWordFix || [], CONFIG.defaultEnWordFixPool);
      console.log(`[Bing Auto Search] 使用本地快取: ${keywordsPool.length} 組`);
      return true;
    }

    // 嘗試從外部 URL 載入最新詞彙（每日一次）
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

        // 合併外部詞彙與預設詞彙（去重）
        keywordsPool = mergeAndDeduplicateKeywords(externalKeywords, CONFIG.defaultKeywordsPool);
        keywordFixPool = mergeAndDeduplicateFixes(externalKeywordFix, CONFIG.defaultKeywordFixPool);
        enWordFixPool = mergeAndDeduplicateFixes(externalEnWordFix, CONFIG.defaultEnWordFixPool);

        // 更新 localStorage 緩存（僅儲存外部詞彙）
        localStorage.setItem(KEYWORDS_CACHE_KEY, JSON.stringify({
          date: today,
          version: data.version || null,
          keywords: externalKeywords,
          keywordFix: externalKeywordFix,
          enWordFix: externalEnWordFix,
          lastFetch: Date.now()
        }));

        console.log(`[Bing Auto Search] 詞彙庫已更新: ${keywordsPool.length} 組`);
        return true;
      } catch (e) {
        console.log('[Bing Auto Search] 外部詞彙載入失敗，使用快取或預設');
      }
    }

    // 若 fetch 失敗，使用舊快取（不限當天）
    if (cacheData && (cacheData.keywords || cacheData.keywordFix || cacheData.enWordFix)) {
      keywordsPool = mergeAndDeduplicateKeywords(cacheData.keywords || [], CONFIG.defaultKeywordsPool);
      keywordFixPool = mergeAndDeduplicateFixes(cacheData.keywordFix || [], CONFIG.defaultKeywordFixPool);
      enWordFixPool = mergeAndDeduplicateFixes(cacheData.enWordFix || [], CONFIG.defaultEnWordFixPool);
      console.log(`[Bing Auto Search] 使用本地快取: ${keywordsPool.length} 組`);
      return true;
    }

    // 完全無法載入時，使用預設詞彙池
    keywordsPool = CONFIG.defaultKeywordsPool;
    keywordFixPool = CONFIG.defaultKeywordFixPool;
    enWordFixPool = CONFIG.defaultEnWordFixPool;
    console.log('[Bing Auto Search] 使用預設詞彙池');
    return false;
  }

  // 從 Bing News 取得熱門搜尋關鍵字（優先使用）
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

      // 解析 HTML 並提取新聞標題
      const parser = new DOMParser();
      const doc = parser.parseFromString(htmlText, 'text/html');

      const titles = [];

      // 方法1: 嘗試原始選擇器 class="na_t news_title" 的 title 屬性
      let items = doc.querySelectorAll('.na_t.news_title');
      items.forEach(item => {
        const title = item.getAttribute('title')?.trim();
        if (title && title.length > 0) {
          titles.push(title);
        }
      });

      // 方法2: 如果方法1沒找到，嘗試從常見的新聞標題元素提取
      if (titles.length === 0) {
        // 嘗試從 b_algo 元素提取
        const bAlgoItems = doc.querySelectorAll('.b_algo a, .b_ans a');
        bAlgoItems.forEach(item => {
          const text = item.textContent?.trim();
          // 過濾掉時間標記、來源名稱等
          if (text && text.length > 4 && text.length < 80 &&
              !text.match(/^\d+\s*(小時|天|分鐘|小時前)$/) &&
              !text.includes('·') &&
              !text.match(/[A-Z][a-z]+\s+[A-Z][a-z]+/)) {
            titles.push(text);
          }
        });
      }

      // 方法3: 從 meta 標籤或標題標籤提取
      if (titles.length === 0) {
        const h2Elements = doc.querySelectorAll('h2, h3');
        h2Elements.forEach(el => {
          const text = el.textContent?.trim();
          if (text && text.length > 4 && text.length < 80 && !text.includes('熱門報導')) {
            titles.push(text);
          }
        });
      }

      // 去重並限制數量
      const uniqueTitles = [...new Set(titles)].slice(0, 30);

      if (uniqueTitles.length > 0) {
        bingNewsKeywords = uniqueTitles;
        console.log(`[Bing Auto Search] 從 Bing News 取得 ${uniqueTitles.length} 組關鍵字: ${uniqueTitles.join(', ')}`);
      }
    } catch (e) {
      console.log('[Bing Auto Search] Bing News 關鍵字載入失敗', e.message);
    }
  }

  // ============================================
  // 初始化
  // ============================================
  function init() {
    // 先檢查登入狀態，若未登入則不執行任務
    const loggedIn = checkLoginStatus();
    if (!loggedIn) {
      console.log('[Bing Auto Search] 未偵測到登入狀態，任務暫停');
      setTabTaskStatus(STATUS_PAUSED);
    }

    // 重置詞綴組合記錄
    resetComboTracking();

    // 只有在已登入狀態才嘗試載入外部詞彙池
    if (loggedIn) {
      // 嘗試載入外部詞彙池（非阻塞，若失敗則使用預設）
      loadExternalKeywords().then(() => {
        console.log('[Bing Auto Search] 外部詞彙載入完成');
      }).catch(err => {
        console.log('[Bing Auto Search] 使用預設詞彙池');
      });

      // 嘗試從 Bing News 載入熱門搜尋關鍵字（優先使用）
      loadPanelKeywords().catch(err => {
        console.log('[Bing Auto Search] Bing News 關鍵字載入失敗');
      });
    }

    // 跨天檢查
    checkAndResetDay();

    // 優先從 sessionStorage 恢復當前分頁的任務狀態
    const savedStatus = getTabTaskStatus();
    if (savedStatus && savedStatus !== STATUS_PAUSED) {
      // 恢復之前保存的狀態
      setTabTaskStatus(savedStatus);
      // 如果是 running 狀態，自動開始搜尋
      if (savedStatus === STATUS_RUNNING) {
        setTimeout(() => startSearch(), 1500);
      }
    } else {
      // 預設為暫停狀態
      setTabTaskStatus(STATUS_PAUSED);
    }

    initStyles();
    initUI();

    // 只有在 running 狀態時才執行滾動
    if (taskStatus === STATUS_RUNNING && window.location.href.includes('bing.com/search')) {
      doAutoScroll();
    }

    // 啟動保活機制
    startKeepAlive();

    // 設置跨天喚醒監聽
    setupCrossDayListener();

    // 頁面載入完成後執行滾動
    if (document.readyState === 'complete') {
      setTimeout(() => {
        if (isTaskRunning() && window.location.href.includes('bing.com/search')) {
          doAutoScroll();
        }
      }, 3000);
    } else {
      window.addEventListener('load', () => {
        setTimeout(() => {
          if (isTaskRunning() && window.location.href.includes('bing.com/search')) {
            doAutoScroll();
          }
        }, 3000);
      });
    }

    // 定時更新狀態徽章（URL 變化監測已由底部 MutationObserver 處理）
    let lastTaskStatus = taskStatus;
    setInterval(() => {
      if (lastTaskStatus !== taskStatus) {
        lastTaskStatus = taskStatus;
        updateStatusBadge(taskStatus);
      }
    }, 500);
  }

  // ============================================
  // 保活機制 - 背景頁支撐
  // ============================================
  function startKeepAlive() {
    // 檢測頁面可見性變化
    document.addEventListener('visibilitychange', handleVisibilityChange);

    // 頁面載入時檢查是否需要執行
    checkScheduledExecution();

    // 定時更新 UI 顯示最新的計數（從 GM_storage 讀取）
    if (checkInterval) clearInterval(checkInterval);
    checkInterval = setInterval(() => {
      checkAndResetDay();
      checkScheduledExecution();
      updateUI();
    }, 10000);
  }

  // 處理頁面可見性變化
  function handleVisibilityChange() {
    if (document.hidden) {
      // 頁面進入背景，記錄當前狀態
      isBackground = true;
    } else {
      // 頁面回到前景，檢查是否需要執行搜尋
      isBackground = false;
      checkScheduledExecution();
      // 恢復滾動（若正在搜尋頁）
      if (isTaskRunning() && window.location.href.includes('bing.com/search')) {
        doAutoScroll();
      }
    }
  }

  // 檢查排程的執行
  function checkScheduledExecution() {
    const now = Date.now();

    // 嘗試從 localStorage 讀取排程時間
    let scheduledTime = nextExecuteTime;
    if (scheduledTime === 0) {
      try {
        const saved = localStorage.getItem('bing_auto_schedule');
        if (saved) {
          const data = JSON.parse(saved);
          // 檢查是否過期（超過1小時視為無效）
          if (data.time > 0 && (now - data.timestamp) < 3600000) {
            scheduledTime = data.time;
          }
        }
      } catch (e) { /* 忽略錯誤 */ }
    }

    if (!isTaskRunning() || taskStatus === STATUS_RESTING) return;

    // 如果有排程的執行時間，且已經到了
    if (scheduledTime > 0 && now >= scheduledTime) {
      nextExecuteTime = 0;
      localStorage.removeItem('bing_auto_schedule');
      performSearch();
      return;
    }

        // 如果沒有排程，但正在運行中，重新計算下次執行時間
        if (isTaskRunning() && scheduledTime === 0) {
          // 從計時器狀態計算剩餘時間（與 timerLoop 相同使用 Date.now() 基準）
          const elapsed = Date.now() - timerStart;
          const remaining = timerInterval - elapsed;

          if (remaining > 0) {
            nextExecuteTime = now + remaining;
            saveScheduleTime(nextExecuteTime);
          }
        }
  }

  // ============================================
  // 跨天重置（支持背景頁）
  // ============================================
  function checkAndResetDay() {
    const stored = getStorageData();
    const today = getToday();

    if (stored && stored.lastDate !== today) {
      // 檢查跨天標記，避免重複觸發
      const crossdayMark = localStorage.getItem(CROSSDAY_CHECK_KEY);
      if (crossdayMark === today) {
        return;  // 已經處理過今天的跨天重置
      }

      console.log('[Bing Auto Search] 檢測到跨天，執行重置...');

      // 重置計數
      const newConfig = {
        date: today,
        lastDate: today,
        pc_count: 0,
        ph_count: 0,
        deviceType: getBingPageType(),
        autoStart: true
      };
      saveConfig(newConfig);

      // 重置詞綴組合記錄
      resetComboTracking();

      // 清除當日關鍵詞記錄
      clearUsedKeywords();

      // 如果是休息中狀態，跨天後自動恢復為進行中
      if (taskStatus === STATUS_RESTING) {
        console.log('[Bing Auto Search] 從休息中狀態恢復為進行中');
        setTabTaskStatus(STATUS_RUNNING);
        startSearchLoop();  // 重新開始搜尋迴圈
        doAutoScroll();     // 恢復滾動
      }

      // 設置跨天標記（當天只觸發一次）
      localStorage.setItem(CROSSDAY_CHECK_KEY, today);

      // 廣播跨天事件喚醒其他分頁
      broadcastWakeup();

      updateUI();
      updateStatus("跨天重置成功! 任務進行中...", "#e67e22");
      console.log('[Bing Auto Search] 跨天重置完成');
    } else if (!stored) {
      // 首次使用，設置跨天標記
      localStorage.setItem(CROSSDAY_CHECK_KEY, today);
    }
  }

  // 廣播喚醒事件到其他分頁
  function broadcastWakeup() {
    try {
      localStorage.setItem(WAKEUP_TRIGGER_KEY, JSON.stringify({
        action: 'WAKEUP',
        timestamp: Date.now()
      }));
      // 觸發 storage 事件讓其他分頁感知
      localStorage.removeItem(WAKEUP_TRIGGER_KEY);
    } catch (e) { /* 忽略錯誤 */ }
  }

  // 設置跨天喚醒監聽
  function setupCrossDayListener() {
    window.addEventListener('storage', (e) => {
      if (e.key === WAKEUP_TRIGGER_KEY && e.newValue) {
        try {
          const data = JSON.parse(e.newValue);
          if (data.action === 'WAKEUP') {
            // 收到喚醒信號，執行跨天重置
            console.log('[Bing Auto Search] 收到喚醒信號');
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

  // ============================================
  // 樣式與 UI
  // ============================================
  function initStyles() {
    GM_addStyle(`
      #br_reward_tool { position: fixed; right: 30px; bottom: 30px; left: auto; top: auto; background: #fff; padding: 0; border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.15); width: 260px; z-index: 9999999; transition: box-shadow 0.2s, opacity 0.2s; cursor: default; user-select: none; border: 1px solid #dcdcdc; box-sizing: border-box; text-align: left; line-height: 1.5; color: #333; }
      #br_reward_tool * { box-sizing: border-box; }
      #br_reward_tool .br_header { position: relative; height: 40px; border-top-left-radius: 8px; border-top-right-radius: 8px; background: #f5f5f5; border-bottom: 1px solid #e0e0e0; display: flex; align-items: center; justify-content: space-between; padding: 0 12px; cursor: move; width: 100%; }
      #br_reward_tool .br_title { font-size: 14px; font-weight: 600; color: #444; }
      #br_reward_tool .br_date { font-size: 11px; color: #888; margin-left: 8px; font-weight: normal; }
      #br_reward_tool .br_minimize-btn { border: none; background: none; cursor: pointer; font-size: 20px; color: #666; padding: 0; width: 24px; height: 24px; display: flex; align-items: center; justify-content: center; }
      #br_reward_tool .br_minimize-btn:hover { color: #0078d4; background: #e0e0e0; border-radius: 4px; }
      #br_reward_tool .br_panel-content { padding: 15px; background: #fff; border-bottom-left-radius: 8px; border-bottom-right-radius: 8px; }
      #br_reward_tool .br_btn { display: block; width: 100%; margin: 8px 0; padding: 8px 0; color: #fff; border-radius: 4px; text-align: center; font-weight: 600; text-decoration: none; font-size: 14px; cursor: pointer; transition: all 0.2s; border: none; outline: none; }
      .br_btn_start { background: #0078d4; }
      .br_btn_start:hover { background: #005bb5; }
      .br_btn_stop { background: #d63031; }
      .br_btn_stop:hover { background: #c0392b; }
      .br_btn_reset { background: #f0f0f0; color: #333 !important; border: 1px solid #ccc !important; font-weight: normal !important; }
      .br_btn_reset:hover { background: #e0e0e0; }
      #br_reward_tool p { margin: 8px 0; color: #444; font-size: 13px; display: flex; justify-content: space-between; align-items: center; }
      #br_reward_tool .br_count { font-weight: bold; color: #0078d4; font-size: 14px; }
      #br_reward_tool #br_status_text { color: #666; font-size: 12px; margin-top: 12px; text-align: center; display: block; background: #f9f9f9; padding: 4px; border-radius: 4px; }
      #br_reward_tool #br_countdown { color: #e67e22; font-weight: bold; }
      #br_reward_tool.br_minimized { width: 50px !important; height: 50px !important; padding: 0 !important; background: transparent !important; box-shadow: none !important; border: none !important; right: 30px !important; bottom: 50px !important; }
      #br_reward_tool.br_minimized .br_header, #br_reward_tool.br_minimized .br_panel-content { display: none !important; }
      #br_reward_tool .br_mini-icon { width: 50px; height: 50px; border-radius: 50%; background: #0078d4; display: flex; align-items: center; justify-content: center; color: #fff; font-size: 12px; cursor: pointer; box-shadow: 0 4px 12px rgba(0,0,0,0.3); font-weight: bold; border: 2px solid #fff; text-align: center; line-height: 1.2; }
      #br_reward_tool .br_mini-icon:hover { transform: scale(1.05); background: #005bb5; }
      #br_reward_tool .br_mini-icon.running { background: #d63031; animation: breathe 2s ease-in-out infinite; }
      @keyframes breathe { 0% { opacity: 1; box-shadow: 0 0 8px rgba(214, 48, 49, 0.5); } 50% { opacity: 0.6; box-shadow: 0 0 16px rgba(214, 48, 49, 0.8); } 100% { opacity: 1; box-shadow: 0 0 8px rgba(214, 48, 49, 0.5); } }
      #br_reward_tool .br_auto-badge { display: inline-block; background: #27ae60; color: #fff; font-size: 10px; padding: 2px 6px; border-radius: 3px; margin-left: 6px; vertical-align: middle; }
      #br_reward_tool .br_live-indicator { display: inline-block; width: 8px; height: 8px; border-radius: 50%; background: #27ae60; margin-right: 6px; animation: pulse 1.5s infinite; }
      @keyframes pulse { 0% { opacity: 1; } 50% { opacity: 0.5; } 100% { opacity: 1; } }
      #br_reward_tool .br_mini-icon.paused { background: #0078d4; }
      #br_reward_tool .br_mini-icon.resting { background: #27ae60; }
      #br_reward_tool .br_status-badge { display: inline-block; font-size: 10px; padding: 2px 6px; border-radius: 3px; margin-left: 6px; vertical-align: middle; }
      #br_reward_tool .br_status-badge.paused { background: #666; color: #fff; }
      #br_reward_tool .br_status-badge.running { background: #e67e22; color: #fff; }
      #br_reward_tool .br_status-badge.resting { background: #27ae60; color: #fff; }
      #br_reward_tool .br_history-accordion { margin-top: 12px; border: 1px solid #e0e0e0; border-radius: 6px; overflow: hidden; }
      #br_reward_tool .br_history-header { display: flex; justify-content: space-between; align-items: center; padding: 8px 12px; background: #f9f9f9; cursor: pointer; font-size: 13px; font-weight: 500; color: #444; user-select: none; }
      #br_reward_tool .br_history-header:hover { background: #f0f0f0; }
      #br_reward_tool .br_history-arrow { transition: transform 0.2s; font-size: 10px; color: #888; }
      #br_reward_tool .br_history-header.expanded .br_history-arrow { transform: rotate(180deg); }
      #br_reward_tool .br_history-content { display: none; max-height: 200px; overflow-y: auto; background: #fff; }
      #br_reward_tool .br_history-content.show { display: block; }
      #br_reward_tool .br_history-list { padding: 8px 12px; }
    `);
  }

  function initUI() {
    const countInfo = getConfig();
    const today = getToday();
    const toolHtml = `
      <div id="br_reward_tool" class="br_minimized">
        <div class="br_header">
          <span class="br_title"><span class="br_live-indicator"></span>隨機搜尋 <span class="br_status-badge" id="br_status_badge">暫停</span></span>
          <span class="br_date">${today}</span>
          <button class="br_minimize-btn" title="最小化">–</button>
        </div>
        <div class="br_panel-content" style="display: none;">
          <button id="br_toggle_btn" class="br_btn br_btn_start">▶ 開始搜尋</button>
          <div style="border-top: 1px solid #eee; margin: 10px 0;"></div>
          <p>桌面版搜尋: <span><span class="br_count" id="pc_count">${countInfo.pc_count}</span> / ${CONFIG.max_pc}</span></p>
          <p>行動版搜尋: <span><span class="br_count" id="ph_count">${countInfo.ph_count}</span> / ${CONFIG.max_ph}</span></p>
          <p>下一次搜尋: <span id="br_countdown">--</span></p>
          <span id="br_status_text">等待開始...</span>
          <button id="br_reset_btn" class="br_btn br_btn_reset" style="margin-top:10px;">↺ 重置今日計數</button>
          <div class="br_history-accordion">
            <div class="br_history-header" id="br_history_header">
              <span>📜 最近搜尋記錄</span>
              <span class="br_history-arrow">▼</span>
            </div>
            <div class="br_history-content" id="br_history_content">
              <div class="br_history-list" id="br_search_history_list">
                <div style="color: #999; font-size: 12px; text-align: center; padding: 8px;">尚無搜尋記錄</div>
              </div>
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
      const panelContent = toolBox.querySelector('.br_panel-content');
      const header = toolBox.querySelector('.br_header');

      minBtn.onclick = (e) => {
        e.stopPropagation();
        toolBox.classList.add('br_minimized');
        panelContent.style.display = 'none';
        header.style.display = 'none';
        miniIcon.style.display = 'flex';
        toolBox.style.right = '30px'; toolBox.style.bottom = '30px'; toolBox.style.left = 'auto'; toolBox.style.top = 'auto';
      };

      miniIcon.onclick = (e) => {
        e.stopPropagation();
        toolBox.classList.remove('br_minimized');
        panelContent.style.display = 'block';
        header.style.display = 'flex';
        miniIcon.style.display = 'none';
        toolBox.style.right = '30px'; toolBox.style.bottom = '30px'; toolBox.style.left = 'auto'; toolBox.style.top = 'auto';
      };

      header.onmousedown = (e) => {
        isDragging = true;
        dragX = e.clientX - toolBox.offsetLeft;
        dragY = e.clientY - toolBox.offsetTop;
        toolBox.style.transition = 'none';
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
      document.addEventListener('mouseup', () => { isDragging = false; toolBox.style.transition = ''; });

      // 歷史搜尋記錄手風琴事件
      const historyHeader = document.getElementById('br_history_header');
      const historyContent = document.getElementById('br_history_content');
      if (historyHeader && historyContent) {
        historyHeader.onclick = () => {
          historyHeader.classList.toggle('expanded');
          historyContent.classList.toggle('show');
        };
        // 初始化時更新歷史記錄 UI
        updateSearchHistoryUI();
      }

      updateStatusAfterInit();
    }, 500);
  }

  function updateStatusAfterInit() {
    const config = getConfig();
    const canRun = canRunSearch(config);

    // 根據任務狀態顯示對應文字
    if (taskStatus === STATUS_PAUSED) {
      updateStatus("等待開始...", "#666");
      updateStatusBadge(STATUS_PAUSED);
    } else if (taskStatus === STATUS_RUNNING && canRun) {
      updateStatus("腳本運行中...", "#e67e22");
      updateStatusBadge(STATUS_RUNNING);
    } else if (taskStatus === STATUS_RUNNING && !canRun) {
      // 已達上限但狀態仍是 running，轉為 resting
      onTaskCompleted();
    } else if (taskStatus === STATUS_RESTING) {
      updateStatus("任務已完成! 等待明日...", "#27ae60");
      updateCountdownUI("完成");
      updateStatusBadge(STATUS_RESTING);
    }
  }

  // ============================================
  // 核心搜尋邏輯
  // ============================================
  function toggleScript() {
    if (!checkLoginStatus()) return;

    const btn = document.getElementById('br_toggle_btn');

    if (isTaskRunning()) {
      setTabTaskStatus(STATUS_PAUSED);
      stopAutoScroll();  // 停止頁面滾動
      stopTimer();  // 停止計時鏈
      btn.textContent = "▶ 繼續搜尋";
      btn.className = "br_btn br_btn_start";
      updateStatus("已暫停", "#666");
      updateCountdownUI("--");
      updateStatusBadge(STATUS_PAUSED);
    } else {
      const config = getConfig();
      const currentPageType = getBingPageType();

      // 如果已達上限，進入休息中狀態
      if (currentPageType === 'pc' && config.pc_count >= CONFIG.max_pc) {
        setTabTaskStatus(STATUS_RESTING);
        updateStatus("桌面版任務已達標", "#27ae60");
        updateCountdownUI("完成");
        return;
      }
      if (currentPageType === 'ph' && config.ph_count >= CONFIG.max_ph) {
        setTabTaskStatus(STATUS_RESTING);
        updateStatus("行動版任務已達標", "#27ae60");
        updateCountdownUI("完成");
        return;
      }

      setTabTaskStatus(STATUS_RUNNING);
      btn.textContent = "⏸ 暫停搜尋";
      btn.className = "br_btn br_btn_stop";
      updateStatus("腳本運行中...", "#e67e22");
      startSearchLoop();
      updateStatusBadge(STATUS_RUNNING);
    }
  }

  function startSearch() {
    if (!checkLoginStatus()) return;

    const config = getConfig();
    const currentPageType = getBingPageType();

    if (currentPageType === 'pc' && config.pc_count >= CONFIG.max_pc) {
      setTabTaskStatus(STATUS_RESTING);
      stopAutoScroll();
      stopTimer();
      updateStatus("桌面版任務已達標", "#27ae60");
      updateCountdownUI("完成");
      updateStatusBadge(STATUS_RESTING);
      return;
    }
    if (currentPageType === 'ph' && config.ph_count >= CONFIG.max_ph) {
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
    if (timerActive) return;  // 已有活動計時鏈，避免重複啟動

    const config = getConfig();
    const currentPageType = getBingPageType();

    if (currentPageType === 'pc' && config.pc_count >= CONFIG.max_pc) { onTaskCompleted(); return; }
    if (currentPageType === 'ph' && config.ph_count >= CONFIG.max_ph) { onTaskCompleted(); return; }

    // 使用 Date.now() 計算間隔（setTimeout 在背景分頁仍能運作，rAF 會被暫停）
    timerStart = Date.now();
    timerInterval = getRandomInterval();

    // 記錄下次執行時間（支持背景執行）
    nextExecuteTime = Date.now() + timerInterval;
    saveScheduleTime(nextExecuteTime);

    updateCountdownUI(Math.ceil(timerInterval / 1000));

    timerActive = true;
    timerLoop();
  }

  // 停止計時鏈
  function stopTimer() {
    timerActive = false;
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

    // 每秒更新倒數
    const elapsed = Date.now() - timerStart;
    const remaining = Math.max(0, Math.ceil((timerInterval - elapsed) / 1000));

    if (remaining !== lastSecondUpdate) {
      lastSecondUpdate = remaining;
      updateCountdownUI(remaining);

      // 即時更新排程時間
      if (remaining > 0) {
        nextExecuteTime = Date.now() + (remaining * 1000);
        saveScheduleTime(nextExecuteTime);
      }
    }

    // 時間到，執行搜尋
    if (elapsed >= timerInterval) {
      stopTimer();
      updateCountdownUI("正在跳轉...");
      lastSecondUpdate = 0;
      nextExecuteTime = 0;
      saveScheduleTime(0);
      performSearch();
      return;
    }

    // 繼續計時
    timerHandle = setTimeout(timerLoop, 250);
  }

  // 保存排程時間到存儲
  function saveScheduleTime(time) {
    try {
      localStorage.setItem('bing_auto_schedule', JSON.stringify({
        time: time,
        timestamp: Date.now()
      }));
    } catch (e) { /* 忽略錯誤 */ }
  }

  function performSearch() {
    if (!checkLoginStatus()) return;
    if (!isTaskRunning()) return;
    // 防止計時鏈與排程輪詢同時觸發造成重複搜尋
    if (Date.now() - lastSearchTime < 2000) return;

    // 跨分頁計數鎖：避免多分頁併發時超過每日上限
    const LOCK_KEY = 'bing_count_lock';
    try {
      const held = localStorage.getItem(LOCK_KEY);
      // 5 秒租約，避免某分頁卡住造成死鎖
      if (held && Number(held) > Date.now() - 5000) return;
      localStorage.setItem(LOCK_KEY, String(Date.now()));
    } catch (e) { /* 忽略錯誤，單分頁場景直接執行 */ }

    const config = getConfig();
    const currentPageType = getBingPageType();

    const releaseLock = () => { try { localStorage.removeItem(LOCK_KEY); } catch (e) { /* 忽略錯誤 */ } };

    if (currentPageType === 'pc' && config.pc_count >= CONFIG.max_pc) { releaseLock(); onTaskCompleted(); return; }
    if (currentPageType === 'ph' && config.ph_count >= CONFIG.max_ph) { releaseLock(); onTaskCompleted(); return; }

    let newConfig = { ...config };
    if (currentPageType === 'pc') newConfig.pc_count++;
    else newConfig.ph_count++;
    saveConfig(newConfig);
    releaseLock();
    updateUI();

    lastSearchTime = Date.now();

    if ((currentPageType === 'pc' && newConfig.pc_count >= CONFIG.max_pc) || (currentPageType === 'ph' && newConfig.ph_count >= CONFIG.max_ph)) {
      onTaskCompleted();
      return;
    }

    getRandomKeyword().then(keyword => {
      // 檢查完整關鍵字是否已送出過，嘗試最多 5 次取得不重複的關鍵字
      let attempts = 0;
      const tryNext = (kw) => {
        if (isFullKeywordUsed(kw) && attempts < 5) {
          attempts++;
          getRandomKeyword().then(tryNext);
          return;
        }
        currentKeyword = kw;
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
        } catch (e) { /* 忽略錯誤 */ }
      }, 300);

      // 強制跳轉後備
      setTimeout(() => {
        if (isTaskRunning() && !window.location.href.includes('bing.com/search?')) {
          window.location.href = 'https://www.bing.com/search?q=' + encodeURIComponent(keyword);
        }
      }, 4000);
    } catch (e) { /* 忽略錯誤 */ }
  }

  function onTaskCompleted() {
    setTabTaskStatus(STATUS_RESTING);  // 進入休息中狀態
    stopAutoScroll();  // 停止頁面滾動
    stopTimer();  // 停止計時鏈

    const btn = document.getElementById('br_toggle_btn');
    if (btn) { btn.textContent = "▶ 開始搜尋"; btn.className = "br_btn br_btn_start"; }

    updateStatus("任務已完成! 等待明日自動重啟...", "#27ae60");
    updateCountdownUI("完成");
    updateStatusBadge(STATUS_RESTING);
  }

  // ============================================
  // 任務狀態管理
  // ============================================
  function isTaskRunning() {
    return taskStatus === STATUS_RUNNING;
  }

  // 檢查是否可執行任務（任務狀態為 running 且未達上限）

  // ============================================
  // 工具函數
  // ============================================
  function getToday() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  }

  function getStorageData() {
    try { const data = GM_getValue(STORAGE_KEY); if (data) return JSON.parse(data); } catch (e) { /* 忽略錯誤 */ }
    return null;
  }

  function saveConfig(config) {
    try { GM_setValue(STORAGE_KEY, JSON.stringify(config)); } catch (e) { /* 忽略錯誤 */ }
  }

  function getConfig() {
    const today = getToday();
    const stored = getStorageData();

    if (!stored || stored.lastDate !== today) {
      return { date: today, lastDate: today, pc_count: 0, ph_count: 0, deviceType: getBingPageType(), autoStart: true };
    }
    return stored;
  }

  function getBingPageType() {
    const url = window.location.href;
    if (url.includes('FORM=MH2MBB') || url.includes('FORM=MBLAD')) return 'ph';
    if (url.includes('FORM=MH16PS') || url.includes('FORM=HDRS2')) return 'pc';
    return isMobile() ? 'ph' : 'pc';
  }

  function isMobile() {
    return /mobile|android|iphone|ipad|touch/i.test(navigator.userAgent.toLowerCase()) || window.innerWidth < 768;
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

    // 移除所有狀態類別
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

    // 更新 mini-icon 樣式
    if (miniIcon) {
      miniIcon.classList.remove('paused', 'running', 'resting');
      miniIcon.classList.add(status);
    }
  }

  function cleanCount(toolBox) {
    if (confirm("確定要重置今日的搜尋計數嗎？")) {
      const today = getToday();
      const currentPageType = getBingPageType();
      saveConfig({ date: today, lastDate: today, pc_count: 0, ph_count: 0, deviceType: currentPageType, autoStart: false });
      setTabTaskStatus(STATUS_PAUSED);
      stopAutoScroll();  // 停止頁面滾動
      stopTimer();  // 停止計時鏈
      updateUI();
      updateStatusBadge(STATUS_PAUSED);
      updateStatus("等待開始...", "#666");
    }
  }

  function doAutoScroll() {
    if (!window.location.href.includes('bing.com/search')) {
      return;
    }
    if (!isTaskRunning()) {
      return;
    }

    // 清除之前的滾動計時器
    stopAutoScroll();

    // 平滑滾動到頂部
    window.scrollTo({ top: 0, behavior: 'smooth' });

    // 延遲後開始滾動
    scrollTimeout = setTimeout(() => {
      startScrollLoop();
    }, 3000);
  }

  function startScrollLoop() {
    if (!isTaskRunning() || !window.location.href.includes('bing.com/search')) {
      stopAutoScroll();
      return;
    }

    // 平滑滾動
    const scrollHeight = Math.max(document.body.scrollHeight, document.documentElement.scrollHeight);
    window.scrollTo({ top: scrollHeight, behavior: 'smooth' });

    scrollTimeout = setTimeout(() => {
      if (isTaskRunning() && window.location.href.includes('bing.com/search')) {
        window.scrollTo({ top: 0, behavior: 'smooth' });

        // 滾動完成後，繼續下一次迴圈
        scrollInterval = setInterval(() => {
          if (!isTaskRunning() || !window.location.href.includes('bing.com/search')) {
            stopAutoScroll();
            return;
          }

          // 到底部
          const sh = Math.max(document.body.scrollHeight, document.documentElement.scrollHeight);
          window.scrollTo({ top: sh, behavior: 'smooth' });

          // 回頂部
          scrollTimeout = setTimeout(() => {
            if (isTaskRunning() && window.location.href.includes('bing.com/search')) {
              window.scrollTo({ top: 0, behavior: 'smooth' });
            }
          }, 2000);
        }, 10000);
      }
    }, 2000);
  }

  function stopAutoScroll() {
    // 清除所有滾動計時器
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
    // 檢查 Bing 登入狀態（多個可能的選擇器）
    const signInSelectors = [
      'span.sw_spd.id_avatar#id_a[aria-label="Sign in"]',
      '#id_a[aria-label*="Sign"]',
      'a[href*="signin"]',
      '.useravatar'
    ];

    for (const selector of signInSelectors) {
      const signInElement = document.querySelector(selector);
      if (signInElement) {
        const computedStyle = window.getComputedStyle(signInElement);
        // 檢查是否可見（未被隱藏）
        const isVisible = computedStyle.display !== 'none' && computedStyle.visibility !== 'hidden';
        if (isVisible) {
          console.log('[Bing Auto Search] 請登入 Bing 以繼續任務');
          updateStatus('請登入 Bing 以繼續任務', '#d63031');
          if (isTaskRunning()) {
            setTabTaskStatus(STATUS_PAUSED);
            stopAutoScroll();
            stopTimer();
          }
          return false;
        }
      }
    }
    return true;
  }

  async function getRandomKeyword() {
    // 優先使用 Bing News 關鍵字
    if (bingNewsKeywords.length > 0) {
      // 嘗試找到未使用過的 keyword
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
        // Bing News keywords 用完了，從頭重置
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

    // 嘗試生成不重複的詞綴組合（最多嘗試 15 次）
    for (let attempt = 0; attempt < 15; attempt++) {
      const baseKeyword = getUniqueKeywordFromPool();
      const positionRoll = Math.random() * 100;
      let selectedFix = null;
      let fixType = 'none';

      if (positionRoll < prefix) {
        // 加前綴（過濾掉與基礎關鍵詞重複的詞綴）
        const availableFixes = filterDuplicateFixes(baseKeyword, keywordFixPool);
        if (availableFixes.length === 0) continue;
        selectedFix = availableFixes[Math.floor(Math.random() * availableFixes.length)];
        fixType = 'prefix';
      } else if (positionRoll < prefix + suffix) {
        // 加後綴（過濾掉與基礎關鍵詞重複的詞綴）
        const availableFixes = filterDuplicateFixes(baseKeyword, keywordFixPool);
        if (availableFixes.length === 0) continue;
        selectedFix = availableFixes[Math.floor(Math.random() * availableFixes.length)];
        fixType = 'suffix';
      }
      // none: 不加詞綴

      // 檢查組合是否已使用
      const keyPrefix = fixType === 'prefix' ? selectedFix : null;
      const keySuffix = fixType === 'suffix' ? selectedFix : null;
      if (!isComboUsed(keyPrefix, keySuffix, baseKeyword)) {
        // 標記為已使用
        markComboUsed(keyPrefix, keySuffix, baseKeyword);
        let result;
        if (fixType === 'prefix') result = `${selectedFix} ${baseKeyword}`;
        else if (fixType === 'suffix') result = `${baseKeyword} ${selectedFix}`;
        else result = baseKeyword;

        // 移除最終關鍵詞中的重複詞彙
        result = removeDuplicateWords(result);

        // 如果仍有非連續重複，嘗試重新生成
        if (hasNonContiguousDuplicates(result)) continue;

        return result;
      }
    }

    // 所有組合都已使用，重置並返回隨機結果
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

    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 5000); // 5秒超時

      const response = await fetch(JOKE_API_URL, { signal: controller.signal });
      clearTimeout(timeoutId);

      if (!response.ok) {
        throw new Error('API 請求失敗: ' + response.status);
      }

      const data = await response.json();

      if (data && data.joke) {
        const validWords = data.joke.split(/\s+/)
          .map(word => word.replace(/[^a-zA-Z]/g, ''))
          .filter(cleanWord => cleanWord.length >= 5);

        if (validWords.length > 0) {
          // 隨機獲取 1~3 組單字（均勻抽樣，避免 sort(shuffle) 偏差）
          const wordCount = Math.floor(Math.random() * 3) + 1;
          const selectedWords = [];
          const pool = [...validWords];
          for (let i = 0; i < wordCount && pool.length > 0; i++) {
            selectedWords.push(pool.splice(Math.floor(Math.random() * pool.length), 1)[0]);
          }
          let enWord = selectedWords.join(' ');

          // 嘗試生成不重複的組合（最多 10 次）
          for (let attempt = 0; attempt < 10; attempt++) {
            const baseWord = selectedWords.join(' ');
            const positionRoll = Math.random() * 100;
            let tempEnWord = baseWord;

            if (positionRoll >= none && positionRoll < none + prefix) {
              // 只在前面加（過濾掉與基礎單字重複的詞綴）
              const availableFixes = enWordFixPool.filter(f => !baseWord.toLowerCase().includes(f.toLowerCase()));
              if (availableFixes.length === 0) continue;
              const p = availableFixes[Math.floor(Math.random() * availableFixes.length)];
              tempEnWord = `${p} ${baseWord}`;
            } else if (positionRoll >= none + prefix && positionRoll < none + prefix + suffix) {
              // 只在後面加（過濾掉與基礎單字重複的詞綴）
              const availableFixes = enWordFixPool.filter(f => !baseWord.toLowerCase().includes(f.toLowerCase()));
              if (availableFixes.length === 0) continue;
              const f = availableFixes[Math.floor(Math.random() * availableFixes.length)];
              tempEnWord = `${baseWord} ${f}`;
            } else if (positionRoll >= none + prefix + suffix) {
              // 兩邊都加（確保前綴和後綴不同，且不與基礎單字重複）
              const baseLower = baseWord.toLowerCase();
              const prefixPool = enWordFixPool.filter(f => !baseLower.includes(f.toLowerCase()));
              const suffixPool = enWordFixPool.filter(f => !baseLower.includes(f.toLowerCase()) && f !== (prefixPool[0] || ''));

              if (prefixPool.length === 0 || suffixPool.length === 0) continue;

              const p = prefixPool[Math.floor(Math.random() * prefixPool.length)];
              const f = suffixPool[Math.floor(Math.random() * suffixPool.length)];
              tempEnWord = `${p} ${baseWord} ${f}`;
            }
            // none: 不加詞綴

            // 移除最終關鍵詞中的重複詞彙
            enWord = removeDuplicateWords(tempEnWord);

            // 如果仍有非連續重複，嘗試重新生成
            if (hasNonContiguousDuplicates(enWord)) continue;

            return enWord;
          }

          // 所有組合都已使用，返回隨機結果（不加重複詞綴）
          const positionRoll = Math.random() * 100;
          if (positionRoll < none) {
            return removeDuplicateWords(enWord);
          } else if (positionRoll < none + prefix) {
            const availableFixes = enWordFixPool.filter(f => !enWord.toLowerCase().includes(f.toLowerCase()));
            if (availableFixes.length > 0) {
              const p = availableFixes[Math.floor(Math.random() * availableFixes.length)];
              return removeDuplicateWords(`${p} ${enWord}`);
            }
          } else if (positionRoll < none + prefix + suffix) {
            const availableFixes = enWordFixPool.filter(f => !enWord.toLowerCase().includes(f.toLowerCase()));
            if (availableFixes.length > 0) {
              const f = availableFixes[Math.floor(Math.random() * availableFixes.length)];
              return removeDuplicateWords(`${enWord} ${f}`);
            }
          }
          return removeDuplicateWords(enWord);
        }
      }
    } catch (e) { /* 忽略錯誤 */ }
    return getRandomKeywordFromPool();
  }

  // 啟動
  if (window.location.href.includes('bing.com/search') || window.location.href.includes('bing.com/')) {
    init();

    // 偵測 URL 變化（處理 SPA 頁面跳轉）
    let lastUrl = window.location.href;
    const urlObserver = new MutationObserver(() => {
      if (window.location.href !== lastUrl) {
        lastUrl = window.location.href;
        // 頁面跳轉後延遲執行滾動和重新開始搜尋
        setTimeout(() => {
          if (isTaskRunning() && window.location.href.includes('bing.com/search')) {
            doAutoScroll();
            // 重新開始搜尋迴圈
            startSearchLoop();
          }
        }, 3000);
      }
    });
    urlObserver.observe(document.body, { childList: true, subtree: true });
  }
})();
