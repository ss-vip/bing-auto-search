// ==UserScript==
// @name         Bing Auto Search
// @version      2026040101
// @description  無人值守 Bing 自動隨機搜尋
// @author       Hank
// @match        https://*.bing.com/*
// @icon         https://www.google.com/s2/favicons?sz=64&domain=bing.com
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_addValueChangeListener
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
    keywordsUrl: '', // 外部關鍵詞池 URL（JSON 格式）
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

  // 預設詞彙池
  let keywordsPool = CONFIG.defaultKeywordsPool;
  let keywordFixPool = CONFIG.defaultKeywordFixPool;
  let enWordFixPool = CONFIG.defaultEnWordFixPool;

  const STORAGE_KEY = 'bingAutoSearch';
  const AUTO_RUN_KEY = 'bing_auto_run';
  const JOKE_API_URL = 'https://v2.jokeapi.dev/joke/Any?blacklistFlags=nsfw,religious,political,racist,sexist,explicit&type=single';
  const TRIGGER_KEY = 'bing_auto_trigger';
  const HEARTBEAT_KEY = 'bing_auto_heartbeat';
  const KEYWORDS_CACHE_KEY = 'bing_keywords_cache';

  let isRunning = false;
  let isCompleted = false;
  let timerStart = 0;
  let timerInterval = 0;
  let lastKeywordFromPool = true;
  let isDragging = false;
  let dragX = 0, dragY = 0;
  let checkInterval = null;
  let heartbeatInterval = null;
  let currentKeyword = '';
  let nextExecuteTime = 0;  // 下次執行時間戳
  let isBackground = false;
  let scrollInterval = null;  // 滾動間隔計時器
  let scrollTimeout = null;  // 滾動超時計時器

  // 外部詞彙池載入（含 localStorage 緩存，當日有效）
  async function loadExternalKeywords() {
    const today = new Date().toISOString().split('T')[0];
    const cached = localStorage.getItem(KEYWORDS_CACHE_KEY);

    // 檢查緩存（當日有效）
    if (cached) {
      try {
        const cacheData = JSON.parse(cached);
        // 日期相同且有版本號或詞彙池時使用緩存
        if (cacheData.date === today && (cacheData.version || cacheData.keywords)) {
          keywordsPool = cacheData.keywords;
          keywordFixPool = cacheData.keywordFix || keywordFixPool;
          enWordFixPool = cacheData.enWordFix || enWordFixPool;
          return true;
        }
      } catch (e) {
        localStorage.removeItem(KEYWORDS_CACHE_KEY);
      }
    }

    if (!CONFIG.keywordsUrl) return false;

    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 5000);

      const response = await fetch(CONFIG.keywordsUrl, { signal: controller.signal });
      clearTimeout(timeoutId);

      if (!response.ok) return false;

      const data = await response.json();

      // 支援 JSON 格式：{ keywords: [], keywordFix: [], enWordFix: [], version: "1.0" }
      if (data.keywords && Array.isArray(data.keywords) && data.keywords.length > 0) {
        keywordsPool = data.keywords;
      }
      if (data.keywordFix && Array.isArray(data.keywordFix) && data.keywordFix.length > 0) {
        keywordFixPool = data.keywordFix;
      }
      if (data.enWordFix && Array.isArray(data.enWordFix) && data.enWordFix.length > 0) {
        enWordFixPool = data.enWordFix;
      }

      // 存入 localStorage（當日有效）
      localStorage.setItem(KEYWORDS_CACHE_KEY, JSON.stringify({
        date: today,
        version: data.version || null,
        keywords: keywordsPool,
        keywordFix: keywordFixPool,
        enWordFix: enWordFixPool
      }));

      return true;
    } catch (e) {
      return false;
    }
  }

  // 手動強制重新整理外部詞彙池
  function refreshExternalKeywords() {
    localStorage.removeItem(KEYWORDS_CACHE_KEY);
    return loadExternalKeywords();
  }

  // ============================================
  // 初始化
  // ============================================
  function init() {
    // 載入外部詞彙池
    loadExternalKeywords();

    // 跨天檢查
    checkAndResetDay();

    // 嘗試恢復排程狀態
    restoreScheduleState();

    // 檢查自動運行狀態
    const shouldAutoRun = GM_getValue(AUTO_RUN_KEY, false);
    const config = getConfig();

    if (shouldAutoRun) {
      const canRun = canRunSearch(config);
      if (canRun) {
        // 自動開始任務
        setTimeout(() => startSearch(), 1500);
      } else {
        // 任務已完成，等待明日
        GM_setValue(AUTO_RUN_KEY, false);
      }
    }

    initStyles();
    initUI();
    doAutoScroll();

    // 啟動保活機制
    startKeepAlive();

    // 監聽跨分頁消息
    setupCrossTabListener();

    // 頁面載入完成後執行滾動
    if (document.readyState === 'complete') {
      setTimeout(() => {
        if (isRunning && window.location.href.includes('bing.com/search')) {
          doAutoScroll();
        }
      }, 3000);
    } else {
      window.addEventListener('load', () => {
        setTimeout(() => {
          if (isRunning && window.location.href.includes('bing.com/search')) {
            doAutoScroll();
          }
        }, 3000);
      });
    }

    // 持續監測 URL 變化
    let lastUrl = window.location.href;
    let lastIsRunning = isRunning;
    setInterval(() => {
      const currentUrl = window.location.href;
      if (currentUrl !== lastUrl) {
        lastUrl = currentUrl;
        setTimeout(() => {
          if (isRunning && window.location.href.includes('bing.com/search')) {
            doAutoScroll();
          }
        }, 6000);
      }
      // 只在 isRunning 狀態改變時更新 mini-icon 樣式
      if (lastIsRunning !== isRunning) {
        lastIsRunning = isRunning;
        const toolBox = document.getElementById('br_reward_tool');
        const miniIcon = toolBox ? toolBox.querySelector('.br_mini-icon') : null;
        if (miniIcon) {
          if (isRunning) {
            miniIcon.style.background = '#d63031';
            miniIcon.classList.add('running');
          } else {
            miniIcon.style.background = '#0078d4';
            miniIcon.classList.remove('running');
          }
        }
      }
    }, 500);

    // 監聽 hashchange 事件
    window.addEventListener('hashchange', () => {
      setTimeout(() => {
        if (isRunning && window.location.href.includes('bing.com/search')) {
          doAutoScroll();
        }
      }, 6000);
    });

    // 監聽 popstate 事件（瀏覽器導航）
    window.addEventListener('popstate', () => {
      setTimeout(() => {
        if (isRunning && window.location.href.includes('bing.com/search')) {
          doAutoScroll();
        }
      }, 6000);
    });

    // 監聽 visibilitychange 恢復滾動
    document.addEventListener('visibilitychange', () => {
      if (!document.hidden && isRunning && window.location.href.includes('bing.com/search')) {
        doAutoScroll();
      }
    });
  }

  // 恢復排程狀態
  function restoreScheduleState() {
    try {
      const saved = localStorage.getItem('bing_auto_schedule');
      if (saved) {
        const data = JSON.parse(saved);
        const now = Date.now();

        // 如果有排程時間且未過期
        if (data.time > 0 && (now - data.timestamp) < 3600000) {
          nextExecuteTime = data.time;
          timerInterval = data.time - now;
          timerStart = performance.now() - (data.time - data.timestamp);
        }
      }

      // 恢復心跳狀態
      const heartbeat = localStorage.getItem(HEARTBEAT_KEY);
      if (heartbeat) {
        const beat = JSON.parse(heartbeat);
        if (beat.isRunning) {
          GM_setValue(AUTO_RUN_KEY, true);
        }
      }
    } catch (e) { /* 忽略錯誤 */ }
  }

  // ============================================
  // 保活機制 - 背景頁支撐
  // ============================================
  function startKeepAlive() {
    // 檢測頁面可見性變化
    document.addEventListener('visibilitychange', handleVisibilityChange);

    // 頁面載入時檢查是否需要執行
    checkScheduledExecution();

    // 每分鐘檢查一次狀態
    if (checkInterval) clearInterval(checkInterval);
    checkInterval = setInterval(() => {
      checkAndResetDay();
      checkAutoRunFromBackground();
      checkScheduledExecution();
    }, 60000);

    // 發送心跳到其他分頁
    if (heartbeatInterval) clearInterval(heartbeatInterval);
    heartbeatInterval = setInterval(() => {
      try {
        const config = getConfig();
        localStorage.setItem(HEARTBEAT_KEY, JSON.stringify({
          timestamp: Date.now(),
          isRunning: isRunning,
          nextExecuteTime: nextExecuteTime,
          pc_count: config.pc_count,
          ph_count: config.ph_count
        }));
      } catch (e) { /* 忽略錯誤 */ }
    }, 30000);
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

    if (!isRunning || isCompleted) return;

    // 如果有排程的執行時間，且已經到了
    if (scheduledTime > 0 && now >= scheduledTime) {
      nextExecuteTime = 0;
      localStorage.removeItem('bing_auto_schedule');
      performSearch();
      return;
    }

    // 如果沒有排程，但正在運行中，重新計算下次執行時間
    if (isRunning && scheduledTime === 0) {
      // 從計時器狀態計算剩餘時間
      const elapsed = performance.now() - timerStart;
      const remaining = timerInterval - elapsed;

      if (remaining > 0) {
        nextExecuteTime = now + remaining;
        saveScheduleTime(nextExecuteTime);
      }
    }
  }

  function checkAutoRunFromBackground() {
    if (isRunning) return;

    const shouldAutoRun = GM_getValue(AUTO_RUN_KEY, false);
    const config = getConfig();

    if (shouldAutoRun && canRunSearch(config)) {
      startSearch();
    }
  }

  // ============================================
  // 跨分頁通訊
  // ============================================
  function setupCrossTabListener() {
    // 監聽 localStorage 變化
    window.addEventListener('storage', (e) => {
      if (e.key === TRIGGER_KEY && e.newValue) {
        try {
          const data = JSON.parse(e.newValue);
          if (data.action === 'EXECUTE_SEARCH') {
            setTimeout(() => executeSearch(data.keyword || currentKeyword), 1000);
          }
        } catch (err) {}
      }
    });

    // GM_addValueChangeListener
    if (typeof GM_addValueChangeListener !== 'undefined') {
      GM_addValueChangeListener(TRIGGER_KEY, (key, oldValue, newValue, remote) => {
        if (newValue && newValue !== oldValue) {
          try {
            const data = JSON.parse(newValue);
            if (data.action === 'EXECUTE_SEARCH') {
              // 收到 GM 搜尋訊號
              setTimeout(() => executeSearch(data.keyword || currentKeyword), 1000);
            }
          } catch (err) {}
        }
      });
    }
  }

  // ============================================
  // 跨天重置
  // ============================================
  function checkAndResetDay() {
    const stored = getStorageData();
    const today = getToday();

    if (stored && stored.lastDate !== today) {
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

      // 設置自動運行
      GM_setValue(AUTO_RUN_KEY, true);

      updateUI();
    }

    // 凌晨自動開始任務（0-2點）
    const hour = new Date().getHours();
    if (hour >= 0 && hour < 2) {
      const config = getConfig();
      if (config.pc_count === 0 && config.ph_count === 0) {
        GM_setValue(AUTO_RUN_KEY, true);
      }
    }
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
      #br_reward_tool .br_mini-icon.running { animation: breathe 2s ease-in-out infinite; }
      @keyframes breathe { 0% { opacity: 1; box-shadow: 0 0 8px rgba(214, 48, 49, 0.5); } 50% { opacity: 0.6; box-shadow: 0 0 16px rgba(214, 48, 49, 0.8); } 100% { opacity: 1; box-shadow: 0 0 8px rgba(214, 48, 49, 0.5); } }
      #br_reward_tool .br_auto-badge { display: inline-block; background: #27ae60; color: #fff; font-size: 10px; padding: 2px 6px; border-radius: 3px; margin-left: 6px; vertical-align: middle; }
      #br_reward_tool .br_live-indicator { display: inline-block; width: 8px; height: 8px; border-radius: 50%; background: #27ae60; margin-right: 6px; animation: pulse 1.5s infinite; }
      @keyframes pulse { 0% { opacity: 1; } 50% { opacity: 0.5; } 100% { opacity: 1; } }
    `);
  }

  function initUI() {
    const countInfo = getConfig();
    const today = getToday();
    const toolHtml = `
      <div id="br_reward_tool" class="br_minimized">
        <div class="br_header">
          <span class="br_title"><span class="br_live-indicator"></span>隨機搜尋 <span class="br_auto-badge" id="br_auto_badge" style="display:none;">AUTO</span></span>
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
        </div>
        <div class="br_mini-icon">Bing</div>
      </div>
    `;

    if (document.body) {
      document.body.insertAdjacentHTML('beforeend', toolHtml);
    } else {
      window.onload = function () { document.body.insertAdjacentHTML('beforeend', toolHtml); }
    }

    setTimeout(() => {
      const toolBox = document.getElementById('br_reward_tool');
      const toggleBtn = document.getElementById('br_toggle_btn');
      const resetBtn = document.getElementById('br_reset_btn');

      if (!toolBox) return;

      toggleBtn.onclick = () => { toggleScript(toolBox); };
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

      document.onmousemove = (e) => {
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
      };
      document.onmouseup = () => { isDragging = false; toolBox.style.transition = ''; };

      updateStatusAfterInit();
    }, 500);
  }

  function updateStatusAfterInit() {
    const config = getConfig();
    const shouldAutoRun = GM_getValue(AUTO_RUN_KEY, false);
    const canRun = canRunSearch(config);

    if (shouldAutoRun && canRun) {
      updateStatus("腳本運行中...", "#e67e22");
      updateAutoBadge(true);
    } else if (shouldAutoRun && !canRun) {
      updateStatus("任務已完成! 等待明日...", "#27ae60");
      updateCountdownUI("完成");
    }
  }

  // ============================================
  // 核心搜尋邏輯
  // ============================================
  function toggleScript(toolBox) {
    if (!checkLoginStatus()) return;

    const btn = document.getElementById('br_toggle_btn');

    if (isRunning) {
      isRunning = false;
      stopAutoScroll();  // 停止頁面滾動
      GM_setValue(AUTO_RUN_KEY, false);
      btn.textContent = "▶ 繼續搜尋";
      btn.className = "br_btn br_btn_start";
      updateStatus("已暫停", "#666");
      updateCountdownUI("--");
      updateAutoBadge(false);
    } else {
      const config = getConfig();
      const currentPageType = getBingPageType();

      if (currentPageType === 'pc' && config.pc_count >= CONFIG.max_pc) {
        updateStatus("桌面版任務已達標", "#27ae60");
        return;
      }
      if (currentPageType === 'ph' && config.ph_count >= CONFIG.max_ph) {
        updateStatus("行動版任務已達標", "#27ae60");
        return;
      }

      isRunning = true;
      isCompleted = false;
      GM_setValue(AUTO_RUN_KEY, true);
      btn.textContent = "⏸ 暫停搜尋";
      btn.className = "br_btn br_btn_stop";
      updateStatus("腳本運行中...", "#e67e22");
      startSearchLoop();
      updateAutoBadge(true);
    }
  }

  function startSearch() {
    if (!checkLoginStatus()) return;

    const config = getConfig();
    const currentPageType = getBingPageType();

    if (currentPageType === 'pc' && config.pc_count >= CONFIG.max_pc) {
      updateStatus("桌面版任務已達標", "#27ae60");
      GM_setValue(AUTO_RUN_KEY, false);
      updateAutoBadge(false);
      return;
    }
    if (currentPageType === 'ph' && config.ph_count >= CONFIG.max_ph) {
      updateStatus("行動版任務已達標", "#27ae60");
      GM_setValue(AUTO_RUN_KEY, false);
      updateAutoBadge(false);
      return;
    }

    isRunning = true;
    isCompleted = false;
    GM_setValue(AUTO_RUN_KEY, true);

    const btn = document.getElementById('br_toggle_btn');
    if (btn) { btn.textContent = "⏸ 暫停搜尋"; btn.className = "br_btn br_btn_stop"; }

    updateStatus("腳本運行中...", "#e67e22");
    startSearchLoop();
    updateAutoBadge(true);
  }

  function startSearchLoop() {
    if (!isRunning) return;

    const config = getConfig();
    const currentPageType = getBingPageType();

    if (currentPageType === 'pc' && config.pc_count >= CONFIG.max_pc) { onTaskCompleted(); return; }
    if (currentPageType === 'ph' && config.ph_count >= CONFIG.max_ph) { onTaskCompleted(); return; }

    // 使用 performance.now() 計算精確間隔
    timerStart = performance.now();
    timerInterval = getRandomInterval();

    // 記錄下次執行時間（支持背景執行）
    nextExecuteTime = Date.now() + timerInterval;
    saveScheduleTime(nextExecuteTime);

    updateCountdownUI(Math.ceil(timerInterval / 1000));

    // 使用 requestAnimationFrame 實現高精度計時
    requestAnimationFrame(timerLoop);
  }

  let lastSecondUpdate = 0;
  function timerLoop(timestamp) {
    if (!isRunning) return;

    // 每秒更新倒數
    const elapsed = timestamp - timerStart;
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
      updateCountdownUI("正在跳轉...");
      lastSecondUpdate = 0;
      nextExecuteTime = 0;
      saveScheduleTime(0);
      performSearch();
      return;
    }

    // 繼續計時
    requestAnimationFrame(timerLoop);
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
    if (!isRunning) return;

    const config = getConfig();
    const currentPageType = getBingPageType();

    if (currentPageType === 'pc' && config.pc_count >= CONFIG.max_pc) { onTaskCompleted(); return; }
    if (currentPageType === 'ph' && config.ph_count >= CONFIG.max_ph) { onTaskCompleted(); return; }

    let newConfig = { ...config };
    if (currentPageType === 'pc') newConfig.pc_count++;
    else newConfig.ph_count++;
    saveConfig(newConfig);
    updateUI();

    if ((currentPageType === 'pc' && newConfig.pc_count >= CONFIG.max_pc) || (currentPageType === 'ph' && newConfig.ph_count >= CONFIG.max_ph)) {
      onTaskCompleted();
      return;
    }

    getRandomKeyword().then(keyword => {
      currentKeyword = keyword;
      executeSearch(keyword);
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

      setTimeout(() => {
        try {
          if (form) form.submit();
          if (!btn) btn = document.querySelector("button.b_searchboxSubmit") || document.querySelector("a[title='Search']") || document.querySelector(".search_icon");
          if (btn) btn.click();
          if (input) input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
        } catch (e) { /* 忽略錯誤 */ }
      }, 300);

      // 強制跳轉後備
      setTimeout(() => {
        if (isRunning && !window.location.href.includes('bing.com/search?')) {
          window.location.href = 'https://www.bing.com/search?q=' + encodeURIComponent(keyword);
        }
      }, 4000);
    } catch (e) { /* 忽略錯誤 */ }
  }

  function onTaskCompleted() {
    isRunning = false;
    isCompleted = true;
    stopAutoScroll();  // 停止頁面滾動
    GM_setValue(AUTO_RUN_KEY, false);

    const btn = document.getElementById('br_toggle_btn');
    if (btn) { btn.textContent = "▶ 開始搜尋"; btn.className = "br_btn br_btn_start"; }

    updateStatus("任務已完成! 等待明日自動重啟...", "#27ae60");
    updateCountdownUI("完成");
    updateAutoBadge(false);
  }

  // ============================================
  // 工具函數
  // ============================================
  function getToday() { return new Date().toISOString().split('T')[0]; }

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

  function updateAutoBadge(show) {
    const badge = document.getElementById('br_auto_badge');
    if (badge) badge.style.display = show ? 'inline-block' : 'none';
  }

  function cleanCount(toolBox) {
    if (confirm("確定要重置今日的搜尋計數嗎？")) {
      const today = getToday();
      const currentPageType = getBingPageType();
      saveConfig({ date: today, lastDate: today, pc_count: 0, ph_count: 0, deviceType: currentPageType, autoStart: false });
      GM_setValue(AUTO_RUN_KEY, false);
      isRunning = false;
      isCompleted = false;
      stopAutoScroll();  // 停止頁面滾動
      updateUI();
      updateAutoBadge(false);
      updateStatus("等待開始...", "#666");
    }
  }

  function doAutoScroll() {
    if (!window.location.href.includes('bing.com/search')) {
      return;
    }
    if (!isRunning) {
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
    if (!isRunning || !window.location.href.includes('bing.com/search')) {
      stopAutoScroll();
      return;
    }

    // 平滑滾動到底部
    const scrollHeight = Math.max(document.body.scrollHeight, document.documentElement.scrollHeight);
    window.scrollTo({ top: scrollHeight, behavior: 'smooth' });

    // 2 秒後滾回頂部
    scrollTimeout = setTimeout(() => {
      if (isRunning && window.location.href.includes('bing.com/search')) {
        window.scrollTo({ top: 0, behavior: 'smooth' });

        // 滾動完成後，繼續下一次迴圈（間隔 8 秒）
        scrollInterval = setInterval(() => {
          if (!isRunning || !window.location.href.includes('bing.com/search')) {
            stopAutoScroll();
            return;
          }

          // 平滑滾動到底部
          const sh = Math.max(document.body.scrollHeight, document.documentElement.scrollHeight);
          window.scrollTo({ top: sh, behavior: 'smooth' });

          // 2 秒後滾回頂部
          scrollTimeout = setTimeout(() => {
            if (isRunning && window.location.href.includes('bing.com/search')) {
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
    const signInElement = document.querySelector('span.sw_spd.id_avatar#id_a[aria-label="Sign in"]');
    if (signInElement) {
      const computedStyle = window.getComputedStyle(signInElement);
      if (computedStyle.display === 'none' || computedStyle.visibility === 'hidden') return true;
      alert('請登入 Bing 以繼續任務');
      if (isRunning) toggleScript();
      return false;
    }
    return true;
  }

  async function getRandomKeyword() {
    lastKeywordFromPool = !lastKeywordFromPool;
    if (lastKeywordFromPool) return getRandomKeywordFromPool();
    else return await getEnWordKeyword();
  }

  function getRandomKeywordFromPool() {
    const baseKeyword = keywordsPool[Math.floor(Math.random() * keywordsPool.length)];

    // 隨機決定加入 keywordFixPool 的位置（前/後/不加）
    const positionRoll = Math.random();
    if (positionRoll < 0.4) {
      // 40% 機率在前面加入 keywordFixPool
      const fix = keywordFixPool[Math.floor(Math.random() * keywordFixPool.length)];
      return `${fix} ${baseKeyword}`;
    } else if (positionRoll < 0.8) {
      // 40% 機率在後面加入 keywordFixPool
      const fix = keywordFixPool[Math.floor(Math.random() * keywordFixPool.length)];
      return `${baseKeyword} ${fix}`;
    }
    // 20% 機率不加

    return baseKeyword;
  }

  async function getEnWordKeyword() {
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
          // 隨機獲取 1~3 組單字
          const wordCount = Math.floor(Math.random() * 3) + 1;
          const shuffled = validWords.sort(() => 0.5 - Math.random());
          const selectedWords = shuffled.slice(0, wordCount);
          let enWord = selectedWords.join(' ');

          // 隨機決定加入 enWordFixPool 的位置（前或後，或兩者）
          const positionRoll = Math.random();
          if (positionRoll < 0.25) {
            // 只在前面加
            const prefix = enWordFixPool[Math.floor(Math.random() * enWordFixPool.length)];
            enWord = `${prefix} ${enWord}`;
          } else if (positionRoll < 0.5) {
            // 只在後面加
            const fix = enWordFixPool[Math.floor(Math.random() * enWordFixPool.length)];
            enWord = `${enWord} ${fix}`;
          } else if (positionRoll < 0.75) {
            // 兩邊都加（確保前綴和後綴不同）
            const prefixPool = enWordFixPool.filter(f => f !== enWord.split(' ')[0]);
            const fixPool = enWordFixPool.filter(f => f !== enWord.split(' ').slice(-1)[0]);
            const prefix = prefixPool[Math.floor(Math.random() * prefixPool.length)] || enWordFixPool[Math.floor(Math.random() * enWordFixPool.length)];
            const fix = fixPool[Math.floor(Math.random() * fixPool.length)] || enWordFixPool[Math.floor(Math.random() * enWordFixPool.length)];
            enWord = `${prefix} ${enWord} ${fix}`;
          }
          // 25% 機率不加任何前綴/後綴

          return enWord;
        }
      }
    } catch (e) { /* 忽略錯誤 */ }
    return getRandomKeywordFromPool();
  }

  // 啟動
  if (window.location.href.includes('bing.com/search') || window.location.href.includes('bing.com/')) {
    init();

    // 監聽 URL 變化（處理 SPA 頁面跳轉）
    let lastUrl = window.location.href;
    const urlObserver = new MutationObserver(() => {
      if (window.location.href !== lastUrl) {
        lastUrl = window.location.href;
        // 頁面跳轉後延遲執行滾動
        setTimeout(() => {
          if (isRunning && window.location.href.includes('bing.com/search')) {
            doAutoScroll();
          }
        }, 1500);
      }
    });
    urlObserver.observe(document.body, { childList: true, subtree: true });

    // 同時監聽 popstate 事件（瀏覽器導航）
    window.addEventListener('popstate', () => {
      setTimeout(() => {
        if (isRunning && window.location.href.includes('bing.com/search')) {
          doAutoScroll();
        }
      }, 1500);
    });
  }
})();