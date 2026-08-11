### 消耗瀏覽器記憶體
- 每日自動進行搜尋任務，達成任務休息至隔日繼續任務
- 每日自動載入外部搜尋詞彙列表，可在 `CONFIG` 修改
- 每日搜尋上限與隨機秒數可在 `CONFIG` 修改
- 每次搜尋的關鍵字組合與出現權重可在 `CONFIG` 修改
- 每次自動取得隨機搜尋關鍵字
- 每次自動搜尋後，持續的滾動頁面

### 載入外部搜尋詞彙檔案請用 JSON 檔案，格式請參考 example.json
1. keywords: 搜尋關鍵字
2. keywordFix: 在關鍵字前面加入辭彙
3. enWordFix: 在關鍵字後面加入辭彙

```json
{
  "keywords": ["程式教學", "作業系統 環境變數"],
  "keywordFix": ["最新", "資訊"],
  "enWordFix": ["英文", "中文"],
  "checkDateTime": "..."
}
```

---

**※ 僅供技術學習測試，使用即了解並同意自行承擔所有後果**

> From [微软Bing 必应积分自动脚本（Microsoft Bing Rewards Script）](https://greasyfork.org/zh-TW/scripts/532315-%E5%BE%AE%E8%BD%AFbing-%E5%BF%85%E5%BA%94%E7%A7%AF%E5%88%86%E8%87%AA%E5%8A%A8%E8%84%9A%E6%9C%AC-microsoft-bing-rewards-script)

---

<p align="center">
  <img src="https://img.shields.io/badge/GNU-GPLv3-blue?logo=gplv3">
</p>