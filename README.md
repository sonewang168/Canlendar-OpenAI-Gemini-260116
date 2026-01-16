# 📅 行程日曆助手 Mobile v3.3.0

> 雙軌 AI 解析系統 - Gemini + OpenAI 並行比對，智能解析研習/畢旅/活動行程

![Version](https://img.shields.io/badge/version-3.3.0-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Platform](https://img.shields.io/badge/platform-Web%20%7C%20Mobile-orange.svg)

## ✨ 功能特色

### 🔀 雙軌 AI 解析（v3.3.0 新功能）
- **Gemini + OpenAI 並行解析**：同時調用兩個 AI，互相驗證結果
- **視覺化比對介面**：並排顯示兩個結果，輕鬆發現差異
- **自由選擇**：選擇較正確的結果匯入日曆
- **共用 LINE BOT**：不管選哪個 AI，都用同一個 LINE BOT 發送提醒

### 📝 多元輸入方式
- **文字輸入**：直接貼上研習公文、活動行程
- **檔案上傳**：支援 Excel (.xlsx/.xls)、CSV、TXT、PDF
- **圖片 OCR**：拍照或上傳圖片，AI 智能辨識

### 📆 日曆整合
- **Google Calendar**：一鍵匯入 Google 日曆（OAuth 2.0 授權）
- **ICS 下載**：匯出標準 ICS 檔案，支援所有日曆軟體
- **自訂提醒**：30分鐘前 / 1小時前 / 2小時前 / 1天前

### 📱 LINE 通知
- **排程提醒**：研習開始前自動推播 LINE 訊息
- **GAS 後端**：透過 Google Apps Script 處理排程
- **Meet 連結偵測**：自動提取 Google Meet 連結

### 🎨 優質 UI/UX
- **冷光科技風**：專業的深色主題設計
- **炫酷開機動畫**：多層軌道粒子效果
- **響應式設計**：完美支援手機與桌面
- **PWA 支援**：可添加到主畫面使用

---

## 📸 截圖預覽

```
┌─────────────────────────────────────┐
│  📅 行程助手          ❓  ⚙️        │
├─────────────────────────────────────┤
│  ┌─────────┐  ┌─────────┐          │
│  │    0    │  │    0    │          │
│  │待匯入事件│  │已上傳檔案│          │
│  └─────────┘  └─────────┘          │
│                                     │
│  ┌────────┐  ┌────────┐            │
│  │  ✏️   │  │  📄   │            │
│  │文字輸入│  │上傳檔案│            │
│  └────────┘  └────────┘            │
│  ┌────────┐  ┌────────┐            │
│  │  📸   │  │  🔑   │            │
│  │拍照辨識│  │API設定│            │
│  └────────┘  └────────┘            │
│                                     │
│  🟣 Gemini 已啟用  🟢 OpenAI 已啟用  │
├─────────────────────────────────────┤
│  🏠     📝     📱     ⚙️           │
│ 首頁   輸入   LINE   設定          │
└─────────────────────────────────────┘
```

---

## 🚀 快速開始

### 方式一：直接使用
1. 下載 `行程日曆助手_Mobile_v3.3.0_雙軌AI.html`
2. 用瀏覽器開啟（支援 Chrome、Safari、Edge）
3. 設定 API Key 後即可使用

### 方式二：部署到網站
```bash
# 上傳到任何靜態網站託管服務
# 例如：Netlify、Vercel、GitHub Pages
```

---

## 🔑 API 設定教學

### Gemini API（Google AI）

1. 前往 [Google AI Studio](https://aistudio.google.com/)
2. 點擊「Get API key」
3. 建立新的 API Key
4. 複製貼到 App 設定頁面的「Gemini API Key」

### OpenAI API

1. 前往 [OpenAI Platform](https://platform.openai.com/)
2. 登入後點擊「API Keys」
3. 點擊「Create new secret key」
4. 複製貼到 App 設定頁面的「OpenAI API Key」

> 💡 **提示**：至少設定一個 API Key 即可使用，建議兩個都設定以啟用雙軌比對功能

### Google OAuth（日曆匯入）

1. 前往 [Google Cloud Console](https://console.cloud.google.com/)
2. 建立新專案或選擇現有專案
3. 啟用「Google Calendar API」
4. 建立 OAuth 2.0 用戶端 ID（應用程式類型：網頁應用程式）
5. 在「已授權的 JavaScript 來源」添加你的網域
6. 複製 Client ID 貼到 App 設定頁面

---

## 📱 LINE 通知設定

### 步驟一：建立 LINE Bot

1. 前往 [LINE Developers Console](https://developers.line.biz/)
2. 建立新的 Provider（如果沒有）
3. 建立新的 Messaging API Channel
4. 在「Messaging API」頁籤取得：
   - **Channel Access Token**（點擊「Issue」產生）
   - **Your User ID**（在「Basic settings」頁籤）

### 步驟二：部署 Google Apps Script

1. 前往 [Google Apps Script](https://script.google.com/)
2. 建立新專案
3. 貼上以下程式碼：

```javascript
// GAS 排程提醒後端 v2.0
const SHEET_NAME = 'Schedules';
const SECRET_KEY = '你的安全金鑰'; // 請自行設定

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    
    // 驗證金鑰
    if (data.secret !== SECRET_KEY) {
      return jsonResponse({ success: false, error: 'Invalid secret key' });
    }
    
    // 處理不同 action
    if (data.action === 'getSchedules') {
      return getSchedules();
    } else if (data.action === 'clearSchedules') {
      return clearSchedules();
    } else if (data.action === 'createSchedules') {
      return createSchedules(data);
    } else if (data.workshops) {
      // 相容舊版：直接發送
      return sendWorkshopNotifications(data);
    }
    
    return jsonResponse({ success: false, error: 'Unknown action' });
  } catch (error) {
    return jsonResponse({ success: false, error: error.message });
  }
}

function createSchedules(data) {
  const sheet = getOrCreateSheet();
  const schedules = data.schedules || [];
  let count = 0;
  
  schedules.forEach(schedule => {
    sheet.appendRow([
      new Date(),                    // 建立時間
      schedule.workshopTitle,        // 研習標題
      schedule.workshopStart,        // 開始時間
      schedule.reminderTime,         // 提醒時間
      schedule.reminderLabel,        // 提醒標籤
      schedule.workshopLocation,     // 地點
      schedule.meetLink || '',       // Meet 連結
      data.token,                    // LINE Token
      data.userId,                   // LINE User ID
      'pending'                      // 狀態
    ]);
    count++;
  });
  
  return jsonResponse({ success: true, count: count });
}

function getSchedules() {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  const schedules = [];
  
  for (let i = 1; i < data.length; i++) {
    schedules.push({
      workshopTitle: data[i][1],
      workshopStart: data[i][2],
      reminderTime: data[i][3],
      reminderLabel: data[i][4],
      status: data[i][9]
    });
  }
  
  return jsonResponse({ success: true, schedules: schedules });
}

function clearSchedules() {
  const sheet = getOrCreateSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
  return jsonResponse({ success: true });
}

function sendWorkshopNotifications(data) {
  const workshops = data.workshops;
  const token = data.token;
  const userId = data.userId;
  
  workshops.forEach(workshop => {
    const message = formatWorkshopMessage(workshop);
    sendLineMessage(token, userId, message);
  });
  
  return jsonResponse({ success: true });
}

function formatWorkshopMessage(workshop) {
  const start = new Date(workshop.startDateTime);
  const dateStr = `${start.getMonth() + 1}/${start.getDate()}`;
  const timeStr = `${String(start.getHours()).padStart(2, '0')}:${String(start.getMinutes()).padStart(2, '0')}`;
  
  let msg = `📚 研習提醒\n\n`;
  msg += `📌 ${workshop.title}\n`;
  msg += `📅 ${dateStr} ${timeStr}\n`;
  if (workshop.location) msg += `📍 ${workshop.location}\n`;
  
  return msg;
}

function sendLineMessage(token, userId, message) {
  const url = 'https://api.line.me/v2/bot/message/push';
  const payload = {
    to: userId,
    messages: [{ type: 'text', text: message }]
  };
  
  UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + token
    },
    payload: JSON.stringify(payload)
  });
}

function checkAndSendReminders() {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][9] === 'pending') {
      const reminderTime = new Date(data[i][3]);
      if (now >= reminderTime) {
        const workshop = {
          title: data[i][1],
          startDateTime: data[i][2],
          location: data[i][5]
        };
        const token = data[i][7];
        const userId = data[i][8];
        
        sendLineMessage(token, userId, formatWorkshopMessage(workshop));
        sheet.getRange(i + 1, 10).setValue('sent');
      }
    }
  }
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['建立時間', '標題', '開始時間', '提醒時間', '提醒標籤', '地點', 'Meet連結', 'Token', 'UserId', '狀態']);
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
```

4. 點擊「部署」→「新增部署」
5. 選擇「網頁應用程式」
6. 設定「執行身分」為「我」
7. 設定「誰可以存取」為「所有人」
8. 點擊「部署」並複製網址

### 步驟三：設定定時觸發器

1. 在 GAS 編輯器中點擊「觸發條件」（時鐘圖示）
2. 新增觸發條件：
   - 選擇函式：`checkAndSendReminders`
   - 事件來源：時間驅動
   - 時間型觸發條件類型：分鐘計時器
   - 間隔：每 5 分鐘

---

## 📋 版本歷史

### v3.3.0（2025-01-17）🆕
- ✨ 新增雙軌 AI 解析系統（Gemini + OpenAI）
- ✨ 並排比對介面，可視化選擇
- ✨ OpenAI API 整合（GPT-4o-mini）
- 🔧 共用 LINE BOT，不浪費點數

### v3.2.3（2025-01-17）
- 🐛 修復西元年份被錯誤轉換的問題
- 🔧 強化年份提取邏輯（中文環境優化）
- 🔧 閉包保存原始年份確保修正生效

### v3.2.2（2025-01-16）
- 🔧 修復民國年轉換問題
- 🔧 增加 Debug Log 方便追蹤

### v3.2.0（2025-01-16）
- ✨ Google OAuth Token 快取（24小時有效）
- ✨ LINE 排程提醒系統
- ✨ GAS 後端整合

### v3.0.0（2025-01-15）
- 🎨 全新 UI 設計（冷光科技風）
- ✨ 炫酷開機動畫
- ✨ 圖片 OCR 辨識
- ✨ Excel/PDF 檔案支援

---

## 🏗️ 技術架構

```
┌─────────────────────────────────────────────────────────┐
│                    前端（單一 HTML 檔案）                 │
├─────────────────────────────────────────────────────────┤
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐     │
│  │   Gemini    │  │   OpenAI    │  │  Google     │     │
│  │   API       │  │   API       │  │  Calendar   │     │
│  │  (解析)     │  │  (解析)     │  │  API        │     │
│  └──────┬──────┘  └──────┬──────┘  └──────┬──────┘     │
│         │                │                │            │
│         └────────┬───────┘                │            │
│                  ▼                        ▼            │
│         ┌───────────────┐        ┌───────────────┐    │
│         │  雙軌比對     │        │  OAuth 2.0   │    │
│         │  選擇介面     │        │  授權流程     │    │
│         └───────┬───────┘        └───────────────┘    │
│                 │                                      │
│                 ▼                                      │
│         ┌───────────────┐                             │
│         │  LINE 通知    │◄─────── GAS 後端            │
│         │  排程系統     │         (Google Sheets)     │
│         └───────────────┘                             │
└─────────────────────────────────────────────────────────┘
```

### 使用的技術
- **前端**：原生 HTML/CSS/JavaScript（無框架）
- **AI API**：Google Gemini 2.0 Flash、OpenAI GPT-4o-mini
- **日曆**：Google Calendar API v3
- **通知**：LINE Messaging API
- **後端**：Google Apps Script + Google Sheets
- **函式庫**：
  - SheetJS (xlsx.js) - Excel 解析
  - PDF.js - PDF 解析

---

## 📁 檔案結構

```
📦 行程日曆助手
├── 📄 行程日曆助手_Mobile_v3.3.0_雙軌AI.html  # 主程式（雙軌版）
├── 📄 行程日曆助手_Mobile_v3.2.3.html         # 單軌版（備用）
├── 📄 README.md                               # 本文件
└── 📄 GAS_Backend.gs                          # GAS 後端程式碼
```

---

## ❓ 常見問題

### Q: 為什麼年份會跑掉？
A: v3.2.3+ 已修復此問題。如果仍有問題，請確認：
1. 輸入的年份格式正確（如「2026年1月18日」）
2. 使用最新版本的 App

### Q: 雙軌解析哪個比較準？
A: 視情況而定：
- **Gemini**：對中文格式支援較好
- **OpenAI**：邏輯推理較穩定
- 建議兩個都試，選擇較正確的

### Q: LINE 通知沒收到？
A: 請檢查：
1. LINE Bot 的 Channel Access Token 是否正確
2. User ID 是否正確（不是 LINE ID）
3. GAS 部署是否設定為「所有人可存取」
4. GAS 觸發器是否正確設定

### Q: 可以只用一個 AI 嗎？
A: 可以！只設定一個 API Key 就會進入單軌模式。

---

## 📄 授權條款

MIT License

Copyright (c) 2025

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

---

## 🙏 致謝

- [Google AI Studio](https://aistudio.google.com/) - Gemini API
- [OpenAI](https://openai.com/) - GPT API
- [LINE Developers](https://developers.line.biz/) - Messaging API
- [SheetJS](https://sheetjs.com/) - Excel 解析
- [PDF.js](https://mozilla.github.io/pdf.js/) - PDF 解析

---

<p align="center">
  Made with ❤️ for Taiwan Teachers
</p>
