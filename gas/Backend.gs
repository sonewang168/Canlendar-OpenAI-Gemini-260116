/**
 * ═══════════════════════════════════════════════════════════════
 * 📅 行程日曆助手 - GAS 後端 v2.0
 * ═══════════════════════════════════════════════════════════════
 * 
 * 功能：
 * - 接收前端的排程建立請求
 * - 儲存排程到 Google Sheets
 * - 定時檢查並發送 LINE 提醒
 * - 查詢/清除排程
 * 
 * 部署步驟：
 * 1. 在 Google Apps Script 建立新專案
 * 2. 貼上此程式碼
 * 3. 修改 SECRET_KEY 為你自己的金鑰
 * 4. 部署 → 新增部署 → 網頁應用程式
 * 5. 執行身分：我、誰可存取：所有人
 * 6. 設定觸發器：checkAndSendReminders，每5分鐘
 * 
 * ═══════════════════════════════════════════════════════════════
 */

// ==================== 設定區 ====================

const SHEET_NAME = 'Schedules';           // 試算表名稱
const SECRET_KEY = 'your-secret-key-here'; // 🔐 請修改為你自己的安全金鑰

// ==================== 主要入口 ====================

/**
 * 處理 POST 請求
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    
    // 驗證金鑰
    if (data.secret !== SECRET_KEY) {
      console.log('❌ 金鑰驗證失敗');
      return jsonResponse({ success: false, error: 'Invalid secret key' });
    }
    
    console.log('📥 收到請求:', data.action || 'send');
    
    // 處理不同 action
    switch (data.action) {
      case 'getSchedules':
        return getSchedules();
      case 'clearSchedules':
        return clearSchedules();
      case 'createSchedules':
        return createSchedules(data);
      default:
        // 相容舊版：直接發送通知
        if (data.workshops) {
          return sendWorkshopNotifications(data);
        }
        return jsonResponse({ success: false, error: 'Unknown action' });
    }
  } catch (error) {
    console.error('❌ 錯誤:', error);
    return jsonResponse({ success: false, error: error.message });
  }
}

/**
 * 處理 GET 請求（用於測試）
 */
function doGet(e) {
  return jsonResponse({ 
    success: true, 
    message: '行程日曆助手 GAS 後端 v2.0',
    status: 'running'
  });
}

// ==================== 排程管理 ====================

/**
 * 建立排程
 */
function createSchedules(data) {
  const sheet = getOrCreateSheet();
  const schedules = data.schedules || [];
  let count = 0;
  
  schedules.forEach(schedule => {
    sheet.appendRow([
      new Date(),                         // A: 建立時間
      schedule.workshopTitle,             // B: 研習標題
      schedule.workshopStart,             // C: 開始時間
      schedule.workshopEnd || '',         // D: 結束時間
      schedule.reminderTime,              // E: 提醒時間
      schedule.reminderMinutes,           // F: 提醒分鐘數
      schedule.reminderLabel,             // G: 提醒標籤
      schedule.workshopLocation || '',    // H: 地點
      schedule.meetLink || '',            // I: Meet 連結
      schedule.workshopDescription || '', // J: 描述
      data.token,                         // K: LINE Token
      data.userId,                        // L: LINE User ID
      'pending'                           // M: 狀態
    ]);
    count++;
  });
  
  console.log(`✅ 已建立 ${count} 個排程`);
  return jsonResponse({ success: true, count: count });
}

/**
 * 取得所有排程
 */
function getSchedules() {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  const schedules = [];
  
  // 跳過標題列
  for (let i = 1; i < data.length; i++) {
    schedules.push({
      createdAt: data[i][0],
      workshopTitle: data[i][1],
      workshopStart: data[i][2],
      workshopEnd: data[i][3],
      reminderTime: data[i][4],
      reminderMinutes: data[i][5],
      reminderLabel: data[i][6],
      workshopLocation: data[i][7],
      meetLink: data[i][8],
      status: data[i][12]
    });
  }
  
  console.log(`📋 查詢到 ${schedules.length} 個排程`);
  return jsonResponse({ success: true, schedules: schedules });
}

/**
 * 清除所有排程
 */
function clearSchedules() {
  const sheet = getOrCreateSheet();
  const lastRow = sheet.getLastRow();
  
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
    console.log(`🗑️ 已清除 ${lastRow - 1} 個排程`);
  }
  
  return jsonResponse({ success: true, deleted: lastRow - 1 });
}

// ==================== LINE 通知 ====================

/**
 * 直接發送研習通知（相容舊版）
 */
function sendWorkshopNotifications(data) {
  const workshops = data.workshops;
  const token = data.token;
  const userId = data.userId;
  let successCount = 0;
  
  workshops.forEach(workshop => {
    try {
      const message = formatWorkshopMessage(workshop);
      sendLineMessage(token, userId, message);
      successCount++;
    } catch (error) {
      console.error('發送失敗:', error);
    }
  });
  
  console.log(`📤 已發送 ${successCount}/${workshops.length} 則通知`);
  return jsonResponse({ success: true, sent: successCount });
}

/**
 * 檢查並發送到期的提醒（由觸發器呼叫）
 */
function checkAndSendReminders() {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  let sentCount = 0;
  
  console.log(`⏰ 檢查提醒... 目前時間: ${now.toLocaleString('zh-TW')}`);
  
  for (let i = 1; i < data.length; i++) {
    const status = data[i][12];
    
    if (status === 'pending') {
      const reminderTime = new Date(data[i][4]);
      
      if (now >= reminderTime) {
        try {
          const workshop = {
            title: data[i][1],
            startDateTime: data[i][2],
            endDateTime: data[i][3],
            location: data[i][7],
            meetLink: data[i][8],
            description: data[i][9],
            reminderLabel: data[i][6]
          };
          
          const token = data[i][10];
          const userId = data[i][11];
          
          const message = formatReminderMessage(workshop);
          sendLineMessage(token, userId, message);
          
          // 更新狀態為已發送
          sheet.getRange(i + 1, 13).setValue('sent');
          sheet.getRange(i + 1, 14).setValue(new Date()); // 發送時間
          
          sentCount++;
          console.log(`✅ 已發送: ${workshop.title}`);
        } catch (error) {
          console.error(`❌ 發送失敗 (row ${i + 1}):`, error);
          sheet.getRange(i + 1, 13).setValue('error');
          sheet.getRange(i + 1, 15).setValue(error.message);
        }
      }
    }
  }
  
  if (sentCount > 0) {
    console.log(`📤 本次共發送 ${sentCount} 則提醒`);
  }
}

/**
 * 格式化研習訊息（直接發送用）
 */
function formatWorkshopMessage(workshop) {
  const start = new Date(workshop.startDateTime);
  const dateStr = formatDate(start);
  const timeStr = formatTime(start);
  const weekday = getWeekday(start);
  
  let msg = `📚 研習通知\n`;
  msg += `━━━━━━━━━━━━━━\n`;
  msg += `📌 ${workshop.title}\n`;
  msg += `📅 ${dateStr} (${weekday})\n`;
  msg += `⏰ ${timeStr}\n`;
  
  if (workshop.location) {
    msg += `📍 ${workshop.location}\n`;
  }
  
  // 檢測 Meet 連結
  const meetLink = extractMeetLink(workshop.location, workshop.description);
  if (meetLink) {
    msg += `\n🔗 會議連結：\n${meetLink}`;
  }
  
  return msg;
}

/**
 * 格式化提醒訊息（排程發送用）
 */
function formatReminderMessage(workshop) {
  const start = new Date(workshop.startDateTime);
  const dateStr = formatDate(start);
  const timeStr = formatTime(start);
  const weekday = getWeekday(start);
  
  let msg = `🔔 研習提醒 (${workshop.reminderLabel})\n`;
  msg += `━━━━━━━━━━━━━━\n`;
  msg += `📌 ${workshop.title}\n`;
  msg += `📅 ${dateStr} (${weekday})\n`;
  msg += `⏰ ${timeStr}\n`;
  
  if (workshop.location) {
    msg += `📍 ${workshop.location}\n`;
  }
  
  // 優先使用已提取的 meetLink
  const meetLink = workshop.meetLink || extractMeetLink(workshop.location, workshop.description);
  if (meetLink) {
    msg += `\n🔗 會議連結：\n${meetLink}`;
  }
  
  return msg;
}

/**
 * 發送 LINE 訊息
 */
function sendLineMessage(token, userId, message) {
  const url = 'https://api.line.me/v2/bot/message/push';
  
  const payload = {
    to: userId,
    messages: [{
      type: 'text',
      text: message
    }]
  };
  
  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + token
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  
  if (responseCode !== 200) {
    throw new Error(`LINE API 錯誤: ${responseCode} - ${response.getContentText()}`);
  }
  
  return true;
}

// ==================== 輔助函數 ====================

/**
 * 取得或建立試算表
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // 設定標題列
    sheet.appendRow([
      '建立時間',      // A
      '研習標題',      // B
      '開始時間',      // C
      '結束時間',      // D
      '提醒時間',      // E
      '提醒分鐘',      // F
      '提醒標籤',      // G
      '地點',          // H
      'Meet連結',     // I
      '描述',          // J
      'Token',        // K
      'UserId',       // L
      '狀態',          // M
      '發送時間',      // N
      '錯誤訊息'       // O
    ]);
    
    // 凍結標題列
    sheet.setFrozenRows(1);
    
    // 設定欄寬
    sheet.setColumnWidth(1, 150);  // 建立時間
    sheet.setColumnWidth(2, 200);  // 標題
    sheet.setColumnWidth(13, 80);  // 狀態
    
    console.log('📊 已建立新試算表');
  }
  
  return sheet;
}

/**
 * 格式化日期
 */
function formatDate(date) {
  const month = date.getMonth() + 1;
  const day = date.getDate();
  return `${month}/${day}`;
}

/**
 * 格式化時間
 */
function formatTime(date) {
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  return `${hours}:${minutes}`;
}

/**
 * 取得星期幾
 */
function getWeekday(date) {
  const weekdays = ['週日', '週一', '週二', '週三', '週四', '週五', '週六'];
  return weekdays[date.getDay()];
}

/**
 * 提取 Google Meet 連結
 */
function extractMeetLink(location, description) {
  const text = (location || '') + ' ' + (description || '');
  const match = text.match(/https?:\/\/meet\.google\.com\/[a-z\-]+/i);
  return match ? match[0] : null;
}

/**
 * 回傳 JSON 格式
 */
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==================== 測試函數 ====================

/**
 * 測試發送 LINE 訊息（手動執行用）
 */
function testSendLine() {
  const token = '你的 LINE Channel Access Token';
  const userId = '你的 LINE User ID';
  const message = '🧪 測試訊息\n\n這是來自 GAS 後端的測試訊息。';
  
  try {
    sendLineMessage(token, userId, message);
    console.log('✅ 測試訊息發送成功！');
  } catch (error) {
    console.error('❌ 發送失敗:', error);
  }
}

/**
 * 手動觸發檢查提醒（測試用）
 */
function manualCheckReminders() {
  console.log('🔧 手動執行提醒檢查...');
  checkAndSendReminders();
  console.log('✅ 檢查完成');
}

