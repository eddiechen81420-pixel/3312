// ================= 全域變數與系統配置 =================
const props = PropertiesService.getScriptProperties();
const GEMINI_API_KEY = props.getProperty('GEMINI_API_KEY');
const SHEET_ID = props.getProperty('SHEET_ID');
const FOLDER_ID = props.getProperty('FOLDER_ID');
const LINE_ACCESS_TOKEN = props.getProperty('LINE_ACCESS_TOKEN');
const EDDIE_LINE_USER_ID = props.getProperty('EDDIE_LINE_USER_ID'); // Eddie 的 LINE ID (審核用)
const ALUN_LINE_USER_ID = props.getProperty('ALUN_LINE_USER_ID');   // 阿倫的 LINE ID (催報與對齊用)

const MIN_API_INTERVAL = 4000; 
const ALLOWED_EXPENSE_CATEGORIES = ['食材', '水電', '包材', '雜支'];
const GEMINI_MODEL = 'gemini-2.5-flash-preview-09-2025'; // 採用此環境最穩定的 2.5 Flash 預覽版

// ================= Webhook 核心分流中樞 =================
function doPost(e) {
  try {
    if (!e || !e.postData) return ContentService.createTextOutput("No data");
    const data = JSON.parse(e.postData.contents);
    if (!data.events || data.events.length === 0) return ContentService.createTextOutput("No events");
    
    // 共用同一試算表連線以降低 API I/O 損耗
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    
    data.events.forEach(event => {
      const replyToken = event.replyToken;
      const cache = CacheService.getScriptCache();
      const userId = event.source.userId;
      const chatId = event.source.groupId || event.source.roomId || userId;
      
      // 1. 處理 Postback 事件 (Eddie 的行動審核卡片點擊)
      if (event.type === 'postback') {
        const postbackData = event.postback.data;
        // Postback 防重入鎖：防止 LINE 逾時重試或使用者高頻重複點擊
        const postbackKey = `PB_${userId}_${postbackData}`;
        if (cache.get(postbackKey)) {
          console.warn(`[偵測到重複 Postback 點擊] User: ${userId}, Data: ${postbackData}`);
          return;
        }
        cache.put(postbackKey, '1', 60); // 鎖定 60 秒
        
        handlePostbackEvent(event, replyToken, spreadsheet);
        return;
      }
      
      if (event.type !== 'message') return;
      const id = event.message.id;
      // 訊息防重入鎖
      if (cache.get(id)) return; 
      cache.put(id, '1', 300);   
      
      // 2. 處理文字訊息 (開班點鈔、打烊點鈔等純文字)
      if (event.message.type === 'text') {
        processTextEvent(event.message.text.trim(), replyToken, userId, chatId, spreadsheet);
      } 
      // 3. 處理圖片訊息 (POS日結單、支出單據)
      else if (event.message.type === 'image') {
        if (chatId) showLoadingAnimation(chatId, 25);
        processImageEvent(id, replyToken, userId, spreadsheet);
      } 
      // 4. 處理語音訊息 (支出語音交代)
      else if (event.message.type === 'audio') {
        if (chatId) showLoadingAnimation(chatId, 30);
        processAudioEvent(id, replyToken, userId, spreadsheet);
      }
    });
    
    return ContentService.createTextOutput("OK");
  } catch (error) {
    console.error('doPost 核心異常 [全域中斷]:', error.toString(), error.stack);
    return ContentService.createTextOutput("Error: " + error.message);
  }
}

// ================= 核心業務邏輯處理模組 =================

// 純文字事件分流
function processTextEvent(text, replyToken, userId, chatId, spreadsheet) {
  // 開班點鈔判定 (匹配純數字)
  if (/^\d+$/.test(text)) {
    const amount = parseInt(text, 10);
    const dateStr = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd');
    const sheet = spreadsheet.getSheetByName("營收與結帳");
    if (!sheet) { 
      console.error(`[工作表缺失] 營收與結帳 (操作者: ${userId})`);
      replyLine(replyToken, "❌ 系統錯誤：找不到『營收與結帳』工作表，請聯絡系統管理員。"); 
      return; 
    }
    
    // 檢查今日是否已建立紀錄，避免重複開班
    const lastRow = sheet.getLastRow();
    let targetRow = lastRow + 1;
    if (lastRow > 1) {
      const lastDateVal = sheet.getRange(lastRow, 1).getValue();
      const lastDate = lastDateVal instanceof Date ? Utilities.formatDate(lastDateVal, 'Asia/Taipei', 'yyyy/MM/dd') : "";
      if (lastDate === dateStr) targetRow = lastRow; // 今日已有打烊或開班，覆寫
    }
    
    sheet.getRange(targetRow, 1).setValue(dateStr); // A: 日期
    sheet.getRange(targetRow, 3).setValue(amount);  // C: 開班金額
    
    replyLine(replyToken, `💰 已為您記錄今日【開班金額】：$${amount.toLocaleString()} 元。\n祝今天真傳好食生意興隆！`);
    return;
  }
  
  // 打烊點鈔判定 (格式：打烊 15800)
  if (text.startsWith("打烊")) {
    const amountStr = text.replace(/^打烊\s*/, "").trim();
    if (/^\d+$/.test(amountStr)) {
      const amount = parseInt(amountStr, 10);
      const dateStr = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd');
      const sheet = spreadsheet.getSheetByName("營收與結帳");
      if (!sheet) { 
        console.error(`[工作表缺失] 營收與結帳 (操作者: ${userId})`);
        replyLine(replyToken, "❌ 系統錯誤：找不到『營收與結帳』工作表。"); 
        return; 
      }
      
      const row = findRowByDate(sheet, dateStr, 1);
      if (row > 0) {
        sheet.getRange(row, 4).setValue(amount); // D: 打烊點鈔金額
        replyLine(replyToken, `🏁 已更新今日【打烊點鈔金額】：$${amount.toLocaleString()} 元。\n請記得上傳 POS 日結單照片以完成最終營收比對。`);
      } else {
        // 若找不到今日開班紀錄，直接新建一列
        const nextRow = sheet.getLastRow() + 1;
        sheet.getRange(nextRow, 1).setValue(dateStr);
        sheet.getRange(nextRow, 4).setValue(amount);
        replyLine(replyToken, `⚠️ 偵測到今日未錄入開班金額。已直接為您建立今日【打烊點鈔金額】：$${amount.toLocaleString()} 元，請補上 POS 日結單照片。`);
      }
    } else {
      replyLine(replyToken, "❌ 格式錯誤。請輸入「打烊 數字」，例如：打烊 18500");
    }
    return;
  }
}

// 圖片事件處理 (POS 日結單 or 支出發票收據)
function processImageEvent(messageId, replyToken, userId, spreadsheet) {
  const blob = getLineContent(messageId);
  if (!blob) { replyLine(replyToken, "❌ 無法下載單據圖片，請重試。"); return; }
  
  // 先判定是否為 POS 日結單
  const isPosCheckPrompt = "請判定這張圖片是一間小吃店的『POS機日結單/當日結帳單報告』，還是一般的『進貨發票/購買收據/免用統一發票收據』？只需回傳 JSON: {\"type\": \"POS\"} 或 {\"type\": \"RECEIPT\"}";
  
  let typeRes;
  try {
    typeRes = callGeminiApiWithRetry(blob, isPosCheckPrompt, true);
  } catch (err) {
    console.error(`[Gemini POS篩選失敗] User: ${userId}`, err);
    replyLine(replyToken, "⚠️ " + err.message);
    return;
  }
  
  if (typeRes && typeRes.type === "POS") {
    // 執行 POS 數據結構化萃取
    const posPrompt = "你是一位精確的財務會計。請分析此張小吃店 POS 日結單，提取『當日營業總額/總銷售金額』。請排除作廢與退組，只拿淨營業額。只回傳純數字之 JSON 結構：{\"posTotal\": 25800}";
    let posData;
    try {
      posData = callGeminiApiWithRetry(blob, posPrompt, true);
    } catch (err) {
      console.error(`[Gemini POS數據萃取失敗] User: ${userId}`, err);
      replyLine(replyToken, "⚠️ POS 辨識失敗：" + err.message);
      return;
    }
    
    if (!posData || !posData.posTotal) { replyLine(replyToken, "⚠️ POS 日結單解析失敗，請確保拍攝清晰並重傳。"); return; }
    
    const dateStr = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd');
    const driveUrl = saveFileToDrive(blob, "日結單", dateStr);
    
    const sheet = spreadsheet.getSheetByName("營收與結帳");
    const row = findRowByDate(sheet, dateStr, 1);
    
    if (row > 0) {
      sheet.getRange(row, 2).setValue(posData.posTotal); // B: POS總額
      sheet.getRange(row, 6).setValue(driveUrl);        // F: 單據連結
    } else {
      const nextRow = sheet.getLastRow() + 1;
      sheet.getRange(nextRow, 1).setValue(dateStr);
      sheet.getRange(nextRow, 2).setValue(posData.posTotal);
      sheet.getRange(nextRow, 6).setValue(driveUrl);
    }
    replyLine(replyToken, `📈 成功辨識 POS 日結單營業額：$${posData.posTotal.toLocaleString()} 元，單據已封存至雲端硬碟。`);
  } else {
    // 視為一般支出單據處理
    const receiptPrompt = `分析此進貨收據發票。規則：1.類別嚴格限制只能是 '食材'、'水電'、'包材'、'雜支' 其中之一。2.品項描述精簡至15字內。3.金額為純數字。結構：{"category":"類別","description":"品項描述","amount":150,"memo":"收據上其他未覆蓋的備註內容"}`;
    let receiptData;
    try {
      receiptData = callGeminiApiWithRetry(blob, receiptPrompt, true);
    } catch (err) {
      console.error(`[Gemini 收據辨識失敗] User: ${userId}`, err);
      replyLine(replyToken, "⚠️ 收據辨識失敗：" + err.message);
      return;
    }
    
    if (!receiptData || !receiptData.amount) { replyLine(replyToken, "⚠️ 無法辨識此支出單據之金額，請補拍或用語音說明。"); return; }
    
    executeExpensePipeline(receiptData, blob, "圖片單據", replyToken, spreadsheet);
  }
}

// 語音事件處理 (阿倫口頭交代開銷)
function processAudioEvent(messageId, replyToken, userId, spreadsheet) {
  const blob = getLineContent(messageId);
  if (!blob) { replyLine(replyToken, "❌ 語音檔案下載失敗。"); return; }
  
  const audioPrompt = `你是一位精確的採購秘書。請傾聽語音內容，提煉出開銷數據。
  規則：
  1. 類別必須嚴格對齊這四個分類：'食材'、'水電'、'包材'、'雜支'（若語意模糊請歸類到雜支）。
  2. 品項描述請去除口語，精簡至15字內（例如：'叫高麗菜3箱'）。
  3. 金額只保留純整數數字。
  4. memo 欄位請完整保留語音的逐字稿內容，不可刪減任何字。
  結構：{"category":"類別","description":"品項描述","amount":500,"memo":"完整語音逐字稿"}`;
  
  let expenseData;
  try {
    expenseData = callGeminiApiWithRetry(blob, audioPrompt, true);
  } catch (err) {
    console.error(`[Gemini 語音辨識失敗] User: ${userId}`, err);
    replyLine(replyToken, "⚠️ 語音辨識失敗：" + err.message);
    return;
  }
  
  if (!expenseData || !expenseData.amount) { replyLine(replyToken, "⚠️ 語音內容未包含明確的開銷金額或項目，請重新說明。"); return; }
  
  executeExpensePipeline(expenseData, blob, "語音備忘", replyToken, spreadsheet);
}

// ================= 支出管線與行動端審核模組 =================
function executeExpensePipeline(data, blob, typeLabel, replyToken, spreadsheet) {
  const dateStr = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd');
  const fileUrl = saveFileToDrive(blob, typeLabel, dateStr);
  
  const sheet = spreadsheet.getSheetByName("支出明細");
  if (!sheet) { 
    console.error(`[工作表缺失] 支出明細 (操作者: ${replyToken})`);
    replyLine(replyToken, "❌ 系統錯誤：找不到『支出明細』工作表，請聯絡系統管理員。"); 
    return; 
  }
  
  const lastRow = sheet.getLastRow();
  
  // 生成安全唯一識別碼 (EXP-yyyyMMdd-XXX)
  const todayPrefix = `EXP-${Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd')}-`;
  let currentSeq = 1;
  if (lastRow > 1) {
    const allIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const todayIds = allIds.filter(id => String(id).startsWith(todayPrefix));
    currentSeq = todayIds.length + 1;
  }
  const expId = todayPrefix + String(currentSeq).padStart(3, '0');
  
  // 欄位防呆：過濾類別全形/半形空白，確保 100% 寫入有效值
  let category = String(data.category || "").trim().replace(/[\s　]+/g, "");
  category = ALLOWED_EXPENSE_CATEGORIES.includes(category) ? category : "雜支";
  
  const desc = data.description || "未命名項目";
  const amount = Number(data.amount);
  if (!Number.isFinite(amount)) {
    replyLine(replyToken, "⚠️ 金額格式異常，請重試。");
    return;
  }

  const rowData = [expId, dateStr, category, desc, amount, fileUrl, data.memo || "", false];
  sheet.appendRow(rowData);
  
  // 回覆阿倫
  replyLine(replyToken, `📝 已幫您錄入一筆【${category}】支出預覽：${desc} $${amount.toLocaleString()} 元。已發送審核卡片給 Eddie 核對。`);
  
  // 推送高級 Flex 輕量審核卡片給 Eddie
  if (EDDIE_LINE_USER_ID) {
    const sheetGid = sheet.getSheetId(); // 動態獲取支出明細的分頁 GID，避免手動對齊出錯
    pushFlexMessage(EDDIE_LINE_USER_ID, `🛒 支出審核提示: $${amount.toLocaleString()}`, buildApprovalFlexCard(expId, dateStr, category, desc, amount, fileUrl, data.memo || "無", sheetGid));
  }
}

// 處理 Eddie 點擊「確認核准」的 Postback 事件
function handlePostbackEvent(event, replyToken, spreadsheet) {
  const postbackData = event.postback.data;
  if (!postbackData.startsWith("action=approve")) return;
  
  // 提取唯一識別碼：收緊 Regex 符合 EXP-yyyyMMdd-XXX 規則
  const match = postbackData.match(/id=(EXP-\d{8}-\d{3})/);
  if (!match) {
    console.error(`[Postback 惡意或錯誤格式識別碼] Data: ${postbackData}`);
    replyLine(replyToken, "❌ 審核失敗：卡片識別碼格式不合法。");
    return;
  }
  const expId = match[1];
  
  const sheet = spreadsheet.getSheetByName("支出明細");
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) { replyLine(replyToken, "❌ 找不到任何明細數據。"); return; }
  
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const index = ids.indexOf(expId);
  
  if (index !== -1) {
    const targetRow = index + 2; // 加上 Header 與 Index 偏移
    sheet.getRange(targetRow, 8).setValue(true); // H欄：將審核狀態變更為 TRUE
    
    // 變更成功後，前端即時反饋
    replyLine(replyToken, `✅ 審核完成！\n編號：${expId} 的支出項目已核准通過，後台損益表已同步動態連動。`);
  } else {
    console.error(`[審核失敗-ID丟失] EXP ID: ${expId} 未在 Sheet 內尋獲。`);
    replyLine(replyToken, `❌ 審核失敗：在資料庫中找不到編號為 ${expId} 的支出項目，可能已被手動刪除。`);
  }
}

// ================= LINE Flex 卡片視覺渲染引擎 =================
function buildApprovalFlexCard(expId, date, category, desc, amount, driveUrl, memo, sheetGid) {
  const sheetUrl = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit#gid=${sheetGid}`; // 採用動態 GID 直達特定分頁
  return {
    "type": "bubble",
    "size": "mega",
    "header": {
      "type": "box",
      "layout": "vertical",
      "backgroundColor": "#FF9800", // 待審核使用警示橘色
      "paddingAll": "md",
      "contents": [
        { "type": "text", "text": "🛒 真傳好食 ‧ 支出明細審核", "color": "#ffffff", "weight": "bold", "size": "sm" },
        { "type": "text", "text": `單號：${expId}`, "color": "#ffe0b2", "size": "xxs", "margin": "xs" }
      ]
    },
    "body": {
      "type": "box",
      "layout": "vertical",
      "spacing": "md",
      "contents": [
        {
          "type": "box",
          "layout": "horizontal",
          "contents": [
            { "type": "text", "text": "金額", "size": "sm", "color": "#888888", "flex": 2 },
            { "type": "text", "text": `$ ${amount.toLocaleString()}`, "size": "xl", "color": "#E53935", "weight": "bold", "flex": 8 }
          ]
        },
        { "type": "separator" },
        {
          "type": "box",
          "layout": "vertical",
          "spacing": "xs",
          "contents": [
            { "type": "box", "layout": "horizontal", "contents": [{ "type": "text", "text": "日期時間", "color": "#888888", "size": "xs", "flex": 3 }, { "type": "text", "text": date, "color": "#333333", "size": "xs", "wrap": true, "flex": 7 }] },
            { "type": "box", "layout": "horizontal", "contents": [{ "type": "text", "text": "費用類別", "color": "#888888", "size": "xs", "flex": 3 }, { "type": "text", "text": category, "color": "#333333", "size": "xs", "wrap": true, "flex": 7 }] },
            { "type": "box", "layout": "horizontal", "contents": [{ "type": "text", "text": "品項描述", "color": "#888888", "size": "xs", "flex": 3 }, { "type": "text", "text": desc, "color": "#333333", "size": "xs", "wrap": true, "flex": 7 }] }
          ]
        },
        { "type": "separator" },
        {
          "type": "box",
          "layout": "vertical",
          "contents": [
            { "type": "text", "text": "📝 原始備註/逐字稿:", "size": "xs", "color": "#888888", "weight": "bold" },
            { "type": "text", "text": memo, "size": "xs", "color": "#555555", "wrap": true, "margin": "xs" }
          ]
        }
      ]
    },
    "footer": {
      "type": "box",
      "layout": "vertical",
      "spacing": "sm",
      "contents": [
        {
          "type": "button",
          "style": "primary",
          "color": "#4CAF50",
          "action": {
            "type": "postback",
            "label": "✅ 快速確認核准",
            "data": `action=approve&id=${expId}`
          }
        },
        {
          "type": "box",
          "layout": "horizontal",
          "spacing": "sm",
          "contents": [
            { "type": "button", "style": "secondary", "height": "sm", "action": { "type": "uri", "label": "📋 進入 Sheet", "uri": sheetUrl } },
            { "type": "button", "style": "secondary", "height": "sm", "action": { "type": "uri", "label": "📁 查看單據", "uri": driveUrl } }
          ]
        }
      ]
    }
  };
}

// ================= 基礎建設: 定時催報廣播模組 =================

// 每日早上 10:00 催報開班點鈔
function sendMorningReminder() {
  if (!ALUN_LINE_USER_ID) return;
  const msg = "🤖【真傳好食 ‧ 開店播報】\n阿倫老板早安！店面準備開張營業囉！請清點收銀機內底現有的【開班金額】，並直接回傳數字（例如：3000）給我，作伙建立起良好的現金流量紀錄吧！加油！";
  pushTextMessage(ALUN_LINE_USER_ID, msg);
}

// 每日晚上 19:30 催報打烊點鈔與日結單
function sendEveningReminder() {
  if (!ALUN_LINE_USER_ID) return;
  const msg = "🤖【真傳好食 ‧ 打烊播報】\n阿倫老闆今日辛苦囉！收攤打烊時，請幫忙清點收銀機內的金額，輸入「打烊 數字」（例如：打烊 16500）回傳；並且記得拍一張【POS日結單】傳上來封存喔！";
  pushTextMessage(ALUN_LINE_USER_ID, msg);
}

// ================= Google Workspace 檔案儲存與安全防禦 =================
function saveFileToDrive(blob, typeLabel, dateStr) {
  try {
    const root = DriveApp.getFolderById(FOLDER_ID);
    const yearStr = dateStr.split('/')[0] + "年";
    const monthStr = dateStr.split('/')[1] + "月";
    
    // 年、月兩層樹狀結構目錄判定
    let yearFolder = root.getFoldersByName(yearStr).hasNext() ? root.getFoldersByName(yearStr).next() : root.createFolder(yearStr);
    let monthFolder = yearFolder.getFoldersByName(monthStr).hasNext() ? yearFolder.getFoldersByName(monthStr).next() : yearFolder.createFolder(monthStr);
    
    // 自動相容多種常見音訊 MIME 格式
    const mimeType = blob.getContentType() || "";
    let ext = "jpg";
    if (mimeType.includes("audio")) {
      if (mimeType.includes("mp4") || mimeType.includes("m4a") || mimeType.includes("x-m4a")) ext = "m4a";
      else if (mimeType.includes("mpeg")) ext = "mp3";
      else if (mimeType.includes("aac")) ext = "aac";
      else ext = "wav";
    }
    
    const timeStamp = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd_HHmmss');
    const fileName = `${dateStr.replace(/\//g, '')}_${typeLabel}_${timeStamp}.${ext}`;
    
    const file = monthFolder.createFile(blob).setName(fileName);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    console.error('saveFileToDrive 異常:', e.toString());
    return "";
  }
}

// 避開標題列優化：從索引 i > 0 (第2列之後) 開始匹配
function findRowByDate(sheet, dateStr, dateColIdx) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;
  const values = sheet.getRange(1, dateColIdx, lastRow, 1).getValues().flat();
  
  // 從最後一列向前匹配，避開索引 0 (第一列標題)
  for (let i = values.length - 1; i > 0; i--) {
    if (values[i]) {
      try {
        const dateObj = values[i] instanceof Date ? values[i] : new Date(values[i]);
        const formatted = Utilities.formatDate(dateObj, 'Asia/Taipei', 'yyyy/MM/dd');
        if (formatted === dateStr) {
          return i + 1;
        }
      } catch(e) {
        // 忽匹配失敗的欄位
      }
    }
  }
  return 0;
}

// ================= 指數退避 API 重試機制 (Exponential Backoff) =================
function callGeminiApiWithRetry(blob, promptText, useJsonMode) {
  const maxRetries = 5;
  const delays = [1000, 2000, 4000, 8000, 16000]; // 1s, 2s, 4s, 8s, 16s 指數型間隔
  
  for (let attempt = 0; attempt < maxRetries; attempt++) {
    try {
      return callGeminiApi(blob, promptText, useJsonMode);
    } catch (error) {
      if (attempt === maxRetries - 1) {
        console.error(`[Gemini API 重試次數耗盡] Attempt: ${attempt + 1}, Error: ${error.toString()}`);
        throw new Error("AI 辨識引擎目前忙碌中（已嘗試重試5次），請稍等 1 分鐘後再次上傳或輸入。");
      }
      Utilities.sleep(delays[attempt]); // 執行退避延遲
    }
  }
}

// 安全的 JSON 解析器：精確過濾 AI 在 Markdown 環境下吐出的代碼圍欄
function parseJsonFromModel(text) {
  const cleaned = String(text)
    .replace(/^```json\s*/i, '')
    .replace(/^```\s*/i, '')
    .replace(/\s*```$/, '')
    .trim();
  return JSON.parse(cleaned);
}

function callGeminiApi(blob, promptText, useJsonMode) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${GEMINI_API_KEY}`;
  let parts = [{ text: promptText }];
  
  if (blob) {
    const mimeType = blob.getContentType() || "image/jpeg";
    parts.push({ 
      inlineData: { 
        mimeType: mimeType, 
        data: Utilities.base64Encode(blob.getBytes()) 
      } 
    });
  }
  
  const payload = { 
    contents: [{ parts: parts }]
  };
  if (useJsonMode) {
    payload.generationConfig = { 
      responseMimeType: "application/json" 
    };
  }

  // 設定防鎖死限流
  const waitTime = waitForMyTurn();
  if (waitTime > 0) Utilities.sleep(waitTime);

  const response = UrlFetchApp.fetch(url, {
    method: 'post', 
    contentType: 'application/json',
    payload: JSON.stringify(payload), 
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() !== 200) {
    throw new Error(`API 伺服器回傳狀態碼 ${response.getResponseCode()}: ${response.getContentText()}`);
  }
  
  const textResult = JSON.parse(response.getContentText())?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!textResult) throw new Error("API 未正常回傳解析文字內容。");
  
  if (!useJsonMode) return textResult;
  
  try {
    return parseJsonFromModel(textResult);
  } catch (e) {
    throw new Error(`JSON 解析失敗: ${e.message}; 原始文本: ${textResult}`);
  }
}

function waitForMyTurn() {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(25000)) {
    try {
      const cache = CacheService.getScriptCache();
      const now = Date.now();
      let next = Number(cache.get('GLOBAL_API_NEXT_TICKET') || 0);
      let sleepMs = Math.max(now, next) - now;
      cache.put('GLOBAL_API_NEXT_TICKET', String(Math.max(now, next) + MIN_API_INTERVAL), 60);
      return sleepMs;
    } finally { lock.releaseLock(); }
  } else { return 2000; }
}

// ================= LINE 底層通訊元件 =================
function replyLine(token, text) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    headers: { 'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN, 'Content-Type': 'application/json' },
    method: 'post', payload: JSON.stringify({ replyToken: token, messages: [{ type: 'text', text }] }), muteHttpExceptions: true
  });
}

function pushTextMessage(toUserId, text) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    headers: { 'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN, 'Content-Type': 'application/json' },
    method: 'post', payload: JSON.stringify({ to: toUserId, messages: [{ type: 'text', text }] }), muteHttpExceptions: true
  });
}

function pushFlexMessage(toUserId, altText, flexContents) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    headers: { 'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN, 'Content-Type': 'application/json' },
    method: 'post', payload: JSON.stringify({ to: toUserId, messages: [{ type: "flex", altText: altText, contents: flexContents }] }), muteHttpExceptions: true
  });
}

function getLineContent(id) {
  const res = UrlFetchApp.fetch(`https://api-data.line.me/v2/bot/message/${id}/content`, { headers: { 'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN }, muteHttpExceptions: true });
  return res.getResponseCode() === 200 ? res.getBlob() : null;
}

function showLoadingAnimation(chatId, sec) {
  try { UrlFetchApp.fetch('https://api.line.me/v2/bot/chat/loading/start', { headers: { 'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN, 'Content-Type': 'application/json' }, method: 'post', payload: JSON.stringify({ "chatId": chatId, "loadingSeconds": sec }), muteHttpExceptions: true }); } catch (e) {}
}
