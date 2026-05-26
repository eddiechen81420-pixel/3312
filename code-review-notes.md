# Gemini 最新版程式碼審查（重點修正清單）

## 先說結論（最重要）
你這份「最新」程式碼目前**不能直接部署**，主要不是邏輯，而是「檔案內容被重複貼上兩次以上」。

---

## P0（必修，先修這些）

1. **整份程式被重複拼接，且尾端文字混入自然語言**
   - 你的檔案在 `showLoadingAnimation` 後面又接了整份 `// ================= 全域變數與系統配置 =================`。
   - 最尾端還有 `，gemini提供最新的程式碼...` 這種自然語言，會直接造成語法錯誤。
   - ✅ 作法：保留第一份完整程式，刪掉後面重複段與所有自然語言。

2. **`replyLine/push*` URL 曾經有 markdown 連結格式殘留風險**
   - 若出現 `"[https://...](https://...)"` 就會壞。
   - ✅ 作法：只保留純字串 URL（你目前修正版已是純 URL，這是正確方向）。

3. **部署前先做最小語法驗證**
   - Apps Script 編輯器先執行一次任意函式（如 `parseJsonFromModel`）確認能編譯。

---

## P1（高價值優化）

1. **`executeExpensePipeline` 回覆文案中 `data.description` 可能是 `undefined`**
   - 你寫 row 時有 fallback：`data.description || "未命名項目"`，但回覆文字使用的是 `${data.description}`。
   - ✅ 建議統一使用 `const desc = data.description || "未命名項目"`。

2. **`data.amount.toLocaleString()` 依賴 amount 一定是 number**
   - AI 可能回字串（例如 "150"）。
   - ✅ 建議：`const amount = Number(data.amount); if (!Number.isFinite(amount)) ...`。

3. **`processTextEvent` 的 `chatId` 參數未使用**
   - 目前僅簽名有 `chatId`。
   - ✅ 建議：移除未用參數，避免混淆。

4. **`findRowByDate` 可直接從第 2 列取 range**
   - 目前用第 1 列再跳過 index 0，能跑但可更直觀。

---

## P2（維運建議）

1. **記錄結構化 log 上下文**
   - 錯誤 log 建議統一帶 `userId`, `messageId`, `expId`, `sheetName`。

2. **`muteHttpExceptions: true` 後可補 API 回應記錄**
   - 對 LINE push/reply 失敗追查會更快。

---

## 可直接套用的小修片段

```javascript
// executeExpensePipeline 內
const desc = data.description || "未命名項目";
const amount = Number(data.amount);
if (!Number.isFinite(amount)) {
  replyLine(replyToken, "⚠️ 金額格式異常，請重試。");
  return;
}

const rowData = [expId, dateStr, category, desc, amount, fileUrl, data.memo || "", false];
sheet.appendRow(rowData);
replyLine(replyToken, `📝 已幫您錄入一筆【${category}】支出預覽：${desc} $${amount.toLocaleString()} 元。已發送審核卡片給 Eddie 核對。`);
```

---

## 總結
你這版其實架構成熟、流程完整（文字/圖片/語音/審核/重試/限流都有），
但目前最大阻塞是**內容重複拼接導致不可編譯**。先清掉重複段與尾端自然語言，再做 amount/description 防呆，就能穩定很多。
