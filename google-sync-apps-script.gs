/**
 * 韓國行前清單 雲端同步用 Apps Script
 * 設定步驟（只要做一次，約 3 分鐘）：
 * 1. 開啟 https://sheets.new 建立一個新的 Google 試算表（名稱隨意，例如「韓國行清單同步」）
 * 2. 上方選單「擴充功能」→「Apps Script」
 * 3. 把這個檔案的全部內容貼上（取代原本的 myFunction），按 Ctrl+S 儲存
 * 4. 右上角「部署」→「新增部署作業」→ 齒輪選「網頁應用程式」
 *    - 執行身分：我
 *    - 誰可以存取：任何人
 *    → 按「部署」→ 授權自己的 Google 帳號 → 複製「網頁應用程式 URL」
 * 5. 回到行程網頁「行前清單」分頁 → 展開「☁️ 雲端同步設定」→ 貼上 URL → 按「啟用同步」
 *    （手機和電腦都貼同一個 URL，勾選就會互相同步）
 */

const SHEET_NAME = "data";

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

// 讀取：回傳上次儲存的 JSON（A1 儲存格）
function doGet(e) {
  const val = getSheet_().getRange("A1").getValue();
  return ContentService
    .createTextOutput(val ? String(val) : "{}")
    .setMimeType(ContentService.MimeType.JSON);
}

// 寫入：把整包 JSON 存到 A1，並在 B1 記錄更新時間
function doPost(e) {
  const sh = getSheet_();
  sh.getRange("A1").setValue(e.postData.contents);
  sh.getRange("B1").setValue(new Date());
  return ContentService
    .createTextOutput('{"ok":true}')
    .setMimeType(ContentService.MimeType.JSON);
}
