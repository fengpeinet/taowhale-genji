/**
 * 廈門AI課 - 代寄信腳本
 *
 * 注意：這不是「轉寄」，收件人收到的是這個新帳號原生寄出的正常信件，
 * 不會有 Fwd: 或轉寄鏈。主帳號只是傳資料（收件人/主旨/內容）過來，
 * 實際寄出的動作由這個帳號執行，目的是讓寄信額度算在這個帳號上，
 * 不會跟主帳號 fengpeinet@gmail.com（掛舊活動）搶額度。
 *
 * 設定步驟（一次性，做完不用再回來管）：
 * 1. 用新帳號登入 https://script.google.com
 * 2. 建立新專案（不用綁定任何試算表）
 * 3. 把這支檔案的內容整個貼進去，儲存
 * 4. 右上角「部署」→「新增部署作業」→ 類型選「網頁應用程式」
 *    - 執行身分：我（新帳號自己）
 *    - 具有存取權的使用者：任何人
 * 5. 部署後會拿到一個網址，把它貼回主帳號那支 google_apps_script.js 裡的 MAIL_RELAY_URL
 */

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    MailApp.sendEmail({
      to: data.to,
      subject: data.subject,
      htmlBody: data.htmlBody,
      name: data.name || "風霈學院"
    });
    return ContentService.createTextOutput(JSON.stringify({ status: "sent" })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}
