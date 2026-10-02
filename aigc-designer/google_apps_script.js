/**
 * FENGPEINET 說明會推廣報名系統 - 圖文配置 + 完整日誌版
 * 廈門AI課（鯨海拾貝代理課程）專用 - 從「蝦皮AI攻略班」專案 fork 過來的骨架
 *
 * ⚠️ 部署注意：這支程式碼不能沿用舊專案的 Apps Script 部署，
 * 必須另外建立一個新的 Google 試算表 → 綁定這支腳本 → 重新部署成 Web App，
 * 拿到新的網址後貼回 index.html 的 SCRIPT_URL。沿用舊部署會把這次的名單寫進舊活動的試算表。
 */

// ==========================================
//   一、 營運設定區 (連結與圖片)
// ==========================================

// key 格式為 "月-日"（例：10-15 代表 10/15），要跟 index.html 表單 radio 的 value 對應邏輯保持一致。
// 四場都先放這裡備用，前台 index.html 一次只顯示其中兩場，場次過了再手動切換前台顯示哪兩場
const ZOOM_LINKS = {
  "10-7": { time: "20:00", url: "https://us06web.zoom.us/j/3909390259?omn=88509412246" },
  "10-14": { time: "20:00", url: "https://us06web.zoom.us/j/3909390259?omn=89432569323" },
  "10-21": { time: "20:00", url: "https://us06web.zoom.us/j/3909390259?omn=81340898316" },
  "10-28": { time: "20:00", url: "https://us06web.zoom.us/j/3909390259?omn=86535584001" },
};

const WEBINAR_URL = "https://fengpeinet.github.io/taowhale-genji/aigc-designer/";
const REBOOK_URL = "https://fengpeinet.github.io/taowhale-genji/aigc-designer/?utm_source=facebook&utm_medium=share&ref=更換場次專用";

const SHEET_NAME = "工作表1";

// AI電話系統只能同步一份 Google Sheet，所以「有填電話」的名單不寫自己試算表的分頁，
// 改成直接寫進 16,800 帶貨起步班那份試算表的「AI電話」分頁（兩門課共用同一份 AI電話名單）
const AI_PHONE_EXTERNAL_SHEET_ID = "15RD8rlyelrBrLIqmCgsjO9g7nD5F4Q0LYMfi9EF_Go0";
const AI_PHONE_SHEET_NAME = "AI電話";
const COURSE_TAG = "AIGC說明會"; // 寫進 AI電話分頁 M 欄，用來區分是哪門課的名單

// TODO：另一個帳號的「代寄信」腳本部署網址（見 google_apps_script_代寄信.js）。
// 這裡只是把信件內容傳過去，實際寄出動作由那個帳號執行（不是轉寄，收件人看到的是正常信件），
// 目的是讓寄信額度算在那個帳號上，不會跟舊活動搶主帳號 fengpeinet@gmail.com 的每日額度
const MAIL_RELAY_URL = "https://script.google.com/macros/s/AKfycbyL4I-qRyr9Xn-RcxOq3xiEykB2DzocdOJjux5BUUAhdvgxAGuJwcl1vWelLGG3ABuo/exec";


// ==========================================
//   一之二、 代寄信（實際發信改由另一個帳號執行）
// ==========================================

function sendViaRelay(to, subject, htmlBody, name) {
  const payload = { to: to, subject: subject, htmlBody: htmlBody, name: name || "風霈學院" };
  const res = UrlFetchApp.fetch(MAIL_RELAY_URL, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  console.log("代寄信結果: " + res.getContentText());
}


// ==========================================
//   二、 工具函式
// ==========================================

function isTruthy(val) {
  return val === true || val === "TRUE";
}

function normalizeDate(rawDate, timeZone) {
  if (rawDate instanceof Date) {
    return Utilities.formatDate(rawDate, timeZone, "yyyy-M-d");
  }
  const s = rawDate.toString().replace(/-/g, "/").split(" ")[0];
  const d = new Date(s);
  return !isNaN(d.getTime()) ? Utilities.formatDate(d, timeZone, "yyyy-M-d") : "";
}

function getDisplayDate(rawDate, timeZone) {
  if (rawDate instanceof Date) {
    return Utilities.formatDate(rawDate, timeZone, "yyyy/MM/dd HH:mm");
  }
  return rawDate.toString().replace(/台北/g, "").trim();
}

function normalizePhone(phone) {
  if (!phone) return '';
  let p = phone.toString().replace(/[\s\-\(\)]/g, ''); // 移除空格、dash、括號
  if (p.startsWith('+886')) p = '0' + p.slice(4);
  else if (p.startsWith('886')) p = '0' + p.slice(3);
  return p;
}


// ==========================================
//   三、 排程發信功能 (含詳細日誌)
// ==========================================

// 欄位對應（0-based index）：
// A=0 時間, B=1 姓名, C=2 電話, D=3 Email, E=4 LINE ID
// F=5 推薦人, G=6 場次, H=7 前天提醒, I=8 當天提醒
// J=9 最後召回日, K=10 直播連結, L=11 電話行銷人員

function sendScheduledEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  console.log("--- 執行排程掃描開始 ---");

  if (!sheet) {
    console.error("錯誤：找不到工作表 [" + SHEET_NAME + "]");
    return;
  }

  const data = sheet.getDataRange().getValues();
  console.log("偵測到總列數 (含標題): " + data.length);

  if (data.length <= 1) {
    console.log("目前沒有報名資料。");
    return;
  }

  const today = new Date();
  const timeZone = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(today, timeZone, "yyyy-M-d");

  const tomorrow = new Date();
  tomorrow.setDate(today.getDate() + 1);
  const tomorrowStr = Utilities.formatDate(tomorrow, timeZone, "yyyy-M-d");

  const [ty, tm, td] = todayStr.split("-").map(Number);
  const todayDateObj = new Date(ty, tm - 1, td);

  console.log("今日基準日期: " + todayStr);

  // PHASE 1: 找下一場（今天以後的最近一場，今天本身不算）
  let nextSessionKey = null;
  let nextSessionEntry = null;
  for (const [k, v] of Object.entries(ZOOM_LINKS)) {
    const [sm, sd] = k.split("-").map(Number);
    const sessionDate = new Date(2026, sm - 1, sd);
    if (sessionDate >= todayDateObj) {
      nextSessionKey = k;
      nextSessionEntry = v;
      break;
    }
  }
  if (nextSessionKey) {
    console.log("下一場說明會: " + nextSessionKey + " " + nextSessionEntry.time);
  } else {
    console.log("找不到下一場說明會，今日跳過召回信。");
  }

  // PHASE 2: 預估今日提醒信數量，計算召回配額
  // 公式：可發召回 = 100 - 提醒信數量 - 30（預留當日新報名）
  let estimatedReminders = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[3] || !row[6]) continue;
    const dateStr = normalizeDate(row[6], timeZone);
    if (dateStr === todayStr && !isTruthy(row[8])) estimatedReminders++;
    if (dateStr === tomorrowStr && !isTruthy(row[7])) estimatedReminders++;
  }
  const recallQuota = Math.max(0, 100 - estimatedReminders - 30);
  console.log("預估提醒信: " + estimatedReminders + " 封，預留報名通知: 30，召回配額: " + recallQuota + " 封");

  // PHASE 3: 發送提醒信 + 召回信
  let remindersSent = 0;
  let recallSent = 0;
  const todayWriteStr = Utilities.formatDate(today, timeZone, "yyyy/MM/dd");

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const userName = row[1];
    const userEmail = row[3];
    const rawDate = row[6];
    const isDayBeforeReminded = row[7]; // H欄
    const isDayOfReminded = row[8];     // I欄
    const lastRecallDate = row[9];      // J欄：最後召回日

    if (!userName || !userEmail || !rawDate) {
      console.log("第 " + (i + 1) + " 列資料不完整，跳過。");
      continue;
    }

    const targetDateStr = normalizeDate(rawDate, timeZone);
    if (!targetDateStr) continue;

    const displayDate = getDisplayDate(rawDate, timeZone);
    const [, rm, rd] = targetDateStr.split("-").map(Number);
    // 優先用這一列自己存的直播連結（K欄，報名當下就寫入），沒有的話才回頭查 ZOOM_LINKS——
    // 這樣就算之後 ZOOM_LINKS 表被清掉/改掉，已經報名的人拿到的連結還是當初報名時對應的那一場
    const zoomLink = row[10] || ZOOM_LINKS[rm + "-" + rd]?.url;

    console.log("正在檢查: " + userName + " | 場次: " + targetDateStr);

    try {
      // --- 今日提醒 ---
      if (targetDateStr === todayStr) {
        if (!isTruthy(isDayOfReminded)) {
          sendReminderEmail(userEmail, userName, displayDate, zoomLink, "今天");
          sheet.getRange(i + 1, 9).setValue(true); // I欄
          remindersSent++;
          console.log("   >>> [成功] 已寄出今日提醒信");
        } else {
          console.log("   (今日提醒已寄過，跳過)");
        }
      }

      // --- 明日提醒 ---
      if (targetDateStr === tomorrowStr) {
        if (!isTruthy(isDayBeforeReminded)) {
          sendReminderEmail(userEmail, userName, displayDate, zoomLink, "明天");
          sheet.getRange(i + 1, 8).setValue(true); // H欄
          remindersSent++;
          console.log("   >>> [成功] 已寄出明日提醒信");
        } else {
          console.log("   (明日提醒已寄過，跳過)");
        }
      }

      // --- 召回信：此人報名的是過去的場次，且今天有下一場可寄 ---
      if (nextSessionEntry && recallSent < recallQuota) {
        const [ry, rm, rd] = targetDateStr.split("-").map(Number);
        if (isNaN(ry)) continue;
        const targetDateObj = new Date(ry, rm - 1, rd);

        if (targetDateObj < todayDateObj) {
          // 檢查 J 欄：30 天內是否已寄過召回信
          let eligible = true;
          if (lastRecallDate) {
            let lastRecall;
            if (lastRecallDate instanceof Date) {
              lastRecall = lastRecallDate;
            } else {
              lastRecall = new Date(lastRecallDate.toString().replace(/-/g, "/"));
            }
            if (!isNaN(lastRecall.getTime())) {
              const daysSince = Math.floor((todayDateObj - lastRecall) / (1000 * 60 * 60 * 24));
              if (daysSince < 30) {
                eligible = false;
                console.log("   (30 天內已寄過召回信，跳過)");
              }
            }
          }

          if (eligible) {
            const [nm, nd] = nextSessionKey.split("-").map(Number);
            const nextDisplayDate = "2026/" + String(nm).padStart(2, "0") + "/" + String(nd).padStart(2, "0") + " " + nextSessionEntry.time;
            sendRecallEmail(userEmail, userName, nextDisplayDate, nextSessionEntry.url);
            sheet.getRange(i + 1, 10).setValue(todayWriteStr); // J欄（1-based 第 10 欄）
            recallSent++;
            console.log("   >>> [成功] 已寄出召回信");
          }
        }
      }

    } catch (err) {
      console.error("   [錯誤] 第 " + (i + 1) + " 列寄信失敗，跳過：" + err.message);
    }
  }

  SpreadsheetApp.flush();
  console.log("總計：提醒信 " + remindersSent + " 封，召回信 " + recallSent + " 封");
  console.log("--- 掃描結束 ---");
}


// ==========================================
//   四、 精美信件 HTML 排版
// ==========================================

function sendReminderEmail(to, name, displayDate, link, dayWord) {
  const subject = `📣 ${dayWord}見！《AI設計師研修班》說明會 活動提醒`;
  const footerUrl = "https://lh3.googleusercontent.com/d/1MAo2woNkcN7-LYjjoL6ZMdlFIToLyz-s";
  const headUrl = "https://lh3.googleusercontent.com/d/1XG6Co4hBQ0r0Yk5c5Zc-I3zhHftUYt3B";
  const brandLogoUrl = "https://lh3.googleusercontent.com/d/12c_FvDGcLxbMlDR4caLLjeElunPRiOrB";
  const htmlBody = `
    <div style="font-family: 'Microsoft JhengHei', Arial, sans-serif; color: #333; line-height: 1.8; max-width: 600px; margin: 0 auto;">

      <img src="${headUrl}" alt="《AI設計師研修班》說明會" style="width: 100%; max-width: 600px; display: block; border-radius: 8px; margin-bottom: 24px;">

      <p style="margin: 0 0 8px 0;"><strong>${name}</strong> 您好：</p>
      <p style="margin: 0 0 20px 0;">《AI設計師研修班》說明會<strong>${dayWord}</strong>就要開始囉！</p>

      <!-- 活動資訊 -->
      <p style="font-weight: bold; margin: 0 0 8px 0;">✅ 活動資訊</p>
      <p style="margin: 4px 0 4px 16px;">活動名稱：<strong>《AI設計師研修班》說明會</strong></p>
      <p style="margin: 4px 0 4px 16px;">活動時間：<strong>${displayDate}</strong></p>
      <p style="margin: 4px 0 20px 16px;">活動形式：線上 Zoom</p>

      <!-- 你將會了解 -->
      <p style="font-weight: bold; margin: 0 0 8px 0;">✅ 這場分享會，你將會了解</p>
      <ul style="margin: 0 0 20px 0; padding-left: 20px;">
        <!-- TODO：這四點是舊「蝦皮AI攻略班」的賣點，需換成《AI設計師研修班》的課程亮點
             （參考專案總覽.md 第四節：AI生文/生圖/生視頻三大模塊，82節錄播＋開營直播＋4節直播課） -->
        <li style="margin-bottom: 6px;">TODO：課程亮點 1</li>
        <li style="margin-bottom: 6px;">TODO：課程亮點 2</li>
        <li style="margin-bottom: 6px;">TODO：課程亮點 3</li>
        <li style="margin-bottom: 6px;">TODO：課程亮點 4</li>
      </ul>

      <!-- 橘色框：三步驟 + Zoom 按鈕 + 重新報名 -->
      <div style="border: 2px solid #ff6600; border-radius: 8px; padding: 16px 20px; margin: 0 0 20px 0; background-color: #fff8f0;">
        <p style="margin: 0 0 10px 0; color: #ff6600; font-weight: bold;">請依照下面簡單三步驟就可參加：</p>
        <ol style="margin: 0 0 20px 0; padding-left: 20px;">
          <li style="margin-bottom: 8px;"><strong>下載會議軟體：</strong>我們使用「Zoom」作為線上遠端教學軟體。</li>
          <li style="margin-bottom: 8px;"><strong>點擊下方按鈕：</strong>直接進入線上研討會教室。</li>
          <li style="margin-bottom: 8px;"><strong>準時出席：</strong>活動當天請提前 10 分鐘進入測試設備。</li>
        </ol>
        <p style="text-align: center; margin: 0 0 12px 0;">
          <a href="${link}" style="background-color: #ff6600; color: #ffffff; padding: 14px 32px; text-decoration: none; border-radius: 6px; display: inline-block; font-weight: bold; font-size: 16px;">立即進入 Zoom 研討會</a>
        </p>
        <p style="text-align: center; margin: 0;">
          <a href="${WEBINAR_URL}" style="color: #1a73e8; font-size: 13px; text-decoration: underline;">這個時間無法上線？看看其他場次！</a>
        </p>
      </div>

      <!-- 參加須知 -->
      <p style="font-weight: bold; margin: 0 0 8px 0;">✅ 參加須知</p>
      <ul style="margin: 0 0 20px 0; padding-left: 20px;">
        <li style="margin-bottom: 6px;">請於活動開始前 <strong>10 分鐘</strong>進入會議室，以免錯過開場內容</li>
        <li style="margin-bottom: 6px;">線上分享，請提前確認網路與設備，<strong>建議使用電腦觀看效果較佳</strong></li>
        <li style="margin-bottom: 6px;">建議準備筆記工具，活動中會分享重點內容與資源連結</li>
      </ul>

      <p style="margin: 0 0 20px 0;">期待這場分享會能帶給您新的啟發與方向，我們活動當天見！</p>

      <div style="border-top: 1px solid #ddd; margin: 20px 0;"></div>

      <table style="width: 100%; max-width: 600px; margin: 0 0 24px 0; border-collapse: collapse;">
        <tr>
          <td style="width: 180px; vertical-align: middle; padding-right: 16px;">
            <img src="${brandLogoUrl}" alt="風霈國際傳媒" style="display: block; width: 160px; height: auto;">
          </td>
          <td style="vertical-align: middle;">
            <p style="margin: 0 0 6px 0; font-size: 14px;">如有問題歡迎詢問官方 LINE：<a href="https://lin.ee/7zmJUX9">https://lin.ee/7zmJUX9</a></p>
            <p style="margin: 0 0 6px 0; font-size: 14px;"><a href="${WEBINAR_URL}" style="color: #1a73e8; text-decoration: underline;">這個時間無法上線？看看其他場次！</a></p>
            <p style="margin: 0; color: #999; font-size: 13px;">※ 此信件為系統自動發送，請勿直接回覆此信件。</p>
          </td>
        </tr>
      </table>

      <img src="${footerUrl}" alt="" style="width: 100%; max-width: 600px; display: block;">

    </div>
  `;
  sendViaRelay(to, subject, htmlBody, "風霈學院"); // TODO：課程正式名稱定案後可改成「風霈學院｜AI設計師研修班」之類
}

function sendRecallEmail(to, name, nextDisplayDate, zoomLink) {
  const subject = `📢 ${name}，還記得我們嗎？`;
  const footerUrl = "https://lh3.googleusercontent.com/d/1MAo2woNkcN7-LYjjoL6ZMdlFIToLyz-s";
  const headUrl = "https://lh3.googleusercontent.com/d/1XG6Co4hBQ0r0Yk5c5Zc-I3zhHftUYt3B";
  const brandLogoUrl = "https://lh3.googleusercontent.com/d/12c_FvDGcLxbMlDR4caLLjeElunPRiOrB";
  const htmlBody = `
    <div style="font-family: 'Microsoft JhengHei', Arial, sans-serif; color: #333; line-height: 1.8; max-width: 600px; margin: 0 auto;">

      <img src="${headUrl}" alt="《AI設計師研修班》說明會" style="width: 100%; max-width: 600px; display: block; border-radius: 8px; margin-bottom: 24px;">

      <p style="margin: 0 0 8px 0;"><strong>${name}</strong> 您好：</p>
      <p style="margin: 0 0 20px 0;">嗨！你之前有報名過我們的「《AI設計師研修班》說明會」，我們還在！下一場說明會即將開始，希望這次能讓你把握機會 😊</p>

      <!-- 活動資訊 -->
      <p style="font-weight: bold; margin: 0 0 8px 0;">✅ 活動資訊</p>
      <p style="margin: 4px 0 4px 16px;">活動名稱：<strong>《AI設計師研修班》說明會</strong></p>
      <p style="margin: 4px 0 4px 16px;">活動時間：<strong>${nextDisplayDate}</strong></p>
      <p style="margin: 4px 0 20px 16px;">活動形式：線上 Zoom</p>

      <!-- 你將會了解 -->
      <p style="font-weight: bold; margin: 0 0 8px 0;">✅ 這場分享會，你將會了解</p>
      <ul style="margin: 0 0 20px 0; padding-left: 20px;">
        <!-- TODO：這四點是舊「蝦皮AI攻略班」的賣點，需換成《AI設計師研修班》的課程亮點
             （參考專案總覽.md 第四節：AI生文/生圖/生視頻三大模塊，82節錄播＋開營直播＋4節直播課） -->
        <li style="margin-bottom: 6px;">TODO：課程亮點 1</li>
        <li style="margin-bottom: 6px;">TODO：課程亮點 2</li>
        <li style="margin-bottom: 6px;">TODO：課程亮點 3</li>
        <li style="margin-bottom: 6px;">TODO：課程亮點 4</li>
      </ul>

      <!-- 橘色框：三步驟 + Zoom 按鈕 -->
      <div style="border: 2px solid #ff6600; border-radius: 8px; padding: 16px 20px; margin: 0 0 20px 0; background-color: #fff8f0;">
        <p style="margin: 0 0 10px 0; color: #ff6600; font-weight: bold;">請依照下面簡單三步驟就可參加：</p>
        <ol style="margin: 0 0 20px 0; padding-left: 20px;">
          <li style="margin-bottom: 8px;"><strong>下載會議軟體：</strong>我們使用「Zoom」作為線上遠端教學軟體。</li>
          <li style="margin-bottom: 8px;"><strong>點擊下方按鈕：</strong>直接進入線上研討會教室。</li>
          <li style="margin-bottom: 8px;"><strong>準時出席：</strong>活動當天請提前 10 分鐘進入測試設備。</li>
        </ol>
        <p style="text-align: center; margin: 0 0 12px 0;">
          <a href="${zoomLink}" style="background-color: #ff6600; color: #ffffff; padding: 14px 32px; text-decoration: none; border-radius: 6px; display: inline-block; font-weight: bold; font-size: 16px;">立即進入 Zoom 研討會</a>
        </p>
        <p style="text-align: center; margin: 0;">
          <a href="${REBOOK_URL}" style="color: #1a73e8; font-size: 13px; text-decoration: underline;">這個時間無法上線？看看其他場次！</a>
        </p>
      </div>

      <!-- 參加須知 -->
      <p style="font-weight: bold; margin: 0 0 8px 0;">✅ 參加須知</p>
      <ul style="margin: 0 0 20px 0; padding-left: 20px;">
        <li style="margin-bottom: 6px;">請於活動開始前 <strong>10 分鐘</strong>進入會議室，以免錯過開場內容</li>
        <li style="margin-bottom: 6px;">線上分享，請提前確認網路與設備，<strong>建議使用電腦觀看效果較佳</strong></li>
        <li style="margin-bottom: 6px;">建議準備筆記工具，活動中會分享重點內容與資源連結</li>
      </ul>

      <p style="margin: 0 0 20px 0;">期待這場分享會能帶給您新的啟發與方向，我們說明會見！</p>

      <div style="border-top: 1px solid #ddd; margin: 20px 0;"></div>

      <table style="width: 100%; max-width: 600px; margin: 0 0 24px 0; border-collapse: collapse;">
        <tr>
          <td style="width: 180px; vertical-align: middle; padding-right: 16px;">
            <img src="${brandLogoUrl}" alt="風霈國際傳媒" style="display: block; width: 160px; height: auto;">
          </td>
          <td style="vertical-align: middle;">
            <p style="margin: 0 0 6px 0; font-size: 14px;">如有問題歡迎詢問官方 LINE：<a href="https://lin.ee/7zmJUX9">https://lin.ee/7zmJUX9</a></p>
            <p style="margin: 0 0 6px 0; font-size: 14px;"><a href="${REBOOK_URL}" style="color: #1a73e8; text-decoration: underline;">這個時間無法上線？看看其他場次！</a></p>
            <p style="margin: 0; color: #999; font-size: 13px;">※ 此信件為系統自動發送，請勿直接回覆此信件。</p>
          </td>
        </tr>
      </table>

      <img src="${footerUrl}" alt="" style="width: 100%; max-width: 600px; display: block;">

    </div>
  `;
  sendViaRelay(to, subject, htmlBody, "風霈學院"); // TODO：課程正式名稱定案後可改成「風霈學院｜AI設計師研修班」之類
}

function sendRegistrationSuccessEmail(to, name, date, zoomUrl) {
  const subject = "✅ 您已成功報名「《AI設計師研修班》說明會」";
  const footerUrl = "https://lh3.googleusercontent.com/d/1MAo2woNkcN7-LYjjoL6ZMdlFIToLyz-s";
  const headUrl = "https://lh3.googleusercontent.com/d/1XG6Co4hBQ0r0Yk5c5Zc-I3zhHftUYt3B";
  const brandLogoUrl = "https://lh3.googleusercontent.com/d/12c_FvDGcLxbMlDR4caLLjeElunPRiOrB";
  const htmlBody = `
    <div style="font-family: 'Microsoft JhengHei', Arial, sans-serif; color: #333; line-height: 1.8; max-width: 600px; margin: 0 auto;">

      <img src="${headUrl}" alt="《AI設計師研修班》說明會" style="width: 100%; max-width: 600px; display: block; border-radius: 8px; margin-bottom: 24px;">

      <p style="margin: 0 0 8px 0;"><strong>${name}</strong> 您好：</p>
      <p style="margin: 0 0 20px 0;">感謝您報名 <strong>《AI設計師研修班》說明會</strong>，您的報名已成功完成！</p>

      <div style="background-color: #f9f9f9; border-left: 4px solid #c8a84b; padding: 16px 20px; border-radius: 4px; margin: 0 0 16px 0;">
        <p style="margin: 0 0 8px 0; color: #a07830; font-weight: bold;">📋 您的報名資訊</p>
        <p style="margin: 4px 0;">報名姓名：<strong>${name}</strong></p>
        <p style="margin: 4px 0;">活動名稱：<strong>《AI設計師研修班》說明會</strong></p>
        <p style="margin: 4px 0;">場次時間：<strong>${date}</strong></p>
      </div>

      <div style="border: 2px solid #ff6600; border-radius: 8px; padding: 16px 20px; margin: 0 0 24px 0; background-color: #fff8f0;">
        <p style="margin: 0 0 8px 0; color: #ff6600; font-weight: bold;">📬 線上會議室連結：</p>
        <p style="margin: 0 0 12px 0;">線上會議室採用 ZOOM，建議提前下載手機 app</p>
        <p style="text-align: center; margin: 0;">
          <a href="${zoomUrl}" style="background-color: #ff6600; color: #ffffff; padding: 14px 32px; text-decoration: none; border-radius: 6px; display: inline-block; font-weight: bold; font-size: 16px;">立即進入 Zoom 研討會</a>
        </p>
      </div>

      <p style="text-align: center; margin: 0 0 12px 0; font-size: 14px;">怕漏接會議室連結嗎？<span style="font-weight: bold; text-decoration: underline; background-color: #ffe0b2; padding: 0 3px;">加入風霈官方 LINE＆傳貼圖給我們！</span>雙重保險不漏接 ✅</p>
      <p style="text-align: center; margin: 0 0 24px 0;">
        <a href="https://lin.ee/7zmJUX9" style="background-color: #06C755; color: #ffffff; padding: 10px 24px; text-decoration: none; border-radius: 6px; display: inline-block; font-weight: bold; font-size: 15px;">加入官方 LINE</a>
      </p>

      <div style="border-top: 1px solid #ddd; margin: 20px 0;"></div>

      <table style="width: 100%; max-width: 600px; margin: 0 0 24px 0; border-collapse: collapse;">
        <tr>
          <td style="width: 180px; vertical-align: middle; padding-right: 16px;">
            <img src="${brandLogoUrl}" alt="風霈國際傳媒" style="display: block; width: 160px; height: auto;">
          </td>
          <td style="vertical-align: middle;">
            <p style="margin: 0 0 6px 0; font-size: 14px;">如有問題歡迎詢問官方 LINE：<a href="https://lin.ee/7zmJUX9">https://lin.ee/7zmJUX9</a></p>
            <p style="margin: 0; color: #999; font-size: 13px;">※ 此信件為系統自動發送，請勿直接回覆此信件。</p>
          </td>
        </tr>
      </table>

      <img src="${footerUrl}" alt="" style="width: 100%; max-width: 600px; display: block;">

    </div>
  `;
  sendViaRelay(to, subject, htmlBody, "風霈學院"); // TODO：課程正式名稱定案後可改成「風霈學院｜AI設計師研修班」之類
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    const data = JSON.parse(e.postData.contents);

    const parts = (data.sessionDate || "").split(" ")[0].split("/");
    const sessionKey = parts.length >= 3
      ? (parseInt(parts[1]) + "-" + parseInt(parts[2]))
      : "";
    const zoomUrl = ZOOM_LINKS[sessionKey]?.url || "";
    const phone = normalizePhone(data.phone);

    // 工作表1：全部都寫
    sheet.appendRow([
      new Date(), data.name, '', data.email,
      data.lineId || '', data.referrer || '', data.sessionDate,
      false, false, '', zoomUrl
    ]);
    const lastRow = sheet.getLastRow();
    // C欄=電話，寫入前先設成純文字格式，避免 Sheets 把開頭的 0 當數字吃掉（例：0912345678 → 912345678）
    sheet.getRange(lastRow, 3).setNumberFormat("@").setValue(phone);

    // 有電話才額外寫一筆到 16,800 那份試算表的「AI電話」分頁（兩門課共用同一份 AI電話名單）
    if (phone) {
      try {
        const aiSheet = SpreadsheetApp.openById(AI_PHONE_EXTERNAL_SHEET_ID).getSheetByName(AI_PHONE_SHEET_NAME);
        if (aiSheet) {
          aiSheet.appendRow([
            new Date(), data.name, '', data.email,
            data.lineId || '', data.referrer || '', data.sessionDate,
            false, false, '', zoomUrl, '', COURSE_TAG
          ]);
          const aiLastRow = aiSheet.getLastRow();
          aiSheet.getRange(aiLastRow, 3).setNumberFormat("@").setValue(phone);
        } else {
          console.error("外部試算表裡找不到分頁 [" + AI_PHONE_SHEET_NAME + "]，AI電話名單沒有同步");
        }
      } catch (aiErr) {
        console.error("寫入外部 AI電話 分頁失敗，不影響本次報名：" + aiErr.message);
      }
    }

    sendRegistrationSuccessEmail(data.email, data.name, data.sessionDate, zoomUrl);
    return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
  } catch (f) { return ContentService.createTextOutput(f.toString()); }
}

function testZoomLinks() {
  const dates = ["2026-6-27", "2026-7-2", "2026-7-4", "2026-7-7", "2026-7-14", "2026-7-16", "2026-7-21", "2026-7-25", "2026-7-27", "2026-8-1", "2026-8-4", "2026-8-6", "2026-8-11", "2026-8-20", "2026-8-25", "2026-8-29"];
  dates.forEach(d => {
    const entry = Object.entries(ZOOM_LINKS).find(([k]) => d.includes("-" + k))?.[1];
    console.log(d + " → " + (entry ? entry.time + " | " + entry.url : "無"));
  });
}

function testEmails() {
  // TODO：ZOOM_LINKS 目前是空的（場次還沒定案），先用假連結測試信件排版，等真實場次填進去後可以改回讀 ZOOM_LINKS
  const testEmail = "fengpeinet@gmail.com";
  const testName = "測試用戶";
  const testDate = "2026/07/25 14:00";
  const testZoomUrl = "https://us06web.zoom.us/j/待補";

  sendRegistrationSuccessEmail(testEmail, testName, testDate, testZoomUrl);
  sendReminderEmail(testEmail, testName, testDate, testZoomUrl, "明天");
  sendRecallEmail(testEmail, testName, "2026/07/25 14:00", testZoomUrl);

  console.log("三封測試信已送出，請檢查信箱");
}
