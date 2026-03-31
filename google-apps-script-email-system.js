/**
 * ============================================================
 *  中華航空 × 2026 高雄旅展  —  電子抽獎券自動寄信系統
 *  Google Apps Script (GAS) + Google Sheets 後端
 * ============================================================
 *
 *  部署步驟：
 *  1. 開啟 Google Sheets → 建立新試算表，命名為「2026高雄旅展_抽獎名單」
 *  2. 建立兩個分頁：
 *     - "名單" : 存放所有報名資料
 *     - "已寄出" : 記錄每一封已發送的 EMAIL
 *  3. 點選「擴充功能」→「Apps Script」→ 貼上本程式碼
 *  4. 點選「部署」→「新增部署作業」→ 類型選「Web 應用程式」
 *     - 執行身分：「我」(你的 Google 帳戶)
 *     - 存取權：「所有人」(允許前端 fetch 呼叫)
 *  5. 複製部署 URL，貼回前端 HTML 的 fetch() 網址
 *  ============================================================
 */

// ─── 全域設定 ───────────────────────────────────────────────
const SPREADSHEET_ID = 'YOUR_GOOGLE_SHEET_ID'; // 試算表 ID（從網址列取得）
const SENDER_NAME    = '中華航空 × 2026 高雄旅展';
const EVENT_DATES    = '2026/05/15 – 05/18';
const EVENT_VENUE    = '高雄展覽館 中華航空攤位';
const RAFFLE_TIME    = '每日 15:30';

// 任務類型 → 中文說明、獎勵張數
const TASK_CONFIG = {
  'dynasty-member': { label: '華夏會員註冊',   tickets: 2, emoji: '✈️' },
  'app-download':   { label: '下載華航 APP',    tickets: 2, emoji: '📱' },
  'line-friend':    { label: '加入 LINE 好友',   tickets: 1, emoji: '💬' },
  'youtube-sub':    { label: '訂閱 YouTube 頻道', tickets: 1, emoji: '▶️' },
  'photo-checkin':  { label: '拍照打卡任務',     tickets: 1, emoji: '📸' },
};

// ─── 主入口：接收前端 POST 請求 ─────────────────────────────
function doPost(e) {
  try {
    const data   = JSON.parse(e.postData.contents);
    const result = processEntry(data);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── 處理報名資料 ────────────────────────────────────────────
function processEntry(data) {
  const { email, name, taskType, questionnaire } = data;

  // 驗證
  if (!email || !taskType) throw new Error('缺少必填欄位');
  if (!TASK_CONFIG[taskType])  throw new Error('不明任務類型');

  const task        = TASK_CONFIG[taskType];
  const ticketNums  = generateTicketNumbers(task.tickets);
  const timestamp   = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });

  // 寫入 Google Sheets
  saveToSheet({
    timestamp,
    email,
    name:         name  || '旅客',
    taskType,
    taskLabel:    task.label,
    tickets:      task.tickets,
    ticketNums:   ticketNums.join(', '),
    q1:           questionnaire?.q1 || '',
    q2:           questionnaire?.q2 || '',
    q3:           questionnaire?.q3 || '',
  });

  // 寄出 HTML 電子抽獎券
  sendTicketEmail({
    email,
    name:       name || '旅客',
    task,
    ticketNums,
    timestamp,
  });

  return { success: true, ticketNums };
}

// ─── 產生抽獎券號碼 ──────────────────────────────────────────
function generateTicketNumbers(count) {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  const nums  = [];
  for (let i = 0; i < count; i++) {
    let num = '';
    for (let j = 0; j < 7; j++) {
      num += chars[Math.floor(Math.random() * chars.length)];
    }
    nums.push('KTF-' + num);
  }
  return nums;
}

// ─── 寫入試算表 ──────────────────────────────────────────────
function saveToSheet(row) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('名單') || ss.insertSheet('名單');

  // 如果是第一列，補上標題
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      '時間戳記','Email','姓名','任務代碼','任務名稱',
      '票數','抽獎券號碼','Q1來源','Q2夢想目的地','Q3會員狀態'
    ]);
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#003087').setFontColor('#ffffff');
  }

  sheet.appendRow([
    row.timestamp, row.email, row.name,
    row.taskType,  row.taskLabel, row.tickets, row.ticketNums,
    row.q1, row.q2, row.q3
  ]);
}

// ─── 寄出 HTML 電子抽獎券 ────────────────────────────────────
function sendTicketEmail({ email, name, task, ticketNums, timestamp }) {
  const subject = `🎫【中華航空 × 2026 高雄旅展】您的電子抽獎券已送達！`;
  const html    = buildEmailHtml({ name, task, ticketNums, timestamp });

  GmailApp.sendEmail(email, subject, '（請以 HTML 格式查看此郵件）', {
    name:     SENDER_NAME,
    htmlBody: html,
  });

  // 記錄已寄出
  logSent(email, ticketNums.join(', '));
}

function logSent(email, ticketNums) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('已寄出') || ss.insertSheet('已寄出');
  sheet.appendRow([
    new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' }),
    email, ticketNums, '已送出'
  ]);
}

// ─── 產生 HTML 電子抽獎券 Email ──────────────────────────────
function buildEmailHtml({ name, task, ticketNums, timestamp }) {
  const ticketBlocks = ticketNums.map((num, i) => {
    const gradients = [
      'background:linear-gradient(135deg,#5B3DC8,#C83D8B);',
      'background:linear-gradient(135deg,#C8102E,#F08020);',
      'background:linear-gradient(135deg,#003087,#3B7FD4);',
    ];
    const grad = gradients[i % gradients.length];
    return `
    <!-- Ticket ${i+1} -->
    <table width="100%" cellpadding="0" cellspacing="0" border="0"
           style="max-width:480px;margin:0 auto 20px;border-radius:20px;overflow:hidden;">
      <tr>
        <td style="${grad}padding:24px 28px;">
          <table width="100%" cellpadding="0" cellspacing="0">
            <tr>
              <td>
                <div style="font-size:10px;letter-spacing:3px;color:rgba(255,255,255,0.7);text-transform:uppercase;">
                  CHINA AIRLINES RAFFLE TICKET
                </div>
                <div style="font-size:22px;font-weight:900;color:#fff;margin-top:2px;letter-spacing:2px;">
                  中華航空 電子抽獎券
                </div>
              </td>
              <td align="right">
                <div style="font-size:10px;color:rgba(255,255,255,0.7);">第 ${i+1} 張</div>
                <div style="font-size:28px;font-weight:900;color:#fff;font-family:monospace;letter-spacing:3px;">${num}</div>
              </td>
            </tr>
          </table>
          <div style="border-top:2px dashed rgba(255,255,255,0.3);margin:16px 0;"></div>
          <table width="100%" cellpadding="0" cellspacing="0">
            <tr>
              <td>
                <div style="font-size:11px;color:rgba(255,255,255,0.7);">任務</div>
                <div style="font-size:14px;font-weight:800;color:#fff;">${task.emoji} ${task.label}</div>
              </td>
              <td align="right">
                <div style="font-size:11px;color:rgba(255,255,255,0.7);">持票人</div>
                <div style="font-size:14px;font-weight:800;color:#fff;">${name}</div>
              </td>
            </tr>
          </table>
          <div style="margin-top:14px;background:rgba(255,255,255,0.15);border-radius:10px;padding:10px 14px;font-size:11px;color:rgba(255,255,255,0.85);line-height:1.8;">
            📅 展期 ${EVENT_DATES}　|　📍 ${EVENT_VENUE}<br>
            🕒 現場抽獎 ${RAFFLE_TIME}，請準時蒞臨對獎
          </div>
        </td>
      </tr>
    </table>`;
  }).join('');

  return `<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
</head>
<body style="margin:0;padding:0;background:#f0e8ff;font-family:'Helvetica Neue',Arial,'Microsoft JhengHei',sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f0e8ff;padding:24px 16px;">
  <tr><td align="center">

    <!-- Header -->
    <table width="100%" cellpadding="0" cellspacing="0" border="0"
           style="max-width:480px;margin:0 auto 20px;background:linear-gradient(135deg,#003087,#5B3DC8);border-radius:24px;overflow:hidden;">
      <tr>
        <td style="padding:32px 28px;text-align:center;">
          <div style="font-size:48px;margin-bottom:12px;">✈️🎫</div>
          <div style="font-size:13px;letter-spacing:3px;color:rgba(255,255,255,0.7);text-transform:uppercase;margin-bottom:8px;">
            CHINA AIRLINES × KTF 2026
          </div>
          <div style="font-size:26px;font-weight:900;color:#fff;letter-spacing:2px;margin-bottom:8px;">
            您的電子抽獎券已送達！
          </div>
          <div style="font-size:13px;color:rgba(255,255,255,0.8);">
            親愛的 ${name}，感謝您參與活動 🎉<br>
            以下是您完成「${task.emoji} ${task.label}」獲得的 ${task.tickets} 張抽獎券
          </div>
        </td>
      </tr>
    </table>

    <!-- Tickets -->
    ${ticketBlocks}

    <!-- Prize Table -->
    <table width="100%" cellpadding="0" cellspacing="0" border="0"
           style="max-width:480px;margin:0 auto 20px;background:#fff;border-radius:20px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,0.08);">
      <tr>
        <td style="padding:24px;">
          <div style="font-size:16px;font-weight:800;color:#003087;margin-bottom:16px;text-align:center;">🏆 獎項說明</div>
          <table width="100%" cellpadding="8" cellspacing="0" style="border-collapse:collapse;font-size:13px;">
            <tr style="background:#003087;color:#fff;">
              <td style="padding:10px 14px;border-radius:8px 0 0 0;font-weight:800;">獎項</td>
              <td style="padding:10px 14px;">獎品內容</td>
              <td style="padding:10px 14px;border-radius:0 8px 0 0;text-align:center;">名額</td>
            </tr>
            <tr style="background:#fff8e8;">
              <td style="padding:10px 14px;font-weight:800;color:#C8102E;">🥇 頭獎</td>
              <td style="padding:10px 14px;">來回機票（高雄出發）</td>
              <td style="padding:10px 14px;text-align:center;font-weight:800;">每日 1 名</td>
            </tr>
            <tr>
              <td style="padding:10px 14px;font-weight:800;color:#B8960C;">🥈 二獎</td>
              <td style="padding:10px 14px;">聯名禮品 / 飯店折扣券</td>
              <td style="padding:10px 14px;text-align:center;font-weight:800;">每日 3 名</td>
            </tr>
            <tr style="background:#f8f0ff;">
              <td style="padding:10px 14px;font-weight:800;color:#8B3DC8;">🥉 三獎</td>
              <td style="padding:10px 14px;">華航抱枕 / 飛機模型</td>
              <td style="padding:10px 14px;text-align:center;font-weight:800;">每日 5 名</td>
            </tr>
            <tr>
              <td style="padding:10px 14px;font-weight:800;color:#333;">✈️ 參加獎</td>
              <td style="padding:10px 14px;">機票折價券（500/1000/1500元）</td>
              <td style="padding:10px 14px;text-align:center;font-weight:800;">數名</td>
            </tr>
          </table>
        </td>
      </tr>
    </table>

    <!-- Reminder -->
    <table width="100%" cellpadding="0" cellspacing="0" border="0"
           style="max-width:480px;margin:0 auto 20px;background:#fff;border-radius:20px;overflow:hidden;">
      <tr>
        <td style="padding:20px 24px;">
          <div style="font-size:14px;font-weight:800;color:#003087;margin-bottom:12px;">📋 重要提醒</div>
          <div style="font-size:13px;color:#555;line-height:2.2;">
            🕒 現場抽獎時間：<strong style="color:#C8102E;">每日 15:30</strong>（共 4 天）<br>
            📍 地點：高雄展覽館 中華航空攤位<br>
            📬 中獎通知將另外以 Email 通知<br>
            📞 請妥善保存此抽獎券號碼<br>
            ⚠️ 每人每個任務限領一次獎勵
          </div>
        </td>
      </tr>
    </table>

    <!-- Footer -->
    <table width="100%" cellpadding="0" cellspacing="0" border="0"
           style="max-width:480px;margin:0 auto;">
      <tr>
        <td style="padding:16px;text-align:center;font-size:11px;color:#aaa;line-height:2;">
          發送時間：${timestamp}<br>
          中華航空股份有限公司 ©2026<br>
          如有疑問請洽展覽現場服務人員
        </td>
      </tr>
    </table>

  </td></tr>
</table>
</body>
</html>`;
}

// ─── 測試函式（在 GAS 編輯器直接執行測試）────────────────────
function testSendEmail() {
  const testData = {
    email:    'test@example.com',  // ← 改成你的 email 測試
    name:     '測試旅客',
    taskType: 'dynasty-member',
    questionnaire: { q1: '🚇 大眾運輸', q2: '🗼 日本', q3: '✅ 是，已是會員' }
  };
  const result = processEntry(testData);
  Logger.log('測試結果：' + JSON.stringify(result));
}
