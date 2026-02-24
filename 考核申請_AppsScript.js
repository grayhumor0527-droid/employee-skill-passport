/**
 * 摩研裝修 考核申請表單 - Apps Script
 * 表單連結：https://forms.gle/nTHruXr1Yfc64CJG7
 *
 * 設定步驟：
 * 1. 開啟 Google Forms 編輯頁面
 * 2. 右上角 ⋮ → 指令碼編輯器
 * 3. 貼上此程式碼
 * 4. 修改 CONFIG.ADMIN_EMAIL 為你的 Email
 * 5. 儲存（Ctrl+S）
 * 6. 執行 → testEmail（測試 Email 是否能寄出）
 * 7. 觸發條件 → 新增觸發器：
 *    - 函式：onFormSubmit
 *    - 事件來源：表單
 *    - 事件類型：提交表單時
 */

// ===== 設定 =====
const CONFIG = {
  ADMIN_EMAIL: 'your-email@gmail.com',  // ← 改成你的 Email
  SHEET_ID: '',  // 選填：如要寫入試算表，填入試算表 ID
  SHEET_NAME: '考核申請'
};

// ===== 表單提交處理 =====
function onFormSubmit(e) {
  try {
    const responses = e.namedValues;

    // 取得表單資料
    const data = {
      email: responses['電子郵件地址'] ? responses['電子郵件地址'][0] : '',
      name: responses['姓名'][0],
      phone4: responses['手機末4碼（驗證用）'][0],
      appType: responses['申請類型'][0],
      currentRank: responses['目前職級'][0],
      scoreRange: responses['技能分數區間'][0],
      projects: responses['主要參與案場'] ? responses['主要參與案場'].join('、') : '',
      skills: responses['技術專長'] ? responses['技術專長'].join('、') : '',
      reason: responses['申請原因'][0],
      note: responses['其他說明'] ? responses['其他說明'][0] : '',
      timestamp: new Date()
    };

    // 產生申請編號
    const appId = 'APP-' + Utilities.formatDate(data.timestamp, 'Asia/Taipei', 'yyyyMMdd-HHmmss');

    // 發送 Email 通知
    sendNotification(appId, data);

    // 寫入試算表（如有設定）
    if (CONFIG.SHEET_ID) {
      logToSheet(appId, data);
    }

    Logger.log('申請處理完成：' + appId);

  } catch (error) {
    Logger.log('錯誤：' + error.message);
    // 發送錯誤通知
    MailApp.sendEmail({
      to: CONFIG.ADMIN_EMAIL,
      subject: '【摩研】考核申請系統錯誤',
      body: '錯誤訊息：' + error.message + '\n\n原始資料：' + JSON.stringify(e.namedValues)
    });
  }
}

// ===== 發送通知 =====
function sendNotification(appId, data) {
  const subject = '【摩研】考核申請 - ' + data.name + '（' + data.appType + '）';

  const body = `
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  考核申請通知
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

申請編號：${appId}
申請時間：${Utilities.formatDate(data.timestamp, 'Asia/Taipei', 'yyyy-MM-dd HH:mm')}

【申請人資料】
姓　　名：${data.name}
手機末4碼：${data.phone4}
電子郵件：${data.email}

【申請內容】
申請類型：${data.appType}
目前職級：${data.currentRank}
技能分數：${data.scoreRange}

【經歷與專長】
代表案場：${data.projects || '未填寫'}
技術專長：${data.skills || '未填寫'}

【申請原因】
${data.reason}

${data.note ? '【其他說明】\n' + data.note + '\n' : ''}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

請至員工主檔確認資格後進行審核。

此為系統自動發送，請勿直接回覆。
`;

  MailApp.sendEmail({
    to: CONFIG.ADMIN_EMAIL,
    subject: subject,
    body: body
  });

  Logger.log('Email 已發送至：' + CONFIG.ADMIN_EMAIL);
}

// ===== 寫入試算表 =====
function logToSheet(appId, data) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  // 如果工作表不存在，建立新的
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.appendRow([
      '申請編號', '申請時間', '姓名', '手機末4碼', '電子郵件',
      '申請類型', '目前職級', '技能分數區間', '代表案場', '技術專長',
      '申請原因', '其他說明', '狀態', '審核者', '審核時間', '審核意見'
    ]);
  }

  sheet.appendRow([
    appId,
    Utilities.formatDate(data.timestamp, 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss'),
    data.name,
    data.phone4,
    data.email,
    data.appType,
    data.currentRank,
    data.scoreRange,
    data.projects,
    data.skills,
    data.reason,
    data.note,
    '待審核',
    '',
    '',
    ''
  ]);

  Logger.log('已寫入試算表');
}

// ===== 測試函式 =====

/**
 * 測試 Email 發送（先執行這個確認能收到信）
 */
function testEmail() {
  MailApp.sendEmail({
    to: CONFIG.ADMIN_EMAIL,
    subject: '【摩研】考核申請系統測試',
    body: '這是一封測試信件。\n\n如果您收到這封信，表示 Email 通知功能正常運作。\n\n測試時間：' + new Date().toLocaleString('zh-TW')
  });

  Logger.log('測試信已發送至：' + CONFIG.ADMIN_EMAIL);
}

/**
 * 模擬表單提交（測試用）
 */
function testFormSubmit() {
  const mockEvent = {
    namedValues: {
      '電子郵件地址': ['test@example.com'],
      '姓名': ['測試員工'],
      '手機末4碼（驗證用）': ['1234'],
      '申請類型': ['升級考核'],
      '目前職級': ['1級 學徒'],
      '技能分數區間': ['81-160分'],
      '主要參與案場': ['德富2期', '瑞典湖濱'],
      '技術專長': ['天花板', '隔間'],
      '申請原因': ['已完成學徒階段所有技能學習，希望申請升級考核。'],
      '其他說明': ['無']
    }
  };

  onFormSubmit(mockEvent);
}
