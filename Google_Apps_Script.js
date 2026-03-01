/**
 * 摩研裝修 員工自助系統 - Google Apps Script
 * 版本：4.0（2026-02-22 更新，薪資自動計算）
 *
 * 功能：
 * 1. 手機號碼 + 密碼登入
 * 2. 員工自行修改密碼
 * 3. 薪資查詢（從出勤紀錄自動計算）
 * 4. 薪資確認簽章
 *
 * 計算規則（來自 config.json）：
 * - 日薪：employee_rates
 * - 外地加給：remote_allowances（德富2期）
 * - 晚餐補貼：100 元（工時 >= 1.5 天）
 *
 * 部署方式：
 * 1. 開啟 Google Sheets「摩研員工主檔」
 * 2. 擴充功能 → Apps Script
 * 3. 貼上此腳本
 * 4. 部署為網頁應用程式
 */

const CONFIG = {
  SHEET_EMPLOYEES: '員工主檔',
  SHEET_SALARY_CONFIRM: '薪資確認紀錄',
  ADMIN_EMAIL: 'your-email@gmail.com',

  // 出勤紀錄試算表（與員工主檔不同）
  ATTENDANCE_SPREADSHEET_ID: '1TVhS5xxugzCQ-CQ5XtbO_IiUgvEhAyaKykgECJBuHR0',
  ATTENDANCE_SHEET_NAME: '表單回覆 1',

  // 員工日薪（來自 config.json employee_rates，已含中餐補貼）
  EMPLOYEE_RATES: {
    '翁建信': 3600, '劉明智': 3600, '陳瑋鑫': 3600, '陳冠興': 2300,
    '林哲立': 3600, '張城少': 3600, '劉哲林': 2000, '許子軒': 1600,
    '陳偉瀚': 3000, '智凱學徒': 2500, '劉星沅': 3500, '王瑞富': 3100,
    '溫智勳': 3500, '徐師傅': 3500, '洪豐洋父子': 3300
  },

  // 外地加給（來自 config.json remote_allowances）
  REMOTE_ALLOWANCES: {
    '翁建信': 450, '劉明智': 450, '陳瑋鑫': 450, '陳冠興': 288,
    '林哲立': 450, '張城少': 450, '劉哲林': 250, '許子軒': 200,
    '劉星沅': 438, '王瑞富': 388
  },

  // 外地案場
  REMOTE_PROJECTS: ['德富2期'],

  // 晚餐補貼（加班 >= 1.5 天時）
  DINNER_ALLOWANCE: 100,
  DINNER_THRESHOLD: 1.5
};

// ===== HTTP 請求處理 =====

function doGet(e) {
  var action = e.parameter.action;
  var phone = e.parameter.phone;
  var password = e.parameter.password;
  var name = e.parameter.name;
  var period = e.parameter.period;

  var result;

  if (action === 'login') {
    result = loginWithPassword(phone, password);
  } else if (action === 'getSalary') {
    result = getSalaryRecords(name, period);
  } else if (action === 'getSalaryPeriods') {
    result = getSalaryPeriods(name);
  } else {
    result = {error: '未知的操作'};
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var action = data.action;

  var result;

  if (action === 'changePassword') {
    result = changePassword(data.phone, data.oldPassword, data.newPassword);
  } else if (action === 'confirmSalary') {
    result = confirmSalary(data);
  } else {
    result = {error: '未知的操作'};
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== 登入功能 =====

/**
 * 手機號碼 + 密碼登入
 * @param {string} phone - 手機號碼
 * @param {string} password - 密碼（預設為手機末4碼）
 */
function loginWithPassword(phone, password) {
  if (!phone || !password) {
    return {error: '請輸入手機號碼和密碼'};
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_EMPLOYEES);
  if (!sheet) return {error: '系統錯誤'};

  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var passwordCol = findColumn(headers, '密碼');

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var empPhone = String(row[2]).trim();  // C欄：手機號碼

    if (empPhone === phone) {
      // 取得密碼（如果有密碼欄則用密碼欄，否則用手機末4碼）
      var correctPassword;
      if (passwordCol >= 0 && row[passwordCol]) {
        correctPassword = String(row[passwordCol]);
      } else {
        correctPassword = phone.slice(-4);  // 預設：手機末4碼
      }

      if (password === correctPassword) {
        // 登入成功，回傳員工資料（排除敏感欄位）
        var emp = {};
        for (var j = 0; j < headers.length; j++) {
          if (headers[j] !== '密碼' && headers[j] !== '身分證字號') {
            emp[headers[j]] = row[j];
          }
        }
        emp.phone = empPhone;
        return {success: true, data: emp};
      } else {
        return {error: '密碼錯誤'};
      }
    }
  }

  return {error: '找不到此手機號碼'};
}

// ===== 密碼修改 =====

/**
 * 修改密碼
 * @param {string} phone - 手機號碼
 * @param {string} oldPassword - 舊密碼
 * @param {string} newPassword - 新密碼
 */
function changePassword(phone, oldPassword, newPassword) {
  if (!phone || !oldPassword || !newPassword) {
    return {error: '請填寫完整資料'};
  }

  if (newPassword.length < 4) {
    return {error: '新密碼至少需要 4 位'};
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_EMPLOYEES);
  if (!sheet) return {error: '系統錯誤'};

  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var passwordCol = findColumn(headers, '密碼');

  // 如果沒有密碼欄位，新增一個
  if (passwordCol < 0) {
    passwordCol = headers.length;
    sheet.getRange(1, passwordCol + 1).setValue('密碼');
  }

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var empPhone = String(row[2]).trim();

    if (empPhone === phone) {
      // 驗證舊密碼
      var currentPassword;
      if (passwordCol < headers.length && row[passwordCol]) {
        currentPassword = String(row[passwordCol]);
      } else {
        currentPassword = phone.slice(-4);  // 預設：手機末4碼
      }

      if (oldPassword !== currentPassword) {
        return {error: '目前密碼錯誤'};
      }

      // 更新密碼
      sheet.getRange(i + 1, passwordCol + 1).setValue(newPassword);
      return {success: true, message: '密碼修改成功'};
    }
  }

  return {error: '找不到此員工'};
}

// ===== 薪資查詢（自動計算） =====

/**
 * 從出勤紀錄自動計算薪資
 * @param {string} name - 員工姓名
 * @param {string} period - 薪資期間（如：2026-02-上）
 */
function getSalaryRecords(name, period) {
  if (!name || !period) {
    return {error: '參數不完整'};
  }

  // 開啟出勤紀錄試算表（與員工主檔不同）
  var ss = SpreadsheetApp.openById(CONFIG.ATTENDANCE_SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET_NAME);

  if (!sheet) {
    return {error: '找不到出勤紀錄工作表'};
  }

  // 解析期間（如：2026-02-上 → year=2026, month=2, half=上）
  var periodMatch = period.match(/(\d{4})-(\d{2})-(上|下)/);
  if (!periodMatch) {
    return {error: '期間格式錯誤，應為 YYYY-MM-上/下'};
  }

  var year = parseInt(periodMatch[1]);
  var month = parseInt(periodMatch[2]);
  var half = periodMatch[3];
  var startDay = (half === '上') ? 1 : 16;
  var endDay = (half === '上') ? 15 : new Date(year, month, 0).getDate();

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var records = [];
  var processed = {};  // 去重用：日期+案場+員工

  // Google Forms 欄位對應
  var colTimestamp = findColumn(headers, '時間戳記');
  var colDate = findColumn(headers, '日期（日期）');
  var colProject = findColumn(headers, '案件名稱');
  var colEmployees = findColumn(headers, '出勤員工');
  var colWorkType = findColumn(headers, '工時類型');
  var colReportType = findColumn(headers, '填報類型');

  // 取得員工日薪
  var dailyRate = CONFIG.EMPLOYEE_RATES[name] || 0;
  var remoteAllowance = CONFIG.REMOTE_ALLOWANCES[name] || 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    // 跳過非出勤紀錄
    var reportType = row[colReportType];
    if (reportType && reportType !== '出勤紀錄') continue;

    // 解析日期
    var dateVal = row[colDate];
    var recordDate;
    if (dateVal instanceof Date) {
      recordDate = dateVal;
    } else {
      recordDate = new Date(dateVal);
    }

    if (isNaN(recordDate.getTime())) continue;

    // 檢查是否在期間內
    var rYear = recordDate.getFullYear();
    var rMonth = recordDate.getMonth() + 1;
    var rDay = recordDate.getDate();

    if (rYear !== year || rMonth !== month) continue;
    if (rDay < startDay || rDay > endDay) continue;

    // 檢查員工是否在出勤名單中
    var employees = String(row[colEmployees] || '');
    if (employees.indexOf(name) === -1) continue;

    // 計算工時（直接解析數字）
    var workType = row[colWorkType];
    var hours = parseFloat(workType) || 1;

    var project = row[colProject] || '';

    // 去重檢查：同一天+同一案場+同一員工 只計算一次
    var uniqueKey = rYear + '-' + rMonth + '-' + rDay + '_' + project + '_' + name;
    if (processed[uniqueKey]) continue;
    processed[uniqueKey] = true;

    // 計算薪資
    var salary = Math.round(dailyRate * hours);

    // 計算外地加給（有出勤即發，不乘工時）
    var remote = 0;
    if (CONFIG.REMOTE_PROJECTS.indexOf(project) >= 0) {
      remote = remoteAllowance;
    }

    // 計算晚餐補貼（加班 >= 1.5 天）
    var dinner = 0;
    if (hours >= CONFIG.DINNER_THRESHOLD) {
      dinner = CONFIG.DINNER_ALLOWANCE;
    }

    records.push({
      date: (rMonth) + '/' + rDay,
      project: project,
      hours: hours,
      salary: salary,
      remote: remote,
      dinner: dinner
    });
  }

  // 處理獎金和代墊款（從支出紀錄中查找）
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var reportType = row[colReportType];

    // 只處理支出報銷紀錄
    if (reportType !== '支出報銷') continue;

    // 解析日期
    var dateVal = row[colDate];
    var recordDate;
    if (dateVal instanceof Date) {
      recordDate = dateVal;
    } else {
      recordDate = new Date(dateVal);
    }

    if (isNaN(recordDate.getTime())) continue;

    // 檢查是否在期間內
    var rYear = recordDate.getFullYear();
    var rMonth = recordDate.getMonth() + 1;
    var rDay = recordDate.getDate();

    if (rYear !== year || rMonth !== month) continue;
    if (rDay < startDay || rDay > endDay) continue;

    // 檢查備註欄是否包含此員工的獎金或代墊
    var colNote = findColumn(headers, '品項說明');
    var colAmount = findColumn(headers, '金額');
    var note = String(row[colNote] || '');
    var amount = parseFloat(row[colAmount]) || 0;

    // 檢查是否是此員工的獎金或代墊（備註中包含員工姓名）
    if (note.indexOf(name) >= 0 || note.indexOf('員工獎金(' + name + ')') >= 0) {
      var project = row[colProject] || '';
      var isBonus = note.indexOf('獎金') >= 0 || note.indexOf('津貼') >= 0 || note.indexOf('補貼') >= 0;
      var isAdvance = note.indexOf('代墊') >= 0;

      if (isBonus) {
        records.push({
          date: rMonth + '/' + rDay,
          project: project,
          hours: 0,
          salary: 0,
          remote: 0,
          dinner: 0,
          bonus: amount,
          remark: note
        });
      } else if (isAdvance || note.indexOf(name + '代墊') >= 0) {
        records.push({
          date: rMonth + '/' + rDay,
          project: project,
          hours: 0,
          salary: 0,
          remote: 0,
          dinner: 0,
          advance: amount,
          remark: note
        });
      }
    }
  }

  // 依日期排序
  records.sort(function(a, b) {
    var dA = parseInt(a.date.split('/')[1]);
    var dB = parseInt(b.date.split('/')[1]);
    return dA - dB;
  });

  return {success: true, data: records};
}

/**
 * 取得員工有出勤紀錄的薪資期間
 * @param {string} name - 員工姓名
 */
function getSalaryPeriods(name) {
  if (!name) {
    return {error: '請提供員工姓名'};
  }

  // 開啟出勤紀錄試算表（與員工主檔不同）
  var ss = SpreadsheetApp.openById(CONFIG.ATTENDANCE_SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET_NAME);

  if (!sheet) {
    return {error: '找不到出勤紀錄工作表'};
  }

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var periods = {};

  var colDate = findColumn(headers, '日期（日期）');
  var colEmployees = findColumn(headers, '出勤員工');
  var colReportType = findColumn(headers, '填報類型');

  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    // 跳過非出勤紀錄
    var reportType = row[colReportType];
    if (reportType && reportType !== '出勤紀錄') continue;

    // 檢查員工是否在出勤名單中
    var employees = String(row[colEmployees] || '');
    if (employees.indexOf(name) === -1) continue;

    // 解析日期
    var dateVal = row[colDate];
    var recordDate;
    if (dateVal instanceof Date) {
      recordDate = dateVal;
    } else {
      recordDate = new Date(dateVal);
    }

    if (isNaN(recordDate.getTime())) continue;

    var year = recordDate.getFullYear();
    var month = recordDate.getMonth() + 1;
    var day = recordDate.getDate();
    var half = (day <= 15) ? '上' : '下';
    var monthStr = (month < 10 ? '0' : '') + month;
    var periodKey = year + '-' + monthStr + '-' + half;

    periods[periodKey] = true;
  }

  // 轉換為陣列並排序（最新的在前）
  var periodList = Object.keys(periods).sort().reverse();

  return {success: true, data: periodList};
}

// ===== 薪資確認 =====

/**
 * 記錄薪資確認簽章
 * @param {Object} data - 確認資料
 */
function confirmSalary(data) {
  if (!data.name || !data.period || !data.amount) {
    return {error: '資料不完整'};
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_SALARY_CONFIRM);

  // 如果工作表不存在，建立新的
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_SALARY_CONFIRM);
    sheet.appendRow(['確認編號', '員工姓名', '薪資期間', '確認金額', '確認時間', 'IP位址', '裝置資訊']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#E8E8E8');
  }

  // 檢查是否已確認過
  var existingData = sheet.getDataRange().getValues();
  for (var i = 1; i < existingData.length; i++) {
    if (existingData[i][1] === data.name && existingData[i][2] === data.period) {
      return {
        success: true,
        message: '此期間已確認過',
        alreadyConfirmed: true,
        timestamp: existingData[i][4]
      };
    }
  }

  // 產生確認編號
  var confirmId = 'SC-' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd-HHmmss');

  // 確認時間
  var timestamp = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');

  // 寫入紀錄
  sheet.appendRow([
    confirmId,
    data.name,
    data.period,
    data.amount,
    timestamp,
    data.ip || '',
    data.device || ''
  ]);

  return {
    success: true,
    message: '薪資確認成功',
    confirmId: confirmId,
    timestamp: timestamp
  };
}

/**
 * 查詢薪資確認紀錄
 * @param {string} name - 員工姓名
 */
function getSalaryConfirmations(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_SALARY_CONFIRM);

  if (!sheet) {
    return {success: true, data: []};
  }

  var data = sheet.getDataRange().getValues();
  var records = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === name) {
      records.push({
        confirmId: data[i][0],
        name: data[i][1],
        period: data[i][2],
        amount: data[i][3],
        timestamp: data[i][4]
      });
    }
  }

  return {success: true, data: records};
}

// ===== 工具函數 =====

/**
 * 尋找欄位位置
 * @param {Array} headers - 標題列
 * @param {string} name - 欄位名稱
 * @returns {number} 欄位索引，找不到則回傳 -1
 */
function findColumn(headers, name) {
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] === name) return i;
  }
  return -1;
}

// ===== 測試函數 =====

/**
 * 測試登入功能（在 Apps Script 編輯器執行）
 */
function testLogin() {
  var result = loginWithPassword('0988139317', '9317');
  Logger.log(result);
}

/**
 * 測試薪資查詢（自動計算）
 */
function testSalary() {
  // 測試取得薪資期間
  var periods = getSalaryPeriods('翁建信');
  Logger.log('可用期間: ' + JSON.stringify(periods));

  // 測試薪資計算
  if (periods.data && periods.data.length > 0) {
    var salary = getSalaryRecords('翁建信', periods.data[0]);
    Logger.log('薪資紀錄: ' + JSON.stringify(salary));

    // 計算總額
    if (salary.data) {
      var total = 0;
      salary.data.forEach(function(r) {
        total += r.salary + r.remote + r.dinner;
      });
      Logger.log('本期總額: ' + total);
    }
  }
}

/**
 * 檢查工作表結構
 */
function testDebug() {
  // 檢查員工主檔
  var empSs = SpreadsheetApp.getActiveSpreadsheet();
  var empSheet = empSs.getSheetByName('員工主檔');
  if (empSheet) {
    var empData = empSheet.getDataRange().getValues();
    Logger.log('員工主檔欄位: ' + empData[0].join(', '));
  }

  // 檢查出勤紀錄（不同試算表）
  try {
    var attSs = SpreadsheetApp.openById(CONFIG.ATTENDANCE_SPREADSHEET_ID);
    var attSheet = attSs.getSheetByName(CONFIG.ATTENDANCE_SHEET_NAME);
    if (attSheet) {
      var attData = attSheet.getDataRange().getValues();
      Logger.log('出勤紀錄欄位: ' + attData[0].join(', '));
      Logger.log('出勤紀錄筆數: ' + (attData.length - 1));
      if (attData.length > 1) {
        Logger.log('第一筆資料: ' + JSON.stringify(attData[1]));
      }
    } else {
      Logger.log('找不到工作表: ' + CONFIG.ATTENDANCE_SHEET_NAME);
    }
  } catch (e) {
    Logger.log('無法開啟出勤紀錄試算表: ' + e.message);
  }
}
