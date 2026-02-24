/**
 * 摩研裝修 學習資源 - Google Apps Script
 * 版本：1.1（2026-02-22）
 *
 * 功能：部署升職階梯架構、木工學習護照、精英合夥計畫為網頁
 *
 * 使用方式：
 * - 升職階梯架構：?page=ladder
 * - 木工學習護照：?page=passport
 * - 精英合夥計畫：?page=elite
 */

function doGet(e) {
  var page = e.parameter.page || 'index';

  if (page === 'ladder') {
    return HtmlService.createHtmlOutputFromFile('ladder')
      .setTitle('摩研裝修 升職階梯架構')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (page === 'passport') {
    return HtmlService.createHtmlOutputFromFile('passport')
      .setTitle('木工學習護照')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (page === 'elite') {
    return HtmlService.createHtmlOutputFromFile('elite')
      .setTitle('摩研裝修 精英合夥計畫')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else {
    return HtmlService.createHtmlOutput(getIndexPage())
      .setTitle('摩研裝修 學習資源')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function getIndexPage() {
  return `
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: 'Microsoft JhengHei', 'Noto Sans TC', sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
        }
        .container {
            background: white;
            border-radius: 20px;
            padding: 40px;
            max-width: 500px;
            width: 100%;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
        }
        h1 {
            text-align: center;
            color: #2C3E50;
            margin-bottom: 30px;
            font-size: 1.8rem;
        }
        .resource-list {
            display: flex;
            flex-direction: column;
            gap: 15px;
        }
        .resource-card {
            display: flex;
            align-items: center;
            padding: 20px;
            background: #F8F9FA;
            border-radius: 12px;
            text-decoration: none;
            color: #2C3E50;
            transition: all 0.3s;
        }
        .resource-card:hover {
            background: #E8E8E8;
            transform: translateX(5px);
        }
        .resource-icon {
            width: 50px;
            height: 50px;
            background: #3498DB;
            border-radius: 12px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.5rem;
            margin-right: 15px;
        }
        .resource-icon.orange { background: #F39C12; }
        .resource-icon.green { background: #27AE60; }
        .resource-icon.purple { background: #9B59B6; }
        .resource-content { flex: 1; }
        .resource-title { font-weight: 600; font-size: 1.1rem; }
        .resource-desc { font-size: 0.85rem; color: #7F8C8D; margin-top: 4px; }
        .arrow { color: #BDC3C7; font-size: 1.5rem; }
    </style>
</head>
<body>
    <div class="container">
        <h1>📚 摩研學習資源</h1>
        <div class="resource-list">
            <a href="?page=ladder" class="resource-card">
                <div class="resource-icon">📊</div>
                <div class="resource-content">
                    <div class="resource-title">升職階梯架構</div>
                    <div class="resource-desc">各職級定義、升級條件、薪資對照</div>
                </div>
                <span class="arrow">›</span>
            </a>
            <a href="?page=passport" class="resource-card">
                <div class="resource-icon orange">📖</div>
                <div class="resource-content">
                    <div class="resource-title">木工學習護照</div>
                    <div class="resource-desc">技能認證、學習進度</div>
                </div>
                <span class="arrow">›</span>
            </a>
            <a href="?page=elite" class="resource-card">
                <div class="resource-icon purple">⭐</div>
                <div class="resource-content">
                    <div class="resource-title">精英合夥計畫</div>
                    <div class="resource-desc">分潤制度、GAINS框架、Kaizen工具管理</div>
                </div>
                <span class="arrow">›</span>
            </a>
        </div>
    </div>
</body>
</html>
  `;
}

function testFiles() {
  try {
    var ladder = HtmlService.createHtmlOutputFromFile('ladder');
    Logger.log('ladder.html: OK');
  } catch (e) {
    Logger.log('ladder.html: 錯誤 - ' + e.message);
  }

  try {
    var passport = HtmlService.createHtmlOutputFromFile('passport');
    Logger.log('passport.html: OK');
  } catch (e) {
    Logger.log('passport.html: 錯誤 - ' + e.message);
  }

  try {
    var elite = HtmlService.createHtmlOutputFromFile('elite');
    Logger.log('elite.html: OK');
  } catch (e) {
    Logger.log('elite.html: 錯誤 - ' + e.message);
  }
}
