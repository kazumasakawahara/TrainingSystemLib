/**
 * =========================================================
 * 研修会管理システム ライブラリ v2.0
 * TrainingSystemLib
 * =========================================================
 * 
 * 【v2.0 新機能】
 * - 未回答者へのリマインダー送信
 * - 年間スケジュール管理
 * - 出欠一覧のPDF出力
 * - 自動送信トリガー
 */

// ============================================================
// グローバル変数（ライブラリ内部で使用）
// ============================================================
let _config = null;
let _spreadsheet = null;

// ============================================================
// 初期化関数
// ============================================================

/**
 * ライブラリを初期化する
 * 
 * @param {Object} config - 設定オブジェクト
 * @param {string} config.parentFolderId - 親フォルダID
 * @param {string} config.systemFolderName - システムフォルダ名
 * @param {string} config.attachmentFolderName - 添付資料フォルダ名
 * @param {string} config.spreadsheetName - スプレッドシート名
 * @param {string} config.senderName - 送信者名
 * @param {number} config.attendanceDeadlineDays - 回答期限（研修日の何日前）
 * @param {number} config.autoSendDaysBefore - 自動送信（研修日の何日前）※オプション
 * @param {Object} config.sheetNames - シート名の設定（オプション）
 * @returns {Object} 初期化結果
 */
function init(config) {
  _config = {
    parentFolderId: config.parentFolderId || '',
    systemFolderName: config.systemFolderName || '研修会管理システム',
    attachmentFolderName: config.attachmentFolderName || '添付資料',
    spreadsheetName: config.spreadsheetName || '研修会管理システム',
    senderName: config.senderName || '研修会事務局',
    attendanceDeadlineDays: config.attendanceDeadlineDays || 3,
    autoSendDaysBefore: config.autoSendDaysBefore || 7,
    reminderDaysBefore: config.reminderDaysBefore || 2,
    sheetNames: {
      participants: (config.sheetNames && config.sheetNames.participants) || '参加者マスター',
      currentTraining: (config.sheetNames && config.sheetNames.currentTraining) || '当日研修会',
      emailTemplate: (config.sheetNames && config.sheetNames.emailTemplate) || 'メールテンプレート',
      history: (config.sheetNames && config.sheetNames.history) || '送信履歴',
      attendance: (config.sheetNames && config.sheetNames.attendance) || '出欠回答',
      settings: (config.sheetNames && config.sheetNames.settings) || '設定',
      yearlySchedule: (config.sheetNames && config.sheetNames.yearlySchedule) || '年間スケジュール'
    }
  };
  
  if (!_config.parentFolderId) {
    throw new Error('parentFolderId は必須です');
  }
  
  return { success: true, config: _config };
}

/**
 * 現在の設定を取得
 */
function getConfig() {
  if (!_config) {
    throw new Error('ライブラリが初期化されていません。init()を先に呼び出してください。');
  }
  return _config;
}

/**
 * スプレッドシートを設定
 */
function setSpreadsheet(spreadsheet) {
  _spreadsheet = spreadsheet;
}

/**
 * 現在のスプレッドシートを取得
 */
function getSpreadsheet() {
  if (_spreadsheet) {
    return _spreadsheet;
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}


// ============================================================
// セットアップ関数
// ============================================================

/**
 * システム全体をセットアップ
 */
function setupSystem() {
  const config = getConfig();
  
  console.log('📦 研修会管理システムのセットアップを開始します...');
  
  try {
    const parentFolder = DriveApp.getFolderById(config.parentFolderId);
    console.log('✅ 親フォルダを確認しました');
    
    const systemFolder = getOrCreateFolder_(parentFolder, config.systemFolderName);
    console.log('✅ システムフォルダを作成/確認しました: ' + systemFolder.getName());
    
    const attachmentFolder = getOrCreateFolder_(systemFolder, config.attachmentFolderName);
    console.log('✅ 添付資料フォルダを作成/確認しました: ' + attachmentFolder.getName());
    
    const spreadsheet = getOrCreateSpreadsheet_(systemFolder, config.spreadsheetName);
    _spreadsheet = spreadsheet;
    console.log('✅ スプレッドシートを作成/確認しました: ' + spreadsheet.getName());
    
    setupAllSheets_(spreadsheet);
    console.log('✅ 全シートをセットアップしました');
    
    saveSystemSettings_(spreadsheet, {
      systemFolderId: systemFolder.getId(),
      attachmentFolderId: attachmentFolder.getId(),
      spreadsheetId: spreadsheet.getId()
    });
    console.log('✅ 設定を保存しました');
    
    const result = {
      success: true,
      spreadsheetId: spreadsheet.getId(),
      spreadsheetUrl: spreadsheet.getUrl(),
      systemFolderId: systemFolder.getId(),
      systemFolderUrl: systemFolder.getUrl(),
      attachmentFolderId: attachmentFolder.getId(),
      attachmentFolderUrl: attachmentFolder.getUrl()
    };
    
    const message = `
🎉 セットアップ完了！

📂 作成されたファイル：
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📁 ${config.systemFolderName}/
├── 📊 ${config.spreadsheetName}
└── 📁 ${config.attachmentFolderName}/
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📍 スプレッドシートURL:
${spreadsheet.getUrl()}

📍 添付資料フォルダURL:
${attachmentFolder.getUrl()}
`;
    console.log(message);
    
    return result;
    
  } catch (error) {
    console.error('❌ セットアップエラー: ' + error.message);
    throw error;
  }
}


// ============================================================
// メール送信機能
// ============================================================

/**
 * 研修会案内メールを送信
 */
function sendNotification(options = {}) {
  const config = getConfig();
  const spreadsheet = getSpreadsheet();
  const testMode = options.testMode || false;
  
  const trainingInfo = getTrainingInfo_(spreadsheet);
  
  if (!trainingInfo.name || !trainingInfo.date) {
    throw new Error('研修会名と開催日を入力してください');
  }
  
  let participants;
  if (testMode) {
    const myEmail = Session.getActiveUser().getEmail();
    participants = [{
      name: 'テストユーザー',
      email: myEmail,
      organization: 'テスト組織',
      position: 'テスト役職'
    }];
  } else {
    participants = getActiveParticipants_(spreadsheet);
    if (participants.length === 0) {
      throw new Error('有効な参加者がいません');
    }
  }
  
  const attachments = getAttachmentFiles_(spreadsheet);
  const template = getEmailTemplate_(spreadsheet);
  
  let successCount = 0;
  let failedEmails = [];
  
  participants.forEach(participant => {
    try {
      const personalizedBody = personalizeTemplate_(template.body, trainingInfo, participant);
      let personalizedSubject = personalizeTemplate_(template.subject, trainingInfo, participant);
      
      if (testMode) {
        personalizedSubject = '【テスト】' + personalizedSubject;
      }
      
      GmailApp.sendEmail(
        participant.email,
        personalizedSubject,
        personalizedBody,
        {
          name: config.senderName,
          attachments: attachments
        }
      );
      successCount++;
    } catch (e) {
      failedEmails.push(participant.email);
      console.error(`Failed to send to ${participant.email}: ${e.message}`);
    }
  });
  
  if (!testMode) {
    recordHistory_(spreadsheet, trainingInfo, participants, attachments, successCount, failedEmails);
  }
  
  return {
    success: true,
    totalCount: participants.length,
    successCount: successCount,
    failedEmails: failedEmails
  };
}


// ============================================================
// 【新機能】未回答者リマインダー送信
// ============================================================

/**
 * 未回答者を取得
 * @returns {Array} 未回答者リスト
 */
function getNoResponseParticipants() {
  const config = getConfig();
  const spreadsheet = getSpreadsheet();
  
  // 全参加者を取得
  const allParticipants = getActiveParticipants_(spreadsheet);
  
  // 回答済みのメールアドレスを取得
  const attendanceSheet = spreadsheet.getSheetByName(config.sheetNames.attendance);
  const lastRow = attendanceSheet.getLastRow();
  const respondedEmails = new Set();
  
  if (lastRow > 1) {
    const responses = attendanceSheet.getRange(2, 4, lastRow - 1, 1).getValues();
    responses.forEach(row => {
      if (row[0]) {
        respondedEmails.add(row[0].toString().toLowerCase().trim());
      }
    });
  }
  
  // 未回答者をフィルタリング
  const noResponseParticipants = allParticipants.filter(p => 
    !respondedEmails.has(p.email.toLowerCase().trim())
  );
  
  return noResponseParticipants;
}

/**
 * 未回答者にリマインダーを送信
 * @param {Object} options - オプション
 * @param {boolean} options.testMode - テストモード
 * @returns {Object} 送信結果
 */
function sendReminder(options = {}) {
  const config = getConfig();
  const spreadsheet = getSpreadsheet();
  const testMode = options.testMode || false;
  
  const trainingInfo = getTrainingInfo_(spreadsheet);
  
  if (!trainingInfo.name || !trainingInfo.date) {
    throw new Error('研修会名と開催日を入力してください');
  }
  
  let participants;
  if (testMode) {
    const myEmail = Session.getActiveUser().getEmail();
    participants = [{
      name: 'テストユーザー',
      email: myEmail,
      organization: 'テスト組織',
      position: 'テスト役職'
    }];
  } else {
    participants = getNoResponseParticipants();
    if (participants.length === 0) {
      return {
        success: true,
        totalCount: 0,
        successCount: 0,
        failedEmails: [],
        message: '未回答者はいません'
      };
    }
  }
  
  // リマインダー用の件名と本文
  const subject = `【リマインダー】出欠確認のお願い - ${trainingInfo.name}（${trainingInfo.date}）`;
  const bodyTemplate = `{{氏名}} 様

いつもお世話になっております。
研修会事務局です。

下記研修会の出欠確認について、まだご回答をいただいておりません。
お忙しいところ恐れ入りますが、{{回答期限}}までにご回答いただけますようお願いいたします。

━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 研修会名：{{研修会名}}
■ 開催日時：{{開催日}} {{開催時間}}
■ 会場：{{会場}}
━━━━━━━━━━━━━━━━━━━━━━━━━━

【出欠確認フォーム】
{{出欠フォームURL}}

何かご不明な点がございましたら、お気軽にお問い合わせください。

━━━━━━━━━━━━━━━━━━━━━━━━━━
研修会事務局
━━━━━━━━━━━━━━━━━━━━━━━━━━`;
  
  let successCount = 0;
  let failedEmails = [];
  
  participants.forEach(participant => {
    try {
      const personalizedBody = personalizeTemplate_(bodyTemplate, trainingInfo, participant);
      let personalizedSubject = subject;
      
      if (testMode) {
        personalizedSubject = '【テスト】' + personalizedSubject;
      }
      
      GmailApp.sendEmail(
        participant.email,
        personalizedSubject,
        personalizedBody,
        { name: config.senderName }
      );
      successCount++;
    } catch (e) {
      failedEmails.push(participant.email);
      console.error(`Failed to send reminder to ${participant.email}: ${e.message}`);
    }
  });
  
  return {
    success: true,
    totalCount: participants.length,
    successCount: successCount,
    failedEmails: failedEmails
  };
}


// ============================================================
// 【新機能】年間スケジュール管理
// ============================================================

/**
 * 年間スケジュールを取得
 * @returns {Array} スケジュールリスト
 */
function getYearlySchedule() {
  const config = getConfig();
  const spreadsheet = getSpreadsheet();
  const sheet = spreadsheet.getSheetByName(config.sheetNames.yearlySchedule);
  
  if (!sheet) {
    throw new Error('年間スケジュールシートが見つかりません');
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const schedule = [];
  
  data.forEach((row, index) => {
    if (row[0]) {  // 回数がある行のみ
      schedule.push({
        rowIndex: index + 2,
        number: row[0],
        name: row[1] || '',
        date: row[2] ? formatDate_(row[2]) : '',
        dateObj: row[2] instanceof Date ? row[2] : null,
        time: row[3] || '',
        venue: row[4] || '',
        instructor: row[5] || '',
        status: row[6] || '予定',
        note: row[7] || ''
      });
    }
  });
  
  return schedule;
}

/**
 * 次回の研修会を取得
 * @returns {Object|null} 次回の研修会情報
 */
function getNextTraining() {
  const schedule = getYearlySchedule();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  // 今日以降で最も近い研修会を探す
  const upcoming = schedule.filter(s => {
    if (!s.dateObj) return false;
    const trainingDate = new Date(s.dateObj);
    trainingDate.setHours(0, 0, 0, 0);
    return trainingDate >= today && s.status !== '完了';
  });
  
  if (upcoming.length === 0) return null;
  
  // 日付順にソート
  upcoming.sort((a, b) => a.dateObj - b.dateObj);
  
  return upcoming[0];
}

/**
 * 年間スケジュールから当日研修会シートにコピー
 * @param {number} scheduleRowIndex - 年間スケジュールの行番号
 */
function copyScheduleToCurrentTraining(scheduleRowIndex) {
  const config = getConfig();
  const spreadsheet = getSpreadsheet();
  
  const scheduleSheet = spreadsheet.getSheetByName(config.sheetNames.yearlySchedule);
  const currentSheet = spreadsheet.getSheetByName(config.sheetNames.currentTraining);
  
  const row = scheduleSheet.getRange(scheduleRowIndex, 1, 1, 8).getValues()[0];
  
  // 当日研修会シートに転記
  currentSheet.getRange('B1').setValue(row[1]);  // 研修会名
  currentSheet.getRange('B2').setValue(row[2]);  // 開催日
  currentSheet.getRange('B3').setValue(row[3]);  // 開催時間
  currentSheet.getRange('B4').setValue(row[4]);  // 会場
  currentSheet.getRange('B6').setValue(row[5]);  // 講師名
  
  return { success: true };
}

/**
 * 年間スケジュールのステータスを更新
 * @param {number} scheduleRowIndex - 行番号
 * @param {string} status - 新しいステータス
 */
function updateScheduleStatus(scheduleRowIndex, status) {
  const config = getConfig();
  const spreadsheet = getSpreadsheet();
  const sheet = spreadsheet.getSheetByName(config.sheetNames.yearlySchedule);
  
  sheet.getRange(scheduleRowIndex, 7).setValue(status);
  
  return { success: true };
}


// ============================================================
// 【新機能】出欠一覧のPDF出力
// ============================================================

/**
 * 出欠一覧をPDFとして出力
 * @returns {Object} PDF情報（URL等）
 */
function exportAttendanceToPdf() {
  const config = getConfig();
  const spreadsheet = getSpreadsheet();
  const settings = loadSystemSettings_(spreadsheet);
  
  const trainingInfo = getTrainingInfo_(spreadsheet);
  const allParticipants = getActiveParticipants_(spreadsheet);
  const attendanceSheet = spreadsheet.getSheetByName(config.sheetNames.attendance);
  
  // 回答データを取得
  const lastRow = attendanceSheet.getLastRow();
  const responseMap = new Map();
  
  if (lastRow > 1) {
    const responses = attendanceSheet.getRange(2, 1, lastRow - 1, 6).getValues();
    responses.forEach(row => {
      const email = row[3] ? row[3].toString().toLowerCase().trim() : '';
      if (email) {
        responseMap.set(email, {
          timestamp: row[0],
          name: row[2],
          attendance: row[4],
          note: row[5]
        });
      }
    });
  }
  
  // 出欠一覧を作成
  const attendanceList = allParticipants.map(p => {
    const response = responseMap.get(p.email.toLowerCase().trim());
    return {
      name: p.name,
      organization: p.organization,
      email: p.email,
      attendance: response ? response.attendance : '未回答',
      note: response ? response.note : ''
    };
  });
  
  // 集計
  const summary = {
    total: attendanceList.length,
    attend: attendanceList.filter(a => a.attendance === '出席').length,
    absent: attendanceList.filter(a => a.attendance === '欠席').length,
    undecided: attendanceList.filter(a => a.attendance === '未定').length,
    noResponse: attendanceList.filter(a => a.attendance === '未回答').length
  };
  
  // HTML形式でレポートを作成
  const html = createAttendanceReportHtml_(trainingInfo, attendanceList, summary);
  
  // PDFに変換
  const blob = Utilities.newBlob(html, 'text/html', 'report.html');
  const pdfBlob = blob.getAs('application/pdf');
  pdfBlob.setName(`出欠一覧_${trainingInfo.name}_${trainingInfo.date.replace(/\//g, '')}.pdf`);
  
  // システムフォルダに保存
  const systemFolder = DriveApp.getFolderById(settings.systemFolderId);
  const pdfFile = systemFolder.createFile(pdfBlob);
  
  return {
    success: true,
    fileId: pdfFile.getId(),
    fileName: pdfFile.getName(),
    fileUrl: pdfFile.getUrl(),
    summary: summary
  };
}

/**
 * 出欠レポートのHTMLを作成（内部関数）
 * @private
 */
function createAttendanceReportHtml_(trainingInfo, attendanceList, summary) {
  const rows = attendanceList.map(a => `
    <tr>
      <td>${a.name}</td>
      <td>${a.organization}</td>
      <td style="text-align:center;background-color:${getAttendanceColor_(a.attendance)}">${a.attendance}</td>
      <td>${a.note || ''}</td>
    </tr>
  `).join('');
  
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body { font-family: "Hiragino Sans", "Yu Gothic", sans-serif; margin: 20px; }
    h1 { font-size: 18px; border-bottom: 2px solid #333; padding-bottom: 10px; }
    .info { margin: 15px 0; }
    .info dt { font-weight: bold; float: left; width: 100px; }
    .info dd { margin-left: 110px; margin-bottom: 5px; }
    .summary { background: #f5f5f5; padding: 15px; margin: 20px 0; border-radius: 5px; }
    .summary span { margin-right: 20px; }
    table { width: 100%; border-collapse: collapse; margin-top: 20px; }
    th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
    th { background: #4a86e8; color: white; }
    tr:nth-child(even) { background: #f9f9f9; }
    .footer { margin-top: 30px; font-size: 12px; color: #666; text-align: right; }
  </style>
</head>
<body>
  <h1>出欠一覧表</h1>
  
  <dl class="info">
    <dt>研修会名</dt><dd>${trainingInfo.name}</dd>
    <dt>開催日</dt><dd>${trainingInfo.date} ${trainingInfo.time}</dd>
    <dt>会場</dt><dd>${trainingInfo.venue}</dd>
  </dl>
  
  <div class="summary">
    <strong>集計：</strong>
    <span>参加者 ${summary.total}名</span>
    <span>出席 ${summary.attend}名</span>
    <span>欠席 ${summary.absent}名</span>
    <span>未定 ${summary.undecided}名</span>
    <span>未回答 ${summary.noResponse}名</span>
  </div>
  
  <table>
    <thead>
      <tr>
        <th>氏名</th>
        <th>所属</th>
        <th>出欠</th>
        <th>備考</th>
      </tr>
    </thead>
    <tbody>
      ${rows}
    </tbody>
  </table>
  
  <div class="footer">
    作成日時: ${Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm')}
  </div>
</body>
</html>
  `;
}

/**
 * 出欠ステータスに応じた背景色を返す
 * @private
 */
function getAttendanceColor_(attendance) {
  switch (attendance) {
    case '出席': return '#d4edda';
    case '欠席': return '#f8d7da';
    case '未定': return '#fff3cd';
    default: return '#e2e3e5';
  }
}


// ============================================================
// 【新機能】自動送信トリガー
// ============================================================

/**
 * 自動送信チェック（トリガーから呼び出し）
 * 年間スケジュールをチェックし、送信タイミングの研修会があれば自動送信
 * @returns {Object} 処理結果
 */
function checkAndAutoSend() {
  const config = getConfig();
  const spreadsheet = getSpreadsheet();
  
  const schedule = getYearlySchedule();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const results = [];
  
  schedule.forEach(training => {
    if (!training.dateObj || training.status !== '予定') return;
    
    const trainingDate = new Date(training.dateObj);
    trainingDate.setHours(0, 0, 0, 0);
    
    const daysUntil = Math.floor((trainingDate - today) / (1000 * 60 * 60 * 24));
    
    // 送信タイミングかどうかチェック
    if (daysUntil === config.autoSendDaysBefore) {
      console.log(`📧 自動送信: ${training.name}（${training.date}）`);
      
      // 当日研修会シートにコピー
      copyScheduleToCurrentTraining(training.rowIndex);
      
      // メール送信
      try {
        const result = sendNotification({ testMode: false });
        updateScheduleStatus(training.rowIndex, '案内済');
        results.push({
          training: training.name,
          status: 'success',
          sent: result.successCount
        });
      } catch (e) {
        results.push({
          training: training.name,
          status: 'error',
          error: e.message
        });
      }
    }
  });
  
  return { processed: results.length, results: results };
}

/**
 * 自動リマインダーチェック（トリガーから呼び出し）
 * @returns {Object} 処理結果
 */
function checkAndAutoRemind() {
  const config = getConfig();
  const spreadsheet = getSpreadsheet();
  
  const trainingInfo = getTrainingInfo_(spreadsheet);
  
  if (!trainingInfo.date) {
    return { skipped: true, reason: '当日研修会の情報がありません' };
  }
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const trainingDate = new Date(trainingInfo.date);
  trainingDate.setHours(0, 0, 0, 0);
  
  const daysUntil = Math.floor((trainingDate - today) / (1000 * 60 * 60 * 24));
  
  // リマインダータイミングかどうかチェック
  if (daysUntil === config.reminderDaysBefore) {
    console.log(`📧 自動リマインダー: ${trainingInfo.name}`);
    return sendReminder({ testMode: false });
  }
  
  return { skipped: true, reason: `リマインダータイミングではありません（残り${daysUntil}日）` };
}

/**
 * 日次トリガーを設定
 * @param {number} hour - 実行する時刻（0-23）
 */
function setupDailyTrigger(hour) {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'dailyAutoProcess') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 新しいトリガーを作成
  ScriptApp.newTrigger('dailyAutoProcess')
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .create();
  
  return { success: true, hour: hour };
}

/**
 * トリガーを削除
 */
function removeDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'dailyAutoProcess') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });
  
  return { success: true, removed: removed };
}


// ============================================================
// 出欠確認フォーム
// ============================================================

/**
 * 出欠確認フォームを作成
 */
function createAttendanceForm() {
  const config = getConfig();
  const spreadsheet = getSpreadsheet();
  
  const trainingInfo = getTrainingInfo_(spreadsheet);
  
  if (!trainingInfo.name || !trainingInfo.date) {
    throw new Error('研修会名と開催日を入力してください');
  }
  
  const settings = loadSystemSettings_(spreadsheet);
  const systemFolder = DriveApp.getFolderById(settings.systemFolderId);
  
  const formTitle = `【出欠確認】${trainingInfo.name}（${trainingInfo.date}）`;
  const form = FormApp.create(formTitle);
  
  const formFile = DriveApp.getFileById(form.getId());
  formFile.moveTo(systemFolder);
  
  form.setDescription(
    `${trainingInfo.name}の出欠確認フォームです。\n\n` +
    `開催日時: ${trainingInfo.date} ${trainingInfo.time || ''}\n` +
    `会場: ${trainingInfo.venue || ''}`
  );
  
  const deadline = new Date(trainingInfo.date);
  deadline.setDate(deadline.getDate() - config.attendanceDeadlineDays);
  
  form.addTextItem()
    .setTitle('氏名')
    .setRequired(true);
  
  form.addTextItem()
    .setTitle('メールアドレス')
    .setRequired(true);
  
  form.addMultipleChoiceItem()
    .setTitle('出欠')
    .setChoiceValues(['出席', '欠席', '未定'])
    .setRequired(true);
  
  form.addParagraphTextItem()
    .setTitle('備考（欠席理由など）')
    .setRequired(false);
  
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());
  
  const currentSheet = spreadsheet.getSheetByName(config.sheetNames.currentTraining);
  currentSheet.getRange('B9').setValue(form.getPublishedUrl());
  currentSheet.getRange('B10').setValue(form.getId());
  
  return {
    success: true,
    formId: form.getId(),
    formUrl: form.getPublishedUrl(),
    editUrl: form.getEditUrl(),
    deadline: Utilities.formatDate(deadline, 'JST', 'yyyy/MM/dd')
  };
}


// ============================================================
// 出欠状況確認
// ============================================================

/**
 * 出欠状況を取得
 */
function getAttendanceStatus() {
  const config = getConfig();
  const spreadsheet = getSpreadsheet();
  
  const attendanceSheet = spreadsheet.getSheetByName(config.sheetNames.attendance);
  const participants = getActiveParticipants_(spreadsheet);
  
  const totalParticipants = participants.length;
  const lastRow = attendanceSheet.getLastRow();
  const responseCount = lastRow > 1 ? lastRow - 1 : 0;
  
  let attendCount = 0;
  let absentCount = 0;
  let undecidedCount = 0;
  
  if (lastRow > 1) {
    const responses = attendanceSheet.getRange(2, 5, lastRow - 1, 1).getValues();
    responses.forEach(row => {
      if (row[0] === '出席') attendCount++;
      else if (row[0] === '欠席') absentCount++;
      else if (row[0] === '未定') undecidedCount++;
    });
  }
  
  return {
    totalParticipants: totalParticipants,
    responseCount: responseCount,
    noResponseCount: totalParticipants - responseCount,
    attendCount: attendCount,
    absentCount: absentCount,
    undecidedCount: undecidedCount
  };
}


// ============================================================
// ユーティリティ関数（公開）
// ============================================================

/**
 * 当日研修会シートをクリア
 */
function clearCurrentTraining() {
  const config = getConfig();
  const spreadsheet = getSpreadsheet();
  const sheet = spreadsheet.getSheetByName(config.sheetNames.currentTraining);
  
  sheet.getRange('B1:B7').clearContent();
  sheet.getRange('B9:B10').clearContent();
  sheet.getRange('B9').setValue('（自動生成されます）');
  sheet.getRange('B10').setValue('（自動生成されます）');
  
  return { success: true };
}

/**
 * 添付資料フォルダのURLを取得
 */
function getAttachmentFolderUrl() {
  const spreadsheet = getSpreadsheet();
  const settings = loadSystemSettings_(spreadsheet);
  const folder = DriveApp.getFolderById(settings.attachmentFolderId);
  return folder.getUrl();
}

/**
 * 研修会情報を取得（公開用）
 */
function getTrainingInfo() {
  const spreadsheet = getSpreadsheet();
  return getTrainingInfo_(spreadsheet);
}

/**
 * 有効な参加者一覧を取得（公開用）
 */
function getActiveParticipants() {
  const spreadsheet = getSpreadsheet();
  return getActiveParticipants_(spreadsheet);
}


// ============================================================
// 内部ヘルパー関数（非公開）
// ============================================================

function getOrCreateFolder_(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(folderName);
}

function getOrCreateSpreadsheet_(folder, name) {
  const files = folder.getFilesByName(name);
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      return SpreadsheetApp.openById(file.getId());
    }
  }
  
  const spreadsheet = SpreadsheetApp.create(name);
  const file = DriveApp.getFileById(spreadsheet.getId());
  file.moveTo(folder);
  
  return spreadsheet;
}

function setupAllSheets_(spreadsheet) {
  setupParticipantsSheet_(spreadsheet);
  setupCurrentTrainingSheet_(spreadsheet);
  setupEmailTemplateSheet_(spreadsheet);
  setupHistorySheet_(spreadsheet);
  setupAttendanceSheet_(spreadsheet);
  setupYearlyScheduleSheet_(spreadsheet);  // 新規追加
  
  const defaultSheet = spreadsheet.getSheetByName('シート1') || spreadsheet.getSheetByName('Sheet1');
  if (defaultSheet && spreadsheet.getSheets().length > 1) {
    spreadsheet.deleteSheet(defaultSheet);
  }
}

function setupParticipantsSheet_(spreadsheet) {
  const config = getConfig();
  let sheet = spreadsheet.getSheetByName(config.sheetNames.participants);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(config.sheetNames.participants);
  }
  
  const headers = ['No', '氏名', 'メールアドレス', '所属', '役職', '備考', '有効'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#4a86e8')
    .setFontColor('white')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 250);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 200);
  sheet.setColumnWidth(7, 60);
  
  sheet.getRange(2, 1, 3, 7).setValues([
    [1, '山田 太郎', 'yamada@example.com', 'A社', '部長', '', '○'],
    [2, '佐藤 花子', 'sato@example.com', 'B社', '課長', '', '○'],
    [3, '鈴木 一郎', 'suzuki@example.com', 'C社', '主任', '', '○']
  ]);
  
  const validRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['○', '×'], true)
    .build();
  sheet.getRange('G2:G100').setDataValidation(validRule);
  
  sheet.setFrozenRows(1);
}

function setupCurrentTrainingSheet_(spreadsheet) {
  const config = getConfig();
  let sheet = spreadsheet.getSheetByName(config.sheetNames.currentTraining);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(config.sheetNames.currentTraining);
  }
  
  const formData = [
    ['研修会名', ''],
    ['開催日', ''],
    ['開催時間', ''],
    ['会場', ''],
    ['会場住所', ''],
    ['講師名', ''],
    ['研修内容', ''],
    ['', ''],
    ['出欠フォームURL', '（自動生成されます）'],
    ['フォームID', '（自動生成されます）']
  ];
  
  sheet.getRange(1, 1, formData.length, 2).setValues(formData);
  
  sheet.getRange('A1:A10')
    .setBackground('#e8f0fe')
    .setFontWeight('bold')
    .setHorizontalAlignment('right');
  
  sheet.getRange('B1:B10')
    .setBackground('#ffffff')
    .setBorder(true, true, true, true, false, false);
  
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 400);
  
  sheet.getRange('B2').setNumberFormat('yyyy/mm/dd');
}

function setupEmailTemplateSheet_(spreadsheet) {
  const config = getConfig();
  let sheet = spreadsheet.getSheetByName(config.sheetNames.emailTemplate);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(config.sheetNames.emailTemplate);
  }
  
  const templateData = [
    ['件名テンプレート', '【研修会のご案内】{{研修会名}}（{{開催日}}）'],
    ['', ''],
    ['本文テンプレート', ''],
    ['', '{{氏名}} 様'],
    ['', ''],
    ['', 'いつもお世話になっております。'],
    ['', '研修会事務局です。'],
    ['', ''],
    ['', '下記の研修会についてご案内申し上げます。'],
    ['', ''],
    ['', '━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    ['', '■ 研修会名：{{研修会名}}'],
    ['', '■ 開催日時：{{開催日}} {{開催時間}}'],
    ['', '■ 会場：{{会場}}'],
    ['', '■ 会場住所：{{会場住所}}'],
    ['', '■ 講師：{{講師名}}'],
    ['', '━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    ['', ''],
    ['', '【研修内容】'],
    ['', '{{研修内容}}'],
    ['', ''],
    ['', '【出欠確認のお願い】'],
    ['', '下記URLより、{{回答期限}}までに出欠をご回答ください。'],
    ['', '{{出欠フォームURL}}'],
    ['', ''],
    ['', '添付資料をご確認の上、ご参加ください。'],
    ['', ''],
    ['', '何かご不明な点がございましたら、お気軽にお問い合わせください。'],
    ['', ''],
    ['', '━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    ['', '研修会事務局'],
    ['', '━━━━━━━━━━━━━━━━━━━━━━━━━━']
  ];
  
  sheet.getRange(1, 1, templateData.length, 2).setValues(templateData);
  
  sheet.getRange('A1').setBackground('#fff2cc').setFontWeight('bold');
  sheet.getRange('A3').setBackground('#fff2cc').setFontWeight('bold');
  
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 500);
}

function setupHistorySheet_(spreadsheet) {
  const config = getConfig();
  let sheet = spreadsheet.getSheetByName(config.sheetNames.history);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(config.sheetNames.history);
  }
  
  const headers = ['送信日時', '研修会名', '開催日', '送信先', '送信者数', '添付ファイル', 'ステータス'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#93c47d')
    .setFontColor('white')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 300);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 200);
  sheet.setColumnWidth(7, 100);
  
  sheet.setFrozenRows(1);
}

function setupAttendanceSheet_(spreadsheet) {
  const config = getConfig();
  let sheet = spreadsheet.getSheetByName(config.sheetNames.attendance);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(config.sheetNames.attendance);
  }
  
  const headers = ['タイムスタンプ', '研修会名', '氏名', 'メールアドレス', '出欠', '備考'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#f6b26b')
    .setFontColor('white')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 250);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 200);
  
  sheet.setFrozenRows(1);
}

/**
 * 年間スケジュールシートをセットアップ（新規追加）
 * @private
 */
function setupYearlyScheduleSheet_(spreadsheet) {
  const config = getConfig();
  let sheet = spreadsheet.getSheetByName(config.sheetNames.yearlySchedule);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(config.sheetNames.yearlySchedule);
  }
  
  const headers = ['回', '研修会名', '開催日', '時間', '会場', '講師', 'ステータス', '備考'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#674ea7')
    .setFontColor('white')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // サンプルデータ
  const sampleData = [
    [1, '第1回 ○○研修', '', '14:00〜17:00', '', '', '予定', ''],
    [2, '第2回 ○○研修', '', '14:00〜17:00', '', '', '予定', ''],
    [3, '第3回 ○○研修', '', '14:00〜17:00', '', '', '予定', '']
  ];
  sheet.getRange(2, 1, sampleData.length, 8).setValues(sampleData);
  
  // 列幅設定
  sheet.setColumnWidth(1, 40);   // 回
  sheet.setColumnWidth(2, 200);  // 研修会名
  sheet.setColumnWidth(3, 100);  // 開催日
  sheet.setColumnWidth(4, 120);  // 時間
  sheet.setColumnWidth(5, 150);  // 会場
  sheet.setColumnWidth(6, 100);  // 講師
  sheet.setColumnWidth(7, 80);   // ステータス
  sheet.setColumnWidth(8, 150);  // 備考
  
  // ステータスのドロップダウン
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['予定', '案内済', '完了', '中止'], true)
    .build();
  sheet.getRange('G2:G50').setDataValidation(statusRule);
  
  // 日付フォーマット
  sheet.getRange('C2:C50').setNumberFormat('yyyy/mm/dd');
  
  sheet.setFrozenRows(1);
}

function saveSystemSettings_(spreadsheet, settings) {
  const config = getConfig();
  let settingsSheet = spreadsheet.getSheetByName(config.sheetNames.settings);
  if (!settingsSheet) {
    settingsSheet = spreadsheet.insertSheet(config.sheetNames.settings);
  }
  
  settingsSheet.clear();
  settingsSheet.getRange('A1:B1').setValues([['設定項目', '値']]);
  settingsSheet.getRange('A2:B4').setValues([
    ['システムフォルダID', settings.systemFolderId],
    ['添付資料フォルダID', settings.attachmentFolderId],
    ['スプレッドシートID', settings.spreadsheetId]
  ]);
  
  settingsSheet.getRange('A1:B1').setBackground('#4a86e8').setFontColor('white').setFontWeight('bold');
  settingsSheet.setColumnWidth(1, 200);
  settingsSheet.setColumnWidth(2, 400);
  
  settingsSheet.hideSheet();
}

function loadSystemSettings_(spreadsheet) {
  const config = getConfig();
  const settingsSheet = spreadsheet.getSheetByName(config.sheetNames.settings);
  
  if (!settingsSheet) {
    throw new Error('設定シートが見つかりません。setupSystem()を実行してください。');
  }
  
  const data = settingsSheet.getRange('A2:B4').getValues();
  return {
    systemFolderId: data[0][1],
    attachmentFolderId: data[1][1],
    spreadsheetId: data[2][1]
  };
}

function getTrainingInfo_(spreadsheet) {
  const config = getConfig();
  const sheet = spreadsheet.getSheetByName(config.sheetNames.currentTraining);
  const data = sheet.getRange('B1:B10').getValues();
  
  let dateStr = '';
  if (data[1][0]) {
    if (data[1][0] instanceof Date) {
      dateStr = Utilities.formatDate(data[1][0], 'JST', 'yyyy/MM/dd');
    } else {
      dateStr = data[1][0].toString();
    }
  }
  
  let deadlineStr = '';
  if (data[1][0]) {
    const deadline = new Date(data[1][0]);
    deadline.setDate(deadline.getDate() - config.attendanceDeadlineDays);
    deadlineStr = Utilities.formatDate(deadline, 'JST', 'yyyy/MM/dd');
  }
  
  return {
    name: data[0][0] || '',
    date: dateStr,
    time: data[2][0] || '',
    venue: data[3][0] || '',
    address: data[4][0] || '',
    instructor: data[5][0] || '',
    content: data[6][0] || '',
    formUrl: data[8][0] || '',
    formId: data[9][0] || '',
    deadline: deadlineStr
  };
}

function getActiveParticipants_(spreadsheet) {
  const config = getConfig();
  const sheet = spreadsheet.getSheetByName(config.sheetNames.participants);
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return [];
  
  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const participants = [];
  
  data.forEach(row => {
    if (row[6] === '○' && row[2]) {
      participants.push({
        name: row[1] || '',
        email: row[2],
        organization: row[3] || '',
        position: row[4] || '',
        note: row[5] || ''
      });
    }
  });
  
  return participants;
}

function getEmailTemplate_(spreadsheet) {
  const config = getConfig();
  const sheet = spreadsheet.getSheetByName(config.sheetNames.emailTemplate);
  const data = sheet.getRange('B1:B32').getValues();
  
  return {
    subject: data[0][0] || '',
    body: data.slice(3).map(row => row[0]).join('\n')
  };
}

function getAttachmentFiles_(spreadsheet) {
  try {
    const settings = loadSystemSettings_(spreadsheet);
    const folder = DriveApp.getFolderById(settings.attachmentFolderId);
    const files = folder.getFiles();
    const attachments = [];
    
    while (files.hasNext()) {
      const file = files.next();
      attachments.push(file.getAs(file.getMimeType()));
    }
    
    return attachments;
  } catch (e) {
    console.error('添付ファイル取得エラー: ' + e.message);
    return [];
  }
}

function personalizeTemplate_(template, trainingInfo, participant) {
  return template
    .replace(/{{氏名}}/g, participant.name)
    .replace(/{{メールアドレス}}/g, participant.email)
    .replace(/{{所属}}/g, participant.organization)
    .replace(/{{役職}}/g, participant.position)
    .replace(/{{研修会名}}/g, trainingInfo.name)
    .replace(/{{開催日}}/g, trainingInfo.date)
    .replace(/{{開催時間}}/g, trainingInfo.time)
    .replace(/{{会場}}/g, trainingInfo.venue)
    .replace(/{{会場住所}}/g, trainingInfo.address)
    .replace(/{{講師名}}/g, trainingInfo.instructor)
    .replace(/{{研修内容}}/g, trainingInfo.content)
    .replace(/{{出欠フォームURL}}/g, trainingInfo.formUrl)
    .replace(/{{回答期限}}/g, trainingInfo.deadline);
}

function recordHistory_(spreadsheet, trainingInfo, participants, attachments, successCount, failedEmails) {
  const config = getConfig();
  const sheet = spreadsheet.getSheetByName(config.sheetNames.history);
  const now = new Date();
  
  const emails = participants.map(p => p.email).join(', ');
  const attachmentNames = attachments.length > 0 
    ? attachments.map(a => a.getName()).join(', ')
    : 'なし';
  const status = failedEmails.length === 0 ? '✅ 成功' : `⚠️ ${failedEmails.length}件失敗`;
  
  sheet.appendRow([
    now,
    trainingInfo.name,
    trainingInfo.date,
    emails,
    successCount,
    attachmentNames,
    status
  ]);
}

function formatDate_(date) {
  if (date instanceof Date) {
    return Utilities.formatDate(date, 'JST', 'yyyy/MM/dd');
  }
  return date ? date.toString() : '';
}
