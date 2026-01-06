/**
 * =========================================================
 * 研修会管理システム v3.0（ライブラリ利用版）
 * =========================================================
 * 
 * 【v3.0 新機能】
 * - 未回答者へのリマインダー送信
 * - 年間スケジュール管理
 * - 出欠一覧のPDF出力
 * - 自動送信トリガー
 */

// ============================================================
// 設定（自分の環境に合わせて変更）
// ============================================================
const CONFIG = {
  // 親フォルダID（ファイルを作成するGoogle DriveフォルダのID）
  parentFolderId: 'ここにフォルダIDを入力',
  
  // システムフォルダ名
  systemFolderName: '研修会管理システム',
  
  // 添付資料フォルダ名
  attachmentFolderName: '添付資料',
  
  // スプレッドシート名
  spreadsheetName: '研修会管理システム',
  
  // 送信者名（メールの差出人として表示）
  senderName: '研修会事務局',
  
  // 出欠回答期限（研修日の何日前まで）
  attendanceDeadlineDays: 3,
  
  // 【新設定】自動送信（研修日の何日前に案内を送信）
  autoSendDaysBefore: 7,
  
  // 【新設定】自動リマインダー（研修日の何日前にリマインダー送信）
  reminderDaysBefore: 2
};


// ============================================================
// 初期化
// ============================================================
function initLibrary_() {
  TrainingSystemLib.init(CONFIG);
}

function isUiAvailable_() {
  try {
    SpreadsheetApp.getUi();
    return true;
  } catch (e) {
    return false;
  }
}


// ============================================================
// メニュー追加（拡張版）
// ============================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📧 研修会管理')
    // 基本機能
    .addItem('📝 出欠確認フォームを作成', 'createAttendanceForm')
    .addSeparator()
    .addItem('📨 案内メールを送信', 'sendTrainingNotification')
    .addItem('📨 テスト送信（自分のみ）', 'sendTestEmail')
    .addSeparator()
    // 出欠管理
    .addItem('📊 出欠状況を確認', 'showAttendanceStatus')
    .addItem('🔔 未回答者にリマインダー送信', 'sendReminderToNoResponse')
    .addItem('📄 出欠一覧をPDF出力', 'exportAttendancePdf')
    .addSeparator()
    // 年間スケジュール
    .addSubMenu(ui.createMenu('📅 年間スケジュール')
      .addItem('次回研修会を当日シートにコピー', 'copyNextTrainingToCurrentSheet')
      .addItem('年間スケジュール一覧を表示', 'showYearlySchedule'))
    .addSeparator()
    // 自動化設定
    .addSubMenu(ui.createMenu('⚙️ 自動化設定')
      .addItem('🟢 自動送信を有効化（毎朝9時）', 'enableAutoSend')
      .addItem('🔴 自動送信を無効化', 'disableAutoSend')
      .addItem('📋 トリガー状態を確認', 'checkTriggerStatus'))
    .addSeparator()
    // その他
    .addItem('🗑️ 当日研修会をクリア', 'clearCurrentTraining')
    .addItem('📁 添付資料フォルダを開く', 'openAttachmentFolder')
    .addToUi();
}


// ============================================================
// セットアップ（初回のみ実行）
// ============================================================
function setupSystem() {
  initLibrary_();
  const result = TrainingSystemLib.setupSystem();
  
  console.log('🎉 セットアップ完了！');
  console.log('📊 スプレッドシートURL: ' + result.spreadsheetUrl);
  console.log('📁 添付資料フォルダURL: ' + result.attachmentFolderUrl);
  
  if (isUiAvailable_()) {
    SpreadsheetApp.getUi().alert('セットアップ完了', 
      `スプレッドシートURL:\n${result.spreadsheetUrl}\n\n` +
      `添付資料フォルダURL:\n${result.attachmentFolderUrl}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
  
  return result;
}


// ============================================================
// 出欠確認フォーム作成
// ============================================================
function createAttendanceForm() {
  initLibrary_();
  
  try {
    const trainingInfo = TrainingSystemLib.getTrainingInfo();
    
    if (!trainingInfo.name || !trainingInfo.date) {
      const msg = '「当日研修会」シートに研修会名と開催日を入力してください。';
      console.error(msg);
      if (isUiAvailable_()) {
        SpreadsheetApp.getUi().alert('エラー', msg, SpreadsheetApp.getUi().ButtonSet.OK);
      }
      return;
    }
    
    if (isUiAvailable_()) {
      const ui = SpreadsheetApp.getUi();
      const confirm = ui.alert(
        '出欠フォーム作成',
        `以下の研修会の出欠フォームを作成しますか？\n\n研修会名: ${trainingInfo.name}\n開催日: ${trainingInfo.date}`,
        ui.ButtonSet.YES_NO
      );
      if (confirm !== ui.Button.YES) return;
    }
    
    const result = TrainingSystemLib.createAttendanceForm();
    
    console.log('✅ フォーム作成完了');
    console.log('フォームURL: ' + result.formUrl);
    
    if (isUiAvailable_()) {
      SpreadsheetApp.getUi().alert('✅ フォーム作成完了',
        `出欠確認フォームを作成しました。\n\n` +
        `フォームURL:\n${result.formUrl}\n\n` +
        `回答期限: ${result.deadline}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
    return result;
    
  } catch (error) {
    console.error('エラー: ' + error.message);
    if (isUiAvailable_()) {
      SpreadsheetApp.getUi().alert('エラー', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}


// ============================================================
// メール送信
// ============================================================
function sendTrainingNotification() {
  initLibrary_();
  
  try {
    const trainingInfo = TrainingSystemLib.getTrainingInfo();
    const participants = TrainingSystemLib.getActiveParticipants();
    
    if (!trainingInfo.name || !trainingInfo.date) {
      const msg = '「当日研修会」シートに研修会名と開催日を入力してください。';
      if (isUiAvailable_()) {
        SpreadsheetApp.getUi().alert('エラー', msg, SpreadsheetApp.getUi().ButtonSet.OK);
      }
      return;
    }
    
    if (participants.length === 0) {
      if (isUiAvailable_()) {
        SpreadsheetApp.getUi().alert('エラー', '有効な参加者がいません。', SpreadsheetApp.getUi().ButtonSet.OK);
      }
      return;
    }
    
    if (isUiAvailable_()) {
      const ui = SpreadsheetApp.getUi();
      const confirm = ui.alert('送信確認',
        `以下の内容でメールを送信しますか？\n\n` +
        `研修会名: ${trainingInfo.name}\n` +
        `開催日: ${trainingInfo.date}\n` +
        `送信先: ${participants.length}名`,
        ui.ButtonSet.YES_NO
      );
      if (confirm !== ui.Button.YES) return;
    }
    
    const result = TrainingSystemLib.sendNotification({ testMode: false });
    
    let message = `✅ ${result.successCount}/${result.totalCount}件 送信完了`;
    if (result.failedEmails.length > 0) {
      message += `\n\n⚠️ 送信失敗:\n${result.failedEmails.join('\n')}`;
    }
    
    console.log(message);
    
    if (isUiAvailable_()) {
      SpreadsheetApp.getUi().alert('送信結果', message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
    return result;
    
  } catch (error) {
    console.error('エラー: ' + error.message);
    if (isUiAvailable_()) {
      SpreadsheetApp.getUi().alert('エラー', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}


// ============================================================
// テスト送信
// ============================================================
function sendTestEmail() {
  initLibrary_();
  
  try {
    const trainingInfo = TrainingSystemLib.getTrainingInfo();
    
    if (!trainingInfo.name || !trainingInfo.date) {
      const msg = '「当日研修会」シートに研修会名と開催日を入力してください。';
      if (isUiAvailable_()) {
        SpreadsheetApp.getUi().alert('エラー', msg, SpreadsheetApp.getUi().ButtonSet.OK);
      }
      return;
    }
    
    if (isUiAvailable_()) {
      const ui = SpreadsheetApp.getUi();
      const confirm = ui.alert('テスト送信',
        '自分のメールアドレスにテスト送信しますか？',
        ui.ButtonSet.YES_NO
      );
      if (confirm !== ui.Button.YES) return;
    }
    
    const result = TrainingSystemLib.sendNotification({ testMode: true });
    
    const msg = `✅ ${Session.getActiveUser().getEmail()} に送信しました。`;
    console.log(msg);
    
    if (isUiAvailable_()) {
      SpreadsheetApp.getUi().alert('✅ テスト送信完了', msg + '\n受信トレイを確認してください。', SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
    return result;
    
  } catch (error) {
    console.error('エラー: ' + error.message);
    if (isUiAvailable_()) {
      SpreadsheetApp.getUi().alert('エラー', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}


// ============================================================
// 出欠状況確認
// ============================================================
function showAttendanceStatus() {
  initLibrary_();
  
  try {
    const status = TrainingSystemLib.getAttendanceStatus();
    
    const message = 
      `📊 出欠状況\n\n` +
      `参加者総数: ${status.totalParticipants}名\n` +
      `回答者数: ${status.responseCount}名\n` +
      `未回答: ${status.noResponseCount}名\n\n` +
      `━━━━━━━━━━━━━━\n` +
      `出席: ${status.attendCount}名\n` +
      `欠席: ${status.absentCount}名\n` +
      `未定: ${status.undecidedCount}名`;
    
    console.log(message);
    
    if (isUiAvailable_()) {
      SpreadsheetApp.getUi().alert('出欠状況', message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
    return status;
    
  } catch (error) {
    console.error('エラー: ' + error.message);
    if (isUiAvailable_()) {
      SpreadsheetApp.getUi().alert('エラー', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}


// ============================================================
// 【新機能】未回答者リマインダー送信
// ============================================================
function sendReminderToNoResponse() {
  initLibrary_();
  
  try {
    // 未回答者を確認
    const noResponseList = TrainingSystemLib.getNoResponseParticipants();
    
    if (noResponseList.length === 0) {
      const msg = '未回答者はいません。全員回答済みです！';
      console.log(msg);
      if (isUiAvailable_()) {
        SpreadsheetApp.getUi().alert('✅ 確認', msg, SpreadsheetApp.getUi().ButtonSet.OK);
      }
      return;
    }
    
    if (isUiAvailable_()) {
      const ui = SpreadsheetApp.getUi();
      const nameList = noResponseList.map(p => p.name).join('\n');
      const confirm = ui.alert('リマインダー送信',
        `以下の${noResponseList.length}名にリマインダーを送信しますか？\n\n${nameList}`,
        ui.ButtonSet.YES_NO
      );
      if (confirm !== ui.Button.YES) return;
    }
    
    const result = TrainingSystemLib.sendReminder({ testMode: false });
    
    let message = `✅ ${result.successCount}/${result.totalCount}件 リマインダー送信完了`;
    if (result.failedEmails && result.failedEmails.length > 0) {
      message += `\n\n⚠️ 送信失敗:\n${result.failedEmails.join('\n')}`;
    }
    
    console.log(message);
    
    if (isUiAvailable_()) {
      SpreadsheetApp.getUi().alert('送信結果', message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
    return result;
    
  } catch (error) {
    console.error('エラー: ' + error.message);
    if (isUiAvailable_()) {
      SpreadsheetApp.getUi().alert('エラー', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}

/**
 * リマインダーテスト送信
 */
function sendReminderTest() {
  initLibrary_();
  
  try {
    if (isUiAvailable_()) {
      const ui = SpreadsheetApp.getUi();
      const confirm = ui.alert('リマインダーテスト送信',
        '自分のメールアドレスにリマインダーをテスト送信しますか？',
        ui.ButtonSet.YES_NO
      );
      if (confirm !== ui.Button.YES) return;
    }
    
    const result = TrainingSystemLib.sendReminder({ testMode: true });
    
    const msg = `✅ ${Session.getActiveUser().getEmail()} にリマインダーを送信しました。`;
    console.log(msg);
    
    if (isUiAvailable_()) {
      SpreadsheetApp.getUi().alert('✅ テスト送信完了', msg, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
    return result;
    
  } catch (error) {
    console.error('エラー: ' + error.message);
    if (isUiAvailable_()) {
      SpreadsheetApp.getUi().alert('エラー', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}


// ============================================================
// 【新機能】出欠一覧PDF出力
// ============================================================
function exportAttendancePdf() {
  initLibrary_();
  
  try {
    const trainingInfo = TrainingSystemLib.getTrainingInfo();
    
    if (!trainingInfo.name) {
      const msg = '「当日研修会」シートに研修会情報を入力してください。';
      if (isUiAvailable_()) {
        SpreadsheetApp.getUi().alert('エラー', msg, SpreadsheetApp.getUi().ButtonSet.OK);
      }
      return;
    }
    
    if (isUiAvailable_()) {
      const ui = SpreadsheetApp.getUi();
      const confirm = ui.alert('PDF出力',
        `「${trainingInfo.name}」の出欠一覧をPDF出力しますか？`,
        ui.ButtonSet.YES_NO
      );
      if (confirm !== ui.Button.YES) return;
    }
    
    const result = TrainingSystemLib.exportAttendanceToPdf();
    
    const message = 
      `✅ PDF出力完了\n\n` +
      `ファイル名: ${result.fileName}\n\n` +
      `【集計】\n` +
      `参加者: ${result.summary.total}名\n` +
      `出席: ${result.summary.attend}名\n` +
      `欠席: ${result.summary.absent}名\n` +
      `未定: ${result.summary.undecided}名\n` +
      `未回答: ${result.summary.noResponse}名`;
    
    console.log(message);
    console.log('PDF URL: ' + result.fileUrl);
    
    if (isUiAvailable_()) {
      const ui = SpreadsheetApp.getUi();
      ui.alert('✅ PDF出力完了', message, ui.ButtonSet.OK);
      
      // PDFを開く
      const html = HtmlService.createHtmlOutput(
        `<script>window.open('${result.fileUrl}', '_blank'); google.script.host.close();</script>`
      ).setWidth(1).setHeight(1);
      ui.showModalDialog(html, 'PDFを開いています...');
    }
    
    return result;
    
  } catch (error) {
    console.error('エラー: ' + error.message);
    if (isUiAvailable_()) {
      SpreadsheetApp.getUi().alert('エラー', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}


// ============================================================
// 【新機能】年間スケジュール管理
// ============================================================

/**
 * 年間スケジュール一覧を表示
 */
function showYearlySchedule() {
  initLibrary_();
  
  try {
    const schedule = TrainingSystemLib.getYearlySchedule();
    
    if (schedule.length === 0) {
      const msg = '年間スケジュールが登録されていません。';
      if (isUiAvailable_()) {
        SpreadsheetApp.getUi().alert('情報', msg, SpreadsheetApp.getUi().ButtonSet.OK);
      }
      return;
    }
    
    let message = '📅 年間スケジュール\n\n';
    schedule.forEach(s => {
      const statusIcon = s.status === '完了' ? '✅' : s.status === '案内済' ? '📧' : s.status === '中止' ? '❌' : '📅';
      message += `${statusIcon} 第${s.number}回 ${s.name}\n`;
      message += `   ${s.date} ${s.time} @ ${s.venue}\n`;
      message += `   ステータス: ${s.status}\n\n`;
    });
    
    console.log(message);
    
    if (isUiAvailable_()) {
      SpreadsheetApp.getUi().alert('年間スケジュール', message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
    return schedule;
    
  } catch (error) {
    console.error('エラー: ' + error.message);
    if (isUiAvailable_()) {
      SpreadsheetApp.getUi().alert('エラー', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}

/**
 * 次回研修会を当日シートにコピー
 */
function copyNextTrainingToCurrentSheet() {
  initLibrary_();
  
  try {
    const next = TrainingSystemLib.getNextTraining();
    
    if (!next) {
      const msg = '今後の研修会がありません。\n年間スケジュールシートを確認してください。';
      if (isUiAvailable_()) {
        SpreadsheetApp.getUi().alert('情報', msg, SpreadsheetApp.getUi().ButtonSet.OK);
      }
      return;
    }
    
    if (isUiAvailable_()) {
      const ui = SpreadsheetApp.getUi();
      const confirm = ui.alert('次回研修会をコピー',
        `以下の研修会を「当日研修会」シートにコピーしますか？\n\n` +
        `第${next.number}回 ${next.name}\n` +
        `開催日: ${next.date}\n` +
        `会場: ${next.venue}`,
        ui.ButtonSet.YES_NO
      );
      if (confirm !== ui.Button.YES) return;
    }
    
    TrainingSystemLib.copyScheduleToCurrentTraining(next.rowIndex);
    
    const msg = `✅ 「${next.name}」を当日研修会シートにコピーしました。`;
    console.log(msg);
    
    if (isUiAvailable_()) {
      SpreadsheetApp.getUi().alert('完了', msg, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
  } catch (error) {
    console.error('エラー: ' + error.message);
    if (isUiAvailable_()) {
      SpreadsheetApp.getUi().alert('エラー', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}


// ============================================================
// 【新機能】自動送信トリガー
// ============================================================

/**
 * 日次自動処理（トリガーから呼び出される）
 * ※この関数名は変更しないでください
 */
function dailyAutoProcess() {
  initLibrary_();
  
  console.log('🤖 日次自動処理を開始...');
  
  // 1. 自動案内送信チェック
  const sendResult = TrainingSystemLib.checkAndAutoSend();
  console.log('案内送信チェック結果:', JSON.stringify(sendResult));
  
  // 2. 自動リマインダーチェック
  const remindResult = TrainingSystemLib.checkAndAutoRemind();
  console.log('リマインダーチェック結果:', JSON.stringify(remindResult));
  
  console.log('🤖 日次自動処理完了');
  
  return { sendResult, remindResult };
}

/**
 * 自動送信を有効化
 */
function enableAutoSend() {
  initLibrary_();
  
  try {
    if (isUiAvailable_()) {
      const ui = SpreadsheetApp.getUi();
      const confirm = ui.alert('自動送信を有効化',
        `毎朝9時に以下を自動実行します：\n\n` +
        `1. 年間スケジュールをチェック\n` +
        `2. ${CONFIG.autoSendDaysBefore}日後に研修会があれば案内メール送信\n` +
        `3. ${CONFIG.reminderDaysBefore}日後に研修会があれば未回答者にリマインダー送信\n\n` +
        `有効化しますか？`,
        ui.ButtonSet.YES_NO
      );
      if (confirm !== ui.Button.YES) return;
    }
    
    const result = TrainingSystemLib.setupDailyTrigger(9);  // 毎朝9時
    
    const msg = `✅ 自動送信を有効化しました（毎朝${result.hour}時に実行）`;
    console.log(msg);
    
    if (isUiAvailable_()) {
      SpreadsheetApp.getUi().alert('完了', msg, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
    return result;
    
  } catch (error) {
    console.error('エラー: ' + error.message);
    if (isUiAvailable_()) {
      SpreadsheetApp.getUi().alert('エラー', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}

/**
 * 自動送信を無効化
 */
function disableAutoSend() {
  initLibrary_();
  
  try {
    if (isUiAvailable_()) {
      const ui = SpreadsheetApp.getUi();
      const confirm = ui.alert('自動送信を無効化',
        '自動送信を無効化しますか？',
        ui.ButtonSet.YES_NO
      );
      if (confirm !== ui.Button.YES) return;
    }
    
    const result = TrainingSystemLib.removeDailyTrigger();
    
    const msg = `✅ 自動送信を無効化しました（${result.removed}件のトリガーを削除）`;
    console.log(msg);
    
    if (isUiAvailable_()) {
      SpreadsheetApp.getUi().alert('完了', msg, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
    return result;
    
  } catch (error) {
    console.error('エラー: ' + error.message);
    if (isUiAvailable_()) {
      SpreadsheetApp.getUi().alert('エラー', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}

/**
 * トリガー状態を確認
 */
function checkTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  let dailyTrigger = null;
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'dailyAutoProcess') {
      dailyTrigger = trigger;
    }
  });
  
  let message;
  if (dailyTrigger) {
    message = `🟢 自動送信: 有効\n\n` +
      `実行関数: ${dailyTrigger.getHandlerFunction()}\n` +
      `種類: ${dailyTrigger.getEventType()}`;
  } else {
    message = `🔴 自動送信: 無効\n\n` +
      `自動送信を有効にするには、メニューから\n` +
      `「⚙️ 自動化設定」→「🟢 自動送信を有効化」\n` +
      `を選択してください。`;
  }
  
  console.log(message);
  
  if (isUiAvailable_()) {
    SpreadsheetApp.getUi().alert('トリガー状態', message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
  
  return { enabled: !!dailyTrigger };
}


// ============================================================
// その他
// ============================================================

/**
 * 当日研修会をクリア
 */
function clearCurrentTraining() {
  initLibrary_();
  
  if (isUiAvailable_()) {
    const ui = SpreadsheetApp.getUi();
    const confirm = ui.alert('確認',
      '「当日研修会」シートの内容をクリアしますか？',
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) return;
  }
  
  TrainingSystemLib.clearCurrentTraining();
  
  console.log('✅ クリア完了');
  
  if (isUiAvailable_()) {
    SpreadsheetApp.getUi().alert('✅ クリア完了', '次回の研修会情報を入力してください。', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 添付資料フォルダを開く
 */
function openAttachmentFolder() {
  initLibrary_();
  
  try {
    const url = TrainingSystemLib.getAttachmentFolderUrl();
    
    console.log('添付資料フォルダURL: ' + url);
    
    if (isUiAvailable_()) {
      const html = HtmlService.createHtmlOutput(
        `<script>window.open('${url}', '_blank'); google.script.host.close();</script>`
      ).setWidth(1).setHeight(1);
      
      SpreadsheetApp.getUi().showModalDialog(html, '添付資料フォルダを開いています...');
    }
    
    return url;
    
  } catch (error) {
    console.error('エラー: ' + error.message);
    if (isUiAvailable_()) {
      SpreadsheetApp.getUi().alert('エラー', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}
