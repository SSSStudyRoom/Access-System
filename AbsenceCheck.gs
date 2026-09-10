// ============================================================
// SSS Education 自習室管理システム
// 欠席確認メール 自動送信モジュール
// ============================================================
// 【前提】
//   - 公開プロフィールシートに「カレンダーID」「メールアドレス」「保護者メール」列を追加済み
//   - 送信済みログシートが存在する（なければ自動作成）
//   - このスクリプトを既存GASプロジェクトに追加してトリガー設定する
// ============================================================

// ========== 設定値（環境に合わせて変更） ==========
const CONFIG = {
  SPREADSHEET_ID: '1HWZDOIJaQB0S3K4a9hXomJZFqmdznmKywQ2Sc_ON1Fo',        // スプレッドシートのID
  SHEET_PROFILE:  '公開プロフィール',               // 生徒マスタシート名
  SHEET_LOG:      '管理シート',                    // 入退室ログシート名
  SHEET_SENT:     '送信済みログ',                  // 送信済みログシート名（新設）
  ABSENCE_MINUTES: 5,                             // 予定開始から何分後に欠席判定するか

  // 公開プロフィールシートのカラム番号（1始まり）
  COL_ID:           1,   // 生徒ID
  COL_NAME:         4,   // 実名
  COL_EMAIL:        2,   // 生徒メールアドレス（★新規追加列）
  COL_PARENT_EMAIL: 3,   // 保護者メールアドレス（★新規追加列、空欄でもOK）
  COL_CALENDAR_ID:  11,   // GoogleカレンダーID（★新規追加列）

  // 管理シートのカラム番号（1始まり）
  LOG_COL_DATE:     1,   // 日付
  LOG_COL_ID:       2,   // 生徒ID
  LOG_COL_CHECKIN:  4,   // 入室時刻
};

// ============================================================
// メイン関数 — 時間トリガーで5分ごとに実行
// ============================================================
function checkAbsenceAndSendMail() {
  const ss     = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const today  = new Date();
  const todayStr = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd');

  // 各シートを取得
  const profileSheet = ss.getSheetByName(CONFIG.SHEET_PROFILE);
  const logSheet     = ss.getSheetByName(CONFIG.SHEET_LOG);
  const sentSheet    = getOrCreateSentSheet(ss);

  // 生徒マスタを全取得（1行目はヘッダーなのでスキップ）
  const profiles = profileSheet.getDataRange().getValues().slice(1);

  // 本日の入室済み生徒IDセットを作成
  const checkedInIds = getCheckedInIds(logSheet, todayStr);

  // 本日の送信済みキーセットを作成（重複送信防止）
  const sentKeys = getSentKeys(sentSheet, todayStr);

  // 各生徒のカレンダーをチェック
  profiles.forEach(row => {
    const studentId   = String(row[CONFIG.COL_ID - 1]).trim();
    const studentName = String(row[CONFIG.COL_NAME - 1]).trim();
    const studentMail = String(row[CONFIG.COL_EMAIL - 1]).trim();
    const parentMail  = String(row[CONFIG.COL_PARENT_EMAIL - 1]).trim();
    const calendarId  = String(row[CONFIG.COL_CALENDAR_ID - 1]).trim();

    // カレンダーIDまたはメールが未設定の生徒はスキップ
    if (!calendarId || calendarId === '' || !studentMail || studentMail === '') return;

    // 生徒のカレンダーから本日の予定を取得
    let calendar;
    try {
      calendar = CalendarApp.getCalendarById(calendarId);
    } catch(e) {
      console.log(`カレンダー取得失敗: ${studentName} (${calendarId})`);
      return;
    }
    if (!calendar) return;

    const startOfDay = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 0, 0, 0);
    const endOfDay   = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59);
    const events     = calendar.getEvents(startOfDay, endOfDay);

    events.forEach(event => {
      const eventStart = event.getStartTime();
      const minutesElapsed = (today - eventStart) / 1000 / 60;

      // 予定開始からCONFIG.ABSENCE_MINUTES分以上経過しているか
      if (minutesElapsed < CONFIG.ABSENCE_MINUTES) return;

      // 送信済みキー（生徒ID + 予定開始時刻）で重複チェック
      const eventStartStr = Utilities.formatDate(eventStart, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
      const sentKey = `${studentId}_${eventStartStr}`;
      if (sentKeys.has(sentKey)) return;

      // 入室済みであればスキップ
      if (checkedInIds.has(studentId)) return;

      // ===== 欠席判定 → メール送信 =====
      const recipients = buildRecipients(studentMail, parentMail);
      const subject    = buildSubject(studentName, eventStart);
      const body       = buildBody(studentName, eventStart);

      try {
        GmailApp.sendEmail(recipients, subject, body);
        logSent(sentSheet, todayStr, sentKey, studentId, studentName, eventStartStr, recipients);
        console.log(`送信完了: ${studentName} (${eventStartStr})`);
      } catch(e) {
        console.error(`送信失敗: ${studentName} — ${e.message}`);
      }
    });
  });
}

// ============================================================
// ヘルパー関数群
// ============================================================

/** 本日入室済みの生徒IDセットを返す */
function getCheckedInIds(logSheet, todayStr) {
  const ids = new Set();
  const data = logSheet.getDataRange().getValues();
  data.slice(1).forEach(row => {
    const rowDate = Utilities.formatDate(new Date(row[CONFIG.LOG_COL_DATE - 1]), 'Asia/Tokyo', 'yyyy/MM/dd');
    const checkin = row[CONFIG.LOG_COL_CHECKIN - 1];
    if (rowDate === todayStr && checkin !== '') {
      ids.add(String(row[CONFIG.LOG_COL_ID - 1]).trim());
    }
  });
  return ids;
}

/** 本日送信済みのキーセットを返す */
function getSentKeys(sentSheet, todayStr) {
  const keys = new Set();
  const data = sentSheet.getDataRange().getValues();
  data.slice(1).forEach(row => {
    const rowDateStr = Utilities.formatDate(row[0],'Asia/Tokyo','yyyy/MM/dd');
    const checkin = row[CONFIG.LOG_COL_CHECKIN - 1];
    if (rowDateStr === todayStr && checkin !== '') keys.add(String(row[1]));
  });
  return keys;
}

/** 送信先メールアドレスを組み立てる（カンマ区切り） */
function buildRecipients(studentMail, parentMail) {
  const list = [studentMail];
  if (parentMail && parentMail !== '' && parentMail !== 'undefined') {
    list.push(parentMail);
  }
  return list.join(',');
}

/** メール件名を組み立てる */
function buildSubject(studentName, eventStart) {
  const timeStr = Utilities.formatDate(eventStart, 'Asia/Tokyo', 'HH:mm');
  return `【自習室】${studentName}さんの来室が確認できません（${timeStr}〜）`;
}

/** メール本文を組み立てる */
function buildBody(studentName, eventStart) {
  const timeStr = Utilities.formatDate(eventStart, 'Asia/Tokyo', 'HH:mm');
  return `${studentName} さん（および保護者の方）

本日 ${timeStr} からの自習室利用予定になっておりますが、
まだ来室が確認できておりません。

欠席の場合はご連絡いただけますと幸いです。
体調など問題ないことをお祈りしております。

──────────────────────
SSS Education
──────────────────────`;
}

/** 送信済みログシートを取得（なければ新規作成） */
function getOrCreateSentSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_SENT);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_SENT);
    sheet.appendRow(['日付', '送信キー', '生徒ID', '生徒名', '予定開始', '送信先']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** 送信済みログに1行追記 */
function logSent(sentSheet, todayStr, sentKey, studentId, studentName, eventStartStr, recipients) {
  sentSheet.appendRow([todayStr, sentKey, studentId, studentName, eventStartStr, recipients]);
}

// ============================================================
// 初回セットアップ用：時間トリガーを自動登録する関数
// GASエディタで一度だけ手動実行してください
// ============================================================
function setupTrigger() {
  // 既存の同名トリガーを削除
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'checkAbsenceAndSendMail') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // 5分ごとのトリガーを新規登録
  ScriptApp.newTrigger('checkAbsenceAndSendMail')
    .timeBased()
    .everyMinutes(5)
    .create();
  console.log('トリガーを登録しました（5分ごと）');
}

// ============================================================
// SSS Education 自習室管理システム
// カレンダー未入力 催促メール 自動送信モジュール
// ============================================================
// 【前提】
//   - absence_check.gs と同じGASプロジェクト内に追加
//   - 公開プロフィールシートに「カレンダーID」「メールアドレス」列が設定済み
//   - 月曜日の朝（例：8:00）に時間トリガーで実行
// ============================================================

// ========== 設定値（absence_check.gsのCONFIGと共通化推奨） ==========
const REMINDER_CONFIG = {
  SPREADSHEET_ID:   '1HWZDOIJaQB0S3K4a9hXomJZFqmdznmKywQ2Sc_ON1Fo',
  SHEET_PROFILE:    '公開プロフィール',
  SHEET_SENT:       '催促送信済みログ',       // 催促用の送信済みログ（別シート）

  COL_ID:           1,
  COL_NAME:         4,
  COL_EMAIL:        2,
  COL_PARENT_EMAIL: 3,
  COL_CALENDAR_ID:  11,

  // VERSION_B用：直近何日間に予定がなければ送るか
  VERSION_B_DAYS_THIS_WEEK: 7,   // 今週（7日間）
  VERSION_B_DAYS_NEXT_WEEK: 14,  // 翌週（14日間）
};

// ============================================================
// Version A：翌週分（月〜日）が1件も入力されていない生徒に送る
// ============================================================
function sendReminderVersionA() {
  const ss = SpreadsheetApp.openById(REMINDER_CONFIG.SPREADSHEET_ID);
  const profiles  = ss.getSheetByName(REMINDER_CONFIG.SHEET_PROFILE)
                      .getDataRange().getValues().slice(1);
  const sentSheet = getOrCreateReminderSentSheet(ss);

  // 翌週の月曜〜日曜を算出
  const today     = new Date();
  const nextMonday = getNextMonday(today);
  const nextSunday = new Date(nextMonday);
  nextSunday.setDate(nextMonday.getDate() + 6);
  nextSunday.setHours(23, 59, 59);

  const weekKey = Utilities.formatDate(nextMonday, 'Asia/Tokyo', 'yyyy/MM/dd'); // 重複防止キー
  const sentKeys = getReminderSentKeys(sentSheet);

  profiles.forEach(row => {
    const studentId  = String(row[REMINDER_CONFIG.COL_ID - 1]).trim();
    const name       = String(row[REMINDER_CONFIG.COL_NAME - 1]).trim();
    const email      = String(row[REMINDER_CONFIG.COL_EMAIL - 1]).trim();
    const parentMail = String(row[REMINDER_CONFIG.COL_PARENT_EMAIL - 1]).trim();
    const calId      = String(row[REMINDER_CONFIG.COL_CALENDAR_ID - 1]).trim();

    if (!calId || !email) return;

    const sentKey = `A_${studentId}_${weekKey}`;
    if (sentKeys.has(sentKey)) return; // 送信済みはスキップ

    let calendar;
    try { calendar = CalendarApp.getCalendarById(calId); } catch(e) { return; }
    if (!calendar) return;

    const events = calendar.getEvents(nextMonday, nextSunday);

    if (events.length === 0) {
      // 翌週に予定が1件もない → 催促メール送信
      const recipients = buildReminderRecipients(email, parentMail);
      const subject    = buildReminderSubjectA(name, nextMonday, nextSunday);
      const body       = buildReminderBodyA(name, nextMonday, nextSunday);
      try {
        GmailApp.sendEmail(recipients, subject, body);
        logReminderSent(sentSheet, sentKey, studentId, name, 'A', weekKey, recipients);
        console.log(`[Version A] 送信: ${name}`);
      } catch(e) {
        console.error(`[Version A] 送信失敗: ${name} — ${e.message}`);
      }
    }
  });
}

// ============================================================
// Version B：今週 OR 翌週のどちらかが空の生徒に送る
// ============================================================
function sendReminderVersionB() {
  const ss = SpreadsheetApp.openById(REMINDER_CONFIG.SPREADSHEET_ID);
  const profiles  = ss.getSheetByName(REMINDER_CONFIG.SHEET_PROFILE)
                      .getDataRange().getValues().slice(1);
  const sentSheet = getOrCreateReminderSentSheet(ss);

  const today      = new Date();
  const thisMonday = getThisMonday(today);
  const thisSunday = new Date(thisMonday);
  thisSunday.setDate(thisMonday.getDate() + 6);
  thisSunday.setHours(23, 59, 59);

  const nextMonday = new Date(thisMonday);
  nextMonday.setDate(thisMonday.getDate() + 7);
  const nextSunday = new Date(nextMonday);
  nextSunday.setDate(nextMonday.getDate() + 6);
  nextSunday.setHours(23, 59, 59);

  const weekKey  = Utilities.formatDate(thisMonday, 'Asia/Tokyo', 'yyyy/MM/dd');
  const sentKeys = getReminderSentKeys(sentSheet);

  profiles.forEach(row => {
    const studentId  = String(row[REMINDER_CONFIG.COL_ID - 1]).trim();
    const name       = String(row[REMINDER_CONFIG.COL_NAME - 1]).trim();
    const email      = String(row[REMINDER_CONFIG.COL_EMAIL - 1]).trim();
    const parentMail = String(row[REMINDER_CONFIG.COL_PARENT_EMAIL - 1]).trim();
    const calId      = String(row[REMINDER_CONFIG.COL_CALENDAR_ID - 1]).trim();

    if (!calId || !email) return;

    const sentKey = `B_${studentId}_${weekKey}`;
    if (sentKeys.has(sentKey)) return;

    let calendar;
    try { calendar = CalendarApp.getCalendarById(calId); } catch(e) { return; }
    if (!calendar) return;

    const thisWeekEvents = calendar.getEvents(thisMonday, thisSunday);
    const nextWeekEvents = calendar.getEvents(nextMonday, nextSunday);

    const thisWeekEmpty = thisWeekEvents.length === 0;
    const nextWeekEmpty = nextWeekEvents.length === 0;

    if (thisWeekEmpty || nextWeekEmpty) {
      const recipients = buildReminderRecipients(email, parentMail);
      const subject    = buildReminderSubjectB(name, thisWeekEmpty, nextWeekEmpty);
      const body       = buildReminderBodyB(name, thisMonday, thisSunday, nextMonday, nextSunday, thisWeekEmpty, nextWeekEmpty);
      try {
        GmailApp.sendEmail(recipients, subject, body);
        logReminderSent(sentSheet, sentKey, studentId, name, 'B', weekKey, recipients);
        console.log(`[Version B] 送信: ${name} (今週空:${thisWeekEmpty}, 翌週空:${nextWeekEmpty})`);
      } catch(e) {
        console.error(`[Version B] 送信失敗: ${name} — ${e.message}`);
      }
    }
  });
}

// ============================================================
// 日付ユーティリティ
// ============================================================

/** 今週の月曜日（0:00:00）を返す */
function getThisMonday(date) {
  const d = new Date(date);
  const day = d.getDay(); // 0=日, 1=月, ..., 6=土
  const diff = (day === 0) ? -6 : 1 - day;
  d.setDate(d.getDate() + diff);
  d.setHours(0, 0, 0, 0);
  return d;
}

/** 翌週の月曜日（0:00:00）を返す */
function getNextMonday(date) {
  const thisMonday = getThisMonday(date);
  const next = new Date(thisMonday);
  next.setDate(thisMonday.getDate() + 7);
  return next;
}

// ============================================================
// メール文面
// ============================================================

function buildReminderRecipients(email, parentMail) {
  const list = [email];
  if (parentMail && parentMail !== '' && parentMail !== 'undefined') list.push(parentMail);
  return list.join(',');
}

function buildReminderSubjectA(name, nextMonday, nextSunday) {
  const from = Utilities.formatDate(nextMonday, 'Asia/Tokyo', 'M/d');
  const to   = Utilities.formatDate(nextSunday,  'Asia/Tokyo', 'M/d');
  return `【自習室】来週（${from}〜${to}）の自習予定を入力してください`;
}

function buildReminderBodyA(name, nextMonday, nextSunday) {
  const from = Utilities.formatDate(nextMonday, 'Asia/Tokyo', 'M月d日');
  const to   = Utilities.formatDate(nextSunday,  'Asia/Tokyo', 'M月d日');
  return `${name} さん

来週（${from}〜${to}）の自習室利用予定がまだ入力されていません。

Googleカレンダーに来週の予定を入力をお願いします。
自習の習慣化のために、先生も一緒に応援しています！

──────────────────────
SSS Education
──────────────────────`;
}

function buildReminderSubjectB(name, thisWeekEmpty, nextWeekEmpty) {
  const parts = [];
  if (thisWeekEmpty) parts.push('今週');
  if (nextWeekEmpty) parts.push('来週');
  return `【自習室】${parts.join('・')}の自習予定を入力してください`;
}

function buildReminderBodyB(name, thisMonday, thisSunday, nextMonday, nextSunday, thisWeekEmpty, nextWeekEmpty) {
  const lines = [`${name} さん\n`];
  if (thisWeekEmpty) {
    const from = Utilities.formatDate(thisMonday, 'Asia/Tokyo', 'M月d日');
    const to   = Utilities.formatDate(thisSunday,  'Asia/Tokyo', 'M月d日');
    lines.push(`・今週（${from}〜${to}）の予定が入力されていません`);
  }
  if (nextWeekEmpty) {
    const from = Utilities.formatDate(nextMonday, 'Asia/Tokyo', 'M月d日');
    const to   = Utilities.formatDate(nextSunday,  'Asia/Tokyo', 'M月d日');
    lines.push(`・来週（${from}〜${to}）の予定が入力されていません`);
  }
  lines.push('\nGoogleカレンダーに予定を入力をお願いします。');
  lines.push('自習の習慣化のために、先生も一緒に応援しています！');
  lines.push('\n──────────────────────\nSSS Education\n──────────────────────');
  return lines.join('\n');
}

// ============================================================
// 送信済みログ管理
// ============================================================

function getOrCreateReminderSentSheet(ss) {
  let sheet = ss.getSheetByName(REMINDER_CONFIG.SHEET_SENT);
  if (!sheet) {
    sheet = ss.insertSheet(REMINDER_CONFIG.SHEET_SENT);
    sheet.appendRow(['送信キー', '生徒ID', '生徒名', 'バージョン', '対象週', '送信先', '送信日時']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getReminderSentKeys(sentSheet) {
  const keys = new Set();
  sentSheet.getDataRange().getValues().slice(1).forEach(row => {
    keys.add(String(row[0]));
  });
  return keys;
}

function logReminderSent(sentSheet, sentKey, studentId, name, version, weekKey, recipients) {
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  sentSheet.appendRow([sentKey, studentId, name, version, weekKey, recipients, now]);
}

// ============================================================
// トリガー登録：月曜日 8:00 に実行（初回のみ手動実行）
// ============================================================
function setupReminderTrigger() {
  // 既存トリガーを削除
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'sendReminderVersionA' || fn === 'sendReminderVersionB') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Version A：毎週月曜 8:00
  ScriptApp.newTrigger('sendReminderVersionA')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();

  // Version B：毎週月曜 8:05（Aと少しずらす）
  ScriptApp.newTrigger('sendReminderVersionB')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();

  console.log('催促メールトリガーを登録しました（毎週月曜 8:00）');
}