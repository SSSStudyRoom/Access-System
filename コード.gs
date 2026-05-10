// ====================================================================
// 【列構成 最新版】インデックス定義（公開プロフィール / 個人情報マスタ共通）
// A:ID / B:生徒メアド / C:保護者メアド / D:氏名 / E:ニックネーム
// F:目標時間 / G:URLトークン / H:模試名 / I:模試日
// ====================================================================
const IDX_PROFILE = {
  ID: 0,             // A列
  STUDENT_EMAIL: 1,  // B列
  PARENT_EMAIL: 2,   // C列
  NAME: 3,           // D列
  NICKNAME: 4,       // E列
  GOAL_HOURS: 5,     // F列
  TOKEN: 6,          // G列
  EXAM_NAME: 7,      // H列
  EXAM_DATE: 8       // I列
};
const IDX_PERSONAL = IDX_PROFILE;

// ====================================================================
// 📊 スプレッドシートのURL定義（一元管理）
//   ★本番移行時はこの2つのURLを書き換えるだけで全関数が連動する★
//
//   - MAIN_SS_URL     : メインスプレッドシート
//                       「公開プロフィール」「管理シート」「生徒設定」を格納
//   - PERSONAL_SS_URL : 個人情報マスタ（極秘情報）
//                       現状は未使用だが、将来的に氏名・住所等をこちらへ分離する
//                       想定で定数として保持
// ====================================================================
const MAIN_SS_URL     = 'https://docs.google.com/spreadsheets/d/1QcruSLwoyPCQvCuaPK9m5Q3mFK2F2pacZPeu6VHEvps/edit';
const PERSONAL_SS_URL = 'https://docs.google.com/spreadsheets/d/1quXMzVooCCFb8Ig4zIgbhlBjkRNtUohWfu_iWomG7tI/edit';

/**
 * メインスプレッドシートを取得する共通ヘルパー。
 * 全関数は必ずこれを使うことで、参照先がブレないようにする。
 */
function getMainSpreadsheet() {
  return SpreadsheetApp.openByUrl(MAIN_SS_URL);
}

/**
 * 極秘情報スプレッドシート（個人情報マスタ）を取得する。
 * 現状未使用だが、将来の機能拡張時に使用。
 */
function getPersonalSpreadsheet() {
  return SpreadsheetApp.openByUrl(PERSONAL_SS_URL);
}

// ====================================================================
// 🔔 Google Chat Webhook URL
//   入退室通知用。WebhookのURLは秘密情報なのでGAS内のみで保管する
//   （リポジトリには絶対にアップロードしないこと）
// ====================================================================
const GCHAT_WEBHOOK_URL = 'https://chat.googleapis.com/v1/spaces/AAQAUHWNd34/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=Cj4ffnZEhNIzEAtmcwWLnKy-TEu_r08HKZsDPVB9FoM';

// ====================================================================
// 🔗 生徒マイページのベースURL
//   本番URLが変わったらここだけ書き換えればOK
// ====================================================================
const STUDENT_DASHBOARD_BASE_URL = 'https://sssstudyroom.github.io/Access-System/dashboard.html?token=';

// ====================================================================
// 📘 小テスト用スプレッドシートの設定
// ====================================================================
const TEST_SS_ID = '15C3TN3oMgx8tEWinCUx_TMg_ZzfgRBnhIOEt3tC82aA';
const TEST_TEMPLATE_NAME = 'テンプレート';

// ====================================================================
// 管理シート（打刻ログ）の列インデックス
// A列:日付 / B列:ID / D列:入室時刻 / E列:退室時刻 / F列:エール数
// ====================================================================
const IDX_LOG = {
  DATE: 0,    // A列
  ID: 1,      // B列
  IN: 3,      // D列
  OUT: 4,     // E列
  CHEERS: 5   // F列
};

// ====================================================================
// ① アクセス振り分け処理
// ====================================================================
function doGet(e) {
  const params = e.parameter || {};
  const action = params.action;
  const token  = params.token;

  if (action === 'stats' && token) {
    try {
      const data = getStudentStats(token);
      return jsonResponse({ ok: true, data: data });
    } catch (err) {
      return jsonResponse({ ok: false, error: String(err && err.message || err) });
    }
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('SSS Education 自習室受付')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = body.action;

    if (action === 'saveSettings') {
      const msg = saveStudentSettings(body.token, body.examName, body.examDate);
      return jsonResponse({ ok: true, message: msg });
    }
    if (action === 'scan') {
      const msg = processScan(body.id);
      return jsonResponse({ ok: true, message: msg });
    }

    return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err && err.message || err) });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ====================================================================
// ② ダッシュボード用データ集計
// ====================================================================
function getStudentStats(studentToken) {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get(studentToken);
  if (cachedData) return JSON.parse(cachedData);

  const ss = getMainSpreadsheet();
  const profileSheet = ss.getSheetByName('公開プロフィール');
  const logSheet = ss.getSheetByName('管理シート');
  const settingsSheet = ss.getSheetByName('生徒設定') || ss.insertSheet('生徒設定');

  const profileData = profileSheet.getDataRange().getValues();
  const logData = logSheet.getDataRange().getValues();
  const settingsData = settingsSheet.getDataRange().getValues();

  let targetId = null;
  let targetNickname = "学習者";
  let weeklyGoalHours = 15;
  let customExam = { name: "", date: "" };
  let nickMap = {};

  for (let i = 1; i < profileData.length; i++) {
    let pId = String(profileData[i][IDX_PROFILE.ID]).trim();
    let pToken = String(profileData[i][IDX_PROFILE.TOKEN]).trim();
    let pName = profileData[i][IDX_PROFILE.NAME] || profileData[i][IDX_PROFILE.NICKNAME] || pId;
    nickMap[pId] = pName;

    if (pToken === studentToken) {
      targetId = pId;
      targetNickname = pName;
      let goalRaw = profileData[i][IDX_PROFILE.GOAL_HOURS];
      if (goalRaw && !isNaN(parseFloat(goalRaw))) weeklyGoalHours = parseFloat(goalRaw);
    }
  }

  if (!targetId) throw new Error("無効なURLです。管理者に連絡してください。");

  for (let i = 1; i < settingsData.length; i++) {
    if (String(settingsData[i][0]).trim() === targetId) {
      customExam.name = settingsData[i][1];
      let sDate = settingsData[i][2];
      customExam.date = sDate instanceof Date ? Utilities.formatDate(sDate, "JST", "yyyy/MM/dd") : sDate;
      break;
    }
  }

  const toDate = (val) => {
    if (!val) return null;
    if (val instanceof Date) return val;
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  };

  const now = new Date();
  const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  const getMonday = (d) => {
    let day = d.getDay(), diff = d.getDate() - day + (day == 0 ? -6 : 1);
    return new Date(d.getFullYear(), d.getMonth(), diff);
  };
  const thisMondayMs = getMonday(new Date()).getTime();

  let totalMs = 0, todayMs = 0, weekMs = 0, todayCheers = 0;
  let dailyMap = {}, weeklyMap = {}, monthlyMap = {}, weeklyRankingMap = {};
  let recentActions = [], historyList = [];
  let uniqueDates = new Set();

  for(let i = logData.length - 1; i >= 1; i--){
    let rowId = String(logData[i][IDX_LOG.ID]).trim();
    let inTime = toDate(logData[i][IDX_LOG.IN]);
    let outTime = toDate(logData[i][IDX_LOG.OUT]);

    if(inTime) {
      let logDateMs = new Date(inTime.getFullYear(), inTime.getMonth(), inTime.getDate()).getTime();
      let diffMs = (outTime) ? outTime.getTime() - inTime.getTime() : 0;
      if (diffMs < 0) diffMs = 0;

      if(rowId === targetId) {
        if(diffMs > 0) {
          totalMs += diffMs;
          if (logDateMs === todayStart) todayMs += diffMs;
          if (logDateMs >= (todayStart - 7*24*60*60*1000)) weekMs += diffMs;

          let dayKey = Utilities.formatDate(inTime, "JST", "MM/dd");
          dailyMap[dayKey] = (dailyMap[dayKey] || 0) + diffMs;

          let monday = getMonday(inTime);
          let weekKey = Utilities.formatDate(monday, "JST", "MM/dd") + "週";
          weeklyMap[weekKey] = (weeklyMap[weekKey] || 0) + diffMs;

          let monthKey = Utilities.formatDate(inTime, "JST", "yyyy/MM");
          monthlyMap[monthKey] = (monthlyMap[monthKey] || 0) + diffMs;

          uniqueDates.add(Utilities.formatDate(inTime, "JST", "yyyy/MM/dd"));

          historyList.push({
            date: Utilities.formatDate(inTime, "JST", "MM/dd"),
            in: Utilities.formatDate(inTime, "JST", "HH:mm"),
            out: outTime ? Utilities.formatDate(outTime, "JST", "HH:mm") : "---",
            time: formatTime(diffMs)
          });
        }
        if (logDateMs === todayStart) {
          todayCheers += parseInt(logData[i][IDX_LOG.CHEERS] || 0);
        }
      }

      // 今週ランキング集計（全員、退室済みのみ）
      if(logDateMs >= thisMondayMs && diffMs > 0) {
        weeklyRankingMap[rowId] = (weeklyRankingMap[rowId] || 0) + diffMs;
      }

      if(recentActions.length < 5) {
        recentActions.push({
          id: rowId,
          name: nickMap[rowId] || "学習者",
          action: (outTime) ? "退室" : "入室",
          time: Utilities.formatDate((outTime ? outTime : inTime), "JST", "HH:mm")
        });
      }
    }
  }

  // ★ ランキング修正：未退室（diffMs=0）も「現在進行中の入室」として含める
  for(let i = logData.length - 1; i >= 1; i--){
    let rowId = String(logData[i][IDX_LOG.ID]).trim();
    let inTime = toDate(logData[i][IDX_LOG.IN]);
    let outTime = toDate(logData[i][IDX_LOG.OUT]);
    if (!inTime) continue;
    let logDateMs = new Date(inTime.getFullYear(), inTime.getMonth(), inTime.getDate()).getTime();
    if (logDateMs < thisMondayMs) break;

    if (!outTime) {
      let liveMs = now.getTime() - inTime.getTime();
      if (liveMs > 0 && liveMs < 12 * 3600000) {
        weeklyRankingMap[rowId] = (weeklyRankingMap[rowId] || 0) + liveMs;
      }
    }
  }

  let top5Ranking = Object.keys(weeklyRankingMap).map(id => {
    return { name: nickMap[id] || "学習者", hours: (weeklyRankingMap[id] / 3600000).toFixed(1) };
  }).filter(r => parseFloat(r.hours) > 0)
    .sort((a, b) => parseFloat(b.hours) - parseFloat(a.hours))
    .slice(0, 5);

  // 小テスト進捗（シートが無ければテンプレートから自動生成）
  const testProgress = getOrCreateTestProgress(targetNickname, targetId);

  const formatChartData = (map) => {
    let labels = Object.keys(map).sort();
    let values = labels.map(k => (map[k]/3600000).toFixed(1));
    return { labels, values };
  };

  const resultData = {
    name: targetNickname,
    id: targetId,
    today: formatTime(todayMs),
    week: formatTime(weekMs),
    total: formatTime(totalMs),
    daily: formatChartData(dailyMap),
    weekly: formatChartData(weeklyMap),
    monthly: formatChartData(monthlyMap),
    rank: calculateRank(totalMs / 3600000),
    community: { ranking: top5Ranking, feed: recentActions },
    weeklyGoal: { currentHours: (weekMs / 3600000).toFixed(1), targetHours: weeklyGoalHours, percent: Math.min(100, ((weekMs/3600000)/weeklyGoalHours)*100).toFixed(1) },
    history: historyList.slice(0, 30),
    streak: { totalDays: uniqueDates.size, active: calculateStreak(uniqueDates) },
    tests: testProgress,
    exam: customExam,
    cheers: todayCheers
  };

  cache.put(studentToken, JSON.stringify(resultData), 900);
  return resultData;
}

// ====================================================================
// ③ 設定保存
// ====================================================================
function saveStudentSettings(token, examName, examDate) {
  const ss = getMainSpreadsheet();
  const profileData = ss.getSheetByName('公開プロフィール').getDataRange().getValues();
  let studentId = null;

  for (let i = 1; i < profileData.length; i++) {
    if (String(profileData[i][IDX_PROFILE.TOKEN]).trim() === token) {
      studentId = String(profileData[i][IDX_PROFILE.ID]).trim();
      break;
    }
  }
  if (!studentId) return "認証エラー";

  const settingsSheet = ss.getSheetByName("生徒設定");
  const settingsData = settingsSheet.getDataRange().getValues();
  let targetRow = -1;

  for (let i = 1; i < settingsData.length; i++) {
    if (String(settingsData[i][0]).trim() === studentId) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow === -1) targetRow = settingsSheet.getLastRow() + 1;

  settingsSheet.getRange(targetRow, 1).setValue(studentId);
  settingsSheet.getRange(targetRow, 2).setValue(examName);
  settingsSheet.getRange(targetRow, 3).setValue(examDate);

  CacheService.getScriptCache().remove(token);

  return "設定を保存しました！";
}

// ====================================================================
// ④ 受付打刻処理（保護者メール通知 + Google Chat通知つき）
// ====================================================================
function processScan(studentId) {
  if (!studentId) return "エラー：IDが読み込めませんでした";
  const cleanTargetId = String(studentId).trim();
  const ss = getMainSpreadsheet();
  const sheet = ss.getSheetByName("管理シート");
  const now = new Date();
  const todayStr = Utilities.formatDate(now, "JST", "yyyy/MM/dd");

  const sheetMaxRow = sheet.getLastRow();

  // 「B列(ID)に実データがある最終行」を求める
  let actualLastDataRow = 1;
  if (sheetMaxRow >= 2) {
    const bColumn = sheet.getRange(2, 2, sheetMaxRow - 1, 1).getValues();
    for (let i = bColumn.length - 1; i >= 0; i--) {
      const v = bColumn[i][0];
      if (v !== "" && v !== null && v !== undefined) {
        actualLastDataRow = i + 2;
        break;
      }
    }
  }

  // 直近100件だけを後ろから取得
  let existingRowIndex = -1;
  let existingInTime = null;
  let searchSize = Math.min(100, Math.max(0, actualLastDataRow - 1));
  let attempt = 0;
  let maxAttempts = 3;

  while (existingRowIndex === -1 && attempt < maxAttempts && searchSize > 0) {
    let startRow = Math.max(2, actualLastDataRow - searchSize + 1);
    let numRows = actualLastDataRow - startRow + 1;
    if (numRows <= 0) break;

    const data = sheet.getRange(startRow, 1, numRows, 5).getValues();

    for (let i = data.length - 1; i >= 0; i--) {
      const rowDate = data[i][IDX_LOG.DATE];
      const rowDateStr = rowDate instanceof Date ? Utilities.formatDate(rowDate, "JST", "yyyy/MM/dd") : "";
      if (rowDateStr !== todayStr) continue;
      if (String(data[i][IDX_LOG.ID]).trim() !== cleanTargetId) continue;
      if (data[i][IDX_LOG.OUT] === "" || data[i][IDX_LOG.OUT] === null) {
        existingRowIndex = startRow + i;
        existingInTime = data[i][IDX_LOG.IN];
        break;
      }
    }

    attempt++;
    searchSize *= 3;
  }

  let actionType = "";
  let studyMs = 0;
  if (existingRowIndex > 0) {
    sheet.getRange(existingRowIndex, IDX_LOG.OUT + 1).setValue(now);
    actionType = "退室";
    if (existingInTime instanceof Date) {
      studyMs = now.getTime() - existingInTime.getTime();
      if (studyMs < 0) studyMs = 0;
    }
  } else {
    // 入室処理：実データ最終行の直下に追記（C列の数式は触らない）
    const targetRow = actualLastDataRow + 1;
    sheet.getRange(targetRow, 1, 1, 2).setValues([[now, cleanTargetId]]);
    sheet.getRange(targetRow, 4).setValue(now);
    actionType = "入室";
  }

  // 【優先度1】画面応答用の情報を確定
  const pData = ss.getSheetByName('公開プロフィール').getDataRange().getValues();
  let studentName = "学習者";
  let studentToken = null;
  let parentEmail = null;

  for (let i = 1; i < pData.length; i++) {
    if (String(pData[i][IDX_PROFILE.ID]).trim() === cleanTargetId) {
      studentName = pData[i][IDX_PROFILE.NAME] || pData[i][IDX_PROFILE.NICKNAME] || "学習者";
      studentToken = String(pData[i][IDX_PROFILE.TOKEN]).trim();
      const rawParentEmail = pData[i][IDX_PROFILE.PARENT_EMAIL];
      if (rawParentEmail && String(rawParentEmail).indexOf('@') > 0) {
        parentEmail = String(rawParentEmail).trim();
      }
      break;
    }
  }

  const responseMessage = `${studentName} さんが ${actionType} しました！`;

  if (studentToken) CacheService.getScriptCache().remove(studentToken);

  // 【優先度2】保護者メール通知
  if (parentEmail) {
    try {
      notifyParent(studentName, parentEmail, actionType, now, studyMs);
    } catch (e) {
      console.error('[優先度2] 保護者メール処理で例外:', e);
    }
  }

  // 【優先度3】Google Chat Webhook通知
  try {
    notifyGoogleChat(studentName, actionType, now, studyMs);
  } catch (e) {
    console.error('[優先度3] Webhook処理で例外:', e);
  }

  return responseMessage;
}

// ====================================================================
// ④-2 保護者向けメール送信
// ====================================================================
// ====================================================================
// ④-2 保護者向けメール送信（入退室の両方に対応）
// ====================================================================
function notifyParent(studentName, parentEmail, actionType, datetime, studyMs) {
  const timeStr = Utilities.formatDate(datetime, "JST", "yyyy/MM/dd HH:mm");
  const subject = `【SSS Education】${studentName}さんが${actionType}しました`;
  
  // 日本語の助詞を調整（入室なら「に」、退室なら「を」）
  const particle = (actionType === "入室") ? "に" : "を";

  let body =
    `${studentName}さんの保護者様\n\n` +
    `お世話になっております。SSS Education自習室管理部です。\n` +
    `${studentName}さんが自習室${particle}${actionType}したことをお知らせいたします。\n\n` +
    `■ ${actionType}時刻： ${timeStr}\n`;

  // 退室時のみ学習時間を表示
  if (actionType === "退室" && studyMs > 0) {
    body += `■ 本日の学習時間： ${formatTime(studyMs)}\n`;
  }

  body +=
    `\n` +
    `こちらのメールは送信専用となります。\n` +
    `よろしくお願いいたします。\n`;

  try {
    MailApp.sendEmail({ to: parentEmail, subject: subject, body: body });
  } catch (err) {
    console.error(`保護者メール送信失敗 (${parentEmail}):`, err);
  }
}

// ====================================================================
// 🔧 デバッグ用：入室メールの送信テスト
// ====================================================================
function testEntryEmailSend() {
  const testEmail = 'Ahmadtanzeel0204@gmail.com'; // ★ご自身のメアドでテストしてください
  notifyParent('テスト 太郎', testEmail, '入室', new Date(), 0);
  console.log(`入室テストメールを ${testEmail} に送信しました`);
}

// ====================================================================
// ④-3 Google Chat通知（Webhook経由）
// ====================================================================
function notifyGoogleChat(studentName, actionType, datetime, studyMs) {
  if (!GCHAT_WEBHOOK_URL) return;

  const timeStr = Utilities.formatDate(datetime, "JST", "HH:mm");
  let message;

  if (actionType === "入室") {
    message = `📢 *${studentName}* さんが入室しました（${timeStr}）`;
  } else if (actionType === "退室") {
    const studyTime = studyMs > 0 ? formatTime(studyMs) : "—";
    message = `🐍 *${studentName}* さんが退室しました（${timeStr} / 学習時間: ${studyTime}）`;
  } else {
    message = `${studentName} さんが${actionType}しました（${timeStr}）`;
  }

  try {
    UrlFetchApp.fetch(GCHAT_WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: message }),
      muteHttpExceptions: true
    });
  } catch (err) {
    console.error('Google Chat通知失敗:', err);
  }
}

// ====================================================================
// ⑥ 生徒マイページURLのメール下書き一括作成
// ====================================================================
function createDraftsByIds(idsString) {
  if (!idsString || !String(idsString).trim()) {
    return { success: false, message: "⚠ IDが入力されていません" };
  }

  const ids = String(idsString)
    .split(/[,、\s]+/)
    .map(s => s.trim())
    .filter(s => s);

  if (ids.length === 0) {
    return { success: false, message: "⚠ 有効なIDがありません" };
  }

  const ss = getMainSpreadsheet();
  const profileData = ss.getSheetByName('公開プロフィール').getDataRange().getValues();

  const profileMap = {};
  for (let i = 1; i < profileData.length; i++) {
    const id = String(profileData[i][IDX_PROFILE.ID]).trim();
    if (!id) continue;
    profileMap[id] = {
      name:         profileData[i][IDX_PROFILE.NAME] || profileData[i][IDX_PROFILE.NICKNAME] || id,
      parentEmail:  String(profileData[i][IDX_PROFILE.PARENT_EMAIL] || '').trim(),
      studentEmail: String(profileData[i][IDX_PROFILE.STUDENT_EMAIL] || '').trim(),
      token:        String(profileData[i][IDX_PROFILE.TOKEN] || '').trim()
    };
  }

  let successList = [];
  let errorList = [];

  ids.forEach(id => {
    const profile = profileMap[id];
    if (!profile) {
      errorList.push(`✗ ${id}：プロフィールが見つかりません`);
      return;
    }
    if (!profile.token) {
      errorList.push(`✗ ${id} (${profile.name})：URLトークン未設定`);
      return;
    }

    // 送信先：必ず生徒メアド（B列）
    const to = profile.studentEmail;
    if (!to || to.indexOf('@') < 0) {
      errorList.push(`✗ ${id} (${profile.name})：生徒メアド(B列)が未設定`);
      return;
    }

    const url = STUDENT_DASHBOARD_BASE_URL + encodeURIComponent(profile.token);
    const subject = `【SSS Education】${profile.name}さん専用 学習ダッシュボードのご案内`;
    const body =
      `${profile.name} さん\n\n` +
      `いつもSSS Educationでの学習、お疲れさまです。\n` +
      `あなた専用の学習ダッシュボードをご用意しました。\n\n` +
      `▼ あなたのマイページURL\n${url}\n\n` +
      `※ このURLはあなた個人専用です。他の方には共有しないようご注意ください。\n` +
      `※ スマートフォンの場合、ホーム画面に追加すると、アプリのように使えます。\n\n` +
      `【マイページでできること】\n` +
      `・自分の自習室入退室履歴の確認\n` +
      `・週間/月間の学習時間グラフ\n` +
      `・小テストの進捗状況\n` +
      `・受付用のQRコード（毎回これをかざせばOK）\n\n` +
      `不明な点があれば、講師までお気軽にお声がけください。\n\n` +
      `──────────────────\n` +
      `SSS Education\n` +
      `──────────────────\n`;

    try {
      GmailApp.createDraft(to, subject, body);
      successList.push(`✓ ${profile.name} (${id}) → ${to}`);
    } catch (e) {
      errorList.push(`✗ ${id} (${profile.name})：下書き作成失敗 - ${e.message}`);
    }
  });

  let message = '';
  if (successList.length > 0) {
    message += `✅ ${successList.length}件の下書きを作成しました\n\n`;
    message += successList.join('\n');
  }
  if (errorList.length > 0) {
    if (message) message += '\n\n';
    message += `⚠ ${errorList.length}件のエラー\n\n`;
    message += errorList.join('\n');
  }
  if (successList.length > 0) {
    message += `\n\n👉 Gmailの「下書き」フォルダで内容を確認してから送信してください。`;
  }

  return { success: successList.length > 0, message: message };
}

// ====================================================================
// ⑦ 小テスト進捗の取得（シートが無ければテンプレートから自動生成）
// ====================================================================
function getOrCreateTestProgress(studentName, studentId) {
  const testProgress = [];

  try {
    const testSs = SpreadsheetApp.openById(TEST_SS_ID);

    let testSheet = testSs.getSheetByName(studentName);
    if (!testSheet) {
      testSheet = testSs.getSheetByName(studentId);
    }

    if (!testSheet) {
      const templateSheet = testSs.getSheetByName(TEST_TEMPLATE_NAME);
      if (!templateSheet) {
        console.error(`テンプレートシート「${TEST_TEMPLATE_NAME}」が見つかりません`);
        return testProgress;
      }

      const lock = LockService.getScriptLock();
      const acquired = lock.tryLock(10000);
      if (!acquired) {
        console.warn('LockService取得失敗。シート自動生成をスキップします。');
        return testProgress;
      }

      try {
        testSheet = testSs.getSheetByName(studentName);
        if (!testSheet) {
          const copied = templateSheet.copyTo(testSs);
          copied.setName(studentName);
          testSheet = copied;
          console.log(`新規シート作成: ${studentName}`);
        }
      } catch (copyErr) {
        console.error('シート作成失敗:', copyErr.message);
        return testProgress;
      } finally {
        lock.releaseLock();
      }
    }

    if (testSheet) {
      const testData = testSheet.getDataRange().getValues();
      for (let i = 1; i < testData.length; i++) {
        const bookName = testData[i][0];
        if (!bookName) continue;
        let totalCount = 0, completedCount = 0;
        const blocks = [];
        for (let j = 1; j <= 30; j++) {
          const cellValue = testData[i][j];
          if (cellValue === true || cellValue === false) {
            totalCount++;
            if (cellValue === true) { completedCount++; blocks.push(true); }
            else { blocks.push(false); }
          }
        }
        if (totalCount > 0) {
          testProgress.push({
            name: bookName,
            total: totalCount,
            completed: completedCount,
            blocks: blocks
          });
        }
      }
    }
  } catch (e) {
    console.error("小テスト取得失敗:", e.message);
  }

  return testProgress;
}

// ====================================================================
// 補助関数群
// ====================================================================
function formatTime(ms) { return `${Math.floor(ms/3600000)}時間${Math.floor((ms%3600000)/60000)}分`; }

function calculateRank(hours) {
  if (hours >= 300) return { current: "SSS MASTER", remainHours: 0 };
  if (hours >= 150) return { current: "PLATINUM", remainHours: (300 - hours).toFixed(1) };
  if (hours >= 50)  return { current: "GOLD", remainHours: (150 - hours).toFixed(1) };
  if (hours >= 10)  return { current: "SILVER", remainHours: (50 - hours).toFixed(1) };
  return { current: "BRONZE", remainHours: (10 - hours).toFixed(1) };
}

function calculateStreak(uniqueDates) {
  let sorted = Array.from(uniqueDates).sort().reverse();
  let streak = 0, check = new Date();
  check.setHours(0,0,0,0);
  for(let i=0; i<sorted.length; i++) {
    let d = new Date(sorted[i]); d.setHours(0,0,0,0);
    if((check - d) / (1000*60*60*24) <= 1) { streak++; check = d; } else break;
  }
  return streak;
}

// ====================================================================
// UI操作系（コンテナバインド時のみ動作）
// ====================================================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('★管理メニュー').addItem('QRメールの下書き作成', 'showDraftDialog').addToUi();
}

function showDraftDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('draftDialog').setWidth(400).setHeight(320),
    'QRコード送付'
  );
}

// ====================================================================
// ⑤ 旧エール送信（廃止済み・互換用に残置）
// ====================================================================
function sendCheer(targetStudentId) {
  return "エール機能は廃止されました";
}

// ====================================================================
// 🔧 デバッグ用：今週のランキング状況を出力
// ====================================================================
function debugWeeklyRanking() {
  const ss = getMainSpreadsheet();
  const logSheet = ss.getSheetByName('管理シート');
  const profileSheet = ss.getSheetByName('公開プロフィール');

  const logData = logSheet.getDataRange().getValues();
  const profileData = profileSheet.getDataRange().getValues();

  let nickMap = {};
  for (let i = 1; i < profileData.length; i++) {
    let pId = String(profileData[i][IDX_PROFILE.ID]).trim();
    let pName = profileData[i][IDX_PROFILE.NAME] || profileData[i][IDX_PROFILE.NICKNAME] || pId;
    nickMap[pId] = pName;
  }

  const getMonday = (d) => {
    let day = d.getDay(), diff = d.getDate() - day + (day == 0 ? -6 : 1);
    return new Date(d.getFullYear(), d.getMonth(), diff);
  };
  const thisMondayMs = getMonday(new Date()).getTime();
  console.log('今週月曜:', new Date(thisMondayMs));

  let weeklyMap = {};
  let weekRecords = [];

  for (let i = logData.length - 1; i >= 1; i--) {
    let rowId = String(logData[i][IDX_LOG.ID]).trim();
    let inTime = logData[i][IDX_LOG.IN];
    let outTime = logData[i][IDX_LOG.OUT];

    if (!(inTime instanceof Date)) continue;
    let logDateMs = new Date(inTime.getFullYear(), inTime.getMonth(), inTime.getDate()).getTime();
    if (logDateMs < thisMondayMs) continue;

    let diffMs = (outTime instanceof Date) ? outTime.getTime() - inTime.getTime() : 0;
    weekRecords.push({
      id: rowId,
      name: nickMap[rowId] || '不明',
      in: inTime,
      out: outTime || '(未退室)',
      hours: diffMs > 0 ? (diffMs / 3600000).toFixed(2) : 0
    });
    if (diffMs > 0) {
      weeklyMap[rowId] = (weeklyMap[rowId] || 0) + diffMs;
    }
  }

  console.log('今週のレコード件数:', weekRecords.length);
  weekRecords.forEach(r => console.log(`  ${r.name}(${r.id}): ${r.in} → ${r.out} (${r.hours}h)`));
  console.log('\n今週ランキング集計結果:');
  Object.keys(weeklyMap).forEach(id => {
    console.log(`  ${nickMap[id]}: ${(weeklyMap[id]/3600000).toFixed(2)}h`);
  });
}

// ====================================================================
// 🔧 デバッグ用：processScanの速度テスト
// ====================================================================
function testProcessScanSpeed() {
  const start = new Date().getTime();
  const result = processScan('テスト用の実在する学習者ID');
  const elapsed = new Date().getTime() - start;
  console.log(`処理時間: ${elapsed}ms / 結果: ${result}`);
}

// ====================================================================
// 🔧 デバッグ用：「公開プロフィール」シートの列構成を確認
// ====================================================================
function debugProfileColumns() {
  const sheet = getMainSpreadsheet().getSheetByName('公開プロフィール');
  if (!sheet) { console.log('シート「公開プロフィール」が見つかりません'); return; }
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  console.log('=== シート上の列構成 ===');
  headers.forEach((h, i) => {
    const col = String.fromCharCode(65 + i);
    console.log(`  ${col}列(${i}): ${h}`);
  });

  console.log('\n=== コード側の認識（IDX_PROFILE） ===');
  Object.keys(IDX_PROFILE).forEach(key => {
    const idx = IDX_PROFILE[key];
    const col = String.fromCharCode(65 + idx);
    const headerVal = idx < headers.length ? headers[idx] : '(範囲外)';
    console.log(`  ${key.padEnd(15)} → ${col}列(${idx})：${headerVal}`);
  });
}

// ====================================================================
// 🔧 デバッグ用：保護者メール送信テスト
// ====================================================================
function testEmailSend() {
  const testEmail = 'Ahmadtanzeel0204@gmail.com';  // ★自分のメアドに変更
  notifyParent('テスト 太郎', testEmail, '退室', new Date(), 3 * 3600000 + 25 * 60000);
  console.log(`テストメールを ${testEmail} に送信しました`);
}

// ====================================================================
// 🔧 デバッグ用：小テストシート自動生成のテスト
// ====================================================================
function testCreateTestSheet() {
  const testName = 'テスト用_' + Utilities.formatDate(new Date(), 'JST', 'MMddHHmm');
  console.log(`テスト名「${testName}」でシート生成を試行します...`);

  const result = getOrCreateTestProgress(testName, 'TEST_ID_999');
  console.log('取得結果:', JSON.stringify(result, null, 2));

  const testSs = SpreadsheetApp.openById(TEST_SS_ID);
  const newSheet = testSs.getSheetByName(testName);
  if (newSheet) {
    console.log(`✓ シート「${testName}」が正常に作成されました（行数: ${newSheet.getLastRow()}）`);
    console.log('  → 動作確認後、不要なら手動で削除してください');
  } else {
    console.log(`✗ シートが作成されませんでした`);
  }
}

// ====================================================================
// 🔧 デバッグ用：スプレッドシート参照のヘルスチェック
//    本番移行後、最初に実行することを強く推奨
// ====================================================================
function debugSpreadsheetConnection() {
  console.log('=== スプレッドシート参照チェック ===\n');

  // メインSS
  try {
    const mainSs = getMainSpreadsheet();
    console.log(`✓ メインSS接続OK`);
    console.log(`  名前: ${mainSs.getName()}`);
    console.log(`  URL: ${mainSs.getUrl()}`);

    const sheets = ['公開プロフィール', '管理シート', '生徒設定'];
    sheets.forEach(name => {
      const sheet = mainSs.getSheetByName(name);
      console.log(`  ${sheet ? '✓' : '✗'} シート「${name}」: ${sheet ? '存在' : '見つかりません'}`);
    });
  } catch (e) {
    console.log(`✗ メインSS接続失敗: ${e.message}`);
  }

  console.log('');

  // 極秘情報SS
  try {
    const personalSs = getPersonalSpreadsheet();
    console.log(`✓ 極秘情報SS接続OK`);
    console.log(`  名前: ${personalSs.getName()}`);
    console.log(`  URL: ${personalSs.getUrl()}`);
  } catch (e) {
    console.log(`✗ 極秘情報SS接続失敗: ${e.message}`);
  }

  console.log('');

  // 小テストSS
  try {
    const testSs = SpreadsheetApp.openById(TEST_SS_ID);
    console.log(`✓ 小テストSS接続OK`);
    console.log(`  名前: ${testSs.getName()}`);
    const tmpl = testSs.getSheetByName(TEST_TEMPLATE_NAME);
    console.log(`  ${tmpl ? '✓' : '✗'} テンプレート「${TEST_TEMPLATE_NAME}」: ${tmpl ? '存在' : '見つかりません'}`);
  } catch (e) {
    console.log(`✗ 小テストSS接続失敗: ${e.message}`);
  }
}