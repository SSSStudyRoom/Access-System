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
// 🔔 Google Chat Webhook URL
// ====================================================================
const GCHAT_WEBHOOK_URL = 'https://chat.googleapis.com/v1/spaces/AAQAUHWNd34/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=YEAO0z4xaLIh9y2bvuVsQqHR0ncqPk0ZdTu-nB3KOcg';

// ====================================================================
// 🔗 生徒マイページのベースURL
// ====================================================================
const STUDENT_DASHBOARD_BASE_URL = 'https://ahmadtanzeel.github.io/-dev-webpages-/dashboard.html?token=';

// ====================================================================
// 📘 小テスト用スプレッドシートの設定
// ====================================================================
const TEST_SS_ID = '1uKXnpKeGCuyPpRAryFP7Ou4K4t3SgTNksZFQAhvj45E';
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
// 🆕 新機能用シート名
// ====================================================================
const SHEET_STUDY_LOG  = '学習ログ';        // スタディログ投稿
const SHEET_REACTIONS  = 'リアクション';    // スタディログへの👍など
const SHEET_RIVALS     = 'ライバル登録';    // 各生徒のライバル設定

// ====================================================================
// ① アクセス振り分け処理
// ====================================================================
function doGet(e) {
  const params = e.parameter || {};
  const action = params.action;
  const token  = params.token;

  try {
    if (action === 'stats' && token) {
      const data = getStudentStats(token);
      return jsonResponse({ ok: true, data: data });
    }
    if (action === 'heatmap' && token) {
      const data = getHeatmap(token);
      return jsonResponse({ ok: true, data: data });
    }
    if (action === 'hourly' && token) {
      const data = getHourlyProfile(token);
      return jsonResponse({ ok: true, data: data });
    }
    if (action === 'studyLogs' && token) {
      const data = getStudyLogs(token);
      return jsonResponse({ ok: true, data: data });
    }
    if (action === 'studentList' && token) {
      const data = getStudentList(token);
      return jsonResponse({ ok: true, data: data });
    }
    if (action === 'rivals' && token) {
      const data = getRivals(token);
      return jsonResponse({ ok: true, data: data });
    }
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err && err.message || err) });
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
    // 🆕 スタディログ投稿
    if (action === 'postStudyLog') {
      const msg = postStudyLog(body.token, body.message, body.hours);
      return jsonResponse({ ok: true, message: msg });
    }
    // 🆕 スタディログにリアクション
    if (action === 'reactStudyLog') {
      const msg = reactStudyLog(body.token, body.logId, body.emoji);
      return jsonResponse({ ok: true, message: msg });
    }
    // 🆕 ライバル登録/解除
    if (action === 'updateRivals') {
      const msg = updateRivals(body.token, body.rivalIds || []);
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
// 🛠 ユーティリティ：トークン→生徒ID変換
// ====================================================================
function findStudentByToken(token) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const profileData = ss.getSheetByName('公開プロフィール').getDataRange().getValues();
  for (let i = 1; i < profileData.length; i++) {
    if (String(profileData[i][IDX_PROFILE.TOKEN]).trim() === token) {
      return {
        id: String(profileData[i][IDX_PROFILE.ID]).trim(),
        name: profileData[i][IDX_PROFILE.NAME] || profileData[i][IDX_PROFILE.NICKNAME] || ''
      };
    }
  }
  return null;
}

// ====================================================================
// 🛠 ユーティリティ：シート確保（無ければ作成）
// ====================================================================
function ensureSheet(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    }
  }
  return sheet;
}

// ====================================================================
// ② ダッシュボード用データ集計（★ メッセージ・ライバル比較を追加）
// ====================================================================
function getStudentStats(studentToken) {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get(studentToken);
  if (cachedData) return JSON.parse(cachedData);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
  const lastMondayMs = thisMondayMs - 7 * 24 * 3600000;

  let totalMs = 0, todayMs = 0, weekMs = 0, lastWeekMs = 0, todayCheers = 0;
  let dailyMap = {}, weeklyMap = {}, monthlyMap = {}, weeklyRankingMap = {};
  let recentActions = [], historyList = [];
  let uniqueDates = new Set();

  // 🆕 ライバル比較用：他生徒の今週時間
  let allWeeklyMap = {};

  // 🆕 過去30日の自分の日別データ（GitHub風ヒートマップではないが比較用）
  // ※ヒートマップ本体は別APIで取る
  let last30dMs = 0;
  let prev30dMs = 0;
  const ms30d = 30 * 24 * 3600000;

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

          // 🆕 過去30日 vs その前の30日（傾向比較）
          if (inTime.getTime() >= now.getTime() - ms30d) {
            last30dMs += diffMs;
          } else if (inTime.getTime() >= now.getTime() - 2 * ms30d) {
            prev30dMs += diffMs;
          }

          // 🆕 先週分（月曜〜日曜）
          if (logDateMs >= lastMondayMs && logDateMs < thisMondayMs) {
            lastWeekMs += diffMs;
          }

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

      // 今週ランキング集計＆全員の今週時間（ライバル機能用）
      if(logDateMs >= thisMondayMs && diffMs > 0) {
        weeklyRankingMap[rowId] = (weeklyRankingMap[rowId] || 0) + diffMs;
        allWeeklyMap[rowId] = (allWeeklyMap[rowId] || 0) + diffMs;
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

  // 入室中（未退室）を暫定加算
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
        allWeeklyMap[rowId] = (allWeeklyMap[rowId] || 0) + liveMs;
      }
    }
  }

  let top5Ranking = Object.keys(weeklyRankingMap).map(id => {
    return { name: nickMap[id] || "学習者", hours: (weeklyRankingMap[id] / 3600000).toFixed(1) };
  }).filter(r => parseFloat(r.hours) > 0)
    .sort((a, b) => parseFloat(b.hours) - parseFloat(a.hours))
    .slice(0, 5);

  // ====================================================================
  // 🆕 ライバル比較データ
  // ====================================================================
  const rivalIds = getRivalIdsFor(targetId);
  const rivalCompare = rivalIds.map(rid => {
    return {
      id: rid,
      name: nickMap[rid] || "学習者",
      weekHours: ((allWeeklyMap[rid] || 0) / 3600000).toFixed(1)
    };
  });

  // ====================================================================
  // 🆕 比較データ（先週比・前月比）
  // ====================================================================
  const compareData = {
    thisWeekHours: (weekMs / 3600000).toFixed(1),
    lastWeekHours: (lastWeekMs / 3600000).toFixed(1),
    last30dHours: (last30dMs / 3600000).toFixed(1),
    prev30dHours: (prev30dMs / 3600000).toFixed(1)
  };

  // 小テスト
  const testProgress = getOrCreateTestProgress(targetNickname, targetId);

  const formatChartData = (map) => {
    let labels = Object.keys(map).sort();
    let values = labels.map(k => (map[k]/3600000).toFixed(1));
    return { labels, values };
  };

  const totalHours = totalMs / 3600000;

  const resultData = {
    name: targetNickname,
    id: targetId,
    today: formatTime(todayMs),
    week: formatTime(weekMs),
    total: formatTime(totalMs),
    totalHours: totalHours,   // 🆕 レベル計算用
    daily: formatChartData(dailyMap),
    weekly: formatChartData(weeklyMap),
    monthly: formatChartData(monthlyMap),
    rank: calculateRank(totalHours),
    community: { ranking: top5Ranking, feed: recentActions },
    weeklyGoal: { currentHours: (weekMs / 3600000).toFixed(1), targetHours: weeklyGoalHours, percent: Math.min(100, ((weekMs/3600000)/weeklyGoalHours)*100).toFixed(1) },
    history: historyList.slice(0, 30),
    streak: { totalDays: uniqueDates.size, active: calculateStreak(uniqueDates) },
    tests: testProgress,
    exam: customExam,
    cheers: todayCheers,
    rivals: rivalCompare,      // 🆕
    compare: compareData,      // 🆕
    level: calculateLevel(totalHours) // 🆕 レベル＆経験値
  };

  cache.put(studentToken, JSON.stringify(resultData), 900);
  return resultData;
}

// ====================================================================
// 🆕 ヒートマップ用データ（過去120日分の日別学習時間）
// ====================================================================
function getHeatmap(token) {
  const student = findStudentByToken(token);
  if (!student) throw new Error("無効なトークンです");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logData = ss.getSheetByName('管理シート').getDataRange().getValues();
  const targetId = student.id;

  const now = new Date();
  const startMs = now.getTime() - 120 * 24 * 3600000;

  const map = {}; // yyyy-MM-dd → 分

  for (let i = 1; i < logData.length; i++) {
    let rowId = String(logData[i][IDX_LOG.ID]).trim();
    if (rowId !== targetId) continue;

    let inT = logData[i][IDX_LOG.IN];
    let outT = logData[i][IDX_LOG.OUT];
    if (!(inT instanceof Date)) continue;
    if (inT.getTime() < startMs) continue;
    if (!(outT instanceof Date)) continue;

    const diffMs = outT.getTime() - inT.getTime();
    if (diffMs <= 0) continue;

    const key = Utilities.formatDate(inT, "JST", "yyyy-MM-dd");
    map[key] = (map[key] || 0) + diffMs / 60000;
  }

  return { days: map };
}

// ====================================================================
// 🆕 時間帯別プロファイル（0〜23時の自分の学習分布）
// ====================================================================
function getHourlyProfile(token) {
  const student = findStudentByToken(token);
  if (!student) throw new Error("無効なトークンです");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logData = ss.getSheetByName('管理シート').getDataRange().getValues();
  const targetId = student.id;

  const hours = new Array(24).fill(0); // 各時間帯の累計学習分

  for (let i = 1; i < logData.length; i++) {
    let rowId = String(logData[i][IDX_LOG.ID]).trim();
    if (rowId !== targetId) continue;

    let inT = logData[i][IDX_LOG.IN];
    let outT = logData[i][IDX_LOG.OUT];
    if (!(inT instanceof Date) || !(outT instanceof Date)) continue;

    let cur = new Date(inT.getTime());
    const end = new Date(outT.getTime());
    if (end <= cur) continue;

    // 1時間ブロックごとに分配
    while (cur < end) {
      const h = cur.getHours();
      const nextHour = new Date(cur.getFullYear(), cur.getMonth(), cur.getDate(), h + 1, 0, 0);
      const sliceEnd = nextHour < end ? nextHour : end;
      const minutes = (sliceEnd.getTime() - cur.getTime()) / 60000;
      hours[h] += minutes;
      cur = sliceEnd;
    }
  }

  // ピーク時間
  let peakHour = 0;
  for (let h = 1; h < 24; h++) {
    if (hours[h] > hours[peakHour]) peakHour = h;
  }

  // 朝(5-11)/昼(11-17)/夜(17-24)/深夜(0-5) の集計でタイプ判定
  const groups = { morning: 0, afternoon: 0, evening: 0, midnight: 0 };
  for (let h = 0; h < 24; h++) {
    if (h >= 5 && h < 11) groups.morning += hours[h];
    else if (h >= 11 && h < 17) groups.afternoon += hours[h];
    else if (h >= 17 && h < 24) groups.evening += hours[h];
    else groups.midnight += hours[h];
  }

  let type = '夜型';
  let max = groups.evening;
  if (groups.morning > max) { max = groups.morning; type = '朝型'; }
  if (groups.afternoon > max) { max = groups.afternoon; type = '昼型'; }
  if (groups.midnight > max) { max = groups.midnight; type = '深夜型'; }

  return {
    hours: hours.map(m => +(m / 60).toFixed(2)), // 0時〜23時の各時間（時間単位）
    peakHour: peakHour,
    type: type
  };
}

// ====================================================================
// 🆕 スタディログ：取得（最新50件）
// ====================================================================
function getStudyLogs(token) {
  const student = findStudentByToken(token);
  if (!student) throw new Error("無効なトークンです");

  ensureSheet(SHEET_STUDY_LOG, ['ID', '投稿日時', '生徒ID', 'メッセージ', '学習時間']);
  ensureSheet(SHEET_REACTIONS, ['ログID', '生徒ID', '絵文字', '日時']);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_STUDY_LOG);
  const reactionSheet = ss.getSheetByName(SHEET_REACTIONS);

  // 名前マップ
  const profileData = ss.getSheetByName('公開プロフィール').getDataRange().getValues();
  const nickMap = {};
  for (let i = 1; i < profileData.length; i++) {
    nickMap[String(profileData[i][IDX_PROFILE.ID]).trim()] = profileData[i][IDX_PROFILE.NAME] || profileData[i][IDX_PROFILE.NICKNAME] || '学習者';
  }

  const logData = logSheet.getDataRange().getValues();
  const reactionData = reactionSheet.getDataRange().getValues();

  // リアクション集計
  const reactionMap = {}; // logId → { emoji → count }
  const myReactions = {}; // logId → [emojis]
  for (let i = 1; i < reactionData.length; i++) {
    const logId = String(reactionData[i][0]).trim();
    const userId = String(reactionData[i][1]).trim();
    const emoji = String(reactionData[i][2]).trim();
    if (!logId || !emoji) continue;
    if (!reactionMap[logId]) reactionMap[logId] = {};
    reactionMap[logId][emoji] = (reactionMap[logId][emoji] || 0) + 1;
    if (userId === student.id) {
      if (!myReactions[logId]) myReactions[logId] = [];
      myReactions[logId].push(emoji);
    }
  }

  const logs = [];
  for (let i = logData.length - 1; i >= 1 && logs.length < 50; i--) {
    const logId = String(logData[i][0]).trim();
    const postedAt = logData[i][1];
    const studentId = String(logData[i][2]).trim();
    const message = String(logData[i][3] || '').trim();
    const hours = logData[i][4];

    if (!logId || !postedAt) continue;

    logs.push({
      id: logId,
      time: postedAt instanceof Date ? Utilities.formatDate(postedAt, "JST", "MM/dd HH:mm") : String(postedAt),
      name: nickMap[studentId] || '学習者',
      isMine: studentId === student.id,
      message: message,
      hours: hours ? String(hours) : '',
      reactions: reactionMap[logId] || {},
      myReactions: myReactions[logId] || []
    });
  }

  return { logs };
}

// ====================================================================
// 🆕 スタディログ：投稿
// ====================================================================
function postStudyLog(token, message, hours) {
  const student = findStudentByToken(token);
  if (!student) throw new Error("無効なトークンです");

  message = String(message || '').trim();
  if (!message) throw new Error("メッセージを入力してください");
  if (message.length > 200) throw new Error("メッセージは200文字以内にしてください");

  // 簡易NGワード（必要に応じて拡張）
  const ngWords = ['死ね', 'バカ', 'うざい', 'きもい'];
  for (let w of ngWords) {
    if (message.indexOf(w) >= 0) throw new Error("不適切な表現が含まれています");
  }

  const sheet = ensureSheet(SHEET_STUDY_LOG, ['ID', '投稿日時', '生徒ID', 'メッセージ', '学習時間']);
  const logId = 'L' + Date.now() + '_' + Math.floor(Math.random() * 1000);
  sheet.appendRow([logId, new Date(), student.id, message, hours || '']);

  return "投稿しました";
}

// ====================================================================
// 🆕 スタディログ：リアクション（トグル：あれば削除、なければ追加）
// ====================================================================
function reactStudyLog(token, logId, emoji) {
  const student = findStudentByToken(token);
  if (!student) throw new Error("無効なトークンです");

  logId = String(logId || '').trim();
  emoji = String(emoji || '').trim();
  if (!logId || !emoji) throw new Error("パラメータが不正です");

  const allowed = ['👍', '🔥', '💪', '👏', '✨'];
  if (allowed.indexOf(emoji) < 0) throw new Error("使えない絵文字です");

  const sheet = ensureSheet(SHEET_REACTIONS, ['ログID', '生徒ID', '絵文字', '日時']);
  const data = sheet.getDataRange().getValues();

  // 既存チェック
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === logId &&
        String(data[i][1]).trim() === student.id &&
        String(data[i][2]).trim() === emoji) {
      // トグル削除
      sheet.deleteRow(i + 1);
      return "removed";
    }
  }

  sheet.appendRow([logId, student.id, emoji, new Date()]);
  return "added";
}

// ====================================================================
// 🆕 ライバル登録用：生徒一覧（自分以外）を返す
// ====================================================================
function getStudentList(token) {
  const student = findStudentByToken(token);
  if (!student) throw new Error("無効なトークンです");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const profileData = ss.getSheetByName('公開プロフィール').getDataRange().getValues();

  const list = [];
  for (let i = 1; i < profileData.length; i++) {
    const id = String(profileData[i][IDX_PROFILE.ID]).trim();
    if (!id || id === student.id) continue;
    const name = profileData[i][IDX_PROFILE.NAME] || profileData[i][IDX_PROFILE.NICKNAME] || id;
    list.push({ id, name });
  }
  list.sort((a, b) => String(a.name).localeCompare(String(b.name), 'ja'));

  return { students: list };
}

// ====================================================================
// 🆕 ライバル登録：取得
// ====================================================================
function getRivals(token) {
  const student = findStudentByToken(token);
  if (!student) throw new Error("無効なトークンです");

  const rivalIds = getRivalIdsFor(student.id);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const profileData = ss.getSheetByName('公開プロフィール').getDataRange().getValues();
  const nickMap = {};
  for (let i = 1; i < profileData.length; i++) {
    nickMap[String(profileData[i][IDX_PROFILE.ID]).trim()] = profileData[i][IDX_PROFILE.NAME] || profileData[i][IDX_PROFILE.NICKNAME] || '学習者';
  }

  return {
    rivals: rivalIds.map(id => ({ id, name: nickMap[id] || '学習者' }))
  };
}

// 内部ヘルパー：指定生徒のライバルID配列
function getRivalIdsFor(studentId) {
  ensureSheet(SHEET_RIVALS, ['生徒ID', 'ライバルIDリスト(カンマ区切り)']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = ss.getSheetByName(SHEET_RIVALS).getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === studentId) {
      const raw = String(data[i][1] || '').trim();
      if (!raw) return [];
      return raw.split(',').map(s => s.trim()).filter(s => s).slice(0, 5);
    }
  }
  return [];
}

// ====================================================================
// 🆕 ライバル登録：更新（最大5人）
// ====================================================================
function updateRivals(token, rivalIds) {
  const student = findStudentByToken(token);
  if (!student) throw new Error("無効なトークンです");

  if (!Array.isArray(rivalIds)) rivalIds = [];
  // 自分自身は除外＆重複排除＆最大5人
  const cleaned = [];
  const seen = {};
  for (let rid of rivalIds) {
    rid = String(rid).trim();
    if (!rid || rid === student.id || seen[rid]) continue;
    seen[rid] = true;
    cleaned.push(rid);
    if (cleaned.length >= 5) break;
  }

  const sheet = ensureSheet(SHEET_RIVALS, ['生徒ID', 'ライバルIDリスト(カンマ区切り)']);
  const data = sheet.getDataRange().getValues();
  let targetRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === student.id) {
      targetRow = i + 1;
      break;
    }
  }
  if (targetRow === -1) targetRow = sheet.getLastRow() + 1;

  sheet.getRange(targetRow, 1).setValue(student.id);
  sheet.getRange(targetRow, 2).setValue(cleaned.join(','));

  // 自分のキャッシュをクリア
  CacheService.getScriptCache().remove(token);

  return "ライバルを更新しました";
}

// ====================================================================
// 🆕 レベル＆経験値計算
//   レベルNから次レベルまでの必要時間 = 5 + N * 2 時間
//   Lv1: 累計0-7h / Lv2: 7-16h / Lv3: 16-27h ...
// ====================================================================
function calculateLevel(totalHours) {
  let level = 1;
  let needed = 5 + level * 2; // Lv1→Lv2は7h
  let acc = 0;

  while (totalHours - acc >= needed) {
    acc += needed;
    level++;
    needed = 5 + level * 2;
  }

  const currentExp = totalHours - acc;
  return {
    level: level,
    currentExp: +currentExp.toFixed(1),
    expForNext: +needed.toFixed(1),
    percent: Math.min(100, +(currentExp / needed * 100).toFixed(1))
  };
}

// ====================================================================
// ③ 設定保存（変更なし）
// ====================================================================
function saveStudentSettings(token, examName, examDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
// ④ 受付打刻処理（変更なし）
// ====================================================================
function processScan(studentId) {
  if (!studentId) return "エラー：IDが読み込めませんでした";
  const cleanTargetId = String(studentId).trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("管理シート");
  const now = new Date();
  const todayStr = Utilities.formatDate(now, "JST", "yyyy/MM/dd");

  const sheetMaxRow = sheet.getLastRow();

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
    const targetRow = actualLastDataRow + 1;
    sheet.getRange(targetRow, 1, 1, 2).setValues([[now, cleanTargetId]]);
    sheet.getRange(targetRow, 4).setValue(now);
    actionType = "入室";
  }

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

  if (parentEmail) {
    try {
      notifyParent(studentName, parentEmail, actionType, now, studyMs);
    } catch (e) {
      console.error('[優先度2] 保護者メール処理で例外:', e);
    }
  }

  try {
    notifyGoogleChat(studentName, actionType, now, studyMs);
  } catch (e) {
    console.error('[優先度3] Webhook処理で例外:', e);
  }

  return responseMessage;
}

// ====================================================================
// ④-2 保護者向けメール送信（変更なし）
// ====================================================================
function notifyParent(studentName, parentEmail, actionType, datetime, studyMs) {
  const timeStr = Utilities.formatDate(datetime, "JST", "yyyy年MM月dd日 HH時mm分");
  const subject = `【SSS Education】${studentName}さんが${actionType}しました`;

  let body =
    `${studentName} さんの保護者様\n\n` +
    `いつもSSS Educationをご利用いただき、ありがとうございます。\n` +
    `お子様が以下の通り自習室に${actionType}されました。\n\n` +
    `  日時：${timeStr}\n` +
    `  行動：${actionType}\n`;

  if (actionType === "退室" && studyMs > 0) {
    body += `  本日の学習時間：${formatTime(studyMs)}\n`;
  }

  body +=
    `\n` +
    `引き続き、お子様の学習を見守ってまいります。\n` +
    `何かございましたら、お気軽に塾までお問い合わせください。\n\n` +
    `──────────────────\n` +
    `SSS Education 自習室管理システム\n` +
    `※このメールは自動送信されています。返信はできませんのでご了承ください。\n` +
    `──────────────────\n`;

  try {
    MailApp.sendEmail({
      to: parentEmail,
      subject: subject,
      body: body
    });
  } catch (err) {
    console.error(`保護者メール送信失敗 (${parentEmail}):`, err);
  }
}

// ====================================================================
// ④-3 Google Chat通知（変更なし）
// ====================================================================
function notifyGoogleChat(studentName, actionType, datetime, studyMs) {
  if (!GCHAT_WEBHOOK_URL) return;

  const timeStr = Utilities.formatDate(datetime, "JST", "HH:mm");
  let message;

  if (actionType === "入室") {
    message = `🟢 *${studentName}* さんが入室しました（${timeStr}）`;
  } else if (actionType === "退室") {
    const studyTime = studyMs > 0 ? formatTime(studyMs) : "—";
    message = `🔴 *${studentName}* さんが退室しました（${timeStr} / 学習時間: ${studyTime}）`;
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
// ⑥ 生徒マイページURLのメール下書き一括作成（変更なし）
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

  const ss = SpreadsheetApp.getActiveSpreadsheet();
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

  return {
    success: successList.length > 0,
    message: message
  };
}

// ====================================================================
// ⑦ 小テスト進捗の取得（変更なし）
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
// 補助関数群（変更なし）
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

function onOpen() {
  SpreadsheetApp.getUi().createMenu('★管理メニュー').addItem('QRメールの下書き作成', 'showDraftDialog').addToUi();
}

function showDraftDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('draftDialog').setWidth(400).setHeight(320),
    'QRコード送付'
  );
}

function sendCheer(targetStudentId) {
  return "エール機能は廃止されました";
}

// ====================================================================
// 🔧 デバッグ用：今週のランキング状況を出力
// ====================================================================
function debugWeeklyRanking() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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

function testProcessScanSpeed() {
  const start = new Date().getTime();
  const result = processScan('テスト用の実在する学習者ID');
  const elapsed = new Date().getTime() - start;
  console.log(`処理時間: ${elapsed}ms / 結果: ${result}`);
}

function debugProfileColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('公開プロフィール');
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

function testEmailSend() {
  const testEmail = 'test@example.com';
  notifyParent('テスト 太郎', testEmail, '退室', new Date(), 3 * 3600000 + 25 * 60000);
  console.log(`テストメールを ${testEmail} に送信しました`);
}

function testCreateTestSheet() {
  const testName = 'テスト用_' + Utilities.formatDate(new Date(), 'JST', 'MMddHHmm');
  console.log(`テスト名「${testName}」でシート生成を試行します...`);

  const result = getOrCreateTestProgress(testName, 'TEST_ID_999');
  console.log('取得結果:', JSON.stringify(result, null, 2));

  const testSs = SpreadsheetApp.openById(TEST_SS_ID);
  const newSheet = testSs.getSheetByName(testName);
  if (newSheet) {
    console.log(`✓ シート「${testName}」が正常に作成されました（行数: ${newSheet.getLastRow()}）`);
  } else {
    console.log(`✗ シートが作成されませんでした`);
  }
}

// ====================================================================
// 🔧 デバッグ用：新機能用シートの初期化（手動実行用）
// ====================================================================
function initNewFeatureSheets() {
  ensureSheet(SHEET_STUDY_LOG, ['ID', '投稿日時', '生徒ID', 'メッセージ', '学習時間']);
  ensureSheet(SHEET_REACTIONS, ['ログID', '生徒ID', '絵文字', '日時']);
  ensureSheet(SHEET_RIVALS, ['生徒ID', 'ライバルIDリスト(カンマ区切り)']);
  console.log('✓ 新機能用シートを初期化しました');
}