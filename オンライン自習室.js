// ==========================================
// 【初期設定】設定済みの値です
// ==========================================

// ① スプレッドシートID
const SPREADSHEET_ID = "1QcruSLwoyPCQvCuaPK9m5Q3mFK2F2pacZPeu6VHEvps";

// ② Google ChatのWebHook URL
const WEBHOOK_URL = "https://chat.googleapis.com/v1/spaces/AAQAeqxmSII/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=0I81GRTsD6DPda0GOcQ8ZrT1RyNMnQ4e-y49blIwBUE";

// ③ 自習室の「Google Meet のURL（リンク）」
const MEET_URL = "https://meet.google.com/nih-zara-ngs";

// ④ ランキングから除外したい人の表示名
const EXCLUDE_NAMES = ["SSS自習室"];

// ==========================================
// 1. 本番用：毎日自動実行するメイン関数（管理者用）
// ==========================================
function createDailyRankingAdmin() {
  const today = new Date();
  const dateStr = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd');
  
  const targetCode = extractTargetMeetingCode(MEET_URL);
  if (!targetCode) return;
  
  const startTimeISO = Utilities.formatDate(today, 'Asia/Tokyo', "yyyy-MM-dd'T'00:00:00+09:00");
  const options = { startTime: startTimeISO, eventName: "call_ended" };
  
  let pageToken;
  const studentData = {};
  
  // ログの収集
  do {
    if (pageToken) options.pageToken = pageToken;
    let meetLogs;
    try {
      meetLogs = AdminReports.Activities.list("all", "meet", options);
    } catch (e) {
      Logger.log("エラー: " + e.message);
      return;
    }
    
    if (meetLogs && meetLogs.items && meetLogs.items.length > 0) {
      for (const item of meetLogs.items) {
        let displayName = "";
        let durationSec = 0;
        let isTargetMeeting = false;
        
        for (const param of item.events[0].parameters) {
          if (param.name === "meeting_code") {
            const currentCode = (param.value || "").replace(/-/g, "").toLowerCase();
            if (currentCode === targetCode) isTargetMeeting = true;
          }
          if (param.name === "display_name") displayName = param.value || "";
          if (param.name === "duration_seconds") durationSec = parseInt(param.intValue, 10);
        }
        
        if (isTargetMeeting && displayName && durationSec > 0 && !EXCLUDE_NAMES.includes(displayName)) {
          if (!studentData[displayName]) {
             studentData[displayName] = { morning: false, night: false, timeSec: 0 };
          }
          studentData[displayName].timeSec += durationSec;
          
          const logDate = new Date(item.id.time);
          const hour = parseInt(Utilities.formatDate(logDate, 'Asia/Tokyo', 'H'), 10);
          
          if (hour >= 4 && hour < 15) {
             studentData[displayName].morning = true;
          } else {
             studentData[displayName].night = true;
          }
        }
      }
    }
    pageToken = meetLogs.nextPageToken;
  } while (pageToken);

  // 称号の付与と振り分け
  const perfects = []; // 朝も夜も参加
  const mornings = []; // 朝のみ参加
  const nights = [];   // 夜のみ参加

  for (let name in studentData) {
    const data = studentData[name];
    const minutes = Math.floor(data.timeSec / 60);

    // 【第一部】時間帯称号の判定
    let timeTitle = "";
    if (data.morning && data.night) timeTitle = "SSS デイ＆ナイト覇者";
    else if (data.morning) timeTitle = "暁の先駆者（パイオニア）";
    else if (data.night) timeTitle = "宵の探求者（シーカー）";

    // 【第二部】SSSランク（学習時間）称号の判定
    let rankTitle = "";
    if (minutes >= 600) rankTitle = "SSS（トリプルエス）・レジェンド";
    else if (minutes >= 480) rankTitle = "SSS プラチナ・エリート";
    else if (minutes >= 360) rankTitle = "SSS ゴールド・ビルダー";
    else if (minutes >= 240) rankTitle = "SSS シルバー・チャレンジャー";
    else if (minutes >= 120) rankTitle = "SSS ブロンズ・ルーキー";

    const studentObj = { name: name, timeTitle: timeTitle, rankTitle: rankTitle, minutes: minutes };

    // カテゴリごとに配列に格納
    if (data.morning && data.night) perfects.push(studentObj);
    else if (data.morning) mornings.push(studentObj);
    else if (data.night) nights.push(studentObj);
  }

  // ▼▼▼ データ確認用のログ出力（ここに配置するのが正解です！） ▼▼▼
  console.log("【本日の生徒別 生データ】\n", JSON.stringify(studentData, null, 2));
  console.log("👑 デイ＆ナイト覇者:\n", JSON.stringify(perfects, null, 2));
  console.log("🌅 暁の先駆者:\n", JSON.stringify(mornings, null, 2));
  console.log("🌙 宵の探求者:\n", JSON.stringify(nights, null, 2));
  // ▲▲▲ ここまで ▲▲▲

  if (perfects.length === 0 && mornings.length === 0 && nights.length === 0) return;

  // スプレッドシートへの記録
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let logSheet = ss.getSheetByName("OL自習室生徒ログ");
  if (!logSheet) {
    setupSpreadsheet();
    logSheet = ss.getSheetByName("OL自習室生徒ログ");
  }
  
  // シートへ書き込み
  for (let name in studentData) {
    const data = studentData[name];
    const minutes = Math.floor(data.timeSec / 60);
    
    let tTitle = "";
    if (data.morning && data.night) tTitle = "SSS デイ＆ナイト覇者";
    else if (data.morning) tTitle = "暁の先駆者";
    else if (data.night) tTitle = "宵の探求者";
    
    let rTitle = "";
    if (minutes >= 600) rTitle = "SSS レジェンド";
    else if (minutes >= 480) rTitle = "SSS プラチナ";
    else if (minutes >= 360) rTitle = "SSS ゴールド";
    else if (minutes >= 240) rTitle = "SSS シルバー";
    else if (minutes >= 120) rTitle = "SSS ブロンズ";
    else rTitle = "（なし）"; 

    logSheet.appendRow([dateStr, name, tTitle, rTitle, minutes]);
  }

  // チャット用メッセージの作成と送信
  const messageText = buildMessage(perfects, mornings, nights);
  sendToGoogleChat(messageText);//
}

function extractTargetMeetingCode(urlOrCode) {
  if (!urlOrCode) return "";
  let code = urlOrCode.split("/").pop();
  code = code.split("?")[0];
  return code.replace(/-/g, "").toLowerCase();
}

// ==========================================
// 2. メッセージの組み立て ★時間を出すように改造！
// ==========================================
function buildMessage(perfects, mornings, nights) {
  let text = `SSS Education オンライン自習室に参加したみなさん、お疲れ様でした！\n\本日参加した皆にそれぞれ称号をお送りいたします！\n\n`;
  text += `━━━━━━━━━━━━━━━━━━━━\n`;

  // 朝も夜も参加した人
  if (perfects.length > 0) {
    text += `👑 *【SSS デイ＆ナイト覇者】*\n`;
    perfects.forEach(s => {
      text += `・${s.name} さん （${s.minutes}分）${s.rankTitle ? `\n   ┗ 🏆 *${s.rankTitle}*` : ""}\n`;
    });
    text += `\n`;
  }

  // 朝参加した人
  if (mornings.length > 0) {
    text += `🌅 *【暁の先駆者】*（朝活☀️）\n`;
    mornings.forEach(s => {
      text += `・${s.name} さん （${s.minutes}分）${s.rankTitle ? `\n   ┗ 🏆 *${s.rankTitle}*` : ""}\n`;
    });
    text += `\n`;
  }

  // 夜参加した人
  if (nights.length > 0) {
    text += `🌙 *【宵の探求者】*（夜の自己投資👍）\n`;
    nights.forEach(s => {
      text += `・${s.name} さん （${s.minutes}分）${s.rankTitle ? `\n   ┗ 🏆 *${s.rankTitle}*` : ""}\n`;
    });
    text += `\n`;
  }

  text += `━━━━━━━━━━━━━━━━━━━━\n\n`;
  text += `今日もそれぞれのペースでよく頑張りました。\n明日も一緒に頑張りましょう！\n\n継続は力なり`;

  return text;
}

function sendToGoogleChat(text) {
  const payload = JSON.stringify({ "text": text });
  const options = { "method": "POST", "contentType": "application/json", "payload": payload, "muteHttpExceptions": true };
  UrlFetchApp.fetch(WEBHOOK_URL, options);
}

// ==========================================
// 3. 設定関数：新しいシートを自動作成
// ==========================================
function setupSpreadsheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let logSheet = ss.getSheetByName("OL自習室生徒ログ");
  if (!logSheet) {
    logSheet = ss.insertSheet("OL自習室生徒ログ");
    logSheet.appendRow(["日付", "名前", "時間帯称号", "時間別ランク", "学習時間(管理用)"]);
    logSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#f3f3f3");
    Logger.log("「OL自習室生徒ログ」シートを新規作成しました！");
  } else {
    Logger.log("「OL自習室生徒ログ」シートは既に存在します。");
  }
}

// ==========================================
// 4. テスト用：今すぐチャットにテスト通知を送る関数 ★テストデータも時間対応！
// ==========================================
function testGoogleChatNotification() {
  const perfects = [
    { name: "テスト生徒A", rankTitle: "SSS ゴールド・ビルダー", minutes: 380 },
    { name: "テスト生徒B", rankTitle: "SSS ブロンズ・ルーキー", minutes: 140 }
  ];
  const mornings = [
    { name: "テスト生徒C", rankTitle: "", minutes: 45 }
  ];
  const nights = [
    { name: "テスト生徒D", rankTitle: "SSS（トリプルエス）・レジェンド", minutes: 620 }
  ];

  const testMessage = buildMessage(perfects, mornings, nights) + "\n\n※これはシステム構築の動作確認テスト送信です。";
  sendToGoogleChat(testMessage);
  Logger.log("テスト送信処理が完了しました！");
}