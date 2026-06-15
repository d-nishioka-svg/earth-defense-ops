function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('地球防衛OPS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// Gemini API ヘルパー
// ============================================================
function callGemini_(systemPrompt, userMessage) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;

  const payload = {
    system_instruction: { parts: [{ text: systemPrompt }] },
    contents: [{ role: 'user', parts: [{ text: userMessage }] }],
    generationConfig: { temperature: 1.2 }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());
  if (json.error) throw new Error('Gemini API エラー: ' + json.error.message);
  return json.candidates[0].content.parts[0].text;
}

function parseJson_(text) {
  return JSON.parse(text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim());
}

// ============================================================
// キャッシュヘルパー
// ============================================================
function cacheGet_(key) {
  try {
    const v = CacheService.getScriptCache().get(key);
    return v ? JSON.parse(v) : null;
  } catch(e) { return null; }
}
function cachePut_(key, val) {
  try { CacheService.getScriptCache().put(key, JSON.stringify(val), 300); } catch(e) {}
}
function cacheRemove_(key) {
  try { CacheService.getScriptCache().remove(key); } catch(e) {}
}

// ============================================================
// Spreadsheet ヘルパー
// ============================================================
function getSpreadsheet_() {
  const ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return SpreadsheetApp.openById(ssId);
}

function getMissionsSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName('missions');
  if (!sheet) {
    sheet = ss.insertSheet('missions');
    sheet.appendRow(['id', 'goalText', 'ifThenTrigger', 'worldStory', 'createdDate', 'status', 'timing', 'worldSetting', 'notifyTime']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getDailyLogsSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName('dailyLogs');
  if (!sheet) {
    sheet = ss.insertSheet('dailyLogs');
    sheet.appendRow(['id', 'missionId', 'date', 'characterName', 'characterPersonality', 'result', 'responseText']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function today_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function timingFromNotifyTime_(notifyTime) {
  if (!notifyTime) return '';
  const parts = notifyTime.split(':');
  const h = parseInt(parts[0], 10);
  const m = parseInt(parts[1] || '0', 10);
  return m === 0 ? h + '時になったら' : h + '時' + m + '分になったら';
}

// ============================================================
// 通知トリガー管理
// ============================================================
function ensureNotifyTrigger_() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkAndNotify') return;
  }
  ScriptApp.newTrigger('checkAndNotify').timeBased().everyHours(1).create();
}

function checkAndNotify() {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const currentHour = parseInt(Utilities.formatDate(now, tz, 'HH'), 10);
  const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');

  const missionsSheet = getMissionsSheet_();
  const missionsData = missionsSheet.getDataRange().getValues();
  const logsSheet = getDailyLogsSheet_();
  const logsData = logsSheet.getDataRange().getValues();

  const loggedToday = new Set();
  for (let i = 1; i < logsData.length; i++) {
    if (String(logsData[i][2]) === todayStr && logsData[i][5] !== '') {
      loggedToday.add(String(logsData[i][1]));
    }
  }

  const props = PropertiesService.getScriptProperties();
  const email = Session.getEffectiveUser().getEmail();

  for (let i = 1; i < missionsData.length; i++) {
    if (missionsData[i][5] !== 'active') continue;
    const notifyTime = String(missionsData[i][8] || '');
    if (!notifyTime) continue;

    const notifyHour = parseInt(notifyTime.split(':')[0], 10);
    if (notifyHour !== currentHour) continue;

    const missionId = String(missionsData[i][0]);
    if (loggedToday.has(missionId)) continue;

    const notifyKey = 'notify_' + missionId + '_' + todayStr;
    if (props.getProperty(notifyKey)) continue;

    const goalText = String(missionsData[i][1]);
    const ifThenTrigger = String(missionsData[i][2]);

    try {
      MailApp.sendEmail({
        to: email,
        subject: '🌍 地球防衛OPS ｜ ミッション通知',
        body: ['大輝へ', '', '時間だ。ミッションを確認しろ。', '',
          '【ミッション】' + goalText, '【If-Then】' + ifThenTrigger, '',
          'アプリを開いて今日のキャラクターの指令を受け取れ。', '', '─ 地球防衛オペレーション'].join('\n'),
      });
      props.setProperty(notifyKey, '1');
    } catch(e) {}
  }
}

// ============================================================
// ミッション生成（保存なし）- step1: ユーザーに3択を提示
// ============================================================
function generateMissionOptions(goalText, notifyTime) {
  const timing = timingFromNotifyTime_(notifyTime);

  const systemPrompt = `あなたは並行世界から大輝に通信を送ってきた存在です。
以下のJSON形式のみで応答してください。他の文章は一切不要です。

{
  "characterName": "キャラの名前",
  "characterPersonality": "口調・性格（一文）",
  "characterIntro": "キャラの自己紹介セリフ（300字程度）",
  "worldSetting": "並行世界の舞台（例：暗黒銀河の古代図書館）",
  "missionText": "ミッション説明＋ペナルティ（350字程度）",
  "ifThenOptions": [
    {"trigger": "トリガー1", "action": "行動1"},
    {"trigger": "トリガー2", "action": "行動2"},
    {"trigger": "トリガー3", "action": "行動3"}
  ]
}

━━ キャラ生成ルール ━━
・既存のSFテンプレート（ロボット・宇宙人・AI・魔法使い）は絶対禁止
・全く異なる2つの概念をサイバー/ファンタジー要素で掛け合わせた、その場で新たに完全自動生成したキャラにすること
 例：「暗黒銀河の読書カマキリ」「並行世界βのサイバーお遍路さん」「量子世界のハードコア盆栽職人」「地下帝国の酔っ払い数学者」「亜空間の元プロ雀士ロボネコ」
・characterIntroでは：①職業 ②現在の状況 ③なぜ大輝に通信してきたか、をキャラ全開の口調で語らせる
・口調・語尾は最初から最後まで完全に崩さないこと（例：だにゃ語尾なら全てだにゃで終わる）

━━ If-Thenトリガーのルール（最重要） ━━
・「大輝が1日に必ず行う無意識な生活行動」を引き金にすること
・以下のレベルの具体性が必要：
  ✓「夜、スマホを充電器に挿した瞬間に」
  ✓「お風呂が沸いたアラームが鳴ったら」
  ✓「朝、コーヒーをマグカップに注ぎ終えたら」
  ✓「YouTubeを開いて最初の広告が流れ始めたら」
  ✓「トイレに座って落ち着いたら」
  ✗「毎日やる」「時間を決めて」→ NG（抽象的すぎる）
・3つは朝・夜・お風呂などバラバラのシーンから選ぶ
・actionには具体的な量を入れる（○回/○分/○ページ）

━━ missionTextのルール ━━
【前半：行動した場合（200字）】
 その行動（例：スクワット）が生み出すエネルギー・現象が、このキャラの並行世界でどう機能するかを、壮大かつバカバカしいほど大真面目に説明する
 世界中の実在する場所・もの・チーム・ブランド（エベレスト、NASA、スタバ、阪神タイガース、任天堂など何でもOK）を最低1つ入れること

【後半：サボった場合のペナルティ（150字）】
 このキャラの世界に何が起きるかをリアルに描写し「俺がやらないと世界が終わる」という変な使命感（損失回避バイアス）を強力に植え付けること`;

  const userMsg = '継続したいこと：' + goalText + (timing ? '\nタイミング：' + timing : '');
  const result = callGemini_(systemPrompt, userMsg);
  return parseJson_(result);
}

// ============================================================
// ミッション保存 - step2: ユーザーが選んだプランを確定保存
// ============================================================
function saveMission(goalText, notifyTime, selectedIndex, generatedData) {
  const timing = timingFromNotifyTime_(notifyTime);
  const option = generatedData.ifThenOptions[selectedIndex];
  const ifThenTrigger = option.trigger + ' → ' + option.action;

  const sheet = getMissionsSheet_();
  const id = Utilities.getUuid();
  const now = today_();
  sheet.appendRow([
    id, goalText, ifThenTrigger, generatedData.missionText,
    now, 'active', timing, generatedData.worldSetting || '', notifyTime || ''
  ]);
  cacheRemove_('missions');

  // 今日のログとして初回キャラを保存（当日はそのままこのキャラが表示される）
  const logsSheet = getDailyLogsSheet_();
  const logId = Utilities.getUuid();
  logsSheet.appendRow([
    logId, id, now,
    generatedData.characterName, generatedData.characterPersonality,
    '', generatedData.characterIntro
  ]);

  if (notifyTime) {
    try { ensureNotifyTrigger_(); } catch(e) {}
  }

  return {
    id: id, goalText: goalText, ifThenTrigger: ifThenTrigger,
    worldStory: generatedData.missionText, createdDate: now, status: 'active',
    timing: timing, worldSetting: generatedData.worldSetting || '', notifyTime: notifyTime || ''
  };
}

// ============================================================
// ミッション一覧取得
// ============================================================
function getMissions() {
  const cached = cacheGet_('missions');
  if (cached) return cached;

  const sheet = getMissionsSheet_();
  const data = sheet.getDataRange().getValues();
  const missions = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][5] === 'active') {
      missions.push({
        id: data[i][0], goalText: data[i][1], ifThenTrigger: data[i][2],
        worldStory: data[i][3], createdDate: data[i][4], status: data[i][5],
        timing: data[i][6] || '', worldSetting: data[i][7] || '', notifyTime: data[i][8] || ''
      });
    }
  }
  cachePut_('missions', missions);
  return missions;
}

// ============================================================
// 今日のミッション（キャラ生成 or 取得）
// ============================================================
function getTodayMission(missionId) {
  const todayStr = today_();
  const cacheKey = 'today_' + missionId + '_' + todayStr;

  const cached = cacheGet_(cacheKey);
  if (cached) return cached;

  const logsSheet = getDailyLogsSheet_();
  const logsData = logsSheet.getDataRange().getValues();

  for (let i = 1; i < logsData.length; i++) {
    if (String(logsData[i][1]) === String(missionId) && String(logsData[i][2]) === todayStr) {
      const entry = {
        logId: logsData[i][0], characterName: logsData[i][3],
        characterPersonality: logsData[i][4], result: logsData[i][5], responseText: logsData[i][6]
      };
      cachePut_(cacheKey, entry);
      return entry;
    }
  }

  const missionsSheet = getMissionsSheet_();
  const missionsData = missionsSheet.getDataRange().getValues();
  let mission = null;
  for (let i = 1; i < missionsData.length; i++) {
    if (String(missionsData[i][0]) === String(missionId)) {
      mission = {
        goalText: missionsData[i][1], ifThenTrigger: missionsData[i][2],
        worldStory: missionsData[i][3], timing: missionsData[i][6] || '',
        worldSetting: missionsData[i][7] || '', notifyTime: missionsData[i][8] || ''
      };
      break;
    }
  }
  if (!mission) return null;

  const totalLogs = logsData.length - 1;
  const isSpecial = (totalLogs % 10 === 0);

  const systemPrompt = `あなたは毎日違うキャラクターで大輝に並行世界の指令を届けるAIです。
以下のJSON形式のみで応答してください。

{
  "characterName": "キャラ名",
  "characterPersonality": "口調・性格（一文）",
  "greeting": "大輝への指令セリフ（200字程度）"
}

greetingのルール：
① キャラが自分の世界から直接話しかけている体で書く（口調・語尾を最初から最後まで崩さない）
② 既存のSFテンプレート禁止。2つの異なる概念を掛け合わせた新キャラ
③ 自分の世界で今何が起きているか（危機の状況）を1文で説明する
④ なぜ大輝の「${mission.goalText}」がその危機を救えるのかを1文で説明する
⑤ 世界中の実在する場所・もの（エベレスト、NASA、スタバ、阪神タイガースなど）を最低1つ使う
⑥ ミッション（${mission.timing ? mission.timing + '、' : ''}${mission.goalText}）を伝えて、キャラらしい一言で締める
⑦ 全部キャラのセリフ。説明文NG` + (isSpecial ? '\n\n【今日は特別通信】キャラ名に称号か肩書きをつけて特別感を出す。' : '');

  const userMsg = `ミッション：${mission.goalText}
If-Thenトリガー：${mission.ifThenTrigger}
並行世界の危機・設定：${mission.worldStory}
舞台設定：${mission.worldSetting || '並行世界'}
タイミング：${mission.timing || 'なし'}`;

  const result = callGemini_(systemPrompt, userMsg);
  const charData = parseJson_(result);

  const logId = Utilities.getUuid();
  logsSheet.appendRow([logId, missionId, todayStr, charData.characterName, charData.characterPersonality, '', charData.greeting]);

  const newEntry = {
    logId: logId, characterName: charData.characterName,
    characterPersonality: charData.characterPersonality,
    result: '', responseText: charData.greeting, isSpecial: isSpecial
  };
  cachePut_('today_' + missionId + '_' + todayStr, newEntry);
  return newEntry;
}

// ============================================================
// 任務完了 / エネルギー不足
// ============================================================
function reportDone(missionId) { return report_(missionId, 'done'); }
function reportSkip(missionId) { return report_(missionId, 'skip'); }

function report_(missionId, type) {
  const todayStr = today_();
  const logsSheet = getDailyLogsSheet_();
  const logsData = logsSheet.getDataRange().getValues();

  let rowIndex = -1, charName = '', charPersonality = '';
  for (let i = 1; i < logsData.length; i++) {
    if (String(logsData[i][1]) === String(missionId) && String(logsData[i][2]) === todayStr) {
      rowIndex = i + 1; charName = logsData[i][3]; charPersonality = logsData[i][4]; break;
    }
  }

  const missionsSheet = getMissionsSheet_();
  const missionsData = missionsSheet.getDataRange().getValues();
  let mission = null;
  for (let i = 1; i < missionsData.length; i++) {
    if (String(missionsData[i][0]) === String(missionId)) {
      mission = { goalText: missionsData[i][1], worldStory: missionsData[i][3] }; break;
    }
  }

  const systemPrompt = `あなたは並行世界から通信してきた「${charName}」です。
口調・性格：${charPersonality}
そのキャラを最初から最後まで崩さず、自分の世界の出来事として話してください。`;

  let userMsg;
  if (type === 'done') {
    userMsg = `大輝が今日「${mission.goalText}」を完了した！
並行世界の設定：${mission.worldStory}
150字以内でリアクション：
・自分の世界で何が救われたか、実在する場所・もの（エベレスト、スタバ、阪神タイガース等、世界中なんでもOK）を1つ使って具体的に伝える
・大輝を称える（キャラの言い方で）
・キャラらしい締めの一言`;
  } else {
    userMsg = `大輝が「今日は無理」と言っている。ミッション：「${mission.goalText}」
並行世界の設定：${mission.worldStory}
200字以内で応答：
・今自分の世界で何がやばいか、実在する場所・もの（エベレスト、スタバ等）を1つ使って具体的に伝える
・「1回だけやってくれたら持ちこたえられる」という最小のお願い
・キャラらしい必死さで締める`;
  }

  const response = callGemini_(systemPrompt, userMsg);

  if (rowIndex > 0) {
    logsSheet.getRange(rowIndex, 6).setValue(type);
    logsSheet.getRange(rowIndex, 7).setValue(response);
  }

  const cacheKey = 'today_' + missionId + '_' + todayStr;
  const cached = cacheGet_(cacheKey);
  if (cached) { cached.result = type; cached.responseText = response; cachePut_(cacheKey, cached); }
  cacheRemove_('stats_' + missionId);

  return { response: response };
}

// ============================================================
// 統計・ストリーク・カレンダーデータ
// ============================================================
function getStats(missionId) {
  const statsCacheKey = 'stats_' + missionId;
  const cached = cacheGet_(statsCacheKey);
  if (cached) return cached;

  const logsSheet = getDailyLogsSheet_();
  const logsData = logsSheet.getDataRange().getValues();
  const missionsSheet = getMissionsSheet_();
  const missionsData = missionsSheet.getDataRange().getValues();

  let createdDate = '';
  for (let i = 1; i < missionsData.length; i++) {
    if (String(missionsData[i][0]) === String(missionId)) { createdDate = missionsData[i][4]; break; }
  }

  const logMap = {};
  for (let i = 1; i < logsData.length; i++) {
    if (String(logsData[i][1]) === String(missionId)) logMap[logsData[i][2]] = logsData[i][5];
  }

  const calendarData = {};
  const tz = Session.getScriptTimeZone();
  for (let d = 0; d < 60; d++) {
    const dt = new Date();
    dt.setDate(dt.getDate() - d);
    const key = Utilities.formatDate(dt, tz, 'yyyy-MM-dd');
    calendarData[key] = logMap[key] || '';
  }

  let currentStreak = 0, checking = true;
  for (let d = 0; d < 365; d++) {
    const dt = new Date();
    dt.setDate(dt.getDate() - d);
    const key = Utilities.formatDate(dt, tz, 'yyyy-MM-dd');
    const result = logMap[key];
    if (d === 0 && result !== 'done') { checking = false; break; }
    if (result === 'done') { if (checking) currentStreak++; }
    else if (result === 'skip') { checking = false; }
  }

  const sortedDates = Object.keys(logMap).filter(function(k){ return logMap[k] === 'done'; }).sort();
  let bestStreak = 0, tempStreak = 0, prevDate = null;
  sortedDates.forEach(function(d) {
    if (prevDate) {
      tempStreak = (new Date(d) - new Date(prevDate)) / 86400000 === 1 ? tempStreak + 1 : 1;
    } else { tempStreak = 1; }
    if (tempStreak > bestStreak) bestStreak = tempStreak;
    prevDate = d;
  });

  const totalDone = Object.values(logMap).filter(function(v){ return v === 'done'; }).length;
  let totalDays = createdDate ? Math.floor((new Date() - new Date(createdDate)) / 86400000) + 1 : 0;

  const stats = { currentStreak: currentStreak, bestStreak: bestStreak, totalDone: totalDone, totalDays: totalDays, calendarData: calendarData, createdDate: String(createdDate) };
  cachePut_(statsCacheKey, stats);
  return stats;
}

// ============================================================
// アーカイブ取得
// ============================================================
function getArchive(missionId) {
  const logsSheet = getDailyLogsSheet_();
  const logsData = logsSheet.getDataRange().getValues();
  const logs = [];
  for (let i = 1; i < logsData.length; i++) {
    if (String(logsData[i][1]) === String(missionId)) {
      logs.push({ date: logsData[i][2], characterName: logsData[i][3], result: logsData[i][5], responseText: logsData[i][6] });
    }
  }
  return logs.sort(function(a, b){ return b.date > a.date ? 1 : -1; });
}

// ============================================================
// ミッション削除（アーカイブ化）
// ============================================================
function deleteMission(missionId) {
  const sheet = getMissionsSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(missionId)) {
      sheet.getRange(i + 1, 6).setValue('archived');
      cacheRemove_('missions');
      return true;
    }
  }
  return false;
}
