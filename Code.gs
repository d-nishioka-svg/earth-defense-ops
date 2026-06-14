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
    generationConfig: { temperature: 1.0 }
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
    // id, goalText, ifThenTrigger, worldStory, createdDate, status, timing, worldSetting, notifyTime
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

// notifyTime "21:30" → "21時30分になったら"
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
  ScriptApp.newTrigger('checkAndNotify')
    .timeBased()
    .everyHours(1)
    .create();
}

// 毎時実行：通知時刻に一致するミッションにメール送信
function checkAndNotify() {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const currentHour = parseInt(Utilities.formatDate(now, tz, 'HH'), 10);
  const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');

  const missionsSheet = getMissionsSheet_();
  const missionsData = missionsSheet.getDataRange().getValues();
  const logsSheet = getDailyLogsSheet_();
  const logsData = logsSheet.getDataRange().getValues();

  // 今日すでにログがあるミッションIDのセット
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

    // 今日すでに通知済みならスキップ
    const notifyKey = 'notify_' + missionId + '_' + todayStr;
    if (props.getProperty(notifyKey)) continue;

    const goalText = String(missionsData[i][1]);
    const ifThenTrigger = String(missionsData[i][2]);

    try {
      MailApp.sendEmail({
        to: email,
        subject: '🌍 地球防衛OPS ｜ ミッション通知',
        body: [
          '大輝へ',
          '',
          '時間だ。ミッションを確認しろ。',
          '',
          '【ミッション】' + goalText,
          '【If-Then】' + ifThenTrigger,
          '',
          'アプリを開いて今日のキャラクターの指令を受け取れ。',
          '',
          '─ 地球防衛オペレーション',
        ].join('\n'),
      });
      props.setProperty(notifyKey, '1');
    } catch(e) {
      // メール送信失敗は無視（ログは残らないが処理を止めない）
    }
  }
}

// ============================================================
// ミッション作成
// ============================================================
function createMission(goalText, notifyTime) {
  const timing = timingFromNotifyTime_(notifyTime);

  const systemPrompt = `あなたは並行世界のミッション設計AIです。
以下のJSON形式のみで応答してください。他の文章は不要です。

{
  "ifThenTrigger": "If-Thenトリガー（例：「21時になったら（If）、スクワットを20回やる（Then）」）",
  "worldStory": "なぜこの行動が並行世界を救うのかの説明（2〜3文）",
  "worldSetting": "並行世界の舞台設定（例：「第七鉱区の鍛冶工房」「銀河系外縁の宇宙ステーション」「地下水路都市の発電所」）"
}

worldStoryのルール：
① 行動（例：スクワット）→その行動が生み出すエネルギー・現象（擬似科学的でOK）→それが並行世界の危機をどう救うか、の流れで書く
② その世界独自の名称・設定を1つ入れる（「聖なる炉」「量子の重低音」「記憶の結晶」など）
③ 実在する場所・人・もの・チーム・ブランドを最低1つ入れること（例：エベレスト、スタバ、バルセロナ、NASA、阪神タイガース、富士山など世界中のなんでもOK）。知っているものが出ると臨場感が生まれる
④ 平易な言葉で。抽象的なSF語禁止
⑤ 全体で2〜3文`;

  const userMsg = '継続したいこと：' + goalText + (timing ? '\nタイミング：' + timing : '');
  const result = callGemini_(systemPrompt, userMsg);
  const data = parseJson_(result);

  const sheet = getMissionsSheet_();
  const id = Utilities.getUuid();
  const now = today_();
  sheet.appendRow([id, goalText, data.ifThenTrigger, data.worldStory, now, 'active', timing, data.worldSetting || '', notifyTime || '']);
  cacheRemove_('missions');

  if (notifyTime) {
    try { ensureNotifyTrigger_(); } catch(e) {}
  }

  return {
    id: id, goalText: goalText, ifThenTrigger: data.ifThenTrigger,
    worldStory: data.worldStory, createdDate: now, status: 'active',
    timing: timing, worldSetting: data.worldSetting || '', notifyTime: notifyTime || ''
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
      const result = {
        logId: logsData[i][0], characterName: logsData[i][3],
        characterPersonality: logsData[i][4], result: logsData[i][5], responseText: logsData[i][6]
      };
      cachePut_(cacheKey, result);
      return result;
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
  "characterName": "キャラの名前（例：「第七鉱区の鍛冶頭ガルドン」「銀河DJニャンコ」「砂漠の交易商マリア」）",
  "characterPersonality": "性格・口調（一文。例：「荒っぽい職人言葉」「だにゃ語尾の元気系」「冷静だが内心あせっている」）",
  "greeting": "大輝への指令セリフ（200字程度）"
}

greetingのルール：
① キャラが自分の世界から直接話しかけている体で書く（口調・語尾を最初から最後まで崩さない）
② 自分の世界で今何が起きているか（危機の状況）を1文で説明する
③ なぜ大輝の「${mission.goalText}」がその危機を救えるのか、世界観に沿ったメカニズムを1文で説明する
④ 実在する場所・もの・チーム・ブランドを最低1つ使うこと（エベレスト、NASA、スタバ、阪神タイガース、富士山など世界中なんでもOK）。知っているものが出ると面白い
⑤ ミッション（${mission.timing ? mission.timing + '、' : ''}${mission.goalText}）を伝えて、最後にキャラらしい一言で締める
⑥ 全部キャラのセリフ。説明文NG` + (isSpecial ? '\n\n【今日は特別通信】キャラ名に称号か肩書きをつけて特別感を出す。' : '');

  const userMsg = `ミッション：${mission.goalText}
If-Thenトリガー：${mission.ifThenTrigger}
並行世界の危機・設定：${mission.worldStory}
舞台設定：${mission.worldSetting || '並行世界'}
タイミング：${mission.timing || 'なし'}

この世界観に合ったキャラクターを作り、指令を届けてください。`;

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

  const systemPrompt = `あなたは並行世界から通信してきた「${charName}」というキャラクターです。
口調・性格：${charPersonality}
そのキャラを最初から最後まで崩さず、自分の世界の出来事として話してください。`;

  let userMsg;
  if (type === 'done') {
    userMsg = `大輝が今日「${mission.goalText}」を完了した！
並行世界の設定：${mission.worldStory}
150字以内でリアクションすること：
・自分の世界で何が救われたか、実在する場所・もの（エベレスト、スタバ、阪神タイガース、富士山など世界中なんでもOK）を1つ使って具体的に伝える
・大輝を称える（キャラの言い方で）
・キャラらしい締めの一言`;
  } else {
    userMsg = `大輝が「今日は無理」と言っている。ミッション：「${mission.goalText}」
並行世界の設定：${mission.worldStory}
200字以内で応答すること：
・今自分の世界で何がやばいか、実在する場所・もの（エベレスト、スタバ、阪神タイガース、富士山など世界中なんでもOK）を1つ使って具体的に伝える
・「1回だけやってくれたら持ちこたえられる」という最小のお願いを伝える
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
// 統計・ストリーク・カレンダーデータ取得
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
    if (String(missionsData[i][0]) === String(missionId)) {
      createdDate = missionsData[i][4]; break;
    }
  }

  const logMap = {};
  for (let i = 1; i < logsData.length; i++) {
    if (String(logsData[i][1]) === String(missionId)) {
      logMap[logsData[i][2]] = logsData[i][5];
    }
  }

  const calendarData = {};
  const tz = Session.getScriptTimeZone();
  for (let d = 0; d < 60; d++) {
    const dt = new Date();
    dt.setDate(dt.getDate() - d);
    const key = Utilities.formatDate(dt, tz, 'yyyy-MM-dd');
    calendarData[key] = logMap[key] || '';
  }

  let currentStreak = 0;
  let checking = true;
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
      const diff = (new Date(d) - new Date(prevDate)) / 86400000;
      tempStreak = (diff === 1) ? tempStreak + 1 : 1;
    } else { tempStreak = 1; }
    if (tempStreak > bestStreak) bestStreak = tempStreak;
    prevDate = d;
  });

  const totalDone = Object.values(logMap).filter(function(v){ return v === 'done'; }).length;
  let totalDays = 0;
  if (createdDate) {
    totalDays = Math.floor((new Date() - new Date(createdDate)) / 86400000) + 1;
  }

  const stats = {
    currentStreak: currentStreak, bestStreak: bestStreak,
    totalDone: totalDone, totalDays: totalDays,
    calendarData: calendarData, createdDate: String(createdDate)
  };
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
