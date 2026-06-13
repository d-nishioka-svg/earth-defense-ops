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
// キャッシュヘルパー（Spreadsheet読み込みを5分間キャッシュ）
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
    sheet.appendRow(['id', 'goalText', 'ifThenTrigger', 'worldStory', 'createdDate', 'status', 'timing']);
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

// ============================================================
// ミッション作成（timing追加 + プロンプト改善）
// ============================================================
function createMission(goalText, timing) {
  const systemPrompt = `あなたは並行世界のミッション設計AIです。
以下のJSON形式のみで応答してください。他の文章は不要です。

{
  "ifThenTrigger": "If-Thenトリガー（例：「お風呂が沸いたアラームが鳴ったら（If）、スクワットを20回やる（Then）」）",
  "worldStory": "なぜこの行動が並行世界を救うのかの説明（2〜3文）",
  "worldSetting": "並行世界の舞台設定（例：「第七鉱区の鍛冶工房」「並行世界Ωのクラブシーン」「星間交易船の機関室」など、キャラが住む世界の具体的な場所・雰囲気）"
}

worldStoryのルール：
① 行動（例：スクワット）→その行動が生み出す何らかのエネルギー・現象（擬似科学的でOK）→それが並行世界の危機をどう救うか、の流れで書く
② 具体的な世界観を作ること：その世界独自の名称・設定・ルールを1つ入れる（「聖なる炉」「量子の重低音」「知識の火花」など）
③ 抽象的なSF語禁止。平易な言葉で書く
④ 全体で2〜3文。テンポよく`;

  const userMsg = '継続したいこと：' + goalText + (timing ? '\nタイミング：' + timing : '');
  const result = callGemini_(systemPrompt, userMsg);
  const data = parseJson_(result);

  const sheet = getMissionsSheet_();
  const id = Utilities.getUuid();
  const now = today_();
  sheet.appendRow([id, goalText, data.ifThenTrigger, data.worldStory, now, 'active', timing || '', data.worldSetting || '']);
  cacheRemove_('missions'); // ミッション追加でキャッシュ無効化

  return { id: id, goalText: goalText, ifThenTrigger: data.ifThenTrigger, worldStory: data.worldStory, createdDate: now, status: 'active', timing: timing || '', worldSetting: data.worldSetting || '' };
}

// ============================================================
// ミッション一覧取得（キャッシュあり）
// ============================================================
function getMissions() {
  const cached = cacheGet_('missions');
  if (cached) return cached;

  const sheet = getMissionsSheet_();
  const data = sheet.getDataRange().getValues();
  const missions = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][5] === 'active') {
      missions.push({ id: data[i][0], goalText: data[i][1], ifThenTrigger: data[i][2], worldStory: data[i][3], createdDate: data[i][4], status: data[i][5], timing: data[i][6] || '', worldSetting: data[i][7] || '' });
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

  // 今日のログキャッシュを確認
  const cached = cacheGet_(cacheKey);
  if (cached) return cached;

  const logsSheet = getDailyLogsSheet_();
  const logsData = logsSheet.getDataRange().getValues();

  for (let i = 1; i < logsData.length; i++) {
    if (String(logsData[i][1]) === String(missionId) && String(logsData[i][2]) === todayStr) {
      const result = { logId: logsData[i][0], characterName: logsData[i][3], characterPersonality: logsData[i][4], result: logsData[i][5], responseText: logsData[i][6] };
      cachePut_(cacheKey, result);
      return result;
    }
  }

  const missionsSheet = getMissionsSheet_();
  const missionsData = missionsSheet.getDataRange().getValues();
  let mission = null;
  for (let i = 1; i < missionsData.length; i++) {
    if (String(missionsData[i][0]) === String(missionId)) {
      mission = { goalText: missionsData[i][1], ifThenTrigger: missionsData[i][2], worldStory: missionsData[i][3], timing: missionsData[i][6] || '', worldSetting: missionsData[i][7] || '' };
      break;
    }
  }
  if (!mission) return null;

  // 10回に1回は「特別通信官」フラグを立てる（間欠強化）
  const totalLogs = logsData.length - 1;
  const isSpecial = (totalLogs % 10 === 0);

  const systemPrompt = `あなたは毎日違うキャラクターで大輝に並行世界の指令を届けるAIです。
以下のJSON形式のみで応答してください。

{
  "characterName": "キャラの名前（例：「第七鉱区の鍛冶頭ガルドン」「並行世界Ωの元DJニャンコ」「星間交易船の機関士マリー」）",
  "characterPersonality": "性格・口調（一文。例：「荒っぽい職人言葉で話す」「だにゃ語尾の元気なネコ系」「冷静だが内心あせっている」）",
  "greeting": "大輝への指令セリフ（200字程度）"
}

greetingのルール：
① キャラが自分の世界から直接話しかけている体で書く（キャラの口調・語尾を最初から最後まで崩さない）
② 自分の世界で今何が起きているか（危機の状況）を1文で説明する
③ なぜ大輝の「${mission.goalText}」がその危機を救えるのか、世界観に沿ったメカニズムを1文で説明する（擬似科学・ファンタジーOK）
④ ミッション（${mission.timing ? mission.timing + 'になったら、' : ''}${mission.goalText}）を伝えて、最後にキャラらしい一言で締める
⑤ 説明的な地の文NG。全部キャラのセリフとして書く` + (isSpecial ? '\n\n【今日は特別通信】キャラ名に称号か肩書きをつけて、特別感を出す。' : '');

  const userMsg = `ミッション：${mission.goalText}
If-Thenトリガー：${mission.ifThenTrigger}
並行世界の危機・設定：${mission.worldStory}
舞台設定：${mission.worldSetting || '並行世界'}
タイミング：${mission.timing || 'なし'}

この世界観に合ったキャラクターを作り、指令を届けてください。キャラはこの世界観に属する住人です。`;

  const result = callGemini_(systemPrompt, userMsg);
  const charData = parseJson_(result);

  const logId = Utilities.getUuid();
  logsSheet.appendRow([logId, missionId, todayStr, charData.characterName, charData.characterPersonality, '', charData.greeting]);

  const newEntry = { logId: logId, characterName: charData.characterName, characterPersonality: charData.characterPersonality, result: '', responseText: charData.greeting, isSpecial: isSpecial };
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
・自分の世界で何が起きたか（危機が救われた具体的な変化）を1文でキャラらしく報告する
・大輝を称える（キャラの言い方で）
・キャラらしい締めの一言`;
  } else {
    userMsg = `大輝が「今日は無理」と言っている。ミッション：「${mission.goalText}」
並行世界の設定：${mission.worldStory}
200字以内で応答すること：
・今自分の世界でどんなやばいことが起きているかをキャラとして伝える（切実だが本気）
・「1回だけやってくれたら持ちこたえられる」という最小のお願いを伝える
・キャラらしい必死さで締める（押しつけではなく、頼んでいる感じ）`;
  }

  const response = callGemini_(systemPrompt, userMsg);

  if (rowIndex > 0) {
    logsSheet.getRange(rowIndex, 6).setValue(type);
    logsSheet.getRange(rowIndex, 7).setValue(response);
  }

  // キャッシュ更新（result と responseText を反映）
  const cacheKey = 'today_' + missionId + '_' + todayStr;
  const cached = cacheGet_(cacheKey);
  if (cached) { cached.result = type; cached.responseText = response; cachePut_(cacheKey, cached); }
  cacheRemove_('stats_' + missionId); // 統計キャッシュも無効化

  return { response: response };
}

// ============================================================
// 統計・ストリーク・カレンダーデータ取得（キャッシュあり）
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

  // dailyLogsから対象ミッションのログを抽出（日付→resultマップ）
  const logMap = {};
  for (let i = 1; i < logsData.length; i++) {
    if (String(logsData[i][1]) === String(missionId)) {
      logMap[logsData[i][2]] = logsData[i][5];
    }
  }

  // 今日から過去60日分のカレンダーデータ
  const calendarData = {};
  const tz = Session.getScriptTimeZone();
  for (let d = 0; d < 60; d++) {
    const dt = new Date();
    dt.setDate(dt.getDate() - d);
    const key = Utilities.formatDate(dt, tz, 'yyyy-MM-dd');
    calendarData[key] = logMap[key] || '';
  }

  // ストリーク計算（今日から遡って done が連続した日数）
  let currentStreak = 0;
  let checking = true;
  for (let d = 0; d < 365; d++) {
    const dt = new Date();
    dt.setDate(dt.getDate() - d);
    const key = Utilities.formatDate(dt, tz, 'yyyy-MM-dd');
    const result = logMap[key];
    if (d === 0 && result !== 'done') { checking = false; break; } // 今日未完了
    if (result === 'done') { if (checking) currentStreak++; }
    else if (result === 'skip') { checking = false; }
    // 未報告の日はスキップして遡る（抜けた日は途切れとみなさない）
  }

  // ベストストリーク計算
  const sortedDates = Object.keys(logMap).filter(function(k){ return logMap[k] === 'done'; }).sort();
  let bestStreak = 0, tempStreak = 0, prevDate = null;
  sortedDates.forEach(function(d) {
    if (prevDate) {
      const prev = new Date(prevDate), cur = new Date(d);
      const diff = (cur - prev) / 86400000;
      tempStreak = (diff === 1) ? tempStreak + 1 : 1;
    } else { tempStreak = 1; }
    if (tempStreak > bestStreak) bestStreak = tempStreak;
    prevDate = d;
  });

  // 累計完了数
  const totalDone = Object.values(logMap).filter(function(v){ return v === 'done'; }).length;

  // 開始からの日数
  let totalDays = 0;
  if (createdDate) {
    totalDays = Math.floor((new Date() - new Date(createdDate)) / 86400000) + 1;
  }

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
