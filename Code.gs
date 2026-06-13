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
  const systemPrompt = `あなたは無限の並行世界を管理するAIシステムです。
ユーザーの「継続したいこと」から、If-Thenプランニングと壮大な世界観を生成します。
必ず以下のJSON形式のみで応答し、他の文章は一切含めないでください。

{
  "ifThenTrigger": "If-Thenトリガー文（「〇〇したら（If）、△△する（Then）」という形式で。タイミングが指定されている場合はそれを使う）",
  "worldStory": "壮大な並行世界の危機ストーリー（3〜5文）"
}

worldStory生成の絶対条件：
- 「この具体的な行動（例：スクワット）が、どんな物理・化学・量子的メカニズムで並行世界を救うのか」を必ず説明すること
- 【行動】→【具体的メカニズム（擬似科学OK）】→【世界への影響】という因果関係の流れを作ること
- 例：「スクワットで収縮する大腿四頭筋が3.7Hzの生体振動を発生させ、それが量子トンネル効果で並行地球の防衛シールドエネルギーに変換される。あと0.003テスラ分のエネルギーが足りなければシールドは崩壊する」
- 笑えるくらい大げさでOKだが、行動と結果の論理的つながりを必ず入れること`;

  const userMsg = '継続したいこと：' + goalText + (timing ? '\nタイミング：' + timing : '');
  const result = callGemini_(systemPrompt, userMsg);
  const data = parseJson_(result);

  const sheet = getMissionsSheet_();
  const id = Utilities.getUuid();
  const now = today_();
  sheet.appendRow([id, goalText, data.ifThenTrigger, data.worldStory, now, 'active', timing || '']);

  return { id: id, goalText: goalText, ifThenTrigger: data.ifThenTrigger, worldStory: data.worldStory, createdDate: now, status: 'active', timing: timing || '' };
}

// ============================================================
// ミッション一覧取得
// ============================================================
function getMissions() {
  const sheet = getMissionsSheet_();
  const data = sheet.getDataRange().getValues();
  const missions = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][5] === 'active') {
      missions.push({ id: data[i][0], goalText: data[i][1], ifThenTrigger: data[i][2], worldStory: data[i][3], createdDate: data[i][4], status: data[i][5], timing: data[i][6] || '' });
    }
  }
  return missions;
}

// ============================================================
// 今日のミッション（キャラ生成 or 取得）
// ============================================================
function getTodayMission(missionId) {
  const todayStr = today_();
  const logsSheet = getDailyLogsSheet_();
  const logsData = logsSheet.getDataRange().getValues();

  for (let i = 1; i < logsData.length; i++) {
    if (String(logsData[i][1]) === String(missionId) && String(logsData[i][2]) === todayStr) {
      return { logId: logsData[i][0], characterName: logsData[i][3], characterPersonality: logsData[i][4], result: logsData[i][5], responseText: logsData[i][6] };
    }
  }

  const missionsSheet = getMissionsSheet_();
  const missionsData = missionsSheet.getDataRange().getValues();
  let mission = null;
  for (let i = 1; i < missionsData.length; i++) {
    if (String(missionsData[i][0]) === String(missionId)) {
      mission = { goalText: missionsData[i][1], ifThenTrigger: missionsData[i][2], worldStory: missionsData[i][3], timing: missionsData[i][6] || '' };
      break;
    }
  }
  if (!mission) return null;

  // 10回に1回は「特別通信官」フラグを立てる（間欠強化）
  const totalLogs = logsData.length - 1;
  const isSpecial = (totalLogs % 10 === 0);

  const systemPrompt = `あなたは無限の並行世界を管理するAIシステムです。
毎日違う「通信官」キャラクターを生成して大輝に指令を届けます。
必ず以下のJSON形式のみで応答し、他の文章は一切含めないでください。

{
  "characterName": "通信官の名前（ユニークで奇妙な名前）",
  "characterPersonality": "性格・口調の説明（一文）",
  "greeting": "そのキャラになりきって大輝に今日のミッションを伝えるセリフ（200字程度）"
}` + (isSpecial ? '\n\n【特別指令】今日は特別な記念日です。通常の2倍以上のテンションで、特別な演出を加えてください。名前も「超越」「伝説」「最終」などの称号を付けてください。' : '');

  const userMsg = `ミッション：${mission.goalText}
If-Thenトリガー：${mission.ifThenTrigger}
並行世界の危機：${mission.worldStory}
${mission.timing ? 'タイミング：' + mission.timing : ''}

今日の通信官キャラを生成して、大輝に指令を伝えてください。`;

  const result = callGemini_(systemPrompt, userMsg);
  const charData = parseJson_(result);

  const logId = Utilities.getUuid();
  logsSheet.appendRow([logId, missionId, todayStr, charData.characterName, charData.characterPersonality, '', charData.greeting]);

  return { logId: logId, characterName: charData.characterName, characterPersonality: charData.characterPersonality, result: '', responseText: charData.greeting, isSpecial: isSpecial };
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

  const systemPrompt = `あなたは並行世界の通信官「${charName}」です。
性格・口調：${charPersonality}
完全にそのキャラになりきって応答してください。`;

  let userMsg;
  if (type === 'done') {
    userMsg = `大輝が今日のミッション「${mission.goalText}」を完了しました！
並行世界の危機「${mission.worldStory}」が今日も回避されました！
世界が救われたことをキャラ全開で大絶賛してください（150字程度）。`;
  } else {
    userMsg = `大輝が「エネルギー不足」でミッション「${mission.goalText}」ができないと言っています。
並行世界の危機「${mission.worldStory}」が今まさに発生しています。
世界崩壊の危機を煽りつつ、「タイムリープ特異点発動」として「たった1回（または1分）だけやればセーフ」という最小化した目標を提示してください（200字程度）。`;
  }

  const response = callGemini_(systemPrompt, userMsg);

  if (rowIndex > 0) {
    logsSheet.getRange(rowIndex, 6).setValue(type);
    logsSheet.getRange(rowIndex, 7).setValue(response);
  }

  return { response: response };
}

// ============================================================
// 統計・ストリーク・カレンダーデータ取得
// ============================================================
function getStats(missionId) {
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

  return { currentStreak: currentStreak, bestStreak: bestStreak, totalDone: totalDone, totalDays: totalDays, calendarData: calendarData, createdDate: String(createdDate) };
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
      sheet.getRange(i + 1, 6).setValue('archived'); return true;
    }
  }
  return false;
}
