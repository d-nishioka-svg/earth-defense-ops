function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('地球防衛オペレーション')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// Secret Manager からシークレットを取得
// ============================================================
const GCP_PROJECT_ID = 'raytech-solutions-development';

function getSecret_(secretName) {
  const token = ScriptApp.getOAuthToken();
  const url = 'https://secretmanager.googleapis.com/v1/projects/' + GCP_PROJECT_ID
    + '/secrets/' + secretName + '/versions/latest:access';

  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });
  const json = JSON.parse(res.getContentText());
  if (json.error) throw new Error('Secret Manager エラー: ' + json.error.message);

  // payload.data は base64 エンコード済み
  return Utilities.newBlob(Utilities.base64Decode(json.payload.data)).getDataAsString();
}

// ============================================================
// Gemini API ヘルパー
// ============================================================
function callGemini_(systemPrompt, userMessage) {
  const apiKey = getSecret_('GEMINI_API_KEY');
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
    sheet.appendRow(['id', 'goalText', 'ifThenTrigger', 'worldStory', 'createdDate', 'status']);
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
// ミッション作成
// ============================================================
function createMission(goalText) {
  const systemPrompt = `あなたは無限の並行世界を管理するAIシステムです。
ユーザーの「継続したいこと」から、If-Thenプランニングと壮大な世界観を生成します。
必ず以下のJSON形式のみで応答し、他の文章は一切含めないでください。

{
  "ifThenTrigger": "具体的なIf-Thenトリガー（例：「お風呂が沸いたら（If）、〇〇を△△する（Then）」という形式で）",
  "worldStory": "壮大な並行世界の危機ストーリー。この継続行動が並行世界をなぜ救うのか、3〜5文で説明。大仰で笑えるくらい大げさに。"
}`;

  const result = callGemini_(systemPrompt, '継続したいこと：' + goalText);
  const data = parseJson_(result);

  const sheet = getMissionsSheet_();
  const id = Utilities.getUuid();
  const now = today_();

  sheet.appendRow([id, goalText, data.ifThenTrigger, data.worldStory, now, 'active']);

  return {
    id: id,
    goalText: goalText,
    ifThenTrigger: data.ifThenTrigger,
    worldStory: data.worldStory,
    createdDate: now,
    status: 'active'
  };
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
      missions.push({
        id: data[i][0],
        goalText: data[i][1],
        ifThenTrigger: data[i][2],
        worldStory: data[i][3],
        createdDate: data[i][4],
        status: data[i][5]
      });
    }
  }

  return missions;
}

// ============================================================
// 今日のミッション（キャラ生成または取得）
// ============================================================
function getTodayMission(missionId) {
  const todayStr = today_();
  const logsSheet = getDailyLogsSheet_();
  const logsData = logsSheet.getDataRange().getValues();

  // 今日のログが既にあれば返す
  for (let i = 1; i < logsData.length; i++) {
    if (String(logsData[i][1]) === String(missionId) && String(logsData[i][2]) === todayStr) {
      return {
        logId: logsData[i][0],
        characterName: logsData[i][3],
        characterPersonality: logsData[i][4],
        result: logsData[i][5],
        responseText: logsData[i][6]
      };
    }
  }

  // ミッション情報を取得
  const missionsSheet = getMissionsSheet_();
  const missionsData = missionsSheet.getDataRange().getValues();
  let mission = null;
  for (let i = 1; i < missionsData.length; i++) {
    if (String(missionsData[i][0]) === String(missionId)) {
      mission = { goalText: missionsData[i][1], ifThenTrigger: missionsData[i][2], worldStory: missionsData[i][3] };
      break;
    }
  }
  if (!mission) return null;

  // 今日限りの新キャラを生成
  const systemPrompt = `あなたは無限の並行世界を管理するAIシステムです。
毎日違う「通信官」キャラクターを生成して大輝に指令を届けます。
必ず以下のJSON形式のみで応答し、他の文章は一切含めないでください。

{
  "characterName": "通信官の名前（ユニークで奇妙な名前）",
  "characterPersonality": "性格・口調の説明（一文。例：「熱血漢で語尾に『だ！』をつける元軍人」）",
  "greeting": "そのキャラになりきって大輝に今日のミッションを伝えるセリフ（200字程度。キャラ全開の口調で。）"
}`;

  const userMsg = `ミッション：${mission.goalText}
If-Thenトリガー：${mission.ifThenTrigger}
並行世界の危機：${mission.worldStory}

今日の通信官キャラを生成して、大輝に指令を伝えてください。`;

  const result = callGemini_(systemPrompt, userMsg);
  const charData = parseJson_(result);

  const logId = Utilities.getUuid();
  logsSheet.appendRow([logId, missionId, todayStr, charData.characterName, charData.characterPersonality, '', charData.greeting]);

  return {
    logId: logId,
    characterName: charData.characterName,
    characterPersonality: charData.characterPersonality,
    result: '',
    responseText: charData.greeting
  };
}

// ============================================================
// 任務完了報告
// ============================================================
function reportDone(missionId) {
  return report_(missionId, 'done');
}

// ============================================================
// エネルギー不足報告
// ============================================================
function reportSkip(missionId) {
  return report_(missionId, 'skip');
}

function report_(missionId, type) {
  const todayStr = today_();
  const logsSheet = getDailyLogsSheet_();
  const logsData = logsSheet.getDataRange().getValues();

  let rowIndex = -1;
  let charName = '';
  let charPersonality = '';

  for (let i = 1; i < logsData.length; i++) {
    if (String(logsData[i][1]) === String(missionId) && String(logsData[i][2]) === todayStr) {
      rowIndex = i + 1;
      charName = logsData[i][3];
      charPersonality = logsData[i][4];
      break;
    }
  }

  const missionsSheet = getMissionsSheet_();
  const missionsData = missionsSheet.getDataRange().getValues();
  let mission = null;
  for (let i = 1; i < missionsData.length; i++) {
    if (String(missionsData[i][0]) === String(missionId)) {
      mission = { goalText: missionsData[i][1], worldStory: missionsData[i][3] };
      break;
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
// アーカイブ取得
// ============================================================
function getArchive(missionId) {
  const logsSheet = getDailyLogsSheet_();
  const logsData = logsSheet.getDataRange().getValues();
  const logs = [];

  for (let i = 1; i < logsData.length; i++) {
    if (String(logsData[i][1]) === String(missionId)) {
      logs.push({
        date: logsData[i][2],
        characterName: logsData[i][3],
        result: logsData[i][5],
        responseText: logsData[i][6]
      });
    }
  }

  return logs.sort(function(a, b) { return b.date > a.date ? 1 : -1; });
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
      return true;
    }
  }
  return false;
}
