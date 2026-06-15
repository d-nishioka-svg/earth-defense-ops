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
  const resp = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  const json = JSON.parse(resp.getContentText());
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
  try { const v = CacheService.getScriptCache().get(key); return v ? JSON.parse(v) : null; } catch(e) { return null; }
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
  return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
}

function getMissionsSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName('missions');
  if (!sh) {
    sh = ss.insertSheet('missions');
    // id(0) goalText(1) ifThenTrigger(2) worldStory(3) createdDate(4) status(5) timing(6) worldSetting(7) notifyTime(8)
    sh.appendRow(['id','goalText','ifThenTrigger','worldStory','createdDate','status','timing','worldSetting','notifyTime']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getDailyLogsSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName('dailyLogs');
  if (!sh) {
    sh = ss.insertSheet('dailyLogs');
    // id(0) missionId(1) date(2) characterName(3) characterPersonality(4)
    // result(5) responseText(6) doneResponse(7) skipResponse(8)
    sh.appendRow(['id','missionId','date','characterName','characterPersonality','result','responseText','doneResponse','skipResponse']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function today_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function timingFromNotifyTime_(t) {
  if (!t) return '';
  const p = t.split(':'), h = parseInt(p[0], 10), m = parseInt(p[1] || '0', 10);
  return m === 0 ? h + '時になったら' : h + '時' + m + '分になったら';
}

// ============================================================
// 通知トリガー
// ============================================================
function ensureNotifyTrigger_() {
  const ts = ScriptApp.getProjectTriggers();
  for (let i = 0; i < ts.length; i++) { if (ts[i].getHandlerFunction() === 'checkAndNotify') return; }
  ScriptApp.newTrigger('checkAndNotify').timeBased().everyHours(1).create();
}

function checkAndNotify() {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const currentHour = parseInt(Utilities.formatDate(now, tz, 'HH'), 10);
  const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const msData = getMissionsSheet_().getDataRange().getValues();
  const lgData = getDailyLogsSheet_().getDataRange().getValues();
  const loggedToday = new Set();
  for (let i = 1; i < lgData.length; i++) {
    if (String(lgData[i][2]) === todayStr && lgData[i][5] !== '') loggedToday.add(String(lgData[i][1]));
  }
  const props = PropertiesService.getScriptProperties();
  const email = Session.getEffectiveUser().getEmail();
  for (let i = 1; i < msData.length; i++) {
    if (msData[i][5] !== 'active') continue;
    const nt = String(msData[i][8] || '');
    if (!nt || parseInt(nt.split(':')[0], 10) !== currentHour) continue;
    const mid = String(msData[i][0]);
    if (loggedToday.has(mid)) continue;
    const key = 'notify_' + mid + '_' + todayStr;
    if (props.getProperty(key)) continue;
    try {
      MailApp.sendEmail({ to: email, subject: '🌍 地球防衛OPS ｜ ミッション通知',
        body: ['大輝へ', '', '時間だ。ミッションを確認しろ。', '',
          '【ミッション】' + String(msData[i][1]), '【If-Then】' + String(msData[i][2]), '',
          'アプリを開いて今日のキャラクターの指令を受け取れ。', '', '─ 地球防衛オペレーション'].join('\n') });
      props.setProperty(key, '1');
    } catch(e) {}
  }
}

// ============================================================
// ミッション生成（保存なし）
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
  "doneResponse": "大輝が任務完了した場合のキャラのリアクション（120字以内）",
  "skipResponse": "大輝がサボった場合のキャラのリアクション（「1回だけでいい」という懇願を含め180字以内）",
  "ifThenOptions": [
    {"trigger": "トリガー1", "action": "行動1"},
    {"trigger": "トリガー2", "action": "行動2"},
    {"trigger": "トリガー3", "action": "行動3"}
  ]
}

━━ キャラ生成ルール ━━
・既存のSFテンプレート（ロボット・宇宙人・AI・魔法使い）は絶対禁止
・全く異なる2つの概念をサイバー/ファンタジー要素で掛け合わせた新キャラ
 例：「暗黒銀河の読書カマキリ」「並行世界βのサイバーお遍路さん」「量子世界のハードコア盆栽職人」
・characterIntroで職業・現在の状況・なぜ大輝に通信したかをキャラ全開の口調で語らせる
・口調・語尾は最初から最後まで完全に崩さないこと

━━ If-Thenトリガールール ━━
・「大輝が1日に必ず行う無意識な生活行動」を引き金にする
・この具体性レベルが必要：
  ✓「夜、スマホを充電器に挿した瞬間に」
  ✓「お風呂が沸いたアラームが鳴ったら」
  ✓「朝、コーヒーをマグカップに注ぎ終えたら」
  ✗「毎日やる」「時間を決めて」→ NG
・3つは朝・夜・お風呂などバラバラのシーン
・actionに具体的な量（○回/○分）を入れる

━━ missionTextルール ━━
・前半：行動が並行世界を救う仕組みを壮大かつ大真面目に（200字）
・後半：サボった場合の損失を具体的に描写（150字）
・実在する場所・もの（エベレスト、NASAなど世界中なんでもOK）を最低1つ入れる

━━ doneResponse / skipResponseルール ━━
・キャラの口調を最初から最後まで崩さない
・doneResponse：世界で何が救われたかを1文 + 大輝を称える + キャラらしい締め
・skipResponse：今何がやばいかを1文 + 「1回だけでいい」という懇願 + キャラらしい必死さ`;

  const userMsg = '継続したいこと：' + goalText + (timing ? '\nタイミング：' + timing : '');
  return parseJson_(callGemini_(systemPrompt, userMsg));
}

// ============================================================
// ミッション作成（1ステップ：生成 + 保存）
// ============================================================
function createMission(goalText, notifyTime) {
  var data = generateMissionOptions(goalText, notifyTime);
  return saveMission(goalText, notifyTime, 0, data);
}

// ============================================================
// ミッション保存（選択確定）
// ============================================================
function saveMission(goalText, notifyTime, selectedIndex, generatedData) {
  const timing = timingFromNotifyTime_(notifyTime);
  const opt = generatedData.ifThenOptions[selectedIndex];
  const ifThenTrigger = opt.trigger + ' → ' + opt.action;

  const msSheet = getMissionsSheet_();
  const id = Utilities.getUuid();
  const now = today_();
  msSheet.appendRow([id, goalText, ifThenTrigger, generatedData.missionText, now, 'active',
    timing, generatedData.worldSetting || '', notifyTime || '']);
  cacheRemove_('missions');

  // 初回（今日）のログ：キャラ自己紹介 + 事前生成済みのdone/skipレスポンスを保存
  const lgSheet = getDailyLogsSheet_();
  const logId = Utilities.getUuid();
  lgSheet.appendRow([logId, id, now,
    generatedData.characterName, generatedData.characterPersonality,
    '', generatedData.characterIntro,
    generatedData.doneResponse || '', generatedData.skipResponse || '']);

  if (notifyTime) { try { ensureNotifyTrigger_(); } catch(e) {} }

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
  const data = getMissionsSheet_().getDataRange().getValues();
  const missions = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][5] === 'active') missions.push({
      id: data[i][0], goalText: data[i][1], ifThenTrigger: data[i][2],
      worldStory: data[i][3], createdDate: data[i][4], status: data[i][5],
      timing: data[i][6] || '', worldSetting: data[i][7] || '', notifyTime: data[i][8] || ''
    });
  }
  cachePut_('missions', missions);
  return missions;
}

// ============================================================
// 今日のミッション
// 【API呼び出しは1日1回のみ】
// greeting + doneResponse + skipResponse を一括生成して保存。
// 報告時（report_）はAPIを呼ばず保存済み応答を返す。
// ============================================================
function getTodayMission(missionId) {
  const todayStr = today_();
  const cacheKey = 'today_' + missionId + '_' + todayStr;
  const cached = cacheGet_(cacheKey);
  if (cached) return cached;

  const lgSheet = getDailyLogsSheet_();
  const lgData = lgSheet.getDataRange().getValues();

  // 今日のログが既にある → キャッシュに入れて返す
  for (let i = 1; i < lgData.length; i++) {
    if (String(lgData[i][1]) === String(missionId) && String(lgData[i][2]) === todayStr) {
      const entry = {
        logId: lgData[i][0], characterName: lgData[i][3],
        characterPersonality: lgData[i][4], result: lgData[i][5], responseText: lgData[i][6]
      };
      cachePut_(cacheKey, entry);
      return entry;
    }
  }

  // 今日のログなし → 新キャラを生成（1日1回のAPI呼び出し）
  const msData = getMissionsSheet_().getDataRange().getValues();
  let mission = null;
  for (let i = 1; i < msData.length; i++) {
    if (String(msData[i][0]) === String(missionId)) {
      mission = { goalText: msData[i][1], ifThenTrigger: msData[i][2],
        worldStory: msData[i][3], timing: msData[i][6] || '', worldSetting: msData[i][7] || '' };
      break;
    }
  }
  if (!mission) return null;

  const totalLogs = lgData.length - 1;
  const isSpecial = (totalLogs > 0 && totalLogs % 10 === 0);

  const systemPrompt = `あなたは毎日違うキャラクターで大輝に並行世界の指令を届けるAIです。
以下のJSON形式のみで応答してください。

{
  "characterName": "キャラ名（今日の新キャラ）",
  "characterPersonality": "口調・性格（一文）",
  "greeting": "今日の指令セリフ（200字程度）",
  "doneResponse": "大輝が任務完了した場合のリアクション（120字以内）",
  "skipResponse": "大輝がサボった場合のリアクション（「1回だけでいい」という懇願を含め180字以内）"
}

キャラ生成ルール：
・既存のSFテンプレート禁止。2つの異なる概念を掛け合わせた新キャラ（毎日違う！）
・口調・語尾を最初から最後まで崩さない（全フィールド共通）
・greeting：自分の世界の危機状況 + なぜ大輝の行動が必要か + 実在する場所・もの（世界中なんでもOK）を1つ + ミッション（${mission.timing ? mission.timing + '、' : ''}${mission.goalText}）を伝える
・doneResponse：世界で何が救われたかを実在する場所・ものを使って + 大輝を称える + キャラらしい締め
・skipResponse：今何がやばいかを + 「1回だけでいい」という必死な懇願 + キャラらしい締め` + (isSpecial ? '\n・【特別通信】キャラ名に称号をつける。' : '');

  const userMsg = `ミッション：${mission.goalText}
If-Then：${mission.ifThenTrigger}
世界設定：${mission.worldStory}
舞台：${mission.worldSetting || '並行世界'}`;

  const charData = parseJson_(callGemini_(systemPrompt, userMsg));

  const logId = Utilities.getUuid();
  // doneResponse(col8) と skipResponse(col9) を事前保存
  lgSheet.appendRow([logId, missionId, todayStr,
    charData.characterName, charData.characterPersonality,
    '', charData.greeting,
    charData.doneResponse || '', charData.skipResponse || '']);

  const entry = {
    logId: logId, characterName: charData.characterName,
    characterPersonality: charData.characterPersonality,
    result: '', responseText: charData.greeting, isSpecial: isSpecial
  };
  cachePut_(cacheKey, entry);
  return entry;
}

// ============================================================
// 任務完了 / エネルギー不足
// 【APIを呼ばず、事前保存済みレスポンスを返す】
// ============================================================
function reportDone(missionId) { return report_(missionId, 'done'); }
function reportSkip(missionId) { return report_(missionId, 'skip'); }

function report_(missionId, type) {
  const todayStr = today_();
  const lgSheet = getDailyLogsSheet_();
  const lgData = lgSheet.getDataRange().getValues();

  let rowIndex = -1, preResponse = '';
  for (let i = 1; i < lgData.length; i++) {
    if (String(lgData[i][1]) === String(missionId) && String(lgData[i][2]) === todayStr) {
      rowIndex = i + 1;
      // col8=doneResponse(index7), col9=skipResponse(index8)
      preResponse = String(type === 'done' ? (lgData[i][7] || '') : (lgData[i][8] || ''));
      break;
    }
  }

  // 事前保存済みレスポンスがある → APIを呼ばない
  if (preResponse) {
    if (rowIndex > 0) {
      lgSheet.getRange(rowIndex, 6).setValue(type);       // result列
      lgSheet.getRange(rowIndex, 7).setValue(preResponse); // responseText列を上書き
    }
    const cacheKey = 'today_' + missionId + '_' + todayStr;
    const cached = cacheGet_(cacheKey);
    if (cached) { cached.result = type; cached.responseText = preResponse; cachePut_(cacheKey, cached); }
    cacheRemove_('stats_' + missionId);
    return { response: preResponse };
  }

  // フォールバック：旧データや初回保存なし → Geminiを呼ぶ
  let charName = '', charPersonality = '';
  if (rowIndex > 0) { charName = lgData[rowIndex - 2 + 1][3]; charPersonality = lgData[rowIndex - 2 + 1][4]; }
  for (let i = 1; i < lgData.length; i++) {
    if (String(lgData[i][1]) === String(missionId) && String(lgData[i][2]) === todayStr) {
      charName = lgData[i][3]; charPersonality = lgData[i][4]; break;
    }
  }
  const msData = getMissionsSheet_().getDataRange().getValues();
  let worldStory = '', goalText = '';
  for (let i = 1; i < msData.length; i++) {
    if (String(msData[i][0]) === String(missionId)) { goalText = msData[i][1]; worldStory = msData[i][3]; break; }
  }

  const sysPrompt = `あなたは「${charName}」です。口調：${charPersonality}。キャラを崩さず話してください。`;
  const userMsg = type === 'done'
    ? `大輝が「${goalText}」を完了！並行世界：${worldStory}\n120字以内でリアクション（何が救われたか・大輝を称える・締め）`
    : `大輝が「今日は無理」と言っている。ミッション：「${goalText}」並行世界：${worldStory}\n180字以内（今の危機・1回だけでいいという懇願・キャラらしい締め）`;

  const response = callGemini_(sysPrompt, userMsg);
  if (rowIndex > 0) {
    lgSheet.getRange(rowIndex, 6).setValue(type);
    lgSheet.getRange(rowIndex, 7).setValue(response);
  }
  const cacheKey = 'today_' + missionId + '_' + todayStr;
  const cached = cacheGet_(cacheKey);
  if (cached) { cached.result = type; cached.responseText = response; cachePut_(cacheKey, cached); }
  cacheRemove_('stats_' + missionId);
  return { response: response };
}

// ============================================================
// 統計・ストリーク・カレンダー
// ============================================================
function getStats(missionId) {
  const key = 'stats_' + missionId;
  const cached = cacheGet_(key);
  if (cached) return cached;

  const lgData = getDailyLogsSheet_().getDataRange().getValues();
  const msData = getMissionsSheet_().getDataRange().getValues();
  let createdDate = '';
  for (let i = 1; i < msData.length; i++) {
    if (String(msData[i][0]) === String(missionId)) { createdDate = msData[i][4]; break; }
  }

  const logMap = {};
  for (let i = 1; i < lgData.length; i++) {
    if (String(lgData[i][1]) === String(missionId)) logMap[lgData[i][2]] = lgData[i][5];
  }

  const tz = Session.getScriptTimeZone();
  const calendarData = {};
  for (let d = 0; d < 60; d++) {
    const dt = new Date(); dt.setDate(dt.getDate() - d);
    const k = Utilities.formatDate(dt, tz, 'yyyy-MM-dd');
    calendarData[k] = logMap[k] || '';
  }

  let currentStreak = 0, checking = true;
  for (let d = 0; d < 365; d++) {
    const dt = new Date(); dt.setDate(dt.getDate() - d);
    const k = Utilities.formatDate(dt, tz, 'yyyy-MM-dd');
    const r = logMap[k];
    if (d === 0 && r !== 'done') { checking = false; break; }
    if (r === 'done') { if (checking) currentStreak++; }
    else if (r === 'skip') { checking = false; }
  }

  const doneDates = Object.keys(logMap).filter(function(k){ return logMap[k] === 'done'; }).sort();
  let bestStreak = 0, tmp = 0, prev = null;
  doneDates.forEach(function(d) {
    tmp = prev && (new Date(d) - new Date(prev)) / 86400000 === 1 ? tmp + 1 : 1;
    if (tmp > bestStreak) bestStreak = tmp;
    prev = d;
  });

  const totalDone = Object.values(logMap).filter(function(v){ return v === 'done'; }).length;
  const totalDays = createdDate ? Math.floor((new Date() - new Date(createdDate)) / 86400000) + 1 : 0;

  const stats = { currentStreak: currentStreak, bestStreak: bestStreak, totalDone: totalDone, totalDays: totalDays, calendarData: calendarData, createdDate: String(createdDate) };
  cachePut_(key, stats);
  return stats;
}

// ============================================================
// アーカイブ取得
// ============================================================
function getArchive(missionId) {
  const lgData = getDailyLogsSheet_().getDataRange().getValues();
  const logs = [];
  for (let i = 1; i < lgData.length; i++) {
    if (String(lgData[i][1]) === String(missionId))
      logs.push({ date: lgData[i][2], characterName: lgData[i][3], result: lgData[i][5], responseText: lgData[i][6] });
  }
  return logs.sort(function(a, b){ return b.date > a.date ? 1 : -1; });
}

// ============================================================
// ミッション削除（アーカイブ化）
// ============================================================
function deleteMission(missionId) {
  const sh = getMissionsSheet_();
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(missionId)) {
      sh.getRange(i + 1, 6).setValue('archived');
      cacheRemove_('missions');
      return true;
    }
  }
  return false;
}
