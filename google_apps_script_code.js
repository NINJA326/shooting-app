/**
 * NINJA SHOOTING AVERAGE v8.5
 * 選手ランキング対応 Apps Script
 */

const SHEET_RECORDS = 'shooting_records';
const SHEET_PLAYERS = 'players';
const SHEET_LOGS = 'logs';
const SHEET_SUMMARY = 'summary';
const RANK_MIN_ATTEMPTS = 500;

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    setupSheets_();

    if (body.action === 'addRecords' || body.action === 'test') {
      appendRecords_(body.records || [], body.action);
      updateSummary_();
      return json_({ status: 'ok', action: body.action, count: (body.records || []).length });
    }

    return json_({ status: 'ok', message: 'no action' });
  } catch (err) {
    log_('ERROR', String(err));
    return json_({ status: 'error', message: String(err) });
  }
}

function doGet(e) {
  const p = e && e.parameter ? e.parameter : {};
  const callback = p.callback;
  let result;
  try {
    setupSheets_();
    const action = p.action;

    if (action === 'coachAddPlayer') result = coachAddPlayer_(p);
    else if (action === 'loginPlayer') result = loginPlayer_(p);
    else if (action === 'changePassword') result = changePassword_(p);
    else if (action === 'rankings') result = getRankings_();
    else if (action === 'listPlayers' || action === 'players') result = listPlayers_();
    else if (action === 'playerDetail') result = playerDetail_(p.playerId);
    else if (action === 'updatePlayerCategory') result = updatePlayerCategory_(p);
    else if (action === 'dashboard') result = dashboard_();
    else {
      updateSummary_();
      result = { status: 'ok', app: 'NINJA SHOOTING AVERAGE v8.5' };
    }
  } catch (err) {
    log_('GET_ERROR', String(err));
    result = { status:'error', message:String(err) };
  }

  if (callback) {
    return ContentService.createTextOutput(`${callback}(${JSON.stringify(result)})`).setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(result);
}

function setupSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let records = ss.getSheetByName(SHEET_RECORDS);
  if (!records) {
    records = ss.insertSheet(SHEET_RECORDS);
    records.appendRow(['送信日時','同期種別','記録ID','日付','選手名','カテゴリー','練習','種目','ポジション','成功数','試投数','成功率','作成日時','更新日時','削除日時']);
    records.setFrozenRows(1);
  }

  let players = ss.getSheetByName(SHEET_PLAYERS);
  if (!players) {
    players = ss.insertSheet(SHEET_PLAYERS);
    players.appendRow(['選手ID','パスワード','選手名','カテゴリー','作成日時','最終更新','メモ']);
    players.setFrozenRows(1);
  } else {
    const first = players.getRange(1,1,1,Math.max(players.getLastColumn(),7)).getValues()[0];
    if (first[0] !== '選手ID') {
      players.clearContents();
      players.appendRow(['選手ID','パスワード','選手名','カテゴリー','作成日時','最終更新','メモ']);
      players.setFrozenRows(1);
    }
  }

  let summary = ss.getSheetByName(SHEET_SUMMARY);
  if (!summary) {
    summary = ss.insertSheet(SHEET_SUMMARY);
    summary.appendRow(['選手名','カテゴリー','種目','ポジション','成功数合計','試投数合計','成功率']);
    summary.setFrozenRows(1);
  }

  let logs = ss.getSheetByName(SHEET_LOGS);
  if (!logs) {
    logs = ss.insertSheet(SHEET_LOGS);
    logs.appendRow(['日時','種別','内容']);
    logs.setFrozenRows(1);
  }
}

function coachAddPlayer_(p) {
  const playerId = String(p.playerId || '').trim();
  const password = String(p.password || '').trim();
  const name = String(p.name || '').trim();
  const category = String(p.category || '').trim();
  if (!playerId || !password || !name || !category) return { status: 'error', message: '入力不足です。' };
  if (password.length < 4) return { status: 'error', message: 'パスワードは4文字以上にしてください。' };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PLAYERS);
  const last = sheet.getLastRow();
  if (last >= 2) {
    const values = sheet.getRange(2,1,last-1,1).getValues();
    const exists = values.some(row => String(row[0] || '').trim() === playerId);
    if (exists) return { status: 'error', message: 'この選手IDはすでに登録されています。' };
  }
  sheet.appendRow([playerId,password,name,category,new Date(),new Date(),'']);
  log_('COACH_ADD_PLAYER', `${playerId} ${name}`);
  return { status: 'ok', player: { playerId, name, category } };
}


function updatePlayerCategory_(p) {
  const playerId = String(p.playerId || '').trim();
  const category = String(p.category || '').trim();
  if (!playerId || !category) return { status:'error', message:'選手IDまたはカテゴリーがありません。' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PLAYERS);
  const last = sheet.getLastRow();
  if (last < 2) return { status:'error', message:'選手が登録されていません。' };
  const values = sheet.getRange(2,1,last-1,4).getValues();
  for (let i=0; i<values.length; i++) {
    if (String(values[i][0] || '').trim() === playerId) {
      sheet.getRange(i+2,4).setValue(category);
      sheet.getRange(i+2,6).setValue(new Date());
      log_('UPDATE_CATEGORY', `${playerId} -> ${category}`);
      return { status:'ok' };
    }
  }
  return { status:'error', message:'選手IDが見つかりません。' };
}

function loginPlayer_(p) {
  const playerId = String(p.playerId || '').trim();
  const password = String(p.password || '').trim();
  if (!playerId || !password) return { status: 'error', message: '選手IDとパスワードを入力してください。' };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PLAYERS);
  const last = sheet.getLastRow();
  if (last < 2) return { status: 'error', message: '選手IDが登録されていません。コーチに確認してください。' };

  const values = sheet.getRange(2,1,last-1,4).getValues();
  const found = values.find(row => String(row[0] || '').trim() === playerId);
  if (!found) return { status: 'error', message: '選手IDが見つかりません。' };
  if (String(found[1] || '').trim() !== password) return { status: 'error', message: 'パスワードが違います。' };
  return { status: 'ok', player: { playerId: found[0], name: found[2], category: found[3] } };
}

function changePassword_(p) {
  const playerId = String(p.playerId || '').trim();
  const oldPassword = String(p.oldPassword || '').trim();
  const newPassword = String(p.newPassword || '').trim();
  if (!playerId || !oldPassword || !newPassword) return { status: 'error', message: '入力不足です。' };
  if (newPassword.length < 4) return { status: 'error', message: '新しいパスワードは4文字以上にしてください。' };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PLAYERS);
  const last = sheet.getLastRow();
  if (last < 2) return { status: 'error', message: '選手IDがありません。' };
  const values = sheet.getRange(2,1,last-1,4).getValues();
  for (let i=0; i<values.length; i++) {
    if (String(values[i][0] || '').trim() === playerId) {
      if (String(values[i][1] || '').trim() !== oldPassword) return { status: 'error', message: '現在のパスワードが違います。' };
      sheet.getRange(i+2,2).setValue(newPassword);
      sheet.getRange(i+2,6).setValue(new Date());
      log_('CHANGE_PASSWORD', playerId);
      return { status: 'ok' };
    }
  }
  return { status: 'error', message: '選手IDが見つかりません。' };
}

function listPlayers_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PLAYERS);
  const last = sheet.getLastRow();
  if (last < 2) return { status: 'ok', players: [] };
  const values = sheet.getRange(2,1,last-1,7).getValues();
  const players = values.filter(row => row[0] && row[2]).map(row => ({
    playerId: String(row[0] || ''),
    name: String(row[2] || ''),
    category: String(row[3] || ''),
    createdAt: formatDateTime_(row[4]),
    updatedAt: formatDateTime_(row[5]),
    memo: String(row[6] || '')
  })).sort((a,b) => a.category.localeCompare(b.category,'ja') || a.name.localeCompare(b.name,'ja'));
  return { status: 'ok', players };
}

function playerDetail_(playerId) {
  playerId = String(playerId || '').trim();
  if (!playerId) return { status: 'error', message: '選手IDがありません。' };
  const players = listPlayers_().players;
  const player = players.find(p => p.playerId === playerId);
  if (!player) return { status: 'error', message: '選手が見つかりません。' };

  const records = getActiveRecords_().filter(r => r.player === player.name && r.category === player.category);
  return buildPlayerDetail_(player, records);
}

function buildPlayerDetail_(player, records) {
  const made = records.reduce((s,r)=>s+r.made,0);
  const attempts = records.reduce((s,r)=>s+r.attempts,0);
  const total = { made, attempts, rate: attempts>0?Math.round(made/attempts*100):0, records: records.length };

  const byTypeMap = {};
  records.forEach(r => {
    byTypeMap[r.type] = byTypeMap[r.type] || { type:r.type, made:0, attempts:0 };
    byTypeMap[r.type].made += r.made;
    byTypeMap[r.type].attempts += r.attempts;
  });
  const byType = Object.values(byTypeMap).map(x => ({ ...x, rate: x.attempts>0?Math.round(x.made/x.attempts*100):0 }));

  const byPosMap = {};
  records.forEach(r => {
    const key = `${r.type}|${r.position}`;
    byPosMap[key] = byPosMap[key] || { type:r.type, position:r.position, made:0, attempts:0 };
    byPosMap[key].made += r.made;
    byPosMap[key].attempts += r.attempts;
  });
  const byPosition = Object.values(byPosMap).map(x => ({ ...x, rate: x.attempts>0?Math.round(x.made/x.attempts*100):0 })).sort((a,b)=>b.rate-a.rate||b.attempts-a.attempts);
  const recent = records.slice().sort((a,b)=>String(b.date).localeCompare(String(a.date))).slice(0,10);
  return { status:'ok', player, total, byType, byPosition, recent };
}

function dashboard_() {
  const players = listPlayers_().players;
  const records = getActiveRecords_();
  const made = records.reduce((s,r)=>s+r.made,0);
  const attempts = records.reduce((s,r)=>s+r.attempts,0);
  const byTypeMap = {};
  const monthlyMap = {};
  records.forEach(r => {
    byTypeMap[r.type] = byTypeMap[r.type] || { type:r.type, made:0, attempts:0 };
    byTypeMap[r.type].made += r.made;
    byTypeMap[r.type].attempts += r.attempts;
    const month = String(r.date || '').slice(0,7);
    if (month) {
      monthlyMap[month] = monthlyMap[month] || { month, made:0, attempts:0 };
      monthlyMap[month].made += r.made;
      monthlyMap[month].attempts += r.attempts;
    }
  });
  const byType = Object.values(byTypeMap).map(x => ({ ...x, rate:x.attempts>0?Math.round(x.made/x.attempts*100):0 }));
  const monthly = Object.values(monthlyMap).sort((a,b)=>String(a.month).localeCompare(String(b.month))).slice(-12).map(x => ({ ...x, rate:x.attempts>0?Math.round(x.made/x.attempts*100):0 }));
  return { status:'ok', playersCount:players.length, totalMade:made, totalAttempts:attempts, totalRate:attempts>0?Math.round(made/attempts*100):0, byType, monthly };
}

function appendRecords_(records, action) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_RECORDS);
  const values = records.map(r => [new Date(), r.syncAction || action || '', r.id || '', r.date || '', r.player || '', r.category || '', r.practice || r.schedule || '', r.type || '', r.position || '', Number(r.made || 0), Number(r.attempts || 0), Number(r.rate || 0), r.createdAt || '', r.updatedAt || '', r.deletedAt || '']);
  if (values.length) sheet.getRange(sheet.getLastRow()+1,1,values.length,values[0].length).setValues(values);
  log_('APPEND', `${values.length} records`);
}

function getActiveRecords_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_RECORDS);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const values = sheet.getRange(2,1,last-1,15).getValues();
  const byId = new Map();
  values.forEach(row => {
    const syncType = String(row[1] || '');
    const id = String(row[2] || '');
    if (!id) return;
    if (syncType === 'delete') { byId.delete(id); return; }
    byId.set(id, { id, date:formatDate_(row[3]), player:String(row[4]||''), category:String(row[5]||''), practice:String(row[6]||''), type:String(row[7]||''), position:String(row[8]||''), made:Number(row[9]||0), attempts:Number(row[10]||0), rate:Number(row[11]||0) });
  });
  return Array.from(byId.values()).filter(r => r.player && r.type && r.position && r.attempts > 0);
}

function getRankings_() {
  const active = getActiveRecords_();
  const groups = new Map();
  active.forEach(r => {
    const key = `${r.type}|${r.position}|${r.player}|${r.category}`;
    if (!groups.has(key)) groups.set(key, { type:r.type, position:r.position, player:r.player, category:r.category, made:0, attempts:0 });
    const g = groups.get(key); g.made += r.made; g.attempts += r.attempts;
  });
  const rankingMap = {};
  Array.from(groups.values()).forEach(g => {
    g.rate = g.attempts>0?Math.round(g.made/g.attempts*100):0;
    const key = `${g.type}|${g.position}`;
    if (!rankingMap[key]) rankingMap[key] = [];
    rankingMap[key].push({ player:g.player, category:g.category, made:g.made, attempts:g.attempts, rate:g.rate });
  });
  Object.keys(rankingMap).forEach(key => { rankingMap[key] = rankingMap[key].filter(r => Number(r.attempts||0) >= RANK_MIN_ATTEMPTS).sort((a,b)=>b.rate-a.rate||b.attempts-a.attempts); });
  
  const overallMap = new Map();
  active.forEach(r => {
    const key = `${r.player}|${r.category}`;
    if (!overallMap.has(key)) overallMap.set(key, { player:r.player, category:r.category, made:0, attempts:0 });
    const g = overallMap.get(key);
    g.made += r.made;
    g.attempts += r.attempts;
  });
  rankingMap.__overall = Array.from(overallMap.values())
    .map(g => ({ ...g, rate: g.attempts>0?Math.round(g.made/g.attempts*100):0 }))
    .filter(g => g.attempts > 0);
updateSummary_();
  return { status:'ok', generatedAt:new Date().toISOString(), rankings:rankingMap };
}

function updateSummary_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName(SHEET_SUMMARY);
  const active = getActiveRecords_();
  const map = new Map();
  active.forEach(r => {
    const key = `${r.player}|${r.category}|${r.type}|${r.position}`;
    if (!map.has(key)) map.set(key, { player:r.player, category:r.category, type:r.type, position:r.position, made:0, attempts:0 });
    const item = map.get(key); item.made += r.made; item.attempts += r.attempts;
  });
  summarySheet.clearContents();
  summarySheet.appendRow(['選手名','カテゴリー','種目','ポジション','成功数合計','試投数合計','成功率']);
  const rows = Array.from(map.values()).map(x => [x.player,x.category,x.type,x.position,x.made,x.attempts,x.attempts>0?Math.round(x.made/x.attempts*100):0]);
  if (rows.length) summarySheet.getRange(2,1,rows.length,rows[0].length).setValues(rows);
}

function formatDate_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(value || '');
}
function formatDateTime_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  return String(value || '');
}
function log_(type, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_LOGS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LOGS);
    sheet.appendRow(['日時','種別','内容']);
  }
  sheet.appendRow([new Date(),type,message]);
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
