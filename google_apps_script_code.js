/**
 * NINJA SHOOTING AVERAGE v12.0
 * 選手用・コーチ用・成長記録・アジリティー記録対応 Apps Script
 */
const SHEET_RECORDS = 'shooting_records';
const SHEET_PLAYERS = 'players';
const SHEET_LOGS = 'logs';
const SHEET_SUMMARY = 'summary';
const SHEET_BODY_MATRIX = '身体測定';
const SHEET_AGILITY_MATRIX = 'アジリティ測定';
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
    if (body.action === 'addGrowthRecords') {
      appendGrowthRecords_(body.records || []);
      return json_({ status: 'ok', action: body.action, count: (body.records || []).length });
    }
    if (body.action === 'addAgilityRecords') {
      appendAgilityRecords_(body.records || []);
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
    else if (action === 'playerRecords') result = playerRecords_(p.playerId);
    else if (action === 'growthRecords') result = growthRecords_(p.playerId);
    else if (action === 'growthSummary') result = growthSummary_();
    else if (action === 'agilityRecords') result = agilityRecords_(p.playerId);
    else if (action === 'agilityRankings') result = agilityRankings_();
    else if (action === 'agilitySummary') result = agilitySummary_();
    else if (action === 'updatePlayerCategory') result = updatePlayerCategory_(p);
    else if (action === 'dashboard') result = dashboard_();
    else { updateSummary_(); result = { status: 'ok', app: 'NINJA SHOOTING AVERAGE v12.0' }; }
  } catch (err) {
    log_('GET_ERROR', String(err));
    result = { status:'error', message:String(err) };
  }
  if (callback) return ContentService.createTextOutput(`${callback}(${JSON.stringify(result)})`).setMimeType(ContentService.MimeType.JAVASCRIPT);
  return json_(result);
}

function setupSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SHEET_RECORDS, ['送信日時','同期種別','記録ID','日付','選手名','カテゴリー','練習','種目','ポジション','成功数','試投数','成功率','作成日時','更新日時','削除日時']);
  ensureSheet_(ss, SHEET_PLAYERS, ['選手ID','パスワード','選手名','カテゴリー','作成日時','最終更新','メモ']);
  ensureSheet_(ss, SHEET_SUMMARY, ['選手名','カテゴリー','種目','ポジション','成功数合計','試投数合計','成功率']);
  ensureMatrixSheets_(ss);
  migrateLegacyGrowthSheets_(ss);
  migrateLegacyAgilitySheets_(ss);
  ensureSheet_(ss, SHEET_LOGS, ['日時','種別','内容']);
}
function ensureSheet_(ss, name, header) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) { sheet = ss.insertSheet(name); sheet.appendRow(header); sheet.setFrozenRows(1); return sheet; }
  const current = sheet.getRange(1,1,1,Math.max(sheet.getLastColumn(), header.length)).getValues()[0];
  let need = false;
  for (let i=0;i<header.length;i++) if (String(current[i]||'') !== header[i]) need = true;
  if (need) { sheet.getRange(1,1,1,header.length).setValues([header]); sheet.setFrozenRows(1); }
  return sheet;
}


function ensureMatrixSheets_(ss) {
  const monthHeaders = [];
  const today = new Date();
  const year = today.getFullYear();
  for (let m = 4; m <= 12; m++) monthHeaders.push(`${year}/${m}`);
  for (let m = 1; m <= 3; m++) monthHeaders.push(`${year + 1}/${m}`);

  let body = ss.getSheetByName(SHEET_BODY_MATRIX);
  if (!body) {
    body = ss.insertSheet(SHEET_BODY_MATRIX);
    body.getRange(1,1,1,2 + monthHeaders.length).setValues([['選手名','項目',...monthHeaders]]);
    body.setFrozenRows(1); body.setFrozenColumns(2);
  }
  syncBodyMatrixPlayers_(body);

  let agility = ss.getSheetByName(SHEET_AGILITY_MATRIX);
  if (!agility) {
    agility = ss.insertSheet(SHEET_AGILITY_MATRIX);
    agility.getRange(1,1,1,2 + monthHeaders.length).setValues([['選手名','種目',...monthHeaders]]);
    agility.setFrozenRows(1); agility.setFrozenColumns(2);
  }
  syncAgilityMatrixPlayers_(agility);
}

function syncBodyMatrixPlayers_(sheet) {
  const players = listPlayers_().players;
  const existing = new Set();
  if (sheet.getLastRow() >= 2) {
    sheet.getRange(2,1,sheet.getLastRow()-1,2).getDisplayValues().forEach(r => {
      const name = String(r[0]||'').trim(), metric = String(r[1]||'').trim();
      if (name && metric) existing.add(`${name}|${metric}`);
    });
  }
  const add = [];
  players.forEach(p => {
    ['身長','体重'].forEach(metric => {
      if (!existing.has(`${p.name}|${metric}`)) add.push([p.name, metric]);
    });
  });
  if (add.length) sheet.getRange(sheet.getLastRow()+1,1,add.length,2).setValues(add);
}

function syncAgilityMatrixPlayers_(sheet) {
  const players = listPlayers_().players;
  const defaults = ['シャトルラン','反復横跳び','L字コーンドリル','垂直跳び'];
  const existing = new Set();
  if (sheet.getLastRow() >= 2) {
    sheet.getRange(2,1,sheet.getLastRow()-1,2).getDisplayValues().forEach(r => {
      const name = String(r[0]||'').trim(), metric = String(r[1]||'').trim();
      if (name && metric) existing.add(`${name}|${metric}`);
    });
  }
  const add = [];
  players.forEach(p => defaults.forEach(metric => {
    if (!existing.has(`${p.name}|${metric}`)) add.push([p.name, metric]);
  }));
  if (add.length) sheet.getRange(sheet.getLastRow()+1,1,add.length,2).setValues(add);
}

function matrixHeaderDate_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const s = String(value || '').trim();
  let m = s.match(/^(\d{4})[\/\-.年](\d{1,2})(?:[\/\-.月](\d{1,2}))?/);
  if (m) return `${m[1]}-${String(Number(m[2])).padStart(2,'0')}-${String(Number(m[3]||1)).padStart(2,'0')}`;
  m = s.match(/^(\d{1,2})[\/\-.月](\d{1,2})/);
  if (m) return `${new Date().getFullYear()}-${String(Number(m[1])).padStart(2,'0')}-${String(Number(m[2])).padStart(2,'0')}`;
  return normalizeDateString_(s);
}

function coachAddPlayer_(p) {
  const playerId = String(p.playerId || '').trim();
  const password = String(p.password || '').trim();
  const name = String(p.name || '').trim();
  const category = String(p.category || '').trim();
  if (!playerId || !password || !name || !category) return { status:'error', message:'入力不足です。' };
  if (password.length < 4) return { status:'error', message:'パスワードは4文字以上にしてください。' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PLAYERS);
  const last = sheet.getLastRow();
  if (last >= 2) {
    const values = sheet.getRange(2,1,last-1,1).getValues();
    if (values.some(row => String(row[0]||'').trim() === playerId)) return { status:'error', message:'この選手IDはすでに登録されています。' };
  }
  sheet.appendRow([playerId,password,name,category,new Date(),new Date(),'']);
  log_('COACH_ADD_PLAYER', `${playerId} ${name}`);
  return { status:'ok', player:{ playerId, name, category } };
}
function updatePlayerCategory_(p) {
  const playerId = String(p.playerId || '').trim();
  const category = String(p.category || '').trim();
  if (!playerId || !category) return { status:'error', message:'選手IDまたはカテゴリーがありません。' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PLAYERS);
  const last = sheet.getLastRow();
  if (last < 2) return { status:'error', message:'選手が登録されていません。' };
  const values = sheet.getRange(2,1,last-1,4).getValues();
  for (let i=0;i<values.length;i++) {
    if (String(values[i][0]||'').trim() === playerId) {
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
  if (!playerId || !password) return { status:'error', message:'選手IDとパスワードを入力してください。' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PLAYERS);
  const last = sheet.getLastRow();
  if (last < 2) return { status:'error', message:'選手IDが登録されていません。コーチに確認してください。' };
  const values = sheet.getRange(2,1,last-1,4).getValues();
  const found = values.find(row => String(row[0]||'').trim() === playerId);
  if (!found) return { status:'error', message:'選手IDが見つかりません。' };
  if (String(found[1]||'').trim() !== password) return { status:'error', message:'パスワードが違います。' };
  return { status:'ok', player:{ playerId: found[0], name: found[2], category: found[3] } };
}
function changePassword_(p) {
  const playerId = String(p.playerId || '').trim();
  const oldPassword = String(p.oldPassword || '').trim();
  const newPassword = String(p.newPassword || '').trim();
  if (!playerId || !oldPassword || !newPassword) return { status:'error', message:'入力不足です。' };
  if (newPassword.length < 4) return { status:'error', message:'新しいパスワードは4文字以上にしてください。' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PLAYERS);
  const last = sheet.getLastRow();
  if (last < 2) return { status:'error', message:'選手IDがありません。' };
  const values = sheet.getRange(2,1,last-1,4).getValues();
  for (let i=0;i<values.length;i++) {
    if (String(values[i][0]||'').trim() === playerId) {
      if (String(values[i][1]||'').trim() !== oldPassword) return { status:'error', message:'現在のパスワードが違います。' };
      sheet.getRange(i+2,2).setValue(newPassword);
      sheet.getRange(i+2,6).setValue(new Date());
      log_('CHANGE_PASSWORD', playerId);
      return { status:'ok' };
    }
  }
  return { status:'error', message:'選手IDが見つかりません。' };
}
function listPlayers_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PLAYERS);
  const last = sheet.getLastRow();
  if (last < 2) return { status:'ok', players:[] };
  const values = sheet.getRange(2,1,last-1,7).getValues();
  const players = values.filter(row => row[0] && row[2]).map(row => ({
    playerId:String(row[0]||''), name:String(row[2]||''), category:String(row[3]||''), createdAt:formatDateTime_(row[4]), updatedAt:formatDateTime_(row[5]), memo:String(row[6]||'')
  })).sort((a,b)=>a.category.localeCompare(b.category,'ja') || a.name.localeCompare(b.name,'ja'));
  return { status:'ok', players };
}
function findPlayer_(playerId) {
  const player = listPlayers_().players.find(p => p.playerId === String(playerId||'').trim());
  return player || null;
}
function playerDetail_(playerId) {
  const player = findPlayer_(playerId);
  if (!player) return { status:'error', message:'選手が見つかりません。' };
  const records = getActiveRecords_().filter(r => r.player === player.name);
  return buildPlayerDetail_(player, records);
}
function playerRecords_(playerId) {
  const player = findPlayer_(playerId);
  if (!player) return { status:'error', message:'選手が見つかりません。' };
  const records = getActiveRecords_().filter(r => r.player === player.name).sort((a,b)=>String(a.date).localeCompare(String(b.date))).map(r => ({...r, syncAction:'cloud'}));
  return { status:'ok', player, records };
}
function buildPlayerDetail_(player, records) {
  const made = records.reduce((s,r)=>s+Number(r.made||0),0);
  const attempts = records.reduce((s,r)=>s+Number(r.attempts||0),0);
  const total = { made, attempts, rate: attempts>0?Math.round(made/attempts*100):0, records: records.length };
  const byTypeMap = {}, byPosMap = {};
  records.forEach(r => {
    byTypeMap[r.type] = byTypeMap[r.type] || { type:r.type, made:0, attempts:0 };
    byTypeMap[r.type].made += Number(r.made||0); byTypeMap[r.type].attempts += Number(r.attempts||0);
    const key = `${r.type}|${r.position}`;
    byPosMap[key] = byPosMap[key] || { type:r.type, position:r.position, made:0, attempts:0 };
    byPosMap[key].made += Number(r.made||0); byPosMap[key].attempts += Number(r.attempts||0);
  });
  const byType = Object.values(byTypeMap).map(x=>({...x, rate:x.attempts>0?Math.round(x.made/x.attempts*100):0}));
  const byPosition = Object.values(byPosMap).map(x=>({...x, rate:x.attempts>0?Math.round(x.made/x.attempts*100):0})).sort((a,b)=>b.rate-a.rate||b.attempts-a.attempts);
  const recent = records.slice().sort((a,b)=>String(b.date).localeCompare(String(a.date))).slice(0,10);
  return { status:'ok', player, total, byType, byPosition, recent };
}
function dashboard_() {
  const players = listPlayers_().players;
  const records = getActiveRecords_();
  const made = records.reduce((s,r)=>s+Number(r.made||0),0);
  const attempts = records.reduce((s,r)=>s+Number(r.attempts||0),0);
  const byTypeMap = {}, monthlyMap = {};
  records.forEach(r => {
    byTypeMap[r.type] = byTypeMap[r.type] || { type:r.type, made:0, attempts:0 };
    byTypeMap[r.type].made += Number(r.made||0); byTypeMap[r.type].attempts += Number(r.attempts||0);
    const month = String(r.date||'').slice(0,7);
    if (month) { monthlyMap[month] = monthlyMap[month] || { month, made:0, attempts:0 }; monthlyMap[month].made += Number(r.made||0); monthlyMap[month].attempts += Number(r.attempts||0); }
  });
  const byType = Object.values(byTypeMap).map(x=>({...x, rate:x.attempts>0?Math.round(x.made/x.attempts*100):0}));
  const monthly = Object.values(monthlyMap).sort((a,b)=>String(a.month).localeCompare(String(b.month))).slice(-12).map(x=>({...x, rate:x.attempts>0?Math.round(x.made/x.attempts*100):0}));
  return { status:'ok', playersCount:players.length, totalMade:made, totalAttempts:attempts, totalRate:attempts>0?Math.round(made/attempts*100):0, byType, monthly };
}

function appendRecords_(records, action) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RECORDS);
  const values = (records||[]).map(r => [new Date(), r.syncAction || action || '', r.id || '', r.date || '', r.player || '', r.category || '', r.practice || r.schedule || '', r.type || '', r.position || '', Number(r.made || 0), Number(r.attempts || 0), Number(r.rate || 0), r.createdAt || '', r.updatedAt || '', r.deletedAt || '']);
  if (values.length) sheet.getRange(sheet.getLastRow()+1,1,values.length,values[0].length).setValues(values);
  log_('APPEND', `${values.length} records`);
}
function getActiveRecords_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RECORDS);
  const last = sheet.getLastRow(); if (last < 2) return [];
  const values = sheet.getRange(2,1,last-1,15).getValues();
  const byId = new Map();
  values.forEach(row => {
    const syncType = String(row[1]||''); const id = String(row[2]||''); if (!id) return;
    if (syncType === 'delete') { byId.delete(id); return; }
    byId.set(id, { id, date:formatDate_(row[3]), player:String(row[4]||''), category:String(row[5]||''), practice:String(row[6]||''), type:String(row[7]||''), position:String(row[8]||''), made:Number(row[9]||0), attempts:Number(row[10]||0), rate:Number(row[11]||0) });
  });
  return Array.from(byId.values()).filter(r=>r.player && r.type && r.position && r.attempts > 0);
}
function getRankings_() {
  const active = getActiveRecords_(); const rankingMap = {}; const groups = new Map();
  active.forEach(r => { const key = `${r.type}|${r.position}|${r.player}|${r.category}`; if (!groups.has(key)) groups.set(key,{type:r.type,position:r.position,player:r.player,category:r.category,made:0,attempts:0}); const g=groups.get(key); g.made += r.made; g.attempts += r.attempts; });
  Array.from(groups.values()).forEach(g => { g.rate = g.attempts>0?Math.round(g.made/g.attempts*100):0; const key = `${g.type}|${g.position}`; if (!rankingMap[key]) rankingMap[key]=[]; rankingMap[key].push({ player:g.player, category:g.category, made:g.made, attempts:g.attempts, rate:g.rate }); });
  Object.keys(rankingMap).forEach(key => rankingMap[key] = rankingMap[key].filter(r=>Number(r.attempts||0)>=RANK_MIN_ATTEMPTS).sort((a,b)=>b.rate-a.rate||b.attempts-a.attempts));
  const overallMap = new Map();
  active.forEach(r => { const key = `${r.player}|${r.category}`; if (!overallMap.has(key)) overallMap.set(key,{player:r.player,category:r.category,made:0,attempts:0}); const g=overallMap.get(key); g.made += r.made; g.attempts += r.attempts; });
  rankingMap.__overall = Array.from(overallMap.values()).map(g=>({...g, rate:g.attempts>0?Math.round(g.made/g.attempts*100):0})).filter(g=>g.attempts>0);
  updateSummary_();
  return { status:'ok', generatedAt:new Date().toISOString(), rankings:rankingMap };
}
function updateSummary_() {
  const summarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SUMMARY);
  const active = getActiveRecords_(); const map = new Map();
  active.forEach(r => { const key = `${r.player}|${r.category}|${r.type}|${r.position}`; if (!map.has(key)) map.set(key,{player:r.player,category:r.category,type:r.type,position:r.position,made:0,attempts:0}); const item=map.get(key); item.made += r.made; item.attempts += r.attempts; });
  summarySheet.clearContents(); summarySheet.appendRow(['選手名','カテゴリー','種目','ポジション','成功数合計','試投数合計','成功率']);
  const rows = Array.from(map.values()).map(x => [x.player,x.category,x.type,x.position,x.made,x.attempts,x.attempts>0?Math.round(x.made/x.attempts*100):0]);
  if (rows.length) summarySheet.getRange(2,1,rows.length,rows[0].length).setValues(rows);
}

function appendGrowthRecords_(records) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_BODY_MATRIX);
  syncBodyMatrixPlayers_(sheet);
  (records || []).forEach(r => writeGrowthRecordToBodyMatrix_(sheet, r));
  normalizeAndSortBodyMonthColumns_(sheet);
  log_('APPEND_GROWTH_BODY_MATRIX', `${(records || []).length} records`);
}

function monthKeyFromValue_(value) {
  const date = normalizeDateString_(value);
  if (!date) return '';
  const m = date.match(/^(\d{4})-(\d{2})/);
  return m ? `${m[1]}/${Number(m[2])}` : '';
}

function monthSortValue_(value) {
  const key = monthKeyFromValue_(value) || String(value || '').trim();
  const m = key.match(/^(\d{4})\/(\d{1,2})$/);
  return m ? Number(m[1]) * 100 + Number(m[2]) : 999999;
}

function findOrCreateBodyMetricRow_(sheet, player, metric) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues();
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0] || '').trim() === player && String(values[i][1] || '').trim() === metric) return i + 2;
    }
  }
  const row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, 2).setValues([[player, metric]]);
  return row;
}

function findOrCreateBodyMonthColumn_(sheet, dateValue) {
  const monthKey = monthKeyFromValue_(dateValue);
  if (!monthKey) return 0;
  const lastCol = Math.max(sheet.getLastColumn(), 2);
  if (lastCol >= 3) {
    const headers = sheet.getRange(1, 3, 1, lastCol - 2).getDisplayValues()[0];
    for (let i = 0; i < headers.length; i++) {
      if (monthKeyFromValue_(headers[i]) === monthKey || String(headers[i] || '').trim() === monthKey) return i + 3;
    }
  }
  const col = lastCol + 1;
  sheet.getRange(1, col).setValue(monthKey);
  return col;
}

function writeGrowthRecordToBodyMatrix_(sheet, record) {
  const player = String(record.player || '').trim();
  const dateValue = record.date || record.measuredAtIso || record.createdAt;
  if (!player || !dateValue) return;
  const col = findOrCreateBodyMonthColumn_(sheet, dateValue);
  if (!col) return;
  const heightRow = findOrCreateBodyMetricRow_(sheet, player, '身長');
  const weightRow = findOrCreateBodyMetricRow_(sheet, player, '体重');
  const action = String(record.syncAction || 'create').toLowerCase();
  if (action === 'delete') {
    sheet.getRange(heightRow, col).clearContent();
    sheet.getRange(weightRow, col).clearContent();
    return;
  }
  const height = Number(record.height || 0);
  const weight = Number(record.weight || 0);
  if (height > 0) sheet.getRange(heightRow, col).setValue(height);
  if (weight > 0) sheet.getRange(weightRow, col).setValue(weight);
}

function normalizeAndSortBodyMonthColumns_(sheet) {
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  if (lastCol < 3 || lastRow < 1) return;
  const headers = sheet.getRange(1, 3, 1, lastCol - 2).getDisplayValues()[0];
  const order = headers.map((h, i) => ({ i, h: monthKeyFromValue_(h) || String(h || '').trim(), sort: monthSortValue_(h) }))
    .filter(x => x.h)
    .sort((a, b) => a.sort - b.sort || a.i - b.i);
  if (!order.length) return;
  const all = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const rewritten = all.map((row, rowIndex) => {
    const base = [row[0], row[1]];
    order.forEach(x => base.push(rowIndex === 0 ? x.h : row[x.i + 2]));
    return base;
  });
  sheet.getRange(1, 1, lastRow, lastCol).clearContent();
  sheet.getRange(1, 1, rewritten.length, rewritten[0].length).setValues(rewritten);
  if (rewritten[0].length < lastCol) sheet.deleteColumns(rewritten[0].length + 1, lastCol - rewritten[0].length);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
}

function migrateLegacyGrowthSheets_(ss) {
  const body = ss.getSheetByName(SHEET_BODY_MATRIX);
  const legacyRecords = ss.getSheetByName('growth_records');
  if (legacyRecords && legacyRecords.getLastRow() >= 2) {
    const byId = new Map();
    legacyRecords.getRange(2, 1, legacyRecords.getLastRow() - 1, 12).getValues().forEach(row => {
      const action = String(row[1] || '');
      const id = String(row[2] || '');
      if (!id) return;
      if (action === 'delete') { byId.delete(id); return; }
      byId.set(id, {
        player: String(row[6] || '').trim(),
        date: row[3] || row[5] || row[10],
        height: Number(row[8] || 0),
        weight: Number(row[9] || 0),
        syncAction: 'create'
      });
    });
    byId.forEach(r => writeGrowthRecordToBodyMatrix_(body, r));
  }
  const legacyInput = ss.getSheetByName('成長記録入力');
  if (legacyInput && legacyInput.getLastRow() >= 2) {
    legacyInput.getRange(2, 1, legacyInput.getLastRow() - 1, 6).getValues().forEach(row => {
      writeGrowthRecordToBodyMatrix_(body, {
        date: row[0], player: String(row[1] || '').trim(),
        height: Number(row[3] || 0), weight: Number(row[4] || 0), syncAction: 'create'
      });
    });
  }
  normalizeAndSortBodyMonthColumns_(body);
  if (legacyRecords) ss.deleteSheet(legacyRecords);
  if (legacyInput) ss.deleteSheet(legacyInput);
}

function appendAgilityRecords_(records) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_AGILITY_MATRIX);
  syncAgilityMatrixPlayers_(sheet);
  (records || []).forEach(r => writeAgilityRecordToMatrix_(sheet, r));
  normalizeAndSortMatrixMonthColumns_(sheet);
  log_('APPEND_AGILITY_MATRIX', `${(records || []).length} records`);
}

function findOrCreateAgilityMetricRow_(sheet, player, metric) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues();
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0] || '').trim() === player && String(values[i][1] || '').trim() === metric) return i + 2;
    }
  }
  const row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, 2).setValues([[player, metric]]);
  return row;
}

function findOrCreateAgilityMonthColumn_(sheet, dateValue) {
  const monthKey = monthKeyFromValue_(dateValue);
  if (!monthKey) return 0;
  const lastCol = Math.max(sheet.getLastColumn(), 2);
  if (lastCol >= 3) {
    const headers = sheet.getRange(1, 3, 1, lastCol - 2).getDisplayValues()[0];
    for (let i = 0; i < headers.length; i++) {
      if (monthKeyFromValue_(headers[i]) === monthKey || String(headers[i] || '').trim() === monthKey) return i + 3;
    }
  }
  const col = lastCol + 1;
  sheet.getRange(1, col).setValue(monthKey);
  return col;
}

function writeAgilityRecordToMatrix_(sheet, record) {
  const player = String(record.player || '').trim();
  const metric = String(record.type || record.metric || '').trim();
  const dateValue = record.date || record.measuredAtIso || record.createdAt;
  if (!player || !metric || !dateValue) return;
  const col = findOrCreateAgilityMonthColumn_(sheet, dateValue);
  if (!col) return;
  const row = findOrCreateAgilityMetricRow_(sheet, player, metric);
  const action = String(record.syncAction || 'create').toLowerCase();
  if (action === 'delete') {
    sheet.getRange(row, col).clearContent();
    return;
  }
  const value = Number(record.value || 0);
  if (value > 0) sheet.getRange(row, col).setValue(value);
}

function normalizeAndSortMatrixMonthColumns_(sheet) {
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  if (lastCol < 3 || lastRow < 1) return;
  const headers = sheet.getRange(1, 3, 1, lastCol - 2).getDisplayValues()[0];
  const order = headers.map((h, i) => ({ i, h: monthKeyFromValue_(h) || String(h || '').trim(), sort: monthSortValue_(h) }))
    .filter(x => x.h)
    .sort((a, b) => a.sort - b.sort || a.i - b.i);
  if (!order.length) return;
  const all = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const rewritten = all.map((row, rowIndex) => {
    const base = [row[0], row[1]];
    order.forEach(x => base.push(rowIndex === 0 ? x.h : row[x.i + 2]));
    return base;
  });
  sheet.getRange(1, 1, lastRow, lastCol).clearContent();
  sheet.getRange(1, 1, rewritten.length, rewritten[0].length).setValues(rewritten);
  if (rewritten[0].length < lastCol) sheet.deleteColumns(rewritten[0].length + 1, lastCol - rewritten[0].length);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
}

function migrateLegacyAgilitySheets_(ss) {
  const matrix = ss.getSheetByName(SHEET_AGILITY_MATRIX);
  const legacyRecords = ss.getSheetByName('agility_records');
  if (legacyRecords && legacyRecords.getLastRow() >= 2) {
    const byId = new Map();
    legacyRecords.getRange(2, 1, legacyRecords.getLastRow() - 1, 11).getValues().forEach(row => {
      const action = String(row[1] || '').toLowerCase();
      const id = String(row[2] || '').trim();
      if (!id) return;
      if (action === 'delete') { byId.delete(id); return; }
      byId.set(id, {
        date: row[3], player: String(row[4] || '').trim(),
        type: String(row[6] || '').trim(), value: Number(row[7] || 0),
        syncAction: 'create'
      });
    });
    byId.forEach(r => writeAgilityRecordToMatrix_(matrix, r));
  }
  const legacyInput = ss.getSheetByName('アジリティ入力');
  if (legacyInput && legacyInput.getLastRow() >= 2) {
    legacyInput.getRange(2, 1, legacyInput.getLastRow() - 1, 7).getValues().forEach(row => {
      writeAgilityRecordToMatrix_(matrix, {
        date: row[0], player: String(row[1] || '').trim(),
        type: String(row[3] || '').trim(), value: Number(row[4] || 0),
        syncAction: 'create'
      });
    });
  }
  normalizeAndSortMatrixMonthColumns_(matrix);
  if (legacyRecords) ss.deleteSheet(legacyRecords);
  if (legacyInput) ss.deleteSheet(legacyInput);
}

function playerCategoryByName_() {
  const map = new Map();
  listPlayers_().players.forEach(p => { if (!map.has(p.name)) map.set(p.name, p.category); });
  return map;
}
function getActiveGrowthRecords_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const matrix = ss.getSheetByName(SHEET_BODY_MATRIX);
  if (!matrix || matrix.getLastRow() < 2 || matrix.getLastColumn() < 3) return [];
  const catMap = playerCategoryByName_();
  const values = matrix.getDataRange().getValues();
  const headers = values[0];
  const combined = new Map();
  for (let r = 1; r < values.length; r++) {
    const player = String(values[r][0] || '').trim();
    const metric = String(values[r][1] || '').trim();
    if (!player || !metric) continue;
    for (let c = 2; c < headers.length; c++) {
      const date = matrixHeaderDate_(headers[c]);
      const value = Number(values[r][c] || 0);
      if (!date || !value) continue;
      const key = `${player}|${date}`;
      const current = combined.get(key) || { player, date, height: 0, weight: 0 };
      if (metric.indexOf('身長') >= 0) current.height = value;
      else if (metric.indexOf('体重') >= 0) current.weight = value;
      combined.set(key, current);
    }
  }
  return Array.from(combined.values()).map(x => ({
    id: `body-matrix-${x.player}-${x.date}`,
    date: x.date,
    measuredAtIso: x.date,
    measuredAtDisplay: x.date,
    player: x.player,
    category: catMap.get(x.player) || '',
    height: x.height,
    weight: x.weight,
    createdAt: ''
  })).filter(r => r.player && (r.height > 0 || r.weight > 0));
}
function getActiveAgilityRecords_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const matrix = ss.getSheetByName(SHEET_AGILITY_MATRIX);
  if (!matrix || matrix.getLastRow() < 2 || matrix.getLastColumn() < 3) return [];
  const catMap = playerCategoryByName_();
  const values = matrix.getDataRange().getValues();
  const headers = values[0];
  const unitMap = {'シャトルラン':'回','反復横跳び':'回','L字コーンドリル':'秒','垂直跳び':'cm'};
  const records = [];
  for (let r = 1; r < values.length; r++) {
    const player = String(values[r][0] || '').trim();
    const type = String(values[r][1] || '').trim();
    if (!player || !type) continue;
    for (let c = 2; c < headers.length; c++) {
      const date = matrixHeaderDate_(headers[c]);
      const value = Number(values[r][c] || 0);
      if (!date || !value) continue;
      records.push({
        id: `agility-matrix-${player}-${type}-${date}`,
        date,
        player,
        category: catMap.get(player) || '',
        type,
        value,
        unit: unitMap[type] || '',
        createdAt: ''
      });
    }
  }
  return records;
}
function growthRecords_(playerId) {
  const player = findPlayer_(playerId); if (!player) return { status:'error', message:'選手が見つかりません。' };
  const records = getActiveGrowthRecords_().filter(r=>r.player===player.name && r.category===player.category).sort((a,b)=>growthSortKey_(a).localeCompare(growthSortKey_(b))).map(r=>({...r, syncAction:'cloud'}));
  return { status:'ok', player, records };
}

function growthSummary_() {
  const players = listPlayers_().players;
  const records = getActiveGrowthRecords_();
  const byPlayer = new Map();
  records.forEach(r => {
    const key = `${r.player}|${r.category}`;
    if (!byPlayer.has(key)) byPlayer.set(key, []);
    byPlayer.get(key).push(r);
  });
  const latestByPlayer = players.map(p => {
    const key = `${p.name}|${p.category}`;
    const list = (byPlayer.get(key) || []).slice().sort((a,b) => growthSortKey_(a).localeCompare(growthSortKey_(b)));
    const latest = list.length ? list[list.length - 1] : null;
    return {
      playerId: p.playerId,
      player: p.name,
      category: p.category,
      height: latest ? Number(latest.height || 0) : 0,
      weight: latest ? Number(latest.weight || 0) : 0,
      date: latest ? normalizeDateString_(latest.date || latest.measuredAtIso || latest.createdAt) : '',
      recordsCount: list.length,
      records: list
    };
  });
  const avg = (items) => {
    const valid = items.filter(x => Number(x.height||0) > 0 && Number(x.weight||0) > 0);
    const height = valid.length ? Math.round((valid.reduce((s,x)=>s+Number(x.height||0),0)/valid.length)*10)/10 : 0;
    const weight = valid.length ? Math.round((valid.reduce((s,x)=>s+Number(x.weight||0),0)/valid.length)*10)/10 : 0;
    return { count: valid.length, height, weight };
  };
  const categories = ['男子U13','男子U14','男子U15','女子U13','女子U14','女子U15'];
  const categoryAverages = categories.map(category => ({ category, ...avg(latestByPlayer.filter(x => x.category === category)) }));
  const genderAverages = [
    { gender:'男子全カテゴリー', ...avg(latestByPlayer.filter(x => String(x.category||'').indexOf('男子') === 0)) },
    { gender:'女子全カテゴリー', ...avg(latestByPlayer.filter(x => String(x.category||'').indexOf('女子') === 0)) }
  ];
  return { status:'ok', generatedAt:new Date().toISOString(), players:latestByPlayer, categoryAverages, genderAverages };
}

function agilityRecords_(playerId) {
  const player = findPlayer_(playerId); if (!player) return { status:'error', message:'選手が見つかりません。' };
  const records = getActiveAgilityRecords_().filter(r=>r.player===player.name && r.category===player.category).sort((a,b)=>String(a.date).localeCompare(String(b.date))).map(r=>({...r, syncAction:'cloud'}));
  return { status:'ok', player, records };
}


function agilitySummary_() {
  const players = listPlayers_().players;
  const records = getActiveAgilityRecords_();
  return {
    status:'ok',
    generatedAt:new Date().toISOString(),
    players:players.map(p => ({
      playerId:p.playerId, player:p.name, category:p.category,
      records:records.filter(r => r.player === p.name && r.category === p.category)
        .sort((a,b)=>String(a.date||'').localeCompare(String(b.date||'')) || String(a.createdAt||'').localeCompare(String(b.createdAt||'')))
    }))
  };
}

function agilityRankings_() {
  const records = getActiveAgilityRecords_();
  const types = Array.from(new Set(records.map(r=>String(r.type||'').trim()).filter(Boolean)));
  const unitMap = {'シャトルラン':'回','反復横跳び':'回','L字コーンドリル':'秒','垂直跳び':'cm'};
  const lowerIsBetter = {'L字コーンドリル': true};
  const result = {};
  types.forEach(type => {
    const typeRecords = records.filter(r => r.type === type && Number(r.value || 0) > 0);
    const grouped = new Map();
    typeRecords.forEach(r => {
      const key = `${r.player}|${r.category}`;
      if (!grouped.has(key)) grouped.set(key, []);
      grouped.get(key).push(r);
    });
    const currentRows = [];
    const previousRows = [];
    grouped.forEach(list => {
      list = list.slice().sort((a,b) => String(a.date || '').localeCompare(String(b.date || '')) || String(a.createdAt || '').localeCompare(String(b.createdAt || '')));
      const latest = list[list.length - 1];
      const previous = list.length >= 2 ? list[list.length - 2] : null;
      if (latest) currentRows.push({ player:latest.player, category:latest.category, type:latest.type, value:Number(latest.value||0), unit:latest.unit||unitMap[type]||'', date:latest.date||'', previousValue:previous?Number(previous.value||0):0, previousDate:previous?(previous.date||''):'' });
      if (previous) previousRows.push({ player:previous.player, category:previous.category, type:previous.type, value:Number(previous.value||0), unit:previous.unit||unitMap[type]||'', date:previous.date||'' });
    });
    const sortFn = lowerIsBetter[type] ? ((a,b)=>Number(a.value||0)-Number(b.value||0)) : ((a,b)=>Number(b.value||0)-Number(a.value||0));
    currentRows.sort(sortFn); previousRows.sort(sortFn);
    const prevRankMap = new Map();
    previousRows.forEach((r,i)=>prevRankMap.set(`${r.player}|${r.category}`, i+1));
    result[type] = currentRows.map((r,i)=>Object.assign({}, r, {rank:i+1, previousRank:prevRankMap.get(`${r.player}|${r.category}`)||0}));
  });
  return { status:'ok', generatedAt:new Date().toISOString(), rankings:result };
}

function normalizeDateString_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const s = String(value || '').trim();
  if (!s) return '';
  let m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if (m) return `${m[1]}-${String(Number(m[2])).padStart(2,'0')}-${String(Number(m[3])).padStart(2,'0')}`;
  m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/);
  if (m) return `${m[3]}-${String(Number(m[1])).padStart(2,'0')}-${String(Number(m[2])).padStart(2,'0')}`;
  const d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return s;
}
function growthSortKey_(r) {
  return normalizeDateString_(r.date || r.measuredAtIso || r.createdAt);
}
function formatDate_(value) { return normalizeDateString_(value); }
function formatDateTime_(value) { if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'); return String(value || ''); }
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function log_(type, message) { try { const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOGS); if (sheet) sheet.appendRow([new Date(), type, message]); } catch(e) {} }
