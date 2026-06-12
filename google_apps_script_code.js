/**
 * NINJA SHOOTING AVERAGE v11.0
 * 選手用・コーチ用・成長記録・アジリティー記録対応 Apps Script
 */
const SHEET_RECORDS = 'shooting_records';
const SHEET_PLAYERS = 'players';
const SHEET_LOGS = 'logs';
const SHEET_SUMMARY = 'summary';
const SHEET_GROWTH = 'growth_records';
const SHEET_AGILITY = 'agility_records';
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
    else if (action === 'updatePlayerCategory') result = updatePlayerCategory_(p);
    else if (action === 'dashboard') result = dashboard_();
    else { updateSummary_(); result = { status: 'ok', app: 'NINJA SHOOTING AVERAGE v11.0' }; }
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
  ensureSheet_(ss, SHEET_GROWTH, ['送信日時','同期種別','記録ID','日付','測定日時表示','測定日時ISO','選手名','カテゴリー','身長cm','体重kg','作成日時','削除日時']);
  ensureSheet_(ss, SHEET_AGILITY, ['送信日時','同期種別','記録ID','日付','選手名','カテゴリー','種目','記録','単位','作成日時','削除日時']);
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_GROWTH);
  const values = (records||[]).map(r => [new Date(), r.syncAction || 'create', r.id || '', r.date || '', r.measuredAtDisplay || '', r.measuredAtIso || r.createdAt || '', r.player || '', r.category || '', Number(r.height || 0), Number(r.weight || 0), r.createdAt || '', r.deletedAt || '']);
  if (values.length) sheet.getRange(sheet.getLastRow()+1,1,values.length,values[0].length).setValues(values);
  log_('APPEND_GROWTH', `${values.length} records`);
}
function appendAgilityRecords_(records) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_AGILITY);
  const values = (records||[]).map(r => [new Date(), r.syncAction || 'create', r.id || '', r.date || '', r.player || '', r.category || '', r.type || '', Number(r.value || 0), r.unit || '', r.createdAt || '', r.deletedAt || '']);
  if (values.length) sheet.getRange(sheet.getLastRow()+1,1,values.length,values[0].length).setValues(values);
  log_('APPEND_AGILITY', `${values.length} records`);
}
function getActiveGrowthRecords_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_GROWTH);
  const last = sheet.getLastRow(); if (last < 2) return [];
  const values = sheet.getRange(2,1,last-1,12).getValues(); const byId = new Map();
  values.forEach(row => { const syncType=String(row[1]||''); const id=String(row[2]||''); if(!id)return; if(syncType==='delete'){byId.delete(id);return;} byId.set(id,{id: id,date:normalizeDateString_(row[3]||row[5]||row[10]),measuredAtDisplay:String(row[4]||''),measuredAtIso:normalizeDateString_(row[5]||row[3]||row[10]),player:String(row[6]||''),category:String(row[7]||''),height:Number(row[8]||0),weight:Number(row[9]||0),createdAt:formatDateTime_(row[10])}); });
  return Array.from(byId.values()).filter(r=>r.player && r.height>0 && r.weight>0);
}
function getActiveAgilityRecords_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_AGILITY);
  const last = sheet.getLastRow(); if (last < 2) return [];
  const values = sheet.getRange(2,1,last-1,11).getValues(); const byId = new Map();
  values.forEach(row => { const syncType=String(row[1]||''); const id=String(row[2]||''); if(!id)return; if(syncType==='delete'){byId.delete(id);return;} byId.set(id,{id,date:formatDate_(row[3]),player:String(row[4]||''),category:String(row[5]||''),type:String(row[6]||''),value:Number(row[7]||0),unit:String(row[8]||''),createdAt:formatDateTime_(row[9])}); });
  return Array.from(byId.values()).filter(r=>r.player && r.type && r.value>0);
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
