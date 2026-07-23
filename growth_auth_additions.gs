/**
 * NINJA 成長記録Webアプリ用追加コード
 * 既存コード.gsの末尾へ追加してください。
 * 既存 doGet(e) の action 判定へ次の3行を追加：
 * else if (action === 'growthLogin') result = growthLogin_(p);
 * else if (action === 'growthDashboard') result = growthDashboard_(p);
 * else if (action === 'growthLogout') result = growthLogout_(p);
 */
const GROWTH_SESSION_TTL_MS_ = 30 * 24 * 60 * 60 * 1000;
const GROWTH_SESSION_PREFIX_ = 'GROWTH_SESSION_';

function growthLogin_(p) {
  const login = loginPlayer_(p);
  if (!login || login.status !== 'ok') return login;
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const expiresAt = Date.now() + GROWTH_SESSION_TTL_MS_;
  PropertiesService.getScriptProperties().setProperty(
    GROWTH_SESSION_PREFIX_ + token,
    JSON.stringify({ playerId:String(login.player.playerId || ''), expiresAt:expiresAt })
  );
  return { status:'ok', token:token, expiresAt:expiresAt, player:login.player };
}

function growthLogout_(p) {
  const token = String(p.token || '').trim();
  if (token) PropertiesService.getScriptProperties().deleteProperty(GROWTH_SESSION_PREFIX_ + token);
  return { status:'ok' };
}

function growthDashboard_(p) {
  const sessionResult = growthSession_(p.token);
  if (!sessionResult.ok) return { status:'error', code:'SESSION_EXPIRED', message:'ログイン期限が切れています。' };

  const playerId = sessionResult.session.playerId;
  const detail = playerDetail_(playerId);
  const growth = growthRecords_(playerId);
  const agility = agilityRecords_(playerId);
  const playerRecords = playerRecords_(playerId);
  if (!detail || detail.status !== 'ok') return detail;
  if (!growth || growth.status !== 'ok') return growth;
  if (!agility || agility.status !== 'ok') return agility;
  if (!playerRecords || playerRecords.status !== 'ok') return playerRecords;

  const shootingRecords = playerRecords.records || [];
  const monthlyMap = {};
  shootingRecords.forEach(function(r){
    const month = String(r.date || '').slice(0,7);
    if (!month) return;
    if (!monthlyMap[month]) monthlyMap[month] = { month:month, made:0, attempts:0 };
    monthlyMap[month].made += Number(r.made || 0);
    monthlyMap[month].attempts += Number(r.attempts || 0);
  });
  const monthly = Object.keys(monthlyMap).sort().slice(-12).map(function(key){
    const x = monthlyMap[key];
    x.rate = x.attempts > 0 ? Math.round(x.made / x.attempts * 100) : 0;
    return x;
  });

  const growthRows = (growth.records || []).slice().sort(function(a,b){ return String(a.date || '').localeCompare(String(b.date || '')); });
  const firstGrowth = growthRows.length ? growthRows[0] : null;
  const latestGrowth = growthRows.length ? growthRows[growthRows.length - 1] : null;

  const latestAgilityByType = {};
  (agility.records || []).forEach(function(r){
    const key = String(r.type || '');
    if (!key) return;
    if (!latestAgilityByType[key] || String(latestAgilityByType[key].date || '') <= String(r.date || '')) latestAgilityByType[key] = r;
  });

  return {
    status:'ok',
    player:detail.player,
    shooting:{
      totalMade:Number(detail.total.made || 0), totalAttempts:Number(detail.total.attempts || 0), totalRate:Number(detail.total.rate || 0),
      latestDate:shootingRecords.length ? shootingRecords.map(function(r){return String(r.date || '')}).sort().pop() : '',
      byType:detail.byType || [], byPosition:detail.byPosition || [], monthly:monthly
    },
    body:{
      latestHeight:latestGrowth ? Number(latestGrowth.height || 0) : null,
      latestWeight:latestGrowth ? Number(latestGrowth.weight || 0) : null,
      heightChange:firstGrowth && latestGrowth ? round1_(Number(latestGrowth.height || 0)-Number(firstGrowth.height || 0)) : null,
      weightChange:firstGrowth && latestGrowth ? round1_(Number(latestGrowth.weight || 0)-Number(firstGrowth.weight || 0)) : null,
      months:growthRows.map(function(r){return String(r.date || '')}),
      heights:growthRows.map(function(r){return Number(r.height || 0)}),
      weights:growthRows.map(function(r){return Number(r.weight || 0)})
    },
    agility:Object.keys(latestAgilityByType).sort().map(function(key){
      const r = latestAgilityByType[key];
      return { event:r.type, value:Number(r.value || 0), unit:String(r.unit || ''), date:String(r.date || '') };
    })
  };
}

function growthSession_(tokenValue) {
  const token = String(tokenValue || '').trim();
  if (!token) return { ok:false };
  const properties = PropertiesService.getScriptProperties();
  const key = GROWTH_SESSION_PREFIX_ + token;
  const raw = properties.getProperty(key);
  if (!raw) return { ok:false };
  try {
    const session = JSON.parse(raw);
    if (!session.playerId || Date.now() >= Number(session.expiresAt || 0)) {
      properties.deleteProperty(key);
      return { ok:false };
    }
    return { ok:true, session:session };
  } catch (e) {
    properties.deleteProperty(key);
    return { ok:false };
  }
}
function round1_(value){ return Math.round(Number(value || 0) * 10) / 10; }
