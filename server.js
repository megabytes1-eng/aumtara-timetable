// Timetable module — REST API server v3 (Express + Postgres)
const express = require('express');
const path = require('path');
const ExcelJS = require('exceljs');
const multer = require('multer');
const upload = multer({ storage: multer.memoryStorage() });
const crypto = require('crypto');
const { q, q1, run, tx, init, loadConfig, getConfigCached, hashPw, verifyPw, seedConfigForSchool } = require('./db');
let webpush=null; try{ webpush=require('web-push'); }catch(e){ console.warn('web-push unavailable — push notifications disabled:', e.message); }
let VAPID_PUB=null;
async function getMeta(k){ const r=await q1('SELECT v FROM tt_appmeta WHERE k=?',[k]); return r?r.v:null; }
async function setMeta(k,v){ await run('INSERT INTO tt_appmeta(k,v) VALUES(?,?) ON CONFLICT(k) DO UPDATE SET v=excluded.v',[k,v]); }
async function ensureVapid(){
  if(!webpush) return;
  try{
    let pub=await getMeta('vapid_pub'), priv=await getMeta('vapid_priv');
    if(!pub||!priv){ const kk=webpush.generateVAPIDKeys(); pub=kk.publicKey; priv=kk.privateKey; await setMeta('vapid_pub',pub); await setMeta('vapid_priv',priv); }
    webpush.setVapidDetails('mailto:notify@aumtara-timetable.onrender.com', pub, priv);
    VAPID_PUB=pub;
  }catch(e){ console.warn('VAPID setup failed:', e.message); }
}
// send a payload to every subscription in a list; prune subscriptions the push service rejects as gone
async function sendToSubs(subs, payload){
  if(!webpush||!subs||!subs.length) return {sent:0,failed:0};
  let sent=0,failed=0; const body=JSON.stringify(payload);
  for(const s of subs){
    try{ await webpush.sendNotification({endpoint:s.endpoint,keys:{p256dh:s.p256dh,auth:s.auth}}, body); sent++; }
    catch(e){ failed++; if(e&&(e.statusCode===404||e.statusCode===410)){ try{ await run('DELETE FROM tt_push_sub WHERE endpoint=?',[s.endpoint]); }catch(_){}} }
  }
  return {sent,failed};
}

const app = express();
// CORS — lets the standalone launch/landing page (hosted on any domain) drive the API (login + owner control panel).
// Tokens are bearer (not cookies), so a wildcard origin exposes nothing beyond what a token-holder can already do.
app.use((req,res,next)=>{ res.header('Access-Control-Allow-Origin','*'); res.header('Access-Control-Allow-Headers','Authorization, Content-Type, X-School-Id'); res.header('Access-Control-Allow-Methods','GET, POST, PUT, DELETE, OPTIONS'); if(req.method==='OPTIONS') return res.sendStatus(204); next(); });
app.use(express.json({ limit: '8mb' }));   // logos (base64 data-URLs) can exceed the 100kb default
// Serve the single-page frontend. index.html lives next to server.js (flat layout
// so the repo uploads cleanly to GitHub's web uploader, which can't preserve subfolders).
// index.html is fully self-contained (inline CSS/JS), so no other static assets are needed.
app.get('/', (_, res) => res.sendFile(path.join(__dirname, 'index.html')));
app.get('/launch', (_, res) => { res.set('Cache-Control','no-cache'); res.sendFile(path.join(__dirname, 'launch.html')); });   // public marketing landing (customers)
app.get('/panel', (_, res) => { res.set('Cache-Control','no-cache'); res.sendFile(path.join(__dirname, 'panel.html')); });    // dedicated owner control panel (separate from the school app login)

// ---------- PWA (installable web app: manifest + service worker + icons) ----------
const MANIFEST = {
  name: 'Aumtara Timetable', short_name: 'Aumtara',
  description: 'School timetable, teacher diary & reports — multi-school.',
  start_url: '/', scope: '/', display: 'standalone', orientation: 'any',
  background_color: '#1f2b57', theme_color: '#1f2b57', lang: 'en',
  icons: [
    { src: '/icon-192.png', sizes: '192x192', type: 'image/png', purpose: 'any' },
    { src: '/icon-512.png', sizes: '512x512', type: 'image/png', purpose: 'any' },
    { src: '/icon-maskable.png', sizes: '512x512', type: 'image/png', purpose: 'maskable' }
  ]
};
const SW_JS = `const CACHE='aumtara-v102';
const SHELL=['/','/manifest.webmanifest','/icon-192.png','/icon-512.png'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(CACHE).then(c=>c.addAll(SHELL)).then(()=>self.skipWaiting()).catch(()=>self.skipWaiting()));});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()));});
self.addEventListener('fetch',e=>{const req=e.request; if(req.method!=='GET') return; let url; try{url=new URL(req.url);}catch(_){return;} if(url.origin!==self.location.origin) return; if(url.pathname.startsWith('/api/')) return;
  if(req.mode==='navigate'){ e.respondWith(fetch(req).then(r=>{const cp=r.clone(); caches.open(CACHE).then(c=>c.put('/',cp)); return r;}).catch(()=>caches.match('/'))); return; }
  e.respondWith(caches.match(req).then(r=>r||fetch(req)));});
self.addEventListener('push',e=>{ let d={}; try{ d=e.data?e.data.json():{}; }catch(_){ d={body:e.data&&e.data.text?e.data.text():''}; }
  const title=d.title||'Aumtara'; const opts={body:d.body||'',icon:'/icon-192.png',badge:'/icon-192.png',data:{url:d.url||'/'}};
  e.waitUntil(self.registration.showNotification(title,opts)); });
self.addEventListener('notificationclick',e=>{ e.notification.close(); const u=(e.notification.data&&e.notification.data.url)||'/';
  e.waitUntil(self.clients.matchAll({type:'window',includeUncontrolled:true}).then(cs=>{ for(const c of cs){ if('focus' in c) return c.focus(); } if(self.clients.openWindow) return self.clients.openWindow(u); })); });`;
app.get('/manifest.webmanifest', (_, res) => { res.type('application/manifest+json'); res.send(JSON.stringify(MANIFEST)); });
app.get('/sw.js', (_, res) => { res.type('application/javascript'); res.set('Cache-Control', 'no-cache'); res.send(SW_JS); });
for (const ic of ['icon-192.png', 'icon-512.png', 'icon-maskable.png'])
  app.get('/' + ic, (_, res) => res.sendFile(path.join(__dirname, ic)));
app.get('/apple-touch-icon.png', (_, res) => res.sendFile(path.join(__dirname, 'icon-192.png')));
app.get('/apple-touch-icon-precomposed.png', (_, res) => res.sendFile(path.join(__dirname, 'icon-192.png')));

// PUBLIC read-only shared timetable (no login). Capability token → one class's weekly grid.
const esc = (s)=>String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
app.get('/share/:token', async (req,res)=>{
  try{
    const link = await q1('SELECT * FROM tt_sharelink WHERE token=?',[req.params.token]);
    const page=(inner)=>`<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Timetable</title>
      <style>body{font-family:'Segoe UI',Arial,sans-serif;background:#f4f7fc;margin:0;color:#222}.wrap{max-width:1000px;margin:0 auto;padding:18px}
      .hd{background:#1F3864;color:#fff;padding:16px 22px;display:flex;align-items:center;gap:12px}.hd img{height:34px;border-radius:6px}.hd h1{font-size:18px;margin:0}.hd .s{font-size:12px;opacity:.85}
      table{border-collapse:collapse;width:100%;font-size:13.5px;background:#fff}th,td{border:1px solid #e2e6ee;padding:8px 9px;text-align:center;vertical-align:middle}
      th{background:#1F3864;color:#fff}td.ph{background:#eef2fb;font-weight:700;color:#1F3864}.subj{font-weight:800;color:#1F3864}.tch{font-size:12px;color:#33415a;font-weight:600}.rm{font-size:11px;color:#7030A0}.free{color:#9aa0ab;font-style:italic}
      .note{color:#6b7280;font-size:12px;margin:10px 2px}.card{background:#fff;border:1px solid #e2e6ee;border-radius:12px;padding:14px;overflow:auto}</style></head>
      <body>${inner}<div class="wrap"><div class="note">Read-only shared timetable · Aumtara</div></div></body></html>`;
    if(!link){ res.status(404).send(page('<div class="wrap"><h2>Link not found or expired</h2></div>')); return; }
    const sid=link.school_id;
    const school=await q1('SELECT name,logo FROM tt_school WHERE id=?',[sid]);
    const cls=await q1('SELECT name FROM tt_class WHERE id=? AND school_id=?',[link.target_id,sid]);
    if(!cls){ res.status(404).send(page('<div class="wrap"><h2>Not found</h2></div>')); return; }
    const subs={}; (await q('SELECT id,name FROM tt_subject WHERE school_id=?',[sid])).forEach(s=>subs[s.id]=s.name);
    const tch={}; (await q('SELECT id,name FROM tt_teacher WHERE school_id=?',[sid])).forEach(x=>tch[x.id]=x.name);
    const rms={}; (await q('SELECT id,name FROM tt_room WHERE school_id=?',[sid])).forEach(r=>rms[r.id]=r.name);
    const cells=await q('SELECT day_of_week,period_index,subject_id,teacher_id,room_id,week_index FROM tt_timetable WHERE class_id=? AND school_id=?',[link.target_id,sid]);
    const DN=['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'];
    const wd=workingDaysArr(sid); let maxP=1; wd.forEach(d=>{maxP=Math.max(maxP, teachingSlots(d,sid).length);});
    const cfg=getConfig(sid)||{}; const cyc = rotMode(sid) ? 1 : Math.max(1,Math.min(4,parseInt(cfg.cycle_weeks,10)||1));
    const wkLabels=(cfg.week_labels||'').split(',').map(x=>x.trim()).filter(Boolean);
    const wkName=i=> wkLabels[i] || (cyc===2?['Week A','Week B'][i] : 'Week '+(i+1));
    const dayLbl = d => rotMode(sid) ? rotLabel(sid,d) : (DN[d]||('Day '+d));
    function gridFor(wk){
      const cmap={}; cells.filter(c=>(c.week_index||0)===wk).forEach(c=>cmap[c.day_of_week+'_'+c.period_index]=c);
      let head='<th>Day</th>'; for(let p=0;p<maxP;p++) head+='<th>P'+(p+1)+'</th>';
      let rows='';
      for(const d of wd){ const ts=teachingSlots(d,sid); let r='<td class="ph">'+esc(dayLbl(d))+'</td>';
        for(let p=0;p<maxP;p++){ if(p>=ts.length){ r+='<td></td>'; continue; } const c=cmap[d+'_'+p];
          if(c&&c.subject_id){ r+='<td><div class="subj">'+esc(subs[c.subject_id]||'')+'</div><div class="tch">'+esc(tch[c.teacher_id]||'—')+(c.room_id?' · '+esc(rms[c.room_id]||''):'')+'</div></td>'; }
          else r+='<td class="free">—</td>'; }
        rows+='<tr>'+r+'</tr>'; }
      const title = cyc>1 ? '<h3 style="margin:0 0 8px;color:#1F3864">'+esc(wkName(wk))+'</h3>' : '';
      return '<div class="card">'+title+'<table><tr>'+head+'</tr>'+rows+'</table></div>';
    }
    let grids=''; for(let wk=0; wk<cyc; wk++) grids+=gridFor(wk);
    const logo=school&&school.logo?`<img src="${esc(school.logo)}" alt="">`:'';
    const inner=`<div class="hd">${logo}<div><h1>${esc((school&&school.name)||'School')}</h1><div class="s">${esc(cls.name)} · Timetable</div></div></div>
      <div class="wrap">${grids}</div>`;
    res.send(page(inner));
  }catch(e){ res.status(500).send('Error'); }
});

// async route wrapper — turns rejected promises into a clean 500 instead of a crash
const h = fn => (req, res) => Promise.resolve(fn(req, res)).catch(e => {
  console.error(e);
  if (!res.headersSent) res.status(500).json({ error: String((e && e.message) || e) });
});

// ========================= AUTH (role-based login) =========================
async function currentUser(req){
  const tok = (req.headers.authorization||'').replace(/^Bearer\s+/i,'').trim() || (req.query && req.query.token) || '';
  if(!tok) return null;
  const s = await q1('SELECT user_id FROM tt_session WHERE token=?',[tok]);
  if(!s) return null;
  const u = await q1('SELECT id,name,role,login_id,email,mobile,qualification,main_subject_id,school_id,is_owner,group_scope FROM tt_user WHERE id=? AND active=1',[s.user_id]);
  // Self-heal: the platform-owner account must always be a master. If it was accidentally demoted, restore it on any request.
  if(u && Number(u.is_owner)===1 && u.role!=='master'){ await run("UPDATE tt_user SET role='master' WHERE id=?",[u.id]); u.role='master'; }
  return u;
}
// public: sign in — identifier can be login_id OR email OR mobile
// ---------- TOTP (RFC 6238) for authenticator-app 2FA ----------
const _B32='ABCDEFGHIJKLMNOPQRSTUVWXYZ234567';
function b32encode(buf){ let bits='',out=''; for(const byte of buf){ bits+=byte.toString(2).padStart(8,'0'); } for(let i=0;i+5<=bits.length;i+=5){ out+=_B32[parseInt(bits.substr(i,5),2)]; } return out; }
function b32decode(s){ s=String(s||'').toUpperCase().replace(/[^A-Z2-7]/g,''); let bits=''; for(const c of s){ const v=_B32.indexOf(c); if(v<0) continue; bits+=v.toString(2).padStart(5,'0'); } const bytes=[]; for(let i=0;i+8<=bits.length;i+=8){ bytes.push(parseInt(bits.substr(i,8),2)); } return Buffer.from(bytes); }
function totpAt(secret, tstep){ const key=b32decode(secret); const buf=Buffer.alloc(8); let c=tstep; for(let i=7;i>=0;i--){ buf[i]=c&0xff; c=Math.floor(c/256); } const hmac=crypto.createHmac('sha1',key).update(buf).digest(); const off=hmac[hmac.length-1]&0xf; const bin=((hmac[off]&0x7f)<<24)|((hmac[off+1]&0xff)<<16)|((hmac[off+2]&0xff)<<8)|(hmac[off+3]&0xff); return String(bin%1000000).padStart(6,'0'); }
function totpVerify(secret, code){ code=String(code||'').replace(/\D/g,''); if(code.length!==6||!secret) return false; const step=Math.floor(Date.now()/1000/30); for(let w=-1;w<=1;w++){ if(totpAt(secret, step+w)===code) return true; } return false; }
function genTotpSecret(){ return b32encode(crypto.randomBytes(20)); }
function genBackupCodes(n){ const codes=[]; for(let i=0;i<(n||8);i++){ codes.push(crypto.randomBytes(5).toString('hex')); } return codes; }

app.post('/api/login', h(async (req,res)=>{
  const b=req.body||{};
  const idf=String(b.login_id||'').trim();
  const u=await q1('SELECT * FROM tt_user WHERE (lower(login_id)=lower(?) OR lower(email)=lower(?) OR mobile=?) AND active=1 LIMIT 1',[idf,idf,idf]);
  if(!u || !verifyPw(b.password, u.password_hash)){ res.status(401).json({error:'invalid credentials'}); return; }
  // Two-factor: if this account has 2FA on, require a valid authenticator code (or a one-time backup code)
  if(Number(u.totp_enabled)===1){
    const code=String(b.code||'').replace(/\s/g,'');
    if(!code){ res.status(401).json({error:'2FA code required', needs_2fa:true}); return; }
    let ok=totpVerify(u.totp_secret, code);
    if(!ok){
      let bc=[]; try{ bc=JSON.parse(u.backup_codes||'[]'); }catch(_){ bc=[]; }
      const idx=bc.findIndex(hc=>verifyPw(code.toLowerCase(), hc));
      if(idx>=0){ ok=true; bc.splice(idx,1); await run('UPDATE tt_user SET backup_codes=? WHERE id=?',[JSON.stringify(bc), u.id]); }
    }
    if(!ok){ res.status(401).json({error:'Invalid 2FA code', needs_2fa:true}); return; }
  }
  // a deactivated (inactive) school suspends all its users' logins — reactivate from the SaaS panel (platform owner not tied to a school, so unaffected)
  if(u.school_id){ const sch=await q1('SELECT active FROM tt_school WHERE id=?',[u.school_id]); if(sch && +sch.active===0){ res.status(403).json({error:'This school account is inactive. Please contact the administrator.'}); return; } }
  // Self-heal: a school must always have at least one Admin. If an accidental role change/deletion left it
  // with none, promote whoever logs in (they are clearly staff trying to administer) so nobody is locked out.
  if(u.school_id && !['admin','master'].includes(u.role)){
    const hasAdmin=await q1("SELECT 1 FROM tt_user WHERE school_id=? AND active=1 AND role IN ('admin','master') LIMIT 1",[u.school_id]);
    if(!hasAdmin){ await run("UPDATE tt_user SET role='admin' WHERE id=?",[u.id]); u.role='admin'; }
  }
  const token=crypto.randomBytes(24).toString('hex');
  await run('INSERT INTO tt_session(token,user_id,created_at) VALUES(?,?,now()::text)',[token,u.id]);
  res.json({ token, user:{ id:u.id, name:u.name, role:u.role, login_id:u.login_id, school_id:u.school_id } });
}));
// public: self sign-up — creates a school admin account and signs in
app.post('/api/signup', h(async (req,res)=>{
  const b=req.body||{};
  const username=String(b.username||'').trim();
  const password=String(b.password||'');
  const school=String(b.school_name||'').trim();
  const email=String(b.email||'').trim()||null;
  const mobile=String(b.mobile||'').trim()||null;
  if(!school || !username || !password){ res.status(400).json({error:'school name, user name and password are required'}); return; }
  if(await q1('SELECT 1 FROM tt_user WHERE lower(login_id)=lower(?)',[username])){ res.status(400).json({error:'user name already taken'}); return; }
  if(email && await q1('SELECT 1 FROM tt_user WHERE lower(email)=lower(?)',[email])){ res.status(400).json({error:'email already registered'}); return; }
  if(mobile && await q1('SELECT 1 FROM tt_user WHERE mobile=?',[mobile])){ res.status(400).json({error:'mobile already registered'}); return; }
  const board=String(b.board||'').trim()||null;
  const medium=String(b.medium||'').trim()||null;
  // pending=true (paid signups from the landing page): create the school INACTIVE — no app access until the owner verifies payment & activates it from the SaaS panel.
  const pending = b.pending===true || b.pending==='true' || b.pending===1;
  const sch=await q1('INSERT INTO tt_school(name,board,medium,active,created_at) VALUES(?,?,?,?,now()::text) RETURNING id',[school,board,medium, pending?0:1]);
  for(const f of ['address','principal_name','pin_code','work_phone']){   // optional school profile fields captured at signup
    if(b[f]!==undefined && String(b[f]).trim()!=='') await run('UPDATE tt_school SET '+f+'=? WHERE id=?',[String(b[f]).trim(), sch.id]);
  }
  await seedConfigForSchool(sch.id, school, board, medium);
  const row=await q1(`INSERT INTO tt_user(name,role,login_id,email,mobile,password_hash,active,created_at,school_id)
     VALUES(?,?,?,?,?,?,1,now()::text,?) RETURNING id`, [b.name||username,'admin',username,email,mobile,hashPw(password),sch.id]);
  if(pending){
    // leave the school inactive; log a pending-payment entry so the owner sees it in the Sales panel
    try{ await run('INSERT INTO tt_enquiry(name,school,phone,email,plan,message,handled,ts) VALUES(?,?,?,?,?,?,0,now()::text)',
      [b.name||username, school, mobile||null, email||null, String(b.plan||'').trim()||null, 'Paid signup — pending verification. Activate school #'+sch.id+' in the SaaS panel after confirming payment.']); }catch(e){}
    res.json({ pending:true, login_id:username, school_id:sch.id });
    return;
  }
  const token=crypto.randomBytes(24).toString('hex');
  await run('INSERT INTO tt_session(token,user_id,created_at) VALUES(?,?,now()::text)',[token,row.id]);
  res.json({ token, user:{ id:row.id, name:b.name||username, role:'admin', login_id:username, school_id:sch.id } });
}));
// public: sign out (harmless if unauthenticated)
app.post('/api/logout', h(async (req,res)=>{
  const tok=(req.headers.authorization||'').replace(/^Bearer\s+/i,'').trim();
  if(tok) await run('DELETE FROM tt_session WHERE token=?',[tok]);
  res.json({ ok:true });
}));
// public: buy / payment settings (only NON-secret values — a Razorpay *page link* URL, UPI id, WhatsApp no., contact email). Used by the landing-page Buy modal.
app.get('/api/pay-settings', h(async (req,res)=>{
  let s={}; try{ const raw=await getMeta('pay_settings'); s=raw?JSON.parse(raw):{}; }catch(_){ s={}; }
  res.json({
    gateway_url: s.gateway_url||'',
    upi_id: s.upi_id||'',
    whatsapp: s.whatsapp||'',
    contact_email: s.contact_email||'',
    payee_name: s.payee_name||'Aumtara Timetable',
    enable_gateway: s.enable_gateway===0?0:1,
    enable_upi: s.enable_upi===0?0:1,
    enable_enquiry: s.enable_enquiry===0?0:1,
    note: s.note||''
  });
}));
// public: submit a buy enquiry (name + phone required); the owner sees these in the control panel
app.post('/api/enquiry', h(async (req,res)=>{
  const b=req.body||{};
  const name=String(b.name||'').trim(), phone=String(b.phone||'').trim();
  if(!name||!phone){ res.status(400).json({error:'name and phone are required'}); return; }
  await run('INSERT INTO tt_enquiry(name,school,phone,email,plan,message,handled,ts) VALUES(?,?,?,?,?,?,0,now()::text)',
    [name, String(b.school||'').trim()||null, phone, String(b.email||'').trim()||null, String(b.plan||'').trim()||null, String(b.message||'').trim()||null]);
  res.json({ ok:true });
}));
// gate: every other /api route requires a valid session
app.use('/api', (req,res,next)=>{
  currentUser(req).then(u=>{ if(!u){ res.status(401).json({error:'auth required'}); return; } req.user=u; next(); })
    .catch(e=>{ console.error(e); res.status(500).json({error:String((e&&e.message)||e)}); });
});
// resolve the current school for this request.
// master can switch schools via ?school= / X-School-Id header; everyone else is locked to their own school.
app.use('/api', (req,res,next)=>{
  (async ()=>{
    const raw = (req.query && req.query.school) || req.headers['x-school-id'];
    let sid;
    if (req.user.role === 'master') sid = (raw!=null && raw!=='') ? Number(raw) : (req.user.school_id || null);
    else if (req.user.group_scope) {
      // institution-scoped admin — may act on any school in their group, verified against group_name
      if (raw!=null && raw!=='') {
        const g = await q1('SELECT group_name FROM tt_school WHERE id=?',[Number(raw)]);
        sid = (g && String(g.group_name||'').trim().toLowerCase() === String(req.user.group_scope).trim().toLowerCase()) ? Number(raw) : (req.user.school_id || null);
      } else sid = req.user.school_id || null;
    }
    else sid = req.user.school_id || null;
    if (sid == null) { const s = await q1('SELECT id FROM tt_school WHERE active=1 ORDER BY id LIMIT 1'); sid = s ? s.id : null; }
    req.sid = sid;
    next();
  })().catch(e=>{ console.error(e); res.status(500).json({error:String((e&&e.message)||e)}); });
});
app.get('/api/me', h(async (req,res)=>res.json(req.user)));

// -------------------- PRODUCT ANALYTICS (self-hosted, consent-gated on the client) --------------------
// Ingest one or more usage events for the current user + school. Best-effort; never blocks the UI.
app.post('/api/analytics/event', h(async (req,res)=>{
  const b=req.body||{};
  const evs=Array.isArray(b.events)?b.events:[b];
  let n=0;
  for(const e of evs.slice(0,50)){
    const name=String((e&&e.name)||'').trim().slice(0,60);
    if(!name) continue;
    let props=null; try{ if(e&&e.props&&typeof e.props==='object') props=JSON.stringify(e.props).slice(0,2000); }catch(_){ props=null; }
    const pth=(e&&e.path)?String(e.path).slice(0,120):null;
    await run('INSERT INTO tt_event(school_id,user_id,name,props,path,ts) VALUES(?,?,?,?::jsonb,?,now())',[req.sid, req.user.id, name, props, pth]);
    n++;
  }
  res.json({ ok:true, stored:n });
}));
// Admin usage summary for the current school.
app.get('/api/analytics/summary', h(async (req,res)=>{
  if(!requireAdmin(req,res)) return;
  const sid=req.sid;
  const total=((await q1('SELECT count(*)::int n FROM tt_event WHERE school_id=?',[sid]))||{}).n||0;
  const last24=((await q1("SELECT count(*)::int n FROM tt_event WHERE school_id=? AND ts > now()-interval '24 hours'",[sid]))||{}).n||0;
  const users7=((await q1("SELECT count(DISTINCT user_id)::int n FROM tt_event WHERE school_id=? AND ts > now()-interval '7 days'",[sid]))||{}).n||0;
  const top=await q("SELECT name, count(*)::int n FROM tt_event WHERE school_id=? GROUP BY name ORDER BY n DESC LIMIT 15",[sid]);
  const daily=await q("SELECT to_char(date_trunc('day',ts),'YYYY-MM-DD') d, count(*)::int n FROM tt_event WHERE school_id=? AND ts > now()-interval '14 days' GROUP BY 1 ORDER BY 1",[sid]);
  const pages=await q("SELECT (props->>'tab') tab, count(*)::int n FROM tt_event WHERE school_id=? AND name='page_view' AND (props->>'tab') IS NOT NULL GROUP BY 1 ORDER BY n DESC LIMIT 12",[sid]);
  res.json({ total, last24, users7, top, daily, pages });
}));

// -------------------- SHAREABLE LINKS (public read-only class timetable) --------------------
app.get('/api/sharelinks', h(async (req,res)=>{ res.json(await q('SELECT * FROM tt_sharelink WHERE school_id=? ORDER BY id DESC',[req.sid])); }));
app.post('/api/sharelinks', h(async (req,res)=>{
  const target_id=req.body&&req.body.target_id?+req.body.target_id:null; if(!target_id){ res.status(400).json({error:'pick a class'}); return; }
  const ex=await q1('SELECT * FROM tt_sharelink WHERE school_id=? AND kind=? AND target_id=?',[req.sid,'class',target_id]);
  if(ex){ res.json({ok:true,token:ex.token,id:ex.id}); return; }
  const token=crypto.randomBytes(18).toString('hex');
  const r=await q1('INSERT INTO tt_sharelink(school_id,token,kind,target_id,created_at) VALUES(?,?,?,?,now()::text) RETURNING id',[req.sid,token,'class',target_id]);
  res.json({ok:true,token,id:r.id});
}));
app.delete('/api/sharelinks/:id', h(async (req,res)=>{ await run('DELETE FROM tt_sharelink WHERE id=? AND school_id=?',[req.params.id,req.sid]); res.json({ok:true}); }));

// -------------------- STUDENTS + STUDENT-LEVEL ELECTIVES --------------------
// Students of a class
app.get('/api/students', h(async (req,res)=>{
  const cid=req.query.class_id?+req.query.class_id:null;
  const rows=cid ? await q('SELECT * FROM tt_student WHERE school_id=? AND class_id=? ORDER BY roll,name,id',[req.sid,cid])
                 : await q('SELECT * FROM tt_student WHERE school_id=? ORDER BY class_id,roll,name,id',[req.sid]);
  res.json(rows);
}));
app.post('/api/students', h(async (req,res)=>{
  const b=req.body||{}; const cid=+b.class_id; const name=(b.name||'').trim();
  if(!cid||!name){ res.status(400).json({error:'class and name required'}); return; }
  const r=await q1('INSERT INTO tt_student(school_id,class_id,name,roll,created_at) VALUES(?,?,?,?,now()::text) RETURNING id',[req.sid,cid,name,(b.roll||'').trim()||null]);
  res.json({ok:true,id:r.id});
}));
app.delete('/api/students/:id', h(async (req,res)=>{
  const id=+req.params.id;
  await run('DELETE FROM tt_student_choice WHERE student_id=? AND school_id=?',[id,req.sid]);
  await run('DELETE FROM tt_student WHERE id=? AND school_id=?',[id,req.sid]);
  res.json({ok:true});
}));

// Electives for a class, each with its options nested
app.get('/api/electives', h(async (req,res)=>{
  const cid=req.query.class_id?+req.query.class_id:null;
  const els=cid ? await q('SELECT * FROM tt_elective WHERE school_id=? AND class_id=? ORDER BY day_of_week,period_index,id',[req.sid,cid])
                : await q('SELECT * FROM tt_elective WHERE school_id=? ORDER BY class_id,day_of_week,period_index,id',[req.sid]);
  const opts=await q('SELECT * FROM tt_elective_option WHERE school_id=? ORDER BY id',[req.sid]);
  els.forEach(e=>{ e.options=opts.filter(o=>o.elective_id===e.id); });
  res.json(els);
}));
app.post('/api/electives', h(async (req,res)=>{
  const b=req.body||{}; const cid=+b.class_id; const name=(b.name||'').trim();
  if(!cid||!name){ res.status(400).json({error:'class and name required'}); return; }
  const r=await q1('INSERT INTO tt_elective(school_id,class_id,name,day_of_week,period_index,week_index,created_at) VALUES(?,?,?,?,?,?,now()::text) RETURNING id',
    [req.sid,cid,name,+b.day_of_week||0,+b.period_index||0,+b.week_index||0]);
  res.json({ok:true,id:r.id});
}));
app.delete('/api/electives/:id', h(async (req,res)=>{
  const id=+req.params.id;
  await run('DELETE FROM tt_student_choice WHERE elective_id=? AND school_id=?',[id,req.sid]);
  await run('DELETE FROM tt_elective_option WHERE elective_id=? AND school_id=?',[id,req.sid]);
  await run('DELETE FROM tt_elective WHERE id=? AND school_id=?',[id,req.sid]);
  res.json({ok:true});
}));
app.post('/api/elective-options', h(async (req,res)=>{
  const b=req.body||{}; const eid=+b.elective_id; const label=(b.label||'').trim();
  if(!eid||!label){ res.status(400).json({error:'elective and label required'}); return; }
  const r=await q1('INSERT INTO tt_elective_option(school_id,elective_id,label,subject_id,teacher_id,room_id,created_at) VALUES(?,?,?,?,?,?,now()::text) RETURNING id',
    [req.sid,eid,label,b.subject_id?+b.subject_id:null,b.teacher_id?+b.teacher_id:null,b.room_id?+b.room_id:null]);
  res.json({ok:true,id:r.id});
}));
app.delete('/api/elective-options/:id', h(async (req,res)=>{
  const id=+req.params.id;
  await run('DELETE FROM tt_student_choice WHERE option_id=? AND school_id=?',[id,req.sid]);
  await run('DELETE FROM tt_elective_option WHERE id=? AND school_id=?',[id,req.sid]);
  res.json({ok:true});
}));

// Student choices: which option each student takes per elective
app.get('/api/student-choices', h(async (req,res)=>{
  const sidq=req.query.student_id?+req.query.student_id:null;
  const rows=sidq ? await q('SELECT elective_id,option_id FROM tt_student_choice WHERE school_id=? AND student_id=?',[req.sid,sidq])
                  : await q('SELECT student_id,elective_id,option_id FROM tt_student_choice WHERE school_id=?',[req.sid]);
  res.json(rows);
}));
app.post('/api/student-choices', h(async (req,res)=>{
  const b=req.body||{}; const stu=+b.student_id, eid=+b.elective_id;
  if(!stu||!eid){ res.status(400).json({error:'student and elective required'}); return; }
  if(b.option_id){
    await run(`INSERT INTO tt_student_choice(school_id,student_id,elective_id,option_id) VALUES(?,?,?,?)
       ON CONFLICT(student_id,elective_id) DO UPDATE SET option_id=excluded.option_id`,[req.sid,stu,eid,+b.option_id]);
  } else {
    await run('DELETE FROM tt_student_choice WHERE school_id=? AND student_id=? AND elective_id=?',[req.sid,stu,eid]);
  }
  res.json({ok:true});
}));

// -------------------- SCHOOLS (registry) --------------------
app.get('/api/schools', h(async (req,res)=>{
  if (req.user.role === 'master') res.json(await q('SELECT * FROM tt_school ORDER BY id'));
  else if (req.user.group_scope) res.json(await q('SELECT * FROM tt_school WHERE lower(trim(group_name))=lower(trim(?)) ORDER BY section_order NULLS LAST, id',[req.user.group_scope]));
  else res.json(await q('SELECT * FROM tt_school WHERE id=? ORDER BY id',[req.user.school_id]));
}));
const SCHOOL_LICENCE_LIMIT=4;   // max schools/sections per account on the current licence
app.post('/api/schools', h(async (req,res)=>{
  if(!requireAdmin(req,res)) return;
  const b=req.body||{}; const name=String(b.name||'').trim();
  if(!name){ res.status(400).json({error:'school name required'}); return; }
  // Licence limit — block adding more than the allowed number of schools/sections.
  const total=(await q1('SELECT COUNT(*)::int n FROM tt_school')).n;
  if(total>=SCHOOL_LICENCE_LIMIT){ res.status(403).json({error:'Licence limit reached — you can have up to '+SCHOOL_LICENCE_LIMIT+' schools/sections. Please purchase a new licence to add more.'}); return; }
  const row=await q1('INSERT INTO tt_school(name,board,medium,active,created_at) VALUES(?,?,?,1,now()::text) RETURNING id',
    [name, b.board||null, b.medium||null]);
  for(const f of ['address','school_code','mobile','email','udise','district','principal_name','pin_code','work_phone','group_name','section_name','section_order']){   // optional profile fields captured at creation
    if(b[f]!==undefined && String(b[f]).trim()!=='') await run('UPDATE tt_school SET '+f+'=? WHERE id=?',[String(b[f]).trim(), row.id]);
  }
  await seedConfigForSchool(row.id, name, b.board||null, b.medium||null);
  res.json({ id:row.id });
}));
app.put('/api/schools/:id', h(async (req,res)=>{
  if(!requireAdmin(req,res)) return;
  const b=req.body||{}, id=req.params.id;
  const cols=['name','board','medium','active','logo','price','paid','valid_till','modules_off','contact','notes','email','mobile','udise','district','deo_reg_code','deo_inspection_ref','deo_officer_name','deo_max_load','address','school_code','principal_name','pin_code','work_phone','group_name','section_name','section_order'];   // SaaS/profile + DEO header fields
  const set=cols.filter(c=>b[c]!==undefined);
  if(set.length) await run(`UPDATE tt_school SET ${set.map(c=>c+'=?').join(',')} WHERE id=?`,[...set.map(c=>b[c]===''?null:b[c]), id]);
  // keep this school's config row in sync for display (school_name/board/medium)
  const sync=[]; const sp=[];
  if(b.name!==undefined){ sync.push('school_name=?'); sp.push(b.name); }
  if(b.board!==undefined){ sync.push('board=?'); sp.push(b.board||null); }
  if(b.medium!==undefined){ sync.push('medium=?'); sp.push(b.medium||null); }
  if(sync.length){ await run(`UPDATE tt_config SET ${sync.join(',')} WHERE school_id=?`,[...sp, id]); await loadConfig(Number(id)); }
  res.json({ ok:true });
}));
app.delete('/api/schools/:id', h(async (req,res)=>{
  if(!requireAdmin(req,res)) return;
  const id=Number(req.params.id);
  const total=(await q1('SELECT COUNT(*)::int AS n FROM tt_school')).n;
  if(total<=1){ res.status(400).json({error:'cannot delete the only school'}); return; }
  // wipe this school's data
  await run('DELETE FROM tt_snapshot_cell WHERE snapshot_id IN (SELECT id FROM tt_snapshot WHERE school_id=?)',[id]);
  await run('DELETE FROM tt_teacher_subject WHERE teacher_id IN (SELECT id FROM tt_teacher WHERE school_id=?)',[id]);
  for(const tbl of ['tt_timetable','tt_quota','tt_chapter','tt_absence','tt_substitution','tt_diary','tt_snapshot','tt_student','tt_elective','tt_elective_option','tt_student_choice','tt_combine','tt_push_sub','tt_class','tt_subject','tt_room','tt_teacher','tt_config'])
    await run(`DELETE FROM ${tbl} WHERE school_id=?`,[id]);
  // detach users of that school (don't delete accounts)
  await run('UPDATE tt_user SET school_id=NULL WHERE school_id=?',[id]);
  await run('DELETE FROM tt_school WHERE id=?',[id]);
  await loadConfig();
  res.json({ ok:true });
}));

// Clone all core setup (subjects, rooms, teachers + subject links, classes, weekly quota, academic-hours config,
// combined periods, availability blocks) from one school into another EMPTY school. Timetable & students are NOT copied.
app.post('/api/schools/:id/clone-into/:target', h(async (req,res)=>{
  if(!requireAdmin(req,res)) return;
  const src=Number(req.params.id), tgt=Number(req.params.target);
  if(!src||!tgt||src===tgt){ res.status(400).json({error:'source and target must be two different schools'}); return; }
  const sOK=await q1('SELECT id FROM tt_school WHERE id=?',[src]);
  const tOK=await q1('SELECT id FROM tt_school WHERE id=?',[tgt]);
  if(!sOK||!tOK){ res.status(404).json({error:'school not found'}); return; }
  if(req.user.role!=='master' && src!==req.sid){ res.status(403).json({error:'forbidden'}); return; }
  const cnt=async(tbl)=>(await q1('SELECT COUNT(*)::int n FROM '+tbl+' WHERE school_id=?',[tgt])).n;
  if((await cnt('tt_class'))+(await cnt('tt_subject'))+(await cnt('tt_teacher'))>0){
    res.status(400).json({error:'target school is not empty — clone only into a fresh school'}); return;
  }
  const subMap={}, roomMap={}, tchMap={}, clsMap={};
  for(const s of await q('SELECT * FROM tt_subject WHERE school_id=? ORDER BY id',[src])){
    const r=await q1('INSERT INTO tt_subject(name,active,double_period,medium,color,school_id) VALUES(?,?,?,?,?,?) RETURNING id',
      [s.name,s.active,s.double_period,s.medium,s.color,tgt]); subMap[s.id]=r.id;
  }
  for(const rm of await q('SELECT * FROM tt_room WHERE school_id=? ORDER BY id',[src])){
    const r=await q1('INSERT INTO tt_room(name,capacity,school_id) VALUES(?,?,?) RETURNING id',[rm.name,rm.capacity,tgt]); roomMap[rm.id]=r.id;
  }
  for(const t of await q('SELECT * FROM tt_teacher WHERE school_id=? ORDER BY id',[src])){
    const r=await q1(`INSERT INTO tt_teacher(name,qualification,main_subject_id,max_load,max_per_day,max_consecutive,can_substitute,designation,sanctioned_load,email,mobile,school_id)
      VALUES(?,?,?,?,?,?,?,?,?,?,?,?) RETURNING id`,
      [t.name,t.qualification,subMap[t.main_subject_id]||null,t.max_load,t.max_per_day,t.max_consecutive,t.can_substitute,t.designation,t.sanctioned_load,t.email,t.mobile,tgt]);
    tchMap[t.id]=r.id;
    for(const ts of await q('SELECT * FROM tt_teacher_subject WHERE teacher_id=? ORDER BY seq NULLS LAST, subject_id',[t.id])){
      if(subMap[ts.subject_id]) await run('INSERT INTO tt_teacher_subject(teacher_id,subject_id,seq) VALUES(?,?,?) ON CONFLICT DO NOTHING',[r.id,subMap[ts.subject_id],ts.seq]);
    }
  }
  for(const c of await q('SELECT * FROM tt_class WHERE school_id=? ORDER BY id',[src])){
    const r=await q1('INSERT INTO tt_class(name,class_teacher_id,board,medium,standard,section,ct_first_period,school_id) VALUES(?,?,?,?,?,?,?,?) RETURNING id',
      [c.name, tchMap[c.class_teacher_id]||null, c.board,c.medium,c.standard,c.section,c.ct_first_period,tgt]); clsMap[c.id]=r.id;
  }
  for(const qr of await q('SELECT * FROM tt_quota WHERE school_id=?',[src])){
    if(clsMap[qr.class_id] && subMap[qr.subject_id])
      await run('INSERT INTO tt_quota(class_id,subject_id,per_week,teacher_id,school_id) VALUES(?,?,?,?,?)',
        [clsMap[qr.class_id],subMap[qr.subject_id],qr.per_week,tchMap[qr.teacher_id]||null,tgt]);
  }
  const sc=await q1('SELECT * FROM tt_config WHERE school_id=?',[src]);
  if(sc){
    const F=['weekday_start','weekday_end','saturday_start','saturday_end','period_minutes','period_durations','num_periods','lunch_after','lunch_minutes','short_break_after','short_break_minutes','working_days','sat_num_periods','sat_period_minutes','sat_period_durations','sat_lunch_after','sat_lunch_minutes','sat_short_break_after','sat_short_break_minutes','academic_session'];
    const present=F.filter(f=>f in sc);
    if(present.length) await run('UPDATE tt_config SET '+present.map(f=>f+'=?').join(',')+' WHERE school_id=?',[...present.map(f=>sc[f]), tgt]);
    try{ await loadConfig(tgt); }catch(e){}   // refresh the in-memory config cache so cloned hours take effect immediately
  }
  for(const cb of await q('SELECT * FROM tt_combine WHERE school_id=?',[src])){
    const ids=String(cb.class_ids||'').split(',').map(x=>x.trim()).filter(Boolean).map(x=>clsMap[Number(x)]).filter(Boolean);
    await run('INSERT INTO tt_combine(school_id,name,subject_id,teacher_id,room_id,day_of_week,period_index,week_index,class_ids,created_at) VALUES(?,?,?,?,?,?,?,?,?,now()::text)',
      [tgt,cb.name,subMap[cb.subject_id]||null,tchMap[cb.teacher_id]||null,roomMap[cb.room_id]||null,cb.day_of_week,cb.period_index,cb.week_index,ids.join(',')]);
  }
  for(const av of await q('SELECT * FROM tt_avail WHERE school_id=?',[src])){
    let eid=null;
    if(av.entity_type==='teacher') eid=tchMap[av.entity_id];
    else if(av.entity_type==='class') eid=clsMap[av.entity_id];
    else if(av.entity_type==='room') eid=roomMap[av.entity_id];
    if(eid) await run('INSERT INTO tt_avail(school_id,entity_type,entity_id,day_of_week,period_index,created_at) VALUES(?,?,?,?,?,now()::text) ON CONFLICT DO NOTHING',
      [tgt,av.entity_type,eid,av.day_of_week,av.period_index]);
  }
  res.json({ ok:true, subjects:Object.keys(subMap).length, rooms:Object.keys(roomMap).length, teachers:Object.keys(tchMap).length, classes:Object.keys(clsMap).length });
}));
// ---- USER MANAGEMENT (create/update/delete restricted to admin & master) ----
const ROLES=['master','admin','principal','supervisor','teacher'];
function requireAdmin(req,res){ if(!['admin','master'].includes(req.user.role)){ res.status(403).json({error:'forbidden'}); return false; } return true; }
app.get('/api/users', h(async (req,res)=>{
  if(!requireAdmin(req,res)) return;
  if(req.user.role==='master') res.json(await q('SELECT id,name,role,login_id,email,mobile,qualification,main_subject_id,active,created_at,school_id,group_scope FROM tt_user ORDER BY id'));
  else res.json(await q('SELECT id,name,role,login_id,email,mobile,qualification,main_subject_id,active,created_at,school_id,group_scope FROM tt_user WHERE school_id=? ORDER BY id',[req.sid]));
}));
app.post('/api/users', h(async (req,res)=>{
  if(!requireAdmin(req,res)) return;
  const b=req.body||{}; const login=String(b.login_id||'').trim();
  if(!login){ res.status(400).json({error:'login_id required'}); return; }
  if(!ROLES.includes(b.role||'teacher')){ res.status(400).json({error:'bad role'}); return; }
  if(await q1('SELECT 1 FROM tt_user WHERE lower(login_id)=lower(?)',[login])){ res.status(400).json({error:'login id already exists'}); return; }
  const pw=(b.password!=null && String(b.password).length)?b.password:'changeme123';
  // new users belong to the current school (master may target a specific school via body.school_id)
  const school = (req.user.role==='master' && b.school_id) ? b.school_id : req.sid;
  const gscope = (b.group_scope!=null && String(b.group_scope).trim()!=='') ? String(b.group_scope).trim() : null;
  const row=await q1(`INSERT INTO tt_user(name,role,login_id,email,mobile,qualification,main_subject_id,password_hash,active,created_at,school_id,group_scope)
     VALUES(?,?,?,?,?,?,?,?,1,now()::text,?,?) RETURNING id`,
    [b.name||login, b.role||'teacher', login, b.email||null, b.mobile||null, b.qualification||null, b.main_subject_id||null, hashPw(pw), school, gscope]);
  res.json({ id:row.id });
}));
app.put('/api/users/:id', h(async (req,res)=>{
  if(!requireAdmin(req,res)) return;
  const b=req.body||{}, id=req.params.id;
  if(req.user.role!=='master'){ const tgt=await q1('SELECT school_id FROM tt_user WHERE id=?',[id]); if(!tgt||tgt.school_id!==req.sid){ res.status(403).json({error:'forbidden'}); return; } }
  if(b.role!==undefined && !ROLES.includes(b.role)){ res.status(400).json({error:'bad role'}); return; }
  // Guard: an admin must not lose their own admin access, and the LAST admin of a school must not be demoted/deactivated (would lock everyone out).
  const demoting = (b.role!==undefined && !['admin','master'].includes(b.role));
  const deactivating = (b.active!==undefined && (Number(b.active)===0 || b.active===false));
  if(demoting && Number(id)===req.user.id){ res.status(400).json({error:'You cannot change your own role. Ask another Admin to do it.'}); return; }
  if(demoting || deactivating){
    const cur=await q1('SELECT role, school_id FROM tt_user WHERE id=?',[id]);
    if(cur && ['admin','master'].includes(cur.role)){
      const others = cur.school_id!=null
        ? (await q1('SELECT COUNT(*)::int AS n FROM tt_user WHERE role IN (?,?) AND id<>? AND active=1 AND school_id=?',['admin','master',id,cur.school_id])).n
        : (await q1('SELECT COUNT(*)::int AS n FROM tt_user WHERE role IN (?,?) AND id<>? AND active=1',['admin','master',id])).n;
      if(!others){ res.status(400).json({error:'This is the only Admin — make another user an Admin first, then you can change this one.'}); return; }
    }
  }
  const cols=['name','role','email','mobile','qualification','main_subject_id','active','group_scope'];
  const set=cols.filter(c=>b[c]!==undefined);
  if(set.length) await run(`UPDATE tt_user SET ${set.map(c=>c+'=?').join(',')} WHERE id=?`,[...set.map(c=>b[c]===''?null:b[c]), id]);
  if(b.login_id){ const login=String(b.login_id).trim(); if(login && !(await q1('SELECT 1 FROM tt_user WHERE lower(login_id)=lower(?) AND id<>?',[login,id]))) await run('UPDATE tt_user SET login_id=? WHERE id=?',[login,id]); }
  if(b.password!=null && String(b.password).length) await run('UPDATE tt_user SET password_hash=? WHERE id=?',[hashPw(b.password), id]);
  // Master may reassign or detach a user's school (school_id=null = not tied to any school, e.g. the platform owner).
  if(req.user.role==='master' && b.school_id!==undefined) await run('UPDATE tt_user SET school_id=? WHERE id=?',[(b.school_id===''||b.school_id==null)?null:Number(b.school_id), id]);
  res.json({ ok:true });
}));
app.delete('/api/users/:id', h(async (req,res)=>{
  if(!requireAdmin(req,res)) return;
  if(+req.params.id===req.user.id){ res.status(400).json({error:'cannot delete yourself'}); return; }
  if(req.user.role!=='master'){ const tgt=await q1('SELECT school_id FROM tt_user WHERE id=?',[req.params.id]); if(!tgt||tgt.school_id!==req.sid){ res.status(403).json({error:'forbidden'}); return; } }
  const target=await q1('SELECT role FROM tt_user WHERE id=?',[req.params.id]);
  if(target && ['admin','master'].includes(target.role)){
    const others=(await q1('SELECT COUNT(*)::int AS n FROM tt_user WHERE role IN (?,?) AND id<>? AND active=1',['admin','master',req.params.id])).n;
    if(!others){ res.status(400).json({error:'cannot remove the last admin'}); return; }
  }
  await run('DELETE FROM tt_session WHERE user_id=?',[req.params.id]);
  await run('DELETE FROM tt_user WHERE id=?',[req.params.id]);
  res.json({ ok:true });
}));

// ===================== PLATFORM OWNER (SaaS super-admin) — cross-school stats & logs =====================
function requireOwner(req,res){ if(!req.user||!req.user.is_owner){ res.status(403).json({error:'platform owner only'}); return false; } return true; }
// per-school overview: schools + counts (users/teachers/classes) + 7-day activity + last-seen
app.get('/api/owner/overview', h(async (req,res)=>{ if(!requireOwner(req,res)) return;
  const schools=await q('SELECT * FROM tt_school ORDER BY id');
  const mapN=(rows)=>{ const m={}; rows.forEach(r=>{ m[r.school_id]=r.n; }); return m; };
  const users=mapN(await q('SELECT school_id, COUNT(*)::int n FROM tt_user GROUP BY school_id'));
  const teachers=mapN(await q('SELECT school_id, COUNT(*)::int n FROM tt_teacher GROUP BY school_id'));
  const classes=mapN(await q('SELECT school_id, COUNT(*)::int n FROM tt_class GROUP BY school_id'));
  const ev7=mapN(await q("SELECT school_id, COUNT(*)::int n FROM tt_event WHERE ts > now() - interval '7 days' GROUP BY school_id"));
  const last={}; (await q('SELECT school_id, MAX(ts) t FROM tt_event GROUP BY school_id')).forEach(r=>{ last[r.school_id]=r.t; });
  res.json({ schools: schools.map(s=>({ ...s, users:users[s.id]||0, teachers:teachers[s.id]||0, classes:classes[s.id]||0, events7:ev7[s.id]||0, last_activity:last[s.id]||null })) });
}));
// cross-school activity log (recent events joined with user + school name)
app.get('/api/owner/logs', h(async (req,res)=>{ if(!requireOwner(req,res)) return;
  const lim=Math.min(500, Math.max(1, parseInt(req.query.limit,10)||150));
  const sc=req.query.school_id?Number(req.query.school_id):null;
  const rows= sc
    ? await q(`SELECT e.id,e.school_id,e.name,e.path,e.ts,s.name school_name,u.login_id,u.name user_name FROM tt_event e LEFT JOIN tt_school s ON s.id=e.school_id LEFT JOIN tt_user u ON u.id=e.user_id WHERE e.school_id=? ORDER BY e.id DESC LIMIT ?`,[sc,lim])
    : await q(`SELECT e.id,e.school_id,e.name,e.path,e.ts,s.name school_name,u.login_id,u.name user_name FROM tt_event e LEFT JOIN tt_school s ON s.id=e.school_id LEFT JOIN tt_user u ON u.id=e.user_id ORDER BY e.id DESC LIMIT ?`,[lim]);
  res.json(rows);
}));
// owner-only: update buy / payment settings
app.put('/api/pay-settings', h(async (req,res)=>{ if(!requireOwner(req,res)) return;
  const b=req.body||{};
  const s={
    gateway_url:String(b.gateway_url||'').trim(),
    upi_id:String(b.upi_id||'').trim(),
    whatsapp:String(b.whatsapp||'').trim().replace(/[^\d+]/g,''),
    contact_email:String(b.contact_email||'').trim(),
    payee_name:String(b.payee_name||'').trim()||'Aumtara Timetable',
    enable_gateway:b.enable_gateway?1:0,
    enable_upi:b.enable_upi?1:0,
    enable_enquiry:b.enable_enquiry?1:0,
    note:String(b.note||'').trim()
  };
  await setMeta('pay_settings', JSON.stringify(s));
  res.json({ ok:true, settings:s });
}));
// owner-only: list buy enquiries + mark handled
app.get('/api/enquiries', h(async (req,res)=>{ if(!requireOwner(req,res)) return;
  res.json(await q('SELECT id,name,school,phone,email,plan,message,handled,ts FROM tt_enquiry ORDER BY id DESC LIMIT 200'));
}));
app.put('/api/enquiry/:id', h(async (req,res)=>{ if(!requireOwner(req,res)) return;
  await run('UPDATE tt_enquiry SET handled=? WHERE id=?',[req.body&&req.body.handled?1:0, Number(req.params.id)]);
  res.json({ ok:true });
}));
// owner-only: record a payment (sales ledger) + optionally mark the school paid
app.post('/api/owner/payment', h(async (req,res)=>{ if(!requireOwner(req,res)) return;
  const b=req.body||{};
  const sid=Number(b.school_id)||null;
  const amount=parseFloat(b.amount); if(!sid || !(amount>0)){ res.status(400).json({error:'school and a positive amount are required'}); return; }
  const method=String(b.method||'other').trim().toLowerCase();
  await run('INSERT INTO tt_payment(school_id,amount,method,note,ts) VALUES(?,?,?,?,now()::text)',[sid, amount, method, String(b.note||'').trim()||null]);
  if(b.mark_paid) await run('UPDATE tt_school SET paid=1 WHERE id=?',[sid]);
  res.json({ ok:true });
}));
// owner-only: sales ledger + totals (all / this-month / by method)
app.get('/api/owner/payments', h(async (req,res)=>{ if(!requireOwner(req,res)) return;
  const rows=await q('SELECT p.id,p.school_id,p.amount::float amount,p.method,p.note,p.ts,s.name school_name FROM tt_payment p LEFT JOIN tt_school s ON s.id=p.school_id ORDER BY p.id DESC LIMIT 200');
  const byMethod=await q("SELECT COALESCE(NULLIF(method,''),'other') method, SUM(amount)::float total, COUNT(*)::int n FROM tt_payment GROUP BY 1 ORDER BY 2 DESC");
  const tot=await q1('SELECT COALESCE(SUM(amount),0)::float total, COUNT(*)::int n FROM tt_payment');
  const month=await q1("SELECT COALESCE(SUM(amount),0)::float total FROM tt_payment WHERE substr(ts,1,7)=to_char(now(),'YYYY-MM')");
  res.json({ payments:rows, byMethod, total:tot.total, count:tot.n, month:month.total });
}));
// owner-only: module packages (named sets of enabled modules) stored in tt_appmeta 'packages'
app.get('/api/owner/packages', h(async (req,res)=>{ if(!requireOwner(req,res)) return;
  let p=[]; try{ const raw=await getMeta('packages'); p=raw?JSON.parse(raw):[]; }catch(_){ p=[]; }
  res.json(Array.isArray(p)?p:[]);
}));
app.put('/api/owner/packages', h(async (req,res)=>{ if(!requireOwner(req,res)) return;
  const arr=Array.isArray(req.body)?req.body:((req.body&&req.body.packages)||[]);
  const clean=arr.filter(x=>x&&x.name).map(x=>({ name:String(x.name).trim().slice(0,40), price:(x.price==null||x.price==='')?null:String(x.price), modules:Array.isArray(x.modules)?x.modules.map(String):[] }));
  await setMeta('packages', JSON.stringify(clean));
  res.json({ ok:true, packages:clean });
}));
// owner-only: reset a school's admin login password (for support)
app.post('/api/owner/school/:id/reset-password', h(async (req,res)=>{ if(!requireOwner(req,res)) return;
  const sid=Number(req.params.id); const pw=String((req.body||{}).password||'');
  if(pw.length<6){ res.status(400).json({error:'password must be at least 6 characters'}); return; }
  const u=await q1("SELECT id,login_id FROM tt_user WHERE school_id=? AND role IN ('admin','master') ORDER BY id LIMIT 1",[sid]);
  if(!u){ res.status(404).json({error:'no admin login found for this school'}); return; }
  await run('UPDATE tt_user SET password_hash=? WHERE id=?',[hashPw(pw), u.id]);
  res.json({ ok:true, login_id:u.login_id });
}));
// owner-only: directly create an admin/principal login for a school (e.g. after verifying a payment)
app.post('/api/owner/school/:id/create-login', h(async (req,res)=>{ if(!requireOwner(req,res)) return;
  const sid=Number(req.params.id); const b=req.body||{};
  const login=String(b.login_id||'').trim();
  const pw=String(b.password||'');
  const name=String(b.name||'').trim()||login;
  const email=String(b.email||'').trim()||null;
  const mobile=String(b.mobile||'').trim()||null;
  const role=(b.role==='principal')?'principal':'admin';
  if(!login){ res.status(400).json({error:'login ID is required'}); return; }
  if(pw.length<6){ res.status(400).json({error:'password must be at least 6 characters'}); return; }
  const sch=await q1('SELECT id,name FROM tt_school WHERE id=?',[sid]);
  if(!sch){ res.status(404).json({error:'school not found'}); return; }
  if(await q1('SELECT 1 FROM tt_user WHERE lower(login_id)=lower(?)',[login])){ res.status(400).json({error:'this login ID is already taken'}); return; }
  if(email && await q1('SELECT 1 FROM tt_user WHERE lower(email)=lower(?)',[email])){ res.status(400).json({error:'this email is already used by another login'}); return; }
  const row=await q1(`INSERT INTO tt_user(name,role,login_id,email,mobile,password_hash,active,created_at,school_id)
     VALUES(?,?,?,?,?,?,1,now()::text,?) RETURNING id`, [name, role, login, email, mobile, hashPw(pw), sid]);
  res.json({ ok:true, id:row.id, login_id:login, role:role, school:sch.name });
}));
// owner-only: list all logins for a school (manage from the owner panel)
app.get('/api/owner/school/:id/logins', h(async (req,res)=>{ if(!requireOwner(req,res)) return;
  const sid=Number(req.params.id);
  const rows=await q('SELECT id,name,login_id,role,email,mobile,active,is_owner FROM tt_user WHERE school_id=? ORDER BY id',[sid]);
  res.json({ ok:true, logins:rows });
}));
// owner-only: delete a login (cannot delete the platform-owner/master account)
app.delete('/api/owner/user/:uid', h(async (req,res)=>{ if(!requireOwner(req,res)) return;
  const uid=Number(req.params.uid);
  const u=await q1('SELECT id,role,is_owner FROM tt_user WHERE id=?',[uid]);
  if(!u){ res.status(404).json({error:'login not found'}); return; }
  if(Number(u.is_owner)===1 || u.role==='master'){ res.status(400).json({error:'cannot delete the platform-owner / master login'}); return; }
  await run('DELETE FROM tt_user WHERE id=?',[uid]);
  res.json({ ok:true });
}));
// owner-only: change a login's role (e.g. demote an admin to principal)
app.put('/api/owner/user/:uid/role', h(async (req,res)=>{ if(!requireOwner(req,res)) return;
  const uid=Number(req.params.uid); const role=(String((req.body||{}).role||'')==='admin')?'admin':'principal';
  const u=await q1('SELECT id,is_owner,role FROM tt_user WHERE id=?',[uid]);
  if(!u){ res.status(404).json({error:'login not found'}); return; }
  if(Number(u.is_owner)===1 || u.role==='master'){ res.status(400).json({error:'cannot change the owner / master login'}); return; }
  await run('UPDATE tt_user SET role=? WHERE id=?',[role, uid]);
  res.json({ ok:true, role });
}));
// any signed-in user: update own profile (email + mobile + optionally login ID)
app.put('/api/me', h(async (req,res)=>{
  const b=req.body||{};
  const email=String(b.email||'').trim()||null, mobile=String(b.mobile||'').trim()||null;
  if(email && await q1('SELECT 1 FROM tt_user WHERE lower(email)=lower(?) AND id<>?',[email, req.user.id])){ res.status(400).json({error:'email already used by another account'}); return; }
  if(mobile && await q1('SELECT 1 FROM tt_user WHERE mobile=? AND id<>?',[mobile, req.user.id])){ res.status(400).json({error:'mobile already used by another account'}); return; }
  let newLogin=null;
  if(b.login_id!==undefined){
    newLogin=String(b.login_id||'').trim();
    if(!newLogin){ res.status(400).json({error:'login ID cannot be empty'}); return; }
    if(await q1('SELECT 1 FROM tt_user WHERE lower(login_id)=lower(?) AND id<>?',[newLogin, req.user.id])){ res.status(400).json({error:'that login ID is already taken'}); return; }
    await run('UPDATE tt_user SET login_id=? WHERE id=?',[newLogin, req.user.id]);
  }
  await run('UPDATE tt_user SET email=?, mobile=? WHERE id=?',[email, mobile, req.user.id]);
  res.json({ ok:true, email, mobile, login_id:newLogin!==null?newLogin:req.user.login_id });
}));
// ---------- 2FA (authenticator app / TOTP) — for the signed-in account ----------
app.get('/api/2fa/status', h(async (req,res)=>{
  const u=await q1('SELECT totp_enabled FROM tt_user WHERE id=?',[req.user.id]);
  res.json({ ok:true, enabled: !!(u && Number(u.totp_enabled)===1) });
}));
// step 1: create a fresh secret (not yet enabled) and return the otpauth URI for the QR
app.post('/api/2fa/setup', h(async (req,res)=>{
  const secret=genTotpSecret();
  await run('UPDATE tt_user SET totp_secret=?, totp_enabled=0 WHERE id=?',[secret, req.user.id]);
  const label=encodeURIComponent('Aumtara ('+(req.user.login_id||'owner')+')');
  const otpauth='otpauth://totp/'+label+'?secret='+secret+'&issuer=Aumtara&period=30&digits=6&algorithm=SHA1';
  res.json({ ok:true, secret, otpauth });
}));
// step 2: verify a code from the app, then turn 2FA on and hand back one-time backup codes
app.post('/api/2fa/enable', h(async (req,res)=>{
  const u=await q1('SELECT totp_secret FROM tt_user WHERE id=?',[req.user.id]);
  if(!u || !u.totp_secret){ res.status(400).json({error:'Please start the setup first.'}); return; }
  if(!totpVerify(u.totp_secret, (req.body||{}).code)){ res.status(400).json({error:'That code is wrong or expired — check your authenticator app and try again.'}); return; }
  const codes=genBackupCodes(8);
  const hashed=JSON.stringify(codes.map(c=>hashPw(c)));
  await run('UPDATE tt_user SET totp_enabled=1, backup_codes=? WHERE id=?',[hashed, req.user.id]);
  res.json({ ok:true, backup_codes:codes });
}));
// turn 2FA off (needs a current code or a backup code)
app.post('/api/2fa/disable', h(async (req,res)=>{
  const u=await q1('SELECT totp_secret,totp_enabled,backup_codes FROM tt_user WHERE id=?',[req.user.id]);
  if(!u || Number(u.totp_enabled)!==1){ res.json({ ok:true }); return; }
  const code=String((req.body||{}).code||'').replace(/\s/g,'');
  let ok=totpVerify(u.totp_secret, code);
  if(!ok){ let bc=[]; try{ bc=JSON.parse(u.backup_codes||'[]'); }catch(_){ bc=[]; } ok=bc.some(hc=>verifyPw(code.toLowerCase(),hc)); }
  if(!ok){ res.status(400).json({error:'Enter a valid current code to turn 2FA off.'}); return; }
  await run('UPDATE tt_user SET totp_enabled=0, totp_secret=NULL, backup_codes=NULL WHERE id=?',[req.user.id]);
  res.json({ ok:true });
}));
// any signed-in user: change own password (needs current password)
app.post('/api/me/password', h(async (req,res)=>{
  const b=req.body||{};
  const cur=String(b.current||''), nw=String(b.new||'');
  if(nw.length<6){ res.status(400).json({error:'new password must be at least 6 characters'}); return; }
  const u=await q1('SELECT password_hash FROM tt_user WHERE id=?',[req.user.id]);
  if(!u || !verifyPw(cur, u.password_hash)){ res.status(401).json({error:'current password is incorrect'}); return; }
  await run('UPDATE tt_user SET password_hash=? WHERE id=?',[hashPw(nw), req.user.id]);
  res.json({ ok:true });
}));

const DAYS = ['Mon','Tue','Wed','Thu','Fri','Sat'];
const addMin = (hhmm,min)=>{const[h,m]=hhmm.split(':').map(Number);const d=new Date(2020,0,1,h,m+min);return String(d.getHours()).padStart(2,'0')+':'+String(d.getMinutes()).padStart(2,'0');};
const getConfig = (sid) => getConfigCached(sid);
const csvNums = s => String(s||'').split(',').map(x=>parseInt(x,10)).filter(n=>!isNaN(n));
// Day-Rotation: N rotation days (A..), each its own daily plan; day_of_week 0..N-1 = rotation day index.
function rotDays(sid){ const c=getConfig(sid)||{}; const n=parseInt(c.rot_days,10)||0; return (n>=2&&n<=8)?n:0; }
function rotMode(sid){ return rotDays(sid)>0; }
function rotLabel(sid,i){ const c=getConfig(sid)||{}; const L=(c.rot_labels||'').split(',').map(x=>x.trim()).filter(Boolean); return L[i] || ('Day '+String.fromCharCode(65+i)); }
// which rotation day is active today, from the anchor date (counts working/calendar days since cycle_start mod N)
function todayRotDay(sid){ const c=getConfig(sid)||{}; const n=rotDays(sid); if(!n||!c.cycle_start) return 0;
  const start=new Date(c.cycle_start+'T00:00:00'); if(isNaN(start.getTime())) return 0;
  const nowD=new Date(); const today=new Date(nowD.getFullYear(),nowD.getMonth(),nowD.getDate());
  let count=0; const d=new Date(start); for(let i=0;i<800 && d<=today;i++){ const wd=(d.getDay()+6)%7; if(wd<6) count++; d.setDate(d.getDate()+1); }  // skip Sundays only
  return ((count-1)%n+n)%n; }
function workingDaysArr(sid){ const rd=rotDays(sid); if(rd) return Array.from({length:rd},(_,i)=>i); const c=getConfig(sid)||{}; const w=csvNums(c.working_days); return w.length?w:[0,1,2,3,4,5]; }
// which cycle-week is active today, from the rotation anchor date (0 when rotation is off)
function curCycleIdx(sid){
  const c=getConfig(sid)||{};
  const n=Math.max(1,Math.min(4,parseInt(c.cycle_weeks,10)||1));
  if(n<=1 || !c.cycle_start) return 0;
  const start=new Date(c.cycle_start+'T00:00:00'); if(isNaN(start.getTime())) return 0;
  const now=new Date(); const today=new Date(now.getFullYear(),now.getMonth(),now.getDate());
  if((c.cycle_mode||'week')==='day'){
    const wd=new Set(workingDaysArr(sid)); let count=0; const d=new Date(start);
    // guard the loop (max ~2 years of days)
    for(let i=0;i<800 && d<=today;i++){ if(wd.has((d.getDay()+6)%7)) count++; d.setDate(d.getDate()+1); }
    return ((count-1)%n+n)%n;   // count includes today; 0-based index
  }
  // week mode: whole weeks between the Monday of start and the Monday of today
  const mondayOf=(x)=>{ const m=new Date(x); m.setDate(m.getDate()-((m.getDay()+6)%7)); m.setHours(0,0,0,0); return m; };
  const weeks=Math.floor((mondayOf(today)-mondayOf(start))/(7*86400000));
  return ((weeks%n)+n)%n;
}
app.get('/api/rotation/today', h(async (req,res)=>{
  const idx=curCycleIdx(req.sid); const c=getConfig(req.sid)||{};
  const n=Math.max(1,Math.min(4,parseInt(c.cycle_weeks,10)||1));
  const labels=(c.week_labels||'').split(',').map(x=>x.trim()).filter(Boolean);
  const label=labels[idx] || (n===2?['Week A','Week B'][idx] : 'Week '+(idx+1));
  res.json({ index:idx, cycle_weeks:n, active: n>1 && !!c.cycle_start, mode:c.cycle_mode||'week', label });
}));
function slotsForDay(dayIdx, sid){
  const c=getConfig(sid);
  if(!c) return [];
  if(!workingDaysArr(sid).includes(dayIdx)) return [];   // non-working day → no periods
  const sat = rotMode(sid) ? false : (dayIdx===5);   // rotation days all use the uniform weekday layout
  const end=sat?c.saturday_end:c.weekday_end;
  const start=(sat&&c.saturday_start)?c.saturday_start:c.weekday_start;
  const periodMin=+(sat&&c.sat_period_minutes!=null?c.sat_period_minutes:c.period_minutes);
  const durs=csvNums(sat&&c.sat_period_durations!=null?c.sat_period_durations:c.period_durations);
  const shortAfter=new Set(csvNums(sat&&c.sat_short_break_after!=null?c.sat_short_break_after:c.short_break_after));
  const shortMin=+(sat&&c.sat_short_break_minutes!=null?c.sat_short_break_minutes:c.short_break_minutes)||0;
  const lunchAfter=+(sat&&c.sat_lunch_after!=null?c.sat_lunch_after:(c.lunch_after!=null?c.lunch_after:c.break_after_period));
  const lunchMin=+(sat&&c.sat_lunch_minutes!=null?c.sat_lunch_minutes:(c.lunch_minutes!=null?c.lunch_minutes:c.break_minutes))||0;
  const N=+(sat&&c.sat_num_periods!=null?c.sat_num_periods:c.num_periods)||0;   // explicit period count; 0 = derive from end time
  const CAP=20;
  const slots=[]; let t=start, idx=0;
  while(true){ const dur=durs[idx]||periodMin; const e=addMin(t,dur);
    slots.push({index:idx,label:'P'+(idx+1),start:t,end:e,is_break:false}); idx++;
    const stop = N>0 ? idx>=Math.min(N,CAP) : e>=end;
    if(stop||idx>=CAP)break; t=e;
    if(lunchAfter&&idx===lunchAfter&&lunchMin>0){ const be=addMin(t,lunchMin); slots.push({index:null,label:'Lunch',start:t,end:be,is_break:true,kind:'lunch'}); t=be; }
    if(shortAfter.has(idx)&&shortMin>0){ const be=addMin(t,shortMin); slots.push({index:null,label:'Break',start:t,end:be,is_break:true,kind:'short'}); t=be; } }
  return slots;
}
const teachingSlots = (di, sid)=>slotsForDay(di, sid).filter(s=>!s.is_break);

// ---------- CLASS / SUBJECT / ROOM (simple CRUD) ----------
function simpleCrud(route, tbl, cols){
  app.get('/api/'+route, h(async (req,res)=>res.json(await q(`SELECT * FROM ${tbl} WHERE school_id=? ORDER BY id`,[req.sid]))));
  app.post('/api/'+route, h(async (req,res)=>{
    const vals=cols.map(c=>req.body[c]??null);
    const row=await q1(`INSERT INTO ${tbl}(${cols.join(',')},school_id) VALUES (${cols.map(()=>'?').join(',')},?) RETURNING id`, [...vals, req.sid]);
    res.json({id:row.id});
  }));
  app.put('/api/'+route+'/:id', h(async (req,res)=>{
    const c=cols.filter(x=>req.body[x]!==undefined);
    if(c.length) await run(`UPDATE ${tbl} SET ${c.map(x=>x+'=?').join(',')} WHERE id=? AND school_id=?`, [...c.map(x=>req.body[x]), req.params.id, req.sid]);
    res.json({ok:true});
  }));
  app.delete('/api/'+route+'/:id', h(async (req,res)=>{ await run(`DELETE FROM ${tbl} WHERE id=? AND school_id=?`, [req.params.id, req.sid]); res.json({ok:true}); }));
}
simpleCrud('classes','tt_class',['name','class_teacher_id','board','medium','standard','section','ct_first_period']);
simpleCrud('subjects','tt_subject',['name','active','double_period','medium','color','is_core']);
simpleCrud('rooms','tt_room',['name','capacity']);

// ---------- SETUP READINESS (checklist %) ----------
app.get('/api/readiness', h(async (req,res)=>{
  const c=getConfig(req.sid)||{};
  const cnt=async (sql)=> (await q1(sql,[req.sid])).n;
  const classes  = await cnt('SELECT COUNT(*)::int n FROM tt_class WHERE school_id=?');
  const subjects = await cnt('SELECT COUNT(*)::int n FROM tt_subject WHERE active=1 AND school_id=?');
  const teachers = await cnt('SELECT COUNT(*)::int n FROM tt_teacher WHERE school_id=?');
  const rooms    = await cnt('SELECT COUNT(*)::int n FROM tt_room WHERE school_id=?');
  const quota    = await cnt('SELECT COUNT(*)::int n FROM tt_quota WHERE per_week>0 AND school_id=?');
  const cells    = await cnt('SELECT COUNT(*)::int n FROM tt_timetable WHERE subject_id IS NOT NULL AND school_id=?');
  const named    = !!(c.school_name && c.school_name!=='Your School Name');
  const hours    = !!(c.weekday_start && c.weekday_end && (c.period_minutes || c.num_periods || (c.period_durations && String(c.period_durations).trim())));
  const steps = [
    { key:'classes',   done: classes>0,  count: classes },
    { key:'subjects',  done: subjects>0, count: subjects },
    { key:'teachers',  done: teachers>0, count: teachers },
    { key:'rooms',     done: rooms>0,    count: rooms },
    { key:'hours',     done: hours },
    { key:'name',      done: named },
    { key:'quota',     done: quota>0,    count: quota, optional:true },
    { key:'generated', done: cells>0,    count: cells },
  ];
  const reqd = steps.filter(s=>!s.optional);
  const pct = Math.round(100 * reqd.filter(s=>s.done).length / reqd.length);
  res.json({ steps, pct });
}));

// ---------- CHAPTERS (per subject) ----------
app.get('/api/chapters', h(async (req,res)=>{
  if(req.query.subject_id) res.json(await q('SELECT * FROM tt_chapter WHERE subject_id=? AND school_id=? ORDER BY seq,id',[req.query.subject_id, req.sid]));
  else res.json(await q('SELECT * FROM tt_chapter WHERE school_id=? ORDER BY subject_id,seq,id',[req.sid]));
}));
app.post('/api/chapters', h(async (req,res)=>{
  const b=req.body; const row=await q1('INSERT INTO tt_chapter(subject_id,name,seq,school_id) VALUES(?,?,?,?) RETURNING id',[b.subject_id, b.name, b.seq||null, req.sid]);
  res.json({id:row.id});
}));
app.delete('/api/chapters/:id', h(async (req,res)=>{ await run('DELETE FROM tt_chapter WHERE id=? AND school_id=?',[req.params.id, req.sid]); res.json({ok:true}); }));

// ---------- TEACHERS (with subject mapping) ----------
app.get('/api/teachers', h(async (req,res)=>{
  const t=await q('SELECT * FROM tt_teacher WHERE school_id=? ORDER BY id',[req.sid]);
  const map=await q('SELECT ts.* FROM tt_teacher_subject ts JOIN tt_teacher t ON t.id=ts.teacher_id WHERE t.school_id=? ORDER BY ts.teacher_id, ts.seq NULLS LAST, ts.subject_id',[req.sid]);
  t.forEach(x=>x.subjects=map.filter(m=>m.teacher_id===x.id).map(m=>m.subject_id));
  res.json(t);
}));
app.post('/api/teachers', h(async (req,res)=>{
  const b=req.body;
  const row=await q1('INSERT INTO tt_teacher(name,qualification,main_subject_id,max_load,school_id) VALUES (?,?,?,?,?) RETURNING id',
    [b.name||'New Teacher', b.qualification||null, b.main_subject_id||null, b.max_load||null, req.sid]);
  const subs=new Set(b.subjects||[]); if(b.main_subject_id)subs.add(+b.main_subject_id);
  await setSubjects(row.id, [...subs]);
  res.json({id:row.id});
}));
app.put('/api/teachers/:id', h(async (req,res)=>{
  const b=req.body; const sid=req.sid;
  if(b.name!==undefined) await run('UPDATE tt_teacher SET name=? WHERE id=? AND school_id=?',[b.name, req.params.id, sid]);
  if(b.qualification!==undefined) await run('UPDATE tt_teacher SET qualification=? WHERE id=? AND school_id=?',[b.qualification, req.params.id, sid]);
  if(b.main_subject_id!==undefined) await run('UPDATE tt_teacher SET main_subject_id=? WHERE id=? AND school_id=?',[b.main_subject_id||null, req.params.id, sid]);
  if(b.max_load!==undefined) await run('UPDATE tt_teacher SET max_load=? WHERE id=? AND school_id=?',[b.max_load||null, req.params.id, sid]);
  if(b.max_per_day!==undefined) await run('UPDATE tt_teacher SET max_per_day=? WHERE id=? AND school_id=?',[b.max_per_day||null, req.params.id, sid]);
  if(b.max_consecutive!==undefined) await run('UPDATE tt_teacher SET max_consecutive=? WHERE id=? AND school_id=?',[b.max_consecutive||null, req.params.id, sid]);
  if(b.can_substitute!==undefined) await run('UPDATE tt_teacher SET can_substitute=? WHERE id=? AND school_id=?',[b.can_substitute?1:0, req.params.id, sid]);
  if(b.is_senior!==undefined) await run('UPDATE tt_teacher SET is_senior=? WHERE id=? AND school_id=?',[b.is_senior?1:0, req.params.id, sid]);
  if(b.designation!==undefined) await run('UPDATE tt_teacher SET designation=? WHERE id=? AND school_id=?',[b.designation||null, req.params.id, sid]);
  if(b.sanctioned_load!==undefined) await run('UPDATE tt_teacher SET sanctioned_load=? WHERE id=? AND school_id=?',[b.sanctioned_load||null, req.params.id, sid]);
  if(b.subjects!==undefined){ const subs=[]; const seen=new Set(); (b.subjects||[]).forEach(x=>{ if(!seen.has(x)){seen.add(x);subs.push(x);} }); if(b.main_subject_id&&!seen.has(+b.main_subject_id))subs.push(+b.main_subject_id); await setSubjects(+req.params.id, subs); }
  res.json({ok:true});
}));
app.delete('/api/teachers/:id', h(async (req,res)=>{
  await run('DELETE FROM tt_teacher WHERE id=? AND school_id=?',[req.params.id, req.sid]);
  await run('DELETE FROM tt_teacher_subject WHERE teacher_id=?',[req.params.id]);
  res.json({ok:true});
}));
async function setSubjects(tid, subs){
  await run('DELETE FROM tt_teacher_subject WHERE teacher_id=?',[tid]);
  let i=0; const seen=new Set();
  for(const s of subs){ if(seen.has(s))continue; seen.add(s);   // keep array order = priority (0 = 1st)
    await run('INSERT INTO tt_teacher_subject(teacher_id,subject_id,seq) VALUES (?,?,?) ON CONFLICT DO NOTHING',[tid, s, i++]); }
}

// ---------- CONFIG / SLOTS ----------
app.get('/api/timetable/config', h(async (req,res)=>res.json(getConfig(req.sid))));
app.put('/api/timetable/config', h(async (req,res)=>{
  const c=getConfig(req.sid)||{}, b=req.body, v=(k)=>b[k]!==undefined?b[k]:c[k];
  await run(`UPDATE tt_config SET weekday_start=?,weekday_end=?,saturday_start=?,saturday_end=?,period_minutes=?,break_after_period=?,break_minutes=?,school_name=?,
       short_break_minutes=?,short_break_after=?,lunch_minutes=?,lunch_after=?,period_durations=?,working_days=?,academic_session=?,
       sat_period_minutes=?,sat_period_durations=?,sat_lunch_after=?,sat_lunch_minutes=?,sat_short_break_after=?,sat_short_break_minutes=?,
       num_periods=?,sat_num_periods=? WHERE school_id=?`,
    [v('weekday_start'),v('weekday_end'),v('saturday_start'),v('saturday_end'),v('period_minutes'),v('break_after_period'),v('break_minutes'),v('school_name'),
     v('short_break_minutes'),v('short_break_after'),v('lunch_minutes'),v('lunch_after'),v('period_durations'),v('working_days'),v('academic_session'),
     v('sat_period_minutes'),v('sat_period_durations'),v('sat_lunch_after'),v('sat_lunch_minutes'),v('sat_short_break_after'),v('sat_short_break_minutes'),
     v('num_periods'),v('sat_num_periods'), req.sid]);
  // keep the school registry name in sync if the school name was edited here
  if(b.school_name!==undefined) await run('UPDATE tt_school SET name=? WHERE id=?',[b.school_name, req.sid]);
  // custom entity labels (Labels & Terminology)
  for(const k of ['label_class','label_teacher','label_room','label_subject'])
    if(b[k]!==undefined) await run(`UPDATE tt_config SET ${k}=? WHERE school_id=?`,[(b[k]||'').trim()||null, req.sid]);
  // Auto Set optimizer toggles
  for(const k of ['opt_spread','opt_balance','opt_morning','opt_gap','opt_solver'])
    if(b[k]!==undefined) await run(`UPDATE tt_config SET ${k}=? WHERE school_id=?`,[b[k]?1:0, req.sid]);
  // Multi-week cycle
  if(b.cycle_weeks!==undefined){ const n=Math.max(1,Math.min(4,parseInt(b.cycle_weeks,10)||1)); await run(`UPDATE tt_config SET cycle_weeks=? WHERE school_id=?`,[n, req.sid]); }
  if(b.week_labels!==undefined) await run(`UPDATE tt_config SET week_labels=? WHERE school_id=?`,[(b.week_labels||'').trim()||null, req.sid]);
  if(b.cycle_start!==undefined) await run(`UPDATE tt_config SET cycle_start=? WHERE school_id=?`,[(b.cycle_start||'').trim()||null, req.sid]);
  if(b.cycle_mode!==undefined) await run(`UPDATE tt_config SET cycle_mode=? WHERE school_id=?`,[b.cycle_mode==='day'?'day':'week', req.sid]);
  if(b.rot_days!==undefined){ const n=parseInt(b.rot_days,10)||0; await run(`UPDATE tt_config SET rot_days=? WHERE school_id=?`,[(n>=2&&n<=8)?n:0, req.sid]); }
  if(b.rot_labels!==undefined) await run(`UPDATE tt_config SET rot_labels=? WHERE school_id=?`,[(b.rot_labels||'').trim()||null, req.sid]);
  await loadConfig(req.sid);
  res.json(getConfig(req.sid));
}));
app.get('/api/timetable/slots', h(async (req,res)=>res.json(slotsForDay(Number(req.query.day||0), req.sid))));

// ---------- ACADEMIC TERMS (structured sessions; one active at a time) ----------
async function syncActiveTerm(sid){
  const t=await q1('SELECT name FROM tt_term WHERE school_id=? AND active=1 ORDER BY id LIMIT 1',[sid]);
  if(t) await run('UPDATE tt_config SET academic_session=? WHERE school_id=?',[t.name, sid]);
}
app.get('/api/terms', h(async (req,res)=> res.json(await q('SELECT * FROM tt_term WHERE school_id=? ORDER BY active DESC, start_date NULLS LAST, id',[req.sid]))));
app.post('/api/terms', h(async (req,res)=>{
  const b=req.body; const active=b.active?1:0;
  const row=await q1('INSERT INTO tt_term(school_id,name,type,start_date,end_date,active,created_at) VALUES(?,?,?,?,?,?,now()::text) RETURNING id',
    [req.sid, b.name||'Untitled', b.type||'regular', b.start_date||null, b.end_date||null, active]);
  if(active){ await run('UPDATE tt_term SET active=0 WHERE school_id=? AND id<>?',[req.sid, row.id]); await syncActiveTerm(req.sid); await loadConfig(req.sid); }
  res.json({id:row.id});
}));
app.put('/api/terms/:id', h(async (req,res)=>{
  const id=req.params.id, b=req.body;
  const own=await q1('SELECT id FROM tt_term WHERE id=? AND school_id=?',[id, req.sid]);
  if(!own){ res.status(404).json({error:'not found'}); return; }
  const cols=['name','type','start_date','end_date','active'].filter(k=>b[k]!==undefined);
  if(cols.length) await run(`UPDATE tt_term SET ${cols.map(c=>c+'=?').join(',')} WHERE id=? AND school_id=?`,
    [...cols.map(c=>c==='active'?(b[c]?1:0):b[c]), id, req.sid]);
  if(b.active){ await run('UPDATE tt_term SET active=0 WHERE school_id=? AND id<>?',[req.sid, id]); await syncActiveTerm(req.sid); await loadConfig(req.sid); }
  res.json({ok:true});
}));
app.delete('/api/terms/:id', h(async (req,res)=>{ await run('DELETE FROM tt_term WHERE id=? AND school_id=?',[req.params.id, req.sid]); res.json({ok:true}); }));

// ---------- QUOTA (subject periods/week per class) ----------
app.get('/api/quota', h(async (req,res)=>res.json(await q('SELECT * FROM tt_quota WHERE class_id=? AND school_id=?',[req.query.class_id, req.sid]))));
app.put('/api/quota', h(async (req,res)=>{
  const {class_id, subject_id, per_week}=req.body;
  const teacher_id = (req.body.teacher_id===undefined||req.body.teacher_id===''||req.body.teacher_id===null) ? null : +req.body.teacher_id;
  await run(`INSERT INTO tt_quota(class_id,subject_id,per_week,teacher_id,school_id) VALUES(?,?,?,?,?)
     ON CONFLICT(class_id,subject_id) DO UPDATE SET per_week=excluded.per_week, teacher_id=excluded.teacher_id`,[class_id,subject_id,per_week,teacher_id, req.sid]);
  res.json({ok:true});
}));
// Copy one class's Weekly Quota to other classes (e.g. all sections of the same standard). Copies subject + per_week
// only (NOT the teacher pin — each section usually has its own teacher), replacing the targets' existing quota.
app.post('/api/quota/copy', h(async (req,res)=>{
  if(!requireAdmin(req,res)) return;
  const from=Number(req.body.from_class_id);
  let to=Array.isArray(req.body.to_class_ids)?req.body.to_class_ids.map(Number).filter(x=>x&&x!==from):[];
  to=[...new Set(to)];
  if(!from||!to.length){ res.status(400).json({error:'from_class_id and to_class_ids required'}); return; }
  const own=await q1('SELECT 1 FROM tt_class WHERE school_id=? AND id=?',[req.sid,from]);
  if(!own){ res.status(403).json({error:'forbidden'}); return; }
  const src=await q('SELECT subject_id,per_week FROM tt_quota WHERE class_id=? AND school_id=?',[from,req.sid]);
  let done=0;
  for(const tid of to){
    const ok=await q1('SELECT 1 FROM tt_class WHERE school_id=? AND id=?',[req.sid,tid]);
    if(!ok) continue;
    await run('DELETE FROM tt_quota WHERE class_id=? AND school_id=?',[tid,req.sid]);
    for(const r of src){
      await run('INSERT INTO tt_quota(class_id,subject_id,per_week,teacher_id,school_id) VALUES(?,?,?,NULL,?)',[tid,r.subject_id,r.per_week,req.sid]);
    }
    done++;
  }
  res.json({ok:true, copied:done, subjects:src.length});
}));
// Teacher workload from PINNED quota rows (teacher-wise bifurcation): per teacher → subject/class breakdown across ALL classes
app.get('/api/teacher-workload', h(async (req,res)=>{
  res.json(await q('SELECT teacher_id, subject_id, class_id, per_week FROM tt_quota WHERE school_id=? AND teacher_id IS NOT NULL AND per_week>0',[req.sid]));
}));

// ---------- GRID ----------
app.get('/api/timetable', h(async (req,res)=>{
  const rows=req.query.class_id
    ? await q('SELECT * FROM tt_timetable WHERE class_id=? AND school_id=?',[req.query.class_id, req.sid])
    : await q('SELECT * FROM tt_timetable WHERE school_id=?',[req.sid]);
  res.json(rows);
}));

// ---------- CONFLICTS (teacher + room) ----------
async function conflictFor(class_id, day, period, teacher_id, room_id, sid, wk){
  const out={}; wk=+wk||0;
  if(teacher_id){ const r=await q1(`SELECT c.name FROM tt_timetable tt JOIN tt_class c ON c.id=tt.class_id
     WHERE tt.day_of_week=? AND tt.period_index=? AND tt.week_index=? AND tt.teacher_id=? AND tt.class_id<>? AND tt.school_id=?`,[day,period,wk,teacher_id,class_id,sid]);
    if(r) out.teacher=r.name; }
  if(room_id){ const r=await q1(`SELECT c.name FROM tt_timetable tt JOIN tt_class c ON c.id=tt.class_id
     WHERE tt.day_of_week=? AND tt.period_index=? AND tt.week_index=? AND tt.room_id=? AND tt.class_id<>? AND tt.school_id=?`,[day,period,wk,room_id,class_id,sid]);
    if(r) out.room=r.name; }
  return out;
}
app.get('/api/timetable/conflicts', h(async (req,res)=>{
  // combined sessions (shared combine_id) count as ONE booking, so they are not flagged as a teacher/room clash
  const teacher=await q(`SELECT day_of_week,period_index,week_index,teacher_id,string_agg(class_id::text,',') classes,COUNT(*)::int n FROM tt_timetable
     WHERE teacher_id IS NOT NULL AND school_id=? GROUP BY day_of_week,period_index,week_index,teacher_id HAVING COUNT(DISTINCT COALESCE(combine_id,-id))>1`,[req.sid]);
  const room=await q(`SELECT day_of_week,period_index,week_index,room_id,string_agg(class_id::text,',') classes,COUNT(*)::int n FROM tt_timetable
     WHERE room_id IS NOT NULL AND school_id=? GROUP BY day_of_week,period_index,week_index,room_id HAVING COUNT(DISTINCT COALESCE(combine_id,-id))>1`,[req.sid]);
  res.json({teacher, room});
}));
// ---------- MERGED / COMBINED CLASSES (one shared session across several classes) ----------
app.get('/api/combines', h(async (req,res)=> res.json(await q('SELECT * FROM tt_combine WHERE school_id=? ORDER BY day_of_week,period_index,id',[req.sid]))));
app.post('/api/combines', h(async (req,res)=>{
  const b=req.body; const sid=req.sid;
  const ids=(Array.isArray(b.class_ids)?b.class_ids:String(b.class_ids||'').split(',')).map(x=>parseInt(x,10)).filter(n=>!isNaN(n));
  if(ids.length<2) return res.status(400).json({error:'Pick at least two classes to combine.'});
  if(b.subject_id==null||b.day==null||b.period==null) return res.status(400).json({error:'Subject, day and period are required.'});
  const wk=+b.week||0;
  const row=await q1(`INSERT INTO tt_combine(school_id,name,subject_id,teacher_id,room_id,day_of_week,period_index,week_index,class_ids,created_at)
     VALUES(?,?,?,?,?,?,?,?,?,now()::text) RETURNING id`,
     [sid, b.name||null, b.subject_id, b.teacher_id||null, b.room_id||null, +b.day, +b.period, wk, ids.join(',')]);
  // write one locked, combine-tagged cell per member class at the shared slot
  for(const cid of ids){
    await run(`INSERT INTO tt_timetable(class_id,day_of_week,period_index,subject_id,teacher_id,room_id,school_id,week_index,locked,combine_id) VALUES(?,?,?,?,?,?,?,?,1,?)
       ON CONFLICT(class_id,day_of_week,period_index,week_index) DO UPDATE SET subject_id=excluded.subject_id,teacher_id=excluded.teacher_id,room_id=excluded.room_id,locked=1,combine_id=excluded.combine_id`,
      [cid, +b.day, +b.period, b.subject_id, b.teacher_id||null, b.room_id||null, sid, wk, row.id]);
  }
  res.json({id:row.id});
}));
app.delete('/api/combines/:id', h(async (req,res)=>{
  await run('DELETE FROM tt_timetable WHERE combine_id=? AND school_id=?',[req.params.id, req.sid]);
  await run('DELETE FROM tt_combine WHERE id=? AND school_id=?',[req.params.id, req.sid]);
  res.json({ok:true});
}));

// ---------- CELL upsert / delete ----------
app.put('/api/timetable/cell', h(async (req,res)=>{
  const {class_id,day,period,subject_id,teacher_id,room_id}=req.body; const wk=+req.body.week||0;
  const hasLocked=('locked' in req.body); const locVal=hasLocked?(req.body.locked?1:0):0;   // only touch locked when explicitly sent
  await run(`INSERT INTO tt_timetable(class_id,day_of_week,period_index,subject_id,teacher_id,room_id,school_id,week_index,locked) VALUES(?,?,?,?,?,?,?,?,?)
     ON CONFLICT(class_id,day_of_week,period_index,week_index) DO UPDATE SET subject_id=excluded.subject_id,teacher_id=excluded.teacher_id,room_id=excluded.room_id,
       locked = CASE WHEN ?=1 THEN excluded.locked ELSE tt_timetable.locked END`,
    [class_id,day,period,subject_id||null,teacher_id||null,room_id||null, req.sid, wk, locVal, hasLocked?1:0]);
  res.json({ok:true, conflict: await conflictFor(class_id,day,period,teacher_id,room_id, req.sid, wk)});
}));
app.delete('/api/timetable/cell', h(async (req,res)=>{
  await run('DELETE FROM tt_timetable WHERE class_id=? AND day_of_week=? AND period_index=? AND week_index=? AND school_id=?',[req.body.class_id,req.body.day,req.body.period,+req.body.week||0, req.sid]);
  res.json({ok:true});
}));

// ---------- SUBJECT PLACEMENT RULES (honoured by auto-generate) ----------
app.get('/api/rules', h(async (req,res)=>{
  res.json(await q('SELECT * FROM tt_rule WHERE school_id=? ORDER BY id DESC',[req.sid]));
}));
app.post('/api/rules', h(async (req,res)=>{
  const b=req.body||{};
  const type=(b.type==='not_consecutive')?'not_consecutive':'not_same_day';
  if(!b.subject_a_id||!b.subject_b_id||+b.subject_a_id===+b.subject_b_id){ res.status(400).json({error:'pick two different subjects'}); return; }
  const r=await q1('INSERT INTO tt_rule(school_id,type,subject_a_id,subject_b_id,class_id,active,created_at) VALUES(?,?,?,?,?,1,now()::text) RETURNING id',
    [req.sid, type, +b.subject_a_id, +b.subject_b_id, b.class_id?+b.class_id:null]);
  res.json({ok:true,id:r.id});
}));
app.delete('/api/rules/:id', h(async (req,res)=>{
  await run('DELETE FROM tt_rule WHERE id=? AND school_id=?',[req.params.id, req.sid]);
  res.json({ok:true});
}));

// ---------- PER-PERIOD AVAILABILITY (blocked slots for teacher/class/room) ----------
app.get('/api/avail', h(async (req,res)=>{
  const et=req.query.entity_type, eid=+req.query.entity_id;
  if(!et||!eid){ res.json([]); return; }
  res.json(await q('SELECT day_of_week,period_index FROM tt_avail WHERE school_id=? AND entity_type=? AND entity_id=?',[req.sid,et,eid]));
}));
app.post('/api/avail', h(async (req,res)=>{
  const b=req.body||{}; const et=b.entity_type, eid=+b.entity_id, di=+b.day_of_week, pi=+b.period_index;
  if(!['teacher','class','room'].includes(et)||!eid){ res.status(400).json({error:'bad entity'}); return; }
  if(b.blocked){ await run('INSERT INTO tt_avail(school_id,entity_type,entity_id,day_of_week,period_index,created_at) VALUES(?,?,?,?,?,now()::text) ON CONFLICT DO NOTHING',[req.sid,et,eid,di,pi]); }
  else { await run('DELETE FROM tt_avail WHERE school_id=? AND entity_type=? AND entity_id=? AND day_of_week=? AND period_index=?',[req.sid,et,eid,di,pi]); }
  res.json({ok:true});
}));

// ---------- AUTO-GENERATE (quota-aware, room-aware) ----------
app.post('/api/timetable/auto-generate', h(async (req,res)=>{
  const sid=req.sid;
  const classes=await q('SELECT * FROM tt_class WHERE school_id=? ORDER BY id',[sid]);
  const subjects=await q('SELECT * FROM tt_subject WHERE active=1 AND school_id=? ORDER BY id',[sid]);
  const activeSet=new Set(subjects.map(s=>s.id));
  const dblSet=new Set(subjects.filter(s=>+s.double_period===1).map(s=>s.id));   // subjects to place as consecutive double periods (labs / lock-together)
  const coreSet=new Set(subjects.filter(s=>+s.is_core===1).map(s=>s.id));   // ⭐ core/main subjects (Math, Science, English…) → generator spreads them across days + biases them to morning slots; secondary/co-curricular subjects are unrestricted
  const seniorClassSet=new Set(classes.filter(c=>{ const st=_stdOf(c.name); return st!=null && st>=9; }).map(c=>c.id));   // ⭐ senior classes (standard 9+, parsed from class name like DEO Reports does) → non-pinned teacher pick prefers a "Senior teacher" when tied on load
  const rooms=await q('SELECT * FROM tt_room WHERE school_id=? ORDER BY id',[sid]);
  const tmap=await q('SELECT ts.* FROM tt_teacher_subject ts JOIN tt_teacher t ON t.id=ts.teacher_id WHERE t.school_id=?',[sid]);
  const quotas=await q('SELECT * FROM tt_quota WHERE school_id=?',[sid]);
  const absent={}; (await q('SELECT * FROM tt_absence WHERE school_id=?',[sid])).forEach(a=>{(absent[a.teacher_id]=absent[a.teacher_id]||new Set()).add(a.day_of_week);});
  // ---- CLASS-TEACHER FIRST-PERIOD: for each class with the flag ON + a class teacher, pin P1 (period 0) of every working day to that teacher (their main subject) as a locked cell ----
  const ctPinPool={};   // 'wk_cid' -> {subj,n} : how many P1 pins used the class teacher's subject (subtract from that class's quota pool so totals stay correct)
  {
    const dl0=workingDaysArr(sid);
    const cw0=rotMode(sid)?1:Math.max(1,Math.min(4,parseInt((getConfig(sid)||{}).cycle_weeks,10)||1));
    await run('DELETE FROM tt_timetable WHERE school_id=? AND COALESCE(ct_pin,0)=1',[sid]);   // rebuild pins fresh each generate (unchecked classes → no pin)
    const ctClasses=classes.filter(c=> +c.ct_first_period===1 && c.class_teacher_id);
    if(ctClasses.length){
      const tsub={}; (await q('SELECT id,main_subject_id FROM tt_teacher WHERE school_id=?',[sid])).forEach(t=>{ tsub[t.id]=t.main_subject_id; });
      const usedAt={};   // 'wk_di_teacher' → a teacher can't take P1 in two classes at once
      for(const c of ctClasses){
        const tId=c.class_teacher_id; const subj=(tsub[tId]!=null && activeSet.has(tsub[tId]))?tsub[tId]:null;   // class teacher's main subject if active, else homeroom (no subject)
        for(let wk=0; wk<cw0; wk++){
          for(const di of dl0){
            if(!teachingSlots(di,sid).length) continue;                 // no P1 that day
            if(absent[tId]&&absent[tId].has(di)) continue;              // class teacher absent that day
            const key=wk+'_'+di+'_'+tId; if(usedAt[key]) continue; usedAt[key]=1;   // teacher already pinned elsewhere this slot
            // P1=CT is compulsory → the pin OWNS this slot: remove whatever is there (old generated cell OR a manual lock), then place the class teacher
            await run('DELETE FROM tt_timetable WHERE class_id=? AND day_of_week=? AND period_index=0 AND week_index=?',[c.id,di,wk]);
            await run('INSERT INTO tt_timetable(class_id,day_of_week,period_index,subject_id,teacher_id,room_id,school_id,week_index,locked,ct_pin) VALUES(?,?,0,?,?,NULL,?,?,1,1)',[c.id,di,subj,tId,sid,wk]);
            if(subj!=null){ const kk=wk+'_'+c.id; (ctPinPool[kk]=ctPinPool[kk]||{subj,n:0}).n++; }
          }
        }
      }
    }
  }
  // locked/pinned cells — preserved as-is; generator never schedules over them, avoids their teacher/room, and pre-counts their teacher load
  const lockedRows = await q('SELECT class_id,day_of_week,period_index,week_index,teacher_id,room_id,subject_id FROM tt_timetable WHERE school_id=? AND COALESCE(locked,0)=1',[sid]);
  const lockSlot=new Set(); const lockUseAt={}; const lockLoadWk={}; const seenLoad=new Set();
  lockedRows.forEach(r=>{ const w=r.week_index||0;
    lockSlot.add(w+'_'+r.class_id+'_'+r.day_of_week+'_'+r.period_index);
    const sk=w+'_'+r.day_of_week+'_'+r.period_index; const u=lockUseAt[sk]||(lockUseAt[sk]={T:new Set(),R:new Set()});
    if(r.teacher_id) u.T.add(r.teacher_id); if(r.room_id) u.R.add(r.room_id);
    if(r.teacher_id){ const lk=sk+'_'+r.teacher_id; if(!seenLoad.has(lk)){ seenLoad.add(lk); (lockLoadWk[w]=lockLoadWk[w]||[]).push({t:r.teacher_id,di:r.day_of_week}); } }   // count a combined teacher once per slot, not per member class
  });
  // group-splits (electives): reserve the class-slot so no normal subject is scheduled there, and mark each option's teacher & room busy so they aren't double-booked
  const elecRows = await q('SELECT e.class_id,e.day_of_week,e.period_index,e.week_index,o.teacher_id,o.room_id FROM tt_elective e LEFT JOIN tt_elective_option o ON o.elective_id=e.id WHERE e.school_id=?',[sid]);
  elecRows.forEach(r=>{ const w=r.week_index||0;
    lockSlot.add(w+'_'+r.class_id+'_'+r.day_of_week+'_'+r.period_index);
    const sk=w+'_'+r.day_of_week+'_'+r.period_index; const u=lockUseAt[sk]||(lockUseAt[sk]={T:new Set(),R:new Set()});
    if(r.teacher_id) u.T.add(r.teacher_id); if(r.room_id) u.R.add(r.room_id);
    if(r.teacher_id){ const lk=sk+'_'+r.teacher_id; if(!seenLoad.has(lk)){ seenLoad.add(lk); (lockLoadWk[w]=lockLoadWk[w]||[]).push({t:r.teacher_id,di:r.day_of_week}); } }
  });
  // per-period availability blocks
  const teacherBlock={}, classBlock={}, roomBlock={};
  (await q('SELECT entity_type,entity_id,day_of_week,period_index FROM tt_avail WHERE school_id=?',[sid])).forEach(a=>{
    const k=a.day_of_week+'_'+a.period_index;
    const m=a.entity_type==='teacher'?teacherBlock:a.entity_type==='class'?classBlock:roomBlock;
    (m[a.entity_id]=m[a.entity_id]||new Set()).add(k);
  });
  const blocked=(m,id,di,pi)=> !!(m[id]&&m[id].has(di+'_'+pi));
  // optimizer preferences (soft): spread subjects across days + balance teacher load + morning-core + minimise gaps
  const optCfg=getConfig(sid)||{};
  const optSpread=(optCfg.opt_spread==null?1:+optCfg.opt_spread)!==0;
  const optBalance=(optCfg.opt_balance==null?1:+optCfg.opt_balance)!==0;
  const optMorning=(+optCfg.opt_morning||0)!==0;
  const optGap=(+optCfg.opt_gap||0)!==0;
  const optSolver=(+optCfg.opt_solver||0)!==0;
  const cycleWeeks= rotMode(sid) ? 1 : Math.max(1,Math.min(4, parseInt(optCfg.cycle_weeks,10)||1));   // day-rotation uses a single week (days = rotation days)
  const dayList = workingDaysArr(sid);   // rotation mode → [0..rotDays-1]; else the school's working weekdays
  const slotCount={}; dayList.forEach(di=>{ slotCount[di]=teachingSlots(di,sid).length; });
  const maxLoad={}, capDay={}, capCons={}, seniorT=new Set();
  (await q('SELECT id,max_load,max_per_day,max_consecutive,is_senior FROM tt_teacher WHERE school_id=?',[sid])).forEach(t=>{
    maxLoad[t.id]=t.max_load||999;
    capDay[t.id]=(t.max_per_day&&t.max_per_day>0)?t.max_per_day:Infinity;
    capCons[t.id]=(t.max_consecutive&&t.max_consecutive>0)?t.max_consecutive:Infinity;
    if(+t.is_senior===1) seniorT.add(t.id);   // ⭐ senior teacher → preferred (over an equally-loaded non-senior teacher) for senior-class (std 9+) assignments
  });
  const load={}, dayCount={}, lastPi={}, consRun={};   // per-teacher weekly/day counts + consecutive tracking
  const predRun=(t,di,pi)=> (lastPi[t+'_'+di]===pi-1 ? (consRun[t+'_'+di]||0)+1 : 1);   // consecutive run if placed at (di,pi)
  const capOk=(t,di,pi)=> (dayCount[t+'_'+di]||0)<capDay[t] && predRun(t,di,pi)<=capCons[t];
  const teachersForSubject=sid=>tmap.filter(m=>m.subject_id===sid).map(m=>m.teacher_id);
  const tsubSeq={}; tmap.forEach(m=>{ tsubSeq[m.teacher_id+'_'+m.subject_id]=(m.seq==null?99:m.seq); });   // teacher's priority for a subject (0=1st) → auto-generate fills higher-priority subjects first

  // subject placement rules (not_same_day / not_consecutive), resolved per class
  const allRules=await q('SELECT * FROM tt_rule WHERE school_id=? AND active=1',[sid]);
  const rkey=(a,b)=>a<b?a+'_'+b:b+'_'+a;
  const clsRules={};
  classes.forEach(c=>{
    const rs=allRules.filter(r=>r.class_id==null||r.class_id===c.id);
    const sameDay=new Set(), notAdj=new Set();
    rs.forEach(r=>{ const k=rkey(r.subject_a_id,r.subject_b_id); if(r.type==='not_consecutive') notAdj.add(k); else sameDay.add(k); });
    clsRules[c.id]={sameDay,notAdj};
  });

  // per-class subject "weight" = weekly quota (higher = more core) for the morning-bias optimizer, boosted for ⭐ core subjects so they always win the early-period tie-break over secondary/co-curricular subjects regardless of their quota count
  const subjWeight={};
  classes.forEach(c=>{ const w={}; quotas.filter(x=>x.class_id===c.id).forEach(x=>{ w[x.subject_id]=(w[x.subject_id]||0)+(+x.per_week||0); }); Object.keys(w).forEach(k=>{ if(coreSet.has(+k)) w[k]+=1000; }); subjWeight[c.id]=w; });

  // A subject is only scheduled if it has at least one teacher connected. Subjects with NO teacher are
  // ignored entirely (not counted, not placed) so extra/unstaffed subjects never leave blank cells.
  const staffedSet=new Set(tmap.map(m=>m.subject_id));
  const staffedSubjects=subjects.filter(s=>staffedSet.has(s.id));
  // TEACHER PIN (teacher-wise bifurcation): a quota row can name a specific teacher → that class+subject's periods MUST use that teacher (never auto-balanced onto someone else). Blank = auto.
  const pinT={}; quotas.forEach(x=>{ if(x.teacher_id!=null) pinT[x.class_id+'_'+x.subject_id]=x.teacher_id; });
  // build per-class subject pool honouring quota (fallback: even rotation) — staffed subjects only
  function poolFor(cid, totalSlots){
    const qs=quotas.filter(x=>x.class_id===cid && activeSet.has(x.subject_id) && staffedSet.has(x.subject_id));
    let pool=[];
    // Start with the class's Weekly Quota (its real curriculum)…
    if(qs.length){ qs.forEach(x=>{ for(let i=0;i<x.per_week;i++) pool.push(x.subject_id); }); }
    // …then FILL the rest of the week using ONLY this class's OWN quota subjects — never school-wide
    // subjects — so a class can never receive a subject outside its curriculum (e.g. Accountancy in a
    // 9th-standard class). A class with no quota set stays blank (a clear signal to set its Weekly Quota).
    const ownSubjects=[...new Set(qs.map(x=>x.subject_id))];
    if(ownSubjects.length && pool.length<totalSlots){ let i=0; while(pool.length<totalSlots){ pool.push(ownSubjects[i%ownSubjects.length]); i++; } }
    return pool.slice(0,totalSlots);
  }

  // compute the assignment in memory, then persist in one transaction
  let totalSlots=0; dayList.forEach(di=>totalSlots+=teachingSlots(di, sid).length);
  const remaining={}, daySubs={}, prevSub={};   // per class: pool left; per class+day: subjects placed + previous subject
  // can this class+subject actually get a teacher in THIS slot right now? (pinned teacher, or any qualified free teacher)
  function canStaffNow(cid, subjId, di, pi, usedT){
    const pinned=pinT[cid+'_'+subjId];
    if(pinned!=null) return !usedT.has(pinned) && !(absent[pinned]&&absent[pinned].has(di)) && !blocked(teacherBlock,pinned,di,pi) && capOk(pinned,di,pi);
    return teachersForSubject(subjId).some(t=>!usedT.has(t)&&!(absent[t]&&absent[t].has(di))&&!blocked(teacherBlock,t,di,pi)&&(load[t]||0)<maxLoad[t]&&capOk(t,di,pi));
  }
  function pickSubject(cid,di,pi,usedT){
    const rem=remaining[cid]; if(!rem.length) return null;
    const R=clsRules[cid]; const ds=daySubs[cid+'_'+di]; const prev=prevSub[cid+'_'+di];
    const ruleOk=(s)=>{ if(ds){ for(const x of ds){ if(R.sameDay.has(rkey(s,x))){ return false; } } if(coreSet.has(s) && ds.has(s)) return false; }   // ⭐ core subjects: max once per day per class (falls back gracefully like any other rule if truly unavoidable)
      if(prev!=null && R.notAdj.has(rkey(s,prev))) return false; return true; };
    // rule-valid candidate indices; then prefer ones not yet placed today (spread)
    const cand=[]; for(let i=0;i<rem.length;i++){ if(ruleOk(rem[i])) cand.push(i); }
    let pool=cand;
    if(optSpread && cand.length){ const fresh=cand.filter(i=>!(ds&&ds.has(rem[i]))); if(fresh.length) pool=fresh; }
    // ⭐ prefer subjects whose teacher is FREE this slot — stops a single-teacher subject (e.g. Drawing→only Dharmesh) being chosen for two classes in the same slot → far fewer unstaffed cells
    if(usedT && pool.length){ const staffable=pool.filter(i=>canStaffNow(cid,rem[i],di,pi,usedT)); if(staffable.length) pool=staffable; }
    let idx;
    if(optMorning && pool.length){   // heavier (higher-quota) subjects earlier, lighter later
      const w=subjWeight[cid]||{}; const n=slotCount[di]||8; const early=pi < n/2;
      idx=pool.reduce((best,i)=>{ const wi=w[rem[i]]||0, wb=w[rem[best]]||0; return (early? wi>wb : wi<wb) ? i : best; }, pool[0]);
    } else { idx = pool.length ? pool[0] : -1; }
    if(idx<0) idx=0;   // nothing satisfies the rules → take next to avoid leaving a gap
    return {s:rem[idx], idx, staffable: (usedT?canStaffNow(cid,rem[idx],di,pi,usedT):true)};   // PEEK only — caller commits (or defers to a later slot if it can't be staffed now)
  }
  function commitPick(cid,di,idx){   // actually consume the peeked subject + record it for spread/rule tracking
    const rem=remaining[cid]; const s=rem.splice(idx,1)[0];
    (daySubs[cid+'_'+di]=daySubs[cid+'_'+di]||new Set()).add(s);
    prevSub[cid+'_'+di]=s;
    return s;
  }
  const clearState=()=>{ [remaining,daySubs,prevSub,load,dayCount,lastPi,consRun].forEach(o=>{ for(const k in o) delete o[k]; }); };
  const toInsert=[];
  for(let wk=0; wk<cycleWeeks; wk++){                       // generate each week of the cycle independently
    clearState();
    // ⭐ seed daySubs with already-LOCKED core subjects (e.g. the class-teacher P1 pin) so the "core subject max once/day" rule in pickSubject knows about them — otherwise a core subject that's also a CT-pin (or any other locked cell) could get placed a 2nd time that same day by the normal generator, since locked cells bypass commitPick entirely
    lockedRows.forEach(r=>{ if((r.week_index||0)===wk && r.subject_id!=null && coreSet.has(r.subject_id)){ (daySubs[r.class_id+'_'+r.day_of_week]=daySubs[r.class_id+'_'+r.day_of_week]||new Set()).add(r.subject_id); } });
    classes.forEach(c=>{ let pool=shuffleStable(poolFor(c.id,totalSlots), c.id + wk*1009).slice();   // vary shuffle per week so weeks differ
      const pp=ctPinPool[wk+'_'+c.id]; if(pp){ let n=pp.n; for(let i=pool.length-1;i>=0 && n>0;i--){ if(pool[i]===pp.subj){ pool.splice(i,1); n--; } } }   // the class-teacher P1 pins already cover n of this subject → drop from pool so quota total is exact
      remaining[c.id]=pool; });
    (lockLoadWk[wk]||[]).forEach(x=>{ load[x.t]=(load[x.t]||0)+1; dayCount[x.t+'_'+x.di]=(dayCount[x.t+'_'+x.di]||0)+1; });   // count locked teacher load up front
    const weekRows=[]; const carry={};
    dayList.forEach(di=>{
      const slots=teachingSlots(di, sid);
      for(const k in carry) delete carry[k];   // consecutive doubles never span across days
      slots.forEach((_,pi)=>{
        const usedT=new Set(), usedR=new Set();
        const lu=lockUseAt[wk+'_'+di+'_'+pi]; if(lu){ lu.T.forEach(x=>usedT.add(x)); lu.R.forEach(x=>usedR.add(x)); }   // locked teacher/room busy this slot
        classes.forEach(c=>{
          const cr=carry[c.id];
          if(cr && cr.di===di && cr.pi===pi){   // second half of a consecutive double period — same subject, teacher & room
            carry[c.id]=null;
            const t2=(cr.teacher_id && !usedT.has(cr.teacher_id) && !blocked(teacherBlock,cr.teacher_id,di,pi) && !(absent[cr.teacher_id]&&absent[cr.teacher_id].has(di))) ? cr.teacher_id : null;
            const r2=(cr.room_id && !usedR.has(cr.room_id)) ? cr.room_id : ((rooms.find(r=>!usedR.has(r.id)&&!blocked(roomBlock,r.id,di,pi))||{}).id ?? null);
            weekRows.push({cid:c.id,di,pi,subject_id:cr.subject_id,teacher_id:t2,room_id:r2,dbl:true});
            if(t2){usedT.add(t2);load[t2]=(load[t2]||0)+1;dayCount[t2+'_'+di]=(dayCount[t2+'_'+di]||0)+1;consRun[t2+'_'+di]=(lastPi[t2+'_'+di]===pi-1?(consRun[t2+'_'+di]||0)+1:1);lastPi[t2+'_'+di]=pi;} if(r2)usedR.add(r2);
            return;
          }
          if(lockSlot.has(wk+'_'+c.id+'_'+di+'_'+pi)) return;   // fixed/locked cell → keep as-is, don't reschedule
          if(blocked(classBlock,c.id,di,pi)) return;   // class marked unavailable this slot → leave empty
          const pick=pickSubject(c.id,di,pi,usedT);
          if(pick==null) return;   // class has no schedulable subjects
          const subjId=pick.s;
          let t;
          const pinned=pinT[c.id+'_'+subjId];
          if(pinned!=null){
            // teacher-wise bifurcation: force the pinned teacher. Use them if free this slot (not double-booked, not absent, not blocked, within day/consecutive caps). If busy, leave the cell unstaffed rather than substituting someone else — it shows in the Generate Log so the user can adjust the split.
            t=(!usedT.has(pinned)&&!(absent[pinned]&&absent[pinned].has(di))&&!blocked(teacherBlock,pinned,di,pi)&&capOk(pinned,di,pi))?pinned:null;
          } else {
            let opts=teachersForSubject(subjId).filter(t=>!usedT.has(t)&&!(absent[t]&&absent[t].has(di))&&!blocked(teacherBlock,t,di,pi)&&(load[t]||0)<maxLoad[t]&&capOk(t,di,pi));
            opts.sort((a,b)=>{
              // LEAST-LOADED FIRST → a subject shared by many teachers (e.g. PT by Dharmesh & Bharat) is spread evenly instead of dumped on one.
              const d=(load[a]||0)-(load[b]||0); if(d) return d;
              if(seniorClassSet.has(c.id)){ const sa=seniorT.has(a)?0:1, sb=seniorT.has(b)?0:1; if(sa!==sb) return sa-sb; }   // ⭐ tie on load, senior class (std 9+) → prefer the senior teacher
              const sp=(tsubSeq[a+'_'+subjId]??99)-(tsubSeq[b+'_'+subjId]??99); if(sp) return sp;   // tie → prefer the teacher whose higher-priority (main) subject it is
              if(optGap){ const ga=(lastPi[a+'_'+di]===pi-1?0:1), gb=(lastPi[b+'_'+di]===pi-1?0:1); if(ga!==gb) return ga-gb; }  // then fewer gaps
              return 0;
            });
            t=opts[0] ?? teachersForSubject(subjId).find(x=>!usedT.has(x)&&!blocked(teacherBlock,x,di,pi)&&(load[x]||0)<maxLoad[x]&&capOk(x,di,pi)) ?? null;
          }
          // ⭐ DEFER unstaffable picks: keep the subject in the pool for a LATER slot where a teacher is free, instead of dumping it here as an unstaffed cell. Every pool subject has ≥1 qualified teacher (poolFor only adds staffed subjects), so t==null just means "teachers busy this slot". Genuine leftovers that never find a free slot are placed as unstaffed in a final pass (below) so they stay visible in the Generate Log. This spreads a single-teacher subject (Drawing→only Dharmesh) across the week instead of clustering it.
          if(t==null) return;
          commitPick(c.id,di,pick.idx);
          const room=rooms.find(r=>!usedR.has(r.id)&&!blocked(roomBlock,r.id,di,pi));
          // start a consecutive double period when this subject wants one and the next teaching period is free for this class
          const canDbl = dblSet.has(subjId) && (pi+1)<slots.length
              && !lockSlot.has(wk+'_'+c.id+'_'+di+'_'+(pi+1)) && !blocked(classBlock,c.id,di,pi+1)
              && remaining[c.id].includes(subjId) && (!t || (capCons[t]||Infinity)>=2);
          weekRows.push({cid:c.id,di,pi,subject_id:subjId,teacher_id:t,room_id:room?room.id:null,dbl:canDbl});
          if(t){usedT.add(t);load[t]=(load[t]||0)+1;
            dayCount[t+'_'+di]=(dayCount[t+'_'+di]||0)+1;
            consRun[t+'_'+di]=(lastPi[t+'_'+di]===pi-1?(consRun[t+'_'+di]||0)+1:1);
            lastPi[t+'_'+di]=pi;} if(room)usedR.add(room.id);
          if(canDbl){   // consume the second half from the pool and carry it into the very next period
            const ix=remaining[c.id].indexOf(subjId); if(ix>=0) remaining[c.id].splice(ix,1);
            (daySubs[c.id+'_'+di]=daySubs[c.id+'_'+di]||new Set()).add(subjId);
            carry[c.id]={di,pi:pi+1,subject_id:subjId,teacher_id:t,room_id:room?room.id:null};
          }
        });
      });
    });
    // FINAL PASS: subjects that could not be staffed in ANY slot this week are still in the pool → drop them into each class's remaining EMPTY teaching slots as unstaffed cells, so a genuine shortage stays visible in the Generate Log (deferring above only stops PREMATURE unstaffed cells; it must not hide a real shortage).
    classes.forEach(c=>{ const rem=remaining[c.id]; if(!rem||!rem.length) return;
      const occupied=new Set(weekRows.filter(r=>r.cid===c.id).map(r=>r.di+'_'+r.pi));
      for(const di of dayList){ const dslots=teachingSlots(di, sid);
        for(let pi=0; pi<dslots.length && rem.length; pi++){ const key=di+'_'+pi;
          if(occupied.has(key)) continue;
          if(lockSlot.has(wk+'_'+c.id+'_'+di+'_'+pi)) continue;
          if(blocked(classBlock,c.id,di,pi)) continue;
          const s=rem.shift(); weekRows.push({cid:c.id,di,pi,subject_id:s,teacher_id:null,room_id:null}); occupied.add(key);
        }
      }
    });
    if(optSolver){ const fixedT={},fixedR={}; for(const k in lockUseAt){ const p=k.split('_'); if(+p[0]===wk){ const sk=p[1]+'_'+p[2]; (fixedT[sk]=fixedT[sk]||new Set()); lockUseAt[k].T.forEach(x=>fixedT[sk].add(x)); (fixedR[sk]=fixedR[sk]||new Set()); lockUseAt[k].R.forEach(x=>fixedR[sk].add(x)); } }
      optimizeWeek(weekRows, {teacherBlock,roomBlock,absent,capCons,subjWeight,clsRules,rkey,fixedT,fixedR}); }   // deep local-search improvement pass (locked cells kept busy)
    for(const r of weekRows) toInsert.push([r.cid,r.di,r.pi,r.subject_id,r.teacher_id,r.room_id,sid,wk]);
  }

  await tx(async (cq)=>{
    await cq('DELETE FROM tt_timetable WHERE school_id=? AND COALESCE(locked,0)=0',[sid]);   // keep locked/pinned cells
    for(const r of toInsert)
      await cq('INSERT INTO tt_timetable(class_id,day_of_week,period_index,subject_id,teacher_id,room_id,school_id,week_index) VALUES(?,?,?,?,?,?,?,?)', r);
  });
  const n=(await q1('SELECT COUNT(*)::int AS n FROM tt_timetable WHERE school_id=?',[sid])).n;
  // subjects that got scheduled but have NO teacher assigned in the master → those periods run teacher-less
  const unstaffed=(await q(`SELECT s.name, COUNT(*)::int AS n FROM tt_timetable tt JOIN tt_subject s ON s.id=tt.subject_id WHERE tt.school_id=? AND tt.teacher_id IS NULL AND tt.subject_id IS NOT NULL GROUP BY s.name ORDER BY n DESC, s.name`,[sid])).map(r=>r.name+' ('+r.n+')');
  res.json({ok:true, cells:n, unstaffed});
}));
function shuffleStable(arr, seed){ // deterministic light shuffle so subjects spread out
  const a=arr.slice(); let s=seed*9301+49297;
  for(let i=a.length-1;i>0;i--){ s=(s*9301+49297)%233280; const j=Math.floor(s/233280*(i+1)); [a[i],a[j]]=[a[j],a[i]]; }
  return a;
}
// Deterministic within-day local-search optimiser. Mutates `rows` (one week) in place, returns swaps applied.
// Swaps the CONTENTS (subject/teacher/room) of two teaching periods of the SAME class on the SAME day when it
// lowers a penalty (teacher idle gaps + heavy-subject-late) AND breaks no hard constraint. Same-day swaps keep
// each teacher's weekly + per-day totals and every same-day rule intact, so only adjacency/consecutive need re-checking.
function optimizeWeek(rows, P){
  const teacherBlock=P.teacherBlock||{}, roomBlock=P.roomBlock||{}, absent=P.absent||{};
  const capCons=P.capCons||{}, subjWeight=P.subjWeight||{}, clsRules=P.clsRules||{}, rkey=P.rkey;
  const Wgap=P.Wgap==null?3:P.Wgap, Wmorn=P.Wmorn==null?1:P.Wmorn, maxPasses=P.maxPasses||8;
  const K=(di,pi)=>di+'_'+pi;
  const tSlot={}, rSlot={};          // slotKey -> Map(entityId -> cid)
  const tDay={};                     // teacher_di -> Set(pi) across all classes
  const byCD={};                     // cid -> di -> [rows]
  for(const r of rows){
    if(r.teacher_id){ (tSlot[K(r.di,r.pi)]=tSlot[K(r.di,r.pi)]||new Map()).set(r.teacher_id,r.cid); (tDay[r.teacher_id+'_'+r.di]=tDay[r.teacher_id+'_'+r.di]||new Set()).add(r.pi); }
    if(r.room_id){ (rSlot[K(r.di,r.pi)]=rSlot[K(r.di,r.pi)]||new Map()).set(r.room_id,r.cid); }
    ((byCD[r.cid]=byCD[r.cid]||{})[r.di]=byCD[r.cid][r.di]||[]).push(r);
  }
  // seed locked/pinned teacher+room usage with sentinel cid -1 so no swap ever collides with a locked cell
  if(P.fixedT) for(const k in P.fixedT){ const m=tSlot[k]=tSlot[k]||new Map(); P.fixedT[k].forEach(id=>{ if(!m.has(id)) m.set(id,-1); }); }
  if(P.fixedR) for(const k in P.fixedR){ const m=rSlot[k]=rSlot[k]||new Map(); P.fixedR[k].forEach(id=>{ if(!m.has(id)) m.set(id,-1); }); }
  const blk=(m,id,di,pi)=> !!(id&&m[id]&&m[id].has(di+'_'+pi));
  const gapOf=(t,di)=>{ const s=tDay[t+'_'+di]; if(!s||s.size<=1)return 0; let mn=Infinity,mx=-Infinity; s.forEach(p=>{if(p<mn)mn=p;if(p>mx)mx=p;}); return (mx-mn+1)-s.size; };
  const maxRunOf=(t,di)=>{ const s=tDay[t+'_'+di]; if(!s||!s.size)return 0; const a=[...s].sort((x,y)=>x-y); let r=1,m=1; for(let i=1;i<a.length;i++){ r=(a[i]===a[i-1]+1)?r+1:1; if(r>m)m=r; } return m; };
  const w=(cid,sub)=> (subjWeight[cid]&&sub!=null)?(subjWeight[cid][sub]||0):0;
  // adjacency (not_consecutive) ok for class C, day di if subject `sub` sits at period pi
  const adjOk=(cid,di,pi,sub)=>{ if(sub==null)return true; const R=clsRules[cid]; if(!R||!R.notAdj||!R.notAdj.size)return true;
    const day=byCD[cid]&&byCD[cid][di]; if(!day)return true;
    for(const nb of day){ if(nb.subject_id==null)continue; if(nb.pi===pi-1||nb.pi===pi+1){ if(R.notAdj.has(rkey(sub,nb.subject_id))) return false; } }
    return true; };
  let swaps=0;
  for(let pass=0; pass<maxPasses; pass++){
    let improved=false;
    for(const cidKey in byCD){ const cid=+cidKey; for(const diKey in byCD[cid]){ const di=+diKey;
      const cells=byCD[cid][di]; if(cells.length<2) continue;
      for(let i=0;i<cells.length;i++) for(let j=i+1;j<cells.length;j++){
        const A=cells[i], B=cells[j]; if(A.pi===B.pi) continue;
        if(A.dbl||B.dbl) continue;   // never break a consecutive double period apart
        const pa=A.pi, pb=B.pi, TA=A.teacher_id, TB=B.teacher_id, RA=A.room_id, RB=B.room_id, SA=A.subject_id, SB=B.subject_id;
        if(TA===TB && RA===RB && SA===SB) continue;
        // hard constraints after moving A->pb, B->pa (same class, same day)
        // teacher TB into pa: no OTHER class uses TB at (di,pa); not blocked/absent
        if(TB){ const m=tSlot[K(di,pa)]; if(m){const u=m.get(TB); if(u!=null&&u!==cid) continue;} if(blk(teacherBlock,TB,+di,pa)) continue; if(absent[TB]&&absent[TB].has(+di)) continue; }
        if(TA){ const m=tSlot[K(di,pb)]; if(m){const u=m.get(TA); if(u!=null&&u!==cid) continue;} if(blk(teacherBlock,TA,+di,pb)) continue; if(absent[TA]&&absent[TA].has(+di)) continue; }
        if(RB){ const m=rSlot[K(di,pa)]; if(m){const u=m.get(RB); if(u!=null&&u!==cid) continue;} if(blk(roomBlock,RB,+di,pa)) continue; }
        if(RA){ const m=rSlot[K(di,pb)]; if(m){const u=m.get(RA); if(u!=null&&u!==cid) continue;} if(blk(roomBlock,RA,+di,pb)) continue; }
        // tentatively update tDay for TA,TB then check consecutive caps + adjacency, measure gap delta
        const before = (TA?gapOf(TA,di):0)+(TB&&TB!==TA?gapOf(TB,di):0);
        const moveT=(t,fromPi,toPi)=>{ if(!t)return; const s=tDay[t+'_'+di]; if(s){ s.delete(fromPi); s.add(toPi); } };
        // apply tentative teacher-day move (A:TA pa->pb, B:TB pb->pa)
        moveT(TA,pa,pb); moveT(TB,pb,pa);
        let ok = ( (!TA||maxRunOf(TA,di)<=(capCons[TA]||Infinity)) && (!TB||maxRunOf(TB,di)<=(capCons[TB]||Infinity)) );
        const after = ok ? ((TA?gapOf(TA,di):0)+(TB&&TB!==TA?gapOf(TB,di):0)) : 0;
        // revert teacher-day move (validity/penalty measured; real apply happens only if accepted)
        moveT(TA,pb,pa); moveT(TB,pa,pb);
        if(!ok) continue;
        // adjacency rule: subject SB now at pa, SA now at pb (temporarily reflect by testing against the OTHER cells)
        // build a quick view excluding A,B then test
        if(!adjOkSwap(byCD[cid][di],A,B,pa,pb,clsRules[cid],rkey)) continue;
        // penalty delta
        const gapDelta = after - before;
        const mornDelta = (pa-pb)*(w(cid,SB)-w(cid,SA));
        const delta = Wgap*gapDelta + Wmorn*mornDelta;
        if(delta < -1e-9){
          // APPLY: swap contents in the row objects, update occupancy maps + tDay
          if(TA){ tSlot[K(di,pa)].delete(TA); } if(RA){ rSlot[K(di,pa)].delete(RA); }
          if(TB){ tSlot[K(di,pb)].delete(TB); } if(RB){ rSlot[K(di,pb)].delete(RB); }
          const a2={s:A.subject_id,t:A.teacher_id,r:A.room_id};
          A.subject_id=B.subject_id; A.teacher_id=B.teacher_id; A.room_id=B.room_id;
          B.subject_id=a2.s; B.teacher_id=a2.t; B.room_id=a2.r;
          if(A.teacher_id){(tSlot[K(di,pa)]=tSlot[K(di,pa)]||new Map()).set(A.teacher_id,cid);} if(A.room_id){(rSlot[K(di,pa)]=rSlot[K(di,pa)]||new Map()).set(A.room_id,cid);}
          if(B.teacher_id){(tSlot[K(di,pb)]=tSlot[K(di,pb)]||new Map()).set(B.teacher_id,cid);} if(B.room_id){(rSlot[K(di,pb)]=rSlot[K(di,pb)]||new Map()).set(B.room_id,cid);}
          moveT(TA,pa,pb); moveT(TB,pb,pa);
          swaps++; improved=true;
        }
      }
    }}
    if(!improved) break;
  }
  return swaps;
}
// adjacency rule check for a tentative same-day swap of cells A(pa) and B(pb): SB goes to pa, SA goes to pb
function adjOkSwap(dayCells, A, B, pa, pb, R, rkey){
  if(!R||!R.notAdj||!R.notAdj.size) return true;
  const SA=A.subject_id, SB=B.subject_id;
  const at=(pi)=>{ if(pi===pa) return SB; if(pi===pb) return SA; const c=dayCells.find(x=>x.pi===pi); return c?c.subject_id:null; };
  const chk=(pi,sub)=>{ if(sub==null)return true; const a=at(pi-1), b=at(pi+1);
    if(a!=null&&R.notAdj.has(rkey(sub,a)))return false; if(b!=null&&R.notAdj.has(rkey(sub,b)))return false; return true; };
  return chk(pa,SB)&&chk(pb,SA);
}

// ---------- TIMETABLE VERSIONS (saved snapshots) = per-term "windows" ----------
const HOURS_FIELDS=['weekday_start','weekday_end','saturday_start','saturday_end','period_minutes','period_durations','lunch_after','lunch_minutes','short_break_after','short_break_minutes','working_days','num_periods','sat_num_periods','sat_period_minutes','sat_period_durations','sat_lunch_after','sat_lunch_minutes','sat_short_break_after','sat_short_break_minutes','academic_session'];
app.get('/api/versions', h(async (req,res)=>{
  res.json(await q(`SELECT s.id,s.name,s.session,s.created_at,s.cell_count,s.term_id,(s.config_json IS NOT NULL) AS has_config, t.name AS term_name
     FROM tt_snapshot s LEFT JOIN tt_term t ON t.id=s.term_id WHERE s.school_id=? ORDER BY s.id DESC`,[req.sid]));
}));
// save the current live timetable + School Hours as a named version, bound to the active term
app.post('/api/versions', h(async (req,res)=>{
  const sid=req.sid;
  const name=(req.body.name||'').trim()||('Version '+((await q1('SELECT COUNT(*)::int n FROM tt_snapshot WHERE school_id=?',[sid])).n+1));
  const cfg=getConfig(sid)||{};
  const term=await q1('SELECT id,name FROM tt_term WHERE school_id=? AND active=1 ORDER BY id LIMIT 1',[sid]);
  const session=(term&&term.name)||cfg.academic_session||null;
  const configSnap={}; HOURS_FIELDS.forEach(k=>configSnap[k]=cfg[k]);
  const config_json=JSON.stringify(configSnap);
  const cells=await q('SELECT class_id,day_of_week,period_index,subject_id,teacher_id,room_id,week_index FROM tt_timetable WHERE school_id=?',[sid]);
  const snap=await tx(async (cq,cq1)=>{
    const s=await cq1('INSERT INTO tt_snapshot(name,session,created_at,cell_count,school_id,term_id,config_json) VALUES(?,?,now()::text,?,?,?,?) RETURNING id',[name,session,cells.length,sid,(term&&term.id)||null,config_json]);
    for(const c of cells)
      await cq('INSERT INTO tt_snapshot_cell(snapshot_id,class_id,day_of_week,period_index,subject_id,teacher_id,room_id,week_index) VALUES(?,?,?,?,?,?,?,?)',
        [s.id,c.class_id,c.day_of_week,c.period_index,c.subject_id,c.teacher_id,c.room_id,c.week_index||0]);
    return s;
  });
  res.json({ok:true, id:snap.id, name, cells:cells.length});
}));
// restore a saved version → overwrites the live timetable AND its School Hours (window switch), and re-activates its term
app.post('/api/versions/:id/restore', h(async (req,res)=>{
  const id=Number(req.params.id), sid=req.sid;
  const snap=await q1('SELECT * FROM tt_snapshot WHERE id=? AND school_id=?',[id,sid]);
  if(!snap){ res.status(404).json({error:'not found'}); return; }
  const cells=await q('SELECT class_id,day_of_week,period_index,subject_id,teacher_id,room_id,week_index FROM tt_snapshot_cell WHERE snapshot_id=?',[id]);
  await tx(async (cq)=>{
    await cq('DELETE FROM tt_timetable WHERE school_id=?',[sid]);
    for(const c of cells)
      await cq('INSERT INTO tt_timetable(class_id,day_of_week,period_index,subject_id,teacher_id,room_id,school_id,week_index) VALUES(?,?,?,?,?,?,?,?)',
        [c.class_id,c.day_of_week,c.period_index,c.subject_id,c.teacher_id,c.room_id,sid,c.week_index||0]);
  });
  // restore the School Hours captured with this window
  let hoursRestored=false;
  if(snap.config_json){ try{ const cs=JSON.parse(snap.config_json); const ks=HOURS_FIELDS.filter(k=>k in cs);
    if(ks.length){ await run(`UPDATE tt_config SET ${ks.map(k=>k+'=?').join(',')} WHERE school_id=?`,[...ks.map(k=>cs[k]), sid]); hoursRestored=true; } }catch(e){} }
  // re-activate this window's term (switch the active term to match)
  if(snap.term_id){ await run('UPDATE tt_term SET active=CASE WHEN id=? THEN 1 ELSE 0 END WHERE school_id=?',[snap.term_id, sid]); await syncActiveTerm(sid); }
  await loadConfig(sid);
  res.json({ok:true, cells:cells.length, hoursRestored});
}));
app.put('/api/versions/:id', h(async (req,res)=>{
  if(req.body.name!==undefined) await run('UPDATE tt_snapshot SET name=? WHERE id=? AND school_id=?',[req.body.name, req.params.id, req.sid]);
  res.json({ok:true});
}));
app.delete('/api/versions/:id', h(async (req,res)=>{
  await run('DELETE FROM tt_snapshot_cell WHERE snapshot_id=?',[req.params.id]);
  await run('DELETE FROM tt_snapshot WHERE id=? AND school_id=?',[req.params.id, req.sid]);
  res.json({ok:true});
}));
// ---------- LOCAL-DRIVE backup (download the live timetable as a restorable JSON file; keeps cloud storage low) ----------
app.get('/api/timetable/backup', h(async (req,res)=>{
  const sid=req.sid;
  const cfg=getConfig(sid)||{};
  const config={}; HOURS_FIELDS.forEach(k=>{ config[k]=cfg[k]; });
  const cells=await q('SELECT class_id,day_of_week,period_index,subject_id,teacher_id,room_id,week_index FROM tt_timetable WHERE school_id=?',[sid]);
  const term=await q1('SELECT id,name FROM tt_term WHERE school_id=? AND active=1 ORDER BY id LIMIT 1',[sid]);
  res.json({ format:'aumtara-timetable-backup', schema:1, name:null, session:(term&&term.name)||cfg.academic_session||null, cell_count:cells.length, config, cells });
}));
// restore the live timetable (+ School Hours) from an uploaded backup JSON — same overwrite behaviour as a cloud version restore, but the file lives on the user's computer
app.post('/api/timetable/restore-backup', h(async (req,res)=>{
  const sid=req.sid;
  const body=req.body||{};
  if(body.format && body.format!=='aumtara-timetable-backup'){ res.status(400).json({error:'not a timetable backup file'}); return; }
  const cells=Array.isArray(body.cells)?body.cells:[];
  await tx(async (cq)=>{
    await cq('DELETE FROM tt_timetable WHERE school_id=?',[sid]);
    for(const c of cells)
      await cq('INSERT INTO tt_timetable(class_id,day_of_week,period_index,subject_id,teacher_id,room_id,school_id,week_index) VALUES(?,?,?,?,?,?,?,?)',
        [c.class_id,c.day_of_week,c.period_index,c.subject_id,c.teacher_id,c.room_id,sid,c.week_index||0]);
  });
  let hoursRestored=false;
  const cfg=body.config;
  if(cfg){ try{ const ks=HOURS_FIELDS.filter(k=>k in cfg); if(ks.length){ await run(`UPDATE tt_config SET ${ks.map(k=>k+'=?').join(',')} WHERE school_id=?`,[...ks.map(k=>cfg[k]), sid]); hoursRestored=true; } }catch(e){} }
  await loadConfig(sid);
  res.json({ok:true, cells:cells.length, hoursRestored});
}));
// restore the live timetable from an Aumtara-exported .xlsx (reads its hidden _aumtara / _aumtara_cfg sheets)
app.post('/api/timetable/restore-xlsx', upload.single('file'), async (req,res)=>{
  try{
    const sid=req.sid;
    if(!req.file){ res.status(400).json({error:'no file'}); return; }
    const wb=new ExcelJS.Workbook(); await wb.xlsx.load(req.file.buffer);
    const dws=wb.getWorksheet('_aumtara');
    if(!dws){ res.status(400).json({error:'This Excel is not an Aumtara backup — no restore data inside.'}); return; }
    const num=v=>{ if(v==null||v==='')return null; if(typeof v==='object'){ if(v.result!=null)return Number(v.result); if(v.text!=null)return Number(v.text); return null; } const n=Number(v); return isNaN(n)?null:n; };
    const cells=[];
    dws.eachRow((row,n)=>{ if(n===1)return; const cid=num(row.getCell(1).value); if(cid==null)return;
      cells.push({class_id:cid,day_of_week:num(row.getCell(2).value),period_index:num(row.getCell(3).value),subject_id:num(row.getCell(4).value),teacher_id:num(row.getCell(5).value),room_id:num(row.getCell(6).value),week_index:num(row.getCell(7).value)||0}); });
    const cfg={}; const cws=wb.getWorksheet('_aumtara_cfg');
    if(cws) cws.eachRow((row)=>{ let k=row.getCell(1).value; const v=row.getCell(2).value; if(k==null)return; k=String(k); if(k==='__format')return;
      cfg[k]=(v==null?'':(typeof v==='object'&&v.text!=null?v.text:String(v))); });
    await tx(async (cq)=>{
      await cq('DELETE FROM tt_timetable WHERE school_id=?',[sid]);
      for(const c of cells)
        await cq('INSERT INTO tt_timetable(class_id,day_of_week,period_index,subject_id,teacher_id,room_id,school_id,week_index) VALUES(?,?,?,?,?,?,?,?)',
          [c.class_id,c.day_of_week,c.period_index,c.subject_id,c.teacher_id,c.room_id,sid,c.week_index||0]);
    });
    let hoursRestored=false;
    try{ const ks=HOURS_FIELDS.filter(k=>k in cfg && cfg[k]!==''); if(ks.length){ await run(`UPDATE tt_config SET ${ks.map(k=>k+'=?').join(',')} WHERE school_id=?`,[...ks.map(k=>cfg[k]), sid]); hoursRestored=true; } }catch(e){}
    await loadConfig(sid);
    res.json({ok:true, cells:cells.length, hoursRestored});
  }catch(e){ res.status(400).json({error:'Could not read that Excel file.'}); }
});

// ---------- TEACHER schedule ----------
app.get('/api/timetable/teacher/:id', h(async (req,res)=>{
  const id=Number(req.params.id);
  const rows=await q(`SELECT tt.*, c.name cls, s.name subj FROM tt_timetable tt
     JOIN tt_class c ON c.id=tt.class_id JOIN tt_subject s ON s.id=tt.subject_id WHERE tt.teacher_id=? AND tt.week_index=${curCycleIdx(req.sid)} AND tt.school_id=?`,[id, req.sid]);
  res.json({periods:rows, weekly_load:rows.length});
}));

// ---------- ABSENCE / SUBSTITUTION ----------
app.get('/api/timetable/absence', h(async (req,res)=>res.json(await q('SELECT * FROM tt_absence WHERE school_id=?',[req.sid]))));
app.post('/api/timetable/absence', h(async (req,res)=>{
  const {teacher_id,day,reason}=req.body;
  if(day===''||day==null) await run('DELETE FROM tt_absence WHERE teacher_id=? AND school_id=?',[teacher_id, req.sid]);
  else await run(`INSERT INTO tt_absence(teacher_id,day_of_week,reason,school_id) VALUES(?,?,?,?)
     ON CONFLICT(teacher_id,day_of_week) DO UPDATE SET reason=excluded.reason`,[teacher_id,day,reason||null, req.sid]);
  res.json({ok:true});
}));
// periods that need cover today (their teacher is absent) + who's free
app.get('/api/timetable/cover', h(async (req,res)=>{
  const day=Number(req.query.day||0), sid=req.sid;
  const absentSet=new Set((await q('SELECT teacher_id FROM tt_absence WHERE day_of_week=? AND school_id=?',[day,sid])).map(r=>r.teacher_id));
  const rows=await q(`SELECT tt.class_id,tt.period_index,tt.subject_id,tt.teacher_id,c.name cls,s.name subj
     FROM tt_timetable tt JOIN tt_class c ON c.id=tt.class_id JOIN tt_subject s ON s.id=tt.subject_id
     WHERE tt.day_of_week=? AND tt.teacher_id IS NOT NULL AND tt.week_index=${curCycleIdx(sid)} AND tt.school_id=?`,[day,sid]);
  const allTeachers=await q('SELECT id,name,can_substitute FROM tt_teacher WHERE school_id=?',[sid]);
  const need=[];
  for(const r of rows){
    if(!absentSet.has(r.teacher_id)) continue;
    const sub=await q1('SELECT proxy_teacher_id FROM tt_substitution WHERE day_of_week=? AND class_id=? AND period_index=?',[day,r.class_id,r.period_index]);
    const busy=new Set((await q(`SELECT teacher_id FROM tt_timetable WHERE day_of_week=? AND period_index=? AND teacher_id IS NOT NULL AND week_index=${curCycleIdx(sid)} AND school_id=?`,[day,r.period_index,sid])).map(x=>x.teacher_id));
    const free=allTeachers.filter(t=>!busy.has(t.id)&&!absentSet.has(t.id)&&+t.can_substitute!==0);
    need.push({...r, proxy_teacher_id: sub?sub.proxy_teacher_id:null, free});
  }
  res.json({day, need});
}));
app.post('/api/timetable/substitute', h(async (req,res)=>{
  const {day,class_id,period,proxy_teacher_id}=req.body;
  if(proxy_teacher_id==null||proxy_teacher_id==='') await run('DELETE FROM tt_substitution WHERE day_of_week=? AND class_id=? AND period_index=? AND school_id=?',[day,class_id,period, req.sid]);
  else await run(`INSERT INTO tt_substitution(day_of_week,class_id,period_index,proxy_teacher_id,school_id) VALUES(?,?,?,?,?)
     ON CONFLICT(day_of_week,class_id,period_index) DO UPDATE SET proxy_teacher_id=excluded.proxy_teacher_id`,[day,class_id,period,proxy_teacher_id, req.sid]);
  res.json({ok:true});
}));

// ---------- PUSH NOTIFICATIONS (web-push) ----------
app.get('/api/push/key', h(async (req,res)=>{ res.json({key: VAPID_PUB, enabled: !!(webpush&&VAPID_PUB)}); }));
app.post('/api/push/subscribe', h(async (req,res)=>{
  const s=req.body||{}; if(!s.endpoint||!s.keys){ res.status(400).json({error:'bad subscription'}); return; }
  await run(`INSERT INTO tt_push_sub(school_id,user_id,endpoint,p256dh,auth,created_at) VALUES(?,?,?,?,?,now()::text)
     ON CONFLICT(endpoint) DO UPDATE SET user_id=excluded.user_id,school_id=excluded.school_id,p256dh=excluded.p256dh,auth=excluded.auth`,
    [req.sid, req.user.id, s.endpoint, s.keys.p256dh||null, s.keys.auth||null]);
  res.json({ok:true});
}));
app.post('/api/push/unsubscribe', h(async (req,res)=>{ const ep=(req.body||{}).endpoint; if(ep) await run('DELETE FROM tt_push_sub WHERE endpoint=? AND user_id=?',[ep, req.user.id]); res.json({ok:true}); }));
app.post('/api/push/test', h(async (req,res)=>{
  if(!webpush) return res.status(400).json({error:'push not available'});
  const subs=await q('SELECT endpoint,p256dh,auth FROM tt_push_sub WHERE user_id=?',[req.user.id]);
  if(!subs.length) return res.status(400).json({error:'no subscription — enable notifications first'});
  const r=await sendToSubs(subs, {title:'Aumtara ✓', body:'Test notification — push is working!', url:'/'});
  res.json({ok:true, ...r});
}));
app.post('/api/push/broadcast', h(async (req,res)=>{
  if(!webpush) return res.status(400).json({error:'push not available'});
  if(!['admin','master','principal'].includes(req.user.role)) return res.status(403).json({error:'not allowed'});
  const title=((req.body||{}).title||'Aumtara').toString().slice(0,80);
  const body=((req.body||{}).body||'').toString().slice(0,240);
  if(!body) return res.status(400).json({error:'message required'});
  const subs=await q('SELECT endpoint,p256dh,auth FROM tt_push_sub WHERE school_id=?',[req.sid]);
  const r=await sendToSubs(subs, {title, body, url:'/'});
  res.json({ok:true, ...r});
}));

// ---------- DATE-BASED SUBSTITUTION (proxy by calendar date + range) ----------
function dowFromISO(s){ const p=String(s).split('-').map(Number); if(p.length<3||p.some(isNaN)) return 0; const dt=new Date(Date.UTC(p[0],p[1]-1,p[2])); return (dt.getUTCDay()+6)%7; }
function datesBetween(from,to){ const out=[]; const a=new Date(from+'T00:00:00Z'), b=new Date(to+'T00:00:00Z'); let g=0; for(let d=a; d<=b && g<62; d.setUTCDate(d.getUTCDate()+1),g++) out.push(d.toISOString().slice(0,10)); return out; }
// list the absent teacher's periods across a date range, with free-teacher candidates + any saved assignment
app.get('/api/datesub/fetch', h(async (req,res)=>{
  const sid=req.sid, teacher_id=Number(req.query.teacher_id), from=req.query.from, to=req.query.to;
  if(!teacher_id||!from||!to){ res.status(400).json({error:'Select a teacher and a date range.'}); return; }
  if(to<from){ res.status(400).json({error:'End date is before start date.'}); return; }
  const dates=datesBetween(from,to); if(dates.length>31){ res.status(400).json({error:'Date range too large (max 31 days).'}); return; }
  const wdays=new Set(workingDaysArr(sid));
  const allTeachers=await q('SELECT id,name,can_substitute,main_subject_id FROM tt_teacher WHERE school_id=? ORDER BY name',[sid]);
  // subject-match ranking data: teacher → {subjectId: seq} + main subject → tier a candidate for a given period's subject
  const tSub={}; (await q('SELECT ts.teacher_id,ts.subject_id,ts.seq FROM tt_teacher_subject ts JOIN tt_teacher t ON t.id=ts.teacher_id WHERE t.school_id=?',[sid])).forEach(r=>{ (tSub[r.teacher_id]=tSub[r.teacher_id]||{})[r.subject_id]=(r.seq==null?99:r.seq); });
  const mainOf={}; allTeachers.forEach(t=>mainOf[t.id]=t.main_subject_id);
  const subTier=(tid,subjId)=>{ if(subjId==null) return 4; const sm=tSub[tid]; const seq=(sm&&sm[subjId]!=null)?sm[subjId]:null;
    if(mainOf[tid]===subjId || seq===0) return 1;   // primary (main / priority-1)
    if(seq!=null && seq<=2) return 2;               // secondary (priority 2-3)
    if(seq!=null) return 3;                          // backup/proxy (teaches it, priority 4+)
    return 4; };                                    // general (can-substitute, doesn't teach it)
  const cells=await q(`SELECT day_of_week,period_index,teacher_id FROM tt_timetable WHERE teacher_id IS NOT NULL AND week_index=${curCycleIdx(sid)} AND school_id=?`,[sid]);
  const busy={}; cells.forEach(c=>{ const k=c.day_of_week+'_'+c.period_index; (busy[k]=busy[k]||new Set()).add(c.teacher_id); });
  const myCells=await q(`SELECT tt.day_of_week,tt.period_index,tt.class_id,tt.subject_id,c.name cls,s.name subj
     FROM tt_timetable tt JOIN tt_class c ON c.id=tt.class_id LEFT JOIN tt_subject s ON s.id=tt.subject_id
     WHERE tt.teacher_id=? AND tt.school_id=?`,[teacher_id,sid]);
  const byDow={}; myCells.forEach(c=>{ (byDow[c.day_of_week]=byDow[c.day_of_week]||[]).push(c); });
  const ex=await q('SELECT sub_date,class_id,period_index,proxy_teacher_id,is_free FROM tt_datesub WHERE school_id=? AND absent_teacher_id=? AND sub_date>=? AND sub_date<=?',[sid,teacher_id,from,to]);
  const exMap={}; ex.forEach(e=>exMap[e.sub_date+'_'+e.class_id+'_'+e.period_index]=e);
  const out=[];
  for(const date of dates){ const dow=dowFromISO(date); if(!wdays.has(dow)) continue; const list=byDow[dow]||[]; if(!list.length) continue;
    const slots=teachingSlots(dow,sid);
    for(const c of list){ const busySet=busy[dow+'_'+c.period_index]||new Set();
      const free=allTeachers.filter(t=>t.id!==teacher_id && !busySet.has(t.id) && +t.can_substitute!==0)
        .map(f=>({id:f.id,name:f.name,tier:subTier(f.id,c.subject_id)}))
        .sort((a,b)=>a.tier-b.tier || String(a.name||'').localeCompare(String(b.name||'')));   // best subject match first
      const e=exMap[date+'_'+c.class_id+'_'+c.period_index]; const sl=slots[c.period_index];
      out.push({ sub_date:date, dow, period_index:c.period_index, class_id:c.class_id, cls:c.cls, subject_id:c.subject_id, subj:c.subj||'',
        start:sl?sl.start:'', end:sl?sl.end:'', free,
        proxy_teacher_id: e?e.proxy_teacher_id:null, is_free: e?+e.is_free:0 }); }
  }
  res.json({periods:out});
}));
// save/clear one date-based assignment (proxy teacher, or mark free, or clear)
app.post('/api/datesub', h(async (req,res)=>{
  const sid=req.sid, b=req.body; const sub_date=b.sub_date, class_id=b.class_id, period_index=b.period_index;
  if(!sub_date||class_id==null||period_index==null){ res.status(400).json({error:'missing fields'}); return; }
  const is_free=b.is_free?1:0; const proxy=(!is_free && b.proxy_teacher_id)?Number(b.proxy_teacher_id):null;
  if(!is_free && !proxy){ await run('DELETE FROM tt_datesub WHERE school_id=? AND sub_date=? AND class_id=? AND period_index=?',[sid,sub_date,class_id,period_index]); res.json({ok:true,cleared:true}); return; }
  await run(`INSERT INTO tt_datesub(school_id,sub_date,class_id,period_index,absent_teacher_id,proxy_teacher_id,is_free,created_at)
     VALUES(?,?,?,?,?,?,?,now()::text)
     ON CONFLICT(school_id,sub_date,class_id,period_index) DO UPDATE SET absent_teacher_id=excluded.absent_teacher_id,proxy_teacher_id=excluded.proxy_teacher_id,is_free=excluded.is_free`,
     [sid,sub_date,class_id,period_index,b.absent_teacher_id||null,proxy,is_free]);
  res.json({ok:true});
}));

// ---------- LEAVE APPLICATIONS (apply → approve → auto proxy cover) ----------
async function myTeacherId(req){ if(!req.user) return null; const t=await q1('SELECT id FROM tt_teacher WHERE lower(name)=lower(?) AND school_id=? LIMIT 1',[req.user.name, req.sid]); return t?t.id:null; }
const isAdminRole = req => req.user && ['master','admin','principal','supervisor'].includes(req.user.role);
app.get('/api/leaves', h(async (req,res)=>{
  let rows;
  if(req.user && req.user.role==='teacher'){ const tid=await myTeacherId(req); rows=tid?await q('SELECT * FROM tt_leave WHERE school_id=? AND teacher_id=? ORDER BY id DESC',[req.sid,tid]):[]; }
  else rows=await q('SELECT * FROM tt_leave WHERE school_id=? ORDER BY id DESC',[req.sid]);
  const tmap={}; (await q('SELECT id,name FROM tt_teacher WHERE school_id=?',[req.sid])).forEach(t=>tmap[t.id]=t.name);
  res.json(rows.map(r=>({...r, teacher_name: tmap[r.teacher_id]||''})));
}));
app.post('/api/leaves', h(async (req,res)=>{
  const b=req.body; let teacher_id=b.teacher_id;
  if(req.user && req.user.role==='teacher'){ teacher_id=await myTeacherId(req); if(!teacher_id){ res.status(400).json({error:'Your teacher profile was not found — your user name must match a teacher name.'}); return; } }
  if(!teacher_id||!b.date_from||!b.date_to){ res.status(400).json({error:'Teacher and both dates are required.'}); return; }
  if(b.date_to<b.date_from){ res.status(400).json({error:'End date is before start date.'}); return; }
  const r=await q1('INSERT INTO tt_leave(school_id,teacher_id,date_from,date_to,reason,leave_type,status,created_at) VALUES(?,?,?,?,?,?,?,now()::text) RETURNING id',[req.sid,teacher_id,b.date_from,b.date_to,b.reason||null,b.leave_type||null,'pending']);
  res.json({ok:true,id:r.id});
}));
// ---------- LEAVE TYPES (config list) ----------
app.get('/api/leave-types', h(async (req,res)=>{ res.json(await q('SELECT * FROM tt_leave_type WHERE school_id=? AND active=1 ORDER BY id',[req.sid])); }));
app.post('/api/leave-types', h(async (req,res)=>{
  const name=(req.body.name||'').trim(); if(!name){ res.status(400).json({error:'name required'}); return; }
  const r=await q1('INSERT INTO tt_leave_type(school_id,name,active,created_at) VALUES(?,?,1,now()::text) RETURNING id',[req.sid,name]);
  res.json({ok:true,id:r.id});
}));
app.delete('/api/leave-types/:id', h(async (req,res)=>{ await run('DELETE FROM tt_leave_type WHERE id=? AND school_id=?',[req.params.id,req.sid]); res.json({ok:true}); }));
app.post('/api/leaves/:id/approve', h(async (req,res)=>{
  if(!isAdminRole(req)){ res.status(403).json({error:'forbidden'}); return; }
  const id=req.params.id, sid=req.sid;
  const lv=await q1('SELECT * FROM tt_leave WHERE id=? AND school_id=?',[id,sid]); if(!lv){ res.status(404).json({error:'not found'}); return; }
  await run('UPDATE tt_leave SET status=? WHERE id=? AND school_id=?',['approved',id,sid]);
  const wdays=new Set(workingDaysArr(sid));
  const myCells=await q(`SELECT day_of_week,period_index,class_id FROM tt_timetable WHERE teacher_id=? AND week_index=${curCycleIdx(sid)} AND school_id=?`,[lv.teacher_id,sid]);
  const byDow={}; myCells.forEach(c=>{ (byDow[c.day_of_week]=byDow[c.day_of_week]||[]).push(c); });
  let created=0;
  for(const date of datesBetween(lv.date_from,lv.date_to)){ const dow=dowFromISO(date); if(!wdays.has(dow)) continue;
    for(const c of (byDow[dow]||[])){ await run(`INSERT INTO tt_datesub(school_id,sub_date,class_id,period_index,absent_teacher_id,proxy_teacher_id,is_free,created_at)
       VALUES(?,?,?,?,?,?,0,now()::text) ON CONFLICT(school_id,sub_date,class_id,period_index) DO NOTHING`,[sid,date,c.class_id,c.period_index,lv.teacher_id,null]); created++; } }
  res.json({ok:true, periods:created, teacher_id:lv.teacher_id, date_from:lv.date_from, date_to:lv.date_to});
}));
app.post('/api/leaves/:id/reject', h(async (req,res)=>{ if(!isAdminRole(req)){ res.status(403).json({error:'forbidden'}); return; } await run('UPDATE tt_leave SET status=? WHERE id=? AND school_id=?',['rejected',req.params.id,req.sid]); res.json({ok:true}); }));
app.delete('/api/leaves/:id', h(async (req,res)=>{ if(!isAdminRole(req)){ res.status(403).json({error:'forbidden'}); return; } await run('DELETE FROM tt_leave WHERE id=? AND school_id=?',[req.params.id,req.sid]); res.json({ok:true}); }));

// ---------- LIVE MONITOR ----------
app.get('/api/timetable/monitor', h(async (req,res)=>{
  let day,hhmm,dateStr=null; const p2=n=>String(n).padStart(2,'0');
  if(req.query.day!=null&&req.query.time){day=Number(req.query.day);hhmm=req.query.time;dateStr=req.query.date||null;}
  else{const d=new Date();day=(d.getDay()+6)%7;if(day>5)day=5;hhmm=p2(d.getHours())+':'+p2(d.getMinutes());dateStr=d.getFullYear()+'-'+p2(d.getMonth()+1)+'-'+p2(d.getDate());}
  if(req.query.day==null && rotMode(req.sid)) day=todayRotDay(req.sid);   // rotation mode: "today" is a rotation-day index
  const sid=req.sid;
  const absent={}; (await q('SELECT * FROM tt_absence WHERE school_id=?',[sid])).forEach(a=>{(absent[a.teacher_id]=absent[a.teacher_id]||new Set()).add(a.day_of_week);});
  // date-based subs for the actual date take precedence over recurring weekday subs
  const dsMap={};
  if(dateStr){ (await q('SELECT class_id,period_index,proxy_teacher_id,is_free FROM tt_datesub WHERE school_id=? AND sub_date=?',[sid,dateStr])).forEach(r=>{ dsMap[r.class_id+'_'+r.period_index]=r; }); }
  const slots=slotsForDay(day, sid);
  const classes=await q('SELECT * FROM tt_class WHERE school_id=? ORDER BY id',[sid]);
  const out=[];
  for(const c of classes){
    let pi=0,status='idle',text='Outside school hours';
    for(const sl of slots){
      if(hhmm>=sl.start&&hhmm<sl.end){
        if(sl.is_break){status='break';text='Break';break;}
        const cell=await q1(`SELECT * FROM tt_timetable WHERE class_id=? AND day_of_week=? AND period_index=? AND week_index=${curCycleIdx(sid)} AND school_id=?`,[c.id,day,pi,sid]);
        if(cell&&cell.subject_id){
          const subj=(await q1('SELECT name FROM tt_subject WHERE id=?',[cell.subject_id])).name;
          const ds=dsMap[c.id+'_'+pi];
          if(ds){
            if(+ds.is_free===1){ status='idle'; text=subj+' · Free (no cover today)'; }
            else if(ds.proxy_teacher_id){ status='proxy'; text=subj+' · Proxy: '+(await q1('SELECT name FROM tt_teacher WHERE id=?',[ds.proxy_teacher_id])).name; }
            else { status='proxy'; text=subj+' · Proxy needed'; }
          } else {
            const sub=await q1('SELECT proxy_teacher_id FROM tt_substitution WHERE day_of_week=? AND class_id=? AND period_index=?',[day,c.id,pi]);
            const absentTeacher=cell.teacher_id&&absent[cell.teacher_id]&&absent[cell.teacher_id].has(day);
            if(absentTeacher && sub && sub.proxy_teacher_id){ status='proxy'; text=subj+' · Proxy: '+(await q1('SELECT name FROM tt_teacher WHERE id=?',[sub.proxy_teacher_id])).name; }
            else if(absentTeacher){ status='proxy'; text=subj+' · Proxy needed'; }
            else { const tch=cell.teacher_id?(await q1('SELECT name FROM tt_teacher WHERE id=?',[cell.teacher_id])).name:'—'; status='live'; text=subj+' · '+tch; }
          }
        } else { status='idle'; text='Free period'; }
        break;
      }
      if(!sl.is_break)pi++;
    }
    out.push({class_id:c.id,name:c.name,status,text});
  }
  res.json({day,time:hhmm,classes:out});
}));
// live "who is free right now" — teachers with no class in the CURRENT period (present, not absent). For the principal's dashboard.
app.get('/api/timetable/free-now', h(async (req,res)=>{
  let day,hhmm; const p2=n=>String(n).padStart(2,'0');
  if(req.query.day!=null&&req.query.time){ day=Number(req.query.day); hhmm=req.query.time; }
  else{ const d=new Date(); day=(d.getDay()+6)%7; if(day>5)day=5; hhmm=p2(d.getHours())+':'+p2(d.getMinutes()); }
  if(req.query.day==null && rotMode(req.sid)) day=todayRotDay(req.sid);
  const sid=req.sid;
  const slots=slotsForDay(day, sid);
  let pi=0, status='closed', period=null;
  for(const sl of slots){
    if(hhmm>=sl.start && hhmm<sl.end){ if(sl.is_break){ status='break'; } else { status='running'; period=pi+1; } break; }
    if(!sl.is_break) pi++;
  }
  const teachers=await q('SELECT id,name FROM tt_teacher WHERE school_id=? ORDER BY name',[sid]);
  const absent=new Set(); (await q('SELECT teacher_id FROM tt_absence WHERE school_id=? AND day_of_week=?',[sid,day])).forEach(a=>absent.add(a.teacher_id));
  let free=[], busyCount=0, classesNow=0, noTeacher=0;
  if(status==='running'){
    // all classes with a subject scheduled this period; a teacher may or may not be assigned
    const cells=await q(`SELECT teacher_id FROM tt_timetable WHERE school_id=? AND day_of_week=? AND period_index=? AND week_index=${curCycleIdx(sid)} AND subject_id IS NOT NULL`,[sid,day,pi]);
    classesNow=cells.length; const busy=new Set();
    cells.forEach(c=>{ if(c.teacher_id) busy.add(c.teacher_id); else noTeacher++; });   // subject scheduled but no teacher = coverage gap
    // group-split (elective) teachers are also teaching this period even though there's no tt_timetable row
    (await q(`SELECT o.teacher_id FROM tt_elective e JOIN tt_elective_option o ON o.elective_id=e.id WHERE e.school_id=? AND e.day_of_week=? AND e.period_index=? AND COALESCE(e.week_index,0)=${curCycleIdx(sid)} AND o.teacher_id IS NOT NULL`,[sid,day,pi])).forEach(r=>busy.add(r.teacher_id));
    busyCount=busy.size;
    teachers.forEach(tt=>{ if(absent.has(tt.id)) return; if(!busy.has(tt.id)) free.push({id:tt.id,name:tt.name}); });
  }
  res.json({ day, time:hhmm, status, period, free, busyCount, absentCount:absent.size, totalTeachers:teachers.length, classesNow, noTeacher });
}));

// ---------- TEACHER DIARY ----------
app.get('/api/diary', h(async (req,res)=>{
  const {teacher_id, date}=req.query;
  res.json(await q('SELECT * FROM tt_diary WHERE teacher_id=? AND entry_date=? AND school_id=?',[teacher_id, date, req.sid]));
}));
app.get('/api/diary/all', h(async (req,res)=>{
  let sql=`SELECT d.*, c.name cls, s.name subj, t.name teacher FROM tt_diary d
     LEFT JOIN tt_class c ON c.id=d.class_id LEFT JOIN tt_subject s ON s.id=d.subject_id JOIN tt_teacher t ON t.id=d.teacher_id`;
  const w=['d.school_id=?'], p=[req.sid];
  if(req.query.teacher_id){w.push('d.teacher_id=?');p.push(req.query.teacher_id);}
  if(req.query.from){w.push('d.entry_date>=?');p.push(req.query.from);}
  if(req.query.to){w.push('d.entry_date<=?');p.push(req.query.to);}
  if(w.length)sql+=' WHERE '+w.join(' AND ');
  sql+=' ORDER BY d.entry_date DESC, d.period_index LIMIT 300';
  res.json(await q(sql,p));
}));
app.post('/api/diary', h(async (req,res)=>{
  const b=req.body; const dow=b.day_of_week!=null?b.day_of_week:null;
  await run(`INSERT INTO tt_diary(teacher_id,entry_date,day_of_week,class_id,subject_id,period_index,lesson,topic,learning_outcome,assessment_lo,homework,teaching_aids,created_at,school_id)
     VALUES(?,?,?,?,?,?,?,?,?,?,?,?, now()::text,?)
     ON CONFLICT(teacher_id,entry_date,class_id,period_index) DO UPDATE SET
       subject_id=excluded.subject_id, day_of_week=excluded.day_of_week,
       lesson=excluded.lesson, topic=excluded.topic, learning_outcome=excluded.learning_outcome,
       assessment_lo=excluded.assessment_lo, homework=excluded.homework, teaching_aids=excluded.teaching_aids`,
    [b.teacher_id,b.entry_date,dow,b.class_id||null,b.subject_id||null,b.period_index??null,
     b.lesson||null,b.topic||null,b.learning_outcome||null,b.assessment_lo||null,b.homework||null,b.teaching_aids||null, req.sid]);
  res.json({ok:true});
}));
app.delete('/api/diary/:id', h(async (req,res)=>{ await run('DELETE FROM tt_diary WHERE id=? AND school_id=?',[req.params.id, req.sid]); res.json({ok:true}); }));

app.get('/api/export/diary.xlsx', h(async (req,res)=>{
  const wb=new ExcelJS.Workbook(); const ws=wb.addWorksheet("Teacher's Diary");
  const school=(getConfig(req.sid)||{}).school_name||'';
  ws.mergeCells('A1:K1'); ws.getCell('A1').value=school; ws.getCell('A1').font={bold:true,size:15,color:{argb:'FF1F3864'}}; ws.getCell('A1').alignment={horizontal:'center'};
  ws.mergeCells('A2:K2'); ws.getCell('A2').value="Teacher's Diary"; ws.getCell('A2').font={bold:true,size:12}; ws.getCell('A2').alignment={horizontal:'center'};
  ws.addRow([]);
  const hdr=ws.addRow(['Date','Day','Teacher','Class','Subject','Lesson','Topic','Learning Outcome','Assessment of LO (Strategies)','Home Work','Teaching Aids used']);
  hdr.eachCell(c=>{c.font={bold:true,color:{argb:'FFFFFFFF'}};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF1F3864'}};c.alignment={horizontal:'center',wrapText:true};c.border={top:{style:'thin'},bottom:{style:'thin'},left:{style:'thin'},right:{style:'thin'}};});
  let sql=`SELECT d.*, c.name cls, s.name subj, t.name teacher FROM tt_diary d
     LEFT JOIN tt_class c ON c.id=d.class_id LEFT JOIN tt_subject s ON s.id=d.subject_id JOIN tt_teacher t ON t.id=d.teacher_id`;
  const w=['d.school_id=?'],p=[req.sid]; if(req.query.teacher_id){w.push('d.teacher_id=?');p.push(req.query.teacher_id);} if(req.query.date){w.push('d.entry_date=?');p.push(req.query.date);} if(req.query.from){w.push('d.entry_date>=?');p.push(req.query.from);} if(req.query.to){w.push('d.entry_date<=?');p.push(req.query.to);} sql+=' WHERE '+w.join(' AND ');
  sql+=' ORDER BY d.entry_date DESC, d.period_index';
  (await q(sql,p)).forEach(d=>{ const r=ws.addRow([d.entry_date, d.day_of_week!=null?DAYS[d.day_of_week]:'', d.teacher, d.cls||'', d.subj||'', d.lesson||'', d.topic||'', d.learning_outcome||'', d.assessment_lo||'', d.homework||'', d.teaching_aids||'']); r.eachCell(c=>{c.alignment={wrapText:true,vertical:'top'};c.border={top:{style:'thin'},bottom:{style:'thin'},left:{style:'thin'},right:{style:'thin'}};}); });
  ws.columns.forEach((c,i)=>c.width=[11,6,15,10,12,8,22,26,26,20,20][i]);
  res.setHeader('Content-Disposition','attachment; filename=teacher-diary.xlsx');
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  await wb.xlsx.write(res); res.end();
}));

// ---------- EXCEL EXPORT ----------
function styleHeader(row){ row.eachCell(c=>{ c.font={bold:true,color:{argb:'FFFFFFFF'}}; c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF1F3864'}}; c.alignment={horizontal:'center',vertical:'middle'}; }); }

app.get('/api/export/timetable.xlsx', h(async (req,res)=>{
  const wb=new ExcelJS.Workbook(); wb.creator='Aumtara'; const sid=req.sid;
  const classes=await q('SELECT * FROM tt_class WHERE school_id=? ORDER BY id',[sid]);
  const subjById={}; (await q('SELECT id,name FROM tt_subject WHERE school_id=?',[sid])).forEach(s=>subjById[s.id]=s.name);
  const tchById={}; (await q('SELECT id,name FROM tt_teacher WHERE school_id=?',[sid])).forEach(t=>tchById[t.id]=t.name);
  const roomById={}; (await q('SELECT id,name FROM tt_room WHERE school_id=?',[sid])).forEach(r=>roomById[r.id]=r.name);
  const cellMap={}; (await q('SELECT * FROM tt_timetable WHERE school_id=?',[sid])).forEach(c=>cellMap[`${c.class_id}|${c.day_of_week}|${c.period_index}|${c.week_index||0}`]=c);
  const cellText=(c)=>{ if(!c||!c.subject_id)return ''; const s=subjById[c.subject_id]||''; const t=c.teacher_id?(tchById[c.teacher_id]||''):''; const r=c.room_id?(roomById[c.room_id]||''):''; return s+(t?'\n'+t:'')+(r?'\n'+r:''); };
  const dayList=workingDaysArr(sid);   // rotation mode → rotation days; else working weekdays
  const dayLbl = d => rotMode(sid) ? rotLabel(sid,d) : DAYS[d];
  const maxSlots=Math.max(1,...dayList.map(d=>teachingSlots(d, sid).length));
  const cfg=getConfig(sid)||{}; const cyc = rotMode(sid) ? 1 : Math.max(1,Math.min(4,parseInt(cfg.cycle_weeks,10)||1));
  const wkLabels=(cfg.week_labels||'').split(',').map(x=>x.trim()).filter(Boolean);
  const wkName=i=> wkLabels[i] || (cyc===2?['Week A','Week B'][i] : 'Week '+(i+1));
  const hdr=['Day',...Array.from({length:maxSlots},(_,i)=>'P'+(i+1))];
  classes.forEach(c=>{
    const ws=wb.addWorksheet(c.name.slice(0,28).replace(/[\\\/\?\*\[\]:]/g,' '));
    for(let wk=0; wk<cyc; wk++){
      if(cyc>1){ const tr=ws.addRow([wkName(wk)]); tr.getCell(1).font={bold:true,color:{argb:'FF1F3864'},size:12}; }
      ws.addRow(hdr); styleHeader(ws.getRow(ws.rowCount));
      dayList.forEach(di=>{ const row=[dayLbl(di)]; for(let p=0;p<maxSlots;p++){ row.push(cellText(cellMap[`${c.id}|${di}|${p}|${wk}`])); }
        const r=ws.addRow(row); r.eachCell(cc=>{cc.alignment={wrapText:true,vertical:'top'};}); r.getCell(1).font={bold:true}; });
      if(cyc>1 && wk<cyc-1) ws.addRow([]);   // spacer between weeks
    }
    ws.columns.forEach((col,i)=>{col.width=i===0?12:16;});
  });
  // hidden machine-readable sheets → this same .xlsx can be re-imported to RESTORE the timetable
  try{
    const dws=wb.addWorksheet('_aumtara'); dws.state='veryHidden';
    dws.addRow(['class_id','day_of_week','period_index','subject_id','teacher_id','room_id','week_index']);
    (await q('SELECT class_id,day_of_week,period_index,subject_id,teacher_id,room_id,week_index FROM tt_timetable WHERE school_id=?',[sid])).forEach(c=>dws.addRow([c.class_id,c.day_of_week,c.period_index,c.subject_id,c.teacher_id,c.room_id,c.week_index||0]));
    const cws=wb.addWorksheet('_aumtara_cfg'); cws.state='veryHidden';
    cws.addRow(['__format','aumtara-timetable-backup']);
    HOURS_FIELDS.forEach(k=>cws.addRow([k, cfg[k]==null?'':String(cfg[k])]));
  }catch(e){}
  res.setHeader('Content-Disposition','attachment; filename=timetable.xlsx');
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  await wb.xlsx.write(res); res.end();
}));

// TEACHER-WISE report: one sheet per teacher = their weekly grid (day × period → class + subject) + total load
app.get('/api/export/teachers-report.xlsx', h(async (req,res)=>{
  const wb=new ExcelJS.Workbook(); wb.creator='Aumtara'; const sid=req.sid;
  const teachers=await q('SELECT * FROM tt_teacher WHERE school_id=? ORDER BY name,id',[sid]);
  const clsById={}; (await q('SELECT id,name FROM tt_class WHERE school_id=?',[sid])).forEach(c=>clsById[c.id]=c.name);
  const subjById={}; (await q('SELECT id,name FROM tt_subject WHERE school_id=?',[sid])).forEach(s=>subjById[s.id]=s.name);
  const roomById={}; (await q('SELECT id,name FROM tt_room WHERE school_id=?',[sid])).forEach(r=>roomById[r.id]=r.name);
  const cells=await q('SELECT * FROM tt_timetable WHERE school_id=? AND week_index=0',[sid]);
  const dayList=workingDaysArr(sid);
  const dayLbl = d => rotMode(sid) ? rotLabel(sid,d) : DAYS[d];
  const maxSlots=Math.max(1,...dayList.map(d=>teachingSlots(d,sid).length));
  const hdr=['Day/Period',...Array.from({length:maxSlots},(_,i)=>'P'+(i+1))];
  teachers.forEach(tc=>{
    const ws=wb.addWorksheet((tc.name||('T'+tc.id)).slice(0,28).replace(/[\\\/\?\*\[\]:]/g,' '));
    ws.addRow(['Teacher: '+(tc.name||'')]).getCell(1).font={bold:true,size:13,color:{argb:'FF1F3864'}};
    ws.addRow(hdr); styleHeader(ws.getRow(ws.rowCount));
    let load=0;
    dayList.forEach(di=>{ const slots=teachingSlots(di,sid); const row=[dayLbl(di)];
      for(let p=0;p<maxSlots;p++){ if(p>=slots.length){ row.push(''); continue; }
        const c=cells.find(x=>x.teacher_id===tc.id && x.day_of_week===di && x.period_index===p);
        if(c){ load++; row.push((clsById[c.class_id]||'')+'\n'+(subjById[c.subject_id]||'')+(c.room_id?'\n'+(roomById[c.room_id]||''):'')); }
        else row.push(''); }
      const r=ws.addRow(row); r.eachCell(cc=>{cc.alignment={wrapText:true,vertical:'top'};}); r.getCell(1).font={bold:true}; });
    const lr=ws.addRow(['Total periods: '+load+(tc.max_load?(' / '+tc.max_load+' max'):'')]); lr.getCell(1).font={bold:true,color:{argb:'FF7030A0'}};
    ws.columns.forEach((col,i)=>{col.width=i===0?14:18;});
  });
  if(!teachers.length){ const ws=wb.addWorksheet('Teachers'); ws.addRow(['No teachers yet']); }
  res.setHeader('Content-Disposition','attachment; filename=teachers-report.xlsx');
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  await wb.xlsx.write(res); res.end();
}));

// CLASS-WISE workload summary: per-class filled/empty + a Class×Subject scheduled-vs-quota table
app.get('/api/export/class-summary.xlsx', h(async (req,res)=>{
  const wb=new ExcelJS.Workbook(); wb.creator='Aumtara'; const sid=req.sid;
  const classes=await q('SELECT * FROM tt_class WHERE school_id=? ORDER BY id',[sid]);
  const subjById={}; (await q('SELECT id,name FROM tt_subject WHERE school_id=?',[sid])).forEach(s=>subjById[s.id]=s.name);
  const cells=await q('SELECT * FROM tt_timetable WHERE school_id=? AND week_index=0',[sid]);
  const quotas=await q('SELECT * FROM tt_quota WHERE school_id=?',[sid]);
  const dayList=workingDaysArr(sid);
  let perClassSlots=0; dayList.forEach(di=>perClassSlots+=teachingSlots(di,sid).length);
  const s1=wb.addWorksheet('Class Summary');
  s1.addRow(['Class','Total slots','Filled','Empty','No teacher','Subjects','Teachers']); styleHeader(s1.getRow(1));
  classes.forEach(c=>{ const cc=cells.filter(x=>x.class_id===c.id && x.subject_id);
    const subs=new Set(cc.map(x=>x.subject_id)), tch=new Set(cc.filter(x=>x.teacher_id).map(x=>x.teacher_id));
    const noT=cc.filter(x=>!x.teacher_id).length;
    s1.addRow([c.name, perClassSlots, cc.length, perClassSlots-cc.length, noT, subs.size, tch.size]); });
  s1.columns.forEach((col,i)=>col.width=[22,12,10,10,12,12,12][i]);
  const s2=wb.addWorksheet('Class x Subject');
  s2.addRow(['Class','Subject','Scheduled/wk','Quota/wk','Difference']); styleHeader(s2.getRow(1));
  classes.forEach(c=>{ const cc=cells.filter(x=>x.class_id===c.id&&x.subject_id);
    const cnt={}; cc.forEach(x=>cnt[x.subject_id]=(cnt[x.subject_id]||0)+1);
    const qs=quotas.filter(x=>x.class_id===c.id); const seen=new Set();
    qs.forEach(qq=>{ seen.add(qq.subject_id); const sch=cnt[qq.subject_id]||0; s2.addRow([c.name, subjById[qq.subject_id]||'', sch, qq.per_week, sch-qq.per_week]); });
    Object.keys(cnt).forEach(k=>{ if(!seen.has(+k)) s2.addRow([c.name, subjById[k]||'', cnt[k], 0, cnt[k]]); }); });
  s2.columns.forEach((col,i)=>col.width=[22,20,14,12,12][i]);
  res.setHeader('Content-Disposition','attachment; filename=class-summary.xlsx');
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  await wb.xlsx.write(res); res.end();
}));

// SUBJECT-WISE coverage: scheduled vs quota across all classes, with coverage %
app.get('/api/export/subject-coverage.xlsx', h(async (req,res)=>{
  const wb=new ExcelJS.Workbook(); wb.creator='Aumtara'; const sid=req.sid;
  const subjects=await q('SELECT * FROM tt_subject WHERE school_id=? ORDER BY name,id',[sid]);
  const cells=await q('SELECT * FROM tt_timetable WHERE school_id=? AND week_index=0 AND subject_id IS NOT NULL',[sid]);
  const quotas=await q('SELECT * FROM tt_quota WHERE school_id=?',[sid]);
  const tchById={}; (await q('SELECT id,name FROM tt_teacher WHERE school_id=?',[sid])).forEach(t=>tchById[t.id]=t.name);
  const ws=wb.addWorksheet('Subject Coverage');
  ws.addRow(['Subject','Quota/wk (all classes)','Scheduled/wk','Coverage %','Classes','Teachers']); styleHeader(ws.getRow(1));
  subjects.forEach(su=>{ const sc=cells.filter(x=>x.subject_id===su.id);
    const quota=quotas.filter(x=>x.subject_id===su.id).reduce((a,b)=>a+(+b.per_week||0),0);
    const sched=sc.length; const cov=quota>0?Math.round(sched/quota*100):(sched>0?100:0);
    const cls=new Set(sc.map(x=>x.class_id)).size;
    const tnames=[...new Set(sc.filter(x=>x.teacher_id).map(x=>tchById[x.teacher_id]||''))].filter(Boolean).join(', ');
    const r=ws.addRow([su.name, quota, sched, cov+'%', cls, tnames]);
    if(quota>0 && sched<quota) r.getCell(4).font={color:{argb:'FFB91C1C'},bold:true};   // under-covered → red
  });
  ws.columns.forEach((col,i)=>col.width=[20,22,14,12,10,36][i]);
  res.setHeader('Content-Disposition','attachment; filename=subject-coverage.xlsx');
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  await wb.xlsx.write(res); res.end();
}));

// ATTENDANCE-STYLE diary register: teachers × dates, cell = entries written that day (compliance view)
app.get('/api/export/diary-summary.xlsx', h(async (req,res)=>{
  const wb=new ExcelJS.Workbook(); wb.creator='Aumtara'; const sid=req.sid;
  const from=req.query.from, to=req.query.to;
  const teachers=await q('SELECT id,name FROM tt_teacher WHERE school_id=? ORDER BY name,id',[sid]);
  const w=['school_id=?'],p=[sid]; if(from){w.push('entry_date>=?');p.push(from);} if(to){w.push('entry_date<=?');p.push(to);}
  const rows=await q('SELECT teacher_id,entry_date,COUNT(*)::int n FROM tt_diary WHERE '+w.join(' AND ')+' GROUP BY teacher_id,entry_date',p);
  let cols=[...new Set(rows.map(r=>r.entry_date))].sort();
  if(from&&to){ cols=[]; let d=new Date(from+'T00:00:00'); const end=new Date(to+'T00:00:00'); for(let i=0;i<45 && d<=end;i++){ cols.push(d.toISOString().slice(0,10)); d.setDate(d.getDate()+1); } }
  const map={}; rows.forEach(r=>{ map[r.teacher_id+'|'+r.entry_date]=r.n; });
  const ws=wb.addWorksheet('Diary Register');
  ws.addRow(['Teacher',...cols.map(d=>d.slice(8)+'/'+d.slice(5,7)),'Total','Days']); styleHeader(ws.getRow(1));
  teachers.forEach(t=>{ let tot=0,days=0; const row=[t.name];
    cols.forEach(d=>{ const n=map[t.id+'|'+d]||0; tot+=n; if(n)days++; row.push(n||''); });
    row.push(tot); row.push(days); ws.addRow(row); });
  ws.columns.forEach((col,i)=>col.width=i===0?20:6);
  res.setHeader('Content-Disposition','attachment; filename=diary-register.xlsx');
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  await wb.xlsx.write(res); res.end();
}));

// ROOM UTILISATION: how much each room is used vs the total slots in the week
app.get('/api/export/room-utilization.xlsx', h(async (req,res)=>{
  const wb=new ExcelJS.Workbook(); wb.creator='Aumtara'; const sid=req.sid;
  const rooms=await q('SELECT * FROM tt_room WHERE school_id=? ORDER BY name,id',[sid]);
  const clsById={}; (await q('SELECT id,name FROM tt_class WHERE school_id=?',[sid])).forEach(c=>clsById[c.id]=c.name);
  const cells=await q('SELECT * FROM tt_timetable WHERE school_id=? AND week_index=0 AND room_id IS NOT NULL',[sid]);
  const dayList=workingDaysArr(sid);
  let totalSlots=0; dayList.forEach(di=>totalSlots+=teachingSlots(di,sid).length);
  const ws=wb.addWorksheet('Room Utilisation');
  ws.addRow(['Room','Capacity','Used periods','Total slots','Utilisation %','Classes using']); styleHeader(ws.getRow(1));
  rooms.forEach(rm=>{ const rc=cells.filter(x=>x.room_id===rm.id);
    const util=totalSlots>0?Math.round(rc.length/totalSlots*100):0;
    const cls=[...new Set(rc.map(x=>clsById[x.class_id]||''))].filter(Boolean).join(', ');
    const r=ws.addRow([rm.name, rm.capacity||'', rc.length, totalSlots, util+'%', cls]);
    if(util<40) r.getCell(5).font={color:{argb:'FF166534'}};           // very free room → green
    else if(util>90) r.getCell(5).font={color:{argb:'FFB91C1C'},bold:true};  // nearly full → red
  });
  ws.columns.forEach((col,i)=>col.width=[20,10,13,11,14,40][i]);
  res.setHeader('Content-Disposition','attachment; filename=room-utilization.xlsx');
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  await wb.xlsx.write(res); res.end();
}));

// TEACHER FREE-PERIOD finder: a teacher × (day-period) matrix — blank = FREE, else the class taught + a Free count
app.get('/api/export/teacher-free.xlsx', h(async (req,res)=>{
  const wb=new ExcelJS.Workbook(); wb.creator='Aumtara'; const sid=req.sid;
  const teachers=await q('SELECT * FROM tt_teacher WHERE school_id=? ORDER BY name,id',[sid]);
  const clsById={}; (await q('SELECT id,name FROM tt_class WHERE school_id=?',[sid])).forEach(c=>clsById[c.id]=c.name);
  const cells=await q('SELECT * FROM tt_timetable WHERE school_id=? AND week_index=0 AND teacher_id IS NOT NULL',[sid]);
  const dayList=workingDaysArr(sid);
  const dayLbl = d => rotMode(sid) ? rotLabel(sid,d) : DAYS[d];
  const slots=[]; dayList.forEach(di=>{ const n=teachingSlots(di,sid).length; for(let p=0;p<n;p++) slots.push([di,p]); });
  const busy={}; cells.forEach(x=>{ busy[x.teacher_id+'|'+x.day_of_week+'|'+x.period_index]=clsById[x.class_id]||'•'; });
  const ws=wb.addWorksheet('Teacher Free Periods');
  ws.addRow(['Teacher',...slots.map(([d,p])=>dayLbl(d)+'-P'+(p+1)),'Free']); styleHeader(ws.getRow(1));
  teachers.forEach(tc=>{ let free=0; const row=[tc.name];
    slots.forEach(([d,p])=>{ const b=busy[tc.id+'|'+d+'|'+p]; if(b) row.push(b); else { row.push(''); free++; } });
    row.push(free); const r=ws.addRow(row);
    r.eachCell((cc,i)=>{ if(i>1 && i<=slots.length+1 && !cc.value){ cc.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFD1FAE5'}}; } }); // free slots green
  });
  ws.columns.forEach((col,i)=>col.width=i===0?18:8); ws.getColumn(1).width=20;
  ws.views=[{state:'frozen',xSplit:1,ySplit:1}];
  res.setHeader('Content-Disposition','attachment; filename=teacher-free-periods.xlsx');
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  await wb.xlsx.write(res); res.end();
}));

// PERIOD-WISE LOAD (heatmap): day × period → how many classes are running; colour by load
app.get('/api/export/period-load.xlsx', h(async (req,res)=>{
  const wb=new ExcelJS.Workbook(); wb.creator='Aumtara'; const sid=req.sid;
  const cells=await q('SELECT day_of_week,period_index,teacher_id,subject_id FROM tt_timetable WHERE school_id=? AND week_index=0 AND subject_id IS NOT NULL',[sid]);
  const dayList=workingDaysArr(sid);
  const dayLbl = d => rotMode(sid) ? rotLabel(sid,d) : DAYS[d];
  const maxP=Math.max(1,...dayList.map(d=>teachingSlots(d,sid).length));
  const count={}; cells.forEach(c=>{ count[c.day_of_week+'|'+c.period_index]=(count[c.day_of_week+'|'+c.period_index]||0)+1; });
  let maxC=1; Object.values(count).forEach(v=>{ if(v>maxC)maxC=v; });
  const hx=n=>{const s=Math.max(0,Math.min(255,Math.round(n))).toString(16);return s.length<2?'0'+s:s;};
  const heat=r=>{ const R=Math.round(200+40*r), G=Math.round(235-150*r), B=Math.round(170-120*r); return 'FF'+hx(R)+hx(G)+hx(B); }; // light green → red
  const ws=wb.addWorksheet('Period Load');
  ws.addRow(['Day \\ Period',...Array.from({length:maxP},(_,i)=>'P'+(i+1)),'Day total']); styleHeader(ws.getRow(1));
  dayList.forEach(di=>{ const n=teachingSlots(di,sid).length; let tot=0; const row=[dayLbl(di)];
    for(let p=0;p<maxP;p++){ if(p>=n){ row.push(''); continue; } const c=count[di+'|'+p]||0; tot+=c; row.push(c); }
    row.push(tot); const r=ws.addRow(row);
    for(let p=0;p<maxP;p++){ if(p>=n)continue; const c=count[di+'|'+p]||0; const cell=r.getCell(p+2); cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:heat(c/maxC)}}; cell.alignment={horizontal:'center'}; }
    r.getCell(1).font={bold:true}; });
  // period totals row
  const totRow=['All days']; for(let p=0;p<maxP;p++){ let s=0; dayList.forEach(di=>{ s+=count[di+'|'+p]||0; }); totRow.push(s); } totRow.push('');
  const tr=ws.addRow(totRow); tr.font={bold:true}; tr.getCell(1).font={bold:true,color:{argb:'FF1F3864'}};
  ws.columns.forEach((col,i)=>col.width=i===0?14:8);
  ws.getCell('A'+(ws.rowCount+2)).value='More classes running in a slot = darker/redder. Use it to balance load across the day.';
  res.setHeader('Content-Disposition','attachment; filename=period-load.xlsx');
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  await wb.xlsx.write(res); res.end();
}));

app.get('/api/export/setup.xlsx', h(async (req,res)=>{
  const wb=new ExcelJS.Workbook(); wb.creator='Aumtara'; const sid=req.sid;
  const c=wb.addWorksheet('Classes'); c.addRow(['Class Name']); styleHeader(c.getRow(1)); (await q('SELECT name FROM tt_class WHERE school_id=? ORDER BY id',[sid])).forEach(x=>c.addRow([x.name])); c.getColumn(1).width=24;
  const s=wb.addWorksheet('Subjects'); s.addRow(['Subject Name']); styleHeader(s.getRow(1)); (await q('SELECT name FROM tt_subject WHERE school_id=? ORDER BY id',[sid])).forEach(x=>s.addRow([x.name])); s.getColumn(1).width=24;
  const r=wb.addWorksheet('Rooms'); r.addRow(['Room Name','Capacity']); styleHeader(r.getRow(1)); (await q('SELECT name,capacity FROM tt_room WHERE school_id=? ORDER BY id',[sid])).forEach(x=>r.addRow([x.name,x.capacity])); r.columns.forEach(cc=>cc.width=18);
  const t=wb.addWorksheet('Teachers'); t.addRow(['Teacher Name','Qualification','Main Subject','Optional Subjects (comma separated)','Weekly Period Load']); styleHeader(t.getRow(1));
  const map=await q('SELECT ts.teacher_id, ts.subject_id, s.name sub FROM tt_teacher_subject ts JOIN tt_subject s ON s.id=ts.subject_id WHERE s.school_id=?',[sid]);
  (await q('SELECT * FROM tt_teacher WHERE school_id=? ORDER BY id',[sid])).forEach(x=>{
    const mine=map.filter(m=>m.teacher_id===x.id);
    const main=x.main_subject_id?((mine.find(m=>m.subject_id===x.main_subject_id)||{}).sub||''):'';
    const opts=mine.filter(m=>m.subject_id!==x.main_subject_id).map(m=>m.sub).join(', ');
    t.addRow([x.name, x.qualification||'', main, opts, x.max_load||'']);
  });
  t.columns.forEach((cc,i)=>cc.width=[24,26,18,34,16][i]);
  const ch=wb.addWorksheet('Chapters'); ch.addRow(['Subject','Chapter']); styleHeader(ch.getRow(1));
  (await q('SELECT ch.name cn, s.name sn FROM tt_chapter ch JOIN tt_subject s ON s.id=ch.subject_id WHERE ch.school_id=? ORDER BY ch.subject_id, ch.seq',[sid])).forEach(r=>ch.addRow([r.sn, r.cn]));
  ch.columns.forEach((cc,i)=>cc.width=[20,30][i]);
  res.setHeader('Content-Disposition','attachment; filename=setup.xlsx');
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  await wb.xlsx.write(res); res.end();
}));

app.get('/api/export/setup-template.xlsx', h(async (_,res)=>{
  const wb=new ExcelJS.Workbook();
  const c=wb.addWorksheet('Classes'); c.addRow(['Class Name']); styleHeader(c.getRow(1)); c.addRow(['Class VI-A']); c.getColumn(1).width=24;
  const s=wb.addWorksheet('Subjects'); s.addRow(['Subject Name']); styleHeader(s.getRow(1)); s.addRow(['Mathematics']); s.getColumn(1).width=24;
  const r=wb.addWorksheet('Rooms'); r.addRow(['Room Name','Capacity']); styleHeader(r.getRow(1)); r.addRow(['Room 101',40]); r.columns.forEach(cc=>cc.width=18);
  const t=wb.addWorksheet('Teachers'); t.addRow(['Teacher Name','Qualification','Main Subject','Optional Subjects (comma separated)','Weekly Period Load']); styleHeader(t.getRow(1)); t.addRow(['R. Kumar','M.Sc, B.Ed','Mathematics','Computer, Science',30]); t.columns.forEach((cc,i)=>cc.width=[24,26,18,34,16][i]);
  const ch=wb.addWorksheet('Chapters'); ch.addRow(['Subject','Chapter']); styleHeader(ch.getRow(1)); ch.addRow(['Mathematics','Fractions']); ch.columns.forEach((cc,i)=>cc.width=[20,30][i]);
  res.setHeader('Content-Disposition','attachment; filename=setup-template.xlsx');
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  await wb.xlsx.write(res); res.end();
}));

// ---------- REPORTS (aSc-style multi-sheet workbook, from the school's data) ----------
function shortCode(name){ if(!name)return ''; const p=String(name).trim().split(/\s+/); if(p.length>1)return p.map(x=>x[0]).join('').toUpperCase().slice(0,4); return (String(name).replace(/[^A-Za-z0-9]/g,'').slice(0,3).toUpperCase())||String(name).slice(0,3); }
app.get('/api/export/reports.xlsx', h(async (req,res)=>{
  const sid=req.sid, cfg=getConfig(sid)||{};
  const classes =await q('SELECT * FROM tt_class WHERE school_id=? ORDER BY id',[sid]);
  const subjects=await q('SELECT * FROM tt_subject WHERE school_id=? ORDER BY id',[sid]);
  const teachers=await q('SELECT * FROM tt_teacher WHERE school_id=? ORDER BY id',[sid]);
  const rooms   =await q('SELECT * FROM tt_room WHERE school_id=? ORDER BY id',[sid]);
  const tmap    =await q('SELECT ts.* FROM tt_teacher_subject ts JOIN tt_teacher t ON t.id=ts.teacher_id WHERE t.school_id=?',[sid]);
  const tt      =await q('SELECT * FROM tt_timetable WHERE week_index=0 AND school_id=?',[sid]);
  const absents =await q('SELECT * FROM tt_absence WHERE school_id=?',[sid]);
  const cById={},sById={},tById={},rById={};
  classes.forEach(c=>cById[c.id]=c); subjects.forEach(s=>sById[s.id]=s); teachers.forEach(t=>tById[t.id]=t); rooms.forEach(r=>rById[r.id]=r);
  const cName=id=>(cById[id]||{}).name||'', sName=id=>(sById[id]||{}).name||'', tName=id=>(tById[id]||{}).name||'';
  const sShort={},tShort={}; subjects.forEach(s=>sShort[s.id]=shortCode(s.name)); teachers.forEach(t=>tShort[t.id]=shortCode(t.name));
  const wb=new ExcelJS.Workbook(); wb.creator='Aumtara';
  const title=(cfg.school_name||'School')+(cfg.board?(' · '+cfg.board):'')+(cfg.academic_session?(' · '+cfg.academic_session):'');
  const used=new Set();
  const uniq=nm=>{ let b=String(nm).slice(0,31).replace(/[\\\/\?\*\[\]:]/g,' ').trim()||'Sheet'; let n=b,i=2; while(used.has(n.toLowerCase())){ n=b.slice(0,27)+' '+i; i++; } used.add(n.toLowerCase()); return n; };
  const addSheet=(name,headers)=>{ const ws=wb.addWorksheet(uniq(name)); const span=Math.max(headers.length,2);
    ws.mergeCells(1,1,1,span); const tc=ws.getCell(1,1); tc.value=title; tc.font={bold:true,size:13,color:{argb:'FF1F3864'}}; tc.alignment={horizontal:'center'};
    ws.mergeCells(2,1,2,span); const sc=ws.getCell(2,1); sc.value=name; sc.font={bold:true,size:11,color:{argb:'FF6B3FA0'}}; sc.alignment={horizontal:'center'};
    const hr=ws.addRow(headers); styleHeader(hr); hr.eachCell(c=>c.border={top:{style:'thin'},bottom:{style:'thin'},left:{style:'thin'},right:{style:'thin'}}); return ws; };
  const widths=(ws,w)=>ws.columns.forEach((c,i)=>c.width=w[i]||14);
  const body=ws=>{ ws.eachRow((r,n)=>{ if(n>3) r.eachCell(c=>{c.alignment={vertical:'top',wrapText:true}; c.border={top:{style:'thin'},bottom:{style:'thin'},left:{style:'thin'},right:{style:'thin'}};}); }); };

  // Teachers list
  let ws=addSheet('Teachers list',['Name','Total lessons/week','Subjects','Classes','Class teacher for']);
  teachers.forEach(t=>{ const mine=tt.filter(x=>x.teacher_id===t.id);
    const subs=[...new Set(tmap.filter(m=>m.teacher_id===t.id).map(m=>sShort[m.subject_id]).filter(Boolean))].join(', ');
    const cls=[...new Set(mine.map(x=>cName(x.class_id)))].filter(Boolean).join(', ');
    const ct=classes.filter(c=>c.class_teacher_id===t.id).map(c=>c.name).join(', ');
    ws.addRow([t.name, mine.length, subs, cls, ct]); });
  widths(ws,[22,16,22,34,20]); body(ws);

  // Lessons
  ws=addSheet('Lessons',['Teacher','Class','Subject','Lessons/week']);
  const grp={}; tt.forEach(x=>{ if(!x.subject_id)return; const k=(x.teacher_id||0)+'|'+x.class_id+'|'+x.subject_id; grp[k]=(grp[k]||0)+1; });
  Object.keys(grp).sort((a,b)=>{const A=a.split('|'),B=b.split('|');return (tName(+A[0])||'~').localeCompare(tName(+B[0])||'~')||cName(+A[1]).localeCompare(cName(+B[1]));}).forEach(k=>{ const[tid,cid,sx]=k.split('|'); ws.addRow([tName(+tid)||'—', cName(+cid), sName(+sx), grp[k]]); });
  widths(ws,[22,12,20,14]); body(ws);

  // Subjects
  ws=addSheet('Subjects',['Subject','Short','Periods/week (all classes)']);
  subjects.forEach(s=>ws.addRow([s.name, sShort[s.id], tt.filter(x=>x.subject_id===s.id).length]));
  widths(ws,[22,10,26]); body(ws);

  // Classes
  ws=addSheet('Classes',['Class','Board','Medium','Standard','Section','Class teacher','Periods/week']);
  classes.forEach(c=>ws.addRow([c.name, c.board||'', c.medium||'', c.standard||'', c.section||'', c.class_teacher_id?tName(c.class_teacher_id):'', tt.filter(x=>x.class_id===c.id&&x.subject_id).length]));
  widths(ws,[16,12,12,10,10,20,14]); body(ws);

  // Teachers
  ws=addSheet('Teachers',['Name','Short','Qualification','Class teacher for','Weekly load','Max load']);
  teachers.forEach(t=>ws.addRow([t.name, tShort[t.id], t.qualification||'', classes.filter(c=>c.class_teacher_id===t.id).map(c=>c.name).join(', '), tt.filter(x=>x.teacher_id===t.id).length, t.max_load||'']));
  widths(ws,[22,10,22,20,12,10]); body(ws);

  // Classrooms
  ws=addSheet('Classrooms',['Name','Capacity']);
  rooms.forEach(r=>ws.addRow([r.name, r.capacity||''])); widths(ws,[20,12]); body(ws);

  const dayP=DAYS.map((_,di)=>teachingSlots(di,sid).length); const maxP=Math.max(1,...dayP);
  const absDay={}; absents.forEach(a=>{(absDay[a.day_of_week]=absDay[a.day_of_week]||new Set()).add(a.teacher_id);});

  // Available (free) teachers per day/period
  ws=addSheet('Available teachers',['Day','Period','Free teachers']);
  DAYS.forEach((d,di)=>{ for(let p=0;p<dayP[di];p++){ const busy=new Set(tt.filter(x=>x.day_of_week===di&&x.period_index===p&&x.teacher_id).map(x=>x.teacher_id)); const ab=absDay[di]||new Set(); ws.addRow([d,p+1,teachers.filter(t=>!busy.has(t.id)&&!ab.has(t.id)).map(t=>tShort[t.id]).join(', ')]); }});
  widths(ws,[10,8,70]); body(ws);

  // Unused classrooms per day/period
  ws=addSheet('Unused classrooms',['Day','Period','Free rooms']);
  DAYS.forEach((d,di)=>{ for(let p=0;p<dayP[di];p++){ const u=new Set(tt.filter(x=>x.day_of_week===di&&x.period_index===p&&x.room_id).map(x=>x.room_id)); ws.addRow([d,p+1,rooms.filter(r=>!u.has(r.id)).map(r=>r.name).join(', ')]); }});
  widths(ws,[10,8,70]); body(ws);

  // Per-subject grids (Day x Period -> classes having that subject)
  subjects.forEach(s=>{ const ws=addSheet(s.name+' .',['Day','Period',...Array.from({length:maxP},(_,i)=>'P'+(i+1))]);
    DAYS.forEach((d,di)=>{ for(let p=0;p<dayP[di];p++){ const cls=tt.filter(x=>x.day_of_week===di&&x.period_index===p&&x.subject_id===s.id).map(x=>cName(x.class_id)); ws.addRow([d,p+1,...cls]); }});
    widths(ws,[10,8,...Array.from({length:maxP},()=>10)]); body(ws); });

  res.setHeader('Content-Disposition','attachment; filename=reports.xlsx');
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  await wb.xlsx.write(res); res.end();
}));

// ---------- REPORT DATA (shared: scalars + collections for template fill) ----------
async function gatherReportData(sid){
  const cfg=getConfig(sid)||{};
  const classes =await q('SELECT * FROM tt_class WHERE school_id=? ORDER BY id',[sid]);
  const subjects=await q('SELECT * FROM tt_subject WHERE school_id=? ORDER BY id',[sid]);
  const teachers=await q('SELECT * FROM tt_teacher WHERE school_id=? ORDER BY id',[sid]);
  const rooms   =await q('SELECT * FROM tt_room WHERE school_id=? ORDER BY id',[sid]);
  const tmap    =await q('SELECT ts.* FROM tt_teacher_subject ts JOIN tt_teacher t ON t.id=ts.teacher_id WHERE t.school_id=?',[sid]);
  const tt      =await q('SELECT * FROM tt_timetable WHERE week_index=0 AND school_id=?',[sid]);
  const cN={},sN={},tN={}; classes.forEach(c=>cN[c.id]=c.name); subjects.forEach(s=>sN[s.id]=s.name); teachers.forEach(t=>tN[t.id]=t.name);
  const dt=new Date(); const dateStr=String(dt.getFullYear())+'-'+String(dt.getMonth()+1).padStart(2,'0')+'-'+String(dt.getDate()).padStart(2,'0');
  const scalars={ school_name:cfg.school_name||'', board:cfg.board||'', medium:cfg.medium||'', session:cfg.academic_session||'', date:dateStr,
    class_count:classes.length, teacher_count:teachers.length, subject_count:subjects.length, room_count:rooms.length };
  const grp={}; tt.forEach(x=>{ if(!x.subject_id)return; const k=(x.teacher_id||0)+'|'+x.class_id+'|'+x.subject_id; grp[k]=(grp[k]||0)+1; });
  const collections={
    teachers: teachers.map(t=>({ name:t.name, short:shortCode(t.name), qualification:t.qualification||'', max_load:t.max_load||'',
      load: tt.filter(x=>x.teacher_id===t.id).length,
      subjects:[...new Set(tmap.filter(m=>m.teacher_id===t.id).map(m=>sN[m.subject_id]).filter(Boolean))].join(', '),
      classes:[...new Set(tt.filter(x=>x.teacher_id===t.id).map(x=>cN[x.class_id]).filter(Boolean))].join(', '),
      class_teacher_for: classes.filter(c=>c.class_teacher_id===t.id).map(c=>c.name).join(', ') })),
    classes: classes.map(c=>({ name:c.name, board:c.board||'', medium:c.medium||'', standard:c.standard||'', section:c.section||'',
      class_teacher: c.class_teacher_id?(tN[c.class_teacher_id]||''):'', periods: tt.filter(x=>x.class_id===c.id&&x.subject_id).length })),
    subjects: subjects.map(s=>({ name:s.name, short:shortCode(s.name), active:(+s.active!==0?'Yes':'No'), periods: tt.filter(x=>x.subject_id===s.id).length })),
    rooms: rooms.map(r=>({ name:r.name, capacity:r.capacity||'' })),
    lessons: Object.keys(grp).map(k=>{ const[tid,cid,sx]=k.split('|'); return { teacher:tN[+tid]||'—', class:cN[+cid]||'', subject:sN[+sx]||'', lessons_week:grp[k] }; }),
  };
  return { scalars, collections };
}

// ---------- REPORT #1: fill an uploaded template with this school's data ----------
app.post('/api/reports/template', upload.single('file'), async (req,res)=>{
  try{
    const sid=req.sid; const { scalars, collections }=await gatherReportData(sid);
    const wb=new ExcelJS.Workbook(); await wb.xlsx.load(req.file.buffer);
    const scal=str=>String(str).replace(/\{\{\s*([a-zA-Z0-9_]+)\s*\}\}/g,(m,k)=> scalars[k]!==undefined?String(scalars[k]):m);
    wb.eachSheet(ws=>{
      // 1) expand repeat rows: a cell containing {{rows:collection}} marks a row template
      const blocks=[];
      ws.eachRow((row,rn)=>{ row.eachCell({includeEmpty:false},cell=>{ if(typeof cell.value==='string'){ const mm=cell.value.match(/\{\{\s*rows:([a-zA-Z0-9_]+)\s*\}\}/); if(mm)blocks.push({rn,coll:mm[1]}); } }); });
      blocks.sort((a,b)=>b.rn-a.rn).forEach(({rn,coll})=>{
        const items=collections[coll]||[];
        const tmplRow=ws.getRow(rn); const tmpl={}; tmplRow.eachCell({includeEmpty:true},(cell,cn)=>{ tmpl[cn]=cell.value; });
        if(items.length>1) ws.duplicateRow(rn, items.length-1, true);
        const fillRow=(r,it)=>{ r.eachCell({includeEmpty:true},(cell,cn)=>{ let v=tmpl[cn]; if(typeof v==='string'){ v=v.replace(/\{\{\s*rows:[a-zA-Z0-9_]+\s*\}\}/g,''); v=v.replace(/\{\{\s*([a-zA-Z0-9_]+)\s*\}\}/g,(m,k)=> it[k]!==undefined?String(it[k]):(scalars[k]!==undefined?String(scalars[k]):'')); cell.value=v; } }); };
        if(items.length===0){ tmplRow.eachCell({includeEmpty:true},(cell,cn)=>{ if(typeof tmpl[cn]==='string') cell.value=tmpl[cn].replace(/\{\{\s*rows:[a-zA-Z0-9_]+\s*\}\}/g,'').replace(/\{\{\s*[a-zA-Z0-9_]+\s*\}\}/g,''); }); }
        else items.forEach((it,i)=> fillRow(ws.getRow(rn+i), it));
      });
      // 2) scalar replace everywhere
      ws.eachRow(row=>row.eachCell({includeEmpty:false},cell=>{ if(typeof cell.value==='string') cell.value=scal(cell.value); }));
    });
    res.setHeader('Content-Disposition','attachment; filename=filled-report.xlsx');
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    await wb.xlsx.write(res); res.end();
  }catch(e){ console.error(e); res.status(400).json({error:String(e.message||e)}); }
});

// ---------- REPORT #1b: sample template so users learn the placeholders ----------
app.get('/api/reports/template-sample.xlsx', h(async (req,res)=>{
  const wb=new ExcelJS.Workbook(); wb.creator='Aumtara';
  const ws=wb.addWorksheet('Template');
  ws.addRow(['{{school_name}}']); ws.getCell('A1').font={bold:true,size:15};
  ws.addRow(['Board: {{board}}   Medium: {{medium}}   Session: {{session}}   Date: {{date}}']);
  ws.addRow(['Totals: {{class_count}} classes · {{teacher_count}} teachers · {{subject_count}} subjects']);
  ws.addRow([]);
  const h1=ws.addRow(['Teacher','Short','Qualification','Subjects','Classes','Weekly load','Class teacher for']); styleHeader(h1);
  ws.addRow(['{{rows:teachers}}{{name}}','{{short}}','{{qualification}}','{{subjects}}','{{classes}}','{{load}}','{{class_teacher_for}}']);
  ws.addRow([]);
  const h2=ws.addRow(['Class','Standard','Section','Class teacher','Periods/week']); styleHeader(h2);
  ws.addRow(['{{rows:classes}}{{name}}','{{standard}}','{{section}}','{{class_teacher}}','{{periods}}']);
  ws.addRow([]);
  const gd=ws.addRow(['GUIDE — scalars: {{school_name}} {{board}} {{medium}} {{session}} {{date}} {{class_count}} {{teacher_count}} {{subject_count}} {{room_count}}']); gd.getCell(1).font={italic:true,color:{argb:'FF8A93B4'}};
  const gd2=ws.addRow(['GUIDE — repeat a row: put {{rows:teachers}} (or classes / subjects / rooms / lessons) in the first cell, then {{field}} tokens in that row. teachers fields: name,short,qualification,subjects,classes,load,max_load,class_teacher_for · classes: name,board,medium,standard,section,class_teacher,periods · subjects: name,short,periods,active · rooms: name,capacity · lessons: teacher,class,subject,lessons_week']); gd2.getCell(1).font={italic:true,color:{argb:'FF8A93B4'}};
  ws.columns.forEach((c,i)=>c.width=[24,10,20,24,24,12,20][i]||18);
  res.setHeader('Content-Disposition','attachment; filename=report-template-sample.xlsx');
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  await wb.xlsx.write(res); res.end();
}));

// ---------- REPORT #3: beautify an uploaded data Excel into a formatted report ----------
app.post('/api/reports/format', upload.single('file'), async (req,res)=>{
  try{
    const sid=req.sid, cfg=getConfig(sid)||{};
    const title=String((req.body&&req.body.title)||'').trim() || (cfg.school_name||'Report');
    const allB={top:{style:'thin'},bottom:{style:'thin'},left:{style:'thin'},right:{style:'thin'}};
    const inWb=new ExcelJS.Workbook(); await inWb.xlsx.load(req.file.buffer);
    const outWb=new ExcelJS.Workbook(); outWb.creator='Aumtara';
    const vtext=v=>{ if(v==null)return ''; if(typeof v==='object')return v.text||v.result||(v.richText&&v.richText.map(x=>x.text).join(''))||''; return v; };
    let sheets=0;
    inWb.eachSheet(ws=>{
      const data=[]; ws.eachRow({includeEmpty:false},row=>{ const vals=[]; row.eachCell({includeEmpty:true},(cell,cn)=>{ vals[cn-1]=vtext(cell.value); }); if(vals.some(v=>String(v).trim()!==''))data.push(vals); });
      if(!data.length)return; sheets++;
      const o=outWb.addWorksheet(ws.name.slice(0,31).replace(/[\\\/\?\*\[\]:]/g,' ')||('Sheet'+sheets));
      const cols=Math.max(1,...data.map(r=>r.length));
      o.mergeCells(1,1,1,cols); const tc=o.getCell(1,1); tc.value=title; tc.font={bold:true,size:14,color:{argb:'FF1F3864'}}; tc.alignment={horizontal:'center'};
      const hdr=o.addRow(data[0]); styleHeader(hdr); hdr.eachCell(c=>c.border=allB);
      for(let i=1;i<data.length;i++){ const r=o.addRow(data[i]); r.eachCell(c=>{ c.alignment={vertical:'top',wrapText:true}; c.border=allB; }); }
      o.columns.forEach((c,i)=>{ let w=10; data.forEach(r=>{ const v=r[i]; if(v!=null) w=Math.max(w, Math.min(42, String(v).length+2)); }); c.width=w; });
      o.views=[{state:'frozen', ySplit:2}];
    });
    if(!sheets){ res.status(400).json({error:'no data found in the uploaded file'}); return; }
    res.setHeader('Content-Disposition','attachment; filename=formatted-report.xlsx');
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    await outWb.xlsx.write(res); res.end();
  }catch(e){ console.error(e); res.status(400).json({error:String(e.message||e)}); }
});

// ---------- EXCEL IMPORT (setup) ----------
app.post('/api/import/setup', upload.single('file'), async (req,res)=>{
  try{
    const sid=req.sid;
    const wb=new ExcelJS.Workbook(); await wb.xlsx.load(req.file.buffer);
    const added={classes:0,subjects:0,rooms:0,teachers:0,chapters:0};
    const val=cell=>{ const v=cell&&cell.value; return v==null?'':(typeof v==='object'&&v.text?v.text:String(v)).trim(); };
    const readSheet=(name)=>{ const wsx=wb.getWorksheet(name); const out=[]; if(wsx) wsx.eachRow((row,n)=>{ if(n===1)return; out.push(row); }); return out; };
    const classRows=readSheet('Classes'), subjRows=readSheet('Subjects'), roomRows=readSheet('Rooms'), teacherRows=readSheet('Teachers'), chapterRows=readSheet('Chapters');

    await tx(async (cq,cq1)=>{
      const subId=async(name)=>{ let r=await cq1('SELECT id FROM tt_subject WHERE lower(name)=lower(?) AND school_id=?',[name,sid]);
        if(!r){ const i=await cq1('INSERT INTO tt_subject(name,school_id) VALUES(?,?) RETURNING id',[name,sid]); added.subjects++; return i.id; } return r.id; };

      for(const row of classRows){ const name=val(row.getCell(1)); if(name && !(await cq1('SELECT 1 FROM tt_class WHERE lower(name)=lower(?) AND school_id=?',[name,sid]))){ await cq('INSERT INTO tt_class(name,school_id) VALUES(?,?)',[name,sid]); added.classes++; } }
      for(const row of subjRows){ const name=val(row.getCell(1)); if(name && !(await cq1('SELECT 1 FROM tt_subject WHERE lower(name)=lower(?) AND school_id=?',[name,sid]))){ await cq('INSERT INTO tt_subject(name,school_id) VALUES(?,?)',[name,sid]); added.subjects++; } }
      for(const row of roomRows){ const name=val(row.getCell(1)); if(name && !(await cq1('SELECT 1 FROM tt_room WHERE lower(name)=lower(?) AND school_id=?',[name,sid]))){ await cq('INSERT INTO tt_room(name,capacity,school_id) VALUES(?,?,?)',[name, Number(val(row.getCell(2)))||null, sid]); added.rooms++; } }

      for(const row of teacherRows){ const name=val(row.getCell(1)); if(!name)continue;
        const qual=val(row.getCell(2))||null;
        const mainName=val(row.getCell(3));
        const optNames=val(row.getCell(4)).split(',').map(x=>x.trim()).filter(Boolean);
        const load=Number(val(row.getCell(5)))||null;
        const mainId=mainName?await subId(mainName):null;
        let tr=await cq1('SELECT id FROM tt_teacher WHERE lower(name)=lower(?) AND school_id=?',[name,sid]);
        let tid;
        if(tr){ tid=tr.id; await cq('UPDATE tt_teacher SET qualification=COALESCE(?,qualification), main_subject_id=COALESCE(?,main_subject_id), max_load=COALESCE(?,max_load) WHERE id=?',[qual,mainId,load,tid]); }
        else { added.teachers++; tid=(await cq1('INSERT INTO tt_teacher(name,qualification,main_subject_id,max_load,school_id) VALUES(?,?,?,?,?) RETURNING id',[name,qual,mainId,load,sid])).id; }
        if(mainId) await cq('INSERT INTO tt_teacher_subject(teacher_id,subject_id) VALUES(?,?) ON CONFLICT DO NOTHING',[tid, mainId]);
        for(const sn of optNames){ await cq('INSERT INTO tt_teacher_subject(teacher_id,subject_id) VALUES(?,?) ON CONFLICT DO NOTHING',[tid, await subId(sn)]); }
      }

      let ci=0;
      for(const row of chapterRows){ ci++; const sn=val(row.getCell(1)), cn=val(row.getCell(2)); if(!sn||!cn)continue;
        const subjid=await subId(sn);
        if(!(await cq1('SELECT 1 FROM tt_chapter WHERE subject_id=? AND lower(name)=lower(?) AND school_id=?',[subjid,cn,sid]))){ await cq('INSERT INTO tt_chapter(subject_id,name,seq,school_id) VALUES(?,?,?,?)',[subjid,cn,ci,sid]); added.chapters++; } }
    });
    res.json({ok:true, added});
  }catch(e){ res.status(400).json({ok:false,error:String(e.message||e)}); }
});

// ---------- DE-DUPLICATE ENTITIES (merge exact same-name rows, keep lowest id, repoint EVERY ref) ----------
// Works for subjects / classes / rooms / teachers. Schema-driven: discovers referencing columns from
// information_schema and unique key-sets from pg_index, so collisions are guarded automatically. Also
// handles the polymorphic tt_avail (entity_type,entity_id) + tt_sharelink (kind,target_id) and the
// tt_combine.class_ids CSV. Timetable cells are preserved (their unique key is the slot, not the entity).
const DEDUPE_CFG={
  subjects:{ table:'tt_subject', colWhere:"column_name IN ('subject_id','main_subject_id')", avail:null,  csvCombine:false },
  classes: { table:'tt_class',   colWhere:"column_name LIKE '%class_id' AND column_name<>'class_ids'",     avail:'class',  csvCombine:true  },
  rooms:   { table:'tt_room',    colWhere:"column_name LIKE '%room_id'",                                   avail:'room',   csvCombine:false },
  teachers:{ table:'tt_teacher', colWhere:"column_name LIKE '%teacher_id'",                                avail:'teacher',csvCombine:false },
};
async function dedupeCounts(qfn, sid){
  const out={};
  for(const k of Object.keys(DEDUPE_CFG)){
    const g=await qfn(`SELECT count(*)::int n FROM (SELECT 1 FROM ${DEDUPE_CFG[k].table} WHERE school_id=? GROUP BY lower(btrim(name)) HAVING count(*)>1) x`,[sid]);
    const rem=await qfn(`SELECT COALESCE(sum(c-1),0)::int r FROM (SELECT count(*) c FROM ${DEDUPE_CFG[k].table} WHERE school_id=? GROUP BY lower(btrim(name)) HAVING count(*)>1) y`,[sid]);
    out[k]={dupNames:(g[0]&&g[0].n)||0, removable:(rem[0]&&rem[0].r)||0};
  }
  return out;
}
app.get('/api/dedupe/count', h(async (req,res)=>{ res.json(await dedupeCounts(q, req.sid)); }));
// kept for backward-compat with the already-shipped subjects button
app.get('/api/dedupe/subjects/count', h(async (req,res)=>{ res.json((await dedupeCounts(q, req.sid)).subjects); }));

async function mergeEntity(cq, sid, cfg){
  if(!cfg) throw new Error('bad entity');
  // referencing columns (in OTHER tables), identifier-guarded
  const refs=(await cq(`SELECT table_name, column_name FROM information_schema.columns
     WHERE table_schema='public' AND (${cfg.colWhere}) AND table_name<>'${cfg.table}' AND table_name LIKE 'tt_%'`))
     .filter(r=>/^[a-z_]+$/.test(r.table_name) && /^[a-z_]+$/.test(r.column_name));
  // unique key-sets per table (covers PK + UNIQUE constraints + unique indexes — all backed by pg_index)
  const ixRows=await cq(`SELECT tbl.relname tname, idx.relname iname, a.attname col
     FROM pg_index ix JOIN pg_class idx ON idx.oid=ix.indexrelid JOIN pg_class tbl ON tbl.oid=ix.indrelid
     JOIN pg_attribute a ON a.attrelid=tbl.oid AND a.attnum=ANY(ix.indkey)
     WHERE ix.indisunique AND tbl.relkind='r' AND tbl.relname LIKE 'tt_%'`);
  const uniq={};   // table -> array of column-sets
  const seenIx={};
  ixRows.forEach(r=>{ const key=r.tname+'|'+r.iname; (seenIx[key]=seenIx[key]||[]).push(r.col); });
  Object.keys(seenIx).forEach(k=>{ const t=k.split('|')[0]; (uniq[t]=uniq[t]||[]).push(seenIx[k]); });
  const ident=x=>/^[a-z_]+$/.test(x);

  const groups=await cq(`SELECT array_agg(id ORDER BY id) ids FROM ${cfg.table} WHERE school_id=? GROUP BY lower(btrim(name)) HAVING count(*)>1`,[sid]);
  let removed=0, merged=0;
  for(const g of groups){
    const ids=g.ids.map(Number); const keep=ids[0];
    for(const dup of ids.slice(1)){
      for(const r of refs){
        if(!ident(r.table_name)||!ident(r.column_name)) continue;
        // guard every unique key-set of this table that includes the ref column
        for(const set of (uniq[r.table_name]||[])){
          if(!set.includes(r.column_name)) continue;
          const others=set.filter(c=>c!==r.column_name);
          if(!others.length || !others.every(ident)) continue;
          const cols=others.join(',');
          await cq(`DELETE FROM ${r.table_name} WHERE ${r.column_name}=? AND (${cols}) IN (SELECT ${cols} FROM ${r.table_name} WHERE ${r.column_name}=?)`,[dup,keep]);
        }
        await cq(`UPDATE ${r.table_name} SET ${r.column_name}=? WHERE ${r.column_name}=?`,[keep,dup]);
      }
      // polymorphic availability (entity_type,entity_id) — guard on its unique slot key
      if(cfg.avail){
        await cq(`DELETE FROM tt_avail WHERE entity_type=? AND entity_id=? AND (school_id,day_of_week,period_index) IN (SELECT school_id,day_of_week,period_index FROM tt_avail WHERE entity_type=? AND entity_id=?)`,[cfg.avail,dup,cfg.avail,keep]);
        await cq(`UPDATE tt_avail SET entity_id=? WHERE entity_type=? AND entity_id=?`,[keep,cfg.avail,dup]);
        await cq(`UPDATE tt_sharelink SET target_id=? WHERE kind=? AND target_id=?`,[keep,cfg.avail,dup]);
      }
      // combined-classes CSV (class merge only): swap dup id -> keep id inside class_ids, de-dup
      if(cfg.csvCombine){
        const rows=await cq(`SELECT id,class_ids FROM tt_combine WHERE school_id=?`,[sid]);
        for(const row of rows){ if(!row.class_ids) continue;
          let arr=String(row.class_ids).split(',').map(x=>x.trim()).filter(Boolean);
          if(arr.includes(String(dup))){ arr=[...new Set(arr.map(x=>x===String(dup)?String(keep):x))];
            await cq(`UPDATE tt_combine SET class_ids=? WHERE id=?`,[arr.join(','),row.id]); }
        }
      }
      await cq(`DELETE FROM ${cfg.table} WHERE id=? AND school_id=?`,[dup,sid]);
      removed++;
    }
    merged++;
  }
  return {removed, merged};
}
app.post('/api/dedupe/:entity', h(async (req,res)=>{
  if(!requireAdmin(req,res)) return;
  const cfg=DEDUPE_CFG[req.params.entity];
  if(!cfg){ res.status(400).json({error:'unknown entity'}); return; }
  const out=await tx(async (cq)=>mergeEntity(cq, req.sid, cfg));
  res.json({ok:true, ...out});
}));

// ---------- ADMIN DATA CLEANUP (clear a field / a section / a master list / everything — current school only) ----------
// Each target = ordered list of statements; every `?` is the school_id. Scoped to req.sid; never touches other schools,
// and (except 'all') never deletes the school, its config or user logins. Admin/master only.
const CLEAN_TARGETS={
  // field clears (blank a column for all rows)
  t_qual:     ['UPDATE tt_teacher SET qualification=NULL WHERE school_id=?'],
  t_caps:     ['UPDATE tt_teacher SET max_load=NULL,max_per_day=NULL,max_consecutive=NULL WHERE school_id=?'],
  t_main:     ['UPDATE tt_teacher SET main_subject_id=NULL WHERE school_id=?'],
  t_canteach: ['DELETE FROM tt_teacher_subject WHERE teacher_id IN (SELECT id FROM tt_teacher WHERE school_id=?)'],
  s_medium:   ['UPDATE tt_subject SET medium=NULL WHERE school_id=?'],
  s_double:   ['UPDATE tt_subject SET double_period=0 WHERE school_id=?'],
  c_struct:   ['UPDATE tt_class SET board=NULL,medium=NULL,standard=NULL,section=NULL WHERE school_id=?'],
  c_teacher:  ['UPDATE tt_class SET class_teacher_id=NULL WHERE school_id=?'],
  r_cap:      ['UPDATE tt_room SET capacity=NULL WHERE school_id=?'],
  // section clears (delete data, keep the master lists)
  d_timetable:['DELETE FROM tt_timetable WHERE school_id=?'],
  d_quota:    ['DELETE FROM tt_quota WHERE school_id=?'],
  d_chapters: ['DELETE FROM tt_chapter WHERE school_id=?'],
  d_avail:    ['DELETE FROM tt_avail WHERE school_id=?'],
  d_subs:     ['DELETE FROM tt_substitution WHERE school_id=?','DELETE FROM tt_datesub WHERE school_id=?','DELETE FROM tt_leave WHERE school_id=?','DELETE FROM tt_absence WHERE school_id=?'],
  d_diary:    ['DELETE FROM tt_diary WHERE school_id=?'],
  d_versions: ['DELETE FROM tt_snapshot_cell WHERE snapshot_id IN (SELECT id FROM tt_snapshot WHERE school_id=?)','DELETE FROM tt_snapshot WHERE school_id=?'],
  d_combines: ['DELETE FROM tt_combine WHERE school_id=?'],
  d_students: ['DELETE FROM tt_student_choice WHERE school_id=?','DELETE FROM tt_student WHERE school_id=?','DELETE FROM tt_elective_option WHERE school_id=?','DELETE FROM tt_elective WHERE school_id=?'],
  d_rules:    ['DELETE FROM tt_rule WHERE school_id=?'],
  // master-list clears (delete the list + its directly-dependent data, to avoid orphans)
  m_subjects: ['DELETE FROM tt_timetable WHERE school_id=?','DELETE FROM tt_quota WHERE school_id=?','DELETE FROM tt_chapter WHERE school_id=?','DELETE FROM tt_rule WHERE school_id=?','DELETE FROM tt_teacher_subject WHERE teacher_id IN (SELECT id FROM tt_teacher WHERE school_id=?)','UPDATE tt_teacher SET main_subject_id=NULL WHERE school_id=?','DELETE FROM tt_subject WHERE school_id=?'],
  m_classes:  ['DELETE FROM tt_timetable WHERE school_id=?','DELETE FROM tt_quota WHERE school_id=?','DELETE FROM tt_student_choice WHERE school_id=?','DELETE FROM tt_student WHERE school_id=?','DELETE FROM tt_elective_option WHERE school_id=?','DELETE FROM tt_elective WHERE school_id=?','DELETE FROM tt_datesub WHERE school_id=?','DELETE FROM tt_substitution WHERE school_id=?','DELETE FROM tt_diary WHERE school_id=?','DELETE FROM tt_rule WHERE school_id=?','DELETE FROM tt_combine WHERE school_id=?','DELETE FROM tt_class WHERE school_id=?'],
  m_teachers: ['UPDATE tt_timetable SET teacher_id=NULL WHERE school_id=?','UPDATE tt_class SET class_teacher_id=NULL WHERE school_id=?','DELETE FROM tt_teacher_subject WHERE teacher_id IN (SELECT id FROM tt_teacher WHERE school_id=?)','DELETE FROM tt_absence WHERE school_id=?','DELETE FROM tt_leave WHERE school_id=?','DELETE FROM tt_datesub WHERE school_id=?','DELETE FROM tt_teacher WHERE school_id=?'],
  m_rooms:    ['UPDATE tt_timetable SET room_id=NULL WHERE school_id=?','UPDATE tt_combine SET room_id=NULL WHERE school_id=?','UPDATE tt_elective_option SET room_id=NULL WHERE school_id=?','DELETE FROM tt_room WHERE school_id=?'],
  // full reset — wipe ALL data for this school, keep the school row, its config and user logins
  all: ['DELETE FROM tt_snapshot_cell WHERE snapshot_id IN (SELECT id FROM tt_snapshot WHERE school_id=?)',
        'DELETE FROM tt_teacher_subject WHERE teacher_id IN (SELECT id FROM tt_teacher WHERE school_id=?)',
        'DELETE FROM tt_timetable WHERE school_id=?','DELETE FROM tt_quota WHERE school_id=?','DELETE FROM tt_chapter WHERE school_id=?','DELETE FROM tt_absence WHERE school_id=?','DELETE FROM tt_substitution WHERE school_id=?','DELETE FROM tt_datesub WHERE school_id=?','DELETE FROM tt_leave WHERE school_id=?','DELETE FROM tt_diary WHERE school_id=?','DELETE FROM tt_snapshot WHERE school_id=?','DELETE FROM tt_student_choice WHERE school_id=?','DELETE FROM tt_student WHERE school_id=?','DELETE FROM tt_elective_option WHERE school_id=?','DELETE FROM tt_elective WHERE school_id=?','DELETE FROM tt_combine WHERE school_id=?','DELETE FROM tt_rule WHERE school_id=?','DELETE FROM tt_avail WHERE school_id=?','DELETE FROM tt_sharelink WHERE school_id=?','DELETE FROM tt_term WHERE school_id=?','DELETE FROM tt_class WHERE school_id=?','DELETE FROM tt_subject WHERE school_id=?','DELETE FROM tt_room WHERE school_id=?','DELETE FROM tt_teacher WHERE school_id=?'],
};
app.post('/api/admin/clean', h(async (req,res)=>{
  if(!requireAdmin(req,res)) return;
  const sid=req.sid;
  if(sid==null){ res.status(400).json({error:'no school selected'}); return; }
  const target=(req.body||{}).target;
  const stmts=CLEAN_TARGETS[target];
  if(!stmts){ res.status(400).json({error:'unknown clean target'}); return; }
  // destructive targets (full reset + clearing a whole list) require typing the exact school name
  if(target==='all' || /^m_/.test(target)){
    const sch=await q1('SELECT name FROM tt_school WHERE id=?',[sid]);
    const want=((sch&&sch.name)||'').trim().toLowerCase();
    const got=(((req.body||{}).confirm)||'').trim().toLowerCase();
    if(!want || got!==want){ res.status(400).json({error:'Type the exact school name to confirm this.'}); return; }
  }
  await tx(async (cq)=>{ for(const s of stmts){ const n=(s.split('?').length-1); await cq(s, Array(n).fill(sid)); } });
  res.json({ok:true});
}));

// ---------- AUTO-FILL HOURS FROM EXCEL (deterministic, no AI) ----------
function hm(v){
  if(v==null) return '';
  if(v instanceof Date){ return String(v.getUTCHours()).padStart(2,'0')+':'+String(v.getUTCMinutes()).padStart(2,'0'); }
  if(typeof v==='object' && v.text!=null) v=v.text;
  if(typeof v==='number'){ let mins=Math.round(v*24*60)%(24*60); if(mins<0)mins+=24*60; return String(Math.floor(mins/60)).padStart(2,'0')+':'+String(mins%60).padStart(2,'0'); }
  const s=String(v).trim(); const m=s.match(/(\d{1,2}):(\d{2})/); return m?String(+m[1]).padStart(2,'0')+':'+m[2]:'';
}
function toMins(hhmm){ const m=String(hhmm).match(/(\d{1,2}):(\d{2})/); return m?(+m[1])*60+(+m[2]):null; }
function parseHoursSheet(ws){
  if(!ws) return null;
  const rows=[];
  ws.eachRow((row,n)=>{ if(n===1) return;
    const lv=row.getCell(1).value; const label=String((lv&&(lv.text!=null?lv.text:lv))||'').trim();
    const start=hm(row.getCell(2).value), end=hm(row.getCell(3).value);
    if(!start||!end) return;
    let kind='period';
    if(/lunch/i.test(label)) kind='lunch';
    else if(/break|recess|short|interval|assembly/i.test(label)) kind='short';
    rows.push({label,start,end,kind,s:toMins(start),e:toMins(end)});
  });
  const periods=rows.filter(r=>r.kind==='period' && r.s!=null && r.e!=null);
  if(!periods.length) return null;
  const startT=periods[0].start, endT=periods[periods.length-1].end;
  const durs=periods.map(r=>Math.max(1,r.e-r.s));
  const freq={}; durs.forEach(d=>freq[d]=(freq[d]||0)+1);
  const periodMin=+Object.keys(freq).sort((a,b)=>freq[b]-freq[a])[0];
  const allSame=durs.every(d=>d===periodMin);
  const afterCount=r=>periods.filter(p=>p.s<r.s).length;
  const lunchRow=rows.find(r=>r.kind==='lunch');
  const breakRows=rows.filter(r=>r.kind==='short');
  return { start:startT, end:endT, num:periods.length, period_minutes:periodMin,
    period_durations: allSame?'':durs.join(','),
    lunch_after: lunchRow?afterCount(lunchRow):null,
    lunch_minutes: lunchRow?Math.max(1,lunchRow.e-lunchRow.s):0,
    short_break_after: breakRows.length?[...new Set(breakRows.map(afterCount))].sort((a,b)=>a-b).join(','):'',
    short_break_minutes: breakRows.length?Math.max(1,breakRows[0].e-breakRows[0].s):0 };
}
// ---------- DEO / GOVT REPORTS: teacher workload statement + teacher requirement (staff pattern) ----------
async function deoData(sid){
  const cfg=getConfig(sid)||{};
  const sch=(await q1('SELECT name,board,medium,udise,district FROM tt_school WHERE id=?',[sid]))||{};
  const teachers=await q('SELECT * FROM tt_teacher WHERE school_id=? ORDER BY name,id',[sid]);
  const subjById={}; (await q('SELECT id,name FROM tt_subject WHERE school_id=?',[sid])).forEach(s=>subjById[s.id]=s.name);
  const clsById={}; (await q('SELECT id,name FROM tt_class WHERE school_id=?',[sid])).forEach(c=>clsById[c.id]=c);
  const cells=await q('SELECT class_id,subject_id,teacher_id FROM tt_timetable WHERE school_id=? AND week_index=0 AND subject_id IS NOT NULL',[sid]);
  const tsub={}; (await q('SELECT ts.teacher_id,ts.subject_id FROM tt_teacher_subject ts JOIN tt_teacher t ON t.id=ts.teacher_id WHERE t.school_id=?',[sid])).forEach(r=>{ (tsub[r.subject_id]=tsub[r.subject_id]||new Set()).add(r.teacher_id); });
  teachers.forEach(t=>{ if(t.main_subject_id){ (tsub[t.main_subject_id]=tsub[t.main_subject_id]||new Set()).add(t.id); } });
  const workload=teachers.map((t,i)=>{
    const tc=cells.filter(c=>c.teacher_id===t.id); const total=tc.length;
    const bySub={}; tc.forEach(c=>{ bySub[c.subject_id]=(bySub[c.subject_id]||0)+1; });
    const subjects=Object.keys(bySub).map(s=>subjById[s]||('#'+s)).join(', ');
    const subjDist=Object.keys(bySub).map(s=>(subjById[s]||('#'+s))+': '+bySub[s]).join(', ');
    const required=(t.sanctioned_load||t.max_load)||null;
    const diff=required!=null?(total-required):null;
    return { sr:i+1, name:t.name||'', designation:t.designation||'', qualification:t.qualification||'', main:subjById[t.main_subject_id]||'', subjects, subjDist, total, required, diff };
  });
  const loads=teachers.map(t=>t.sanctioned_load||t.max_load).filter(x=>x>0);
  const NORM=loads.length?(function(){const f={};loads.forEach(x=>f[x]=(f[x]||0)+1);return +Object.keys(f).sort((a,b)=>f[b]-f[a])[0];})():30;
  const subjTotals={}; cells.forEach(c=>{ subjTotals[c.subject_id]=(subjTotals[c.subject_id]||0)+1; });
  const requirement=Object.keys(subjById).map(s=>{ const total=subjTotals[s]||0; const available=tsub[s]?tsub[s].size:0; const required=total>0?Math.ceil(total/NORM):0; return { subject:subjById[s], total, norm:NORM, required, available, gap:required-available }; }).filter(x=>x.total>0||x.available>0).sort((a,b)=>b.total-a.total);
  return { school:{ name:cfg.school_name||sch.name||'', board:sch.board||'', medium:sch.medium||'', udise:sch.udise||'', district:sch.district||'', session:cfg.academic_session||'' }, norm:NORM, generated: cells.length>0, workload, requirement };
}
app.get('/api/reports/deo', h(async (req,res)=>{ res.json(await deoData(req.sid)); }));
const _XB={top:{style:'thin'},bottom:{style:'thin'},left:{style:'thin'},right:{style:'thin'}};
app.get('/api/export/teacher-workload.xlsx', h(async (req,res)=>{
  const d=await deoData(req.sid); const wb=new ExcelJS.Workbook(); wb.creator='Aumtara'; const ws=wb.addWorksheet('Teacher Workload');
  ws.columns=[{width:5},{width:22},{width:16},{width:16},{width:14},{width:26},{width:30},{width:9},{width:9},{width:15}];
  let r=1; const H=(txt,f)=>{const c=ws.getCell('A'+r);c.value=txt;if(f)c.font=f;ws.mergeCells('A'+r+':J'+r);r++;};
  H(d.school.name,{bold:true,size:14,color:{argb:'FF1F3864'}});
  H('TEACHER WORKLOAD STATEMENT'+(d.school.session?(' — '+d.school.session):''),{bold:true,size:12});
  H([d.school.udise&&('UDISE: '+d.school.udise),d.school.district&&('Dist/Taluka: '+d.school.district),d.school.board,d.school.medium].filter(Boolean).join('   ·   '),{size:10,color:{argb:'FF555555'}});
  r++;
  const hr=ws.getRow(r); hr.values=['Sr','Teacher Name','Designation','Qualification','Main Subject','Subjects taught','Subject-wise periods/week','Total/wk','Reqd','Surplus(+)/Deficit(−)'];
  hr.eachCell(c=>{c.font={bold:true,color:{argb:'FFFFFFFF'}};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF1F3864'}};c.alignment={horizontal:'center',wrapText:true,vertical:'middle'};c.border=_XB;}); r++;
  d.workload.forEach(w=>{ const row=ws.getRow(r); row.values=[w.sr,w.name,w.designation,w.qualification,w.main,w.subjects,w.subjDist,w.total,(w.required==null?'':w.required),(w.diff==null?'':(w.diff>0?'+'+w.diff:''+w.diff))];
    row.eachCell(c=>{c.border=_XB;c.alignment={vertical:'top',wrapText:true};}); r++; });
  r+=2; ws.getCell('A'+r).value='Date of issue: ______________'; ws.getCell('H'+r).value='Signature (Head Master)';
  res.setHeader('Content-Disposition','attachment; filename=teacher-workload.xlsx');
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  await wb.xlsx.write(res); res.end();
}));
app.get('/api/export/teacher-requirement.xlsx', h(async (req,res)=>{
  const d=await deoData(req.sid); const wb=new ExcelJS.Workbook(); wb.creator='Aumtara'; const ws=wb.addWorksheet('Teacher Requirement');
  ws.columns=[{width:5},{width:26},{width:16},{width:18},{width:16},{width:16},{width:18}];
  let r=1; const H=(txt,f)=>{const c=ws.getCell('A'+r);c.value=txt;if(f)c.font=f;ws.mergeCells('A'+r+':G'+r);r++;};
  H(d.school.name,{bold:true,size:14,color:{argb:'FF1F3864'}});
  H('TEACHER REQUIREMENT — STAFF PATTERN'+(d.school.session?(' — '+d.school.session):''),{bold:true,size:12});
  H('Norm used: '+d.norm+' periods per teacher per week'+(d.generated?'':'   (⚠ timetable not generated — totals may be 0)'),{size:10,color:{argb:'FF555555'}});
  r++;
  const hr=ws.getRow(r); hr.values=['Sr','Subject','Total periods/week','Periods/teacher (norm)','Teachers required','Teachers available','Shortfall(+)/Surplus(−)'];
  hr.eachCell(c=>{c.font={bold:true,color:{argb:'FFFFFFFF'}};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF1F3864'}};c.alignment={horizontal:'center',wrapText:true,vertical:'middle'};c.border=_XB;}); r++;
  d.requirement.forEach((x,i)=>{ const row=ws.getRow(r); row.values=[i+1,x.subject,x.total,x.norm,x.required,x.available,(x.gap>0?'+'+x.gap:''+x.gap)];
    row.eachCell(c=>{c.border=_XB;c.alignment={horizontal:'center'};}); row.getCell(2).alignment={horizontal:'left'}; r++; });
  res.setHeader('Content-Disposition','attachment; filename=teacher-requirement.xlsx');
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  await wb.xlsx.write(res); res.end();
}));

// ============================================================================
// DEO / GUJARAT PATRAK full report pack (Patrak A/B/K × Secondary/Higher-Sec + DEO Format 1-5)
// ============================================================================
function _stdOf(name){ const m=String(name||'').match(/\d+/); return m?+m[0]:null; }
function _sectionOf(std){ if(std!=null && std>=11) return 'hsec'; if(std!=null && std>=9) return 'sec'; if(std!=null && std>=1) return 'pri'; return 'sec'; }  // 1-8 Primary · 9-10 Secondary · 11-12 Higher-Sec · unnumbered → Secondary
async function deoPack(sid){
  const cfg=getConfig(sid)||{};
  const sch=(await q1('SELECT * FROM tt_school WHERE id=?',[sid]))||{};
  const teachers=await q('SELECT * FROM tt_teacher WHERE school_id=? ORDER BY name,id',[sid]);
  const subjects=await q('SELECT id,name FROM tt_subject WHERE school_id=? ORDER BY id',[sid]);
  const subjById={}; subjects.forEach(s=>subjById[s.id]=s.name);
  const classes=await q('SELECT * FROM tt_class WHERE school_id=? ORDER BY id',[sid]);
  const clsById={}; classes.forEach(c=>clsById[c.id]=c);
  const rooms=await q('SELECT * FROM tt_room WHERE school_id=? ORDER BY name,id',[sid]);
  const quotas=await q('SELECT * FROM tt_quota WHERE school_id=?',[sid]);
  const cells=await q('SELECT class_id,subject_id,teacher_id,room_id,day_of_week,period_index FROM tt_timetable WHERE school_id=? AND week_index=0',[sid]);
  const tchById={}; teachers.forEach(t=>tchById[t.id]=t);
  const header={ school_name:cfg.school_name||sch.name||'', address:sch.address||'', school_code:sch.school_code||'', board:sch.board||'', medium:sch.medium||'', udise:sch.udise||'', district:sch.district||'', email:sch.email||'', mobile:sch.mobile||'', session:cfg.academic_session||'', reg_code:sch.deo_reg_code||'', inspection_ref:sch.deo_inspection_ref||'', officer_name:sch.deo_officer_name||'' };
  const generated=cells.some(c=>c.subject_id);

  const secClasses={pri:[],sec:[],hsec:[]};
  classes.forEach(c=>{ const st=_stdOf(c.name); secClasses[_sectionOf(st)].push({id:c.id,name:c.name,std:st}); });
  function buildSection(key){
    const scls=secClasses[key]; if(!scls.length) return {key,present:false};
    const clsIds=new Set(scls.map(c=>c.id));
    const secCells=cells.filter(c=>clsIds.has(c.class_id) && c.subject_id);
    const subjSet=new Set();
    quotas.forEach(qq=>{ if(clsIds.has(qq.class_id) && +qq.per_week>0) subjSet.add(qq.subject_id); });
    secCells.forEach(c=>subjSet.add(c.subject_id));
    const subjList=subjects.filter(s=>subjSet.has(s.id)).map(s=>({id:s.id,name:s.name}));
    const stds=[...new Set(scls.map(c=>c.std==null?0:c.std))].sort((a,b)=>a-b);
    const patrakA={ standards:[], colTotal:{}, grand:0 };
    stds.forEach(st=>{
      const divs=scls.filter(c=>(c.std==null?0:c.std)===st); const divisions=divs.length; const divIds=new Set(divs.map(c=>c.id));
      const row={ std:st, divisions, cols:[], total:0 };
      subjList.forEach(su=>{
        const qv=quotas.filter(qq=>divIds.has(qq.class_id) && qq.subject_id===su.id).map(x=>+x.per_week||0).filter(x=>x>0);
        let tas;
        if(qv.length){ const f={}; qv.forEach(x=>f[x]=(f[x]||0)+1); tas=+Object.keys(f).sort((a,b)=>f[b]-f[a])[0]; }
        else { const tc=secCells.filter(c=>divIds.has(c.class_id) && c.subject_id===su.id).length; tas=divisions?Math.round(tc/divisions):0; }
        const total=tas*divisions;
        row.cols.push({subj:su.id, tas, varg:divisions, total}); row.total+=total;
        patrakA.colTotal[su.id]=(patrakA.colTotal[su.id]||0)+total; patrakA.grand+=total;
      });
      patrakA.standards.push(row);
    });
    const secTeacherIds=[...new Set(secCells.filter(c=>c.teacher_id).map(c=>c.teacher_id))];
    const patrakB={ teachers:[] }; const colB={};
    secTeacherIds.map(id=>tchById[id]).filter(Boolean).sort((a,b)=>String(a.name||'').localeCompare(String(b.name||''))).forEach(t=>{
      const per={}; let tot=0;
      subjList.forEach(su=>{ const n=secCells.filter(c=>c.teacher_id===t.id && c.subject_id===su.id).length; per[su.id]=n; tot+=n; if(n) colB[su.id]=(colB[su.id]||0)+n; });
      patrakB.teachers.push({ id:t.id, name:t.name||'', per, total:tot });
    });
    const patrakK={ rows: subjList.map(su=>{ const a=patrakA.colTotal[su.id]||0, b=colB[su.id]||0; return { subj:su.id, a, b, remain:a-b }; }) };
    return { key, present:true, subjects:subjList, patrakA, patrakB, patrakK,
             totals:{ A:patrakA.grand, B:subjList.reduce((s,su)=>s+(colB[su.id]||0),0) } };
  }
  const sections=[buildSection('pri'),buildSection('sec'),buildSection('hsec')];

  const maxLoad=+sch.deo_max_load||32;
  const format1=teachers.map((t,i)=>{ const total=cells.filter(c=>c.teacher_id===t.id && c.subject_id).length;
    const pct=maxLoad>0?Math.round(total/maxLoad*100):0; const status=total===0?'unassigned':(pct<60?'under':(pct>100?'over':'optimal'));
    return { sr:i+1, name:t.name||'', designation:t.designation||'', total, max:maxLoad, pct, status }; });

  const format2=[]; let f2sr=0;
  classes.forEach(c=>{ const cc=cells.filter(x=>x.class_id===c.id && x.subject_id);
    const cnt={}, tset={}; cc.forEach(x=>{ cnt[x.subject_id]=(cnt[x.subject_id]||0)+1; if(x.teacher_id)(tset[x.subject_id]=tset[x.subject_id]||new Set()).add(x.teacher_id); });
    const qs=quotas.filter(x=>x.class_id===c.id); const seen=new Set();
    const addRow=(subjId,mandated)=>{ const sched=cnt[subjId]||0; const teachs=tset[subjId]?[...tset[subjId]].map(id=>(tchById[id]||{}).name).filter(Boolean).join(', '):'';
      const pct=mandated>0?Math.round(sched/mandated*100):(sched>0?100:0); const status=mandated>0?(sched>=mandated?'compliant':'short'):(sched>0?'extra':'none');
      format2.push({ sr:++f2sr, cls:c.name, subject:subjById[subjId]||('#'+subjId), teacher:teachs||'—', mandated:mandated||0, scheduled:sched, pct, status }); };
    qs.forEach(qq=>{ seen.add(qq.subject_id); addRow(qq.subject_id, +qq.per_week||0); });
    Object.keys(cnt).forEach(k=>{ if(!seen.has(+k)) addRow(+k, 0); }); });

  let totalSlots=0; workingDaysArr(sid).forEach(di=>totalSlots+=teachingSlots(di,sid).length);
  const format3=rooms.map((rm,i)=>{ const used=cells.filter(x=>x.room_id===rm.id).length; const pct=totalSlots>0?Math.round(used/totalSlots*100):0;
    const status=pct===0?'unused':(pct<40?'under':(pct>90?'high':'optimal'));
    return { sr:i+1, room:rm.name||'', capacity:rm.capacity||'', used, slots:totalSlots, pct, status }; });

  const dmin=(a,b)=>{ try{const x=a.split(':').map(Number),y=b.split(':').map(Number);return (y[0]*60+y[1])-(x[0]*60+x[1]);}catch(e){return 0;} };
  const format4=[]; const wd=workingDaysArr(sid);
  const mkShift=(dayIdx,label)=>{ const slots=slotsForDay(dayIdx,sid); if(!slots.length)return; const teach=slots.filter(s=>!s.is_break);
    const instr=teach.reduce((s,x)=>s+dmin(x.start,x.end),0); const first=slots[0], last=slots[slots.length-1]; const dur=teach.length?Math.round(instr/teach.length):0;
    format4.push({ shift:label, board:[sch.board,sch.medium].filter(Boolean).join(' / ')||'—', timings:(first.start||'')+' – '+(last.end||''), periods:teach.length, duration:dur, dailyMin:instr }); };
  const wdFirst=wd.find(d=>d!==5); if(wdFirst!=null) mkShift(wdFirst,'weekdays'); if(wd.includes(5)) mkShift(5,'saturday');

  const ds=await q('SELECT * FROM tt_datesub WHERE school_id=? ORDER BY sub_date DESC, id DESC LIMIT 300',[sid]);
  const format5=ds.map((r,i)=>({ sr:i+1, id:'ABS-'+String(r.id).padStart(3,'0'), date:r.sub_date||'', absent:(tchById[r.absent_teacher_id]||{}).name||'—',
    cls:(clsById[r.class_id]||{}).name||'', period:(r.period_index!=null?('P'+(r.period_index+1)):''),
    sub:r.is_free?'(free)':((tchById[r.proxy_teacher_id]||{}).name||'—'), status:r.is_free?'free':(r.proxy_teacher_id?'assigned':'pending') }));

  return { header, generated, subjById, sections, format1, format2, format3, format4, format5, maxLoad };
}
app.get('/api/reports/deo-pack', h(async (req,res)=>{ res.json(await deoPack(req.sid)); }));

// ---- Excel export for any DEO format (fmt query), plus a combined pack ----
function _deoXlHead(ws, d, lastCol, title, subtitle){
  let r=1; const merge=(txt,font,center)=>{ const c=ws.getCell('A'+r); c.value=txt; if(font)c.font=font; c.alignment={horizontal:center?'center':'left'}; ws.mergeCells('A'+r+':'+lastCol+r); r++; };
  const H=d.header;
  merge(H.school_name||'', {bold:true,size:14,color:{argb:'FF1F3864'}}, true);
  if(H.address) merge(H.address, {size:9,color:{argb:'FF555555'}}, true);
  const meta=[H.school_code&&('Code: '+H.school_code),H.udise&&('UDISE: '+H.udise),H.district&&('Dist/Taluka: '+H.district),H.board,H.medium,H.session&&('A.Y. '+H.session)].filter(Boolean).join('   ·   ');
  if(meta) merge(meta, {size:9,color:{argb:'FF555555'}}, true);
  const contact=[H.mobile&&('Mob: '+H.mobile),H.email&&('Email: '+H.email)].filter(Boolean).join('   ·   ');
  if(contact) merge(contact, {size:9,color:{argb:'FF555555'}}, true);
  merge(title, {bold:true,size:12}, true); if(subtitle) merge(subtitle,{size:10,color:{argb:'FF444444'}},true);
  r++; return r;
}
function _deoXlTable(ws, startRow, headers, rows, widths){
  if(widths) ws.columns=widths.map(w=>({width:w}));
  let r=startRow; const hr=ws.getRow(r); hr.values=headers;
  hr.eachCell(c=>{c.font={bold:true,color:{argb:'FFFFFFFF'}};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF1F3864'}};c.alignment={horizontal:'center',wrapText:true,vertical:'middle'};c.border=_XB;}); r++;
  rows.forEach(row=>{ const rr=ws.getRow(r); rr.values=row; rr.eachCell(c=>{c.border=_XB;c.alignment={vertical:'middle',wrapText:true};}); r++; });
  return r;
}
function _deoSig(ws, r, officer){ r+=2; ws.getCell('A'+r).value='__________________________'; ws.getCell('E'+r).value='__________________________'; r++;
  ws.getCell('A'+r).value='Signature of School Principal'; ws.getCell('E'+r).value='Signature of District Education Officer'; r++;
  ws.getCell('A'+r).value='Institutional Stamp & Seal'; ws.getCell('E'+r).value='Government Department Verification Seal'; return r; }
function _secLabel(key){ return key==='hsec'?'Higher Secondary Section':(key==='pri'?'Primary Section':'Secondary Section'); }
function buildDeoSheet(wb, d, fmt){
  const secByKey={}; (d.sections||[]).forEach(s=>secByKey[s.key]=s);
  const sn=(/^patrak_[abk]_(pri|sec|hsec)$/.test(fmt))?fmt.split('_').pop():null;
  if(fmt.startsWith('patrak_a')){ const sec=secByKey[sn]||{present:false}; const ws=wb.addWorksheet('Patrak A '+(sn==='hsec'?'(HS)':sn==='pri'?'(Pri)':'(Sec)'));
    if(!sec.present){ ws.addRow(['No '+_secLabel(sn)+' classes.']); return; }
    const subs=sec.subjects; const start=_deoXlHead(ws,d,String.fromCharCode(66+subs.length),'કાર્યભાર પત્રક-અ · Class Allocation','('+_secLabel(sn)+')');
    const headers=['ધોરણ / Std', ...subs.map(s=>s.name), 'કુલ / Total']; const rows=[];
    sec.patrakA.standards.forEach(st=>{ rows.push(['Std '+(st.std||'-')+'  ·  Periods (તાસ)', ...st.cols.map(c=>c.tas), '']);
      rows.push(['   Divisions (વર્ગ)', ...st.cols.map(c=>c.varg), st.divisions]);
      rows.push(['   Total (કુલ તાસ)', ...st.cols.map(c=>c.total), st.total]); });
    rows.push(['GRAND TOTAL', ...subs.map(s=>sec.patrakA.colTotal[s.id]||0), sec.patrakA.grand]);
    const end=_deoXlTable(ws,start,headers,rows,[22,...subs.map(()=>9),10]); _deoSig(ws,end,d.header.officer_name); return; }
  if(fmt.startsWith('patrak_b')){ const sec=secByKey[sn]||{present:false}; const ws=wb.addWorksheet('Patrak B '+(sn==='hsec'?'(HS)':sn==='pri'?'(Pri)':'(Sec)'));
    if(!sec.present){ ws.addRow(['No '+_secLabel(sn)+' classes.']); return; }
    const subs=sec.subjects; const start=_deoXlHead(ws,d,String.fromCharCode(66+subs.length),'પત્રક-બ · Teacher Workload','('+_secLabel(sn)+')');
    const headers=['કર્મચારીનો ક્રમ / Teacher', ...subs.map(s=>s.name), 'કુલ / Total']; const rows=[];
    sec.patrakB.teachers.forEach((t,i)=>rows.push([(i+1)+'. '+t.name, ...subs.map(s=>t.per[s.id]||'-'), t.total]));
    const end=_deoXlTable(ws,start,headers,rows,[24,...subs.map(()=>9),10]); _deoSig(ws,end,d.header.officer_name); return; }
  if(fmt.startsWith('patrak_k')){ const sec=secByKey[sn]||{present:false}; const ws=wb.addWorksheet('Patrak K '+(sn==='hsec'?'(HS)':sn==='pri'?'(Pri)':'(Sec)'));
    if(!sec.present){ ws.addRow(['No '+_secLabel(sn)+' classes.']); return; }
    const subs=sec.subjects; const start=_deoXlHead(ws,d,String.fromCharCode(66+subs.length),'કાર્યભાર પત્રક-ક · Workload Certificate','('+_secLabel(sn)+')');
    const headers=['વિષય / Subject-wise', ...subs.map(s=>s.name)];
    const rows=[["પત્રક 'અ' મુજબ (as per A)", ...sec.patrakK.rows.map(r=>r.a)],
                ["પત્રક 'બ' મુજબ (as per B)", ...sec.patrakK.rows.map(r=>r.b)],
                ["બાકી રહેતું (remaining)", ...sec.patrakK.rows.map(r=>r.remain)]];
    const end=_deoXlTable(ws,start,headers,rows,[26,...subs.map(()=>9)]);
    let r=end+1; ws.getCell('A'+r).value='આથી પ્રમાણપત્ર આપવામાં આવે છે કે ઉપર્યુક્ત પત્રકમાં દર્શાવ્યા મુજબનો કાર્યભાર શાળાના સમયપત્રકમાં દર્શાવેલ છે.'; ws.mergeCells('A'+r+':'+String.fromCharCode(66+subs.length)+r); r++;
    _deoSig(ws,r,d.header.officer_name); return; }
  if(fmt==='format1'){ const ws=wb.addWorksheet('Format 1'); const start=_deoXlHead(ws,d,'G','FORMAT 1 — TEACHER WEEKLY WORKLOAD & DUTY SUMMARY','');
    const rows=d.format1.map(w=>[w.sr,w.name,w.designation,w.total,w.max,w.pct+'%',_deoStatusText(w.status)]);
    const end=_deoXlTable(ws,start,['S.No','Faculty Name','Designation / Cadre','Total Weekly Load','Max','Capacity %','Status'],rows,[6,26,20,14,10,11,16]); _deoSig(ws,end,d.header.officer_name); return; }
  if(fmt==='format2'){ const ws=wb.addWorksheet('Format 2'); const start=_deoXlHead(ws,d,'G','FORMAT 2 — CLASS SUBJECT PERIOD ALLOCATION & CURRICULUM COVERAGE','');
    const rows=d.format2.map(w=>[w.sr,w.cls,w.subject,w.teacher,w.mandated+' p/wk',w.scheduled+' p/wk',w.pct+'% '+_deoStatusText(w.status)]);
    const end=_deoXlTable(ws,start,['S.No','Class & Section','Subject','Assigned Faculty','Mandated','Scheduled','Compliance'],rows,[6,18,22,24,11,11,20]); _deoSig(ws,end,d.header.officer_name); return; }
  if(fmt==='format3'){ const ws=wb.addWorksheet('Format 3'); const start=_deoXlHead(ws,d,'G','FORMAT 3 — ROOM & LAB INFRASTRUCTURE UTILIZATION AUDIT','');
    const rows=d.format3.map(w=>[w.sr,w.room,w.capacity!==''?(w.capacity+' seats'):'—',w.used+' / '+w.slots+' slots',w.pct+'%',_deoStatusText(w.status)]);
    const end=_deoXlTable(ws,start,['S.No','Room / Name','Seating Capacity','Occupied Periods/Wk','Occupancy %','Audit Status'],rows,[6,24,16,20,13,18]); _deoSig(ws,end,d.header.officer_name); return; }
  if(fmt==='format4'){ const ws=wb.addWorksheet('Format 4'); const start=_deoXlHead(ws,d,'G','FORMAT 4 — BELL SCHEDULE & INSTRUCTIONAL MINUTES','');
    const rows=d.format4.map((w,i)=>[i+1,_deoStatusText(w.shift),w.board,w.timings,w.periods+' periods',w.duration+' min',w.dailyMin+' min ('+(w.dailyMin/60).toFixed(1)+' hrs)']);
    const end=_deoXlTable(ws,start,['S.No','Shift / Day','Board Affiliation','Timings','Periods/Day','Period Duration','Daily Instructional Time'],rows,[6,16,20,16,13,14,22]); _deoSig(ws,end,d.header.officer_name); return; }
  if(fmt==='format5'){ const ws=wb.addWorksheet('Format 5'); const start=_deoXlHead(ws,d,'G','FORMAT 5 — FACULTY ABSENCE & SUBSTITUTE DUTY REGISTER','');
    const rows=d.format5.map(w=>[w.sr,w.id+'  '+w.date,w.absent,(w.cls?(w.cls+' '):'')+w.period,w.sub,_deoStatusText(w.status)]);
    const end=_deoXlTable(ws,start,['S.No','Absence ID & Date','Absent Faculty','Class / Period','Substitute Faculty','Duty Status'],rows,[6,20,22,16,22,14]); _deoSig(ws,end,d.header.officer_name); return; }
}
function _deoStatusText(s){ return ({under:'Underloaded',over:'Overloaded',optimal:'Optimal',unassigned:'Unassigned',compliant:'Compliant',short:'Shortfall',extra:'Extra',none:'—',unused:'Unused',high:'High Demand',assigned:'Assigned',pending:'Pending',free:'Free Period',weekdays:'Weekdays',saturday:'Saturday'}[s]||s); }
app.get('/api/export/deo.xlsx', h(async (req,res)=>{
  const d=await deoPack(req.sid); const wb=new ExcelJS.Workbook(); wb.creator='Aumtara';
  const fmt=String(req.query.fmt||'pack');
  const present={}; (d.sections||[]).forEach(s=>present[s.key]=s.present);
  const ALL=['patrak_a_pri','patrak_a_sec','patrak_a_hsec','patrak_b_pri','patrak_b_sec','patrak_b_hsec','patrak_k_pri','patrak_k_sec','patrak_k_hsec','format1','format2','format3','format4','format5'];
  if(fmt==='pack') ALL.filter(f=>{ const m=f.match(/^patrak_[abk]_(pri|sec|hsec)$/); return m?present[m[1]]:true; }).forEach(f=>buildDeoSheet(wb,d,f)); else buildDeoSheet(wb,d,fmt);
  if(!wb.worksheets.length) wb.addWorksheet('DEO').addRow(['No data']);
  res.setHeader('Content-Disposition','attachment; filename=deo-'+fmt+'.xlsx');
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  await wb.xlsx.write(res); res.end();
}));
// Printable/exportable "Daily Bell Timings" notice — period start/end times computed from Academic Hours (weekday + Saturday)
app.get('/api/export/bell-timings.xlsx', h(async (req,res)=>{
  const sid=req.sid; const cfg=getConfig(sid)||{};
  const school=(await q1('SELECT name,board,medium FROM tt_school WHERE id=?',[sid]))||{};
  const clip=(s,n)=>String(s==null?'':s).slice(0,n);
  const title=clip(req.query.title,80)||'DAILY BELL TIMINGS';
  const wef=clip(req.query.wef,20), issue=clip(req.query.issue,20), sign=clip(req.query.sign,60)||'Principal', notes=clip(req.query.notes,1200);
  const mins=(a,b)=>{ try{ const x=a.split(':').map(Number),y=b.split(':').map(Number); return (y[0]*60+y[1])-(x[0]*60+x[1]); }catch(e){ return ''; } };
  const wb=new ExcelJS.Workbook(); wb.creator='Aumtara';
  const ws=wb.addWorksheet('Bell Timings'); ws.columns=[{width:24},{width:14},{width:12},{width:12}];
  let r=1; const bd={bottom:{style:'thin'},top:{style:'thin'},left:{style:'thin'},right:{style:'thin'}};
  const merge=(txt,font,center)=>{ const c=ws.getCell('A'+r); c.value=txt; if(font)c.font=font; c.alignment={horizontal:center?'center':'left'}; ws.mergeCells('A'+r+':D'+r); r++; };
  merge(cfg.school_name||school.name||'', {bold:true,size:14,color:{argb:'FF1F3864'}});
  if(school.board||school.medium) merge([school.board,school.medium].filter(Boolean).join(' · '), {size:10,color:{argb:'FF666666'}});
  merge(title, {bold:true,size:12}, true);
  if(wef) merge('With effect from: '+wef, {size:11});
  r++;
  const addTable=(dayIdx,label)=>{
    const slots=slotsForDay(dayIdx,sid); if(!slots.length)return;
    merge(label, {bold:true,color:{argb:'FF1F3864'}});
    const hr=ws.getRow(r); hr.values=['Period / Event','Duration','From','To'];
    hr.eachCell(c=>{ c.font={bold:true,color:{argb:'FFFFFFFF'}}; c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF1F3864'}}; c.alignment={horizontal:'center'}; c.border=bd; }); r++;
    slots.forEach(s=>{ const nm=s.is_break?(s.kind==='lunch'?'Lunch':'Break'):('Period '+(s.index+1)); const row=ws.getRow(r);
      row.values=[nm, mins(s.start,s.end)+' min', s.start, s.end];
      row.eachCell(c=>{ c.border=bd; c.alignment={horizontal:'center'}; }); row.getCell(1).alignment={horizontal:'left'};
      if(s.is_break) row.eachCell(c=>{ c.font={italic:true,color:{argb:'FF666666'}}; c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFF3F4F6'}}; });
      r++; });
    r++;
  };
  const wdFirst=workingDaysArr(sid).find(d=>d!==5);
  if(wdFirst!=null) addTable(wdFirst,'Weekdays');
  if(workingDaysArr(sid).includes(5)) addTable(5,'Saturday');
  if(notes){ merge('Notes:', {bold:true}); notes.split('\n').filter(x=>x.trim()).forEach(n=>merge('• '+n.trim())); r++; }
  merge('Date of issue: '+(issue||''));
  merge('Signature: '+sign);
  res.setHeader('Content-Disposition','attachment; filename=bell-timings.xlsx');
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  await wb.xlsx.write(res); res.end();
}));
app.get('/api/export/hours-template.xlsx', h(async (_,res)=>{
  const wb=new ExcelJS.Workbook();
  const mk=(name,rows)=>{ const w=wb.addWorksheet(name); w.addRow(['Label','Start','End']); styleHeader(w.getRow(1));
    rows.forEach(r=>w.addRow(r)); [16,12,12].forEach((wd,i)=>w.getColumn(i+1).width=wd);
    w.getColumn(2).numFmt='@'; w.getColumn(3).numFmt='@'; };
  mk('Weekday',[['Period 1','07:30','08:05'],['Period 2','08:05','08:40'],['Period 3','08:40','09:15'],['Lunch','09:15','09:40'],['Period 4','09:40','10:15'],['Period 5','10:15','10:50'],['Break','10:50','11:00'],['Period 6','11:00','11:35']]);
  mk('Saturday',[['Period 1','08:00','08:35'],['Period 2','08:35','09:10'],['Period 3','09:10','09:45'],['Period 4','09:45','10:20']]);
  const g=wb.addWorksheet('How to use'); g.addRow(['Fill the Weekday sheet (and Saturday if different). One row per period/break.']);
  g.addRow(['Column A = Label: write "Period 1", "Period 2"... for classes; "Lunch" for lunch; "Break" for a short break.']);
  g.addRow(['Column B/C = Start/End time as HH:MM (24-hour), e.g. 07:30 and 08:05.']);
  g.addRow(['The app auto-detects start, end, number of periods, period length, lunch and breaks.']);
  g.getColumn(1).width=95;
  res.setHeader('Content-Disposition','attachment; filename=hours-template.xlsx');
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  await wb.xlsx.write(res); res.end();
}));
app.post('/api/import/hours', upload.single('file'), async (req,res)=>{
  try{
    const sid=req.sid;
    const wb=new ExcelJS.Workbook(); await wb.xlsx.load(req.file.buffer);
    const findWs=(names)=>{ for(const w of wb.worksheets){ if(names.includes(w.name.trim().toLowerCase())) return w; } return null; };
    const wdWs=findWs(['weekday','weekdays','regular','mon-fri','mon–fri'])||wb.worksheets[0];
    const satWs=findWs(['saturday','saturdays','sat']);
    const wd=parseHoursSheet(wdWs);
    if(!wd) return res.status(400).json({ok:false,error:'No period rows found. Use the template: columns Label, Start, End; label rows Period 1, Period 2, Lunch, Break.'});
    await run(`UPDATE tt_config SET weekday_start=?,weekday_end=?,num_periods=?,period_minutes=?,period_durations=?,lunch_after=?,lunch_minutes=?,short_break_after=?,short_break_minutes=? WHERE school_id=?`,
      [wd.start,wd.end,wd.num,wd.period_minutes,wd.period_durations,wd.lunch_after,wd.lunch_minutes,wd.short_break_after,wd.short_break_minutes, sid]);
    const sat=(satWs&&satWs!==wdWs)?parseHoursSheet(satWs):null;
    if(sat){
      await run(`UPDATE tt_config SET saturday_start=?,saturday_end=?,sat_num_periods=?,sat_period_minutes=?,sat_period_durations=?,sat_lunch_after=?,sat_lunch_minutes=?,sat_short_break_after=?,sat_short_break_minutes=? WHERE school_id=?`,
        [sat.start,sat.end,sat.num,sat.period_minutes,sat.period_durations,sat.lunch_after,sat.lunch_minutes,sat.short_break_after,sat.short_break_minutes, sid]);
    }
    await loadConfig(sid);
    res.json({ok:true, weekday:wd, saturday:sat});
  }catch(e){ res.status(400).json({ok:false,error:String(e.message||e)}); }
});

// ---------- STARTUP ----------
async function ensurePayDefaults(){
  try{
    let s={}; try{ const raw=await getMeta('pay_settings'); s=raw?JSON.parse(raw):{}; }catch(_){ s={}; }
    let changed=false;
    if(!s.upi_id){ s.upi_id='spaumtaraent3107261120@spcb'; changed=true; }
    if(!s.payee_name || s.payee_name==='Aumtara Timetable'){ s.payee_name='AUMTARA ENTERPRISE'; changed=true; }
    if(s.enable_upi!==0 && s.enable_upi!==1){ s.enable_upi=1; changed=true; }
    if(changed){ await setMeta('pay_settings', JSON.stringify(s)); console.log('Seeded pay_settings UPI defaults'); }
  }catch(e){ console.error('ensurePayDefaults failed', e); }
}
// Break-glass owner recovery: if the env var OWNER_PW_RESET is set (only the account owner can set it in Render),
// reset the platform-owner login's password to that value on startup and clear any 2FA. Remove the env var afterwards.
async function ensureOwnerReset(){
  try{
    const pw=String(process.env.OWNER_PW_RESET||'').trim();
    if(pw.length<6) return;
    let owner=await q1('SELECT id,login_id FROM tt_user WHERE is_owner=1 ORDER BY id LIMIT 1');
    if(!owner) owner=await q1("SELECT id,login_id FROM tt_user WHERE role='master' ORDER BY id LIMIT 1");
    if(!owner){ console.log('OWNER_PW_RESET set but no owner/master account found'); return; }
    await run('UPDATE tt_user SET password_hash=?, active=1, totp_enabled=0, totp_secret=NULL, backup_codes=NULL WHERE id=?',[hashPw(pw), owner.id]);
    console.log('OWNER_PW_RESET applied to owner login "'+owner.login_id+'" (2FA cleared). Remove the OWNER_PW_RESET env var now.');
  }catch(e){ console.error('ensureOwnerReset failed', e); }
}
(async () => {
  await init();
  await ensureVapid();
  await ensurePayDefaults();
  await ensureOwnerReset();
  const PORT=process.env.PORT||4100;
  app.listen(PORT,()=>console.log(`Timetable module running → http://localhost:${PORT}`));
})().catch(e=>{ console.error('Startup failed:', e); process.exit(1); });
