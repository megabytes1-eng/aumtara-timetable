// Timetable module — REST API server v3 (Express + Postgres)
const express = require('express');
const path = require('path');
const ExcelJS = require('exceljs');
const multer = require('multer');
const upload = multer({ storage: multer.memoryStorage() });
const crypto = require('crypto');
const { q, q1, run, tx, init, loadConfig, getConfigCached, hashPw, verifyPw, seedConfigForSchool } = require('./db');

const app = express();
app.use(express.json({ limit: '8mb' }));   // logos (base64 data-URLs) can exceed the 100kb default
// Serve the single-page frontend. index.html lives next to server.js (flat layout
// so the repo uploads cleanly to GitHub's web uploader, which can't preserve subfolders).
// index.html is fully self-contained (inline CSS/JS), so no other static assets are needed.
app.get('/', (_, res) => res.sendFile(path.join(__dirname, 'index.html')));

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
const SW_JS = `const CACHE='aumtara-v2';
const SHELL=['/','/manifest.webmanifest','/icon-192.png','/icon-512.png'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(CACHE).then(c=>c.addAll(SHELL)).then(()=>self.skipWaiting()).catch(()=>self.skipWaiting()));});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()));});
self.addEventListener('fetch',e=>{const req=e.request; if(req.method!=='GET') return; let url; try{url=new URL(req.url);}catch(_){return;} if(url.origin!==self.location.origin) return; if(url.pathname.startsWith('/api/')) return;
  if(req.mode==='navigate'){ e.respondWith(fetch(req).then(r=>{const cp=r.clone(); caches.open(CACHE).then(c=>c.put('/',cp)); return r;}).catch(()=>caches.match('/'))); return; }
  e.respondWith(caches.match(req).then(r=>r||fetch(req)));});`;
app.get('/manifest.webmanifest', (_, res) => { res.type('application/manifest+json'); res.send(JSON.stringify(MANIFEST)); });
app.get('/sw.js', (_, res) => { res.type('application/javascript'); res.set('Cache-Control', 'no-cache'); res.send(SW_JS); });
for (const ic of ['icon-192.png', 'icon-512.png', 'icon-maskable.png'])
  app.get('/' + ic, (_, res) => res.sendFile(path.join(__dirname, ic)));
app.get('/apple-touch-icon.png', (_, res) => res.sendFile(path.join(__dirname, 'icon-192.png')));
app.get('/apple-touch-icon-precomposed.png', (_, res) => res.sendFile(path.join(__dirname, 'icon-192.png')));

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
  return await q1('SELECT id,name,role,login_id,email,mobile,qualification,main_subject_id,school_id FROM tt_user WHERE id=? AND active=1',[s.user_id]);
}
// public: sign in — identifier can be login_id OR email OR mobile
app.post('/api/login', h(async (req,res)=>{
  const b=req.body||{};
  const idf=String(b.login_id||'').trim();
  const u=await q1('SELECT * FROM tt_user WHERE (lower(login_id)=lower(?) OR lower(email)=lower(?) OR mobile=?) AND active=1 LIMIT 1',[idf,idf,idf]);
  if(!u || !verifyPw(b.password, u.password_hash)){ res.status(401).json({error:'invalid credentials'}); return; }
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
  // each signup creates its own school + config + admin account for it
  const sch=await q1('INSERT INTO tt_school(name,board,medium,active,created_at) VALUES(?,?,?,1,now()::text) RETURNING id',[school,board,medium]);
  await seedConfigForSchool(sch.id, school, board, medium);
  const row=await q1(`INSERT INTO tt_user(name,role,login_id,email,mobile,password_hash,active,created_at,school_id)
     VALUES(?,?,?,?,?,?,1,now()::text,?) RETURNING id`, [b.name||username,'admin',username,email,mobile,hashPw(password),sch.id]);
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
    else sid = req.user.school_id || null;
    if (sid == null) { const s = await q1('SELECT id FROM tt_school WHERE active=1 ORDER BY id LIMIT 1'); sid = s ? s.id : null; }
    req.sid = sid;
    next();
  })().catch(e=>{ console.error(e); res.status(500).json({error:String((e&&e.message)||e)}); });
});
app.get('/api/me', h(async (req,res)=>res.json(req.user)));

// -------------------- SCHOOLS (registry) --------------------
app.get('/api/schools', h(async (req,res)=>{
  if (req.user.role === 'master') res.json(await q('SELECT * FROM tt_school ORDER BY id'));
  else res.json(await q('SELECT * FROM tt_school WHERE id=? ORDER BY id',[req.user.school_id]));
}));
app.post('/api/schools', h(async (req,res)=>{
  if(!requireAdmin(req,res)) return;
  const b=req.body||{}; const name=String(b.name||'').trim();
  if(!name){ res.status(400).json({error:'school name required'}); return; }
  const row=await q1('INSERT INTO tt_school(name,board,medium,active,created_at) VALUES(?,?,?,1,now()::text) RETURNING id',
    [name, b.board||null, b.medium||null]);
  await seedConfigForSchool(row.id, name, b.board||null, b.medium||null);
  res.json({ id:row.id });
}));
app.put('/api/schools/:id', h(async (req,res)=>{
  if(!requireAdmin(req,res)) return;
  const b=req.body||{}, id=req.params.id;
  const cols=['name','board','medium','active','logo'];
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
  for(const tbl of ['tt_timetable','tt_quota','tt_chapter','tt_absence','tt_substitution','tt_diary','tt_snapshot','tt_class','tt_subject','tt_room','tt_teacher','tt_config'])
    await run(`DELETE FROM ${tbl} WHERE school_id=?`,[id]);
  // detach users of that school (don't delete accounts)
  await run('UPDATE tt_user SET school_id=NULL WHERE school_id=?',[id]);
  await run('DELETE FROM tt_school WHERE id=?',[id]);
  await loadConfig();
  res.json({ ok:true });
}));

// ---- USER MANAGEMENT (create/update/delete restricted to admin & master) ----
const ROLES=['master','admin','principal','supervisor','teacher'];
function requireAdmin(req,res){ if(!['admin','master'].includes(req.user.role)){ res.status(403).json({error:'forbidden'}); return false; } return true; }
app.get('/api/users', h(async (req,res)=>{
  if(!requireAdmin(req,res)) return;
  if(req.user.role==='master') res.json(await q('SELECT id,name,role,login_id,email,mobile,qualification,main_subject_id,active,created_at,school_id FROM tt_user ORDER BY id'));
  else res.json(await q('SELECT id,name,role,login_id,email,mobile,qualification,main_subject_id,active,created_at,school_id FROM tt_user WHERE school_id=? ORDER BY id',[req.sid]));
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
  const row=await q1(`INSERT INTO tt_user(name,role,login_id,email,mobile,qualification,main_subject_id,password_hash,active,created_at,school_id)
     VALUES(?,?,?,?,?,?,?,?,1,now()::text,?) RETURNING id`,
    [b.name||login, b.role||'teacher', login, b.email||null, b.mobile||null, b.qualification||null, b.main_subject_id||null, hashPw(pw), school]);
  res.json({ id:row.id });
}));
app.put('/api/users/:id', h(async (req,res)=>{
  if(!requireAdmin(req,res)) return;
  const b=req.body||{}, id=req.params.id;
  if(req.user.role!=='master'){ const tgt=await q1('SELECT school_id FROM tt_user WHERE id=?',[id]); if(!tgt||tgt.school_id!==req.sid){ res.status(403).json({error:'forbidden'}); return; } }
  if(b.role!==undefined && !ROLES.includes(b.role)){ res.status(400).json({error:'bad role'}); return; }
  const cols=['name','role','email','mobile','qualification','main_subject_id','active'];
  const set=cols.filter(c=>b[c]!==undefined);
  if(set.length) await run(`UPDATE tt_user SET ${set.map(c=>c+'=?').join(',')} WHERE id=?`,[...set.map(c=>b[c]===''?null:b[c]), id]);
  if(b.login_id){ const login=String(b.login_id).trim(); if(login && !(await q1('SELECT 1 FROM tt_user WHERE lower(login_id)=lower(?) AND id<>?',[login,id]))) await run('UPDATE tt_user SET login_id=? WHERE id=?',[login,id]); }
  if(b.password!=null && String(b.password).length) await run('UPDATE tt_user SET password_hash=? WHERE id=?',[hashPw(b.password), id]);
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

const DAYS = ['Mon','Tue','Wed','Thu','Fri','Sat'];
const addMin = (hhmm,min)=>{const[h,m]=hhmm.split(':').map(Number);const d=new Date(2020,0,1,h,m+min);return String(d.getHours()).padStart(2,'0')+':'+String(d.getMinutes()).padStart(2,'0');};
const getConfig = (sid) => getConfigCached(sid);
const csvNums = s => String(s||'').split(',').map(x=>parseInt(x,10)).filter(n=>!isNaN(n));
function workingDaysArr(sid){ const c=getConfig(sid)||{}; const w=csvNums(c.working_days); return w.length?w:[0,1,2,3,4,5]; }
function slotsForDay(dayIdx, sid){
  const c=getConfig(sid);
  if(!c) return [];
  if(!workingDaysArr(sid).includes(dayIdx)) return [];   // non-working day → no periods
  const sat=dayIdx===5;
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
simpleCrud('classes','tt_class',['name','class_teacher_id','board','medium','standard','section']);
simpleCrud('subjects','tt_subject',['name','active']);
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
  const hours    = !!(c.weekday_start && c.weekday_end && c.period_minutes);
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
  const map=await q('SELECT ts.* FROM tt_teacher_subject ts JOIN tt_teacher t ON t.id=ts.teacher_id WHERE t.school_id=?',[req.sid]);
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
  if(b.subjects!==undefined){ const subs=new Set(b.subjects); if(b.main_subject_id)subs.add(+b.main_subject_id); await setSubjects(+req.params.id, [...subs]); }
  res.json({ok:true});
}));
app.delete('/api/teachers/:id', h(async (req,res)=>{
  await run('DELETE FROM tt_teacher WHERE id=? AND school_id=?',[req.params.id, req.sid]);
  await run('DELETE FROM tt_teacher_subject WHERE teacher_id=?',[req.params.id]);
  res.json({ok:true});
}));
async function setSubjects(tid, subs){
  await run('DELETE FROM tt_teacher_subject WHERE teacher_id=?',[tid]);
  for(const s of subs) await run('INSERT INTO tt_teacher_subject VALUES (?,?) ON CONFLICT DO NOTHING',[tid, s]);
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
  await run(`INSERT INTO tt_quota(class_id,subject_id,per_week,school_id) VALUES(?,?,?,?)
     ON CONFLICT(class_id,subject_id) DO UPDATE SET per_week=excluded.per_week`,[class_id,subject_id,per_week, req.sid]);
  res.json({ok:true});
}));

// ---------- GRID ----------
app.get('/api/timetable', h(async (req,res)=>{
  const rows=req.query.class_id
    ? await q('SELECT * FROM tt_timetable WHERE class_id=? AND school_id=?',[req.query.class_id, req.sid])
    : await q('SELECT * FROM tt_timetable WHERE school_id=?',[req.sid]);
  res.json(rows);
}));

// ---------- CONFLICTS (teacher + room) ----------
async function conflictFor(class_id, day, period, teacher_id, room_id, sid){
  const out={};
  if(teacher_id){ const r=await q1(`SELECT c.name FROM tt_timetable tt JOIN tt_class c ON c.id=tt.class_id
     WHERE tt.day_of_week=? AND tt.period_index=? AND tt.teacher_id=? AND tt.class_id<>? AND tt.school_id=?`,[day,period,teacher_id,class_id,sid]);
    if(r) out.teacher=r.name; }
  if(room_id){ const r=await q1(`SELECT c.name FROM tt_timetable tt JOIN tt_class c ON c.id=tt.class_id
     WHERE tt.day_of_week=? AND tt.period_index=? AND tt.room_id=? AND tt.class_id<>? AND tt.school_id=?`,[day,period,room_id,class_id,sid]);
    if(r) out.room=r.name; }
  return out;
}
app.get('/api/timetable/conflicts', h(async (req,res)=>{
  const teacher=await q(`SELECT day_of_week,period_index,teacher_id,string_agg(class_id::text,',') classes,COUNT(*)::int n FROM tt_timetable
     WHERE teacher_id IS NOT NULL AND school_id=? GROUP BY day_of_week,period_index,teacher_id HAVING COUNT(*)>1`,[req.sid]);
  const room=await q(`SELECT day_of_week,period_index,room_id,string_agg(class_id::text,',') classes,COUNT(*)::int n FROM tt_timetable
     WHERE room_id IS NOT NULL AND school_id=? GROUP BY day_of_week,period_index,room_id HAVING COUNT(*)>1`,[req.sid]);
  res.json({teacher, room});
}));

// ---------- CELL upsert / delete ----------
app.put('/api/timetable/cell', h(async (req,res)=>{
  const {class_id,day,period,subject_id,teacher_id,room_id}=req.body;
  await run(`INSERT INTO tt_timetable(class_id,day_of_week,period_index,subject_id,teacher_id,room_id,school_id) VALUES(?,?,?,?,?,?,?)
     ON CONFLICT(class_id,day_of_week,period_index) DO UPDATE SET subject_id=excluded.subject_id,teacher_id=excluded.teacher_id,room_id=excluded.room_id`,
    [class_id,day,period,subject_id||null,teacher_id||null,room_id||null, req.sid]);
  res.json({ok:true, conflict: await conflictFor(class_id,day,period,teacher_id,room_id, req.sid)});
}));
app.delete('/api/timetable/cell', h(async (req,res)=>{
  await run('DELETE FROM tt_timetable WHERE class_id=? AND day_of_week=? AND period_index=? AND school_id=?',[req.body.class_id,req.body.day,req.body.period, req.sid]);
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

// ---------- AUTO-GENERATE (quota-aware, room-aware) ----------
app.post('/api/timetable/auto-generate', h(async (req,res)=>{
  const sid=req.sid;
  const classes=await q('SELECT * FROM tt_class WHERE school_id=? ORDER BY id',[sid]);
  const subjects=await q('SELECT * FROM tt_subject WHERE active=1 AND school_id=? ORDER BY id',[sid]);
  const activeSet=new Set(subjects.map(s=>s.id));
  const rooms=await q('SELECT * FROM tt_room WHERE school_id=? ORDER BY id',[sid]);
  const tmap=await q('SELECT ts.* FROM tt_teacher_subject ts JOIN tt_teacher t ON t.id=ts.teacher_id WHERE t.school_id=?',[sid]);
  const quotas=await q('SELECT * FROM tt_quota WHERE school_id=?',[sid]);
  const absent={}; (await q('SELECT * FROM tt_absence WHERE school_id=?',[sid])).forEach(a=>{(absent[a.teacher_id]=absent[a.teacher_id]||new Set()).add(a.day_of_week);});
  const maxLoad={}, capDay={}, capCons={};
  (await q('SELECT id,max_load,max_per_day,max_consecutive FROM tt_teacher WHERE school_id=?',[sid])).forEach(t=>{
    maxLoad[t.id]=t.max_load||999;
    capDay[t.id]=(t.max_per_day&&t.max_per_day>0)?t.max_per_day:Infinity;
    capCons[t.id]=(t.max_consecutive&&t.max_consecutive>0)?t.max_consecutive:Infinity;
  });
  const load={}, dayCount={}, lastPi={}, consRun={};   // per-teacher weekly/day counts + consecutive tracking
  const predRun=(t,di,pi)=> (lastPi[t+'_'+di]===pi-1 ? (consRun[t+'_'+di]||0)+1 : 1);   // consecutive run if placed at (di,pi)
  const capOk=(t,di,pi)=> (dayCount[t+'_'+di]||0)<capDay[t] && predRun(t,di,pi)<=capCons[t];
  const teachersForSubject=sid=>tmap.filter(m=>m.subject_id===sid).map(m=>m.teacher_id);

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

  // build per-class subject pool honouring quota (fallback: even rotation)
  function poolFor(cid, totalSlots){
    const qs=quotas.filter(x=>x.class_id===cid && activeSet.has(x.subject_id));
    let pool=[];
    if(qs.length){ qs.forEach(x=>{ for(let i=0;i<x.per_week;i++) pool.push(x.subject_id); }); }
    if(!subjects.length) return [];   // no active subjects → nothing to schedule
    if(pool.length<totalSlots){ let i=0; while(pool.length<totalSlots){ pool.push(subjects[i%subjects.length].id); i++; } }
    return pool.slice(0,totalSlots);
  }

  // compute the assignment in memory, then persist in one transaction
  let totalSlots=0; DAYS.forEach((_,di)=>totalSlots+=teachingSlots(di, sid).length);
  const remaining={}; classes.forEach(c=>{ remaining[c.id]=shuffleStable(poolFor(c.id,totalSlots),c.id).slice(); });
  const daySubs={}, prevSub={};   // per class+day: subjects already placed, and the previous period's subject
  function pickSubject(cid,di,pi){
    const rem=remaining[cid]; if(!rem.length) return null;
    const R=clsRules[cid]; const ds=daySubs[cid+'_'+di]; const prev=prevSub[cid+'_'+di];
    let idx=-1;
    for(let i=0;i<rem.length;i++){ const s=rem[i]; let ok=true;
      if(ds){ for(const x of ds){ if(R.sameDay.has(rkey(s,x))){ ok=false; break; } } }
      if(ok && prev!=null && R.notAdj.has(rkey(s,prev))) ok=false;
      if(ok){ idx=i; break; }
    }
    if(idx<0) idx=0;   // nothing satisfies the rules → take next to avoid leaving a gap
    const s=rem.splice(idx,1)[0];
    (daySubs[cid+'_'+di]=daySubs[cid+'_'+di]||new Set()).add(s);
    prevSub[cid+'_'+di]=s;
    return s;
  }
  const toInsert=[];
  DAYS.forEach((_,di)=>{
    const slots=teachingSlots(di, sid);
    slots.forEach((_,pi)=>{
      const usedT=new Set(), usedR=new Set();
      classes.forEach(c=>{
        const subjId=pickSubject(c.id,di,pi);
        if(subjId==null) return;   // class has no schedulable subjects
        let opts=teachersForSubject(subjId).filter(t=>!usedT.has(t)&&!(absent[t]&&absent[t].has(di))&&(load[t]||0)<maxLoad[t]&&capOk(t,di,pi));
        opts.sort((a,b)=>(load[a]||0)-(load[b]||0));
        let t=opts[0] ?? teachersForSubject(subjId).find(x=>!usedT.has(x)&&(load[x]||0)<maxLoad[x]&&capOk(x,di,pi)) ?? null;
        const room=rooms.find(r=>!usedR.has(r.id));
        toInsert.push([c.id,di,pi,subjId,t,room?room.id:null,sid]);
        if(t){usedT.add(t);load[t]=(load[t]||0)+1;
          dayCount[t+'_'+di]=(dayCount[t+'_'+di]||0)+1;
          consRun[t+'_'+di]=(lastPi[t+'_'+di]===pi-1?(consRun[t+'_'+di]||0)+1:1);
          lastPi[t+'_'+di]=pi;} if(room)usedR.add(room.id);
      });
    });
  });

  await tx(async (cq)=>{
    await cq('DELETE FROM tt_timetable WHERE school_id=?',[sid]);
    for(const r of toInsert)
      await cq('INSERT INTO tt_timetable(class_id,day_of_week,period_index,subject_id,teacher_id,room_id,school_id) VALUES(?,?,?,?,?,?,?)', r);
  });
  const n=(await q1('SELECT COUNT(*)::int AS n FROM tt_timetable WHERE school_id=?',[sid])).n;
  res.json({ok:true, cells:n});
}));
function shuffleStable(arr, seed){ // deterministic light shuffle so subjects spread out
  const a=arr.slice(); let s=seed*9301+49297;
  for(let i=a.length-1;i>0;i--){ s=(s*9301+49297)%233280; const j=Math.floor(s/233280*(i+1)); [a[i],a[j]]=[a[j],a[i]]; }
  return a;
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
  const cells=await q('SELECT class_id,day_of_week,period_index,subject_id,teacher_id,room_id FROM tt_timetable WHERE school_id=?',[sid]);
  const snap=await tx(async (cq,cq1)=>{
    const s=await cq1('INSERT INTO tt_snapshot(name,session,created_at,cell_count,school_id,term_id,config_json) VALUES(?,?,now()::text,?,?,?,?) RETURNING id',[name,session,cells.length,sid,(term&&term.id)||null,config_json]);
    for(const c of cells)
      await cq('INSERT INTO tt_snapshot_cell(snapshot_id,class_id,day_of_week,period_index,subject_id,teacher_id,room_id) VALUES(?,?,?,?,?,?,?)',
        [s.id,c.class_id,c.day_of_week,c.period_index,c.subject_id,c.teacher_id,c.room_id]);
    return s;
  });
  res.json({ok:true, id:snap.id, name, cells:cells.length});
}));
// restore a saved version → overwrites the live timetable AND its School Hours (window switch), and re-activates its term
app.post('/api/versions/:id/restore', h(async (req,res)=>{
  const id=Number(req.params.id), sid=req.sid;
  const snap=await q1('SELECT * FROM tt_snapshot WHERE id=? AND school_id=?',[id,sid]);
  if(!snap){ res.status(404).json({error:'not found'}); return; }
  const cells=await q('SELECT class_id,day_of_week,period_index,subject_id,teacher_id,room_id FROM tt_snapshot_cell WHERE snapshot_id=?',[id]);
  await tx(async (cq)=>{
    await cq('DELETE FROM tt_timetable WHERE school_id=?',[sid]);
    for(const c of cells)
      await cq('INSERT INTO tt_timetable(class_id,day_of_week,period_index,subject_id,teacher_id,room_id,school_id) VALUES(?,?,?,?,?,?,?)',
        [c.class_id,c.day_of_week,c.period_index,c.subject_id,c.teacher_id,c.room_id,sid]);
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

// ---------- TEACHER schedule ----------
app.get('/api/timetable/teacher/:id', h(async (req,res)=>{
  const id=Number(req.params.id);
  const rows=await q(`SELECT tt.*, c.name cls, s.name subj FROM tt_timetable tt
     JOIN tt_class c ON c.id=tt.class_id JOIN tt_subject s ON s.id=tt.subject_id WHERE tt.teacher_id=? AND tt.school_id=?`,[id, req.sid]);
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
     WHERE tt.day_of_week=? AND tt.teacher_id IS NOT NULL AND tt.school_id=?`,[day,sid]);
  const allTeachers=await q('SELECT id,name FROM tt_teacher WHERE school_id=?',[sid]);
  const need=[];
  for(const r of rows){
    if(!absentSet.has(r.teacher_id)) continue;
    const sub=await q1('SELECT proxy_teacher_id FROM tt_substitution WHERE day_of_week=? AND class_id=? AND period_index=?',[day,r.class_id,r.period_index]);
    const busy=new Set((await q('SELECT teacher_id FROM tt_timetable WHERE day_of_week=? AND period_index=? AND teacher_id IS NOT NULL AND school_id=?',[day,r.period_index,sid])).map(x=>x.teacher_id));
    const free=allTeachers.filter(t=>!busy.has(t.id)&&!absentSet.has(t.id));
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
  const allTeachers=await q('SELECT id,name FROM tt_teacher WHERE school_id=? ORDER BY name',[sid]);
  const cells=await q('SELECT day_of_week,period_index,teacher_id FROM tt_timetable WHERE teacher_id IS NOT NULL AND school_id=?',[sid]);
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
      const free=allTeachers.filter(t=>t.id!==teacher_id && !busySet.has(t.id));
      const e=exMap[date+'_'+c.class_id+'_'+c.period_index]; const sl=slots[c.period_index];
      out.push({ sub_date:date, dow, period_index:c.period_index, class_id:c.class_id, cls:c.cls, subject_id:c.subject_id, subj:c.subj||'',
        start:sl?sl.start:'', end:sl?sl.end:'', free:free.map(f=>({id:f.id,name:f.name})),
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
  const r=await q1('INSERT INTO tt_leave(school_id,teacher_id,date_from,date_to,reason,status,created_at) VALUES(?,?,?,?,?,?,now()::text) RETURNING id',[req.sid,teacher_id,b.date_from,b.date_to,b.reason||null,'pending']);
  res.json({ok:true,id:r.id});
}));
app.post('/api/leaves/:id/approve', h(async (req,res)=>{
  if(!isAdminRole(req)){ res.status(403).json({error:'forbidden'}); return; }
  const id=req.params.id, sid=req.sid;
  const lv=await q1('SELECT * FROM tt_leave WHERE id=? AND school_id=?',[id,sid]); if(!lv){ res.status(404).json({error:'not found'}); return; }
  await run('UPDATE tt_leave SET status=? WHERE id=? AND school_id=?',['approved',id,sid]);
  const wdays=new Set(workingDaysArr(sid));
  const myCells=await q('SELECT day_of_week,period_index,class_id FROM tt_timetable WHERE teacher_id=? AND school_id=?',[lv.teacher_id,sid]);
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
        const cell=await q1('SELECT * FROM tt_timetable WHERE class_id=? AND day_of_week=? AND period_index=? AND school_id=?',[c.id,day,pi,sid]);
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
  const w=['d.school_id=?'],p=[req.sid]; if(req.query.teacher_id){w.push('d.teacher_id=?');p.push(req.query.teacher_id);} if(req.query.date){w.push('d.entry_date=?');p.push(req.query.date);} sql+=' WHERE '+w.join(' AND ');
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
  const cellMap={}; (await q('SELECT * FROM tt_timetable WHERE school_id=?',[sid])).forEach(c=>cellMap[`${c.class_id}|${c.day_of_week}|${c.period_index}`]=c);
  const cellText=(c)=>{ if(!c||!c.subject_id)return ''; const s=subjById[c.subject_id]||''; const t=c.teacher_id?(tchById[c.teacher_id]||''):''; const r=c.room_id?(roomById[c.room_id]||''):''; return s+(t?'\n'+t:'')+(r?'\n'+r:''); };
  const maxSlots=Math.max(1,...DAYS.map((_,i)=>teachingSlots(i, sid).length));
  classes.forEach(c=>{
    const ws=wb.addWorksheet(c.name.slice(0,28).replace(/[\\\/\?\*\[\]:]/g,' '));
    const hdr=['Day',...Array.from({length:maxSlots},(_,i)=>'P'+(i+1))]; ws.addRow(hdr); styleHeader(ws.getRow(1));
    DAYS.forEach((d,di)=>{ const row=[d]; for(let p=0;p<maxSlots;p++){ row.push(cellText(cellMap[`${c.id}|${di}|${p}`])); }
      const r=ws.addRow(row); r.eachCell(cc=>{cc.alignment={wrapText:true,vertical:'top'};}); r.getCell(1).font={bold:true}; });
    ws.columns.forEach((col,i)=>{col.width=i===0?10:16;}); ws.getColumn(1).font={bold:true};
  });
  res.setHeader('Content-Disposition','attachment; filename=timetable.xlsx');
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
  const tt      =await q('SELECT * FROM tt_timetable WHERE school_id=?',[sid]);
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
  const tt      =await q('SELECT * FROM tt_timetable WHERE school_id=?',[sid]);
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
        if(mainId) await cq('INSERT INTO tt_teacher_subject VALUES(?,?) ON CONFLICT DO NOTHING',[tid, mainId]);
        for(const sn of optNames){ await cq('INSERT INTO tt_teacher_subject VALUES(?,?) ON CONFLICT DO NOTHING',[tid, await subId(sn)]); }
      }

      let ci=0;
      for(const row of chapterRows){ ci++; const sn=val(row.getCell(1)), cn=val(row.getCell(2)); if(!sn||!cn)continue;
        const subjid=await subId(sn);
        if(!(await cq1('SELECT 1 FROM tt_chapter WHERE subject_id=? AND lower(name)=lower(?) AND school_id=?',[subjid,cn,sid]))){ await cq('INSERT INTO tt_chapter(subject_id,name,seq,school_id) VALUES(?,?,?,?)',[subjid,cn,ci,sid]); added.chapters++; } }
    });
    res.json({ok:true, added});
  }catch(e){ res.status(400).json({ok:false,error:String(e.message||e)}); }
});

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
(async () => {
  await init();
  const PORT=process.env.PORT||4100;
  app.listen(PORT,()=>console.log(`Timetable module running → http://localhost:${PORT}`));
})().catch(e=>{ console.error('Startup failed:', e); process.exit(1); });
