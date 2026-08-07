// Timetable module — REST API server v3 (Express + Postgres)
const express = require('express');
const path = require('path');
const ExcelJS = require('exceljs');
const multer = require('multer');
const upload = multer({ storage: multer.memoryStorage() });
const crypto = require('crypto');
const { q, q1, run, tx, init, loadConfig, getConfigCached, hashPw, verifyPw, seedConfigForSchool } = require('./db');

const app = express();
app.use(express.json());
// Serve the single-page frontend. index.html lives next to server.js (flat layout
// so the repo uploads cleanly to GitHub's web uploader, which can't preserve subfolders).
// index.html is fully self-contained (inline CSS/JS), so no other static assets are needed.
app.get('/', (_, res) => res.sendFile(path.join(__dirname, 'index.html')));

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
  const end=dayIdx===5?c.saturday_end:c.weekday_end;
  const durs=csvNums(c.period_durations);
  const shortAfter=new Set(csvNums(c.short_break_after));
  const shortMin=+c.short_break_minutes||0;
  const lunchAfter=(c.lunch_after!=null?+c.lunch_after:+c.break_after_period);
  const lunchMin=(c.lunch_minutes!=null?+c.lunch_minutes:+c.break_minutes)||0;
  const slots=[]; let t=c.weekday_start, idx=0;
  while(true){ const dur=durs[idx]||c.period_minutes; const e=addMin(t,dur);
    slots.push({index:idx,label:'P'+(idx+1),start:t,end:e,is_break:false}); idx++;
    if(e>=end||idx>=12)break; t=e;
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
  await run(`UPDATE tt_config SET weekday_start=?,weekday_end=?,saturday_end=?,period_minutes=?,break_after_period=?,break_minutes=?,school_name=?,
       short_break_minutes=?,short_break_after=?,lunch_minutes=?,lunch_after=?,period_durations=?,working_days=?,academic_session=? WHERE school_id=?`,
    [v('weekday_start'),v('weekday_end'),v('saturday_end'),v('period_minutes'),v('break_after_period'),v('break_minutes'),v('school_name'),
     v('short_break_minutes'),v('short_break_after'),v('lunch_minutes'),v('lunch_after'),v('period_durations'),v('working_days'),v('academic_session'), req.sid]);
  // keep the school registry name in sync if the school name was edited here
  if(b.school_name!==undefined) await run('UPDATE tt_school SET name=? WHERE id=?',[b.school_name, req.sid]);
  await loadConfig(req.sid);
  res.json(getConfig(req.sid));
}));
app.get('/api/timetable/slots', h(async (req,res)=>res.json(slotsForDay(Number(req.query.day||0), req.sid))));

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
  const maxLoad={}; (await q('SELECT id,max_load FROM tt_teacher WHERE school_id=?',[sid])).forEach(t=>maxLoad[t.id]=t.max_load||999);
  const load={};
  const teachersForSubject=sid=>tmap.filter(m=>m.subject_id===sid).map(m=>m.teacher_id);

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
  const pool={}; classes.forEach(c=>{ pool[c.id]=shuffleStable(poolFor(c.id,totalSlots),c.id); });
  const ptr={}; classes.forEach(c=>ptr[c.id]=0);
  const toInsert=[];
  DAYS.forEach((_,di)=>{
    const slots=teachingSlots(di, sid);
    slots.forEach((_,pi)=>{
      const usedT=new Set(), usedR=new Set();
      classes.forEach(c=>{
        const subjId=pool[c.id][ptr[c.id]++];
        if(subjId==null) return;   // class has no schedulable subjects
        let opts=teachersForSubject(subjId).filter(t=>!usedT.has(t)&&!(absent[t]&&absent[t].has(di))&&(load[t]||0)<maxLoad[t]);
        opts.sort((a,b)=>(load[a]||0)-(load[b]||0));
        let t=opts[0] ?? teachersForSubject(subjId).find(x=>!usedT.has(x)&&(load[x]||0)<maxLoad[x]) ?? null;
        const room=rooms.find(r=>!usedR.has(r.id));
        toInsert.push([c.id,di,pi,subjId,t,room?room.id:null,sid]);
        if(t){usedT.add(t);load[t]=(load[t]||0)+1;} if(room)usedR.add(room.id);
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

// ---------- TIMETABLE VERSIONS (saved snapshots) ----------
app.get('/api/versions', h(async (req,res)=>{
  res.json(await q('SELECT id,name,session,created_at,cell_count FROM tt_snapshot WHERE school_id=? ORDER BY id DESC',[req.sid]));
}));
// save the current live timetable as a named version
app.post('/api/versions', h(async (req,res)=>{
  const sid=req.sid;
  const name=(req.body.name||'').trim()||('Version '+((await q1('SELECT COUNT(*)::int n FROM tt_snapshot WHERE school_id=?',[sid])).n+1));
  const session=(getConfig(sid)||{}).academic_session||null;
  const cells=await q('SELECT class_id,day_of_week,period_index,subject_id,teacher_id,room_id FROM tt_timetable WHERE school_id=?',[sid]);
  const snap=await tx(async (cq,cq1)=>{
    const s=await cq1('INSERT INTO tt_snapshot(name,session,created_at,cell_count,school_id) VALUES(?,?,now()::text,?,?) RETURNING id',[name,session,cells.length,sid]);
    for(const c of cells)
      await cq('INSERT INTO tt_snapshot_cell(snapshot_id,class_id,day_of_week,period_index,subject_id,teacher_id,room_id) VALUES(?,?,?,?,?,?,?)',
        [s.id,c.class_id,c.day_of_week,c.period_index,c.subject_id,c.teacher_id,c.room_id]);
    return s;
  });
  res.json({ok:true, id:snap.id, name, cells:cells.length});
}));
// restore a saved version → overwrites the live timetable (only if the snapshot belongs to this school)
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
  res.json({ok:true, cells:cells.length});
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

// ---------- LIVE MONITOR ----------
app.get('/api/timetable/monitor', h(async (req,res)=>{
  let day,hhmm;
  if(req.query.day!=null&&req.query.time){day=Number(req.query.day);hhmm=req.query.time;}
  else{const d=new Date();day=(d.getDay()+6)%7;if(day>5)day=5;hhmm=String(d.getHours()).padStart(2,'0')+':'+String(d.getMinutes()).padStart(2,'0');}
  const sid=req.sid;
  const absent={}; (await q('SELECT * FROM tt_absence WHERE school_id=?',[sid])).forEach(a=>{(absent[a.teacher_id]=absent[a.teacher_id]||new Set()).add(a.day_of_week);});
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
          const sub=await q1('SELECT proxy_teacher_id FROM tt_substitution WHERE day_of_week=? AND class_id=? AND period_index=?',[day,c.id,pi]);
          const absentTeacher=cell.teacher_id&&absent[cell.teacher_id]&&absent[cell.teacher_id].has(day);
          if(absentTeacher && sub && sub.proxy_teacher_id){ status='proxy'; text=subj+' · Proxy: '+(await q1('SELECT name FROM tt_teacher WHERE id=?',[sub.proxy_teacher_id])).name; }
          else if(absentTeacher){ status='proxy'; text=subj+' · Proxy needed'; }
          else { const tch=cell.teacher_id?(await q1('SELECT name FROM tt_teacher WHERE id=?',[cell.teacher_id])).name:'—'; status='live'; text=subj+' · '+tch; }
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

// ---------- STARTUP ----------
(async () => {
  await init();
  const PORT=process.env.PORT||4100;
  app.listen(PORT,()=>console.log(`Timetable module running → http://localhost:${PORT}`));
})().catch(e=>{ console.error('Startup failed:', e); process.exit(1); });
