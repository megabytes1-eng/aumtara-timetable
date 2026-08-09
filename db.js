// Postgres persistence layer for the Timetable module (v3 — Neon/Render ready)
// Migrated from better-sqlite3 → pg so data persists on free cloud hosts.
const { Pool } = require('pg');
const crypto = require('crypto');

// ---- password hashing (built-in scrypt; no native deps) ----
function hashPw(pw){ const salt=crypto.randomBytes(16).toString('hex'); const h=crypto.scryptSync(String(pw),salt,32).toString('hex'); return salt+':'+h; }
function verifyPw(pw,stored){ try{ if(!stored||!stored.includes(':'))return false; const [salt,h]=stored.split(':'); const hh=crypto.scryptSync(String(pw),salt,32).toString('hex'); const a=Buffer.from(h,'hex'),b=Buffer.from(hh,'hex'); return a.length===b.length && crypto.timingSafeEqual(a,b); }catch(e){ return false; } }

const connectionString =
  process.env.DATABASE_URL ||
  'postgres://postgres:postgres@127.0.0.1:5432/timetable';

// Neon (and most managed Postgres) require SSL. Local dev does not.
const needSSL = /neon\.tech|render\.com|supabase|sslmode=require/i.test(connectionString)
  || process.env.PGSSL === '1';
const pool = new Pool({
  connectionString,
  ssl: needSSL ? { rejectUnauthorized: false } : false,
  max: 5,
});

// Convert better-sqlite3 style "?" placeholders to Postgres "$1,$2,..."
function toPg(sql) { let i = 0; return sql.replace(/\?/g, () => '$' + (++i)); }

// Query helpers (async). Keep call-sites close to the old .all()/.get()/.run() shape.
async function q(sql, params = [])  { const r = await pool.query(toPg(sql), params); return r.rows; }        // → array
async function q1(sql, params = []) { const r = await pool.query(toPg(sql), params); return r.rows[0]; }      // → first row | undefined
async function run(sql, params = []){ return pool.query(toPg(sql), params); }                                 // → result

// Transaction: fn receives a scoped query helper cq(sql, params)
async function tx(fn) {
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    const cq  = async (sql, params = []) => (await client.query(toPg(sql), params)).rows;
    const cq1 = async (sql, params = []) => (await client.query(toPg(sql), params)).rows[0];
    const out = await fn(cq, cq1);
    await client.query('COMMIT');
    return out;
  } catch (e) {
    await client.query('ROLLBACK');
    throw e;
  } finally {
    client.release();
  }
}

// ---- in-memory config cache (per school) so slot math can stay synchronous ----
let CONFIGS = {};   // school_id -> config row
async function loadConfig(sid) {
  if (sid == null) { CONFIGS = {}; (await q('SELECT * FROM tt_config')).forEach(r => { if (r.school_id != null) CONFIGS[r.school_id] = r; }); return CONFIGS; }
  const c = await q1('SELECT * FROM tt_config WHERE school_id=?', [sid]); if (c) CONFIGS[sid] = c; return c;
}
function getConfigCached(sid) { return (sid != null ? CONFIGS[sid] : null) || null; }

async function init() {
  await run(`
  CREATE TABLE IF NOT EXISTS tt_class   (id SERIAL PRIMARY KEY, name TEXT NOT NULL);
  CREATE TABLE IF NOT EXISTS tt_subject (id SERIAL PRIMARY KEY, name TEXT NOT NULL);
  CREATE TABLE IF NOT EXISTS tt_teacher (id SERIAL PRIMARY KEY, name TEXT NOT NULL, qualification TEXT, main_subject_id INTEGER, max_load INTEGER);
  CREATE TABLE IF NOT EXISTS tt_room    (id SERIAL PRIMARY KEY, name TEXT NOT NULL, capacity INTEGER);
  CREATE TABLE IF NOT EXISTS tt_teacher_subject (teacher_id INTEGER, subject_id INTEGER, PRIMARY KEY (teacher_id, subject_id));

  CREATE TABLE IF NOT EXISTS tt_config (
     id INTEGER PRIMARY KEY CHECK (id=1),
     weekday_start TEXT, weekday_end TEXT, saturday_end TEXT,
     period_minutes INTEGER, break_after_period INTEGER, break_minutes INTEGER,
     school_name TEXT);

  CREATE TABLE IF NOT EXISTS tt_timetable (
     id SERIAL PRIMARY KEY,
     class_id INTEGER NOT NULL,
     day_of_week INTEGER NOT NULL,
     period_index INTEGER NOT NULL,
     subject_id INTEGER, teacher_id INTEGER, room_id INTEGER,
     UNIQUE (class_id, day_of_week, period_index));
  CREATE INDEX IF NOT EXISTS ix_tt_slot ON tt_timetable (day_of_week, period_index, teacher_id);
  CREATE INDEX IF NOT EXISTS ix_tt_room ON tt_timetable (day_of_week, period_index, room_id);

  CREATE TABLE IF NOT EXISTS tt_quota (
     class_id INTEGER, subject_id INTEGER, per_week INTEGER,
     PRIMARY KEY (class_id, subject_id));

  CREATE TABLE IF NOT EXISTS tt_chapter (
     id SERIAL PRIMARY KEY, subject_id INTEGER NOT NULL, name TEXT NOT NULL, seq INTEGER);

  CREATE TABLE IF NOT EXISTS tt_absence (
     id SERIAL PRIMARY KEY,
     teacher_id INTEGER NOT NULL, day_of_week INTEGER NOT NULL, reason TEXT,
     UNIQUE (teacher_id, day_of_week));

  CREATE TABLE IF NOT EXISTS tt_substitution (
     id SERIAL PRIMARY KEY,
     day_of_week INTEGER NOT NULL, class_id INTEGER NOT NULL, period_index INTEGER NOT NULL,
     proxy_teacher_id INTEGER,
     UNIQUE (day_of_week, class_id, period_index));

  CREATE TABLE IF NOT EXISTS tt_diary (
     id SERIAL PRIMARY KEY,
     teacher_id INTEGER NOT NULL,
     entry_date TEXT NOT NULL,
     day_of_week INTEGER,
     class_id INTEGER, subject_id INTEGER, period_index INTEGER,
     topic TEXT, observation TEXT, plan TEXT, remarks TEXT,
     created_at TEXT,
     lesson TEXT, learning_outcome TEXT, assessment_lo TEXT, homework TEXT, teaching_aids TEXT,
     UNIQUE (teacher_id, entry_date, class_id, period_index));
  `);

  // idempotent safety migrations (no-ops on a fresh DB)
  await run(`ALTER TABLE tt_config ADD COLUMN IF NOT EXISTS school_name TEXT`);
  for (const c of ['lesson','learning_outcome','assessment_lo','homework','teaching_aids'])
    await run(`ALTER TABLE tt_diary ADD COLUMN IF NOT EXISTS ${c} TEXT`);
  // Phase 1 — richer hours: multiple short breaks + separate lunch, per-period durations, working days
  await run(`ALTER TABLE tt_config ADD COLUMN IF NOT EXISTS short_break_minutes INTEGER`);
  await run(`ALTER TABLE tt_config ADD COLUMN IF NOT EXISTS short_break_after TEXT`);   // CSV of period numbers, e.g. "6"
  await run(`ALTER TABLE tt_config ADD COLUMN IF NOT EXISTS lunch_minutes INTEGER`);
  await run(`ALTER TABLE tt_config ADD COLUMN IF NOT EXISTS lunch_after INTEGER`);       // period number after which lunch falls
  await run(`ALTER TABLE tt_config ADD COLUMN IF NOT EXISTS period_durations TEXT`);     // CSV per-period minutes, blank = use period_minutes
  await run(`ALTER TABLE tt_config ADD COLUMN IF NOT EXISTS working_days TEXT`);         // CSV day indices 0=Mon..5=Sat
  await run(`ALTER TABLE tt_config ADD COLUMN IF NOT EXISTS saturday_start TEXT`);       // Saturday-only start time; blank = same as weekday_start
  // Saturday-independent timing (NULL on any column = fall back to the weekday equivalent)
  await run(`ALTER TABLE tt_config ADD COLUMN IF NOT EXISTS sat_period_minutes INTEGER`);
  await run(`ALTER TABLE tt_config ADD COLUMN IF NOT EXISTS sat_period_durations TEXT`);
  await run(`ALTER TABLE tt_config ADD COLUMN IF NOT EXISTS sat_lunch_after INTEGER`);
  await run(`ALTER TABLE tt_config ADD COLUMN IF NOT EXISTS sat_lunch_minutes INTEGER`);
  await run(`ALTER TABLE tt_config ADD COLUMN IF NOT EXISTS sat_short_break_after TEXT`);
  await run(`ALTER TABLE tt_config ADD COLUMN IF NOT EXISTS sat_short_break_minutes INTEGER`);
  // Phase 2 — class-teacher assignment + subject active/inactive
  await run(`ALTER TABLE tt_class   ADD COLUMN IF NOT EXISTS class_teacher_id INTEGER`); // designated class teacher
  await run(`ALTER TABLE tt_subject ADD COLUMN IF NOT EXISTS active INTEGER DEFAULT 1`); // 1=schedulable, 0=archived
  await run(`UPDATE tt_subject SET active=1 WHERE active IS NULL`);
  // Phase 3 — optional class structure (Board / Medium / Standard / Section). All additive & optional.
  for (const c of ['board','medium','standard','section'])
    await run(`ALTER TABLE tt_class ADD COLUMN IF NOT EXISTS ${c} TEXT`);
  // Phase 4 — academic session label + saved timetable versions (snapshots)
  await run(`ALTER TABLE tt_config ADD COLUMN IF NOT EXISTS academic_session TEXT`);
  await run(`ALTER TABLE tt_config ADD COLUMN IF NOT EXISTS board TEXT`);   // school board (CBSE/State/IB/…)
  await run(`CREATE TABLE IF NOT EXISTS tt_snapshot (
     id SERIAL PRIMARY KEY, name TEXT NOT NULL, session TEXT, created_at TEXT, cell_count INTEGER)`);
  await run(`CREATE TABLE IF NOT EXISTS tt_snapshot_cell (
     snapshot_id INTEGER NOT NULL, class_id INTEGER NOT NULL,
     day_of_week INTEGER NOT NULL, period_index INTEGER NOT NULL,
     subject_id INTEGER, teacher_id INTEGER, room_id INTEGER)`);
  await run(`CREATE INDEX IF NOT EXISTS ix_snap_cell ON tt_snapshot_cell (snapshot_id)`);
  // Phase 5 — role-based users + login sessions
  await run(`CREATE TABLE IF NOT EXISTS tt_user (
     id SERIAL PRIMARY KEY, name TEXT, role TEXT NOT NULL DEFAULT 'teacher',
     login_id TEXT UNIQUE NOT NULL, email TEXT, mobile TEXT, qualification TEXT,
     main_subject_id INTEGER, password_hash TEXT NOT NULL, active INTEGER DEFAULT 1, created_at TEXT)`);
  await run(`CREATE TABLE IF NOT EXISTS tt_session (token TEXT PRIMARY KEY, user_id INTEGER NOT NULL, created_at TEXT)`);
  // seed a default super-admin so the very first login is possible (must be changed after first login)
  const uc = (await q1('SELECT COUNT(*)::int AS n FROM tt_user')).n;
  if (!uc)
    await run(`INSERT INTO tt_user(name,role,login_id,password_hash,active,created_at) VALUES(?,?,?,?,1,now()::text)`,
      ['Administrator','master','admin',hashPw('admin123')]);

  // ===================== MULTI-SCHOOL (multi-tenancy) schema =====================
  await run(`CREATE TABLE IF NOT EXISTS tt_school (
     id SERIAL PRIMARY KEY, name TEXT NOT NULL, board TEXT, medium TEXT, active INTEGER DEFAULT 1, created_at TEXT)`);
  await run(`ALTER TABLE tt_school ADD COLUMN IF NOT EXISTS logo TEXT`);   // data-URL (jpg/png) or null
  // let each school own a tt_config row (drop the single-row id=1 restriction)
  await run(`ALTER TABLE tt_config DROP CONSTRAINT IF EXISTS tt_config_id_check`);
  await run(`ALTER TABLE tt_config ADD COLUMN IF NOT EXISTS school_id INTEGER`);
  await run(`ALTER TABLE tt_config ADD COLUMN IF NOT EXISTS medium TEXT`);
  // school_id on every per-school data table + on users
  for (const tbl of ['tt_class','tt_subject','tt_teacher','tt_room','tt_quota','tt_chapter','tt_absence','tt_substitution','tt_diary','tt_timetable','tt_snapshot','tt_user'])
    await run(`ALTER TABLE ${tbl} ADD COLUMN IF NOT EXISTS school_id INTEGER`);

  const n = (await q1('SELECT COUNT(*)::int AS n FROM tt_class')).n;
  if (!n) await seed();

  const cfg = await q1('SELECT COUNT(*)::int AS n FROM tt_config');
  if (!cfg.n)
    await run(`INSERT INTO tt_config(id,weekday_start,weekday_end,saturday_end,period_minutes,break_after_period,break_minutes)
               VALUES (1,?,?,?,?,?,?)`, ['07:30','13:00','11:30',35,3,20]);

  const sn = await q1('SELECT school_name FROM tt_config WHERE id=1');
  if (!sn || !sn.school_name)
    await run('UPDATE tt_config SET school_name=? WHERE id=1', ['Your School Name']);

  // backfill Phase-1 hours defaults: map the old single break → Lunch; sensible defaults for the rest
  await run(`UPDATE tt_config SET
     working_days       = COALESCE(working_days, '0,1,2,3,4,5'),
     lunch_after        = COALESCE(lunch_after, break_after_period),
     lunch_minutes      = COALESCE(lunch_minutes, break_minutes),
     short_break_minutes= COALESCE(short_break_minutes, 0),
     short_break_after  = COALESCE(short_break_after, ''),
     period_durations   = COALESCE(period_durations, ''),
     saturday_start     = COALESCE(saturday_start, weekday_start),
     sat_period_minutes = COALESCE(sat_period_minutes, period_minutes),
     sat_period_durations = COALESCE(sat_period_durations, period_durations, ''),
     sat_lunch_after    = COALESCE(sat_lunch_after, lunch_after, break_after_period),
     sat_lunch_minutes  = COALESCE(sat_lunch_minutes, lunch_minutes, break_minutes),
     sat_short_break_after = COALESCE(sat_short_break_after, short_break_after, ''),
     sat_short_break_minutes = COALESCE(sat_short_break_minutes, short_break_minutes, 0)
   WHERE id=1`);

  // ---- MULTI-SCHOOL seed: create "School 1" from the current config and adopt all existing data ----
  const scn = (await q1('SELECT COUNT(*)::int AS n FROM tt_school')).n;
  if (!scn) {
    const cur = await q1('SELECT school_name, board, medium FROM tt_config WHERE id=1');
    const nm = (cur && cur.school_name && cur.school_name !== 'Your School Name') ? cur.school_name : 'School 1';
    const s1 = await q1('INSERT INTO tt_school(name,board,medium,active,created_at) VALUES(?,?,?,1,now()::text) RETURNING id',
      [nm, (cur && cur.board) || null, (cur && cur.medium) || null]);
    const sid = s1.id;
    await run('UPDATE tt_config SET school_id=? WHERE id=1', [sid]);
    // adopt every existing row (data + users) into School 1
    for (const tbl of ['tt_class','tt_subject','tt_teacher','tt_room','tt_quota','tt_chapter','tt_absence','tt_substitution','tt_diary','tt_timetable','tt_snapshot','tt_user'])
      await run(`UPDATE ${tbl} SET school_id=? WHERE school_id IS NULL`, [sid]);
  }

  await loadConfig();
}

// create a fresh config row (defaults) for a newly-added school
async function seedConfigForSchool(schoolId, schoolName, board, medium) {
  const nid = (await q1('SELECT COALESCE(MAX(id),0)+1 AS n FROM tt_config')).n;
  await run(`INSERT INTO tt_config(id,school_id,weekday_start,weekday_end,saturday_end,period_minutes,break_after_period,break_minutes,
       school_name,board,medium,working_days,lunch_after,lunch_minutes,short_break_minutes,short_break_after,period_durations,saturday_start,
       sat_period_minutes,sat_period_durations,sat_lunch_after,sat_lunch_minutes,sat_short_break_after,sat_short_break_minutes)
     VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    [nid, schoolId, '07:30','13:00','11:30',35,3,20, schoolName||'New School', board||null, medium||null, '0,1,2,3,4,5', 3, 20, 0, '', '', '07:30',
     35, '', 3, 20, '', 0]);
  await loadConfig(schoolId);
  return nid;
}

async function seed() {
  const cls = ['Class VI-A','Class VII-A','Class VIII-A'];
  const sub = ['English','Maths','Science','Social','Hindi','Computer','PE'];
  const rooms = [['Room 101',40],['Room 102',40],['Lab-1',30],['Computer Lab',30]];
  // name, qualification, main subject, [optional subjects], max weekly load
  const tch = [
    ['R. Kumar','M.Sc, B.Ed','Maths',['Computer'],30],
    ['A. Sharma','M.Sc Physics, B.Ed','Science',[],28],
    ['P. Desai','M.A, B.Ed','English',['Hindi'],30],
    ['S. Nair','M.A History, B.Ed','Social',[],26],
    ['V. Iyer','B.P.Ed','PE',['Science'],24],
  ];
  const chapters = {
    English:['Grammar','Prose','Poetry','Writing Skills'],
    Maths:['Fractions','Decimals','Geometry','Algebra Basics','Mensuration'],
    Science:['Living World','Matter Around Us','Force & Motion','Light','The Human Body'],
    Social:['Our Past','Maps & Geography','Civics Basics'],
    Hindi:['व्याकरण','गद्य','पद्य','रचना'],
    Computer:['Intro to Computers','MS Word','Scratch Programming'],
    PE:['Fitness','Athletics','Team Games'],
  };
  await tx(async (cq, cq1) => {
    for (const c of cls) await cq('INSERT INTO tt_class(name) VALUES (?)', [c]);
    for (const s of sub) await cq('INSERT INTO tt_subject(name) VALUES (?)', [s]);
    for (const r of rooms) await cq('INSERT INTO tt_room(name,capacity) VALUES (?,?)', [r[0], r[1]]);

    const subId = {};
    (await cq('SELECT * FROM tt_subject')).forEach(r => subId[r.name] = r.id);

    for (const [name, qual, main, opts, load] of tch) {
      const mid = subId[main];
      const row = await cq1('INSERT INTO tt_teacher(name,qualification,main_subject_id,max_load) VALUES (?,?,?,?) RETURNING id',
        [name, qual, mid, load]);
      await cq('INSERT INTO tt_teacher_subject VALUES (?,?) ON CONFLICT DO NOTHING', [row.id, mid]);
      for (const sn of opts) await cq('INSERT INTO tt_teacher_subject VALUES (?,?) ON CONFLICT DO NOTHING', [row.id, subId[sn]]);
    }

    for (const [sn, list] of Object.entries(chapters)) {
      if (!subId[sn]) continue;
      let i = 0;
      for (const c of list) { i++; await cq('INSERT INTO tt_chapter(subject_id,name,seq) VALUES(?,?,?)', [subId[sn], c, i]); }
    }
  });
}

module.exports = { pool, q, q1, run, tx, init, loadConfig, getConfigCached, hashPw, verifyPw, seedConfigForSchool };
