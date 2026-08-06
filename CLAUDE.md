# CLAUDE.md — Timetable Module

Guidance for any AI agent (or dev) working in this repo. Read this before editing.

## What this is
A **complete, real, working** timetable module for the Aumtara school ERP — **not a prototype/mockup**. Node.js + Express + **Postgres** (`pg`). Everything persists to the database and survives restarts. Migrated from `better-sqlite3` → Postgres so it deploys free to cloud hosts (Neon + Render) where a file-based SQLite would reset. The user explicitly rejected front-end-only prototypes: keep it full-stack and functional.

## Run / test
```bash
npm install          # express, pg, exceljs, multer
export DATABASE_URL="postgres://USER:PASS@HOST:5432/DBNAME"   # defaults to local postgres if unset
npm start            # http://localhost:4100  (PORT=5000 to change)
```
- Needs a Postgres. `DATABASE_URL` unset → defaults to `postgres://postgres:postgres@127.0.0.1:5432/timetable`.
- Tables auto-create on first start; first run seeds sample data (3 classes, 7 subjects, 5 teachers, 4 rooms, 27 chapters).
- Reset: drop & recreate the DB (or `DROP SCHEMA public CASCADE; CREATE SCHEMA public;`) and restart.
- Free deploy steps: see `DEPLOY_GUIDE.md` (Neon Postgres + Render, no credit card).

## Layout
```
server.js          Express app + ALL REST endpoints + scheduling logic (async / pg)
db.js              Postgres pool + q/q1/run/tx helpers, schema, one-time seed, config cache
public/index.html  Whole frontend SPA (vanilla JS, no build step, ~48KB)
render.yaml        Render Blueprint (free web service)
DEPLOY_GUIDE.md    Step-by-step free deploy
README.md          User-facing feature list
package.json
```

## Architecture notes (read before changing)
- **Frontend is one file, no framework.** Vanilla JS, template literals, tab-based render. Globals: `classes, subjects, teachers, rooms, config, TT, absence, conflicts, chaptersAll`. `loadAll()` fetches everything on boot. 13 tabs: dash, setup, hours, quota, build, class, teacher, diary, school, sub, monitor, data, help.
- **Pulldown-with-typing** = HTML `<datalist>` combo boxes (`<input list="chap_<sid>">`). Diary uses these for Lesson/Chapter; Subject is a real `<select>`; Day is a `<select>` that auto-jumps Date to that weekday via `dateForDay(dow)`.
- **DB access is async** — all handlers are `async`, wrapped in `h()` (catches rejects → 500). Use `q()`→rows, `q1()`→first row, `run()`→exec, `tx(fn)`→transaction (fn gets `cq`/`cq1`). SQL keeps `?` placeholders; `db.js` `toPg()` rewrites them to `$1,$2…`. Config is cached in memory (`getConfigCached()`) so slot math stays sync; call `loadConfig()` after any config write.
- **DB migrations are additive & idempotent** — tables use `CREATE TABLE IF NOT EXISTS`; new columns via `ALTER TABLE ... ADD COLUMN IF NOT EXISTS`. Never drop/rename columns on existing DBs.
- **Upserts** use Postgres `INSERT ... ON CONFLICT(cols) DO UPDATE SET x=excluded.x`. Key uniqueness: `tt_timetable(class_id,day_of_week,period_index)`, `tt_diary(teacher_id,entry_date,class_id,period_index)`. Dialect notes: `string_agg(x::text,',')` (not GROUP_CONCAT), `HAVING COUNT(*)>1` (no alias in HAVING), `COUNT(*)::int` (else bigint returns as string), `now()::text` for timestamps, `RETURNING id` for inserted PKs.
- **Auto-generate** is quota-aware + room-aware + respects each teacher's `max_load` (weekly period cap) and absences; avoids teacher/room double-booking. If you touch it, re-verify: 0 conflicts, load caps respected.
- **Conflict detection** = same teacher OR same room in same day+period. Server returns `{teacher:[],room:[]}`; UI highlights.
- **Excel** via `exceljs`; upload via `multer` memoryStorage. Import/export/template share the SAME sheet set: Classes, Subjects, Rooms, Teachers (Name, Qualification, Main Subject, Optional Subjects, Weekly Period Load), Chapters (Subject, Chapter). Keep them in sync if you add a sheet.

## Key tables (full schema in db.js)
`tt_class, tt_subject, tt_teacher(qualification,main_subject_id,max_load), tt_room(capacity), tt_teacher_subject, tt_config(+school_name), tt_timetable, tt_quota, tt_chapter(subject_id,name,seq), tt_absence, tt_substitution, tt_diary`.

`tt_diary` official columns: `lesson, topic, learning_outcome, assessment_lo, homework, teaching_aids` (+ legacy `topic/observation/plan/remarks` kept for back-compat).

## Conventions / gotchas
- **School name is data, not hardcoded.** Default is `"Your School Name"` (set in Academic Hours). Do NOT hardcode any real/example school name (a prior "Mount Maurya International School" placeholder was removed on user request). Same for any competitor's product wording.
- Diary Report renders the **official format**: School name → "Teacher's Diary" → Class/Subject/Date/Day → columns Lesson, Topic, Learning Outcome, Assessment of LO (Strategies), Home Work, Teaching Aids used. Has Export Excel / Print-PDF / Edit buttons + single-class filter.
- **Print CSS** hides `.row,.desc,button,.note,.card>h2`; `printPlain()` clears printhead first. Don't break this when restyling.
- When returning data to the Chrome-automation layer, **return labels/text only, never hrefs/tokens** — the privacy filter blocks results containing query strings/PII (`[BLOCKED]`).
- `index.html` must end with `</script></body></html>` — a past truncation left `setTab` undefined and the app stuck on "Loading…". Verify the tail after big edits.

## Integrating into Aumtara (not yet done)
Point `tt_class/tt_subject/tt_teacher/tt_room` at Aumtara's existing tables; keep `tt_config/tt_timetable/tt_quota/tt_absence/tt_substitution/tt_chapter/tt_diary`. Scope every query by `school_id + academic_year_id`; add auth. Cannot inject into Aumtara's live codebase from here — this ships as a standalone module + integration notes.
