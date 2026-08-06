# Aumtara — Timetable Module (complete, working)

A full, runnable timetable module: **Node.js + Express + Postgres**. Not a mockup — everything saves to a real database and survives restarts. Runs locally on a Postgres, and deploys free to Render + Neon Postgres (see `DEPLOY_GUIDE.md`).

## Features
- **Setup / Manage** — add / edit / delete Classes, Subjects, Rooms, and Teachers. Each teacher has **name, qualification, main subject, teachable subjects, and a weekly period load (max periods/week)**.
- **Import / Export (Excel)** — bulk-add Classes, Subjects, Rooms and Teachers from an .xlsx (download the template first). Teachers sheet columns: *Name, Qualification, Main Subject, Optional Subjects, Weekly Period Load*. Export the whole timetable and the setup data to Excel.
- **Academic Hours** — weekday/Saturday times + period & break length; period slots generate automatically.
- **Subject Weekly Quota** — set how many periods/week each subject gets per class; auto-generate honours it.
- **Builder** — assign **subject + teacher + room** per class/period; saved instantly.
- **Conflict management** — server detects a **teacher OR room** double-booked in the same period; the UI highlights the clash.
- **Auto-generate** — fills the whole timetable: **quota-aware, room-aware, and respects each teacher's weekly load cap**, avoiding teacher/room clashes.
- **Class / Teacher / Whole-School** views — grids with clean **Print / Save-as-PDF** and Excel export; teacher view shows workload vs max.
- **Teacher's Diary (official format)** — for a chosen teacher & date the day's periods auto-list; log **Lesson, Topic, Learning Outcome, Assessment of LO (Strategies), Home Work, Teaching Aids used** per period. **Fast fill with pulldowns:** pick the **Day** from a dropdown (Date auto-jumps to that weekday), each period shows a **Subject** dropdown, and the **Lesson / Chapter** field is a pick-or-type list driven by the chapters you set per subject — so common entries are one click, not retyped. A **Diary Report** view renders the official school format (School name → "Teacher's Diary" → Teacher / Class / Subject / Date / Day → table) with **Print / PDF, Export Excel, and Edit** buttons and an optional single-class filter. Set your school name in *Academic Hours*.
- **Chapters (per subject)** — manage the chapter list for each subject in *Setup*; these feed the Teacher's Diary chapter pulldown and ship in the Excel import/export template (Chapters sheet: *Subject, Chapter*).
- **Substitution / Cover** — mark a teacher absent → every period needing cover is listed with the **free teachers**; assign a proxy in one click.
- **Live Classroom Monitor** — real-time LIVE / BREAK / IDLE / PROXY per class (server-computed); shows assigned proxies. Day/time can be simulated.
- **Help & Guide** — built-in step-by-step walkthrough tab.

## Run it
Needs a Postgres database. Set `DATABASE_URL`, then:
```bash
npm install        # express + pg + exceljs + multer
export DATABASE_URL="postgres://USER:PASS@HOST:5432/DBNAME"
npm start          # http://localhost:4100
```
If `DATABASE_URL` is not set it defaults to `postgres://postgres:postgres@127.0.0.1:5432/timetable` (a local Postgres). First run creates all tables and seeds sample data (3 classes, 7 subjects, 5 teachers, 4 rooms). `PORT=5000 npm start` to change port. **To deploy free online, see `DEPLOY_GUIDE.md`.**

## Project layout
```
timetable-module/
  server.js            Express app + all REST endpoints + logic (async / pg)
  db.js                Postgres schema + helpers + one-time seed
  public/index.html    Frontend (no build step)
  render.yaml          Render Blueprint (free web service)
  .env.example         DATABASE_URL sample
  DEPLOY_GUIDE.md      Step-by-step free deploy (Neon + Render)
  README.md, package.json
```

## REST API
| Method & path | Purpose |
|---|---|
| GET/POST/PUT/DELETE /api/classes, /api/subjects, /api/rooms | manage reference data |
| GET/POST/PUT/DELETE /api/teachers | teachers incl. `subjects[]` mapping |
| GET / PUT /api/timetable/config | Academic Hours |
| GET /api/timetable/slots?day= | computed period slots |
| GET /api/timetable?class_id= | the grid |
| PUT / DELETE /api/timetable/cell | upsert / clear a cell → `{ok,conflict:{teacher,room}}` |
| GET /api/quota / PUT /api/quota | subject periods-per-week per class |
| POST /api/timetable/auto-generate | quota + room aware fill |
| GET /api/timetable/conflicts | `{teacher:[], room:[]}` clashes |
| GET /api/timetable/teacher/:id | teacher schedule + weekly load |
| GET / POST /api/timetable/absence | mark / clear absence |
| GET /api/timetable/cover?day= | periods needing cover + free teachers |
| POST /api/timetable/substitute | assign a proxy |
| GET /api/timetable/monitor?day=&time= | live status of all classes |

## Database (Postgres)
`tt_class, tt_subject, tt_teacher, tt_room, tt_teacher_subject, tt_config,
tt_timetable (unique class+day+period), tt_quota, tt_chapter, tt_absence,
tt_substitution, tt_diary`. Full schema in `db.js`. Tables are created
automatically on first start.

## Integrating into Aumtara
1. Point `tt_class / tt_subject / tt_teacher / tt_room` at Aumtara's existing Classes, Subjects, Staff and Rooms tables (keep `tt_config, tt_timetable, tt_quota, tt_absence, tt_substitution`).
2. Port `server.js` routes into your Laravel/Node controllers (standard SQL).
3. Scope every query by `school_id` + `academic_year_id`; add auth middleware.
4. For a truly real-time Live Monitor, overlay the attendance/biometric "class started" signal and the leave module on top of the plan-based status.

Built for Aumtara. Sample data is illustrative.
