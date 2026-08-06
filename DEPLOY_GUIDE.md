# Aumtara Timetable — Free Online Deploy Guide (Hinglish)

Is guide se aapka Timetable + Diary module **internet pe live** ho jayega — koi bhi
teacher apne phone/laptop se URL kholke use kar sakta hai, aur **data hamesha
save** rahega. Sab kuch **free** hai, **credit card nahi** chahiye.

**3 free accounts banenge:**
1. **Neon** — free Postgres database (jahan data save hoga)
2. **GitHub** — jahan code rakhna hai
3. **Render** — jahan app chalegi (live URL)

Total time: ~20–30 minute. Ek-ek step follow karo.

---

## STEP 1 — Neon pe free database banao (~5 min)

1. Browser mein kholo: **https://neon.com** → **Sign up** (Google se sign up sabse aasan).
2. Sign-up ke baad wo poochhega project banao — **"Create project"**:
   - Project name: `aumtara-timetable` (kuch bhi)
   - Postgres version: default rehne do
   - Region: aapke paas ka koi bhi (e.g. Singapore / Mumbai agar dikhe)
   - **Create** dabao.
3. Ab wo ek **connection string** dikhayega. Ye sabse zaroori cheez hai. Dikhega aisa:
   ```
   postgresql://neondb_owner:XXXXXXXX@ep-xxxx-xxxx.ap-southeast-1.aws.neon.tech/neondb?sslmode=require
   ```
   - "**Connection string**" / "**Connect**" wale box mein milega.
   - **Poori line copy karo** aur kahin note/notepad mein paste karke rakh lo.
     (Isme aapka password hai — kisi ko mat bhejo.)

> Bas Neon ka kaam ho gaya. Tables khud ban jayenge jab app pehli baar chalegi.

---

## STEP 2 — Code GitHub pe daalo (~8 min)

1. **https://github.com** → **Sign up** (agar account nahi hai).
2. Login ke baad upar-right **"+"** → **"New repository"**:
   - Repository name: `aumtara-timetable`
   - **Private** select karo (aapka code private rahega)
   - Baaki default. **"Create repository"** dabao.
3. Agli screen pe **"uploading an existing file"** link par click karo
   (ya URL: `github.com/aapka-username/aumtara-timetable/upload`).
4. Ab **is module ke andar ki saari files** drag-and-drop karo:
   - Wo folder kholo jahan zip extract kiya hai (`timetable-module`).
   - Andar ki **saari files select karo** — `server.js`, `db.js`, `package.json`,
     `package-lock.json`, `render.yaml`, `.gitignore`, `README.md`,
     `DEPLOY_GUIDE.md`, aur **`public` folder**.
   - ⚠️ `node_modules` folder **mat** daalna (agar hai bhi to skip — bhaari aur useless).
   - Inhe GitHub ke upload box mein drop kar do.
5. Neeche **"Commit changes"** (green button) dabao.

> Files upload ho gayi. Ab GitHub pe aapko `server.js`, `db.js` etc. dikhne chahiye.

---

## STEP 3 — Render pe app deploy karo (~8 min)

1. **https://render.com** → **Get Started** → **GitHub se sign up** karo
   (isse Render aapke repo dekh paayega).
2. Dashboard pe **"New +"** → **"Web Service"**.
3. **"Build and deploy from a Git repository"** → **Connect** → apna
   `aumtara-timetable` repo select karo. (Pehli baar Render se GitHub access
   "Configure" karna pad sakta hai — allow kar do.)
4. Settings form aayega. Ye bharo:
   - **Name:** `aumtara-timetable` (aapka URL isi se banega)
   - **Region:** koi bhi paas wala (e.g. Singapore)
   - **Branch:** `main`
   - **Runtime / Language:** **Node**
   - **Build Command:** `npm install`
   - **Start Command:** `npm start`
   - **Instance Type:** **Free** select karo.
5. Neeche **"Environment Variables"** / **"Advanced"** → **"Add Environment Variable"**:
   - **Key:** `DATABASE_URL`
   - **Value:** wahi **Neon connection string** jo Step 1 mein copy ki thi
     (poori `postgresql://...sslmode=require` wali line paste karo).
6. **"Create Web Service"** (ya "Deploy") dabao.
7. Render ab build karega — logs chalte dikhenge (~2–4 min). Jab
   **"Timetable module running"** aur **"Live"** (green) dikhe → ho gaya! 🎉
8. Upar aapka URL milega, jaise:
   ```
   https://aumtara-timetable.onrender.com
   ```
   Use kholo — aapka Timetable + Diary module **live** hai.

---

## Bas! Ab kya-kya dhyaan rakhna hai

- **Data save rehta hai** — Neon Postgres pe, restart/deploy ke baad bhi. Teachers
  jo diary daalenge wo tikega. ✅
- **Pehli baar khulne mein slow** — Render ka free plan 15 min inactivity ke baad
  app ko "sula" deta hai. Agli baar koi kholega to ~40–50 sec lagega jaagne mein,
  phir normal speed. (Data safe rehta hai, sirf pehla load slow.)
- **School ka naam set karo** — app mein **Academic Hours** tab mein school name
  daal do; wo diary/print pe dikhega.
- **Sample data hata do** — pehli baar demo classes/teachers seed hote hain.
  Setup tab se delete karke apne real classes/subjects/teachers daal do, ya
  Excel template se import kar lo.

## Update kaise karein (baad mein)
Code badalna ho to GitHub pe nayi file upload/commit karo — Render **apne aap**
dobara deploy kar dega. Kuch aur karne ki zaroorat nahi.

## Kuch atke to
Render dashboard → aapki service → **"Logs"** tab kholo. Wahan error dikhega.
Wo error mujhe bhej dena, main fix bata dunga. Sabse common galti:
`DATABASE_URL` galat/adhoora paste hona — Neon se poori line dobara copy karke
Render → **Environment** mein update kar do, phir **"Manual Deploy"** → **"Deploy latest"**.
