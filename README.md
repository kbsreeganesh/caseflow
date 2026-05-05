# ⚕ CaseFlow — Medical Case Management

**Live URL (after deploy):** https://kbsreeganesh.github.io/caseflow/

---

## One-time Setup — do this in order

### Step 1 — Create the GitHub repository

1. Go to https://github.com/new
2. Repository name: **caseflow** (must be exactly this)
3. Set to **Public**
4. Leave everything else blank — do NOT add README or .gitignore
5. Click **Create repository**

---

### Step 2 — Set up Supabase (free shared database)

1. Go to https://supabase.com → **Start your project** → sign in with GitHub
2. Click **New project**
   - Name: `caseflow`
   - Region: pick closest to India (e.g. Singapore)
   - Generate a strong password — save it somewhere
3. Wait ~2 minutes for the project to spin up
4. Go to **SQL Editor** (left sidebar) → click **New query**
5. Paste the contents of `supabase/schema.sql` → click **Run**
6. Go to **Settings → API** (left sidebar)
7. Copy two values:
   - **Project URL** → looks like `https://abcdefgh.supabase.co`
   - **anon public** key → long JWT string starting with `eyJ...`

---

### Step 3 — Add secrets to GitHub

1. Go to https://github.com/kbsreeganesh/caseflow/settings/secrets/actions
2. Click **New repository secret** — add these two:

   | Name | Value |
   |------|-------|
   | `VITE_SUPABASE_URL` | your Project URL from step 2 |
   | `VITE_SUPABASE_ANON_KEY` | your anon key from step 2 |

---

### Step 4 — Push the code

Open a terminal, navigate to this folder, then run:

```bash
git init
git add .
git commit -m "Initial CaseFlow deployment"
git branch -M main
git remote add origin https://github.com/kbsreeganesh/caseflow.git
git push -u origin main
```

If Git asks for credentials:
- Username: `kbsreeganesh`
- Password: use a **GitHub Personal Access Token** (not your password)
  - Create one at: https://github.com/settings/tokens → Generate new token (classic)
  - Scopes needed: `repo`

---

### Step 5 — Enable GitHub Pages

1. Go to https://github.com/kbsreeganesh/caseflow/settings/pages
2. Under **Source** select: **GitHub Actions**
3. Click **Save**

The deploy action will run automatically. Check progress at:
https://github.com/kbsreeganesh/caseflow/actions

First deploy takes ~2 minutes. After that, every `git push` auto-deploys.

---

### Step 6 — Test locally (optional but recommended)

```bash
# Create your local .env file
cp .env.example .env
# Edit .env — paste your Supabase URL and anon key

npm install
npm run dev
# → http://localhost:5173/caseflow/
```

---

## Your live app

**URL:** https://kbsreeganesh.github.io/caseflow/

**Default admin password:** `admin123`  
→ Change this immediately in Settings after first login

---

## First login checklist

- [ ] Admin → Settings → Change admin password
- [ ] User Master → Upload your user Excel (Name, Role, PIN, User ID 1, Password 1, ...)
- [ ] Import Cases → Upload your case spreadsheet
- [ ] Settings → Add Google Form URLs (Analyser + Supervisor)
- [ ] Settings → Add Form Response Sheet URLs
- [ ] Settings → Configure sync column mapping (HrnId / Case Key)
- [ ] Dashboard → Sync Now

---

## How data is shared

All data (cases, users, progress) is stored in **Supabase** — a shared cloud
database. Every doctor logging in from any device will see the same cases and
the same progress. Admin changes apply immediately for everyone.

---

## Project layout

```
caseflow/
├── src/
│   ├── main.jsx       ← mounts React + loads Supabase adapter
│   ├── storage.js     ← Supabase adapter (window.storage API)
│   └── App.jsx        ← full CaseFlow application
├── supabase/
│   └── schema.sql     ← run once in Supabase SQL editor
├── .env.example       ← copy to .env for local dev
├── vite.config.js
├── package.json
└── .github/
    └── workflows/
        └── deploy.yml ← auto-deploys to GitHub Pages on push
```
