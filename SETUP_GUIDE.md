# Gavakshi Baddies — Setup Guide (V5)
## Fresh setup from scratch · ~10 minutes

---

## PART A — Google Sheet + Apps Script

### Step 1 — Create the Google Sheet
1. Go to **sheets.google.com** → **Blank**
2. Rename it: **Gavakshi Baddies**
3. Keep the tab open

### Step 2 — Add the Apps Script
1. **Extensions** → **Apps Script**
2. Delete the default code
3. Paste everything from **Code.gs**
4. **Ctrl+S** → name it **GavakshiBaddies**

### Step 3 — Deploy as Web App
1. **Deploy** → **New deployment**
2. ⚙️ gear → **Web app**
3. Execute as: **Me** · Who has access: **Anyone**
4. **Deploy** → **Authorize access** → allow
5. **Copy the Web App URL** (ends in `/exec`)
6. **Done**

---

## PART B — GitHub Pages

### Step 4 — Create repo
1. github.com → log in as **robinairepo**
2. **+** → **New repository**
3. Name: **gavakshibaddies** · **Public** · Add README
4. **Create repository**

### Step 5 — Upload files
1. **Add file** → **Upload files**
2. Upload: **index.html**, **manifest.json**, **sw.js**
3. **Commit changes**

### Step 6 — Enable Pages
1. **Settings** → **Pages**
2. Source: **Deploy from a branch** · Branch: **main** · Folder: **/ (root)**
3. **Save** → wait ~1 min
4. Live at: **https://robinairepo.github.io/gavakshibaddies/**

---

## PART C — Install on Phone
- **Android (Chrome):** open URL → ⋮ → **Add to Home screen**
- **iPhone (Safari):** open URL → Share → **Add to Home Screen**

Then paste your Apps Script URL → **Connect & Setup**.

---

## What's new in V5
- Clean **light theme** (no dark mode)
- **Empty date pickers** — tap to pick, today is the default highlighted date
- **Confirmation popup** after every save — button locks during save to prevent double/triple entry
- Existing data loads when you pick a date (who paid, amount, players) — blank if none
- Tab renamed to **Record Attendance**
- **Simple numeric player IDs** (#1, #2, #3…) shown on each player
- Everything branded **Gavakshi Baddies**

---

## Changing the URL later
Tap the **⚙ icon** (top-right) → confirm → paste new URL. Works on any browser.

---

## Your 5 sheets (auto-created)
| Sheet | Contents |
|---|---|
| Players | id, name, status, dates |
| Sessions | daily session + who paid |
| PlayLog | per-player cost per session |
| Expenses | incidental + group expenses |
| Meta | internal counter for player IDs |

*Never delete these sheets.*
