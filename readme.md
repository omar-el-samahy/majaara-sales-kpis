# Majaara KPI Dashboard | لوحة مؤشرات الأداء

Bilingual (English / العربية) KPI dashboard implementing the official Majaara sales KPI system:

- **Sales Representative scorecard** — 100 pts (revenue, effective meetings, closure rate, cycle efficiency, CRM, CSAT)
- **Team Leader scorecard** — 110 pts (team goals, leader meetings, member development, team closure, CRM tools, initiatives)

Data lives in a **Supabase cloud database**, so daily entries are permanent and available for long-term analysis from any PC.

---

## Files

| File | Purpose |
|---|---|
| `index.html` | The full web dashboard (open in any browser) |
| `supabase_setup.sql` | Run once in Supabase to create the tables |
| `KPI_Scorecard.xlsx` | Standalone Excel version with auto-formulas (offline use) |
| `tools/gen_excel.py` | Script that regenerates the Excel file |

---

## Setup (one time, ~10 minutes)

### 1. Create the Supabase project
1. Go to [supabase.com](https://supabase.com) → **Sign up** (free).
2. Click **New project** → name it `majaara-kpi`, pick a password, choose the region closest to you → **Create**.
3. Wait ~2 minutes for provisioning.

### 2. Create the tables
1. In the left sidebar open **SQL Editor** → **New query**.
2. Open `supabase_setup.sql` (this folder), copy **all** its contents, paste into the editor.
3. Click **Run**. You should see "Success".

### 3. Get your credentials
1. Go to **Project Settings** (gear icon) → **API**.
2. Copy **Project URL** (looks like `https://abcd1234.supabase.co`).
3. Copy the **anon public** key (a long `eyJ...` string). Do NOT use the `service_role` key.

### 4. Connect the dashboard
1. Open `index.html` in Chrome/Edge (double-click works; if your browser blocks it, run `python -m http.server` in this folder and visit `http://localhost:8000`).
2. The **Connect Database** dialog opens on first launch — paste the URL and anon key → **Test Connection** → **Save**.

Done. Credentials are stored in that browser only.

> **Security note:** the anon key is public by design; anyone holding both URL + key can read/write data. For internal HR use this is usually acceptable. To lock it down later, enable Supabase Auth and tighten the RLS policies in `supabase_setup.sql`.

---

## Daily workflow

1. **Daily Entry** tab → pick employee + date → fill the day's numbers → **Save Entry**.
   - *Rep fields:* sales amount, meetings held, deals closed, avg cycle days (optional), avg customer rating 0–10 (optional), CRM fully updated checkbox.
   - *Leader fields:* team sales, meetings led, team totals, members achieving goals, CRM status & initiative status dropdowns.
2. **Scorecards** tab → pick month + employee → **Calculate**. Shows every KPI's actual vs target, tier reached, points, total score ring and rating badge (ممتاز / جيد جداً / يحتاج تحسين / مراجعة).
3. **Analysis & Export** tab:
   - Monthly score trend chart per employee (dots colored by rating band).
   - Month ranking for all reps and leaders.
   - **Export to Excel** (current ranking) and **Backup JSON** (full raw-data backup — do this monthly).

## Employees tab

Add each rep/leader once with their **monthly targets**: sales target, meetings target, target cycle days (rep only). Edit/delete anytime; deleting removes their entries too.

---

## Scoring rules (as per the official document)

Tiers are implemented verbatim from the PDF. Documented interpretations where the source was ambiguous or inconsistent:

- **Revenue "exceed by ≥15%" tier awards 40 pts** although the section summary caps revenue at 35 — kept verbatim, so a rep can reach 105. Rating thresholds still apply as written (>85 excellent etc.).
- **Sales cycle:** faster than target = 15 pts, ≈ target (±0.5%) = 10, up to 120% of target = 8, beyond = 3 (the 100–110% gap is folded into the 8-pt tier).
- **CRM compliance <60%** and **no recorded meetings (closure rate)** award 0/5 respectively — the PDF lists no tier below its lowest stated band.
- **CRM status & initiative status** (leader) are admin assessments entered with the daily entry; the latest value of each month is scored.

## Excel version

`KPI_Scorecard.xlsx` mirrors the same logic offline: fill the yellow cells, points/rating compute automatically with color-coded rating bands and dropdown lists on the Team Leader sheet. One workbook evaluates one person-month — duplicate the sheet file per evaluation or copy the sheet.

To regenerate after editing: `python tools/gen_excel.py`
