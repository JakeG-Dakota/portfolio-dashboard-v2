# Portfolio Dashboard — 17 Castlereagh & 1 Elizabeth Plaza

One URL. Both assets. PIN protected.

## What's in the dashboard

| View | Contents |
|------|----------|
| **Portfolio** | Combined KPIs, joint expiry profile table, mini stack bars |
| **Castlereagh — Dashboard** | KPIs, expiry profile, leases, market rents, by floor, critical dates, valuation |
| **Castlereagh — Stack Plan** | Interactive building stack with suite details panel |
| **Elizabeth — Dashboard** | Same sections as Castlereagh |
| **Elizabeth — Stack Plan** | Interactive building stack with suite details panel |

---

## Setup (one time, ~15 minutes)

### Step 1 — GitHub
1. Create a free account at github.com if you don't have one
2. **New repository** → name `portfolio-dashboard` → Create
3. Open **GitHub Desktop** → File → Clone Repository → `portfolio-dashboard`
4. Copy all files from this folder into the cloned repo folder
5. Copy your Excel files into `public/data/`:
   - `17_Castlereagh_v14.xlsx` → rename to **`17_castlereagh.xlsx`**
   - `1_Elizabeth_Plaza_v2.xlsx` → rename to **`1_elizabeth_plaza.xlsx`**
   - Open each in Excel and Save before copying (caches formula results)
6. GitHub Desktop: commit "Initial build" → **Push origin**

### Step 2 — Vercel
1. vercel.com → sign in with GitHub → **Add New Project** → Import `portfolio-dashboard`
2. Click **Deploy** (no settings changes needed)
3. **Settings → Environment Variables** → add:
   - Name: `NEXT_PUBLIC_DASHBOARD_PIN`
   - Value: your chosen PIN (e.g. `7291`)
4. **Redeploy** so the PIN takes effect
5. Your dashboard is live — share the URL

---

## Monthly update workflow

For **either** asset (or both):

1. Update your Excel file on your laptop as normal
2. **Open and Save in Excel** (important — caches formula results for the web reader)
3. Copy the file to `public/data/` in your repo folder, replacing the old version
4. GitHub Desktop: you'll see the file listed as changed
5. Write commit message (e.g. "May 2026 — Castlereagh update") → **Commit** → **Push**
6. Vercel rebuilds automatically in ~60 seconds
7. Refresh the dashboard URL — done

---

## File names (must be exact)
```
public/data/17_castlereagh.xlsx
public/data/1_elizabeth_plaza.xlsx
```

## Changing the PIN
Vercel → your project → Settings → Environment Variables → edit `NEXT_PUBLIC_DASHBOARD_PIN` → Redeploy

---

## Data reading approach
The dashboard reads **cached cell values** (not live formulas) from each workbook.
- Hardcoded columns (expiry source dates, PSM source, display names) are always reliable
- All KPIs, WALE, expiry buckets and stack plans are computed in JavaScript from the raw data
- Critical dates are calculated from lease expiry and rent review dates in Input Data — no Dashboard formulas needed
- Valuation data reads from the Valuations tab yellow input cells
