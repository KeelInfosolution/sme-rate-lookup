# SME Compensation Matrix — Rate Lookup Tool

A single self-contained webpage. All 26,818 rate rows are embedded directly in
`index.html` as JSON — there is **no database, no backend, no build step, and
no external dependencies to install**. Whatever serves this one file as a
static webpage is all the "hosting" it needs.

File size: ~3 MB (this is normal — it's the embedded data, not bloat).

---

## Option 1 — Netlify (free, fastest, no account strictly required)

1. Go to **https://app.netlify.com/drop**
2. Drag `index.html` straight into the browser window
3. Netlify gives you a live public URL immediately (e.g.
   `https://random-name-123.netlify.app`)
4. Optional: create a free account first if you want to keep the site,
   rename the subdomain, or update the file later without generating a new
   random URL each time.

## Option 2 — Vercel (free)

1. Create a free account at **https://vercel.com**
2. New Project → Deploy → drag and drop the folder containing `index.html`
   (or use the Vercel CLI: `npx vercel` from inside this folder)
3. Vercel serves `index.html` automatically as the site root

## Option 3 — GitHub Pages (free, good if you already use GitHub)

1. Create a new GitHub repository
2. Upload `index.html` to the repo (rename is not required — GitHub Pages
   will serve `index.html` as the homepage automatically)
3. Repo Settings → Pages → set source to the `main` branch, root folder
4. GitHub gives you a URL like `https://yourusername.github.io/reponame/`

## Option 4 — Cloudflare Pages (free)

1. Create a free account at **https://pages.cloudflare.com**
2. Create a project → Direct Upload → drag in `index.html`
3. Cloudflare gives you a `https://yourproject.pages.dev` URL

## Option 5 — Your own server (Apache / Nginx / IIS / any static host)

This file needs zero special server configuration — just place it where your
web server serves static files.

**Nginx example:**
```nginx
server {
    listen 80;
    server_name your-domain.com;
    root /var/www/sme-rate-lookup;
    index index.html;
}
```
Then copy `index.html` into `/var/www/sme-rate-lookup/`.

**Apache:** drop `index.html` into your `public_html` (or equivalent)
directory — no `.htaccess` changes needed.

**Internal company server / intranet:** same idea — any server that can
serve a static `.html` file over HTTP works. No PHP, Node, or database
runtime is required.

## Option 6 — Quick local test before you deploy anywhere

From a terminal, inside this folder:
```bash
python3 -m http.server 8000
```
Then open `http://localhost:8000` in a browser. This is just for previewing
— it only works on your own machine, not for sharing with others.

---

## Updating the data later

You have two ways to refresh this tool whenever the underlying SME Excel
grid changes:

### Option A — Do it yourself (recommended, no need to come back to Claude)

This package includes everything needed to regenerate `index.html` yourself:

- `generate_lookup_tool.py` — a script that reads the current SME Excel file
  and rebuilds `index.html`
- `index_template.html` — the page shell (design + search logic) with a
  placeholder where the data gets inserted. **Keep this file in the same
  folder as the script** — don't rename or move it separately.

**Steps:**
1. Make sure you have Python 3 installed, and the `openpyxl` package:
   ```bash
   pip install openpyxl
   ```
2. Run the script, pointing it at your updated Excel file:
   ```bash
   python3 generate_lookup_tool.py "SME_Upload_Grid_UPDATED.xlsx"
   ```
3. This overwrites `index.html` in the same folder with fresh data pulled
   straight from the Excel file.
4. Re-upload/redeploy that `index.html` to your host (same process as your
   original deployment — e.g. drag it into Netlify, push to your GitHub
   Pages repo, or copy it onto your server, replacing the old file).

The script expects the Excel file's `SME Compensation Matrix` sheet to keep
the same column layout it has today: Grid ID, Company, Category, Product,
Occupancy Code, Occupancy Name, Recivable Net, Payable Net (columns A
through H). If those columns ever get reordered or renamed, the script will
need a small update — see Option B below.

Use `-o` if you want to write to a different filename instead of overwriting
`index.html` directly, e.g.:
```bash
python3 generate_lookup_tool.py "SME_Upload_Grid_UPDATED.xlsx" -o index_v2.html
```

### Option B — Ask Claude to do it

Upload the updated SME Excel file in a new message and ask Claude to
regenerate the hosting package. This is the better option if the sheet's
column structure has changed, if you want the design or search behavior
adjusted at the same time, or if you'd just rather not run scripts yourself.

## Notes

- Works fully offline once loaded — no API calls, no internet connection
  needed after the page loads (all data is embedded).
- No login, tracking, or analytics of any kind are included.
- If your host has a file-size limit under ~3 MB (some free tiers occasionally
  do for a single file, though all six options above comfortably support this),
  let me know and I can look at trimming or compressing the embedded dataset.
