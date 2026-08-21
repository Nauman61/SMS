# Habib Shining Star School — Management System

A school management app covering staff records, admissions, a fee ledger
(with discounts, optional papers fund and late fee, and free/exempt
students), uniforms & books issuance, petty cash, and staff salary — with
Admin/Staff login, an audit trail on every payment entry, and Excel export.

---

## ⚠️ About data storage

This app can run in two modes:

- **Cloud mode (recommended)** — connected to a free Supabase database.
  Every device (admin's phone, staff's phone, the office PC) reads and
  writes the **same live data**. Set this up once, below.
- **Local mode (default if you skip setup)** — data is stored only in
  each device's own browser. Fine for testing or a single shared
  computer, but different devices will each have their **own separate
  copy** of the data, and clearing browser data erases it.

The sidebar footer always shows which mode you're currently in
("☁ Synced across devices" or "💻 Local only").

---

## 1. Requirements

- [Node.js](https://nodejs.org) version 18 or newer (includes `npm`)

Check your version:
```bash
node -v
```

## 2. Install dependencies

From inside this project folder:
```bash
npm install
```

## 3. Attach a shared database (Supabase) — do this before going live

This is what makes every device see the same live data. It's free and
takes about five minutes.

1. Go to [supabase.com](https://supabase.com) and sign up (free tier).
2. Click **New project**. Pick any name/region, set a database password
   (you won't need to remember it — Supabase manages this), and wait
   ~1 minute for it to finish setting up.
3. In your new project, open the **SQL Editor** (left sidebar) → **New
   query**, paste in the entire contents of `supabase/schema.sql` from
   this project, and click **Run**. This creates the one table the app
   needs.
4. Go to **Settings → API** in your Supabase project. You'll see a
   **Project URL** and an **anon public** API key — copy both.
5. In this project folder, copy `.env.example` to a new file named
   `.env`, and paste your values in:
   ```
   VITE_SUPABASE_URL=https://your-project-id.supabase.co
   VITE_SUPABASE_ANON_KEY=your-anon-key-here
   ```
6. That's it. Rebuild (`npm run build`) and redeploy — the app will
   automatically detect these values and switch to cloud mode.

**Security note:** the anon key above is meant to be included in the
built website's code (that's normal for Supabase), but it does mean
anyone who has it can read/write your data directly, bypassing the
app's Admin/Staff login. For a small private school tool this is a
reasonable trade-off, but don't publish your `.env` file or post your
Supabase URL/key publicly (e.g. in a public GitHub repo). If you need
stronger security later, that requires a custom backend with real
server-side authentication — let me know if you want to go there.

## 4. Run it locally (to test before going live)

```bash
npm run dev
```
This starts a local server, usually at `http://localhost:5173`. Open that
in your browser to use the app.

## 5. Build for production

```bash
npm run build
```
This creates a `dist/` folder containing the finished static website
(HTML, CSS, JS) — this is what you deploy. **Make sure your `.env` file
is in place before running this** — Vite bakes those values into the
build at this step.

You can preview the production build locally before deploying:
```bash
npm run preview
```

## 6. Deploy it live

Any static hosting service works, since this is a plain static site after
building. A few easy options:

### Option A — Netlify (free, easiest)
1. Go to [netlify.com](https://www.netlify.com) and sign up/log in.
2. Run `npm run build` locally.
3. Drag and drop the `dist/` folder onto Netlify's "Deploy" page.
4. Netlify gives you a live URL immediately. You can add a custom domain
   in Netlify's site settings later.

### Option B — Vercel
1. Go to [vercel.com](https://vercel.com) and sign up/log in.
2. Install the Vercel CLI: `npm i -g vercel`
3. Run `vercel` from this project folder and follow the prompts.

### Option C — Your own web hosting (e.g. shared hosting / cPanel)
1. Run `npm run build`.
2. Upload everything **inside** the `dist/` folder to your hosting's
   public web folder (often called `public_html`) via FTP or your host's
   file manager.
3. Visit your domain — the site should load directly.

### Option D — GitHub Pages
1. Push this project to a GitHub repository.
2. Run `npm run build`.
3. Use a GitHub Pages deploy action, or manually push the contents of
   `dist/` to a `gh-pages` branch.

**Using the drag-and-drop options above (A/C)?** Just make sure your
`.env` file exists locally with your Supabase values *before* running
`npm run build` — Vite reads it at build time.

**Connecting a Git repo instead (Netlify/Vercel's "import from GitHub"
flow)?** Don't commit your `.env` file. Instead add `VITE_SUPABASE_URL`
and `VITE_SUPABASE_ANON_KEY` as environment variables in that platform's
site settings (Netlify: Site settings → Environment variables. Vercel:
Project settings → Environment Variables), then redeploy.

---

## 7. Before you open it to real use

1. **Change the default passwords.** Log in as Admin with the default
   password `admin123`, then go to the sidebar → **⚙ Settings & access**
   and change both the Admin password and the Staff access code.
2. **Set your papers fund and late fee amounts** in the same Settings
   panel if the defaults (Rs 1000 / Rs 200) don't match your school.
3. **Do a test run**: add a test staff member, a test student, generate a
   month's fees, mark a payment, and download the Excel report to confirm
   everything looks right before entering real data.
4. **Back up regularly.** Since data lives only in the browser, use the
   dashboard's "⬇ Download full report (Excel)" button often — daily or
   weekly, depending on how much data changes — and save that file
   somewhere safe (email it to yourself, save to a USB drive, etc.).

---

## Mobile use

The app is fully responsive — on a phone the sidebar becomes a slide-out
menu (tap the ☰ button top-left), tables scroll horizontally with a swipe,
and forms stack into a single column.

It's also installable as an app on a phone's home screen (a Progressive
Web App): open the deployed site in Chrome (Android) or Safari (iPhone)
and use "Add to Home Screen" / "Install app" from the browser menu. It
will then open full-screen with its own icon, without browser address
bars, just like a native app.

**Note:** this is a installable *web app* (PWA), not a native iOS/Android
app submitted to the App Store or Play Store. It works great installed
from the browser, but if you specifically need Play Store/App Store
listings, that's a separate, larger project (e.g. wrapping this with
Capacitor or rebuilding in React Native) — happy to help with that later
if it becomes a requirement.

## Project structure

```
├── index.html            Entry HTML page
├── package.json          Dependencies and scripts
├── vite.config.js        Build tool configuration
├── .env.example           Template for your Supabase credentials
├── supabase/
│   └── schema.sql          SQL to run once in Supabase to create the data table
├── public/
│   ├── manifest.webmanifest   PWA install metadata
│   ├── sw.js                  Service worker (offline app-shell + installability)
│   ├── icon-192.png, icon-512.png, icon-maskable-512.png, apple-touch-icon.png, favicon.png
└── src/
    ├── main.jsx          Mounts the app into the page, registers the service worker
    ├── App.jsx           The entire application (all tabs/features)
    ├── storageShim.js    Talks to Supabase if configured, else falls back to local browser storage
    └── index.css         Minimal page styling
```

## Default login credentials

| Role  | Password    |
|-------|-------------|
| Admin | `admin123`  |
| Staff | `staff123`  |

**Change these immediately after your first login** — see step 7 above.
