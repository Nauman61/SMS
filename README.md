# Habib Shining Star School — Management System

A school management app covering staff records, admissions, a fee ledger
(with discounts, optional papers fund and late fee, and free/exempt
students), uniforms & books issuance, petty cash, and staff salary — with
Admin/Staff login, an audit trail on every payment entry, and Excel export.

---

## ⚠️ Please read before going live

This build stores all data in **your browser's local storage** on
whichever computer you run it from. That means:

- Data does **not** sync between different computers, laptops, or phones.
  If your front desk uses one PC and the office uses another, they will
  each have their **own separate copy** of the data.
- Clearing browser history/site data, or switching browsers, will **erase
  all school data** on that device.
- Anyone with physical/browser access to that computer can open the
  browser's developer tools and read the stored data (including the
  admin/staff passwords), since there is no real server-side security.

This is a perfectly reasonable way to get live quickly if the school will
run everything from **one shared computer** (e.g. the office PC), and you
export an Excel backup regularly using the "Download full report" button.

**If you need multiple computers/devices to see the same live data**, you
need a real backend with a shared database instead of this local-storage
version. That's a bigger project — let me know and I can help build it
(for example, a small Node.js + PostgreSQL or Firebase backend) as a next
step.

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

## 3. Run it locally (to test before going live)

```bash
npm run dev
```
This starts a local server, usually at `http://localhost:5173`. Open that
in your browser to use the app.

## 4. Build for production

```bash
npm run build
```
This creates a `dist/` folder containing the finished static website
(HTML, CSS, JS) — this is what you deploy.

You can preview the production build locally before deploying:
```bash
npm run preview
```

## 5. Deploy it live

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

---

## 6. Before you open it to real use

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

## Project structure

```
├── index.html          Entry HTML page
├── package.json         Dependencies and scripts
├── vite.config.js        Build tool configuration
└── src/
    ├── main.jsx          Mounts the app into the page
    ├── App.jsx           The entire application (all tabs/features)
    ├── storageShim.js    Makes the app's data storage work in a normal browser
    └── index.css         Minimal page styling
```

## Default login credentials

| Role  | Password    |
|-------|-------------|
| Admin | `admin123`  |
| Staff | `staff123`  |

**Change these immediately after your first login** — see step 6 above.
