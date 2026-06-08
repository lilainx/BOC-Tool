# Deployment Guide

This takes the add-in from "runs on my laptop via `npm run dev:server`" to
"hosted on a real URL and available to the whole accounting team."

There are two distinct jobs:

1. **Hosting** — put the built files (`dist/`) on an always-on HTTPS URL.
2. **Distribution** — make the add-in appear for your team via the Microsoft 365 Admin Center.

Hosting is something you (or whoever set up the project) can do. Distribution
usually requires a **Microsoft 365 / Global or Exchange admin** — likely your IT
team. Loop them in early.

---

## Part 1 — Host the files

The add-in is a static site (HTML/JS/CSS) plus a JSON manifest, so any static
HTTPS host works.

### Option A — Cloudflare Pages (recommended: free, private repos OK)

Free, works with **private** GitHub repos, and serves from the root domain (no
sub-path), which keeps the manifest simple. Cloudflare builds the project itself
from your repo — no GitHub Actions workflow or secrets to manage.

Prereq: the project pushed to a GitHub repo (private is fine) and a free
Cloudflare account (sign up at https://dash.cloudflare.com — no card required).

1. In the Cloudflare dashboard: **Workers & Pages → Create → Pages → Connect to Git**.
2. Authorize Cloudflare to access your GitHub account and pick the repo.
3. Set the build configuration:
   - **Framework preset**: None
   - **Build command**: `npm run build`
   - **Build output directory**: `dist`
4. **Save and Deploy.** Cloudflare builds and gives you a URL like
   **`https://boc-tool.pages.dev`** (files live at the root, e.g.
   `https://boc-tool.pages.dev/taskpane.html`).
5. Every push to the repo redeploys automatically.

> No public/private worry here — Cloudflare Pages is free either way, so keep the
> repo private if that's your preference for company code.

### Option B — Azure Static Web Apps

Prereqs: an Azure account (GeoComply almost certainly has one — ask IT), and
this project pushed to a GitHub repo.

1. In the [Azure Portal](https://portal.azure.com), create a **Static Web App**.
2. Connect it to this GitHub repo. For build settings:
   - **App location**: `/`
   - **Output location**: `dist`
   - Build command: `npm run build`
3. Azure provisions a URL like `https://boc-rates-xxxx.azurestaticapps.net`
   (or attach a custom domain).
4. Every push to the repo redeploys automatically.

### Option B — Any static host

Run `npm run build` and upload the contents of `dist/` to your host
(Netlify, Cloudflare Pages, an internal IIS/nginx server, etc.). The only
requirements: **HTTPS** and CORS not blocking the files.

---

## Part 2 — Point the manifest at your host

The committed `manifest.xml` is the **dev** manifest (hardcoded to
`https://localhost:3000`). Generate a production copy pointed at your real URL:

```bash
npm run make:prod-manifest https://YOUR-HOSTED-URL
```

This writes `manifest.prod.xml`. Validate it:

```bash
npx office-addin-manifest validate manifest.prod.xml
```

> Note: the production icon URLs now point at `YOUR-HOSTED-URL/assets/icon-*.png`,
> so make sure the `assets/` folder is included in what you deploy (it is, via the
> webpack copy step → `dist/assets/`).

---

## Part 3 — Get it to your user(s)

### Just one (or a few) users — no admin needed

For a single accountant, skip the M365 Admin Center entirely. That person
sideloads the production manifest themselves, exactly like during development:

1. Send them `manifest.prod.xml` (the one pointing at your hosted URL).
2. They open **Excel on the web** → **Insert → Add-ins → Upload My Add-in** →
   browse to `manifest.prod.xml` → **Upload**.
3. Because the files are hosted (not on localhost), it keeps working with no dev
   server and nothing running on anyone's machine.

That's the fastest way to get your first user going. Move to org distribution
below once you want it to appear automatically for the whole team.

### Whole team — via M365 admin

Hand `manifest.prod.xml` to your M365 admin (or do it yourself if you have rights):

1. Go to the [Microsoft 365 Admin Center](https://admin.microsoft.com) →
   **Settings → Integrated apps**.
2. Click **Upload custom apps** → **Office Add-in** → upload `manifest.prod.xml`.
3. Choose who gets it: **specific users/groups** (e.g. the Accounting group) or
   the whole org.
4. Deploy. It appears on the Excel **Home** tab for those users automatically —
   no per-person sideloading.

Propagation across all users can take up to ~24 hours on first deploy.

---

## Updating later

- **Code/UI changes**: push to the repo (Azure auto-redeploys) or re-upload
  `dist/`. Users get the new version on next load — no admin action needed,
  because the manifest just points at your URL.
- **Manifest changes** (new functions, new permissions, version bump): regenerate
  `manifest.prod.xml`, bump the `<Version>`, and re-upload it in Integrated Apps.

---

## Checklist

- [ ] Project pushed to a Git remote (GitHub)
- [ ] Files hosted on an always-on HTTPS URL
- [ ] `manifest.prod.xml` generated and validated against that URL
- [ ] `assets/` icons reachable at the hosted URL
- [ ] Manifest uploaded via M365 Integrated Apps, assigned to the Accounting group
- [ ] Verified by a teammate from a different machine (the real test)
