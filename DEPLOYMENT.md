# Deployment Guide

This takes the add-in from "runs on my laptop via `npm start`" to
"hosted on a real URL and available to the whole accounting team — **no local
dev server required**."

## Recommended: GitHub Pages (100% free)

Your repo already includes a GitHub Actions workflow that builds and deploys the
add-in for free. No credit card, no server to maintain.

### One-time setup (~2 minutes)

1. **Push this repo to GitHub** (if you haven't already).
2. On GitHub, open the repo → **Settings → Pages**.
3. Under **Build and deployment → Source**, choose **GitHub Actions**.
4. Go to **Actions** → run **Deploy add-in to GitHub Pages** (or push to `main`).
5. When it finishes, your add-in lives at:
   **`https://lilainx.github.io/BOC-Tool/`**

### Generate the production manifest

```bash
npm run make:github-manifest
```

This writes `manifest.prod.xml` pointed at your GitHub Pages URL. Validate it:

```bash
npx office-addin-manifest validate manifest.prod.xml
```

### Install in Excel (desktop or web — no local server)

1. In Excel: **Insert → Add-ins → My Add-ins → Upload My Add-in**
2. Select `manifest.prod.xml` → **Upload**
3. Done. The add-in loads from GitHub Pages forever — nothing runs on your laptop.

Every push to `main` auto-redeploys the latest code. Users get updates next time
they open Excel (no re-sideloading needed unless the manifest itself changes).

---

## Other free options

| Host | Cost | Best for |
|---|---|---|
| **GitHub Pages** | Free | Already set up in this repo |
| **Cloudflare Pages** | Free | Private company repos, custom domain |
| **Netlify** | Free tier | Alternative static host |

See below for Cloudflare / Azure if you prefer those over GitHub Pages.

---

## Desktop Excel vs Excel on the web

The same add-in works on **both**:

| Platform | Supported? | Notes |
|---|---|---|
| Excel on the web | Yes | What you tested during dev sideloading |
| Excel desktop (Windows) | Yes | Microsoft 365 subscription required for custom functions |
| Excel desktop (Mac) | Yes | Same requirement |

During development, `npm start` serves files from `https://localhost:3000`.
That is why the add-in stops working when you close the terminal — Excel is
still pointing at your laptop. Once you host the files on a permanent HTTPS URL
and sideload/distribute a production manifest, **nothing needs to run locally**.

The manifest already includes a `DesktopFormFactor` (ribbon button on the Home
tab) and `AllFormFactors` custom functions, so desktop and web use the same build.

---

There are two distinct jobs:

1. **Hosting** — put the built files (`dist/`) on an always-on HTTPS URL.
2. **Distribution** — sideload the production manifest (one user) or deploy via
   the Microsoft 365 Admin Center (whole team).

---

## Part 1 — Host the files

The add-in is a static site (HTML/JS/CSS) plus a JSON manifest, so any static
HTTPS host works.

### Option A — GitHub Pages (easiest if the repo is already on GitHub)

This repo includes a GitHub Actions workflow (`.github/workflows/deploy-pages.yml`)
that builds and deploys `dist/` on every push to `main`.

1. On GitHub: **Settings → Pages → Build and deployment → Source: GitHub Actions**.
2. Push to `main` (or run the workflow manually under **Actions**).
3. Your URL will be `https://<username>.github.io/BOC-Tool/` (project site).

Then generate the production manifest:

```bash
npm run make:prod-manifest https://<username>.github.io/BOC-Tool
npx office-addin-manifest validate manifest.prod.xml
```

### Option B — Cloudflare Pages (recommended for private company repos)

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

### Option C — Azure Static Web Apps

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

### Option D — Any static host

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
sideloads the production manifest themselves:

1. Send them `manifest.prod.xml` (the one pointing at your hosted URL, **not**
   the dev `manifest.xml` with localhost).
2. They sideload it in Excel:

   **Excel desktop (Windows or Mac)**
   - **Insert → Add-ins → My Add-ins**
   - **Upload My Add-in** (bottom of the panel) → browse to `manifest.prod.xml`
     → **Upload**
   - The **BOC Exchange Rates** button appears on the **Home** tab.

   **Excel on the web**
   - **Insert → Add-ins → Upload My Add-in** → browse to `manifest.prod.xml`
     → **Upload**

3. Because the files are hosted (not on localhost), it keeps working with no dev
   server and nothing running on anyone's machine.

That's the fastest way to get your first user going. Move to org distribution
below once you want it to appear automatically for the whole team.

> **Custom functions on desktop** require a **Microsoft 365** subscription
> (Excel build 16.0.11629+). The task pane works on older perpetual licenses,
> but `=BOC.RATE(...)` formulas need M365.

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
