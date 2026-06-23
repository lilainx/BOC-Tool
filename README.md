# BOC Exchange Rates — Excel Add-in

Bank of Canada exchange rate tools for accountants, built as a Microsoft Excel add-in.

## Features

| | |
|---|---|
| **Task pane** | Interactive rate lookups, currency converter, trend viewer with mini-chart |
| **Custom functions** | Use BOC rates directly in spreadsheet formulas |
| **Data source** | Bank of Canada Valet API — official rates, no API key required |

---

## Custom Functions

| Formula | Description |
|---|---|
| `=BOC.RATE("USD")` | Latest USD/CAD daily rate |
| `=BOC.RATE("EUR", "2024-03-31")` | EUR/CAD rate on a specific date |
| `=BOC.MONTHLY_AVG("USD", 2024, 3)` | March 2024 monthly average |
| `=BOC.ANNUAL_AVG("USD", 2023)` | 2023 annual average |
| `=BOC.TO_CAD(5000, "USD", "2024-12-31")` | Convert $5,000 USD → CAD |
| `=BOC.FROM_CAD(10000, "EUR")` | Convert $10,000 CAD → EUR |
| `=BOC.RATE_DATE("USD", "2024-12-28")` | Actual date of rate (e.g. nearest business day) |

> Functions live under the **`BOC`** namespace — type `=BOC.` in any cell to see the full list with autocomplete.
> All rates are expressed as **CAD per 1 unit of foreign currency**, matching the BOC convention.

---

## Development Setup

### Prerequisites
- Node.js 18+
- Microsoft Excel (desktop) or Microsoft 365

### Install dependencies

```bash
cd "BOC Tool"
npm install
```

### Run locally (development only)

```bash
npm start
```

This starts a dev server at `https://localhost:3000` and sideloads the add-in.
**The add-in only works while this server is running** — that is expected for
local development.

### Run without a local server (free — GitHub Pages)

1. Push to GitHub and enable **Settings → Pages → Source: GitHub Actions**
2. After the workflow deploys:
   ```bash
   npm run make:github-manifest
   ```
3. In Excel: **Insert → Add-ins → Upload My Add-in** → select `manifest.prod.xml`

Full details in **[DEPLOYMENT.md](./DEPLOYMENT.md)**.

### Build for production

```bash
npm run build
```

Output goes to `dist/`.

---

## Project Structure

```
BOC Tool/
├── src/
│   ├── boc-api.ts              # Bank of Canada Valet API client
│   ├── functions/
│   │   ├── functions.ts        # Custom Excel function definitions
│   │   └── functions.html      # Custom functions runtime shim
│   └── taskpane/
│       ├── taskpane.html       # Task pane markup
│       ├── taskpane.css        # Styles
│       └── taskpane.ts         # Task pane logic
├── assets/
│   └── functions.json          # Custom function metadata (for Excel)
├── manifest.xml                # Office Add-in manifest
├── webpack.config.js
├── tsconfig.json
└── package.json
```

---

## Deployment

See **[DEPLOYMENT.md](./DEPLOYMENT.md)** for full instructions (GitHub Pages,
Cloudflare, Azure, desktop sideloading, and org-wide M365 Admin Center deploy).

---

## Notes

- Rates are published on **business days only**. Date lookups on weekends/holidays automatically return the nearest prior business day rate.
- All rates are sourced from the [Bank of Canada Valet API](https://www.bankofcanada.ca/valet/docs) and are subject to their [terms of use](https://www.bankofcanada.ca/terms/).
