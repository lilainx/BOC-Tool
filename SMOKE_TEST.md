# Smoke Test — BOC Exchange Rates Add-in

Run through this after `npm start` sideloads the add-in into Excel. Expected
values are real rates pulled from the live Bank of Canada API on 2026-06-08, so
you're checking for **correct** numbers, not just non-error output.

> ⚠️ Daily rates change. The "latest" values below will differ on your run —
> verify those against [the BOC site](https://www.bankofcanada.ca/rates/exchange/daily-exchange-rates/).
> The **dated / monthly / annual** values are historical and should match exactly.

---

## 0. Sideload & load

- [ ] `npm start` launches without errors; dev server reachable at `https://localhost:3000`
- [ ] Excel opens. If prompted to trust the localhost dev certificate, accept it.
- [ ] On the **Home** tab, a **"BOC Exchange Rates"** group with an **"Exchange Rates"** button appears.
- [ ] Clicking it opens the task pane on the right (red header, 4 tabs).

If the button/pane doesn't appear: check the dev-server console, and confirm
`dist/` contains `taskpane.html`, `functions.js`, `functions.json`, `functions.html`,
and `assets/icon-*.png`.

---

## 1. Custom functions (type these into cells)

> All functions live under the **`BOC`** namespace — type `=BOC.` to autocomplete.

| Cell formula | Expected result |
|---|---|
| `=BOC.RATE("USD")` | Today's USD/CAD rate (~1.39, compare to BOC site) |
| `=BOC.RATE("USD","2024-03-28")` | **1.3550** (exact) |
| `=BOC.RATE("USD","2024-03-31")` | **1.3550** — Mar 31 is a Sunday; should fall back to Fri Mar 28 |
| `=BOC.MONTHLY_AVG("USD",2024,3)` | **1.3539** (exact) |
| `=BOC.ANNUAL_AVG("USD",2023)` | **1.3497** (exact) |
| `=BOC.TO_CAD(1000,"USD","2024-03-28")` | **1355.00** (1000 × 1.3550) |
| `=BOC.FROM_CAD(1355,"USD","2024-03-28")` | **~1000.00** (1355 ÷ 1.3550) |
| `=BOC.RATE_DATE("USD","2024-03-31")` | **2024-03-28** (proves weekend fallback) |
| `=BOC.RATE("EUR")` | Today's EUR/CAD rate (sanity-check a 2nd currency) |

Checks:
- [ ] Typing `=BOC.` autocompletes the list of functions, each with a parameter tooltip.
- [ ] No `#VALUE!` / `#NAME?` errors.
- [ ] Dated/monthly/annual values match the **exact** numbers above.
- [ ] A function recalculates when you press F9 / edit the cell.
- [ ] Error path: `=BOC.RATE("ZZZ")` returns a clean error, not a crash.

---

## 2. Lookup tab

- [ ] **Latest**: pick USD → "Get Rate" → shows `1 USD = X.XXXXXX CAD` + a date + "Bank of Canada".
- [ ] **Specific Date**: enter `2024-03-28` → result shows **1.3550**.
- [ ] **Specific Date** weekend: enter `2024-03-31` → still **1.3550**, meta date reads `2024-03-28`.
- [ ] **Monthly Avg**: 2024 / March → **1.3539**.
- [ ] **Annual Avg**: 2023 → **1.3497**.
- [ ] **Insert into Cell**: select a cell, click it → the numeric rate lands in that cell, formatted to 6 decimals.

---

## 3. Convert tab

- [ ] `1000` USD → CAD (blank date = latest) → result reads `1,000.00 USD = X.XXXX CAD`.
- [ ] Swap button (⇄) flips From/To.
- [ ] CAD → USD direction works (division path).
- [ ] Cross-rate: USD → EUR (neither is CAD) returns a sensible number (goes via CAD).
- [ ] Same currency both sides → friendly error, no crash.
- [ ] Non-numeric / empty amount → friendly error.
- [ ] **Insert into Cell** drops the converted amount into the selected cell.

---

## 4. Trend tab

- [ ] USD, `2024-03-01` → `2024-03-28` → "Load Trend".
- [ ] Summary cards populate (High / Low / Avg / Change). Change is colored green/red by sign.
- [ ] The mini line chart renders with Y-axis rate labels and a few date labels on X.
- [ ] **Export to Sheet** creates a new worksheet `BOC USD Rates` with Date / Rate / Source columns,
      red header row, rates formatted to 6 decimals, columns auto-fit.
- [ ] start > end date → friendly error.

---

## 5. Formulas tab

- [ ] All 6 function reference cards render with signatures and examples.

---

## Known-good reference values (from live API, 2026-06-08)

| Series | Date | Value |
|---|---|---|
| FXUSDCAD (daily) | 2024-03-28 | 1.3550 |
| FXUSDCAD (daily) | 2024-03-27 | 1.3587 |
| FXMUSDCAD (monthly) | 2024-03 | 1.3539 |
| FXAUSDCAD (annual) | 2023 | 1.3497 |
| FXUSDCAD (latest) | 2026-06-05 | 1.3924 *(will have moved)* |
