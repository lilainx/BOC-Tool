import "./taskpane.css";
import {
  getLatestRate,
  getRateOnDate,
  getMonthlyAvgRate,
  getAnnualAvgRate,
  getRateRange,
  convertToCAD,
  convertFromCAD,
  SUPPORTED_CURRENCIES,
  RateObservation,
} from "../boc-api";

/* global Office, Excel */

Office.onReady(() => {
  initCurrencySelects();
  initTabs();
  initLookup();
  initConvert();
  initTrend();
  setDefaultDates();
});

// ─── Currency selects ─────────────────────────────────────────────────────────

function initCurrencySelects() {
  const selects = [
    document.getElementById("lookup-currency") as HTMLSelectElement,
    document.getElementById("convert-from") as HTMLSelectElement,
    document.getElementById("convert-to") as HTMLSelectElement,
    document.getElementById("trend-currency") as HTMLSelectElement,
  ];

  const cad = document.createElement("option");
  cad.value = "CAD";
  cad.textContent = "CAD — Canadian Dollar";

  selects.forEach((sel) => {
    if (!sel) return;
    // Add CAD only for convert selects (not lookup/trend)
    if (sel.id === "convert-from" || sel.id === "convert-to") {
      sel.appendChild(cad.cloneNode(true));
    }
    for (const [code, name] of Object.entries(SUPPORTED_CURRENCIES)) {
      const opt = document.createElement("option");
      opt.value = code;
      opt.textContent = `${code} — ${name}`;
      if (code === "USD") opt.selected = true;
      sel.appendChild(opt);
    }
  });

  // Default convert-to to CAD
  const convertTo = document.getElementById("convert-to") as HTMLSelectElement;
  if (convertTo) convertTo.value = "CAD";
}

function setDefaultDates() {
  const today = new Date().toISOString().slice(0, 10);
  const monthAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString().slice(0, 10);

  (document.getElementById("lookup-date") as HTMLInputElement).value = today;
  (document.getElementById("convert-date") as HTMLInputElement).value = "";
  (document.getElementById("trend-end") as HTMLInputElement).value = today;
  (document.getElementById("trend-start") as HTMLInputElement).value = monthAgo;

  const now = new Date();
  (document.getElementById("lookup-year") as HTMLInputElement).value = String(now.getFullYear());
  (document.getElementById("lookup-month") as HTMLSelectElement).value = String(now.getMonth() + 1);
  (document.getElementById("lookup-annual-year") as HTMLInputElement).value = String(now.getFullYear());
}

// ─── Tabs ─────────────────────────────────────────────────────────────────────

function initTabs() {
  document.querySelectorAll(".tab").forEach((tab) => {
    tab.addEventListener("click", () => {
      const target = (tab as HTMLElement).dataset.tab!;
      document.querySelectorAll(".tab").forEach((t) => t.classList.remove("active"));
      document.querySelectorAll(".tab-content").forEach((c) => c.classList.remove("active"));
      tab.classList.add("active");
      document.getElementById(`tab-${target}`)?.classList.add("active");
    });
  });
}

// ─── Lookup tab ───────────────────────────────────────────────────────────────

function initLookup() {
  const radios = document.querySelectorAll<HTMLInputElement>('input[name="lookup-type"]');
  radios.forEach((r) => {
    r.addEventListener("change", () => {
      document.getElementById("lookup-date-row")?.classList.toggle("hidden", r.value !== "daily");
      document.getElementById("lookup-monthly-row")?.classList.toggle("hidden", r.value !== "monthly");
      document.getElementById("lookup-annual-row")?.classList.toggle("hidden", r.value !== "annual");
    });
  });

  document.getElementById("lookup-btn")?.addEventListener("click", async () => {
    const currency = (document.getElementById("lookup-currency") as HTMLSelectElement).value;
    const type = (document.querySelector<HTMLInputElement>('input[name="lookup-type"]:checked'))?.value ?? "latest";

    hideResult("lookup");
    showLoading();

    try {
      let obs: RateObservation;
      if (type === "latest") {
        obs = await getLatestRate(currency);
      } else if (type === "daily") {
        const date = (document.getElementById("lookup-date") as HTMLInputElement).value;
        if (!date) throw new Error("Please select a date.");
        obs = await getRateOnDate(currency, date);
      } else if (type === "monthly") {
        const year = parseInt((document.getElementById("lookup-year") as HTMLInputElement).value);
        const month = parseInt((document.getElementById("lookup-month") as HTMLSelectElement).value);
        obs = await getMonthlyAvgRate(currency, year, month);
      } else {
        const year = parseInt((document.getElementById("lookup-annual-year") as HTMLInputElement).value);
        obs = await getAnnualAvgRate(currency, year);
      }
      showLookupResult(obs, type);
    } catch (e) {
      showError("lookup", (e as Error).message);
    } finally {
      hideLoading();
    }
  });

  document.getElementById("lookup-insert-btn")?.addEventListener("click", () => {
    const rateText = document.getElementById("lookup-rate-display")?.dataset.rate;
    if (rateText) insertValueIntoCell(parseFloat(rateText));
  });
}

function showLookupResult(obs: RateObservation, type: string) {
  const typeLabel: Record<string, string> = {
    latest: "Latest daily rate",
    daily: "Daily rate",
    monthly: "Monthly average",
    annual: "Annual average",
  };
  const rateDisplay = document.getElementById("lookup-rate-display")!;
  const metaDisplay = document.getElementById("lookup-meta-display")!;

  rateDisplay.textContent = `1 ${obs.currency} = ${obs.rate.toFixed(6)} CAD`;
  rateDisplay.dataset.rate = String(obs.rate);
  metaDisplay.textContent = `${typeLabel[type] ?? type} · Date: ${obs.date} · Source: Bank of Canada`;

  document.getElementById("lookup-result")?.classList.remove("hidden");
  document.getElementById("lookup-error")?.classList.add("hidden");
}

// ─── Convert tab ──────────────────────────────────────────────────────────────

function initConvert() {
  document.getElementById("convert-swap-btn")?.addEventListener("click", () => {
    const from = document.getElementById("convert-from") as HTMLSelectElement;
    const to = document.getElementById("convert-to") as HTMLSelectElement;
    const tmp = from.value;
    from.value = to.value;
    to.value = tmp;
  });

  document.getElementById("convert-btn")?.addEventListener("click", async () => {
    const amount = parseFloat((document.getElementById("convert-amount") as HTMLInputElement).value);
    const from = (document.getElementById("convert-from") as HTMLSelectElement).value;
    const to = (document.getElementById("convert-to") as HTMLSelectElement).value;
    const date = (document.getElementById("convert-date") as HTMLInputElement).value || undefined;

    if (isNaN(amount)) { showError("convert", "Please enter a valid amount."); return; }
    if (from === to) { showError("convert", "From and To currencies must be different."); return; }

    hideResult("convert");
    showLoading();

    try {
      let result: number;
      let rateObs: RateObservation;
      let rateDate: string;

      if (from === "CAD") {
        rateObs = date ? await getRateOnDate(to, date) : await getLatestRate(to);
        result = convertFromCAD(amount, rateObs.rate);
        rateDate = rateObs.date;
        const display = document.getElementById("convert-result-display")!;
        display.textContent = `${amount.toLocaleString("en-CA", { minimumFractionDigits: 2 })} CAD = ${result.toLocaleString("en-CA", { minimumFractionDigits: 4 })} ${to}`;
        display.dataset.rate = String(result);
      } else if (to === "CAD") {
        rateObs = date ? await getRateOnDate(from, date) : await getLatestRate(from);
        result = convertToCAD(amount, rateObs.rate);
        rateDate = rateObs.date;
        const display = document.getElementById("convert-result-display")!;
        display.textContent = `${amount.toLocaleString("en-CA", { minimumFractionDigits: 2 })} ${from} = ${result.toLocaleString("en-CA", { minimumFractionDigits: 4 })} CAD`;
        display.dataset.rate = String(result);
      } else {
        // Cross-rate via CAD
        const fromObs = date ? await getRateOnDate(from, date) : await getLatestRate(from);
        const toObs = date ? await getRateOnDate(to, date) : await getLatestRate(to);
        const amountInCAD = convertToCAD(amount, fromObs.rate);
        result = convertFromCAD(amountInCAD, toObs.rate);
        rateDate = fromObs.date;
        const display = document.getElementById("convert-result-display")!;
        display.textContent = `${amount.toLocaleString("en-CA", { minimumFractionDigits: 2 })} ${from} = ${result.toLocaleString("en-CA", { minimumFractionDigits: 4 })} ${to}`;
        display.dataset.rate = String(result);
      }

      const metaDisplay = document.getElementById("convert-meta-display")!;
      metaDisplay.textContent = `Rate date: ${rateDate} · via Bank of Canada`;
      document.getElementById("convert-result")?.classList.remove("hidden");
      document.getElementById("convert-error")?.classList.add("hidden");
    } catch (e) {
      showError("convert", (e as Error).message);
    } finally {
      hideLoading();
    }
  });

  document.getElementById("convert-insert-btn")?.addEventListener("click", () => {
    const rateText = document.getElementById("convert-result-display")?.dataset.rate;
    if (rateText) insertValueIntoCell(parseFloat(rateText));
  });
}

// ─── Trend tab ────────────────────────────────────────────────────────────────

let trendData: RateObservation[] = [];

function initTrend() {
  document.getElementById("trend-btn")?.addEventListener("click", async () => {
    const currency = (document.getElementById("trend-currency") as HTMLSelectElement).value;
    const start = (document.getElementById("trend-start") as HTMLInputElement).value;
    const end = (document.getElementById("trend-end") as HTMLInputElement).value;

    if (!start || !end) { showError("trend", "Please select a start and end date."); return; }
    if (start > end) { showError("trend", "Start date must be before end date."); return; }

    document.getElementById("trend-summary")?.classList.add("hidden");
    document.getElementById("trend-chart")?.classList.add("hidden");
    document.getElementById("trend-export-btn")?.classList.add("hidden");
    showLoading();

    try {
      trendData = await getRateRange(currency, start, end);
      if (trendData.length === 0) throw new Error("No data available for this date range.");

      const rates = trendData.map((d) => d.rate);
      const high = Math.max(...rates);
      const low = Math.min(...rates);
      const avg = rates.reduce((a, b) => a + b, 0) / rates.length;
      const change = rates[rates.length - 1] - rates[0];
      const changePct = (change / rates[0]) * 100;

      document.getElementById("trend-high")!.textContent = high.toFixed(4);
      document.getElementById("trend-low")!.textContent = low.toFixed(4);
      document.getElementById("trend-avg")!.textContent = avg.toFixed(4);
      const changeEl = document.getElementById("trend-change")!;
      changeEl.textContent = `${change >= 0 ? "+" : ""}${change.toFixed(4)} (${changePct.toFixed(2)}%)`;
      changeEl.style.color = change >= 0 ? "var(--green)" : "var(--error)";

      document.getElementById("trend-summary")?.classList.remove("hidden");
      drawTrendChart(trendData);
      document.getElementById("trend-export-btn")?.classList.remove("hidden");
      document.getElementById("trend-error")?.classList.add("hidden");
    } catch (e) {
      showError("trend", (e as Error).message);
    } finally {
      hideLoading();
    }
  });

  document.getElementById("trend-export-btn")?.addEventListener("click", () => {
    exportTrendToSheet(trendData);
  });
}

function drawTrendChart(data: RateObservation[]) {
  const canvas = document.getElementById("trend-chart") as HTMLCanvasElement;
  if (!canvas) return;
  canvas.classList.remove("hidden");

  const dpr = window.devicePixelRatio || 1;
  const rect = canvas.getBoundingClientRect();
  canvas.width = rect.width * dpr;
  canvas.height = 160 * dpr;
  const ctx = canvas.getContext("2d")!;
  ctx.scale(dpr, dpr);

  const W = rect.width;
  const H = 160;
  const PAD = { top: 12, right: 10, bottom: 24, left: 44 };
  const chartW = W - PAD.left - PAD.right;
  const chartH = H - PAD.top - PAD.bottom;

  const rates = data.map((d) => d.rate);
  const minRate = Math.min(...rates) * 0.9995;
  const maxRate = Math.max(...rates) * 1.0005;

  const xScale = (i: number) => PAD.left + (i / (data.length - 1)) * chartW;
  const yScale = (r: number) => PAD.top + chartH - ((r - minRate) / (maxRate - minRate)) * chartH;

  // Clear
  ctx.clearRect(0, 0, W, H);

  // Grid lines
  ctx.strokeStyle = "#e5e7eb";
  ctx.lineWidth = 1;
  for (let i = 0; i <= 4; i++) {
    const y = PAD.top + (i / 4) * chartH;
    ctx.beginPath(); ctx.moveTo(PAD.left, y); ctx.lineTo(W - PAD.right, y); ctx.stroke();
    const val = maxRate - (i / 4) * (maxRate - minRate);
    ctx.fillStyle = "#9ca3af";
    ctx.font = "10px 'Segoe UI', sans-serif";
    ctx.textAlign = "right";
    ctx.fillText(val.toFixed(4), PAD.left - 4, y + 3);
  }

  // Fill area under line
  ctx.beginPath();
  ctx.moveTo(xScale(0), yScale(rates[0]));
  for (let i = 1; i < data.length; i++) {
    ctx.lineTo(xScale(i), yScale(rates[i]));
  }
  ctx.lineTo(xScale(data.length - 1), PAD.top + chartH);
  ctx.lineTo(PAD.left, PAD.top + chartH);
  ctx.closePath();
  const grad = ctx.createLinearGradient(0, PAD.top, 0, PAD.top + chartH);
  grad.addColorStop(0, "rgba(200,16,46,.18)");
  grad.addColorStop(1, "rgba(200,16,46,0)");
  ctx.fillStyle = grad;
  ctx.fill();

  // Line
  ctx.beginPath();
  ctx.strokeStyle = "#c8102e";
  ctx.lineWidth = 1.5;
  ctx.lineJoin = "round";
  ctx.moveTo(xScale(0), yScale(rates[0]));
  for (let i = 1; i < data.length; i++) {
    ctx.lineTo(xScale(i), yScale(rates[i]));
  }
  ctx.stroke();

  // X-axis date labels (first, middle, last)
  ctx.fillStyle = "#9ca3af";
  ctx.font = "10px 'Segoe UI', sans-serif";
  ctx.textAlign = "center";
  [[0, data[0].date], [Math.floor(data.length / 2), data[Math.floor(data.length / 2)].date], [data.length - 1, data[data.length - 1].date]].forEach(([i, d]) => {
    ctx.fillText(String(d).slice(5), xScale(Number(i)), H - 6);
  });
}

// ─── Export to sheet ──────────────────────────────────────────────────────────

async function exportTrendToSheet(data: RateObservation[]) {
  if (!data.length) return;
  showLoading();
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.add(`BOC ${data[0].currency} Rates`);

      // Headers
      const header = sheet.getRange("A1:C1");
      header.values = [["Date", `${data[0].currency}/CAD Rate`, "Source"]];
      header.format.font.bold = true;
      header.format.fill.color = "#c8102e";
      header.format.font.color = "#ffffff";

      // Data rows
      const rows = data.map((d) => [d.date, d.rate, "Bank of Canada"]);
      const dataRange = sheet.getRange(`A2:C${rows.length + 1}`);
      dataRange.values = rows;

      // Format rate column
      sheet.getRange(`B2:B${rows.length + 1}`).numberFormat = [["0.000000"]];

      // Auto-fit
      sheet.getUsedRange().format.autofitColumns();
      sheet.activate();

      await context.sync();
    });
  } catch (e) {
    showError("trend", `Export failed: ${(e as Error).message}`);
  } finally {
    hideLoading();
  }
}

// ─── Insert into active cell ──────────────────────────────────────────────────

async function insertValueIntoCell(value: number) {
  try {
    await Excel.run(async (context) => {
      const cell = context.workbook.getSelectedRange();
      cell.values = [[value]];
      cell.numberFormat = [["0.000000"]];
      await context.sync();
    });
  } catch {
    // If not in a spreadsheet context, silently ignore
  }
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function hideResult(prefix: string) {
  document.getElementById(`${prefix}-result`)?.classList.add("hidden");
  document.getElementById(`${prefix}-error`)?.classList.add("hidden");
}

function showError(prefix: string, msg: string) {
  const el = document.getElementById(`${prefix}-error`);
  if (el) { el.textContent = msg; el.classList.remove("hidden"); }
  document.getElementById(`${prefix}-result`)?.classList.add("hidden");
}

function showLoading() {
  document.getElementById("loading-overlay")?.classList.remove("hidden");
}

function hideLoading() {
  document.getElementById("loading-overlay")?.classList.add("hidden");
}
