/**
 * Bank of Canada Valet API client
 * Docs: https://www.bankofcanada.ca/valet/docs
 *
 * All rates are expressed as: 1 unit of foreign currency = X CAD
 */

const BASE_URL = "https://www.bankofcanada.ca/valet";

export type FrequencyPrefix = "FX" | "FXM" | "FXA";

/** All currencies published by the Bank of Canada */
export const SUPPORTED_CURRENCIES: Record<string, string> = {
  USD: "US Dollar",
  EUR: "Euro",
  GBP: "British Pound",
  JPY: "Japanese Yen",
  CHF: "Swiss Franc",
  AUD: "Australian Dollar",
  HKD: "Hong Kong Dollar",
  MXN: "Mexican Peso",
  BRL: "Brazilian Real",
  CNY: "Chinese Renminbi",
  INR: "Indian Rupee",
  IDR: "Indonesian Rupiah",
  KRW: "South Korean Won",
  MYR: "Malaysian Ringgit",
  NZD: "New Zealand Dollar",
  NOK: "Norwegian Krone",
  PEN: "Peruvian Sol",
  SAR: "Saudi Riyal",
  SGD: "Singapore Dollar",
  ZAR: "South African Rand",
  SEK: "Swedish Krona",
  TWD: "New Taiwan Dollar",
  THB: "Thai Baht",
  TRY: "Turkish Lira",
};

export interface RateObservation {
  date: string;       // YYYY-MM-DD
  rate: number;       // CAD per 1 unit of foreign currency
  currency: string;
}

interface ValetObservation {
  d: string;
  [series: string]: { v: string } | string;
}

interface ValetResponse {
  observations: ValetObservation[];
}

function seriesName(prefix: FrequencyPrefix, currency: string): string {
  return `${prefix}${currency.toUpperCase()}CAD`;
}

async function fetchValet(url: string): Promise<ValetResponse> {
  const res = await fetch(url);
  if (!res.ok) {
    throw new Error(`BOC API error ${res.status}: ${res.statusText}`);
  }
  const data = (await res.json()) as ValetResponse;
  if (!data.observations || data.observations.length === 0) {
    throw new Error("No data returned for the requested date/currency.");
  }
  return data;
}

function parseObservation(
  obs: ValetObservation,
  series: string,
  currency: string
): RateObservation | null {
  const raw = obs[series];
  if (!raw || typeof raw !== "object") return null;
  const v = parseFloat((raw as { v: string }).v);
  if (isNaN(v)) return null;
  return { date: obs.d, rate: v, currency: currency.toUpperCase() };
}

/**
 * Get the most recent daily rate for a currency.
 * Returns the latest available observation (rates are only published on business days).
 */
export async function getLatestRate(currency: string): Promise<RateObservation> {
  const series = seriesName("FX", currency);
  // `recent=1` returns only the most recent observation. (Note: the BOC Valet
  // API does NOT support a `limit` parameter — it returns HTTP 400.)
  const data = await fetchValet(
    `${BASE_URL}/observations/${series}/json?recent=1`
  );
  const obs = data.observations[0];
  const result = parseObservation(obs, series, currency);
  if (!result) throw new Error(`Rate not available for ${currency}`);
  return result;
}

/**
 * Get the daily rate for a currency on a specific date.
 * The BOC only publishes rates on business days; if the date falls on a weekend
 * or holiday, this returns the nearest prior business day.
 */
export async function getRateOnDate(
  currency: string,
  date: string // YYYY-MM-DD
): Promise<RateObservation> {
  const series = seriesName("FX", currency);
  // Fetch a small window ending on the requested date to handle weekends/holidays
  const target = new Date(date + "T12:00:00Z");
  const windowStart = new Date(target);
  windowStart.setDate(windowStart.getDate() - 7);
  const startStr = windowStart.toISOString().slice(0, 10);

  // order_dir=desc puts the newest in-range observation first; we take [0].
  // (No `limit` param — it is unsupported and returns HTTP 400.)
  const data = await fetchValet(
    `${BASE_URL}/observations/${series}/json?start_date=${startStr}&end_date=${date}&order_dir=desc`
  );
  const obs = data.observations[0];
  const result = parseObservation(obs, series, currency);
  if (!result) throw new Error(`Rate not available for ${currency} on or before ${date}`);
  return result;
}

/**
 * Get the monthly average rate for a currency.
 * @param year  e.g. 2024
 * @param month 1–12
 */
export async function getMonthlyAvgRate(
  currency: string,
  year: number,
  month: number
): Promise<RateObservation> {
  const series = seriesName("FXM", currency);
  const mm = String(month).padStart(2, "0");
  const dateStr = `${year}-${mm}-01`;
  // Monthly averages are keyed to the first of the month in BOC data
  const endDate = `${year}-${mm}-28`;

  const data = await fetchValet(
    `${BASE_URL}/observations/${series}/json?start_date=${dateStr}&end_date=${endDate}&order_dir=desc`
  );
  const obs = data.observations[0];
  const result = parseObservation(obs, series, currency);
  if (!result) throw new Error(`Monthly average not available for ${currency} ${year}-${mm}`);
  return result;
}

/**
 * Get the annual average rate for a currency.
 */
export async function getAnnualAvgRate(
  currency: string,
  year: number
): Promise<RateObservation> {
  const series = seriesName("FXA", currency);
  const data = await fetchValet(
    `${BASE_URL}/observations/${series}/json?start_date=${year}-01-01&end_date=${year}-12-31&order_dir=desc`
  );
  const obs = data.observations[0];
  const result = parseObservation(obs, series, currency);
  if (!result) throw new Error(`Annual average not available for ${currency} ${year}`);
  return result;
}

/**
 * Get daily rates for a currency over a date range (for trend views).
 */
export async function getRateRange(
  currency: string,
  startDate: string,
  endDate: string
): Promise<RateObservation[]> {
  const series = seriesName("FX", currency);
  const data = await fetchValet(
    `${BASE_URL}/observations/${series}/json?start_date=${startDate}&end_date=${endDate}`
  );
  return data.observations
    .map((obs) => parseObservation(obs, series, currency))
    .filter((r): r is RateObservation => r !== null);
}

/**
 * Convert an amount from a foreign currency to CAD using a specific rate.
 */
export function convertToCAD(amount: number, rate: number): number {
  return amount * rate;
}

/**
 * Convert an amount from CAD to a foreign currency using a specific rate.
 */
export function convertFromCAD(amount: number, rate: number): number {
  if (rate === 0) throw new Error("Rate cannot be zero");
  return amount / rate;
}
