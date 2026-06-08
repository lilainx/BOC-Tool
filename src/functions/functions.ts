/* global CustomFunctions */

// CustomFunctions is injected by the Office.js runtime
declare const CustomFunctions: {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  associate(id: string, func: (...args: any[]) => any): void;
};

import {
  getLatestRate,
  getRateOnDate,
  getMonthlyAvgRate,
  getAnnualAvgRate,
  convertToCAD,
  convertFromCAD,
} from "../boc-api";

// All functions are invoked in Excel under the "BOC" namespace, e.g. =BOC.RATE("USD").

// ─── Rate Lookups ─────────────────────────────────────────────────────────────

/**
 * Returns the CAD exchange rate for a currency from the Bank of Canada.
 * @customfunction RATE
 * @param currency Currency code, e.g. "USD", "EUR", "GBP"
 * @param [date] Optional date in YYYY-MM-DD format. Omit for today's (latest) rate.
 * @returns Exchange rate: how many CAD per 1 unit of the given currency
 * @helpurl https://www.bankofcanada.ca/rates/exchange/daily-exchange-rates/
 */
export async function RATE(currency: string, date?: string): Promise<number> {
  const obs = date
    ? await getRateOnDate(currency, date)
    : await getLatestRate(currency);
  return obs.rate;
}

/**
 * Returns the monthly average CAD exchange rate from the Bank of Canada.
 * @customfunction MONTHLY_AVG
 * @param currency Currency code, e.g. "USD"
 * @param year Four-digit year, e.g. 2024
 * @param month Month number 1–12
 * @returns Monthly average exchange rate (CAD per 1 unit of currency)
 */
export async function MONTHLY_AVG(
  currency: string,
  year: number,
  month: number
): Promise<number> {
  const obs = await getMonthlyAvgRate(currency, year, month);
  return obs.rate;
}

/**
 * Returns the annual average CAD exchange rate from the Bank of Canada.
 * @customfunction ANNUAL_AVG
 * @param currency Currency code, e.g. "USD"
 * @param year Four-digit year, e.g. 2024
 * @returns Annual average exchange rate (CAD per 1 unit of currency)
 */
export async function ANNUAL_AVG(currency: string, year: number): Promise<number> {
  const obs = await getAnnualAvgRate(currency, year);
  return obs.rate;
}

// ─── Conversions ─────────────────────────────────────────────────────────────

/**
 * Converts a foreign currency amount to CAD using the Bank of Canada rate.
 * @customfunction TO_CAD
 * @param amount Amount in the foreign currency
 * @param currency Currency code, e.g. "USD"
 * @param [date] Optional date in YYYY-MM-DD format. Omit for today's rate.
 * @returns Amount in CAD
 */
export async function TO_CAD(
  amount: number,
  currency: string,
  date?: string
): Promise<number> {
  const obs = date
    ? await getRateOnDate(currency, date)
    : await getLatestRate(currency);
  return convertToCAD(amount, obs.rate);
}

/**
 * Converts a CAD amount to a foreign currency using the Bank of Canada rate.
 * @customfunction FROM_CAD
 * @param amount Amount in CAD
 * @param currency Currency code, e.g. "USD"
 * @param [date] Optional date in YYYY-MM-DD format. Omit for today's rate.
 * @returns Amount in the foreign currency
 */
export async function FROM_CAD(
  amount: number,
  currency: string,
  date?: string
): Promise<number> {
  const obs = date
    ? await getRateOnDate(currency, date)
    : await getLatestRate(currency);
  return convertFromCAD(amount, obs.rate);
}

/**
 * Returns the date of the most recent available rate for a currency.
 * Useful for confirming which business day's rate is being used.
 * @customfunction RATE_DATE
 * @param currency Currency code, e.g. "USD"
 * @param [date] Optional date in YYYY-MM-DD; returns the actual date of the nearest available rate.
 * @returns Date string in YYYY-MM-DD format
 */
export async function RATE_DATE(currency: string, date?: string): Promise<string> {
  const obs = date
    ? await getRateOnDate(currency, date)
    : await getLatestRate(currency);
  return obs.date;
}

CustomFunctions.associate("RATE", RATE);
CustomFunctions.associate("MONTHLY_AVG", MONTHLY_AVG);
CustomFunctions.associate("ANNUAL_AVG", ANNUAL_AVG);
CustomFunctions.associate("TO_CAD", TO_CAD);
CustomFunctions.associate("FROM_CAD", FROM_CAD);
CustomFunctions.associate("RATE_DATE", RATE_DATE);
