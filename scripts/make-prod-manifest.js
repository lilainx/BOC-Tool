#!/usr/bin/env node
/**
 * Generates a production manifest from the dev manifest by swapping the
 * localhost dev URL for your real hosting URL.
 *
 * Usage:
 *   node scripts/make-prod-manifest.js https://boc-rates.example.com
 *
 * Output: manifest.prod.xml (upload THIS one for production / org distribution)
 */
const fs = require("fs");
const path = require("path");

const DEV_URL = "https://localhost:3000";
const prodUrl = process.argv[2];

if (!prodUrl) {
  console.error("Error: provide your production base URL.");
  console.error("  Example: node scripts/make-prod-manifest.js https://boc-rates.example.com");
  process.exit(1);
}
if (!/^https:\/\//.test(prodUrl)) {
  console.error("Error: the production URL must use https://");
  process.exit(1);
}

// Strip any trailing slash so we don't produce https://host//taskpane.html
const cleanUrl = prodUrl.replace(/\/+$/, "");

const root = path.resolve(__dirname, "..");
const src = path.join(root, "manifest.xml");
const out = path.join(root, "manifest.prod.xml");

let xml = fs.readFileSync(src, "utf8");
const count = (xml.match(new RegExp(DEV_URL, "g")) || []).length;
xml = xml.split(DEV_URL).join(cleanUrl);

fs.writeFileSync(out, xml);
console.log(`✓ Wrote manifest.prod.xml`);
console.log(`  Replaced ${count} occurrences of ${DEV_URL} → ${cleanUrl}`);
console.log(`  Next: validate it with  npx office-addin-manifest validate manifest.prod.xml`);
