#!/usr/bin/env node
// Writes manifest.xml from manifest.template.xml using the hosting URL you pass in.
// Usage: node scripts/build-manifest.js https://yourname.github.io/deal-team-scheduler

const fs = require("fs");
const path = require("path");

const raw = process.argv[2];
if (!raw) {
  console.error("Missing URL. Example:\n  npm run manifest -- https://yourname.github.io/deal-team-scheduler");
  process.exit(1);
}

let url;
try {
  url = new URL(raw);
} catch (e) {
  console.error(`"${raw}" is not a valid URL.`);
  process.exit(1);
}
if (url.protocol !== "https:") {
  console.error("Outlook only loads add-ins over HTTPS. Use a URL that starts with https://");
  process.exit(1);
}

const baseUrl = url.toString().replace(/\/+$/, "");
const root = path.join(__dirname, "..");
const template = fs.readFileSync(path.join(root, "manifest.template.xml"), "utf8");
const output = template.split("{{BASE_URL}}").join(baseUrl);

fs.writeFileSync(path.join(root, "manifest.xml"), output);
console.log(`Wrote manifest.xml pointing to ${baseUrl}`);
console.log("Next: confirm " + baseUrl + "/taskpane.html opens in your browser, then upload manifest.xml to Outlook.");
