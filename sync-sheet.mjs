import { mkdir, writeFile } from "node:fs/promises";

const SHEET_ID = "1pFzKzVdvlfXfXQyZybL6aYHBR2bTCvbYPgK2uQYBFQg";
const GID = "0";
const CALLBACK_NAME = "snapshotCallback";
const url =
  `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq` +
  `?tqx=out:json;responseHandler:${CALLBACK_NAME}&gid=${GID}&headers=1&tq=${encodeURIComponent("select *")}`;

const response = await fetch(url);

if (!response.ok) {
  throw new Error(`Unable to download the Google Sheet (${response.status} ${response.statusText}).`);
}

const body = await response.text();
const start = body.indexOf("(");
const end = body.lastIndexOf(");");

if (start === -1 || end === -1) {
  throw new Error("Unexpected Google Visualization response format.");
}

const payload = JSON.parse(body.slice(start + 1, end));

if (payload.status !== "ok" || !payload.table) {
  throw new Error("Google Sheet response did not include a usable table.");
}

const snapshot = {
  generatedAt: new Date().toISOString(),
  source: {
    sheetId: SHEET_ID,
    gid: GID,
    title: "LRMS Inventory System",
  },
  table: payload.table,
};

await mkdir("data", { recursive: true });
await writeFile(
  "data/snapshot.js",
  `window.DASHBOARD_SNAPSHOT = ${JSON.stringify(snapshot, null, 2)};\n`,
  "utf8"
);

console.log(
  `Saved snapshot with ${payload.table.rows.length.toLocaleString()} rows to data/snapshot.js`
);
