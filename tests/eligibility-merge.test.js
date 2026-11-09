// The Eligibility feed flow's { csv, rows } answer: Sohar webhook rows merged into the rotation CSV by Sohar ID. FICTIONAL data.
const test = require("node:test"), assert = require("node:assert/strict");
const C = require("../core.js");
const H = C.FEED_COLS, q = v => /[",\r\n]/.test(v) ? `"${String(v).replace(/"/g, '""')}"` : String(v);
const row = o => H.map(h => q(o[h] ?? "")).join(",");
const csv = "﻿" + [H.join(","), row({ "Pulled At": "2026-11-09T10:00:00Z", MRN: "900", Kind: "Verification", "Sohar Status": "complete.eligible", "Sohar ID": "v-1", "Projected Patient Responsibility": "$20.00" })].join("\r\n");
const rows = [
  { "Pulled At": "2026-11-09T10:00:05Z", MRN: "900", Kind: "Verification", "Sohar Status": "complete.eligible", "Sohar ID": "v-1", Source: "Sohar webhook" },
  { "Pulled At": "2026-11-09T16:00:00Z", MRN: "901", Kind: "Verification", "Sohar Status": "complete.eligible", "Sohar ID": "v-2", Copay: "15", "@odata.etag": "x" },
  { "Pulled At": "2026-11-09T17:00:00Z", MRN: "901", Kind: "Verification", "Sohar Status": "complete.eligible", "Sohar ID": "v-2", Copay: "25" },
  { "Pulled At": "", MRN: "", "Sohar ID": "" }];
test("webhook rows merge into the feed by Sohar ID; the CSV's copy wins, the newest webhook copy wins", () => {
  const out = C.parseFeed(JSON.stringify({ csv, rows })), id = H.indexOf("Sohar ID");
  assert.equal(out.length, 2);
  assert.equal(out.find(r => r[id] === "v-1")[H.indexOf("Projected Patient Responsibility")], "$20.00");
  assert.equal(out.find(r => r[id] === "v-2")[H.indexOf("Copay")], "25");
  assert.equal(C.parseFeed(JSON.stringify({ csv: "", rows })).length, 2);
});
