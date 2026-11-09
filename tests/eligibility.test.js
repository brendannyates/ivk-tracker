// node --test tests/   (the eligibility feed; invented codes, MRNs and Sohar ids only)
const test = require("node:test"), assert = require("node:assert/strict");
const C = require("../core.js"), N = require("../notify.js");

const H = C.FEED_COLS;
const id = n => `00000000-0000-4000-8000-${String(n).padStart(12, "0")}`;
/* one feed row by column name */
const row = o => H.map(h => o[h] ?? "");
const q = v => /[",\r\n]/.test(v) ? `"${String(v).replace(/"/g, '""')}"` : String(v);
const csv = rows => "﻿" + [H, ...rows].map(r => r.map(q).join(",")).join("\r\n") + "\r\n";

const FEED = [
  row({ "Pulled At": "2026-11-09T15:02:00Z", LASFIR: "TESANN", MRN: "00990001", Kind: "Verification", "Sohar Status": "complete.ineligible", "Sohar Payer": "Test Health",
    "Sohar ID": id(3), "Sohar Link": `https://app.soharhealth.com/verifications/${id(3)}` }),
  row({ "Pulled At": "2026-11-02T15:00:00Z", LASFIR: "TESANN", MRN: "990001", Kind: "Verification", "Sohar Status": "complete.eligible", "Sohar Payer": "Test Health", Plan: "PPO, Gold",
    "Coverage Start": "2026-01-01", "Coverage End": "12/31/2026", Copay: "30", Coinsurance: "20%", Deductible: "1500", "Deductible Remaining": "250", "Deductible Status": "Approaching",
    "OOP Max": "5000", "OOP Remaining": "4100", "OOP Status": "Remaining", "Benefit Flag": "Deductible approaching", "Sohar ID": id(2),
    "Sohar Link": `https://app.soharhealth.com/verifications/${id(2)}`, "Related Sohar IDs": `${id(1)}; ${id(9)}`, "Projected Patient Responsibility": "530.12",
    "Estimate Note": "99214 estimate + $400 self-pay medication (infusion 3)", Source: "Rotation" }),
  row({ "Pulled At": "2026-10-30T15:00:00Z", LASFIR: "TESANN", MRN: "990001", Kind: "Discovery", "Sohar Status": "complete.eligible", "Sohar ID": id(1) }),
  row({ "Pulled At": "2026-11-08T15:00:00Z", LASFIR: "DOEJON", MRN: "990002", Kind: "Discovery", "Sohar Status": "No coverage found", "Sohar ID": id(4) }),
  row({ "Pulled At": "2026-11-01T15:00:00Z", LASFIR: "DOEJON", MRN: "990002", Kind: "Verification", "Sohar Status": "error.member.id", "Sohar ID": id(5) }),
  row({ "Pulled At": "2026-11-09T15:00:00Z", LASFIR: "ELSPAT", MRN: "880001", Kind: "Verification", "Sohar Status": "complete.eligible", "Sohar ID": id(6) })
];

test("parseCsv: BOM, quotes, doubled quotes, commas and line breaks inside quotes, CRLF", () => {
  assert.deepEqual(C.parseCsv('﻿a,b,c\r\n"x, y","say ""hi""","two\nlines"\r\n'), [["a", "b", "c"], ["x, y", 'say "hi"', "two\nlines"]]);
  assert.deepEqual(C.parseCsv("a,b\n1,\n"), [["a", "b"], ["1", ""]]);
  assert.deepEqual(C.parseCsv(""), []);
});

test("parseFeed reads the feed by header, and refuses anything that isn't it", () => {
  const rows = C.parseFeed(csv(FEED));
  assert.equal(rows.length, 6); assert.equal(rows[1][H.indexOf("Plan")], "PPO, Gold");
  assert.equal(rows[1][H.indexOf("Estimate Note")], "99214 estimate + $400 self-pay medication (infusion 3)");
  /* columns in another order still land in FEED_COLS order */
  const back = H.slice().reverse(), text = [back, ...FEED.slice(0, 1).map(r => r.slice().reverse())].map(r => r.map(q).join(",")).join("\n");
  assert.deepEqual(C.parseFeed(text)[0], FEED[0]);
  /* wrapped by the flow: { body }, a JSON string */
  assert.equal(C.parseFeed(JSON.stringify({ body: csv(FEED) })).length, 6);
  assert.equal(C.parseFeed(JSON.stringify(csv(FEED))).length, 6);
  assert.equal(C.parseFeed("MRN,Deductible Remaining\n990001,250\n"), null);
  assert.equal(C.parseFeed("<html>Sign in</html>"), null);
  assert.equal(C.parseFeed(""), null);
});

test("feedMrn: digits only, leading zeros stripped", () => {
  assert.equal(C.feedMrn("00990001"), "990001"); assert.equal(C.feedMrn(990001), "990001"); assert.equal(C.feedMrn("990001.0"), "990001");
  assert.equal(C.feedMrn(" MRN-0099 0001 "), "990001"); assert.equal(C.feedMrn(""), "");
});

test("pickElig: the newest eligible verification, else the newest pull", () => {
  const L = C.eligRows([H, ...FEED]);
  const tes = L.filter(r => r.mrn === "990001"), doe = L.filter(r => r.mrn === "990002");
  assert.equal(tes[0].id, id(3));                         // newest first
  assert.equal(C.pickElig(tes).id, id(2));                // the eligible verification, not the newer ineligible one or the discovery
  assert.equal(C.pickElig(doe).id, id(4));                // none eligible: the newest of any kind
  assert.equal(C.pickElig([]), null);
  assert.equal(C.isEligibleVer({ kind: "Discovery", status: "complete.eligible" }), false);
});

test("planEligibility keeps this tracker's rows and fills the profile's Sohar columns only", () => {
  const mrns = new Map([["990001", "990001"], ["990002", "990002"]]), codes = new Map([["990001", "TESANN"], ["990002", "DOEJON"]]);
  const PROF = [C.PROFILE_COLS, C.PROFILE_COLS.map(c => ({ MRN: "990001", Patient: "TESANN", ProvCover: "Insurance", ProvPayer: "Test Health", BidfPrinted: "2026-10-01", Note: "staff note" }[c] ?? ""))];
  const plan = C.planEligibility(C.parseFeed(csv(FEED)), { mrns, codes, profile: PROF, now: "2026-11-09T16:00:00.000Z" });
  assert.equal(plan.total, 6); assert.equal(plan.rows.length, 5); assert.equal(plan.patients, 2);
  assert.ok(plan.rows.every(r => r[H.indexOf("MRN")] !== "880001"));
  const tes = plan.items.find(i => i.mrn === "990001"), doe = plan.items.find(i => i.mrn === "990002");
  assert.equal(tes.isNew, false); assert.equal(tes.patch.SoharRef, id(2)); assert.equal(tes.patch.SoharStatus, "complete.eligible");
  assert.equal(tes.patch.CovStart, "2026-01-01"); assert.equal(tes.patch.CovEnd, "2026-12-31"); assert.equal(tes.patch.SoharDate, C.isoOf(C.serialOf(new Date(Date.parse("2026-11-02T15:00:00Z")))));
  assert.equal(tes.patch["Related Sohar IDs"], `${id(1)}; ${id(9)}`); assert.equal(tes.patch["Projected Patient Responsibility"], "530.12");
  assert.equal(tes.patch.By, C.FEED_BY); assert.equal(tes.patch.Updated, "2026-11-09T16:00:00.000Z");
  for (const k of ["ProvCover", "ProvPayer", "InfCover", "InfPayer", "BidfPrinted", "Note", "Patient"]) assert.ok(!(k in tes.patch), k);
  assert.equal(doe.isNew, true); assert.equal(doe.patch.Patient, "DOEJON"); assert.equal(doe.patch.SoharStatus, "No coverage found");
});

test("planEligibility: an unchanged profile is left alone, and so is a newer one entered by hand", () => {
  const mrns = new Map([["990001", "990001"]]), rows = C.parseFeed(csv(FEED));
  const first = C.planEligibility(rows, { mrns, profile: [C.PROFILE_COLS], now: "x" }).items[0];
  const head = [...C.PROFILE_COLS, ...C.PROFILE_FEED_COLS], at = r => head.map(c => r[c] ?? "");
  const same = C.planEligibility(rows, { mrns, profile: [head, at(Object.assign({ MRN: "990001" }, first.patch))] });
  assert.equal(same.items.length, 0); assert.equal(same.same, 1);
  const byHand = C.planEligibility(rows, { mrns, profile: [head, at({ MRN: "990001", SoharDate: "2026-11-20", By: "Sam" })] });
  assert.equal(byHand.items.length, 0); assert.equal(byHand.newer, 1);
});

test("Sohar links only point at Sohar; related ids link as verifications", () => {
  assert.equal(C.soharLink(`https://app.soharhealth.com/discovery/${id(1)}`), `https://app.soharhealth.com/discovery/${id(1)}`);
  assert.equal(C.soharLink("javascript:alert(1)"), ""); assert.equal(C.soharLink("https://example.com/verifications/x"), "");
  assert.equal(C.soharVerLink(id(1)), `https://app.soharhealth.com/verifications/${id(1)}`); assert.equal(C.soharVerLink("a b"), "");
});

test("eligDue: once a day", () => {
  const T = C.serialOf(new Date(2026, 10, 9));
  assert.equal(C.eligDue("", T), true);
  assert.equal(C.eligDue(new Date(2026, 10, 8, 23, 0).toISOString(), T), true);
  assert.equal(C.eligDue(new Date(2026, 10, 9, 7, 5).toISOString(), T), false);
});

test("pullFeed posts {tracker} as text/plain (no preflight) and reads the answer as text", async () => {
  const URL = "https://default0000.00.environment.api.powerplatform.com/powerautomate/automations/direct/workflows/abc/triggers/manual/paths/invoke?sig=FAKE";
  const calls = [], f = (status, body) => async (u, o) => { calls.push({ u, o }); return { ok: status < 300, status, text: async () => body }; };
  const ok = await N.pullFeed(URL, { tracker: "IVK" }, f(200, csv(FEED)));
  assert.equal(ok.ok, true); assert.equal(C.parseFeed(ok.text).length, 6);
  assert.equal(calls[0].o.method, "POST"); assert.equal(calls[0].o.headers["Content-Type"], "text/plain;charset=UTF-8"); assert.equal(calls[0].o.body, '{"tracker":"IVK"}');
  const bad = await N.pullFeed(URL, { tracker: "IVK" }, f(404, "")); assert.equal(bad.ok, false); assert.match(bad.reason, /404.*Nothing changed/);
  const none = await N.pullFeed("https://example.com/x", {}, f(200, "")); assert.equal(none.ok, false); assert.match(none.reason, /Settings › Eligibility feed/);
  const down = await N.pullFeed(URL, {}, async () => { throw new Error("net"); }); assert.equal(down.ok, false); assert.match(down.reason, /couldn’t be reached/);
});
