// node --test tests/   (invented codes and MRNs only): the index-booking flag on a treatment start, transfers, changed payments
const test = require("node:test"), assert = require("node:assert/strict");
const C = require("../core.js");

const T = C.serialOf(new Date(2026, 10, 4));   // Wednesday Nov 4 2026
const cfg = Object.assign({}, C.DEFAULTS, { Clinic: "Castro" });   // a workbook that saved its clinic (nothing is built in)
const AP_H = ["MRN", "Patient", "Date", "Time", "Provider", "Event", "Status", "Location"];
const ap = (mrn, date, min, event = "IV Ketamine Follow Up", status = "Expected") => [mrn, "", date, min / 1440, "Rivera MD, Ana Maria", event, status, "Castro Street"];
const P = (o) => Object.assign({ mrn: "990101", code: "TESANN", status: "Induction", freq: "2x per week", loc: "Castro", copay: "40", infusion: "300", prefs: [] }, o);
const START = "IV Ketamine Treatment Start";

test("a start with fewer than 5 visits booked behind it is flagged, only for Induction on the index protocol", () => {
  const A = C.appts([AP_H, ap("990101", T, 600, START), ap("990101", T + 2, 600), ap("990101", T + 7, 600), ap("990101", T + 9, 600, undefined, "Cancelled")]);
  const s = C.indexShort(P(), A, A[0], cfg);
  assert.deepEqual([s.booked, s.need], [2, 5]);
  assert.equal(C.indexShort(P({ freq: "" }), A, A[0], cfg).booked, 2);          // no protocol set yet: the index default
  assert.equal(C.indexShort(P({ freq: "1x per week" }), A, A[0], cfg), null);    // protocol changed
  assert.equal(C.indexShort(P({ status: "Ongoing" }), A, A[0], cfg), null);
  assert.equal(C.indexShort(P(), A, A[1], cfg), null);                           // a follow-up
  const full = C.appts([AP_H, ap("990101", T, 600, START), ...[2, 7, 9, 14, 16].map(d => ap("990101", T + d, 600))]);
  assert.equal(C.indexShort(P(), full, full[0], cfg), null);
  const n = C.needs(P(), { today: T, cfg, appts: A, clinic: "Castro" });
  assert.ok(n.find(y => y.k === "idxbook" && /2 of 5/.test(y.t)));
});

test("a scheduled transfer: last visit here, visits after the date, when to confirm", () => {
  const A = C.appts([AP_H, ap("990101", T, 600), ap("990101", T + 7, 600), ap("990101", T + 21, 600)]);
  const XF = C.xferRows([C.XFER_COLS, ["990101", "TESANN", "Oakland", C.isoOf(T + 14), "scheduled", "", "", "BY", ""]]);
  const xs = C.xferState(XF.get("990101"), A, T);
  assert.equal(xs.last.date, T + 7); assert.equal(xs.after.length, 1); assert.equal(xs.due, false);
  assert.equal(C.xferState(XF.get("990101"), A, T + 7).due, true);               // the last visit here is today
  const n = C.needs(P({ status: "Ongoing", freq: "1x per week" }), { today: T, cfg, appts: A, clinic: "Castro", XF });
  assert.ok(n.find(y => y.k === "xferafter")); assert.ok(n.find(y => y.k === "xferq"));
  assert.equal(C.xferState(Object.assign({}, XF.get("990101"), { status: "done" }), A, T), null);
  assert.equal(C.elsewhere({ loc: "Oakland" }, "Castro"), true); assert.equal(C.elsewhere({ loc: "Castro" }, "Castro"), false);
});

test("a payment changed after check-out: the latest answer counts, collected isn't counted twice", () => {
  const k = C.partKey(C.visitKey(T, "990101", 600), "prov");
  const row = (id, kind, amount) => ({ id, date: T, mrn: "990101", key: k, kind, amount });
  let L = C.ledger("990101", [row("a", "missed", 40)]); assert.equal(L.balance, 40);
  L = C.ledger("990101", [row("a", "missed", 40), row("b", "collected", 40)]); assert.equal(L.balance, 0); assert.equal(L.collected, 40);
  L = C.ledger("990101", [row("a", "collected", 40), row("b", "collected", 25)]); assert.equal(L.collected, 25);
});

test("the tracker's AM/PM columns are read and written by position, so a mistyped header still counts", () => {
  const H = ["MRN", "Patient Name", "Status", "AM", "PM", "AM2", "PM3", "AM4", "PM5", "AM6", "PM7", "AM8", "PM2", "Clinic Notes"];
  const [p] = C.patients([H, [990101, "Test, Ann", "Ongoing", 1, 0, 0, 0, 0, 0, 0, 0, 0, 1, ""]]);
  assert.equal(p.prefs[0].am, true); assert.equal(p.prefs[4].pm, true);          // "PM2" in the Friday PM place
  const g = C.DAYS5.map(() => ({ am: "", pm: "" })); g[4].pm = "P";
  const cells = C.prefCells(g, H);
  assert.equal(cells["@12"], 1); assert.equal(cells["@3"], 0); assert.equal(Object.keys(cells).length, 10);
});

test("a wider report: other clinics' visits kept apart and used to check the hand-off", () => {
  assert.equal(C.isOurs("Castro Street", cfg), true); assert.equal(C.isOurs("", cfg), true); assert.equal(C.isOurs("MHS Oakland", cfg), false);
  const c2 = Object.assign({}, cfg, { Locations: "SF Main; Castro IVK" });
  assert.equal(C.isOurs("SF Main", c2), true); assert.equal(C.isOurs("Castro Street", c2), false);
  const A = C.appts([AP_H, ap("990101", T, 600), ap("990101", T + 7, 600)]);
  const XF = C.xferRows([C.XFER_COLS, ["990101", "TESANN", "Oakland", C.isoOf(T + 14), "scheduled", "", "", "BY", ""]]);
  const E0 = C.elseRows([C.ELSE_COLS]);
  const E = C.elseRows([C.ELSE_COLS, ["990101", "TESANN", C.isoOf(T + 15), "10:00 AM", "Chen MD, Paul", "IV Ketamine Follow Up", "Expected", "MHS Oakland"],
    ["990101", "TESANN", C.isoOf(T + 16), "10:00 AM", "Chen MD, Paul", "IV Ketamine Follow Up", "Cancelled", "MHS Oakland"]]);
  assert.equal(C.xferState(XF.get("990101"), A, T).there, undefined);           // no wider report: nothing said
  assert.equal(C.xferState(XF.get("990101"), A, T, E0).there, null);
  assert.equal(C.xferState(XF.get("990101"), A, T, E).there.date, T + 15);
  const x = { today: T, cfg, appts: A, clinic: "Castro", XF }, pt = P({ status: "Ongoing", freq: "1x per week" });
  assert.ok(C.needs(pt, Object.assign({ E: E0 }, x)).find(y => y.k === "xfernone"));
  assert.ok(C.needs(pt, Object.assign({ E }, x)).find(y => y.k === "xferthere"));
  assert.ok(!C.needs(pt, x).find(y => /^xfer(none|there)$/.test(y.k)));
});

test("NextGen location names: home clinic and appointment counts here and at other clinics", () => {
  const c2 = Object.assign({}, cfg, { Locations: "MHS SF Castro IVK" });
  assert.equal(C.elsewhere({ loc: "MHS SF Castro IVK" }, c2), false); assert.equal(C.elsewhere({ loc: "MHS Oakland IVK" }, c2), true);
  assert.equal(C.elsewhere({ loc: "MHS Castro" }, cfg), false);                 // no Locations set: names the clinic
  const A = C.appts([AP_H, ap("990101", T - 7, 600), ap("990101", T - 3, 600, undefined, "No Show"), ap("990101", T + 2, 600), ap("990101", T + 4, 600, undefined, "Cancelled")]);
  const E = C.elseRows([C.ELSE_COLS, ["990101", "TESANN", C.isoOf(T + 9), "10:00 AM", "Chen MD, Paul", "IV Ketamine Follow Up", "Expected", "MHS Oakland IVK"]]);
  assert.deepEqual(C.visitCounts("990101", A, E, T), { kept: 1, booked: 1, away: [{ loc: "MHS Oakland IVK", n: 1 }] });
});

test("the sooner waitlist: openings before their visit, on their days, new ones only once", () => {
  const SO = C.soonerRows([C.SOONER_COLS, ["990101", "TESANN", C.isoOf(T + 14), "10:00 AM", "Rivera MD, Ana Maria", C.isoOf(T + 2), "1am,3pm", "TRUE", "waiting", "", "", "BY", ""]]);
  const r = SO.get("990101"); assert.equal(r.days.size, 2); assert.equal(r.sameProv, true);
  const mon = T + 5, wed = T + 7;                                                 // T is a Wednesday
  const cands = [{ date: mon, time: 600, provider: "Rivera MD, Ana Maria" }, { date: mon, time: 800, provider: "Rivera MD, Ana Maria" },   // Mon PM: not a day they said
    { date: wed, time: 780, provider: "Chen MD, Paul" }, { date: wed, time: 780, provider: "Rivera MD, Ana Maria" }, { date: T + 1, time: 600, provider: "Rivera MD, Ana Maria" },   // before From
    { date: T + 19, time: 600, provider: "Rivera MD, Ana Maria" }];             // after their visit
  assert.deepEqual(C.soonerMatch(r, cands, T, false).map(c => [c.date, c.time]), [[mon, 600], [wed, 780]]);
  assert.equal(C.soonerMatch(Object.assign({}, r, { status: "done" }), cands, T, false).length, 0);
  /* a cancellation on the new list frees its time, unless someone else is booked in it */
  const old = C.appts([AP_H, ap("990202", mon, 600), ap("990303", wed, 600), ap("990404", T + 8, 600)]);
  const now = C.appts([AP_H, ap("990202", mon, 600, undefined, "Cancelled"), ap("990505", wed, 600), ap("990404", T + 8, 600)]);
  assert.deepEqual(C.freedTimes(old, now, T).map(c => c.date), [mon]);
});

test("older NextGen exports: written-out dates, the time inside the date-time, older header names", () => {
  assert.equal(C.toSerial("Wednesday, November 04, 2026"), T); assert.equal(C.toSerial("Nov 4, 2026"), T);
  assert.equal(C.toMin("11/04/2026 10:30 AM"), 630); assert.equal(C.toSerial("11/04/2026 10:30 AM"), T);
  const aoa = [["Appointments Listing"], [], ["Med Rec Nbr", "Patient Name", "Appt Date", "Provider Name", "Event Name", "Appt Status", "Location Name"],
    ["990101", "Test, Ann", "11/04/2026 10:30 AM", "Rivera MD, Ana Maria", "IV Ketamine Follow Up", "Confirmed", "Castro Street"], ["Total: 1"]];
  const rows = C.parseReport(aoa);
  assert.equal(rows.length, 1); assert.equal(rows[0][2], T); assert.equal(Math.round(rows[0][3] * 1440), 630); assert.equal(rows[0][7], "Castro Street");
});
