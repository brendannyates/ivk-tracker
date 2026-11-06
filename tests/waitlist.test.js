// node --test tests/   (invented codes and MRNs only): the cancellation waitlist (sooner, or a better day and time) and offers
const test = require("node:test"), assert = require("node:assert/strict");
const C = require("../core.js");

const T = C.serialOf(new Date(2026, 10, 4));   // Wednesday Nov 4 2026
const AP_H = ["MRN", "Patient", "Date", "Time", "Provider", "Event", "Status", "Location"];
const RIV = "Rivera MD, Ana Maria", CHEN = "Chen MD, Paul";
const ap = (mrn, date, min, status = "Expected", prov = RIV) => [mrn, "", date, min / 1440, prov, "IV Ketamine Follow Up", status, "Castro Street"];
const row = o => C.SOONER_COLS.map(c => o[c] ?? "");
const one = o => C.soonerRows([C.SOONER_COLS, row(Object.assign({ MRN: "990101", Patient: "TESANN", Status: "waiting" }, o))]).get("990101");
const at = (d, t, p = RIV) => ({ date: d, time: t, provider: p });
const hits = (r, cands, busy, start) => C.soonerMatch(r, cands, T, !!start, busy).map(c => [c.date - T, c.time]);
const thu = T + 1, mon = T + 5, tue = T + 6, thu2 = T + 8;   // T is a Wednesday

test("the Days cell: weekdays, particular dates, a time window and a last day, read and written back the same", () => {
  const w = C.parseWants(`3pm,1am,${C.isoOf(thu2)},${C.isoOf(tue)}pm,${C.isoOf(mon)}@840,t600-900,to${C.isoOf(T + 20)},junk`);
  assert.deepEqual([...w.days].sort(), ["1am", "3pm"]);
  assert.deepEqual(w.dates.map(x => [x.d - T, x.h, x.t]), [[5, "", 840], [6, "pm", null], [8, "", null]]);
  assert.deepEqual(w.tw, { from: 600, to: 900 }); assert.equal(w.to, T + 20);
  assert.equal(C.wantsCell(w), `1am,3pm,${C.isoOf(mon)}@840,${C.isoOf(tue)}pm,${C.isoOf(thu2)},t600-900,to${C.isoOf(T + 20)}`);
  assert.equal(C.parseWants("t900-600").tw, null);   // backwards window ignored
});

test("old rows (sooner only) still read and match as before", () => {
  const r = one({ Booked: C.isoOf(T + 14), BookedTime: "10:00 AM", Provider: RIV, From: C.isoOf(T + 2), Days: "1am,3pm", SameProv: "TRUE" });
  assert.equal(C.soonerKind(r), "sooner"); assert.equal(C.soonerLast(r), T + 13);
  assert.deepEqual(hits(r, [at(mon, 600), at(mon, 800), at(T + 7, 780, CHEN), at(T + 7, 780), at(thu, 600), at(T + 19, 600)]), [[5, 600], [7, 780]]);
});

test("a better time, not just sooner: a latest date after their visit, never the visit they have or a day they already come", () => {
  const r = one({ Booked: C.isoOf(tue), BookedTime: "8:00 AM", Provider: RIV, Days: "4pm", Note: "" });
  assert.equal(C.soonerKind(r), "sooner");
  assert.deepEqual(hits(r, [at(thu, 840), at(thu2, 840)]), [[1, 840]]);           // sooner only: Thu Nov 12 is after Tue
  const r2 = Object.assign({}, r, { to: T + 10 });
  assert.equal(C.soonerKind(r2), "better");
  assert.deepEqual(hits(r2, [at(thu, 840), at(thu2, 840), at(thu2, 600)]), [[1, 840], [8, 840]]);
  assert.deepEqual(hits(r2, [at(thu, 840), at(thu2, 840)], new Set([thu])), [[8, 840]]);   // already in on Thu the 5th
  const r3 = Object.assign({}, r, { to: T + 10, days: new Set(), dates: [] });
  assert.deepEqual(hits(r3, [at(tue, 480), at(tue, 720)]), [[6, 720]]);             // the visit they have isn't an opening
});

test("particular dates: any time, a half, or one time, and they stretch the window past the visit", () => {
  const r = one({ Booked: C.isoOf(mon), BookedTime: "9:00 AM", Provider: RIV, Days: `${C.isoOf(thu2)}@840,${C.isoOf(thu)}am` });
  assert.equal(C.soonerLast(r), thu2); assert.equal(C.soonerKind(r), "better");
  assert.deepEqual(hits(r, [at(thu, 600), at(thu, 780), at(thu2, 840), at(thu2, 870), at(T + 7, 600)]), [[1, 600], [8, 840]]);
});

test("a start-time window applies to weekdays and dates, not to an exact time they named", () => {
  const r = one({ Days: `1am,1pm,${C.isoOf(thu2)}@480,t600-780` });
  assert.equal(r.booked, null); assert.equal(C.soonerKind(r), "better"); assert.equal(C.soonerLast(r), thu2);
  assert.deepEqual(hits(r, [at(mon, 540), at(mon, 600), at(mon, 840), at(thu2, 480), at(thu2, 600)]), [[5, 600], [8, 480]]);
});

test("not tied to a visit, no dates: no end; same provider only counts when tied to a visit", () => {
  const r = one({ Days: "1am", SameProv: "TRUE", Provider: RIV });
  assert.equal(r.sameProv, false); assert.equal(C.soonerLast(r), null);
  assert.deepEqual(hits(r, [at(mon, 600, CHEN), at(T + 40, 600)]), [[5, 600], [40, 600]]);
});

test("a start only takes an index start slot", () => {
  const r = one({ Booked: C.isoOf(T + 14), BookedTime: "10:00 AM" });
  assert.deepEqual(hits(r, [Object.assign(at(mon, 600), { index: false }), at(tue, 600)], null, true), [[6, 600]]);
});

test("offers: rows, keys, and where each one is against the list", () => {
  const H = C.OFFER_COLS, o = (outcome, extra) => Object.assign({ Key: "", MRN: "990101", Patient: "TESANN", Date: C.isoOf(thu), Time: "2:00 PM", Provider: RIV, Source: "freed",
    Booked: C.isoOf(tue), BookedTime: "8:00 AM", Outcome: outcome, Note: "", By: "Sam", Updated: "2026-11-04T10:00:00" }, extra);
  const rows = (...L) => C.offerRows([H, ...L.map(x => H.map(c => x[c] ?? ""))]);
  const [x] = rows(o(""));
  assert.equal(x.key, `990101|${C.isoOf(thu)}|840|${C.soonerKey(x).split("|")[2]}`); assert.equal(x.key, C.offerKey("990101", at(thu, 840)));
  const base = C.appts([AP_H, ap("990101", tue, 480)]);
  assert.equal(C.offerState(x, base, T, []), "offer");
  assert.equal(C.offerState(rows(o("noanswer"))[0], base, T, []), "retry");
  assert.equal(C.offerState(rows(o("declined"))[0], base, T, []), "declined");
  assert.equal(C.offerState(x, C.appts([AP_H, ap("990101", tue, 480), ap("990707", thu, 840)]), T, []), "taken");                // someone booked it
  assert.equal(C.offerState(x, C.appts([AP_H, ap("990101", tue, 480), ap("990707", thu, 840, "Expected", CHEN)]), T, []), "offer");   // another provider's time
  assert.equal(C.offerState(x, base, thu, []), "past");
  const other = rows(o("accepted", { MRN: "990808" }))[0];
  assert.equal(C.offerState(x, base, T, [x, other]), "taken");                                                                     // another patient is taking it
  const acc = rows(o("accepted"))[0], mov = rows(o("moved"))[0];
  assert.equal(C.offerState(acc, base, T, []), "book");
  assert.equal(C.offerState(mov, base, T, []), "moved");
  const listAt = C.serialOf(new Date(2026, 10, 4)) + 0.5;                                                                         // uploaded at noon, after 10 AM
  assert.equal(C.offerState(mov, base, T, [], listAt), "missing");
  const both = C.appts([AP_H, ap("990101", tue, 480), ap("990101", thu, 840)]);
  assert.equal(C.offerState(mov, both, T, [], listAt), "cancelold");
  const moved = C.appts([AP_H, ap("990101", tue, 480, "Cancelled"), ap("990101", thu, 840)]);
  assert.equal(C.offerState(mov, moved, T, [], listAt), "done");
  assert.equal(C.offerState(x, moved, T, []), "done");                                                                             // they booked it some other way
});
