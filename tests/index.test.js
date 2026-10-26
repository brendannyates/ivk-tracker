// Index planner rules (invented codes, MRNs and providers only)
const test = require("node:test"), assert = require("node:assert/strict");
const C = require("../core.js");

const T = C.serialOf(new Date(2026, 9, 21));   // Wednesday Oct 21 2026
const WED = T + 7, cfg = Object.assign({}, C.DEFAULTS);
const RIV = "Rivera MD, Ana Maria", CHEN = "Chen, MD Paul";
const SL_H = ["Id", "Provider", "Day", "Time", "Type"];
const slots = C.slotRows([SL_H,
  ...["Mon", "Tue", "Wed", "Thu", "Fri"].map((d, i) => ["s" + i, RIV, d, "9:00 AM", "Index"]),
  ["s5", RIV, "Wednesday", "12:30 PM", "Flex"], ["s6", RIV, "Fri", "12:30 PM", "Flex"], ["s7", CHEN, "Thu", "12:00 PM", "Index"],
  ["s8", "", "Tue", "3:00 PM", "Follow-up"]]);
const AP_H = ["MRN", "Patient", "Date", "Time", "Provider", "Event", "Status", "Location", "Name", "Lookup Key"];
const ap = (mrn, code, date, min, prov = RIV, status = "Expected") => [mrn, "", date, min / 1440, prov, "IV Ketamine Follow Up", status, "Castro Street", code, ""];
const X = (appts = [], holds = []) => ({ slots, appts: C.appts([AP_H, ...appts]), holds, cfg, chairs: 4 });
const plan = (o = {}) => C.planIndex(Object.assign({ start: WED, today: T, perWeek: 2, weeks: 3, minGap: 2, maxGap: 5, avail: null, X: X() }, o));
const gaps = r => r.items.slice(1).map((x, i) => x.date - r.items[i].date);

test("template rows read by header; days and types tolerate spelling", () => {
  assert.equal(slots.length, 9);
  assert.deepEqual(slots.find(s => s.id === "s5"), { id: "s5", provider: RIV, day: 3, time: 750, type: "Flex" });
  assert.equal(C.slotType("follow up"), "Follow-up"); assert.equal(C.slotType(""), "Index"); assert.equal(C.dayNum("thurs"), 4);
});

test("standard index: 6 visits, 2 in each 7 days from treatment 1, never back to back, at most 5 days apart", () => {
  const r = plan();
  assert.equal(r.items.length, 6); assert.deepEqual(r.notes, []);
  const t1 = r.items[0];
  assert.equal(t1.date, WED); assert.equal(t1.time, 540); assert.equal(t1.type, "Index"); assert.equal(t1.post, 630, "post check in the last 30 min of 2 hours");
  assert.ok(gaps(r).every(g => g >= 2 && g <= 5), gaps(r).join());
  for (const w of [1, 2, 3]) assert.equal(r.items.filter(x => x.week === w).length, 2, `week ${w}`);
  /* week 1 is Wednesday to Tuesday: a Wednesday start puts the following Wednesday in week 2 */
  assert.ok(r.items.filter(x => x.week === 1).every(x => x.date - WED <= 6)); assert.ok(r.items.filter(x => x.week === 2).every(x => x.date - WED >= 7 && x.date - WED <= 13));
  assert.ok(r.items.every(x => C.weekday(x.date) >= 1 && C.weekday(x.date) <= 5));
});

test("spacing is as even as the slots allow: weekends are closed, so after a Wednesday start 3 and 4 day gaps follow", () => {
  const g = gaps(plan()), spread = g.reduce((s, x) => s + (x - 3.5) ** 2, 0);
  assert.ok([2, 5].includes(g[0]) && g.slice(1).every(x => x === 3 || x === 4), `gaps ${g.join()}`); assert.ok(spread <= 2.25 + 4 * 0.25 + 1e-9);
});

test("a custom protocol: 3 a week for 2 weeks", () => {
  const r = plan({ perWeek: 3, weeks: 2 });
  assert.equal(r.items.length, 6); assert.ok(gaps(r).every(g => g >= 2)); assert.equal(r.items.filter(x => x.week === 1).length, 3);
});

test("availability: only Monday mornings means week gaps over 5 days, each flagged, and the weeks are short", () => {
  const r = plan({ avail: C.availOf({ avail: "Mon AM" }) });
  assert.ok(r.items.every(x => C.weekday(x.date) === 1));
  assert.ok(r.items.slice(1).every(x => x.long && x.gap === 7));
  assert.ok(r.notes.some(n => /Week 1 has 1 visit/.test(n)), r.notes.join(" | "));
});

test("a taken slot and a full room move treatment 1", () => {
  const taken = plan({ X: X([ap("990201", "AAABBB", WED, 540)]) });
  assert.equal(taken.items[0].date, WED); assert.equal(taken.items[0].time, 750, "Wednesday's 9:00 is taken, the 12:30 flex is next");
  assert.equal(taken.items[0].type, "Flex");
  const full = plan({ X: X([ap("990201", "AAABBB", WED, 540), ...["990202", "990203", "990204", "990205"].map(m => ap(m, "CCCDDD", WED, 660, CHEN))]) });
  assert.notEqual(full.items[0].date, WED, "4 chairs busy from 11:00 to 1:00, so 12:30 can't start");
  const d = C.daySlots(WED, X([ap("990201", "AAABBB", WED, 540), ...["990202", "990203", "990204", "990205"].map(m => ap(m, "CCCDDD", WED, 660, CHEN))]));
  assert.equal(d.slots.find(s => s.time === 540).by.code, "AAABBB"); assert.equal(d.slots.find(s => s.time === 750).full, true);
});

test("a visit already booked is kept in the plan", () => {
  const fixed = [{ date: WED + 2, time: 540, provider: RIV }];
  const r = plan({ fixed });
  const b = r.items.find(x => x.date === WED + 2); assert.ok(b && b.fixed && b.type === "Booked");
});

test("holds take their slot; a released one doesn't", () => {
  const h = { id: "h1", mrn: "990301", code: "EEEFFF", date: WED, time: 540, provider: RIV, status: "held", at: 0 };
  assert.equal(plan({ X: X([], [h]) }).items[0].time, 750);
  assert.equal(plan({ X: X([], [Object.assign({}, h, { status: "released" })]) }).items[0].time, 540);
});

test("pins and alternatives", () => {
  const r = plan(), alt = C.altSlots(r, 1, 2);
  assert.ok(alt.length > 0);
  assert.ok(alt.every(s => s.date - r.items[0].date >= 2 && r.items[2].date - s.date >= 2), "alternatives keep a day off either side");
  const pick = alt[alt.length - 1], r2 = plan({ pins: { 0: r.items[0], 1: pick } });
  assert.equal(r2.items[1].date, pick.date); assert.equal(r2.items[1].time, pick.time); assert.ok(r2.items[1].pinned);
  assert.ok(gaps(r2).every(g => g >= 2));
});

test("hold state: 10 minutes to confirm, then red; on the list counts as booked", () => {
  const now = Date.UTC(2026, 9, 21, 20, 0), h = { mrn: "990301", date: WED, time: 540, status: "held", at: now - 5 * 60000 };
  const A = C.appts([AP_H, ap("990301", "EEEFFF", WED, 540)]);
  assert.equal(C.holdState(h, [], now, cfg), "held");
  assert.equal(C.holdState(Object.assign({}, h, { at: now - 11 * 60000 }), [], now, cfg), "late");
  assert.equal(C.holdState(Object.assign({}, h, { at: now - 11 * 60000 }), A, now, cfg), "booked");
  assert.equal(C.holdState(Object.assign({}, h, { status: "released" }), A, now, cfg), "released");
  const rows = C.holdRows([C.HOLD_COLS, ["h1", "990301", "EEEFFF", C.isoOf(WED), "9:00 AM", RIV, "1", "held", "Sam", new Date(now).toISOString(), "", ""]]);
  assert.equal(rows[0].date, WED); assert.equal(rows[0].time, 540); assert.equal(rows[0].at, now); assert.equal(rows[0].n, 1);
});

test("availability text: tracker preferences, typed days, any", () => {
  const p = { prefs: [{ am: true, pm: false }, { am: false, pm: false }, { am: false, pm: true }, { am: false, pm: false }, { am: false, pm: false }] };
  assert.equal(C.availText(C.availOf({ avail: "" }, p)), "Mon AM, Wed PM");
  assert.equal(C.availOf({ avail: "any" }, p), null);
  assert.equal(C.availText(C.availOf({ avail: "tue pm; thu am" }, null)), "Tue PM, Thu AM");
  assert.equal(C.fits(C.availOf({ avail: "Mon AM" }), 1, 540), true); assert.equal(C.fits(C.availOf({ avail: "Mon AM" }), 1, 750), false);
});

test("Index Planner tab: layout found from the Date row; cells get codes and holds", () => {
  const V = [["Index Week", "Week 1", "", ""], ["Day", "Mon", "Tues", "Wed"], ["Date", T + 5, T + 6, T + 7], [0.375, "", "", ""], ["9:30 AM", "", "", ""], ["% Scheduled", "", "", ""],
    ["", "", "", ""], ["Index Planning", "", "", ""], ["Index Eligible", "Index Slot", "Post Check", "Slot Type"], ["", 0.375, 0.4375, "Flex"], ["", "9:30 AM", "11:00 AM", ""], ["", "", "", ""]];
  const L = C.plannerLayout(V);
  assert.deepEqual(L.rows, [[3, 540], [4, 570]]); assert.deepEqual(L.cols, [[1, T + 5], [2, T + 6], [3, T + 7]]);
  const A = C.appts([AP_H, ap("990401", "GGGHHH", T + 7, 540), ap("990402", "IIIJJJ", T + 7, 555), ap("990403", "KKKLLL", T + 5, 570, RIV, "Cancelled")]);
  const g = C.plannerGrid(L, A, [{ code: "MMMNNN", date: T + 6, time: 570, status: "held", mrn: "990404" }]);
  assert.deepEqual(g, [["", "", "GGGHHH, IIIJJJ"], ["", "HOLD MMMNNN", ""]]);
  assert.deepEqual(C.tabSlots(V), [{ time: 540, post: 630 }, { time: 570, post: 660 }]);
});

test("a hold marked booked stops keeping its time once a later appointment list is in", () => {
  const done = new Date(2026, 9, 21, 14, 0).getTime(), h = { mrn: "990301", date: WED, time: 540, provider: RIV, status: "booked", doneAt: done };
  assert.equal(C.holding(h, [], null), true, "no upload since: still keeps the time");
  assert.equal(C.holding(h, [], T + 13 / 24), true, "list from before the booking");
  assert.equal(C.holding(h, [], T + 15 / 24), false, "list uploaded after the booking: NextGen is the record");
  assert.equal(C.holding(Object.assign({}, h, { status: "held" }), [], T + 15 / 24), true, "never confirmed: keeps it (and stays red)");
  assert.equal(plan({ X: Object.assign(X([], [h]), { listAt: T + 15 / 24 }) }).items[0].time, 540);
});

test("a long custom protocol plans quickly", () => {
  const t0 = Date.now(), r = plan({ perWeek: 3, weeks: 6 });
  assert.equal(r.items.length, 18); assert.ok(gaps(r).every(g => g >= 2)); assert.ok(Date.now() - t0 < 1500, `${Date.now() - t0} ms`);
});
