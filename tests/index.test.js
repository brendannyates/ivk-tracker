// Index planner rules on provider templates (invented codes, MRNs and providers only)
const test = require("node:test"), assert = require("node:assert/strict");
const C = require("../core.js"), TM = require("../template.js"), P = require("./pattern.js");

const T = C.serialOf(new Date(2026, 9, 21));   // Wednesday Oct 21 2026
const WED = T + 7, cfg = Object.assign({}, C.DEFAULTS), types = TM.DEFAULT_TYPES;
const RIV = "Rivera MD, Ana Maria", CHEN = "Chen, MD Paul";
/* Rivera every weekday: a start at 9:00 ending on an IVK post check, a start at 12:30 ending on a flex post check, a follow-up at 10:00.
   Chen on Thursdays: a start at 12:00 with its post check at 1:30. */
const every = new Map(); for (let d = 1; d <= 5; d++) { every.set(`${d}|540`, "IV Ketamine Start/Index"); every.set(`${d}|630`, "IVK Post Check"); every.set(`${d}|600`, "IV Ketamine Follow Up");
  every.set(`${d}|750`, "IV Ketamine Start/Index"); every.set(`${d}|840`, "Flex Post Check"); every.set(`${d}|870`, "Telepsych"); }
const gRiv = TM.layout({ label: "Rivera", copy: every, types }, 1);
const gChen = TM.layout({ label: "Chen", copy: new Map([["4|720", "IV Ketamine Start/Index"], ["4|810", "IVK Post Check"]]), types }, TM.nextRow(TM.parseGrids("t", P.sheet(TM, [gRiv]))));
const GRIDS = TM.parseGrids("IVK Template", P.sheet(TM, [gRiv, gChen]));
const AP_H = ["MRN", "Patient", "Date", "Time", "Provider", "Event", "Status", "Location", "Name", "Lookup Key"];
const ap = (mrn, code, date, min, prov = RIV, event = "IV Ketamine Follow Up", status = "Expected") => [mrn, "", date, min / 1440, prov, event, status, "Castro Street", code, ""];
const X = (appts = [], holds = [], o = {}) => Object.assign({ appts: C.appts([AP_H, ...appts]), holds, others: [], cfg, chairs: 4, needsPost: (m, ev) => /start/i.test(ev) }, TM.onDays(o.grids || GRIDS, types, 90), o);
const plan = (o = {}) => C.planIndex(Object.assign({ start: WED, today: T, perWeek: 2, weeks: 3, minGap: 2, maxGap: 5, avail: null, X: X() }, o));
const gaps = r => r.items.slice(1).map((x, i) => x.date - r.items[i].date);

test("standard index: 6 visits, 2 in each 7 days from treatment 1, never back to back, at most 5 days apart", () => {
  const r = plan();
  assert.equal(r.items.length, 6); assert.deepEqual(r.notes, []);
  const t1 = r.items[0];
  assert.equal(t1.date, WED); assert.equal(t1.time, 540, "a start ending on an IVK post check comes before one ending on a flex post check");
  assert.equal(t1.postUse, "post"); assert.equal(t1.post, 630);
  assert.ok(gaps(r).every(g => g >= 2 && g <= 5), gaps(r).join());
  for (const w of [1, 2, 3]) assert.equal(r.items.filter(x => x.week === w).length, 2, `week ${w}`);
  /* week 1 is Wednesday to Tuesday; week 2 starts the next Wednesday */
  assert.ok(r.items.filter(x => x.week === 1).every(x => x.date - WED <= 6)); assert.ok(r.items.filter(x => x.week === 2).every(x => x.date - WED >= 7 && x.date - WED <= 13));
});

test("spacing is as even as the slots allow: weekends are closed, so after a Wednesday start 3 and 4 day gaps follow", () => {
  const g = gaps(plan());
  assert.ok([2, 5].includes(g[0]) && g.slice(1).every(x => x === 3 || x === 4), `gaps ${g.join()}`);
});

test("a custom protocol: 3 a week for 2 weeks", () => {
  const r = plan({ perWeek: 3, weeks: 2 });
  assert.equal(r.items.length, 6); assert.ok(gaps(r).every(g => g >= 2)); assert.equal(r.items.filter(x => x.week === 1).length, 3);
});

test("availability: only Monday mornings means week gaps over 5 days, each flagged, and the weeks are short", () => {
  const r = plan({ avail: C.availOf({ avail: "Mon AM" }) });
  assert.ok(r.items.every(x => C.weekday(x.date) === 1 && x.time < 720));
  assert.ok(r.items.slice(1).every(x => x.long && x.gap === 7));
  assert.ok(r.notes.some(n => /Week 1 has 1 visit/.test(n)), r.notes.join(" | "));
});

test("a booked start, a full room, a used post check each move treatment 1", () => {
  assert.equal(plan({ X: X([ap("990201", "AAABBB", WED, 540)]) }).items[0].time, 750, "9:00 is booked, 12:30 is next");
  const full = X([ap("990201", "AAABBB", WED, 540), ...["990202", "990203", "990204", "990205"].map(m => ap(m, "CCCDDD", WED, 660, CHEN))]);
  assert.notEqual(plan({ X: full }).items[0].date, WED, "4 chairs busy 11:00 to 1:00, so 12:30 can't start");
  assert.equal(C.daySlots(WED, full).slots.find(s => s.time === 750).full, true);
  /* a tele visit in the 2:00 flex post check leaves the 12:30 start for follow-ups only */
  const tele = X([ap("990201", "AAABBB", WED, 540)], [], { others: [{ date: WED, time: 840, provider: RIV, event: "Telepsych", status: "Expected" }] });
  const s = C.daySlots(WED, tele).slots.find(x => x.time === 750);
  assert.equal(s.postBy.kind, "other"); assert.equal(s.open, true); assert.equal(s.openIndex, false);
  assert.notEqual(plan({ X: tele }).items[0].date, WED);
});

test("post checks: index visits, holds and patients who require one use them; follow-ups don't", () => {
  const fu = C.daySlots(WED, X([ap("990201", "AAABBB", WED, 540)]));
  assert.equal(fu.slots.find(s => s.time === 540).by.code, "AAABBB"); assert.equal(fu.cells.find(c => c.time === 630).by, null, "a follow-up leaves the post check free");
  const st = C.daySlots(WED, X([ap("990201", "AAABBB", WED, 540, RIV, "IV Ketamine Treatment Start")]));
  assert.equal(st.cells.find(c => c.time === 630).by.kind, "post");
  const flagged = C.daySlots(WED, X([ap("990201", "AAABBB", WED, 540)], [], { needsPost: m => m === "990201" }));
  assert.equal(flagged.cells.find(c => c.time === 630).by.code, "AAABBB");
  const held = C.daySlots(WED, X([], [{ id: "h1", mrn: "990301", code: "EEEFFF", date: WED, time: 750, provider: "Rivera", status: "held", at: 0 }]));
  assert.equal(held.slots.find(s => s.time === 750).by.kind, "hold"); assert.equal(held.cells.find(c => c.time === 840).by.kind, "holdpost");
  /* an other visit lands in the cell it starts in, and a NextGen post check hold doesn't count */
  const o = C.daySlots(WED, X([], [], { others: [{ date: WED, time: 885, provider: RIV, event: "Telepsych", status: "Kept" }, { date: WED, time: 630, provider: RIV, event: "IVK Post Check Hold", status: "Expected" }] }));
  assert.equal(o.cells.find(c => c.time === 870).by.event, "Telepsych"); assert.equal(o.cells.find(c => c.time === 630).by, null);
});

test("a patient who requires a post check: open starts that end on a free post check, index or follow-up", () => {
  const L = C.postCheckSlots(X(), WED, 4);
  assert.equal(L.length, 4); assert.ok(L.every(s => s.postUse && s.open && !s.postBy));
  assert.deepEqual(L.slice(0, 2).map(s => s.time), [540, 750]);
});

test("the real template's shape: Wednesday's index starts, IVK post checks first", () => {
  const real = TM.parseGrids("t", P.sheet(TM, [TM.layout({ label: "Rivera", copy: P.copyMap(TM), types }, 1)]));
  const r = plan({ X: X([], [], { grids: real, ...TM.onDays(real, types, 90) }) });
  assert.equal(r.items.length, 6); assert.equal(r.items[0].time, 750, "Wed 12:30 ends on the 2:00 IVK post check");
  assert.ok(r.items.every(x => [3, 4, 5].includes(C.weekday(x.date))), "Rivera only has index starts Wednesday to Friday");
  assert.ok(gaps(r).every(g => g >= 2 && g <= 5), gaps(r).join());
});

test("a visit already booked is kept in the plan", () => {
  const b = plan({ fixed: [{ date: WED + 2, time: 540, provider: RIV }] }).items.find(x => x.date === WED + 2);
  assert.ok(b && b.fixed && b.type === "Booked");
});

test("holds take their slot; a released one doesn't", () => {
  const h = { id: "h1", mrn: "990301", code: "EEEFFF", date: WED, time: 540, provider: "Rivera", status: "held", at: 0 };
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

test("capacity: how many more patients can start index in the next 4 weeks", () => {
  const o = { today: T, perWeek: 2, weeks: 3, minGap: 2, maxGap: 5 };
  const free = C.capacity(Object.assign({ X: X() }, o), 4);
  assert.ok(free.starts >= 4, JSON.stringify(free)); assert.ok(free.open > 0);
  const busy = C.capacity(Object.assign({ X: X([], Array.from({ length: 20 }, (_, i) => ({ id: "h" + i, mrn: "99" + i, code: "X", date: T + 1 + i, time: 540, provider: "Rivera", status: "held" }))) }, o), 4);
  assert.ok(busy.starts < free.starts && busy.open < free.open, JSON.stringify([free, busy]));
});

test("hold state: 10 minutes to confirm, then red; on the list counts as booked", () => {
  const now = Date.UTC(2026, 9, 21, 20, 0), h = { mrn: "990301", date: WED, time: 540, status: "held", at: now - 5 * 60000 };
  const A = C.appts([AP_H, ap("990301", "EEEFFF", WED, 540)]);
  assert.equal(C.holdState(h, [], now, cfg), "held");
  assert.equal(C.holdState(Object.assign({}, h, { at: now - 11 * 60000 }), [], now, cfg), "late");
  assert.equal(C.holdState(Object.assign({}, h, { at: now - 11 * 60000 }), A, now, cfg), "booked");
  assert.equal(C.holdState(Object.assign({}, h, { status: "released" }), A, now, cfg), "released");
  const rows = C.holdRows([C.HOLD_COLS, ["h1", "990301", "EEEFFF", C.isoOf(WED), "9:00 AM", "Rivera", "1", "held", "Sam", new Date(now).toISOString(), "", ""]]);
  assert.equal(rows[0].date, WED); assert.equal(rows[0].time, 540); assert.equal(rows[0].at, now); assert.equal(rows[0].n, 1);
});

test("a hold marked booked stops keeping its time once a later appointment list is in", () => {
  const done = new Date(2026, 9, 21, 14, 0).getTime(), h = { mrn: "990301", date: WED, time: 540, provider: "Rivera", status: "booked", doneAt: done };
  assert.equal(C.holding(h, [], null), true); assert.equal(C.holding(h, [], T + 13 / 24), true);
  assert.equal(C.holding(h, [], T + 15 / 24), false); assert.equal(C.holding(Object.assign({}, h, { status: "held" }), [], T + 15 / 24), true);
});

test("the wider report: IVK visits stay on the Appt List, the rest are other visits with no patient on them", () => {
  assert.ok(C.isIvkVisit("IV Ketamine Follow Up") && C.isIvkVisit("IV Ketamine Treatment Start"));
  assert.ok(!C.isIvkVisit("IV Ketamine Consult") && !C.isIvkVisit("Telepsych") && !C.isIvkVisit("TMS Follow Up"));
  const o = C.otherRows([C.OTHER_COLS, [C.isoOf(WED), "2:00 PM", RIV, "Telepsych", "Expected"]]);
  assert.deepEqual(o[0], { date: WED, time: 840, provider: RIV, event: "Telepsych", status: "Expected" });
  assert.equal(C.flagRows([C.FLAG_COLS, ["990201", "AAABBB", "TRUE", "Sam", ""], ["990202", "CCCDDD", "FALSE", "Sam", ""]]).get("990201"), true);
});

test("availability text: tracker preferences, typed days, any", () => {
  const p = { prefs: [{ am: true, pm: false }, { am: false, pm: false }, { am: false, pm: true }, { am: false, pm: false }, { am: false, pm: false }] };
  assert.equal(C.availText(C.availOf({ avail: "" }, p)), "Mon AM, Wed PM");
  assert.equal(C.availOf({ avail: "any" }, p), null);
  assert.equal(C.availText(C.availOf({ avail: "tue pm; thu am" }, null)), "Tue PM, Thu AM");
});

test("Index Planner tab: layout found from the Date row; cells get codes and holds", () => {
  const V = [["Index Week", "Week 1", "", ""], ["Day", "Mon", "Tues", "Wed"], ["Date", T + 5, T + 6, T + 7], [0.375, "", "", ""], ["9:30 AM", "", "", ""], ["% Scheduled", "", "", ""]];
  const L = C.plannerLayout(V);
  assert.deepEqual(L.rows, [[3, 540], [4, 570]]); assert.deepEqual(L.cols, [[1, T + 5], [2, T + 6], [3, T + 7]]);
  const A = C.appts([AP_H, ap("990401", "GGGHHH", T + 7, 540), ap("990402", "IIIJJJ", T + 7, 555), ap("990403", "KKKLLL", T + 5, 570, RIV, "IV Ketamine Follow Up", "Cancelled")]);
  assert.deepEqual(C.plannerGrid(L, A, [{ code: "MMMNNN", date: T + 6, time: 570, status: "held", mrn: "990404" }]), [["", "", "GGGHHH, IIIJJJ"], ["", "HOLD MMMNNN", ""]]);
});

test("a long custom protocol plans quickly", () => {
  const t0 = Date.now(), r = plan({ perWeek: 3, weeks: 6 });
  assert.equal(r.items.length, 18); assert.ok(gaps(r).every(g => g >= 2)); assert.ok(Date.now() - t0 < 1500, `${Date.now() - t0} ms`);
});

test("treatment 1 moves a day or two later when the first open start would force a gap over 5 days", () => {
  const real = TM.parseGrids("t", P.sheet(TM, [TM.layout({ label: "Rivera", copy: P.copyMap(TM), types }, 1)]));
  /* from Thu Oct 22 the next index start is Wed Oct 28 (6 days); starting Fri Oct 23 keeps every gap at 2 or 5 */
  const r = plan({ start: T + 1, X: X([], [], { grids: real, ...TM.onDays(real, types, 90) }) });
  assert.equal(r.items[0].date, T + 2); assert.ok(gaps(r).every(g => g >= 2 && g <= 5), gaps(r).join()); assert.deepEqual(r.notes, []);
  const pinned = plan({ start: T + 1, pins: { 0: { date: T + 1, time: 540, provider: "Rivera", pkey: "rivera" } }, X: X([], [], { grids: real, ...TM.onDays(real, types, 90) }) });
  assert.equal(pinned.items[0].date, T + 1, "a hand-picked treatment 1 stays"); assert.ok(pinned.items[1].long);
});

test("review: a 9:15 visit takes the 9:00 start, and a booked hold keeps its post check whatever NextGen calls it", () => {
  const late = C.daySlots(WED, X([ap("990201", "AAABBB", WED, 555, RIV, "IV Ketamine Treatment Start")]));
  assert.equal(late.slots.find(s => s.time === 540).by.code, "AAABBB"); assert.equal(late.cells.find(c => c.time === 630).by.kind, "post");
  const h = { id: "h1", mrn: "990301", code: "EEEFFF", date: WED, time: 540, provider: "Rivera", status: "booked", at: 0 };
  const booked = C.daySlots(WED, X([ap("990301", "EEEFFF", WED, 540)], [h], { needsPost: () => false }));
  assert.equal(booked.cells.find(c => c.time === 630).by.code, "EEEFFF", "index visit 3 booked as a follow-up still uses its post check");
});
