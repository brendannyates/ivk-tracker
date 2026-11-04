// Access demand, index-eligible starts, profiles, per-part payments, template mismatches (invented codes, MRNs and providers only)
const test = require("node:test"), assert = require("node:assert/strict");
const C = require("../core.js"), TM = require("../template.js"), P = require("./pattern.js");

const T = C.serialOf(new Date(2026, 9, 21));   // Wednesday Oct 21 2026
const cfg = Object.assign({}, C.DEFAULTS), types = TM.DEFAULT_TYPES;
const RIV = "Rivera MD, Ana Maria";
const every = new Map(); for (let d = 1; d <= 5; d++) { every.set(`${d}|540`, "IV Ketamine Start/Index"); every.set(`${d}|630`, "IVK Post Check"); every.set(`${d}|600`, "IV Ketamine Follow Up"); }
const GRIDS = TM.parseGrids("IVK Template", P.sheet(TM, [TM.layout({ label: "Rivera", copy: every, types }, 1)]));
const AP_H = ["MRN", "Patient", "Date", "Time", "Provider", "Event", "Status", "Location", "Name", "Lookup Key"];
const ap = (mrn, code, date, min, event = "IV Ketamine Follow Up", prov = RIV) => [mrn, "", date, min / 1440, prov, event, "Expected", "Castro Street", code, ""];
const keys = new Set(GRIDS.map(g => g.key));
const X = (appts = [], holds = []) => Object.assign({ appts: C.appts([AP_H, ...appts]), holds, others: [], cfg, chairs: 4, needsPost: (m, ev) => /start/i.test(ev), hasGrid: k => keys.has(k) }, TM.onDays(GRIDS, types, 90));
const MPT_H = ["MRN", "Patient Name", "Status", "Frequency", "Copay", "Infusion", "Primary Encounter Insurance", "Infusion Insurance", "Tx Start Date", "Protocol Updated"];
const pts = rows => C.patients([MPT_H, ...rows]);

test("an index-eligible start has all 6 index visits open in their windows after it", () => {
  const E = C.eligibleStarts({ today: T, perWeek: 2, weeks: 3, minGap: 2, maxGap: 5, X: X() }, T + 1, 3, 14);
  assert.equal(E.length, 3);
  for (const e of E) { assert.equal(e.plan.items.length, 6); assert.ok(C.fullPlan(e.plan)); assert.equal(e.plan.items[0].date, e.slot.date); }
});

test("a start whose later weeks are booked full isn't eligible, even though it's open", () => {
  /* every 9:00 start booked from the Monday after next for two weeks: a start this week can't fit weeks 2 and 3 */
  const busy = []; for (let d = T + 5; d < T + 19; d++) { const wd = C.weekday(d); if (wd >= 1 && wd <= 5) busy.push(ap("990900" + d, "BUSPAT", d, 540)); }
  const x = X(busy), o = { today: T, perWeek: 2, weeks: 3, minGap: 2, maxGap: 5, X: x };
  const thu = C.daySlots(T + 1, x).slots.find(s => s.openIndex);
  assert.ok(thu, "Thursday's start is open");
  const E = C.eligibleStarts(o, T + 1, 1, 2);
  assert.equal(E.length, 0, "no full index fits from Thursday or Friday");
});

test("a start booked in a follow-up-only slot, or off the template, is flagged; a follow-up in a start slot isn't", () => {
  const x = X([ap("990101", "ONEPAT", T + 1, 600, "IV Ketamine Start"), ap("990102", "TWOPAT", T + 1, 540), ap("990103", "THRPAT", T + 1, 780)]);
  const B = C.daySlots(T + 1, x).bookings, by = m => B.find(b => b.mrn === m);
  assert.equal(by("990101").miss, true, "start in a follow-up slot");
  assert.equal(by("990102").miss, false, "follow-up in a start/index slot");
  assert.equal(by("990103").miss, true, "1:00 PM has no IVK slot on the template");
  assert.equal(by("990101").postAt, 690, "a start's post check lands 90 min after it");
});

test("profiles: coverage from the tracker's columns; only new patients (on the list, no tracker row) are prompted, starts first", () => {
  const P1 = pts([["990201", "ALPPAT", "Ongoing", "1x per week", "40", "300", "Blue Shield", "Self Pay", "", ""]]);
  const cv = C.coverage(P1[0], null);
  assert.equal(cv.mode, "provins"); assert.equal(cv.prov.payer, "Blue Shield"); assert.equal(cv.inf.cover, "selfpay"); assert.equal(cv.prov.due, 40); assert.equal(cv.inf.due, 300);
  const A = C.appts([AP_H, ap("990201", "ALPPAT", T + 2, 540), ap("990299", "NEWPAT", T + 5, 540, "IV Ketamine Start"), ap("990297", "OTHPAT", T + 3, 540), ap("990298", "LATPAT", T + 40, 540)]);
  const G = C.profileGaps(A, new Map(P1.map(p => [p.mrn, p])), new Map(), T, 28);
  assert.deepEqual(G.map(g => g.code), ["NEWPAT", "OTHPAT"], "the start first; the tracker patient isn't prompted; nobody past 4 weeks");
  assert.ok(G.every(g => !g.onMpt));
  const PF = C.profileRows([C.PROFILE_COLS, ["990201", "ALPPAT", "Insurance", "Blue Shield", "Self pay", "", "Blue Shield", "PPO", "In", "", "", "40", "20%", "1000", "250", "5000", "4100", "complete.eligible", "2026-10-20", "r1", "", "", "Sam", ""]]);
  assert.equal(PF.get("990201").sohar.dedRem, "250"); assert.equal(PF.get("990201").sohar.date, C.fromIso("2026-10-20"));
  assert.deepEqual(C.modeParts("infins"), { prov: "selfpay", inf: "insurance" }); assert.equal(C.modeOf("insurance", "insurance"), "both");
});

test("the Sohar benefits CSV: latest verification per MRN, discovery rows skipped", () => {
  const H = ["Per Nbr", "MRN", "LASFIR", "Sohar Payer", "Plan", "Network", "Coverage Level", "Coverage Start", "Coverage End", "Copay", "Coinsurance", "Deductible", "Deductible Remaining",
    "Deductible Status", "OOP Max", "OOP Remaining", "OOP Status", "Benefit Flag", "Sohar Status", "Sohar Date", "Sohar Ref", "Sent As"];
  const row = (mrn, date, rem, sent = "Verification") => ["1", mrn, "ALPPAT", "Blue Shield", "PPO", "In", "", "2026-01-01", "2026-12-31", "40", "20%", "1000", rem, "", "5000", "4000", "", "", "complete.eligible", date, "r", sent];
  const M = C.parseSohar([["Benefits export"], H, row("990201", "2026-10-01", "500"), row("990201", "2026-10-15", "250"), row("990202", "2026-10-16", "0", "Discovery")]);
  assert.equal(M.size, 1); assert.equal(M.get("990201").dedRem, "250");
  assert.equal(C.parseSohar([["MRN", "Name"], ["1", "x"]]), null);
});

test("payments by part: the provider visit and the infusion are answered on their own; old whole-visit rows still count", () => {
  const p = pts([["990201", "ALPPAT", "Ongoing", "1x per week", "40", "300", "Blue Shield", "Self Pay", "", ""]])[0];
  assert.deepEqual(C.dueParts(p), [["prov", 40], ["inf", 300]]);
  const k = C.visitKey(T - 2, "990201", 540), pay = (key, kind, n) => ({ id: key + kind, key, kind, amount: n, mrn: "990201", date: T - 2 });
  assert.equal(C.visitAnswered([pay(C.partKey(k, "prov"), "collected", 40)], k, p), false);
  assert.equal(C.visitAnswered([pay(C.partKey(k, "prov"), "collected", 40), pay(C.partKey(k, "inf"), "missed", 300)], k, p), true);
  assert.equal(C.visitAnswered([pay(k, "collected", 340)], k, p), true, "a row for the whole visit from before the split");
  const L = C.ledger("990201", [pay(C.partKey(k, "prov"), "collected", 40), pay(C.partKey(k, "inf"), "missed", 300)]);
  assert.equal(L.balance, 300);
  const A = C.appts([AP_H, ap("990201", "ALPPAT", T - 2, 540)]);
  assert.equal(C.toConfirm(p, A, [pay(C.partKey(k, "prov"), "collected", 40)], T, T - 10).length, 1, "the infusion is still to confirm");
});

test("demand by protocol: needed at the protocol, booked, to book; paused and unscheduled filters", () => {
  const P1 = pts([["990301", "WEEPAT", "Ongoing", "1x per week", "", "", "", "", "", ""], ["990302", "TWOPAT", "Ongoing", "every 2 weeks", "", "", "", "", "", ""],
    ["990303", "PAUPAT", "Paused", "1x per week", "", "", "", "", "", ""], ["990304", "PRNPAT", "Ongoing", "PRN", "", "", "", "", "", ""]]);
  const A = C.appts([AP_H, ap("990301", "WEEPAT", T + 1, 540), ap("990301", "WEEPAT", T + 8, 540), ap("990302", "TWOPAT", T + 9, 540)]);
  const D = C.demand({ patients: P1, appts: A, holds: [], today: T, cfg, weeks: 8, paused: false, sched: "all" });
  assert.deepEqual(D.rows.map(r => r.code).sort(), ["PRNPAT", "TWOPAT", "WEEPAT"], "paused left out");
  const g = Object.fromEntries(D.groups.map(x => [x.freq, x]));
  assert.equal(g["1x per week"].rate, 1); assert.equal(g["every 2 weeks"].rate, 0.5); assert.equal(g["PRN"].rate, 0);
  assert.ok(Math.abs(g["1x per week"].d8 - (7 + 4 / 7)) < 1e-9, "this week counts from today (Wednesday): 4 of 7 days");
  assert.equal(g["1x per week"].b8, 2);
  assert.equal(C.demand({ patients: P1, appts: A, holds: [], today: T, cfg, weeks: 8, paused: true, sched: "all" }).rows.length, 4);
  assert.deepEqual(C.demand({ patients: P1, appts: A, holds: [], today: T, cfg, weeks: 8, sched: "only" }).rows.map(r => r.code), ["PRNPAT"]);
  assert.equal(C.demand({ patients: P1, appts: A, holds: [], today: T, cfg, weeks: 8, sched: "exclude" }).rows.length, 2);
});

test("index patients come 2 a week to treatment 6, then once a week: the step-down frees access from that week", () => {
  const start = T - 3, P1 = pts([["990401", "IDXPAT", "Induction", "2x per week", "", "", "", "", C.isoOf(start), ""]]);
  /* treatments 1 and 2 done (Sun-based list starts at the course start), 3 to 6 booked over the next two weeks */
  const days = [start, start + 2, T + 1, T + 5, T + 8, T + 12];
  const A = C.appts([AP_H, ...days.map((d, i) => ap("990401", "IDXPAT", d, 540, i ? "IV Ketamine Follow Up" : "IV Ketamine Start"))]);
  const D = C.demand({ patients: P1, appts: A, holds: [], today: T, cfg, weeks: 6, sched: "all" });
  const r = D.rows[0];
  assert.equal(r.idx, true); assert.equal(r.idxEnd, T + 12, "treatment 6 as booked"); assert.equal(r.afterRate, 1);
  const w = D.weeks.find(x => x.start <= T + 12 && T + 12 < x.start + 7);
  assert.ok(w.steps.some(x => x.code === "IDXPAT"));
  assert.ok(D.weeks[D.weeks.length - 1].freed > 0.99, "a full visit a week free once they're past it");
  assert.equal(D.weeks[0].freed, 0);
  /* only Induction is index: the same course on Ongoing counts at its own protocol */
  const O = pts([["990401", "IDXPAT", "Ongoing", "1x per week", "", "", "", "", C.isoOf(start), ""]]);
  assert.equal(C.demand({ patients: O, appts: A, holds: [], today: T, cfg, weeks: 6, sched: "all" }).rows[0].idx, false);
});

test("index done: an Induction patient who has kept 6 visits is asked to move to Ongoing; no-shows don't count", () => {
  const start = T - 20, P1 = pts([["990501", "SIXPAT", "Induction", "2x per week", "", "", "", "", C.isoOf(start), ""]]);
  const days = [start, start + 3, start + 7, start + 10, start + 14, start + 17];
  const A = C.appts([AP_H, ...days.map((d, i) => ap("990501", "SIXPAT", d, 540, i ? "IV Ketamine Follow Up" : "IV Ketamine Start"))]);
  const x = { today: T, cfg, appts: A, patients: P1 };
  assert.equal(C.indexState(P1[0], A, T, cfg).complete, true);
  assert.ok(C.needs(P1[0], x).some(n => n.k === "indexdone" && n.act[1] === "idxDone"));
  const ns = A.map((a, i) => i === 5 ? Object.assign({}, a, { status: "No Show" }) : a);
  assert.equal(C.indexState(P1[0], ns, T, cfg).complete, false, "5 kept");
  /* today's 6th visit counts once it's checked out */
  const today6 = C.appts([AP_H, ...days.slice(0, 5).map((d, i) => ap("990501", "SIXPAT", d, 540, i ? "IV Ketamine Follow Up" : "IV Ketamine Start")), ap("990501", "SIXPAT", T, 540)]);
  assert.equal(C.indexState(P1[0], today6, T, cfg).complete, false);
  assert.equal(C.indexState(P1[0], today6, T, cfg, a => a.date === T).complete, true);
  assert.equal(C.indexState(pts([["990502", "ONGPAT", "Ongoing", "1x per week", "", "", "", "", C.isoOf(start), ""]])[0], A, T, cfg), null);
});
