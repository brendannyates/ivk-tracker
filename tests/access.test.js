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
  const na = C.coverage(pts([["990203", "NAPAT", "Ongoing", "", "30", "", "Kaiser", "N/A", "", ""]])[0], null);
  assert.equal(na.inf.cover, "", "N/A is no payer"); assert.equal(na.inf.payer, ""); assert.equal(na.mode, "", "no mode until both parts are set");
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
  assert.ok(Math.abs(g["1x per week"].t8 - (7 + 4 / 7 - 2)) < 1e-9, "to book = need less booked, per patient");
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

test("the Sohar IVN export (whole file): matched patients get their profile's Sohar fields, older and unchanged rows are left alone", () => {
  /* the Sohar Eligibility add-in's IVN export as it writes it (v0.27, invented patients) */
  const H = "Per Nbr,MRN,LASFIR,Sohar Payer,Plan,Network,Coverage Level,Coverage Start,Coverage End,Copay,Coinsurance,Deductible,Deductible Remaining,Deductible Status,OOP Max,OOP Remaining,OOP Status,Benefit Flag,Sohar Status,Sohar Date,Sohar Ref,Service,Source,Exported,Secondary Payer,Cost-share Flags,Clinic".split(",");
  const r = (mrn, code, date, ref, rem) => ["50", mrn, code, "Cigna", "Open Access", "", "", "", "", "40", "20%", "1000", rem, "Not met", "5000", "4000", "Not met", "", "complete.eligible", date, ref, "IVK", "Verification", "11/4/2026 9:00 AM", "", "", "Castro Street"];
  const M = C.parseSohar([H, r("990201", "ALPPAT", "2026-10-14T01:00:00Z", "q-1", "250"), r("990202", "BETPAT", "2026-10-14T01:00:00Z", "q-2", "100"),
    r("990203", "GAMPAT", "2026-10-01T01:00:00Z", "q-3", "900"), r("990204", "DELPAT", "2026-10-14T01:00:00Z", "q-4", "0"), r("990205", "EPSPAT", "2026-10-14T01:00:00Z", "q-5", "10")]);
  assert.equal(M.size, 5); assert.equal(M.get("990201").date, C.serialOf(new Date(2026, 9, 14)));
  const sh = (date, ref) => ({ sohar: { date, ref } });
  const PM = new Map([["990201", { mrn: "990201", code: "ALPPAT" }], ["990203", { mrn: "990203", code: "GAMPAT" }], ["990204", { mrn: "990204", code: "DELPAT" }]]);
  const PF = new Map([["990201", Object.assign({ code: "ALPPAT" }, sh(C.serialOf(new Date(2026, 9, 1)), "old"))], ["990203", Object.assign({ code: "GAMPAT" }, sh(C.serialOf(new Date(2026, 9, 10)), "newer"))],
    ["990204", Object.assign({ code: "DELPAT" }, sh(C.serialOf(new Date(2026, 9, 14)), "q-4"))]]);
  const A = [{ mrn: "990202", code: "BETPAT" }];
  const out = C.planSoharImport(M, { PM, PF, A, by: "Desk", now: "2026-11-04T17:00:00.000Z" });
  assert.deepEqual(out.items.map(i => [i.mrn, i.isNew]), [["990201", false], ["990202", true]]);
  assert.equal(out.older, 1, "a newer saved verification wins"); assert.equal(out.same, 1, "same date and ref is unchanged"); assert.equal(out.notHere, 1, "not on this tracker");
  const a = out.items[0].patch;
  assert.equal(a.DedRemaining, "250"); assert.equal(a.SoharDate, "2026-10-14"); assert.equal(a.SoharRef, "q-1"); assert.equal(a.By, "Desk"); assert.ok(!("Patient" in a), "an existing profile keeps its code");
  assert.equal(out.items[1].patch.Patient, "BETPAT", "a new profile row takes the appointment list's code");
  assert.ok(Object.keys(a).every(k => C.PROFILE_COLS.includes(k)), "only profile columns");
});

test("scheduled through: share of active patients with a protocol booked that week or later; PRN and paused don't count", () => {
  const P1 = pts([["990801", "WEKONE", "Ongoing", "1x per week"], ["990802", "BIWTWO", "Ongoing", "every 2 weeks"], ["990803", "PRNTHR", "Ongoing", "PRN"],
    ["990804", "PAUFOU", "Paused", "1x per week"], ["990805", "NONFIV", "Ongoing", "every 4 weeks"]]);
  const w0 = T - C.weekday(T);   // this week's Sunday
  const A = C.appts([AP_H, ap("990801", "WEKONE", w0 + 29, 600), ap("990802", "BIWTWO", w0 + 10, 600), ap("990803", "PRNTHR", w0 + 40, 600)]);
  const D = C.demand({ patients: P1, appts: A, holds: [], today: T, cfg, weeks: 6, paused: true, sched: "all" });
  assert.deepEqual(D.weeks.map(w => w.den), [3, 3, 3, 3, 3, 3]);                     // WEKONE, BIWTWO, NONFIV
  assert.deepEqual(D.weeks.map(w => w.thruN), [2, 2, 1, 1, 1, 0]);
  assert.equal(Math.round(D.weeks[4].thruPct * 100), 33);
});

test("projected need: each patient's cadence past their last booked visit; index weekly after treatment 6", () => {
  const w0 = T - C.weekday(T);
  const P1 = pts([["990811", "WEKONE", "Ongoing", "1x per week"], ["990812", "BIWTWO", "Ongoing", "every 2 weeks"], ["990813", "INDTHR", "Induction", "2x per week", "", "", "", "", T - 14]]);
  /* WEKONE booked through next Wednesday; BIWTWO nothing booked, last seen 10 days ago (due in 4 days); INDTHR has 4
     index treatments kept and 1 booked: 1 more index visit, then weekly */
  const A = C.appts([AP_H, ap("990811", "WEKONE", T + 7, 600), ap("990812", "BIWTWO", T - 10, 600),
    ...[14, 11, 7, 4].map(d => ap("990813", "INDTHR", T - d, 600)), ap("990813", "INDTHR", T + 1, 600)]);
  const D = C.demand({ patients: P1, appts: A, holds: [], today: T, cfg, weeks: 5, sched: "all" });
  const r = m => D.rows.find(x => x.mrn === m).proj;
  assert.deepEqual(r("990811"), [0, 0, 1, 1, 1]);   // booked this week and next; then weekly
  assert.deepEqual(r("990812"), [0, 1, 0, 1, 0]);   // due Sunday, then every 2 weeks
  assert.deepEqual(r("990813"), [0, 1, 1, 1, 1]);   // treatment 6 three and a half days after the 5th, then weekly
  assert.deepEqual(D.weeks.map(w => w.projNeed), [0, 2, 2, 3, 2]);
  assert.equal(D.proj4, 7);
  /* overdue: last seen 3 weeks ago on every 2 weeks counts this week */
  const O = C.demand({ patients: pts([["990814", "OVRFOU", "Ongoing", "every 2 weeks"]]), appts: C.appts([AP_H, ap("990814", "OVRFOU", T - 21, 600)]), holds: [], today: T, cfg, weeks: 4, sched: "all" });
  assert.deepEqual(O.rows[0].proj, [1, 0, 1, 0]);
  assert.equal(D.weeks[0].listed, 1);                // live visits on the list from today this week: INDTHR Thursday (WEKONE is next week)
});

test("no-shows and late cancels by patient over 30, 60 and 90 days", () => {
  const L = [{ mrn: "1", code: "AAA", kind: "noshow", date: T - 5 }, { mrn: "1", code: "AAA", kind: "late", date: T - 45 }, { mrn: "1", code: "AAA", kind: "noshow", date: T - 80 },
    { mrn: "2", code: "BBB", kind: "late", date: T - 20, exempt: "kpva" }, { mrn: "3", code: "CCC", kind: "noshow", date: T - 120 }, { mrn: "4", code: "DDD", kind: "review", date: T - 1 }];
  const W = C.missWindows(L, T);
  assert.deepEqual(W.map(x => [x.code, x.ns, x.lc, x.exempt]), [["AAA", [1, 1, 2], [0, 1, 1], ""], ["BBB", [0, 0, 0], [1, 1, 1], "kpva"]]);
});

test("Access weeks: every IVK consult at this clinic counts in its week but cancelled ones; the tile counts only those still expected", () => {
  const A = C.appts([AP_H, ap("990301", "ONEPAT", T + 1, 540, "IV Ketamine Start"), ap("990302", "TWOPAT", T + 8, 540, "IV Ketamine Start"), ap("990303", "THRPAT", T + 8, 600),
    ["990304", "", T + 9, 540 / 1440, RIV, "IV Ketamine Start", "Cancelled", "Castro Street", "FOUPAT", ""]]);
  const O = C.otherRows([C.OTHER_COLS, [C.isoOf(T + 1), "10:00 AM", RIV, "IVK Consult", "Expected", "Castro Street", "990311", "CONONE"], [C.isoOf(T + 2), "11:00 AM", RIV, "IV Ketamine Consult", "Expected", "", "990312", "CONTWO"],
    [C.isoOf(T + 2), "1:00 PM", RIV, "IV Ketamine Consult", "No Show", "Castro Street", "990313", "CONTRE"], [C.isoOf(T + 2), "2:00 PM", RIV, "IVK Consult", "Cancelled", "Castro Street", "990314", "CONFOU"],
    [C.isoOf(T + 3), "1:00 PM", RIV, "Esketamine Consult", "Expected", "Castro Street", "", ""],
    [C.isoOf(T + 8), "1:00 PM", RIV, "IVK Consult", "Expected", "Oakland", "990315", "CONFIV"], [C.isoOf(T + 15), "1:00 PM", RIV, "TMS Consult", "Expected", "Castro Street", "", ""]]);
  const B = C.weekBookings({ appts: A, others: O, today: T, weeks: 4, cfg: Object.assign({}, cfg, { Clinic: "Castro" }) });
  assert.deepEqual(B.weeks.map(w => [w.consults, w.starts]), [[3, 1], [0, 1], [0, 0], [0, 0]], "the no-show counts; cancelled, Oakland, esketamine and TMS don't");
  assert.deepEqual(B.weeks[0].CL.map(x => x.code), ["CONONE", "CONTWO", "CONTRE"], "each consult keeps its patient, by day and time");
  assert.equal(B.ahead, 2, "the tile: only the ones still expected"); assert.equal(B.ahead4, 2); assert.ok(B.any);
  assert.deepEqual(B.weeks[1].SL.map(x => x.code), ["TWOPAT"], "the week's starts, with who");
  assert.ok(!C.weekBookings({ appts: A, others: [], today: T, weeks: 2, cfg }).any, "no consults on the report");
});

test("custom protocols: any steps, read back from the cell, as a schedule with the first 4 weeks counted", () => {
  const P = C.protoParse(C.protoText({ name: "Slow start", note: "", steps: [{ n: 2, every: 1, weeks: 1 }, { n: 1, every: 1, weeks: 4 }, { n: 1, every: 2, weeks: null }] }));
  assert.equal(P.name, "Slow start"); assert.equal(P.steps.length, 3);
  const Sc = C.protoSchedule(P, 4);
  assert.deepEqual(Sc.lines.map(l => [l.from, l.to, l.text]), [[1, 1, "2 treatments in 1 week"], [2, 5, "1 a week for 4 weeks"], [6, null, "1 every 2 weeks, ongoing"]]);
  assert.equal(Sc.inHorizon, 5); assert.equal(Sc.total, null, "ongoing has no total");
  const trial = C.protoSchedule(C.protoParse('{"steps":[{"n":1,"every":1,"weeks":1}]}'));
  assert.deepEqual([trial.lines[0].text, trial.total], ["1 treatment", 1]);
  assert.equal(C.protoParse("not json"), null); assert.equal(C.protoParse('{"steps":[]}'), null);
  assert.equal(C.protoParse('{"steps":[{"n":1,"every":1,"weeks":null},{"n":1,"every":1,"weeks":2}]}').steps.length, 1, "only the last step can be ongoing");
  assert.equal(C.stageNo("3. PA Submitted - Not Yet Approved"), 3); assert.equal(C.stageNo("Unclear"), null);
});

test("approved not scheduled: Approved on the Consults tab and not started; Not Yet Approved, started and closed aside", () => {
  const H = ["Location", "Referral Date", "Start Date", "Patient", "Patient MRN", "Consult Date", "Provider", "Patient Journey Stage", "Individual PJ Stage", "Days in current Stage", "Kept Start"];
  const CS = C.consults([H, ["Castro", "", "", "Alp, Ann", "990401", T - 20, RIV, "4. PA Approved - Not Yet Scheduled", "", 5, ""],
    ["Castro", "", "", "Bet, Bo", "990402", T - 20, RIV, "4. PA Approved - Not Yet Scheduled", "", 5, ""],
    ["Castro", "", "", "Gam, Gil", "990403", T - 20, RIV, "3. PA Submitted - Not Yet Approved", "", 5, ""],
    ["Castro", "", "", "Del, Dee", "990404", T - 30, RIV, "4. PA Approved - Not Yet Scheduled", "", 5, ""],
    ["Castro", "", "", "Eps, Eve", "990405", T - 30, RIV, "7. Not Proceeding", "", 5, ""]], T);
  const A = C.appts([AP_H, ap("990402", "BETBO", T + 3, 540, "IV Ketamine Start"), ["990404", "", T - 5, 540 / 1440, RIV, "IV Ketamine Start", "Kept", "Castro Street", "DELDEE", ""]]);
  C.consultStarts(CS, A, new Map(), cfg);
  const W = C.approvedWaiting(CS, A, T);
  assert.deepEqual(W.map(x => x.c.mrn), ["990401", "990402"], "DELDEE started; GAMGIL isn't approved; EPSEVE is closed");
  assert.equal(W.find(x => x.c.mrn === "990402").booked.date, T + 3, "BETBO's start is booked, the stage lags");
});

test("Access weeks: a start held in the Index planner (treatment 1) counts as a start until NextGen has it", () => {
  const A = C.appts([AP_H, ap("990501", "ONEPAT", T + 8, 540, "IV Ketamine Start")]);
  const H = [{ id: "h1", mrn: "990502", date: T + 9, time: 540, n: 1, status: "held" }, { id: "h2", mrn: "990502", date: T + 12, time: 540, n: 2, status: "held" },
    { id: "h3", mrn: "990501", date: T + 8, time: 540, n: 1, status: "held" }, { id: "h4", mrn: "990503", date: T + 9, time: 600, n: 1, status: "released" }];
  const B = C.weekBookings({ appts: A, others: [], holds: H, listAt: null, today: T, weeks: 3, cfg });
  assert.deepEqual(B.weeks.map(w => [w.starts, w.held]), [[0, 0], [2, 1], [0, 0]], "the held start counts; treatment 2, the one NextGen has and the released one don't");
});
