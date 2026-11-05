// Clinic closures, no-shows and late cancels, step-down to 1x a week (invented codes, MRNs and providers only)
const test = require("node:test"), assert = require("node:assert/strict");
const C = require("../core.js"), TM = require("../template.js"), P = require("./pattern.js");

const T = C.serialOf(new Date(2026, 9, 21));   // Wednesday Oct 21 2026
const cfg = Object.assign({}, C.DEFAULTS), types = TM.DEFAULT_TYPES;
const RIV = "Rivera MD, Ana Maria";
const every = new Map(); for (let d = 1; d <= 5; d++) { every.set(`${d}|540`, "IV Ketamine Start/Index"); every.set(`${d}|630`, "IVK Post Check"); every.set(`${d}|600`, "IV Ketamine Follow Up"); }
const GRIDS = TM.parseGrids("IVK Template", P.sheet(TM, [TM.layout({ label: "Rivera", copy: every, types }, 1)]));
const AP_H = ["MRN", "Patient", "Date", "Time", "Provider", "Event", "Status", "Location", "Name", "Lookup Key"];
const ap = (mrn, code, date, min, status = "Expected", event = "IV Ketamine Follow Up") => [mrn, "", date, min / 1440, RIV, event, status, "Castro Street", code, ""];
const EX_H = C.EX_COLS;
const ex = (id, prov, from, to, note) => [id, prov, C.isoOf(from), to == null ? "" : C.isoOf(to), "", "", note || "", "", "", ""];
const MPT_H = ["MRN", "Patient Name", "Status", "Frequency", "Copay", "Infusion", "Primary Encounter Insurance", "Infusion Insurance", "Tx Start Date", "Protocol Updated"];
const pts = rows => C.patients([MPT_H, ...rows]);
const MS_H = C.MISS_COLS;
const ms = (date, mrn, time, kind, exempt = "", reason = "") => [C.visitKey(date, mrn, time), C.isoOf(date), C.fmtT(time), mrn, "", RIV, kind, exempt, reason, "desk", "", ""];
const X = (appts, EX, late) => Object.assign({ appts: C.appts([AP_H, ...appts]), holds: [], others: [], cfg, chairs: 4, needsPost: () => false, hasGrid: () => true, late },
  C.withExceptions(TM.onDays(GRIDS, types, 90), EX ? C.exRows([EX_H, ...EX]) : []));

test("a clinic closure closes every provider's slots and flags what's still booked those days", () => {
  const EX = C.exRows([EX_H, ex("c1", "Clinic closed", T + 8, T + 9, "Thanksgiving"), ex("p1", RIV, T + 1)]);
  assert.equal(EX[0].closed, true); assert.equal(EX[0].provider, C.CLOSED);
  const x = X([ap("990201", "ONEPAT", T + 8, 600)], [ex("c1", "Clinic closed", T + 8, T + 9, "Thanksgiving")]);
  assert.equal(C.daySlots(T + 8, x).slots.length, 0, "no slots on the closed Thursday");
  assert.equal(C.daySlots(T + 9, x).slots.length, 0, "or the Friday");
  assert.ok(C.daySlots(T + 7, x).slots.length > 0, "the Wednesday before is open");
  assert.equal(C.closureOn(EX, T + 8).note, "Thanksgiving");
  const A = C.appts([AP_H, ap("990201", "ONEPAT", T + 8, 600), ap("990202", "TWOPAT", T + 2, 600)]);
  const cc = C.closureClashes(EX, A, [], T);
  assert.equal(cc.length, 1); assert.deepEqual(cc[0].list.map(a => a.mrn), ["990201"]);
  const p = pts([["990201", "ONEPAT", "Ongoing", "1x per week", "", "", "Aetna", "", "", ""]])[0];
  const n = C.needs(p, { today: T, cfg, appts: A, EX });
  assert.ok(n.some(y => y.k === "closed" && y.bad), "the patient is flagged red");
});

test("no-shows come from NextGen, late cancels from the desk and uploads; Kaiser and VA are exempt", () => {
  const A = C.appts([AP_H, ap("990301", "KAIPAT", T - 2, 540, "No Show"), ap("990302", "AETPAT", T - 1, 540, "Cancelled"), ap("990303", "CLNPAT", T - 3, 600, "No Show"), ap("990304", "NOTPAT", T - 1, 600, "Cancelled")]);
  const PM = new Map(pts([["990301", "KAIPAT", "Ongoing", "1x per week", "", "", "Kaiser Permanente", "", "", ""], ["990302", "AETPAT", "Ongoing", "1x per week", "", "", "Aetna", "", "", ""],
    ["990303", "CLNPAT", "Ongoing", "1x per week", "", "", "TriWest VA CCN", "", "", ""]]).map(p => [p.mrn, p]));
  const MR = C.missRows([MS_H, ms(T - 1, "990302", 540, "late"), ms(T - 3, "990303", 600, "noshow", "clinical", "provider advised"), ms(T - 1, "990304", 600, "notlate")]);
  const L = C.missList({ appts: A, MR, PM, PF: new Map() }), by = m => L.find(x => x.mrn === m);
  assert.equal(L.length, 3, "the not-late cancel isn't counted");
  assert.equal(by("990301").kind, "noshow"); assert.equal(by("990301").exempt, "kpva");
  assert.equal(by("990302").kind, "late"); assert.equal(by("990302").exempt, "");
  assert.equal(by("990303").exempt, "kpva", "VA (TriWest) wins over the clinical reason");
  assert.ok(C.isKpVa({ ins1: "VA Community Care" })); assert.ok(!C.isKpVa({ ins1: "Nevada Medicaid" }));
  const bp = C.missByPatient(L); assert.equal(bp.length, 3);
});

test("a late cancel keeps its time on the schedule; an early cancel frees it", () => {
  const rows = [ap("990401", "LATPAT", T + 1, 540, "Cancelled"), ap("990402", "ERLPAT", T + 1, 600, "Cancelled")];
  const late = C.lateKeys(C.missRows([MS_H, ms(T + 1, "990401", 540, "late")]));
  const S = C.daySlots(T + 1, X(rows, null, late)).slots, at = t => S.find(s => s.time === t);
  assert.equal(at(540).open, false, "9:00 is held by the late cancel");
  assert.equal(at(540).by.kind, "late");
  assert.equal(at(600).open, true, "10:00 was cancelled early: open");
});

test("an upload finds cancellations close to the visit: same day is late, the rest go to the desk", () => {
  const oldA = C.appts([AP_H, ap("990501", "TODPAT", T, 540), ap("990502", "TOMPAT", T + 1, 540), ap("990503", "FARPAT", T + 9, 540), ap("990504", "GONPAT", T + 1, 600)]);
  const newA = C.appts([AP_H, ap("990501", "TODPAT", T, 540, "Cancelled"), ap("990502", "TOMPAT", T + 1, 540, "Rescheduled"), ap("990503", "FARPAT", T + 9, 540, "Cancelled"), ap("990505", "NEWPAT", T + 2, 540)]);
  const morning = T + 0.3, NC = C.newCancels(oldA, newA, T, cfg, morning), by = m => NC.find(x => x.a.mrn === m);
  assert.equal(by("990501").sameDay, true, "today's visit, old list from this morning: late, no question");
  assert.equal(by("990502").sameDay, false, "tomorrow: the desk says");
  assert.ok(!by("990503"), "nine days out isn't close");
  assert.ok(by("990504"), "gone from a list that still covers that day");
  assert.equal(C.newCancels(oldA, newA, T, cfg, T - 1 + 0.5).find(x => x.a.mrn === "990501").sameDay, false, "old list from yesterday: the desk says");
});

test("what no-shows and late cancels cost: past weeks booked, lost and left open", () => {
  const w = C.weekStart(T) - 7, rows = [];
  for (let i = 1; i <= 5; i++) rows.push(ap("9906" + i, "PASPAT", w + i, 540, i === 2 ? "No Show" : i === 4 ? "Cancelled" : "Kept"));
  const A = C.appts([AP_H, ...rows]), MR = C.missRows([MS_H, ms(w + 4, "99064", 540, "late")]);
  const L = C.missList({ appts: A, MR, PM: new Map(), PF: new Map() });
  const M = C.missImpact({ appts: A, X: X([]), missed: L, today: T, weeks: 2 });
  assert.equal(M.rows.length, 1, "only the week with visits on the list");
  const r = M.rows[0];
  assert.equal(r.kept, 3); assert.equal(r.noshow, 1); assert.equal(r.late, 1); assert.equal(r.booked, 5);
  assert.equal(r.slots, 10, "two IVK starts a day, Monday to Friday"); assert.equal(r.endOpen, 7);
  assert.equal(M.rate, 2 / 5);
});

test("kept all 6 at 2x a week and the protocol not updated since: step down to 1x a week", () => {
  const rows = []; [0, 3, 7, 10, 14, 17].forEach((d, i) => rows.push(ap("990701", "STPPAT", T - 30 + d, 540, "Kept", i === 0 ? "IV Ketamine Start" : undefined)));
  const A = C.appts([AP_H, ...rows]), start = T - 30;
  const mk = (st, f, proto) => pts([["990701", "STPPAT", st, f, "", "", "Aetna", "", C.isoOf(start), proto == null ? "" : C.isoOf(proto)]])[0];
  const sd = C.stepDue(mk("Ongoing", "2x per week"), A, T, cfg);
  assert.ok(sd); assert.equal(sd.N, 6); assert.equal(sd.to, "1x per week");
  assert.equal(C.stepDue(mk("Ongoing", "2x per week", T - 2), A, T, cfg), null, "protocol updated after the 6th");
  assert.ok(C.stepDue(mk("Ongoing", "2x per week", T - 25), A, T, cfg), "updated before the 6th still asks");
  assert.equal(C.stepDue(mk("Ongoing", "1x per week"), A, T, cfg), null, "already stepped down");
  assert.equal(C.stepDue(mk("Induction", "2x per week"), A, T, cfg), null, "Induction gets Move to Ongoing instead");
  const few = C.appts([AP_H, ...rows.slice(0, 5)]);
  assert.equal(C.stepDue(mk("Ongoing", "2x per week"), few, T, cfg), null, "only 5 kept");
  const n = C.needs(mk("Ongoing", "2x per week"), { today: T, cfg, appts: A });
  assert.ok(n.some(y => y.k === "stepdown" && y.act[1] === "stepDown"));
});
