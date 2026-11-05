// node --test tests/   (invented codes and MRNs only): the treatment start checklist and Case Management after a kept start
const test = require("node:test"), assert = require("node:assert/strict");
const C = require("../core.js");

const T = C.serialOf(new Date(2026, 10, 4));   // Wednesday Nov 4 2026
const cfg = Object.assign({}, C.DEFAULTS);
const AP_H = ["MRN", "Patient", "Date", "Time", "Provider", "Event", "Status", "Location"];
const START = "IV Ketamine Treatment Start", FU = "IV Ketamine Follow Up";
const ap = (mrn, code, date, min, event = FU, status = "Expected") => [mrn, code, date, min / 1440, "Rivera MD, Ana Maria", event, status, "Castro Street"];
const CS_H = ["Location", "Referral Date", "Start Date", "Patient", "Patient MRN", "Consult Date", "Provider", "Patient Journey Stage", "Individual PJ Stage", "Days in current Stage",
  "Case Management Latest Status Note", "Kept Start", "CM Updated", "CM Flag"];
const cs = (mrn, code, consult, stage, extra = {}) => ["Castro Street", consult, extra.start ?? "", code, mrn, consult, "Rivera MD, Ana Maria", stage, stage, 5, "", extra.kept ?? "", extra.cm ?? "", ""];

test("Consults: a kept start since the consult takes the row off the waitlist and flags Case Management until stage 6", () => {
  const A = C.appts([AP_H,
    ap("990201", "TESAAA", T - 20, 600, START, "Kept"),          // kept start, CM still at 5
    ap("990202", "TESBBB", T - 10, 600, START, "Kept"),          // kept, CM already at 6
    ap("990203", "TESCCC", T + 21, 600, START),                  // booked, not kept yet
    ap("990204", "TESDDD", T - 400, 600, START, "Kept"),         // an old start, before this consult
    ap("990205", "TESEEE", T - 3, 600, START, "Kept")]);         // kept, CM Updated on the tab
  const CS = C.consults([CS_H,
    cs("990201", "TESAAA", T - 40, "5. Scheduled to Start"),
    cs("990202", "TESBBB", T - 30, "6. Started"),
    cs("990203", "TESCCC", T - 15, "5. Scheduled to Start"),
    cs("990204", "TESDDD", T - 30, "4. PA Approved - Not Yet Scheduled"),
    cs("990205", "TESEEE", T - 30, "4. PA Approved - Not Yet Scheduled", { cm: T - 2 }),
    cs("990206", "TESFFF", T - 30, "7. Not Proceeding")], T);
  C.consultStarts(CS, A, new Map(), cfg);
  const by = Object.fromEntries(CS.map(c => [c.code, c]));
  assert.equal(by.TESAAA.state, "cm"); assert.equal(by.TESAAA.kept, T - 20); assert.match(by.TESAAA.flag, /update Case Management to 6\. Started/);
  assert.equal(by.TESBBB.state, "started"); assert.equal(by.TESBBB.flag, "");
  assert.equal(by.TESCCC.kept, null); assert.equal(by.TESCCC.state, "done");             // still "Scheduled", on no list
  assert.equal(by.TESDDD.kept, null); assert.equal(by.TESDDD.state, "act");               // the old start doesn't count
  assert.equal(by.TESEEE.state, "started");
  assert.equal(by.TESFFF.state, "closed");
  const W = C.waitlist(CS, new Map());
  assert.deepEqual(W.list.map(w => w.code), ["TESDDD"]);
  assert.equal(C.isStartedStage("5. Scheduled to Start / 6. Started"), false);
  assert.equal(C.isStartedStage("6. Started"), true);
});

test("Consults: the tab's Kept Start counts when the start has dropped off the appointment list; a checklist tick clears the flag", () => {
  const CS = C.consults([CS_H, cs("990211", "TESGGG", T - 90, "5. Scheduled to Start", { kept: T - 60 })], T);
  C.consultStarts(CS, C.appts([AP_H]), new Map(), cfg);
  assert.equal(CS[0].state, "cm"); assert.equal(CS[0].kept, T - 60);
  const ST = C.startRows([C.START_COLS, [C.startKey(T - 60, "990211"), C.isoOf(T - 60), "990211", "TESGGG", C.isoOf(T), "BY", ""]]);
  const CS2 = C.consults([CS_H, cs("990211", "TESGGG", T - 90, "5. Scheduled to Start", { kept: T - 60 })], T);
  C.consultStarts(CS2, C.appts([AP_H]), ST, cfg);
  assert.equal(CS2[0].state, "started"); assert.equal(CS2[0].cmAt, T);
});

test("start checklist: a newly booked start needs a profile; on the day, Case Management", () => {
  const A = C.appts([AP_H,
    ap("990301", "TESHHH", T + 9, 900, START),                   // new patient, booked: profile to do
    ap("990302", "TESIII", T + 9, 600, START),                   // on the tracker already: nothing to do yet
    ap("990303", "TESJJJ", T, 540, START, "Kept"),                // today, kept
    ap("990304", "TESKKK", T - 5, 540, START, "Kept"),            // kept 5 days ago, not ticked
    ap("990305", "TESLLL", T - 5, 600, START, "No Show"),         // no-show: nothing to update
    ap("990306", "TESMMM", T - 30, 600, START, "Kept"),           // past StartBackDays
    ap("990307", "TESNNN", T - 2, 600, START, "Cancelled"),
    ap("990303", "TESJJJ", T + 2, 540, FU)]);
  const PM = new Map([["990302", {}], ["990303", {}], ["990304", {}]]);
  const L = C.startList({ today: T, appts: A, PM, PF: new Map(), ST: new Map(), CS: [], visits: new Map(), cfg });
  const by = Object.fromEntries(L.map(s => [s.code, s]));
  assert.deepEqual(Object.keys(by).sort(), ["TESHHH", "TESJJJ", "TESKKK"]);
  assert.equal(by.TESHHH.profDue, true); assert.equal(by.TESHHH.cmDue, false);
  assert.equal(by.TESJJJ.today, true); assert.equal(by.TESJJJ.cmDue, true); assert.equal(by.TESJJJ.profDue, false);
  assert.equal(by.TESKKK.cmDue, true); assert.equal(by.TESKKK.kept, true);
  /* ticked on tblIvkStarts, or the Consults tab at stage 6: done; today's start still listed (ticked) */
  const ST = C.startRows([C.START_COLS, [C.startKey(T, "990303"), C.isoOf(T), "990303", "TESJJJ", C.isoOf(T), "BY", ""]]);
  const CS = C.consults([CS_H, cs("990304", "TESKKK", T - 40, "6. Started")], T);
  const L2 = C.startList({ today: T, appts: A, PM, PF: new Map(), ST, CS, visits: new Map(), cfg });
  assert.deepEqual(L2.map(s => s.code), ["TESJJJ", "TESHHH"]);
  assert.equal(L2[0].cmDone, true); assert.equal(L2[0].cmAt, T); assert.equal(L2[0].open, false);
  /* a profile row clears the profile item */
  const L3 = C.startList({ today: T, appts: A, PM, PF: new Map([["990301", {}]]), ST, CS, visits: new Map(), cfg });
  assert.equal(L3.find(s => s.code === "TESHHH"), undefined);
});

test("a start today that is still Expected: Case Management is due on the day; the Patients list says so", () => {
  const A = C.appts([AP_H, ap("990401", "TESOOO", T, 600, START)]);
  const P = { mrn: "990401", code: "TESOOO", status: "Induction", freq: "2x per week", loc: "Castro", copay: "0", infusion: "0", prefs: [] };
  const SL = C.startList({ today: T, appts: A, PM: new Map([["990401", P]]), PF: new Map(), ST: new Map(), CS: [], visits: new Map(), cfg });
  assert.equal(SL[0].cmDue, true); assert.equal(SL[0].kept, false);
  const n = C.needs(P, { today: T, cfg, appts: A, clinic: "Castro", SL });
  assert.ok(n.find(y => y.k === "cm" && /today/.test(y.t) && y.act[1] === "cmDone"));
  /* a past start that was never marked kept isn't asked about */
  const SLp = C.startList({ today: T + 1, appts: A, PM: new Map([["990401", P]]), PF: new Map(), ST: new Map(), CS: [], visits: new Map(), cfg });
  assert.equal(SLp.length, 0);
  /* unless the desk checked them out */
  const V = new Map([[C.visitKey(T, "990401", 600), { out: 720 }]]);
  const SLv = C.startList({ today: T + 1, appts: A, PM: new Map([["990401", P]]), PF: new Map(), ST: new Map(), CS: [], visits: V, cfg });
  assert.equal(SLv[0].cmDue, true);
});
