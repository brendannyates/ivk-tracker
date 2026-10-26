// node --test tests/   (invented codes and MRNs only)
const test = require("node:test"), assert = require("node:assert/strict");
const C = require("../core.js");

const T = C.serialOf(new Date(2026, 9, 21));   // Wednesday Oct 21 2026
const MPT_H = ["MRN", "Patient Name", "Status", "Auth Exp Date", "Encounters Remaining", "Copay", "Infusion", "Protocol Updated", "Frequency", "Primary Location",
  "Last Treatment ", "Next Treatment", "Scheduled Thru", "Primary Encounter Insurance", "Infusion Insurance", "AM", "PM", "AM2", "PM3", "AM4", "PM5", "AM6", "PM7", "AM8", "PM9",
  "Tx Start Date", "2026 BIDF", "Auth Number", "Auth Start", "Encounters\r\nAuthorized", "PCN", "PCN\r\nManager", "Task Status", "PCN Notes", "Clinic Notes"];
const pt = o => MPT_H.map(h => ({ MRN: o.mrn, "Patient Name": o.code, Status: o.status ?? "Ongoing", "Auth Exp Date": o.authExp ?? "", "Encounters Remaining": o.left ?? "N/A",
  "Protocol Updated": o.proto ?? "", Frequency: o.freq ?? "every 2 weeks", "Primary Location": o.loc ?? "Castro", "Infusion Insurance": o.ins2 ?? "Self Pay", AM: 1, PM4: 0,
  "Auth Start": o.authStart ?? "", "Encounters\r\nAuthorized": o.auth ?? "N/A" }[h] ?? ""));
const AP_H = ["MRN", "Patient", "Date", "Time", "Provider", "Event", "Status", "Location", "Name", "Lookup Key"];
const ap = (mrn, code, date, min, status = "Expected") => [mrn, "", date, min / 1440, "Rivera MD, Ana Maria", "IV Ketamine Follow Up", status, "Castro Street", code, ""];
const cfg = Object.assign({}, C.DEFAULTS);

function world(extra = {}) {
  const P = C.patients([MPT_H,
    pt({ mrn: "990001", code: "TESANN" }),                                         // today, nothing after -> book + proto
    pt({ mrn: "990002", code: "DOEJON", proto: T - 10 }),                           // today, booked 6 weeks out, proto fresh
    pt({ mrn: "990003", code: "ROEMAR", freq: "PRN" }),                             // PRN: never asked to book
    pt({ mrn: "990004", code: "LEEKIM", authExp: T - 1, left: 0, auth: 16, authStart: T - 90 }), // expired auth, none left
    pt({ mrn: "990005", code: "PARSAM", status: "Discontinued" }),
    pt({ mrn: "990006", code: "OAKAVA", loc: "Oakland", status: "Paused" })]);
  const A = C.appts([AP_H,
    ap("990001", "TESANN", T, 13 * 60 + 30), ap("990002", "DOEJON", T, 12 * 60 + 30), ap("990002", "DOEJON", T + 42, 12 * 60 + 30),
    ap("990003", "ROEMAR", T, 9 * 60), ap("990004", "LEEKIM", T + 7, 9 * 60), ap("990001", "TESANN", T - 14, 13 * 60 + 30, "Kept"),
    ap("990006", "OAKAVA", T, 15 * 60, "Cancelled")]);
  return Object.assign({ today: T, cfg, appts: A, patients: P, visits: new Map(), clinic: "Castro" }, extra);
}

test("helpers: codes, providers, times, dates", () => {
  assert.equal(C.lasfir("Smith, Jane"), "SMIJAN");
  assert.equal(C.lasfir("Jane Smith"), "SMIJAN");
  assert.equal(C.lasfir("TESANN"), "TESANN");
  assert.equal(C.lasfir("O'Neil, Mary Ann"), "ONEMAR");
  assert.equal(C.provName("Rivera MD, Ana Maria"), "Dr. Rivera");
  assert.equal(C.provName("Chen, MD Paul"), "Dr. Chen");
  assert.equal(C.provName("Chen, Paul MD"), "Dr. Chen");
  assert.equal(C.toMin("1:30 PM"), 810); assert.equal(C.toMin(0.5625), 810); assert.equal(C.toMin("13:30"), 810); assert.equal(C.toMin("12:05 AM"), 5);
  assert.equal(C.toSerial("10/21/2026"), T); assert.equal(C.toSerial("2026-10-21"), T); assert.equal(C.toSerial(T + 0.4), T); assert.equal(C.toSerial("none scheduled"), null);
  assert.equal(C.fmtT(545), "9:05"); assert.equal(C.fmtD(T), "Oct 21"); assert.equal(C.fmtDay(T), "Wednesday, October 21");
  assert.equal(C.weekStart(T), T - 3);   // Sunday Oct 18
  assert.equal(C.freqDays("every 3 weeks"), 21); assert.equal(C.freqDays("PRN"), null);
  assert.equal(C.mrnKey("0099001"), "99001"); assert.equal(C.mrnKey(99001), "99001");
});

test("Master Patient Tracker rows read by header", () => {
  const P = world().patients, a = P.find(p => p.code === "LEEKIM");
  assert.equal(P.length, 6);
  assert.equal(a.mrn, "990004"); assert.equal(a.encLeft, 0); assert.equal(a.encAuth, 16); assert.equal(a.authExp, T - 1);
  assert.equal(P[0].freq, "every 2 weeks"); assert.equal(P[0].prefs[0].am, true); assert.equal(P[0].prefs[0].pm, false);
});

test("the team's ask reads the way they send it", () => {
  const x = world(), p = x.patients[0], a = C.asks(p, x, null);
  assert.deepEqual([a.book, a.proto], [true, true]);
  assert.equal(C.askText(p, a), "Can we please confirm every 2 week protocol for TESANN and send to the desk at discharge for scheduling?");
  assert.equal(C.askText(p, { book: false, proto: true }), "Can we please confirm every 2 week protocol for TESANN?");
  assert.equal(C.askText(p, { book: true, proto: false }), "Can we please send TESANN to the desk at discharge for scheduling?");
  assert.equal(C.askText(Object.assign({}, p, { freq: "" }), a), "Can we please confirm the protocol for TESANN and send to the desk at discharge for scheduling?");
  assert.equal(C.askText(p, a, { protocol: "Confirm {protocol} for {patient}, desk at discharge." }), "Confirm every 2 week for TESANN, desk at discharge.");
});

test("asks: booking window, PRN, protocol age, and what the desk already did", () => {
  const x = world(), P = n => x.patients.find(p => p.code === n);
  assert.equal(C.asks(P("DOEJON"), x, null).book, false, "booked 6 weeks out is enough");
  assert.equal(C.asks(P("DOEJON"), x, null).proto, false, "confirmed 10 days ago");
  assert.equal(C.asks(P("DOEJON"), Object.assign({}, x, { cfg: Object.assign({}, cfg, { AheadWeeks: 7 }) }), null).book, true);
  assert.equal(C.asks(P("ROEMAR"), x, null).book, false, "PRN is never asked to book");
  assert.equal(C.asks(P("PARSAM"), x, null).any, false, "discontinued");
  assert.equal(C.asks(P("TESANN"), x, { booked: "yes", proto: "" }).book, false);
  assert.equal(C.asks(P("TESANN"), x, { booked: "later", proto: "" }).book, true, "book later keeps reminding");
  assert.equal(C.asks(P("TESANN"), x, { booked: "", proto: "every 2 weeks" }).proto, false);
  const old = Object.assign({}, P("DOEJON"), { protoAt: T - 91 }); assert.equal(C.asks(old, x, null).proto, true, "older than ProtocolDays");
});

test("today: phases from the desk's rows, cancelled visits dropped", () => {
  const V = new Map(), k1 = C.visitKey(T, "990001", 810), k2 = C.visitKey(T, "990002", 750), k3 = C.visitKey(T, "990003", 540);
  assert.equal(k1, "2026-10-21|990001|13:30");
  V.set(k1, { key: k1, arr: 805, chair: "", seat: null, out: null, booked: "", asked: null, proto: "", status: "" });
  V.set(k2, { key: k2, arr: 740, chair: "2", seat: 745, out: null, booked: "", asked: 746, proto: "", status: "" });
  V.set(k3, { key: k3, arr: 530, chair: "1", seat: 535, out: 630, booked: "no", asked: null, proto: "", status: "" });
  const L = C.today(world({ visits: V })), ph = Object.fromEntries(L.map(v => [C.code(v), v.ph]));
  assert.deepEqual(ph, { ROEMAR: "done", DOEJON: "chair", TESANN: "waiting" });
  assert.equal(L.find(v => v.a.code === "OAKAVA"), undefined);
  assert.deepEqual(L.map(v => C.code(v)), ["ROEMAR", "DOEJON", "TESANN"], "sorted by visit time");
  /* a desk entry whose visit was moved in NextGen still shows while the patient is here */
  const k4 = C.visitKey(T, "990004", 600); V.set(k4, { key: k4, arr: 600, chair: "3", seat: 601, out: null, booked: "", asked: null, proto: "", status: "" });
  const L2 = C.today(world({ visits: V })), off = L2.find(v => v.key === k4);
  assert.equal(off.ph, "chair"); assert.equal(off.off, true); assert.equal(C.code(off), "LEEKIM");
  V.set(k1, Object.assign({}, V.get(k1), { status: "cancelled" })); assert.equal(C.today(world({ visits: V })).find(v => v.key === k1).ph, "gone");
});

test("chair states: open, track, soon, ready; waiting is late after LateMin", () => {
  const V = new Map(), k = C.visitKey(T, "990002", 750);
  V.set(k, { key: k, arr: 740, chair: "2", seat: 745, out: null, booked: "", asked: null, proto: "", status: "" });
  const L = C.today(world({ visits: V }));
  assert.equal(C.chairState("1", L, 800, cfg).state, "open");
  const tr = C.chairState("2", L, 800, cfg); assert.equal(tr.state, "track"); assert.equal(tr.end, 865); assert.equal(C.chairStatus(tr), "Out around 14:25"); assert.equal(tr.check, 775);
  const so = C.chairState("2", L, 858, cfg); assert.equal(so.state, "soon"); assert.equal(C.chairStatus(so), "Out in 7 min");
  const re = C.chairState("2", L, 866, cfg); assert.equal(re.state, "ready"); assert.equal(C.chairStatus(re), "Ready to check out");
  assert.equal(C.chairStatus(C.chairState("2", L, 875, cfg)), "Due out 10 min ago");
  const kw = C.visitKey(T, "990001", 810); V.set(kw, { key: kw, arr: 800, chair: "", seat: null, out: null, booked: "", asked: null, proto: "", status: "" });
  const w = C.today(world({ visits: V })).find(v => v.key === kw);
  assert.equal(C.isLate(w, 809, cfg), false); assert.equal(C.isLate(w, 810, cfg), true);
  assert.deepEqual(C.chairList(cfg, null), ["1", "2", "3", "4"]);
  assert.deepEqual(C.chairList(cfg, { bays: [{ id: "A" }, { id: "B" }] }), ["A", "B"]);
});

test("needs: red only for what blocks the next infusion, sorted red, amber, quiet", () => {
  const x = world(), P = n => x.patients.find(p => p.code === n);
  const lee = C.needs(P("LEEKIM"), x);
  assert.deepEqual(lee.filter(y => y.bad).map(y => y.k), ["auth", "enc"]);
  assert.equal(lee[0].t, "Auth ended Oct 20");
  const tes = C.needs(P("TESANN"), x);
  assert.equal(tes[0].k, "book"); assert.equal(!!tes[0].bad, false, "here today: the desk can book at discharge"); assert.equal(tes[0].t, "Nothing booked after today");
  const gone = C.needs(P("TESANN"), Object.assign({}, x, { appts: x.appts.filter(a => a.mrn !== "990001") }));
  assert.equal(gone[0].k, "book"); assert.equal(gone[0].bad, true); assert.equal(gone[0].t, "Nothing booked");
  assert.ok(tes.some(y => y.k === "proto" && !y.bad));
  assert.ok(tes.some(y => y.k === "selfpay" && y.quiet));
  assert.deepEqual(C.needs(P("PARSAM"), x), []);
  const oak = C.needs(P("OAKAVA"), x); assert.ok(oak.every(y => y.quiet)); assert.deepEqual(oak.map(y => y.k), ["paused", "home", "selfpay"]);
  assert.equal(C.loud(C.needs(P("ROEMAR"), x)).filter(y => y.k === "book").length, 0, "PRN");
  const soon = Object.assign({}, P("DOEJON"), { authExp: T + 10, encLeft: 2 });
  assert.deepEqual(C.loud(C.needs(soon, x)).map(y => [y.k, !!y.bad]), [["auth", false], ["enc", false]]);
});

test("arrival message", () => {
  const L = C.today(world()), v = L.find(y => y.a.code === "TESANN");
  assert.equal(C.arrivalText(v, "2"), "TESANN is here for their 13:30 IVK visit with Dr. Rivera. Chair 2.");
  assert.equal(C.arrivalText(v, ""), "TESANN is here for their 13:30 IVK visit with Dr. Rivera. Waiting for a chair.");
  assert.equal(C.fill(C.TEMPLATES.seated, { patient: "TESANN", chair: "3" }), "TESANN is in Chair 3.");
});

test("NextGen Appointments Listing export -> Appt List rows", () => {
  const aoa = [[], ["", "", "", "", "", "Mindful Health Solutions"], [], ["", "", "", "", "", "Appointments Listing"], [46199.41], [], [],
    ["", "MRN", "Location", "Appointment Date", "Time", "Provider", "Appointment Type", "Patient", "Status"],
    ["", "990001", "Castro Street", "10/21/2026", "1:30 PM", "Rivera MD, Ana Maria", "IV Ketamine Follow Up", "Test, Ann", "Expected"],
    ["", 990002, "Castro Street", T + 28, 0.5208333, "Chen, MD Paul", "IV Ketamine Treatment Start", "Doe, Jon", "Kept"],
    ["", "", "", "", "", "", "", "", ""], ["", "Total: 2", "", "", "", "", "", "", ""]];
  const rows = C.parseReport(aoa);
  assert.equal(rows.length, 2);
  assert.deepEqual(rows[0].slice(0, 3), ["990001", "Test, Ann", T]);
  assert.equal(Math.round(rows[0][3] * 1440), 810);
  assert.deepEqual(rows[1].slice(4), ["Chen, MD Paul", "IV Ketamine Treatment Start", "Kept", "Castro Street"]);
  assert.equal(C.parseReport([["a", "b"], [1, 2]]), null);
  const sh = C.shrink(rows, Array(40).fill(T + 60)); assert.equal(sh.flag, true); assert.equal(sh.few, true); assert.equal(sh.half, true); assert.equal(sh.short, true);
  assert.equal(C.shrink(Array(30).fill(["1", "", T]), Array(30).fill(T)).flag, false);
});

test("consults: approved-not-scheduled and unclear referrals are to act on", () => {
  const H = ["Location", "Referral Date", "Start Date", "Patient", "Patient MRN", "Consult Date", "Provider", "Patient Journey Stage", "Individual PJ Stage", "Days in current Stage", "Case Management Latest Status Note"];
  const L = C.consults([H, ["Castro", "", "", "Test, Ann", "990001", T - 20, "Chen, Paul MD", "4. PA Approved - Not Yet Scheduled", "4. PA Approved - Not Yet Scheduled", 14, ""],
    ["Castro", "", "", "Doe, Jon", "990002", T - 40, "Chen, MD Paul", "3. PA Submitted - Not Yet Approved", "", 35, ""],
    ["Castro", "", "", "Roe, Mar", "990003", T - 5, "Chen, MD Paul", "6. Approved - Scheduled", "", "", ""],
    ["", "", "", "", "", "", "", "", "", "", ""]], T);
  assert.equal(L.length, 3);
  assert.deepEqual(L.map(c => [c.code, c.state]), [["TESANN", "act"], ["DOEJON", "wait"], ["ROEMAR", "done"]]);
  assert.equal(L[0].flag, "Book the treatment start"); assert.equal(L[0].provider, "Dr. Chen");
});

test("week: Sunday-to-Saturday week, Mon-Fri columns, 30-min rows", () => {
  const x = world(), W = C.week(x.appts, C.weekStart(T), null);
  assert.deepEqual(W.days, [T - 2, T - 1, T, T + 1, T + 2]);
  assert.equal(W.rows[0], 540); assert.equal(W.rows[W.rows.length - 1], 930);
  assert.deepEqual(W.cell(T, 810).map(a => a.code), ["TESANN"]);
  assert.equal(W.count, 3, "the cancelled visit isn't counted");
});

test("a single word that isn't a code is never shown as one", () => {
  assert.equal(C.lasfir("Johnson"), ""); assert.equal(C.lasfir("JOHN"), "JOHN"); assert.equal(C.lasfir("TESANN"), "TESANN");
});

test("a patient listed twice: the active row wins, else the later one", () => {
  const rows = [["990001", "Discontinued"], ["990001", "Ongoing"], ["990001", "Paused"], ["990002", "Paused"], ["990002", "Discontinued"]];
  assert.equal(C.pickRow(rows, r => r[0], r => r[1], "990001"), 1);
  assert.equal(C.pickRow(rows, r => r[0], r => r[1], "990002"), 4);
  assert.equal(C.pickRow(rows, r => r[0], r => r[1], "990009"), -1);
  const P = C.patients([MPT_H, pt({ mrn: "990001", code: "TESANN", status: "Ongoing", freq: "every 2 weeks" }), pt({ mrn: "990001", code: "TESANN", status: "Discontinued", freq: "PRN" })]);
  assert.equal(P.length, 1); assert.equal(P[0].freq, "every 2 weeks"); assert.equal(P[0].i, 0);
});
