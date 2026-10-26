// Provider templates in the Clinic Template Tool's format (invented providers only)
const test = require("node:test"), assert = require("node:assert/strict");
const TM = require("../template.js"), P = require("./pattern.js");
const types = TM.DEFAULT_TYPES;
const lay = TM.layout({ title: "Ana Rivera, MD", label: "Rivera", location: "Castro", modified: 46300, from: 46290, days: [2, 3, 4, 5], copy: P.copyMap(TM), types }, 1);
const V = P.sheet(TM, [lay]), [g] = TM.parseGrids("IVK Template", V);

test("provider keys: NextGen names, template labels and titles meet on the last name", () => {
  for (const n of ["Rivera MD, Ana Maria", "Ana Rivera, MD", "Rivera", "Dr. Rivera", "Rivera, Ana MD", "Ana Rivera"]) assert.equal(TM.provKey(n), "rivera", n);
  assert.equal(TM.provKey("Chen, MD Paul"), "chen"); assert.equal(TM.provKey(""), "");
});

test("a written grid reads back the way the Clinic Template Tool reads one", () => {
  assert.ok(g, "grid found from its Monday..Friday header");
  assert.equal(g.label, "Rivera"); assert.equal(g.title, "Ana Rivera, MD"); assert.equal(g.key, "rivera"); assert.equal(g.location, "Castro");
  assert.equal(g.effFrom, 46290); assert.equal(g.modified, 46300); assert.equal(g.step, 30);
  assert.equal(g.rows[0].t, 480); assert.equal(g.rows[g.rows.length - 1].t, 1050, "8:00 AM to 5:30 PM");
  assert.equal(g.cells.get("3|540").type, "IV Ketamine Start/Index"); assert.equal(g.cells.has("1|540"), false, "Monday off");
  assert.deepEqual(V[lay.first].slice(0, 2), [480 / 1440, ""], "time column holds day fractions");
});

test("the tally matches the template the clinic built (slots, hours, by day, clinical, total)", () => {
  const t = TM.tally(g, types), row = n => t.rows.find(r => r.type === n);
  assert.deepEqual([row("IV Ketamine Start/Index").slots, row("IV Ketamine Start/Index").hours, row("IV Ketamine Start/Index").byDay], [11, 5.5, [0, 0, 5, 1, 5]]);
  assert.deepEqual(row("IV Ketamine Follow Up").byDay, [0, 0, 3, 3, 3]); assert.equal(row("IVK Post Check").slots, 4); assert.equal(row("Flex Post Check").slots, 8);
  /* the source sheet shows 60 (0, 15, 15, 16, 14): its Clinical row sums rows 5-14 for Monday but 5-15 (with Meeting) for Tue-Fri */
  assert.deepEqual([t.clinical.slots, t.clinical.hours, t.clinical.byDay], [57, 28.5, [0, 14, 15, 14, 14]]);
  assert.deepEqual([t.total.slots, t.total.byDay], [66, [0, 16, 17, 17, 16]], "Lunch isn't in the total");
});

test("the tab's tally is live formulas over the grid", () => {
  const f = lay.put.filter(p => p[1] === 9 && /^=COUNTIF/.test(p[2]));
  assert.equal(f.length, types.length);
  const first = lay.first + 1, last = lay.last + 1;
  assert.equal(f[0][2], `=COUNTIF($B$${first}:$F$${last},$I${f[0][0] + 1})`);
  const clin = lay.put.find(p => p[2] === "Clinical"), tot = lay.put.find(p => p[2] === "Total:");
  assert.ok(clin && tot && tot[0] === clin[0] + 1);
  const totF = lay.put.find(p => p[0] === tot[0] && p[1] === 9)[2], lunchRow = lay.put.find(p => p[2] === "Lunch" && p[1] === 8)[0] + 1;
  assert.ok(!totF.includes("J" + lunchRow), "Total skips Lunch: " + totF);
});

test("a start is an index start only when a post check lines up 90 min later", () => {
  const wed = TM.startsOn(g, 3, types, 90), at = t => wed.find(s => s.time === t);
  assert.equal(wed.filter(s => s.index).length, 5);
  assert.deepEqual([at(540).index, at(540).postUse], [true, "flexpost"]);
  assert.deepEqual([at(750).index, at(750).postUse], [true, "post"]);
  assert.equal(at(600).use, "followup"); assert.equal(at(600).index, false);
  assert.equal(TM.startsOn(g, 4, types, 90).filter(s => s.index).length, 1, "Thursday: only 9:00 has its post check");
});

test("checks name a start with no post check, and a type not on the Slot Types tab", () => {
  const g2 = TM.parseGrids("t", P.sheet(TM, [TM.layout({ label: "Chen", copy: new Map([["2|540", "IV Ketamine Start/Index"], ["2|600", "Mystery Block"]]), types }, 1)]))[0];
  const w = TM.checks(g2, types, 90);
  assert.ok(w.some(x => /Tuesday 9:00 AM: no post check at 10:30 AM/.test(x)), w.join(" | "));
  assert.ok(w.some(x => /Mystery Block/.test(x)));
  assert.deepEqual(TM.checks(g, types, 90), []);
});

test("two grids on one tab; the next goes under the lowest; Effective From picks the one in effect", () => {
  const lay2 = TM.layout({ label: "Rivera", from: 46400, copy: new Map([["3|540", "IV Ketamine Follow Up"]]), types }, TM.nextRow([g]));
  const gs = TM.parseGrids("t", P.sheet(TM, [lay, lay2]));
  assert.equal(gs.length, 2); assert.ok(gs[1].top > gs[0].bottom);
  assert.equal(TM.gridFor(gs, "rivera", 46350), gs[0]); assert.equal(TM.gridFor(gs, "rivera", 46401), gs[1]); assert.equal(TM.gridFor(gs, "rivera", 46200), null);
});

test("slot types: the tab's columns, guessed uses, accepts", () => {
  const T = TM.typeRows([TM.TYPE_COLS, ["IV Ketamine Start/Index", "Clinical", "IV Ketamine Treatment Start, IV Ketamine Follow Up", "", "start", "166569", "#ffffff"], ["Flex PC", "Clinical", "TMS Follow Up + Tele", "", "", "", ""], ["Admin", "Blocked", "", "", "", "", ""]]);
  assert.equal(T[0].fill, "#166569"); assert.equal(T[0].font, "#FFFFFF");
  assert.equal(T[1].use, "flexpost"); assert.equal(T[2].kind, "Blocked");
  assert.ok(TM.accepts(T[0], "IV Ketamine Follow Up")); assert.ok(!TM.accepts(T[0], "Telepsych"));
  assert.ok(TM.accepts({ accepts: "IVK" }, "IV Ketamine Follow Up"), "service line word");
  assert.ok(TM.accepts({ accepts: "*" }, "Anything"));
  assert.deepEqual(TM.eventsOf(T[0]), ["iv ketamine treatment start", "iv ketamine follow up"]);
});

test("onDays gives each day's starts and cells from the grid in effect", () => {
  const D = TM.onDays([g], types, 90), wed = 46316;   // Wed Oct 21 2026
  assert.equal(TM.dow(wed), 3);
  assert.equal(D.slotsOn(wed).filter(s => s.index).length, 5); assert.equal(D.slotsOn(wed)[0].pkey, "rivera");
  assert.equal(D.cellsOn(wed).length, 18); assert.equal(D.cellsOn(wed - 2).length, 0, "Monday off");
});

test("review: more name forms; Do and Pa are surnames unless written in capitals as credentials", () => {
  for (const [n, k] of [["Ellis Lark, M.D.", "lark"], ["Ellis Lark, MD, PhD", "lark"], ["Ellis Lark Jr.", "lark"], ["Do, Anh", "do"], ["Do", "do"], ["Pa, Lin DO", "pa"], ["Lin Pa, DO", "pa"]]) assert.equal(TM.provKey(n), k, n);
});

test("review: a blank label never reads a value from the tally or the next label", () => {
  const l2 = TM.layout({ label: "Chen", copy: new Map([["2|540", "IV Ketamine Start/Index"]]), types }, 1);   // no modified, no from
  const g2 = TM.parseGrids("t", P.sheet(TM, [l2]))[0];
  assert.equal(g2.modified, null); assert.deepEqual(g2.at.modified, [3, 5], "the cell right of Last Modified:, not the Event Types header");
  assert.equal(g2.effFrom, null); assert.deepEqual(g2.at.from, [2, 5]);
  /* a clinic file: "Last Modified: [blank] Effective From: 1/5/2026" on one row */
  const V = [["Lark, Ellis MD Template - Provider Configurable"], ["Location:", "Castro", "", "Last Modified:", "", "Effective From:", "1/5/2026"], ["Provider:", "", "Lark", "Lark", "Lark", "Lark"],
    ["Time", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"], [480 / 1440, "", "Admin"], [510 / 1440, "", "New Patient"]];
  const g3 = TM.parseGrids("t", V)[0];
  assert.equal(g3.modified, null); assert.deepEqual(g3.at.modified, [1, 4]); assert.equal(g3.effFrom, 46027);
  assert.equal(g3.title, "Lark, Ellis MD Template - Provider Configurable", "a title with Template in it is still the title"); assert.equal(g3.top, 0);
});

test("review: the next grid clears a long tally", () => {
  assert.ok(TM.nextRow([g], 30) >= g.anchor[0] - 2 + 32 + 4);
});
