/* excel.js: every workbook read and write the IVK Tracker makes. Runs in the task pane only (full screen asks the pane, see app.js).
   The clinic's own tabs are found by sheet name (Master Patient Tracker, Appt List, Consults), so the table names don't matter.
   The add-in's own tables live on one hidden sheet, IVK Tracker Data, created by setup() and never overwritten:
     tblIvkVisits     A1   Key | Date | MRN | Patient | Appt | Arrived | Chair | Seated | Out | Booked | Asked | Protocol | Status | By | Updated
     tblIvkConfig     Q1   Key | Value
     tblIvkRoutes     T1   Event | Sends to | FlowUrl | Enabled
     tblIvkTemplates  Y1   Event | Template
     tblIvkLog        AB1  Id | Event | Patient | Chair | Message | SentBy | SentAt | Result
     tblIvkSlots      AK1  Id | Provider | Day | Time | Type                (the index template)
     tblIvkPlans      AQ1  MRN | Patient | Start | PerWeek | Weeks | Avail | Note | By | Updated
     tblIvkHolds      BA1  Id | MRN | Patient | Date | Time | Provider | Treatment | Status | PlacedBy | PlacedAt | DoneBy | DoneAt
   The Index Planner tab's grid is rewritten as plain values (fillPlanner), replacing the lookups into another workbook.
   Text cells only on that sheet ("1:22 PM" stays text). The chair map is a document setting (a map file can be larger than a cell holds). */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory(require("./core.js"));
  else root.IvkExcel = factory(root.IvkCore);
})(typeof self !== "undefined" ? self : this, function (C) {
  "use strict";

  const DATA_SHEET = "IVK Tracker Data";
  const SHEETS = { mpt: "Master Patient Tracker", appt: "Appt List", consults: "Consults", access: "Access" }, PLANNER = "Index Planner";
  const CH = 250;   // rows per write: big single syncs fail in Excel on the web
  const VISIT_READ = 400;   // the newest visit rows read on each refresh (today's are always among them)
  const SPEC = [
    { name: "tblIvkVisits", anchor: "A", columns: C.VISIT_COLS, rows: [C.VISIT_COLS.map(() => "")] },
    { name: "tblIvkConfig", anchor: "Q", columns: ["Key", "Value"], rows: Object.entries(C.DEFAULTS).map(([k, v]) => [k, String(v)]) },
    { name: "tblIvkRoutes", anchor: "T", columns: ["Event", "Sends to", "FlowUrl", "Enabled"], rows: [["arrival", "the RN", "", "TRUE"], ["protocol", "Castro IVK group chat", "", "TRUE"]] },
    { name: "tblIvkTemplates", anchor: "Y", columns: ["Event", "Template"], rows: Object.entries(C.TEMPLATES) },
    { name: "tblIvkLog", anchor: "AB", columns: ["Id", "Event", "Patient", "Chair", "Message", "SentBy", "SentAt", "Result"], rows: [["", "", "", "", "", "", "", ""]] },
    { name: "tblIvkSlots", anchor: "AK", columns: ["Id", "Provider", "Day", "Time", "Type"], rows: [["", "", "", "", ""]] },
    { name: "tblIvkPlans", anchor: "AQ", columns: C.PLAN_COLS, rows: [C.PLAN_COLS.map(() => "")] },
    { name: "tblIvkHolds", anchor: "BA", columns: C.HOLD_COLS, rows: [C.HOLD_COLS.map(() => "")] }
  ];
  const BIG = ["tblIvkVisits", "tblIvkLog", "tblIvkHolds"];
  const IP_RANGE = "A1:AZ60";   // the Index Planner tab: 9 weeks of Mon-Fri across, half hours down, the slot list below
  const colIdx = l => l.split("").reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
  const colL = i => { let s = ""; for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s; return s; };
  const isF = v => typeof v === "string" && v.startsWith("=");
  const sameName = (a, b) => String(a).trim().toLowerCase() === String(b).trim().toLowerCase();

  /* the first table on each of the clinic's tabs, plus the add-in's own tables by name */
  async function tableMap(ctx) {
    const ts = ctx.workbook.tables; ts.load("items/name,items/worksheet/name"); await ctx.sync();
    const m = { own: {} };
    for (const t of ts.items) {
      for (const [k, s] of Object.entries(SHEETS)) if (!m[k] && sameName(t.worksheet.name, s)) m[k] = t.name;
      if (/^tblIvk/.test(t.name)) m.own[t.name] = t.name;
    }
    return m;
  }
  const readTable = (ctx, name) => { const r = ctx.workbook.tables.getItem(name).getRange(); r.load("values"); return r; };

  /* Everything the screens need, in a few syncs. Missing pieces come back null so the app can say what's missing. */
  async function load() {
    return Excel.run(async ctx => {
      const tm = await tableMap(ctx), out = { tables: tm, missing: SPEC.filter(s => !tm.own[s.name]).map(s => s.name) };
      const wb = ctx.workbook; wb.load("name");
      const small = {};
      for (const k of ["mpt", "consults"]) if (tm[k]) small[k] = readTable(ctx, tm[k]);
      const TAIL = BIG;
      for (const s of SPEC) if (tm.own[s.name] && !TAIL.includes(s.name)) small[s.name] = readTable(ctx, s.name);
      const apSheet = wb.worksheets.getItemOrNullObject(SHEETS.appt), acSheet = wb.worksheets.getItemOrNullObject(SHEETS.access), ipSheet = wb.worksheets.getItemOrNullObject(PLANNER);
      apSheet.load("isNullObject"); acSheet.load("isNullObject"); ipSheet.load("isNullObject");
      const bodies = TAIL.filter(n => tm.own[n]).map(n => { const b = wb.tables.getItem(n).getDataBodyRange(); b.load("rowCount"); return [n, b]; });
      await ctx.sync();
      out.wbName = wb.name;
      for (const [k, r] of Object.entries(small)) out[k] = r.values;
      out.apptUpdated = null; out.freqs = []; out.planner = null;
      if (!ipSheet.isNullObject) { const r = ipSheet.getRange(IP_RANGE); r.load("values"); await ctx.sync(); out.planner = r.values; }
      if (!apSheet.isNullObject || !acSheet.isNullObject) {
        const upd = apSheet.isNullObject ? null : apSheet.getRange("L1:M1"), fr = acSheet.isNullObject ? null : acSheet.getRange("R5:R14");
        if (upd) upd.load("values"); if (fr) fr.load("values"); await ctx.sync();
        if (upd) out.apptUpdated = C.num(upd.values[0][1]);
        if (fr) out.freqs = fr.values.map(r => C.txt(r[0])).filter(Boolean);
      }
      if (tm.appt) { const r = readTable(ctx, tm.appt); await ctx.sync(); out.appt = r.values; }
      /* the newest rows only: today's check-ins and Teams posts are always among them */
      for (const [n, b] of bodies) {
        const h = wb.tables.getItem(n).getHeaderRowRange(); h.load("values");
        const cnt = b.rowCount, from = Math.max(0, cnt - VISIT_READ), part = b.getRow(from).getResizedRange(cnt - from - 1, 0); part.load("values");
        await ctx.sync();
        out[n === "tblIvkVisits" ? "visits" : n] = cnt ? [h.values[0], ...part.values] : [h.values[0]];
      }
      return out;
    });
  }

  /* Creates the data sheet and whichever of the add-in's tables are missing. Never writes over a cell that holds something. */
  async function setup() {
    return Excel.run(async ctx => {
      const tm = await tableMap(ctx);
      let ws = ctx.workbook.worksheets.getItemOrNullObject(DATA_SHEET); ws.load("isNullObject"); await ctx.sync();
      const isNew = ws.isNullObject; if (isNew) ws = ctx.workbook.worksheets.add(DATA_SHEET);
      const todo = SPEC.filter(s => !tm.own[s.name]), checks = todo.map(s => {
        const c = colIdx(s.anchor), addr = `${s.anchor}1:${colL(c + s.columns.length - 1)}${s.rows.length + 1}`, r = ws.getRange(addr); r.load("values"); return { s, addr, r };
      });
      await ctx.sync();
      const made = [], blocked = [];
      for (const { s, addr, r } of checks) {
        if (!isNew && r.values.some(row => row.some(v => v !== "" && v != null))) { blocked.push(s.name); continue; }
        const c = colIdx(s.anchor), last = colL(c + s.columns.length - 1);
        const n = BIG.includes(s.name) ? 2000 : 200;
        ws.getRange(`${s.anchor}2:${last}${n + 1}`).numberFormat = Array.from({ length: n }, () => s.columns.map(() => "@"));
        const t = ws.tables.add(addr, true); t.name = s.name; t.style = "TableStyleLight1";
        t.getHeaderRowRange().values = [s.columns];
        t.getDataBodyRange().values = s.rows.map(r => r.map(String));
        made.push(s.name);
      }
      if (made.length) { ws.getRange("A:BL").format.autofitColumns(); ws.visibility = Excel.SheetVisibility.hidden; }
      await ctx.sync();
      return { made, blocked };
    });
  }

  /* Upsert one visit by Key. patch = { Arrived: "13:22", Chair: "2", ... } by column name. Re-reads the rows inside the same run
     so two people at the desk don't overwrite each other's fields. blankFirst = a column that must still be empty (Arrived for a
     check-in): if someone else already filled it, nothing is written and the error says who. Returns the row as saved. */
  async function saveVisit(key, base, patch, blankFirst) {
    return Excel.run(async ctx => {
      const t = ctx.workbook.tables.getItem("tblIvkVisits"), h = t.getHeaderRowRange(), b = t.getDataBodyRange();
      h.load("values"); b.load("values"); await ctx.sync();
      const cols = h.values[0].map(C.norm), at = n => cols.indexOf(C.norm(n));
      if (at("Key") < 0) throw new Error("The IVK Visits table has no Key column. Set up the workbook again in Settings.");
      let i = -1; for (let j = b.values.length - 1; j >= 0; j--) if (String(b.values[j][at("Key")]) === key) { i = j; break; }
      if (blankFirst && i >= 0 && C.txt(b.values[i][at(blankFirst)])) {
        const by = C.txt(b.values[i][at("By")]);
        throw new Error(`${blankFirst} was already saved at ${C.txt(b.values[i][at(blankFirst)])}${by ? ` by ${by}` : ""}. The list refreshes in a moment`);
      }
      /* only the blank starter row from setup is reused; any other blank row could sit where load() doesn't read */
      const blank = i < 0 && b.values.length === 1 && b.values[0].every(v => v === "" || v == null) ? 0 : -1;
      const row = i >= 0 ? b.values[i].slice() : cols.map(() => "");
      for (const [k, v] of Object.entries(Object.assign({ Key: key }, i < 0 ? base : {}, patch))) { const c = at(k); if (c >= 0) row[c] = v == null ? "" : String(v); }
      const out = row.map(v => v == null ? "" : String(v));
      if (i >= 0) b.getRow(i).values = [out];
      else if (blank >= 0) b.getRow(blank).values = [out];
      else t.rows.add(null, [out]);
      await ctx.sync();
      return out;
    });
  }

  /* Master Patient Tracker: Protocol Updated = the day it was confirmed, and Frequency when it changed. The MRN is checked again first. */
  async function confirmProtocol(mrn, freq, serial) {
    return Excel.run(async ctx => {
      const tm = await tableMap(ctx); if (!tm.mpt) throw new Error("No Master Patient Tracker table in this workbook.");
      const t = ctx.workbook.tables.getItem(tm.mpt), h = t.getHeaderRowRange(), b = t.getDataBodyRange(); h.load("values"); b.load("values"); await ctx.sync();
      const cols = h.values[0].map(C.norm), cM = cols.indexOf("mrn"), cP = cols.indexOf("protocolupdated"), cF = cols.indexOf("frequency");
      if (cM < 0 || cP < 0) throw new Error("The Master Patient Tracker needs MRN and Protocol Updated columns.");
      const i = C.pickRow(b.values, r => r[cM], r => r[cols.indexOf("status")], mrn);
      if (i < 0) throw new Error("That patient isn’t on the Master Patient Tracker.");
      const pc = b.getCell(i, cP); pc.numberFormat = [["m/d/yyyy"]]; pc.values = [[serial]];
      let wroteFreq = false;
      if (cF >= 0 && freq && C.txt(b.values[i][cF]) !== C.txt(freq)) { b.getCell(i, cF).values = [[freq]]; wroteFreq = true; }
      await ctx.sync();
      return { row: i, wroteFreq };
    });
  }

  /* NextGen upload: rows = [MRN, Patient, Date, Time, Provider, Event, Status, Location] into the Appt List table's first 8
     columns. Grows the table carrying the formula columns (Name, Lookup Key) down; a shorter list clears the rows left over.
     L1:M1 gets "Last Updated:" and the time. Same steps as the ESK Tracker's upload. */
  const APPT_HDR = ["mrn", "patient", "date", "time", "provider", "event", "status", "location"];
  async function replaceAppts(rows, stamp) {
    return Excel.run(async ctx => {
      const tm = await tableMap(ctx); if (!tm.appt) throw new Error("No table on the Appt List tab.");
      const t = ctx.workbook.tables.getItem(tm.appt), h = t.getHeaderRowRange(); h.load("values");
      let b = t.getDataBodyRange(); b.load("rowCount"); await ctx.sync();
      const hn = h.values[0].map(C.norm), c0 = hn.indexOf("mrn"), had = b.rowCount, W = APPT_HDR.length;
      if (APPT_HDR.some((k, i) => hn[c0 + i] !== k)) throw new Error("Stopped: the Appt List table should start with MRN, Patient, Date, Time, Provider, Event, Status, Location, in that order.");
      const top = b.getCell(0, c0).getResizedRange(0, W - 1), last = b.getRow(had - 1); top.load("formulas"); last.load("formulasR1C1"); await ctx.sync();
      if (top.formulas[0].some(isF)) throw new Error("Stopped: one of the MRN to Location columns holds a formula.");
      if (rows.length > had) {
        t.resize(t.getRange().getResizedRange(rows.length - had, 0)); await ctx.sync();
        b = t.getDataBodyRange(); b.load("rowCount"); await ctx.sync();
        last.formulasR1C1[0].forEach((f, c) => { if (isF(f) && (c < c0 || c >= c0 + W)) b.getCell(had, c).getResizedRange(rows.length - had - 1, 0).formulasR1C1 = Array.from({ length: rows.length - had }, () => [f]); });
        await ctx.sync();
      }
      for (let s = 0; s < rows.length; s += CH) {
        const part = rows.slice(s, s + CH), r = b.getCell(s, c0).getResizedRange(part.length - 1, W - 1);
        r.values = part; b.getCell(s, c0 + 2).getResizedRange(part.length - 1, 0).numberFormat = part.map(() => ["m/d/yyyy"]);
        b.getCell(s, c0 + 3).getResizedRange(part.length - 1, 0).numberFormat = part.map(() => ["h:mm AM/PM"]);
        await ctx.sync();
      }
      if (had > rows.length) { b.getCell(rows.length, c0).getResizedRange(had - rows.length - 1, W - 1).clear(Excel.ClearApplyTo.contents); await ctx.sync(); }
      const ws = t.worksheet, u = ws.getRange("L1:M1"); u.values = [["Last Updated: ", stamp]]; ws.getRange("M1").numberFormat = [["m/d/yyyy h:mm AM/PM"]];
      await ctx.sync();
      return { n: rows.length, had };
    });
  }

  /* key/value rows (tblIvkConfig, tblIvkTemplates): change the ones given, add the rest */
  async function saveKeyed(table, pairs) {
    return Excel.run(async ctx => {
      const t = ctx.workbook.tables.getItem(table), b = t.getDataBodyRange(); b.load("values"); await ctx.sync();
      const add = [];
      for (const [k, v] of Object.entries(pairs)) {
        const i = b.values.findIndex(r => sameName(r[0], k));
        if (i >= 0) b.getCell(i, 1).values = [[String(v)]]; else add.push([k, String(v)]);
      }
      if (add.length) t.rows.add(null, add);
      await ctx.sync();
    });
  }
  /* tblIvkRoutes: one row per event, found by Event; columns by header */
  async function saveRoutes(routes) {
    return Excel.run(async ctx => {
      const t = ctx.workbook.tables.getItem("tblIvkRoutes"), h = t.getHeaderRowRange(), b = t.getDataBodyRange(); h.load("values"); b.load("values"); await ctx.sync();
      const cols = h.values[0].map(C.norm), cE = cols.indexOf("event"), add = [];
      for (const [ev, r] of Object.entries(routes)) {
        const row = cols.map(() => ""), i = b.values.findIndex(x => sameName(x[cE], ev)), base = i >= 0 ? b.values[i].map(String) : row;
        const put = (n, v) => { const c = cols.indexOf(n); if (c >= 0) base[c] = String(v); };
        put("event", ev); put("sendsto", r.to); put("flowurl", r.url); put("enabled", r.on ? "TRUE" : "FALSE");
        if (i >= 0) b.getRow(i).values = [base]; else add.push(base);
      }
      if (add.length) t.rows.add(null, add);
      await ctx.sync();
    });
  }
  async function appendLog(row) {
    return Excel.run(async ctx => {
      const t = ctx.workbook.tables.getItem("tblIvkLog"), b = t.getDataBodyRange(); b.load("values,rowCount"); await ctx.sync();
      const blank = b.values.length === 1 && b.values[0].every(v => v === "" || v == null);
      if (blank) b.getRow(0).values = [row.map(String)]; else t.rows.add(null, [row.map(String)]);
      await ctx.sync();
    });
  }
  /* Show the patient's row on the Master Patient Tracker (or a Consults row) */
  async function openRow(kind, mrn) {
    return Excel.run(async ctx => {
      const tm = await tableMap(ctx), name = tm[kind]; if (!name) return false;
      const t = ctx.workbook.tables.getItem(name), h = t.getHeaderRowRange(), b = t.getDataBodyRange(); h.load("values"); b.load("values"); await ctx.sync();
      const cols = h.values[0].map(C.norm), c = cols.indexOf(kind === "consults" ? "patientmrn" : "mrn"), i = C.pickRow(b.values, r => r[c], r => r[cols.indexOf("status")], mrn);
      t.worksheet.activate(); if (i >= 0) b.getRow(i).select();
      await ctx.sync();
      return i >= 0;
    });
  }
  /* ---------- Index planner ---------- */
  const isBlank = r => r.every(v => v === "" || v == null);
  /* rows by column name; the blank starter row from setup is used first */
  function addRows(t, cols, b, objs) {
    const rows = objs.map(o => cols.map(c => { const k = Object.keys(o).find(x => C.norm(x) === c); return k == null || o[k] == null ? "" : String(o[k]); }));
    if (!rows.length) return;
    if (b.values.length === 1 && isBlank(b.values[0])) { b.getRow(0).values = [rows[0]]; rows.shift(); }
    if (rows.length) t.rows.add(null, rows);
  }
  const openTable = async (ctx, name) => {
    const t = ctx.workbook.tables.getItemOrNullObject(name); t.load("isNullObject"); await ctx.sync();
    if (t.isNullObject) throw new Error("The Index Planner tables aren’t in this workbook yet. Use Update the workbook in Settings.");
    const h = t.getHeaderRowRange(), b = t.getDataBodyRange(); h.load("values"); b.load("values"); await ctx.sync();
    return { t, b, cols: h.values[0].map(C.norm) };
  };
  /* template slots: rows = [{ Provider, Day, Time, Type }]; a slot the provider already has on that day and time is skipped */
  async function addSlots(rows) {
    return Excel.run(async ctx => {
      const { t, b, cols } = await openTable(ctx, "tblIvkSlots"), at = n => cols.indexOf(n);
      const have = b.values.filter(r => !isBlank(r)).map(r => `${C.norm(C.provName(r[at("provider")]))}|${C.dayNum(r[at("day")])}|${C.toMin(r[at("time")])}`);
      const add = rows.filter(r => !have.includes(`${C.norm(C.provName(r.Provider))}|${C.dayNum(r.Day)}|${C.toMin(r.Time)}`))
        .map((r, i) => Object.assign({ Id: "s" + Date.now().toString(36) + i }, r));
      addRows(t, cols, b, add); await ctx.sync();
      return { added: add.length, skipped: rows.length - add.length };
    });
  }
  async function removeSlots(ids) {
    return Excel.run(async ctx => {
      const { t, b, cols } = await openTable(ctx, "tblIvkSlots"), c = cols.indexOf("id"), idx = [];
      b.values.forEach((r, i) => { if (ids.includes(String(r[c]))) idx.push(i); });
      /* a table keeps one row, so the last slot is cleared rather than deleted */
      for (const i of idx.sort((x, y) => y - x)) { if (b.values.length - idx.length === 0 && i === 0) b.getRow(0).clear(Excel.ClearApplyTo.contents); else t.rows.getItemAt(i).delete(); }
      await ctx.sync();
      return idx.length;
    });
  }
  /* one plan row per patient, by MRN */
  async function savePlan(mrn, patch) {
    return Excel.run(async ctx => {
      const { t, b, cols } = await openTable(ctx, "tblIvkPlans"), cM = cols.indexOf("mrn");
      const i = b.values.findIndex(r => C.mrnKey(r[cM]) === C.mrnKey(mrn));
      if (i < 0) { addRows(t, cols, b, [Object.assign({ MRN: mrn }, patch)]); await ctx.sync(); return true; }
      const row = b.values[i].map(v => v == null ? "" : String(v));
      for (const [k, v] of Object.entries(patch)) { const c = cols.indexOf(C.norm(k)); if (c >= 0) row[c] = v == null ? "" : String(v); }
      b.getRow(i).values = [row]; await ctx.sync(); return true;
    });
  }
  /* Holds: [{ MRN, Patient, Date (iso), Time, Provider, Treatment }]. The table is read again first: if someone else has just held one
     of these times, or this patient already has holds waiting, nothing is written and the error says which. */
  async function placeHolds(rows, by) {
    return Excel.run(async ctx => {
      const { t, b, cols } = await openTable(ctx, "tblIvkHolds"), H = C.holdRows([cols.map(c => c), ...b.values]);
      const live = H.filter(h => h.status !== "released");
      const mine = live.filter(h => h.status === "held" && rows.some(r => C.mrnKey(r.MRN) === h.mrn));
      if (mine.length) throw new Error(`${mine[0].code} already has ${mine.length} hold${mine.length === 1 ? "" : "s"} waiting. Book or release those first`);
      for (const r of rows) {
        const d = C.toSerial(r.Date), tm = C.toMin(r.Time), x = live.find(h => h.date === d && h.time === tm && C.sameProv(h.provider, r.Provider) && h.mrn !== C.mrnKey(r.MRN));
        if (x) throw new Error(`${x.by || "Someone"} just held ${C.fmtD(d)} at ${C.fmtT(tm)} for ${x.code}. The plan refreshes; check the times again`);
      }
      const at = new Date().toISOString(), base = Date.now().toString(36);
      addRows(t, cols, b, rows.map((r, i) => Object.assign({ Id: `h${base}${i}`, Status: "held", PlacedBy: by, PlacedAt: at }, r)));
      await ctx.sync();
      return { placed: rows.length, at };
    });
  }
  async function setHolds(ids, status, by) {
    return Excel.run(async ctx => {
      const { b, cols } = await openTable(ctx, "tblIvkHolds"), c = n => cols.indexOf(n), at = new Date().toISOString(); let n = 0;
      b.values.forEach((r, i) => { if (!ids.includes(String(r[c("id")]))) return; const row = r.map(v => v == null ? "" : String(v));
        row[c("status")] = status; if (c("doneby") >= 0) row[c("doneby")] = by || ""; if (c("doneat") >= 0) row[c("doneat")] = at; b.getRow(i).values = [row]; n++; });
      await ctx.sync();
      return n;
    });
  }
  /* Rewrites the Index Planner tab's grid as plain values: every IVK visit on the Appt List in its half hour, and holds as
     "HOLD CODE". The lookups into the other workbook go; the dates row, times, colours and the planning block stay. */
  async function fillPlanner(A, holds, listAt) {
    return Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItemOrNullObject(PLANNER); ws.load("isNullObject"); await ctx.sync();
      if (ws.isNullObject) return { skipped: "no Index Planner tab" };
      const r = ws.getRange(IP_RANGE); r.load("values"); await ctx.sync();
      const L = C.plannerLayout(r.values); if (!L) return { skipped: "no Date row on the Index Planner tab" };
      const g = C.plannerGrid(L, A, holds || [], listAt), r0 = L.rows[0][0], c0 = L.cols[0][0], c1 = L.cols[L.cols.length - 1][0];
      const block = ws.getRange(IP_RANGE).getCell(r0, c0).getResizedRange(L.rows.length - 1, c1 - c0); block.load("formulas"); await ctx.sync();
      const colAt = new Map(L.cols.map(([c], j) => [c, j]));
      block.formulas = block.formulas.map((row, i) => row.map((f, k) => colAt.has(c0 + k) ? g[i][colAt.get(c0 + k)] : f));
      await ctx.sync();
      return { cells: g.flat().filter(Boolean).length, from: L.cols[0][1], to: L.cols[L.cols.length - 1][1] };
    });
  }

  /* document settings: the chair map */
  const getDoc = k => { try { return Office.context.document.settings.get(k); } catch (e) { return null; } };
  /* pick up a chair map saved on another computer (the pane otherwise keeps the copy it opened with) */
  const refreshDoc = () => new Promise(res => { try { Office.context.document.settings.refreshAsync(() => res()); } catch (e) { res(); } });
  function setDoc(k, v) {
    return new Promise((res, rej) => {
      const s = Office.context.document.settings; if (v == null) s.remove(k); else s.set(k, v);
      s.saveAsync(r => r.status === Office.AsyncResultStatus.Succeeded ? res(true) : rej(new Error("Couldn’t save the setting to the workbook. Try again.")));
    });
  }

  return { DATA_SHEET, SHEETS, PLANNER, SPEC, load, setup, saveVisit, confirmProtocol, replaceAppts, saveKeyed, saveRoutes, appendLog, openRow,
    addSlots, removeSlots, savePlan, placeHolds, setHolds, fillPlanner, getDoc, setDoc, refreshDoc };
});
