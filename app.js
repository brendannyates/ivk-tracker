/* app.js: the IVK Tracker screens. Used at the Castro front desk; the RN never opens it and gets Teams posts instead.
   Today:     chair map (a Bay Map Builder file, else chair tiles), who's waiting and still to come, and one visit card
   Schedule:  the week by day and time, with open template slots and holds
   Index planner: the template by provider, and each new patient's index visits proposed from it, held, then booked in NextGen
   Patients:  everyone on the Master Patient Tracker, sorted by what they need (auth, booking, protocol)
   Consults:  the Consults tab, approved-not-scheduled first
   Settings:  your name, the NextGen upload, Teams workflows and messages, chair map, timing, appearance, workbook setup
   The task pane reads and writes the workbook (excel.js). Full screen is an Office dialog floating over the Excel page; it holds a
   copy of what the pane read and asks the pane for every write (callPane). Teams posts go straight from either one. */
(() => {
  "use strict";
  const C = window.IvkCore, N = window.IvkNotify, X = window.IvkExcel, TM = window.IvkTemplate;
  const FULL = new URLSearchParams(location.search).has("full");
  const MAP_KEY = "ivkBayPlan", POLL_MS = 15000, TICK_MS = 20000, CHUNK = 8000, CALL_MS = 120000;
  const BUILDER = "https://brendannyates.github.io/bay-map-builder/";
  const $ = s => document.querySelector(s);
  const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const LS = { get: (k, d) => { try { const v = localStorage.getItem(k); return v == null ? d : v; } catch (e) { return d; } }, set: (k, v) => { try { localStorage.setItem(k, v); } catch (e) { /* private mode */ } } };
  const nowMin = () => { const d = new Date(); return d.getHours() * 60 + d.getMinutes(); };
  const todayS = () => C.serialOf(new Date());
  const hm = () => C.fmtT(nowMin());
  const plural = (n, a, b) => `${n} ${n === 1 ? a : b || a + "s"}`;

  const S = { view: "today", raw: null, sig: "", live: false, err: "", sel: null, pSel: null, chip: "need", q: "", wk: 0, ipView: "plan", ipSel: null, ipPins: {}, ipAlt: null, ipWk: 0, tg: 0, tsel: null, te: null,
    modal: null, m: {}, loading: false, name: LS.get("ivkName", ""), theme: LS.get("ivkTheme", "light") };
  document.documentElement.dataset.theme = S.theme;

  /* ---------- derived model ---------- */
  const kv = vals => { const o = {}; (vals || []).slice(1).forEach(r => { const k = C.txt(r[0]); if (k) o[k] = r[1]; }); return o; };
  function derive() {
    const r = S.raw || {}, T = todayS(), cfg = Object.assign({}, C.DEFAULTS);
    for (const [k, v] of Object.entries(kv(r.tblIvkConfig))) if (C.txt(v) !== "") cfg[k] = typeof C.DEFAULTS[k] === "number" && C.num(v) != null ? C.num(v) : C.txt(v);
    S.cfg = cfg;
    S.tpl = Object.assign({}, C.TEMPLATES, C.START_TPL); for (const [k, v] of Object.entries(kv(r.tblIvkTemplates))) if (C.txt(v)) S.tpl[k] = C.txt(v);
    S.routes = {}; C.byHeader(r.tblIvkRoutes).rows.forEach(x => { const ev = C.txt(x.get("event")).toLowerCase(); if (ev) S.routes[ev] = { to: C.txt(x.get("sendsto")), url: C.txt(x.get("flowurl")), on: N.truthy(x.get("enabled")) }; });
    S.log = C.byHeader(r.tblIvkLog).rows.map(x => ({ ev: C.txt(x.get("event")), code: C.txt(x.get("patient")), chair: C.txt(x.get("chair")), msg: C.txt(x.get("message")), at: C.txt(x.get("sentat")), by: C.txt(x.get("sentby")), result: C.txt(x.get("result")) })).filter(l => l.ev);
    S.P = C.patients(r.mpt); S.PM = new Map(S.P.map(p => [p.mrn, p]));
    S.A = C.appts(r.appt);
    /* a tracker row whose Patient Name isn't a code takes the code from the Appt List's Name column, never the name itself */
    const codeOf = new Map(S.A.filter(a => a.code).map(a => [a.mrn, a.code]));
    for (const p of S.P) if (!p.code) p.code = codeOf.get(p.mrn) || "(no code)"; S.V = C.visitRows(r.visits); S.CS = C.consults(r.consults, T);
    /* and an Appt List row with a blank Patient takes the tracker's code (the Schedule showed only the time for these) */
    for (const a of S.A) if (!a.code && S.PM.get(a.mrn)?.code && S.PM.get(a.mrn).code !== "(no code)") a.code = S.PM.get(a.mrn).code;
    let plan = null; try { plan = r.plan ? bpCheck(r.plan).plan || null : null; } catch (e) { plan = null; }
    S.plan = plan; S.planErr = r.plan && !plan ? "The saved chair map can’t be read. Load the map file again." : "";
    S.x = { today: T, cfg, appts: S.A, patients: S.P, visits: S.V, clinic: cfg.Clinic };
    S.list = C.today(S.x); S.chairs = C.chairList(plan);
    /* a protocol ask posted today counts as asked even if saving the time on the visit failed */
    for (const v of S.list) if (v.st.asked == null) { const l = sentToday("protocol", C.code(v)); if (l) v.st = Object.assign({}, v.st, { asked: C.toMin(timeOf(l.at)) }); }
    S.plans = C.planRows(r.tblIvkPlans); S.holds = C.holdRows(r.tblIvkHolds); S.others = C.otherRows(r.tblIvkOther);
    /* provider templates (the IVK Template tab) and their slot types; the clinic template's types until a Slot Types tab exists */
    S.types = (r.slotTypes && TM.typeRows(r.slotTypes).length) ? TM.typeRows(r.slotTypes) : TM.DEFAULT_TYPES.slice();
    S.grids = TM.parseGrids(X.TPL_SHEET, r.template || []);
    S.flags = C.flagRows(r.tblIvkFlags); for (const p of S.P) if (S.flags.has(p.mrn)) p.reqPost = S.flags.get(p.mrn);
    /* per patient (payment flags, Oakland, availability, waitlist) and the payment log; visits before the log started are never asked about */
    S.PT = C.ptRows(r.tblIvkPatient); S.pays = C.payRows(r.tblIvkPay); S.payOn = !!r.tables?.own?.tblIvkPay;
    const firstPay = S.pays.map(x => x.date).filter(d => d != null).sort((a, b) => a - b)[0];
    S.payStart = C.toSerial(cfg.PayStart) ?? firstPay ?? T;
    Object.assign(S.x, { PT: S.PT, pays: S.payOn ? S.pays : null, payStart: S.payStart });
    /* the start checklist (tblIvkStarts) and the Consults rows against the appointment list: a kept start takes them off the waitlist */
    S.ST = C.startRows(r.tblIvkStarts); S.stOn = !!r.tables?.own?.tblIvkStarts;
    C.consultStarts(S.CS, S.A, S.ST, cfg);
    S.WL = C.waitlist(S.CS, S.PT);
    /* transfers to another clinic (scheduled, confirmed at the last visit here) */
    S.XF = C.xferRows(r.tblIvkTransfers); S.xferOn = !!r.tables?.own?.tblIvkTransfers; S.x.XF = S.XF;
    /* IVK visits at the other clinics, from a wider report: whether a transfer or an Oakland hand-off is booked there.
       E stays null until one has been uploaded, so "nothing booked there" is never said without a report to say it from */
    S.elseOn = !!r.tables?.own?.tblIvkElsewhere; S.elseAt = C.toSerial(cfg.ElsewhereAt);
    S.E = C.elseRows(r.tblIvkElsewhere); S.x.E = S.elseOn && C.txt(cfg.ElsewhereAt) ? S.E : null;
    /* the sooner waitlist: patients booked on a date who'd take an earlier one */
    S.SO = C.soonerRows(r.tblIvkSooner); S.soonerOn = !!r.tables?.own?.tblIvkSooner;
    /* patient profiles: who pays the provider visit and the infusion, the latest Sohar benefits. Anyone on the appointment list in
       the next 4 weeks without one (not on the tracker, or no profile row) is prompted, treatment starts first. */
    S.PF = C.profileRows(r.tblIvkProfile); S.profOn = !!r.tables?.own?.tblIvkProfile;
    S.gaps = C.profileGaps(S.A, S.PM, S.PF, T, 28);
    S.x.noProfile = new Set();   // patients on the tracker have their profile in its columns
    /* every treatment start: a profile once it's booked, Case Management in NextGen on the day */
    S.SL = C.startList({ today: T, appts: S.A, PM: S.PM, PF: S.PF, ST: S.ST, CS: S.CS, visits: S.V, cfg }); S.x.SL = S.SL;
    /* a visit takes its post check when it's an index visit (a treatment start, an Induction patient) or the patient requires one */
    const needsPost = (mrn, ev) => { const p = S.PM.get(mrn); return /start/i.test(ev) || !!p?.reqPost || /induction/i.test(p?.status || ""); };
    /* provider time off: their template slots go on those days, or become the covering provider's */
    S.EX = C.exRows(r.tblIvkExceptions);
    const gridKeys = new Set(S.grids.map(g => g.key));
    S.X = Object.assign({ appts: S.A, holds: S.holds, others: S.others, cfg, chairs: C.chairCount(cfg, plan), listAt: r.apptUpdated, needsPost, hasGrid: k => gridKeys.has(k) }, C.withExceptions(TM.onDays(S.grids, S.types, +cfg.PostCheckMin), S.EX));
    const pv = new Map(), addPv = n => { const k = TM.provKey(n); if (n && k && !pv.has(k)) pv.set(k, n); };
    S.A.forEach(a => addPv(a.provider)); S.others.forEach(o => addPv(o.provider));
    S.providers = [...pv.values()].sort((a, b) => TM.provKey(a).localeCompare(TM.provKey(b)));
    S.events = [...new Set(S.A.map(a => a.event).concat(S.others.map(o => o.event)).filter(Boolean))].sort();
    S.freqs = (r.freqs && r.freqs.length ? r.freqs : ["2x per week", "1x per week", "every 2 weeks", "every 3 weeks", "every 4 weeks", "every 6 weeks", "every 2 months", "every 3 months", "PRN"]);
  }
  /* worked out once per workbook read (the long scans: eligible starts, demand) */
  const memo = (k, fn) => { if (!S.memoR || S.memoR.raw !== S.raw) S.memoR = { raw: S.raw, v: {} }; const v = S.memoR.v; if (!(k in v)) v[k] = fn(); return v[k]; };
  /* last Teams post of an event for a patient today (from tblIvkLog) */
  const sentToday = (ev, code) => { const iso = C.isoOf(todayS()); return S.log.filter(l => l.ev === ev && l.code === code && new Date(l.at).toString() !== "Invalid Date" && C.isoOf(C.serialOf(new Date(l.at))) === iso && !/^failed/.test(l.result)).pop() || null; };
  const lastSent = (ev, code) => S.log.filter(l => l.ev === ev && l.code === code && !/^failed/.test(l.result)).pop() || null;
  const timeOf = iso => { const d = new Date(iso); return isNaN(d) ? "" : C.fmtT(d.getHours() * 60 + d.getMinutes()); };
  const dateOfIso = iso => { const d = new Date(iso); return isNaN(d) ? "" : C.fmtD(C.serialOf(d)); };
  const teamsReady = ev => !!(S.routes[ev] && S.routes[ev].on && N.flowUrl(S.routes[ev].url) && N.truthy(S.cfg.Enabled));
  const routeTo = ev => S.routes[ev]?.to || (ev === "arrival" ? "the RN" : "the group chat");

  function floor() {
    const now = nowMin(), L = S.list;
    const chairs = S.chairs.map(id => Object.assign({ id }, C.chairState(id, L)));
    const by = ph => L.filter(v => v.ph === ph);
    const tx = by("tx").sort((a, b) => a.st.seat - b.st.seat), inn = by("in").sort((a, b) => a.st.arr - b.st.arr);
    /* a chair kept on a visit that isn't on the loaded map (the map changed under it) */
    const off = S.plan ? L.filter(v => (v.ph === "tx" || v.ph === "in") && v.st.chair && !S.chairs.includes(v.st.chair)) : [];
    const asks = L.filter(v => !["gone", "noshow", "done"].includes(v.ph) && v.ask.any && v.st.asked == null);
    const openBook = by("done").filter(v => v.st.booked === "later" && v.ask.book);
    return { now, chairs, tx, inn, off, coming: by("booked"), done: by("done"), noshow: by("noshow"), gone: by("gone"), asks, openBook, askNow: tx.filter(C.askDue) };
  }
  /* the visit to show when nothing is picked: an ask due in treatment, checked in, in treatment, next to arrive */
  function pickVisit(F) {
    return F.askNow[0] || F.inn[0] || F.tx[0] || F.coming.find(v => (v.a.time ?? 0) >= F.now - 30) || F.coming[0] || F.done[F.done.length - 1] || null;
  }
  const visitBy = key => S.list.find(v => v.key === key) || null;

  /* ---------- writes (pane does them; full screen asks the pane) ---------- */
  const Local = {
    saveVisit: (k, b, p, g) => X.saveVisit(k, b, p, g), confirmProtocol: (m, f, s) => X.confirmProtocol(m, f, s), replaceAppts: (r, s) => X.replaceAppts(r, s),
    saveKeyed: (t, p) => X.saveKeyed(t, p), saveRoutes: r => X.saveRoutes(r), appendLog: r => X.appendLog(r), openRow: (k, m) => X.openRow(k, m),
    setup: () => X.setup(), setDoc: (k, v) => X.setDoc(k, v),
    savePlan: (m, p) => X.savePlan(m, p), saveFlag: (m, p) => X.saveFlag(m, p), placeHolds: (r, by) => X.placeHolds(r, by), setHolds: (ids, st, by) => X.setHolds(ids, st, by),
    readRow: (k, m) => X.readRow(k, m), saveRow: (k, m, p) => X.saveRow(k, m, p), savePt: (m, p) => X.savePt(m, p), addPays: r => X.addPays(r), saveRanks: (o, by) => X.saveRanks(o, by), saveException: (i, p) => X.saveException(i, p),
    setupTemplate: () => X.setupTemplate(), addGrid: (o, t) => X.addGrid(o, t), setCells: (a, r, c, v, m, s) => X.setCells(a, r, c, v, m, s),
    saveTypes: (t, g) => X.saveTypes(t, g), replaceOther: r => X.replaceOther(r), replaceElse: r => X.replaceElse(r), saveSooner: (m, p) => X.saveSooner(m, p), saveProfile: (m, p) => X.saveProfile(m, p), saveXfer: (m, p) => X.saveXfer(m, p), addMpt: f => X.addMpt(f), saveStart: (k, p) => X.saveStart(k, p),
    /* the pane reads first so the tab gets every hold anyone has placed, not just this screen's copy */
    fillPlanner: async () => { await refresh(true); return X.fillPlanner(S.A, S.holds, S.raw.apptUpdated); }
  };
  const QUIET = ["openRow", "readRow"];   // reads: no refresh after them
  const Store = {};
  for (const k of Object.keys(Local)) Store[k] = FULL ? (...a) => callPane(k, a) : async (...a) => { const r = await Local[k](...a); if (!QUIET.includes(k)) await refresh(true); return r; };
  /* Excel's refusals in plain words: AccessDenied is a protected sheet or a workbook open for viewing only */
  const whyFailed = e => { const m = String(e?.message || e);
    return e?.code === "AccessDenied" || /cannot perform the requested operation|can't perform|could not perform/i.test(m) ? `Excel refused the change. Check the workbook is open for Editing (not Viewing) and the sheet isn't protected (Review › Protect), then try again. (${m})`
      : `${m}. Try again.`; };
  async function write(label, fn, done) {
    S.busy = label; render();
    try { const r = await fn(); if (done) toast(typeof done === "function" ? done(r) : done); return r || true; }
    catch (e) { console.error(e); toast(`Couldn’t save to the workbook (${label.replace(/…$/, "").toLowerCase()}): ${whyFailed(e)}`, true); return null; }
    finally { S.busy = null; render(); }
  }
  async function teams(ev, text, facts, code, chair) {
    const r = await N.send({ event: ev, text, facts, patient: code, chair, sentBy: S.name, clinic: S.cfg.Clinic }, S.routes, S.cfg.Enabled, row => Store.appendLog(row));
    return r;
  }
  const teamsWord = (r, who) => r.ok ? (r.unconfirmed ? `Posted to ${who}; the workflow didn’t confirm it, so check the chat.` : `Posted to ${who} in Teams.`) : r.reason;

  /* ---------- pane <-> full screen ---------- */
  const pop = { dlg: null }, rpc = { id: 0, wait: new Map() }, asm = {};
  const parts = s => { const n = Math.ceil(s.length / CHUNK) || 1, seq = Date.now() + Math.random(); return Array.from({ length: n }, (_, i) => JSON.stringify({ k: "part", seq, i, n, part: s.slice(i * CHUNK, (i + 1) * CHUNK) })); };
  function reassemble(raw, cb) {
    let m; try { m = JSON.parse(raw); } catch (e) { return; }
    if (m.k !== "part") return cb(m);
    const b = asm[m.seq] || (asm[m.seq] = { n: m.n, p: [] }); b.p[m.i] = m.part;
    if (b.p.filter(x => x != null).length === m.n) { delete asm[m.seq]; let o; try { o = JSON.parse(b.p.join("")); } catch (e) { return; } cb(o); }
  }
  function toChild(o) { if (!pop.dlg) return; for (const p of parts(JSON.stringify(o))) try { pop.dlg.messageChild(p); } catch (e) { /* window closed */ } }
  function toPane(o) { for (const p of parts(JSON.stringify(o))) Office.context.ui.messageParent(p); }
  /* the pane pings "wait" every 15 s while it works, so a long write (a big upload) doesn't time out; silence for CALL_MS does */
  function callPane(fn, args) {
    return new Promise((res, rej) => {
      const id = ++rpc.id, lost = () => { if (rpc.wait.delete(id)) rej(new Error("full screen lost touch with the task pane. Keep the pane open, then check the workbook before trying again")); };
      const w = { t: setTimeout(lost, CALL_MS), alive() { clearTimeout(w.t); w.t = setTimeout(lost, CALL_MS); }, res: v => { clearTimeout(w.t); res(v); }, rej: e => { clearTimeout(w.t); rej(e); } };
      rpc.wait.set(id, w);
      try { toPane({ k: "call", id, fn, args }); } catch (e) { clearTimeout(w.t); rpc.wait.delete(id); rej(e); }
    });
  }
  const PANE = Object.fromEntries(Object.keys(Local).map(k => [k, async (...a) => { const r = await Local[k](...a); if (!QUIET.includes(k)) await refresh(true); return r; }]));
  function onChild(arg) {
    reassemble(arg.message, async m => {
      if (window.DevUpdates && DevUpdates.receive(m)) return;   // a Dev Update dismissed in full screen
      if (m.k === "ready") return toChild({ k: "snap", raw: S.raw });
      if (m.k === "close") return closeFull();
      if (m.k === "refresh") { await refresh(true); return; }
      if (m.k !== "call" || !PANE[m.fn]) return;
      let ok = true, err = "", result = null;
      const ping = setInterval(() => toChild({ k: "wait", id: m.id }), 15000);
      try { result = await PANE[m.fn](...(m.args || [])); } catch (e) { ok = false; err = e.message || String(e); }
      finally { clearInterval(ping); }
      toChild({ k: "ack", id: m.id, ok, err, result: result === undefined ? null : result });
    });
  }
  function onPane(m) {
    if (window.DevUpdates && DevUpdates.receive(m)) return;
    if (m.k === "snap") { S.raw = m.raw; S.live = true; derive(); if (!S.modal && !typing()) render(); }
    else if (m.k === "wait") rpc.wait.get(m.id)?.alive();
    else if (m.k === "ack") { const w = rpc.wait.get(m.id); if (!w) return; rpc.wait.delete(m.id); if (m.ok) w.res(m.result); else w.rej(new Error(m.err || "the task pane couldn’t do that")); }
  }
  function openFull() {
    if (FULL) { try { toPane({ k: "close" }); } catch (e) { /* pane gone */ } return; }
    if (pop.dlg) return toast("Full screen is already open.");
    if (!Office.context.requirements.isSetSupported("DialogApi", "1.2")) return toast("This version of Excel can’t open full screen.", true);
    const url = new URL(location.href); url.search = "?full=1"; url.hash = "";
    /* the Dev Updates already dismissed here go along, so full screen doesn't show them again */
    Office.context.ui.displayDialogAsync(window.DevUpdates ? DevUpdates.dialogUrl(url.href) : url.href, { height: 95, width: 95, displayInIframe: true }, res => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) return toast(res.error.code === 12007 ? "Close the other add-in window first." : "Full screen couldn’t open. Try again.", true);
      pop.dlg = res.value; render(); if (window.DevUpdates) DevUpdates.setSend(m => toChild(m));
      pop.dlg.addEventHandler(Office.EventType.DialogMessageReceived, onChild);
      pop.dlg.addEventHandler(Office.EventType.DialogEventReceived, () => { pop.dlg = null; if (window.DevUpdates) DevUpdates.setSend(null); render(); });
    });
  }
  function closeFull() { if (pop.dlg) { try { pop.dlg.close(); } catch (e) { /* closed */ } pop.dlg = null; } if (window.DevUpdates) DevUpdates.setSend(null); render(); }

  /* ---------- reading ---------- */
  const hash = s => { let h = 0; for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) | 0; return s.length + ":" + h; };
  /* One read at a time. A refresh asked for while one is running (a write landing during the 15 s poll) waits for it and then
     reads again, so nobody is left looking at the rows from before the write. */
  let readQ = Promise.resolve(), queued = null;
  function refresh(force) {
    if (FULL) { try { toPane({ k: "refresh" }); } catch (e) { /* pane gone */ } return Promise.resolve(); }
    if (queued) { queued.force = queued.force || !!force; return queued.p; }
    const q = queued = { force: !!force };
    q.p = readQ = readQ.then(() => { queued = null; return readOnce(q.force); });
    return q.p;
  }
  async function readOnce(force) {
    try {
      await X.refreshDoc();
      const raw = await X.load(S.raw ? null : bootStep); raw.plan = X.getDoc(MAP_KEY) || null;
      const sig = hash(JSON.stringify(raw));
      S.live = true; S.err = "";
      if (force || sig !== S.sig) { const first = !S.raw; S.sig = sig; S.raw = raw; derive(); if (first) bootDone(); toChild({ k: "snap", raw }); if (!S.modal && !typing()) render(); }
    } catch (e) { console.error(e); S.err = "Couldn’t read the workbook. Check you’re still signed in; the tracker tries again in a few seconds."; render(); }
  }

  /* the loading screen on open: each step ticks as its read finishes (excel.js load), then the chair map is drawn */
  const BOOT = ["Connecting to the workbook", "Reading Master Patient Tracker", "Reading Appt List", "Reading IVK Tracker Data", "Drawing the chair map"];
  const rowsIn = v => v && v.length > 1 ? v.length - 1 : 0;
  function bootStep(i, o) {
    const B = S.boot || (S.boot = { i: 0, m: [] });
    B.m[i] = i === 0 ? "Connected" : i === 1 ? (o.mpt ? plural(rowsIn(o.mpt), "patient") : "No Master Patient Tracker table") : i === 2 ? (o.appt ? plural(rowsIn(o.appt), "visit") + " on the list" : "No Appt List table")
      : "Check-ins, holds, payments";
    B.i = i + 1; render();
  }
  function bootDone() { const B = S.boot || (S.boot = { i: 0, m: [] }); B.m[4] = S.plan ? plural(S.chairs.length, "chair") : "No map loaded"; B.i = 5; }
  function bootHTML() {
    const B = S.boot || { i: 0, m: [] }, steps = FULL ? ["Getting the workbook from the task pane"] : BOOT;
    return `<div class="boot" role="status" aria-live="polite"><div class="bootin"><div class="spinner" aria-hidden="true"></div><h1>IVK Tracker</h1><p class="muted">${S.err ? esc(S.err) : "Reading the workbook."}</p>
      <ol class="steps">${steps.map((t, i) => `<li class="step ${B.i > i ? "done" : B.i === i ? "now" : "next"}"><span class="k">${B.i > i ? "✓" : ""}</span><span><span class="t">${t}</span><span class="m">${esc(B.m[i] || "")}</span></span><span></span></li>`).join("")}</ol></div></div>`;
  }

  /* ---------- frame ---------- */
  const NAV = [["today", "Today"], ["schedule", "Schedule"], ["access", "Access"], ["patients", "Patients"], ["avail", "Availability"]];
  /* Access holds four pages, picked with the chips at its top: demand by protocol, the waitlist, the index planner, the templates */
  const ACCESS = [["access", "Demand"], ["consults", "Waitlist"], ["index", "Index planner"], ["template", "Templates"]];
  const inAccess = v => ACCESS.some(a => a[0] === v);
  const navOn = k => k === "access" ? inAccess(S.view) : S.view === k;
  function badges() {
    const F = floor(), b = {};
    /* Today counts the visits in treatment whose providers still need asking */
    const cmN = S.SL.filter(s => s.cmDue).length;
    if (F.askNow.length || cmN) b.today = { n: F.askNow.length + cmN };
    /* Patients counts only what blocks the next infusion; the amber needs live under the Needs attention chip */
    const pn = S.P.filter(p => C.needs(p, S.x).some(y => y.bad)).length; if (pn) b.patients = { n: pn, bad: true };
    const ca = S.CS.filter(c => c.state === "act" || c.state === "cm").length; if (ca) b.consults = { n: ca };
    if (S.gaps.length) b.patients = b.patients ? { n: b.patients.n + S.gaps.filter(g => !g.onMpt).length, bad: true } : { n: S.gaps.length };
    /* holds waiting to be booked; red once one is past HoldMin without being confirmed */
    /* and booked holds the uploaded report hasn't confirmed; red when a newer report doesn't have them */
    const hs = liveHolds(), U = unconfirmed(), np = new Set(hs.map(h => h.mrn).concat(U.map(h => h.mrn))).size;
    if (np) b.index = { n: np, bad: hs.some(h => h.state === "late") || U.some(h => h.conf === "missing") };
    if (b.index || b.consults) b.access = { n: (b.index?.n || 0) + (b.consults?.n || 0), bad: !!b.index?.bad };
    const set = (S.raw?.missing || []).length + (teamsReady("arrival") && teamsReady("protocol") ? 0 : 1); if (set) b.settings = { n: set };
    return b;
  }
  function render() {
    const page = $(".page"), top = page ? page.scrollTop : 0;
    if (!S.raw) { $("#app").innerHTML = bootHTML(); return; }
    const B = badges(), item = ([k, label]) => `<button class="item" data-act="nav" data-v="${k}"${navOn(k) ? ' aria-current="page"' : ""}>${label}${B[k] ? `<span class="count${B[k].bad ? " bad" : ""}" aria-label="${B[k].n} need attention">${B[k].n}</span>` : ""}</button>`;
    /* the patient view takes the page while it's open (as the TMS Tracker's does); the rail closes it */
    const pv = S.pv && (S.PM.get(S.pv) || null);
    if (S.pv && !pv) S.pv = null;
    const body = pv ? ptViewHTML(pv) : { today: todayHTML, schedule: scheduleHTML, access: demandHTML, index: indexHTML, template: templatePageHTML, patients: patientsHTML, consults: consultsHTML, avail: availPageHTML, settings: settingsHTML }[S.view]();
    /* Settings sits at the bottom, above full screen (in full screen) or on its own (in the pane, where full screen is with
       refresh and dark mode at the top right). While a refresh reads the workbook the page dims and can't be pressed. */
    $("#app").innerHTML = `<div class="app">
      <nav class="rail" aria-label="IVK Tracker">
        <div class="brand">IVK Tracker<small>${esc(S.cfg.Clinic)} · IV ketamine</small></div>
        ${NAV.map(item).join("")}
        <div class="spacer"></div>
        ${item(["settings", "Settings"])}
        ${FULL ? `<button class="item" data-act="full">Exit full screen</button>` : ""}
        <div class="conn${S.err ? " off" : ""}"><span class="cdot"></span>${S.err ? "Not connected" : S.busy ? esc(S.busy) : FULL ? "Saving through the task pane" : esc(C.txt(S.raw.wbName).replace(/\.xls[xmb]?$/i, "") || `${S.cfg.Clinic}: connected to the workbook`)}</div>
      </nav>
      <main class="page" id="main">${S.reading ? `<div class="pbar" aria-hidden="true"><span></span></div>` : ""}${S.err ? `<div class="msg bad">${esc(S.err)}</div>` : ""}<div class="pagein"${S.reading ? " inert" : ""}>${body}</div></main></div>
      ${S.modal ? `<div class="scrim" data-act="scrim"><div class="modal" role="dialog" aria-modal="true" aria-labelledby="mh">${MODALS[S.modal]()}</div></div>` : ""}`;
    const p2 = $(".page"); if (p2) p2.scrollTop = top;
    if (S.modal) { const f = $(".modal [autofocus]") || $(".modal button.primary"); if (f && !$(".modal").contains(document.activeElement)) f.focus(); }
  }
  let toastT = null;
  function toast(msg, bad) {
    const t = $("#toast"); if (!t) return; t.textContent = msg; t.classList.toggle("bad", !!bad); t.classList.add("show");
    clearTimeout(toastT); toastT = setTimeout(() => t.classList.remove("show"), bad ? 6000 : 3600);
  }
  /* top right of every screen: read the workbook again (circular arrow), and light or dark (the moon switches to dark, the sun back) */
  const ICON = {
    refresh: `<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 12a8 8 0 1 1-2.34-5.66"/><path d="M20 4v5h-5"/></svg>`,
    moon: `<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 14.5A8 8 0 0 1 9.5 4a8 8 0 1 0 10.5 10.5z"/></svg>`,
    full: `<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5"/></svg>`,
    sun: `<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/></svg>`
  };
  const tools = () => { const dark = S.theme === "dark", busy = !!S.reading;
    const full = FULL ? "" : `<button class="ibtn" data-act="full" title="${pop.dlg ? "Full screen is open" : "Full screen"}" aria-label="${pop.dlg ? "Full screen is open" : "Full screen"}">${ICON.full}</button>`;
    return `<div class="tools">${busy ? `<span class="busyl">Reading the workbook…</span>` : ""}${full}<button class="ibtn${busy ? " spin" : ""}" data-act="reload" title="Refresh: read the workbook again" aria-label="Refresh from the workbook"${busy ? " disabled" : ""}>${ICON.refresh}</button>
      <button class="ibtn" data-act="theme" title="${dark ? "Light mode" : "Dark mode"}" aria-label="${dark ? "Switch to light mode" : "Switch to dark mode"}" aria-pressed="${dark}">${dark ? ICON.sun : ICON.moon}</button></div>`; };
  const head = (title, sub, clock) => `<div class="top"><div><h1>${title}</h1>${sub ? `<p class="sub">${sub}</p>` : ""}</div><div class="topr">${tools()}${clock ? `<div class="clock"><div class="t">${hm()}</div><div class="d">${esc(C.fmtDay(todayS()))}</div></div>` : ""}</div></div>`;
  const TABLE_WHAT = { tblIvkPlans: "index plans", tblIvkHolds: "holds", tblIvkOther: "the providers’ other visits", tblIvkFlags: "post-check flags", tblIvkPatient: "payment flags, Oakland, availability and the waitlist order", tblIvkPay: "payments and balance reminders", tblIvkExceptions: "provider time off and coverage", tblIvkProfile: "patient profiles (who pays, Sohar benefits)", tblIvkTransfers: "transfers to another clinic", tblIvkElsewhere: "visits at the other clinics (transfer and Oakland hand-offs)", tblIvkSooner: "the sooner waitlist", tblIvkStarts: "the treatment start checklist" };
  function setupBanner() {
    const miss = S.raw.missing || [], noTabs = ["mpt", "appt"].filter(k => !S.raw.tables?.[k]);
    if (noTabs.length) return `<div class="msg warn"><b>This workbook has no ${noTabs.map(k => k === "mpt" ? "Master Patient Tracker" : "Appt List").join(" or ")} table.</b> Open the IVK Patient Tracker workbook, then the add-in.</div>`;
    const NEW = ["tblIvkPlans", "tblIvkHolds", "tblIvkOther", "tblIvkFlags", "tblIvkPatient", "tblIvkPay", "tblIvkExceptions", "tblIvkProfile", "tblIvkTransfers", "tblIvkElsewhere", "tblIvkSooner", "tblIvkStarts"];
    if (miss.length && miss.every(n => NEW.includes(n))) return `<div class="msg warn"><b>This version needs ${miss.length === 1 ? "one more table" : `${miss.length} more tables`}.</b> Updating adds them to the hidden ${esc(X.DATA_SHEET)} tab (${miss.map(n => TABLE_WHAT[n]).join(", ")}); nothing else changes. <div class="actions"><button class="btn primary inl" data-act="setup">Update the workbook</button></div></div>`;
    if (miss.length) return `<div class="msg warn"><b>This workbook isn’t set up for the add-in yet.</b> Setting up adds a hidden IVK Tracker Data tab for check-ins and Teams settings; nothing else changes. <div class="actions"><button class="btn primary inl" data-act="setup">Set up this workbook</button></div></div>`;
    return "";
  }

  /* ---------- Today ---------- */
  function todayHTML() {
    const F = floor();
    const sub = S.list.length ? `${F.tx.length} in treatment, ${F.inn.length} checked in, ${F.coming.length} still to come.` : "No IVK visits on the list today.";
    if (!S.sel || !visitBy(S.sel)) { const v = pickVisit(F); S.sel = v ? v.key : null; }
    return `${setupBanner()}${apptStale()}${head("Today", sub, true)}
      <div class="today"><div class="tmain">${S.plan ? chairsHTML(F) : planErrHTML()}
        ${listHTML("In treatment", `${F.tx.length} now`, F.tx, "Nobody is in treatment.")}
        ${listHTML("Checked in", `${F.inn.length} today`, F.inn, "Nobody is checked in.")}
        ${listHTML("Still to come", F.coming.length ? `${F.coming.length} more today` : "", F.coming, "Nobody else is booked today.")}
        ${startsHTML()}${doneHTML(F)}${gapsHTML(6)}</div>
      <div class="tside">${visitCardHTML(visitBy(S.sel), F)}</div></div>`;
  }
  const planErrHTML = () => S.planErr ? `<div class="msg warn">${esc(S.planErr)}</div>` : "";
  /* patients on the appointment list without a profile (treatment starts first): the tech makes one before they come */
  /* the start checklist: a newly booked start needs a profile; on the day, Case Management in NextGen goes to CmStage.
     Today's starts always show; the rest only while something is still to do. */
  const cmStage = () => C.txt(S.cfg.CmStage) || C.DEFAULTS.CmStage;
  function startsHTML() {
    const L = S.SL; if (!L.length) return "";
    const box = (on, attrs, label, small, dis) => `<label class="opt"><input type="checkbox" ${attrs}${on ? " checked" : ""}${dis ? " disabled" : ""}> ${label}${small ? `<small>${small}</small>` : ""}</label>`;
    const rows = L.map(s => { const m = esc(s.mrn), k = esc(s.key), when = s.today ? "today" : dshort(s.date), cmLater = s.date > todayS();
      const prof = box(s.profDone, `data-prof="${m}"`, "Profile created", s.profDone ? (S.PM.has(s.mrn) ? "On the Master Patient Tracker" : "Profile saved") : "Opens the profile form: who pays, Sohar benefits", s.profDone);
      const cmWhy = s.cmDone ? (s.cmAt != null ? `Marked ${C.fmtD(s.cmAt)}` : `Case Management shows ${esc(s.c?.stage || cmStage())}`)
        : cmLater ? "On the day of the start" : s.noshow ? "No-show: nothing to update" : s.today ? "Once the start is kept" : `Start kept ${C.fmtD(s.date)}`;
      const cm = box(s.cmDone, `data-cm="${k}" data-m="${m}"`, `Case Management updated in NextGen to “${esc(cmStage())}”`, cmWhy,
        cmLater || (!s.cmDone && s.noshow) || (s.cmDone && s.cmAt == null) || !(S.stOn || s.c));
      return `<div class="nrow startck"><b>${esc(s.code || "—")}</b>
        <span class="i"><span class="ktab st"></span>Treatment start ${when} ${C.fmtT(s.a.time)} · ${esc(C.provName(s.a.provider) || "")} · MRN ${m}${s.kept ? " · kept" : ""}</span>
        <div class="ck">${prof}${cm}</div></div>`; }).join("");
    const open = L.filter(s => s.open).length;
    return `<section class="card list"><div class="lh"><h2>Treatment starts</h2><span>${open ? `${plural(open, "start")} with something to do` : "All done"}</span></div>${rows}
      ${S.stOn ? "" : `<p class="note warn">Ticking Case Management needs the newer table. Use <b>Update the workbook</b> in Settings.</p>`}</section>`;
  }
  function gapsHTML(max) {
    /* new patients with a treatment start are on the start checklist above */
    const G = S.gaps.filter(g => !S.SL.some(s => s.mrn === g.mrn && s.profDue)); if (!G.length) return "";
    const rows = G.slice(0, max).map(g => `<div class="nrow"><b>${esc(g.code || "—")}</b>
      <span class="i">${g.start ? `<span class="ktab st"></span>Treatment start` : "Visit"} ${dshort(g.first.date)} ${C.fmtT(g.first.time)} · ${esc(C.provName(g.first.provider) || "")} · MRN ${esc(g.mrn)}</span>
      <button class="btn line sm" data-act="profile" data-m="${esc(g.mrn)}">Create profile</button></div>`).join("");
    return `<section class="card list"><div class="lh"><h2>New patients to add</h2><span>On the appointment list in the next 4 weeks, not on the tracker</span></div>${rows}
      ${G.length > max ? `<p class="note">And ${plural(G.length - max, "more")} in the next 4 weeks. <button class="btn mute sm" data-act="nav" data-v="patients">See them on Patients</button></p>` : ""}
      ${S.profOn ? "" : `<p class="note warn">Profiles need the newer table. Use <b>Update the workbook</b> in Settings.</p>`}</section>`;
  }
  function apptStale() {
    const u = S.raw.apptUpdated, T = todayS();
    if (u != null && Math.floor(u) >= T - 1) return "";
    return `<div class="msg warn"><b>${u == null ? "No record of when the appointment list was last updated." : `The appointment list was last updated ${C.fmtD(Math.floor(u))}.`}</b> Upload a fresh NextGen export so today’s list is right. <button class="btn ghost sm" data-act="nav" data-v="settings">Upload in Settings</button></div>`;
  }
  const chairTxt = id => `Chair ${id}`;
  /* the chair map: only once a Bay Map Builder file is loaded (Settings › Chair map) */
  function chairsHTML(F) {
    const P = S.plan, used = F.chairs.filter(c => c.v).length, W = P.w, H = P.h;
    const label = c => `${chairTxt(c.id)}, ${c.v ? `${C.code(c.v)}, ${C.chairStatus(c)}` : "open"}`;
    const cur = c => !!(c.v && c.v.key === S.sel);
    const chairOf = id => F.chairs.find(c => c.id === id) || { id, state: "open", v: null };
    const off = F.off.length ? `<p class="boff">In a chair that isn’t on the map: ${F.off.map(v => `${esc(C.code(v))} (${esc(chairTxt(v.st.chair))})`).join(", ")}.</p>` : "";
    const bays = P.bays.map(b => { const c = chairOf(b.id); return Object.assign({}, b, { c, st: c.state, title: c.v ? C.code(c.v) : "Open", s: C.chairStatus(c) }); });
    const lit = b => !!b.c.v, usedB = id => lit(bays.find(b => b.id === id) || { c: {} });
    const pos = b => `left:${bpPct(b.cx - b.r, W)};top:${bpPct(b.cy - b.r, H)};width:${bpPct(2 * b.r, W)}`;
    const tag = b => `<div class="btag" data-state="${b.st}" style="left:${bpPct(b.tag[0], W)};top:${bpPct(b.tag[1], H)}" aria-hidden="true"><div class="n"><span class="num">${esc(b.id)}</span>${esc(b.title)}</div>${b.s ? `<div class="s"><span class="dot"></span><span class="tx">${esc(b.s)}</span></div>` : ""}</div>`;
    const bay = b => `<button class="bay" data-act="bSel" data-b="${esc(b.id)}" data-state="${b.st}" aria-current="${cur(b.c)}" aria-label="${esc(label(b.c))}" style="${pos(b)}"></button>`;
    return `<section class="card mapcard"><div class="hd"><h2>Chairs</h2><span>${used} of ${F.chairs.length} in use · tap a chair for its visit</span></div>
      <div class="bplan" style="aspect-ratio:${W}/${H}">
      <svg viewBox="0 0 ${W} ${H}" aria-hidden="true">${bpUnder(P, usedB)}</svg>
      ${bays.map(bay).join("")}
      <svg viewBox="0 0 ${W} ${H}" aria-hidden="true">${bpOver(P, bays, lit)}</svg>
      ${bpRooms(P)}${bays.map(tag).join("")}</div>
      <ul class="bkey" aria-label="Key"><li><span class="bsw" data-k="open"></span>Open</li><li><span class="bsw" data-k="track"></span>Checked in or in treatment</li><li><span class="bsw" data-k="soon"></span>Ask the providers</li><li><span class="lamp"></span>Lantern on: someone is in the chair</li></ul>${off}</section>`;
  }
  /* what the row says it needs, in amber */
  function rowFlag(v) {
    if (!v.p) return "Not on the Master Patient Tracker";
    if (C.askDue(v)) return "Ask providers";
    const ish = C.indexShort(v.p, S.A, v.a, S.cfg); if (ish) return `Index not fully booked (${ish.booked} of ${ish.need})`;
    if (v.ph === "booked" && v.st.asked == null) return v.ask.book ? "Book out" : v.ask.proto ? "Confirm protocol" : "";
    return "";
  }
  /* treatment start (blue tab), index (lavender), follow-up (grey) and "Index 3 of 6" for an Induction patient */
  const kindOf = v => visitKind(v.a) || "fu";
  const KIND_TXT = { st: "Treatment start", ix: "Index", fu: "Follow-up" };
  function kindTxt(v) {
    const ix = v.p ? C.indexState(v.p, S.A, todayS(), S.cfg) : null;
    return ix && !ix.complete ? `Index ${Math.min(ix.N, ix.done + 1)} of ${ix.N}` : KIND_TXT[kindOf(v)];
  }
  const hasPc = v => !!S.X.needsPost(v.a.mrn, v.a.event || "");
  function rowHTML(v) {
    const w = v.ph === "tx" ? `Since ${C.fmtT(v.st.seat)}` : v.ph === "in" ? `At ${C.fmtT(v.st.arr)}` : v.ph === "done" ? `Out ${C.fmtT(v.st.out)}` : v.ph === "noshow" ? "No-show" : v.ph === "gone" ? "Not coming" : inTxt(v.a.time);
    const info = [kindTxt(v), C.provName(v.a.provider), hasPc(v) ? "PC" : "", v.p && ptOf(v.p.mrn)?.oak ? "Oakland" : ""].filter(Boolean).join(" · ");
    const flag = v.ph === "done" ? (v.st.booked === "later" && v.ask.book ? "Next visit not booked" : "") : rowFlag(v);
    const right = S.plan && v.st.chair && (v.ph === "tx" || v.ph === "in") ? chairTxt(v.st.chair) : "";
    return `<button class="row${v.ph === "done" || v.ph === "gone" || v.ph === "noshow" ? " past" : ""}" data-act="sel" data-k="${esc(v.key)}" aria-pressed="${S.sel === v.key}">
      <span class="tm">${C.fmtT(v.a.time).replace(/ ?[AP]M$/, "")}</span><span class="w">${w}</span>
      <span class="i"><span class="ktab ${kindOf(v)}"></span><b>${esc(C.code(v))}</b>${mrnTag(v.a.mrn)}<span class="m">${esc(info)}</span>${flag ? `<span class="f">${esc(flag)}</span>` : ""}</span>
      <span class="r">${esc(right)}</span></button>`;
  }
  const inTxt = t => { if (t == null) return ""; const d = t - nowMin(); return d <= 0 ? "Due now" : d >= 60 ? `In ${Math.floor(d / 60)} h${d % 60 ? ` ${d % 60} min` : ""}` : `In ${d} min`; };
  const listHTML = (title, count, L, empty) => `<section class="card list"><div class="lh"><h2>${title}</h2><span>${count}</span></div>${L.map(rowHTML).join("") || `<p class="empty">${empty}</p>`}</section>`;
  const mrnTag = m => m ? `<small class="pm">${esc(m)}</small>` : "";
  /* in a window title */
  const mrnH = m => m ? ` <span class="mrn">MRN ${esc(m)}</span>` : "";
  function doneHTML(F) {
    const L = [...F.done, ...F.noshow, ...F.gone];
    if (!L.length) return "";
    return `<section class="card list"><div class="lh"><h2>Checked out</h2><span>${F.done.length} today${F.openBook.length ? `, ${F.openBook.length} still to book` : ""}</span></div>${L.map(rowHTML).join("")}</section>`;
  }

  /* the visit card: status, tags, lead, steps, facts and the one next step */
  function visitState(v) {
    const p = v.p, pc = hasPc(v), post = v.st.seat != null && pc ? v.st.seat + +S.cfg.PostCheckMin : null, prov = C.provName(v.a.provider) || "the provider";
    const why = !p ? "" : v.ask.proto ? `The protocol hasn’t been confirmed since ${p.protoAt != null ? C.fmtD(p.protoAt) : "it was set"}.` : v.ask.thru != null ? `Booked through ${C.fmtD(v.ask.thru)}.` : "Nothing is booked after today.";
    if (v.ph === "booked") { const late = v.a.time != null && nowMin() - v.a.time > 15;
      return { lead: late ? `${C.fmtT(v.a.time)} visit, not checked in yet` : inTxt(v.a.time) || `Booked at ${C.fmtT(v.a.time)}`, body: `Expected at ${C.fmtT(v.a.time)}.${p && v.ask.any ? " The group chat gets the ask at check-in." : ""}`, pri: ["checkin", "Check in"] }; }
    if (v.ph === "in") return { lead: "Checked in", body: `Checked in at ${C.fmtT(v.st.arr)} for a ${C.fmtT(v.a.time)} visit.`, pri: ["start", "Start treatment"] };
    if (v.ph === "tx") {
      if (C.askDue(v)) return { warn: true, lead: "Ask the group chat", body: `${why} Ask now so the provider can answer before discharge.`, pri: ["ask", "Ask the group chat"] };
      return { lead: "In treatment", body: `Since ${C.fmtT(v.st.seat)}.${post != null ? ` Requires a post-check at ${C.fmtT(post)} with ${prov}.` : " Nothing needed from the desk until check-out."}`, pri: ["checkout", "Check out"] };
    }
    if (v.ph === "done") { const open = v.st.booked === "later" && v.ask.book;
      return { warn: open, lead: `Checked out at ${C.fmtT(v.st.out)}`, body: open ? "Next visit not booked yet." : "Nothing left for this visit.", pri: open ? ["markBooked", "Mark next visit booked"] : v.p ? ["printSched", "Print schedule"] : null }; }
    if (v.ph === "noshow") return { lead: "NextGen has this visit as a no-show", body: "", pri: ["checkin", "Check in anyway"] };
    return { lead: "Taken off today’s list", body: "NextGen wasn’t changed.", pri: ["undoGone", "Put back on today’s list"] };
  }
  function visitCardHTML(v) {
    if (!v) return `<aside class="card detail"><div class="name">No visit picked</div><p class="body">${S.list.length ? "Pick a row or a chair to see the visit." : "Nobody is on today’s IVK list. Upload a fresh appointment list in Settings if that looks wrong."}</p></aside>`;
    const vs = visitState(v), k = esc(v.key), c = C.code(v), p = v.p, pc = hasPc(v), T = todayS();
    const sv = /start/i.test(v.a.event || "") ? S.SL.find(s => s.key === C.startKey(v.a.date, v.a.mrn)) || null : null;
    const where = ({ in: "Checked in", tx: "In treatment", done: "Checked out", noshow: "No-show", gone: "Not coming" })[v.ph] || "";
    const loc = where + (S.plan && v.st.chair && (v.ph === "in" || v.ph === "tx") ? ` · ${chairTxt(v.st.chair)}` : "");
    const meta = [`${C.fmtT(v.a.time)} ${kindTxt(v).toLowerCase()}`, C.provName(v.a.provider), `MRN ${v.a.mrn}`].filter(Boolean).join(" · ");
    const ix = p ? C.indexState(p, S.A, T, S.cfg) : null, kd = kindOf(v), pt = p ? ptOf(p.mrn) : null;
    const tags = [kd === "st" ? ["st", "Treatment start"] : null, ix && !ix.complete ? ["ix", `Index ${Math.min(ix.N, ix.done + 1)} of ${ix.N}`] : null, pc ? ["ix", "Requires post-check"] : null,
      pt?.oak ? ["", "Oakland"] : null, p?.freq ? ["", p.freq] : null].filter(Boolean);
    const arr = sentToday("arrival", c), ask = sentToday("protocol", c), post = v.st.seat != null && pc ? v.st.seat + +S.cfg.PostCheckMin : null;
    const st = (cls, t, m, w) => `<li class="step ${cls}"><span class="k">${cls === "done" ? "✓" : ""}</span><span><span class="t">${t}</span>${m ? `<span class="m">${m}</span>` : ""}</span><span class="w">${w || ""}</span></li>`;
    const steps = [
      v.st.arr != null ? st("done", "Checked in", arr ? `${esc(routeTo("arrival"))} told in Teams at ${timeOf(arr.at)}` : teamsReady("arrival") ? `${esc(routeTo("arrival"))} not told in Teams` : "Teams not set up for arrivals", C.fmtT(v.st.arr))
        : st(v.ph === "booked" || v.ph === "noshow" ? "now" : "next", "Check in", `${esc(routeTo("arrival"))} is told in Teams`, C.fmtT(v.a.time)),
      v.st.seat != null ? st("done", "In treatment", [S.plan && v.st.chair ? chairTxt(v.st.chair) : "", post != null ? `Post check at ${C.fmtT(post)}` : ""].filter(Boolean).join(" · "), `since ${C.fmtT(v.st.seat)}`)
        : st(v.ph === "in" ? "now" : "next", "Start treatment", "", ""),
      p && (v.ask.any || v.st.asked != null) ? (v.st.asked != null ? st("done", "Asked the providers", ask ? esc(routeTo("protocol")) : "Marked asked", C.fmtT(v.st.asked))
        : st(v.ph === "tx" ? "soon" : "next", "Ask the providers", v.ask.proto ? `Protocol not confirmed since ${p.protoAt != null ? C.fmtD(p.protoAt) : "it was set"}` : v.ask.thru != null ? `Booked through ${C.fmtD(v.ask.thru)}` : "Nothing booked after today", "")) : "",
      v.st.out != null ? st("done", "Checked out", v.st.booked === "yes" ? "Next visit booked" : v.st.booked === "no" ? "No next visit needed" : v.ask.book ? "Next visit not booked yet" : "Next visit on the list", C.fmtT(v.st.out))
        : st(v.ph === "tx" && !C.askDue(v) ? "now" : "next", "Check out", "Next visit, protocol, payment", ""),
      sv ? (sv.cmDone ? st("done", "Case Management updated", `In NextGen, ${esc(cmStage())}`, sv.cmAt != null ? C.fmtD(sv.cmAt) : "")
        : st(v.ph === "done" ? "soon" : "next", "Update Case Management", `In NextGen, to “${esc(cmStage())}”, once the start is kept`, "")) : ""
    ].join("");
    /* secondary steps as quiet buttons under the primary; corrections and lookups underneath them */
    const sec = [], quiet = [];
    if (v.ph === "tx" && vs.pri?.[0] === "ask") sec.push(["checkout", "Check out"]);
    if ((v.ph === "tx" || v.ph === "in") && vs.pri?.[0] !== "ask" && p) sec.push(["ask", "Ask the group chat"]);
    if (p && vs.pri?.[0] !== "printSched") sec.push(["printSched", "Print schedule"]);
    if (v.ph === "in") sec.push(["remind", `Remind ${routeTo("arrival")}`]);
    if (v.ph === "tx" && S.plan) sec.push(["seat", v.st.chair ? "Move to another chair" : "Pick a chair"]);
    if (p && ix?.complete) sec.push(["idxDone", "Move to Ongoing"]);
    if (sv && sv.cmDue && v.ph !== "booked") sec.push(["cmDone", "Case Management updated"]);
    const oak = p ? C.oakState(p, pt, S.A, T, S.cfg) : null;
    if (oak && oak.task && v.ph !== "booked") sec.push(["oakTask", "Mark OAK tasked"]);
    const xs = p ? C.xferState(S.XF.get(p.mrn), S.A, T, S.x.E) : null;
    if (xs && xs.due && v.ph !== "booked") sec.push(["xfer", `Confirm transfer to ${xs.to}`]);
    if (v.ph === "booked") quiet.push(["gone", "Not coming today"]);
    if (v.ph === "in") quiet.push(["undoIn", "Undo check-in"]);
    if (v.ph === "tx") quiet.push(["undoSeat", "Back to checked in"]);
    if (v.ph === "done") quiet.push(["undoOut", "Undo check-out"]);
    /* start paperwork is for a treatment start only; the start email goes out before it, so it's on the patient view */
    if (p && /start/i.test(v.a.event)) quiet.push(["startForms", "Start paperwork"]);
    /* today's payment: recorded at check-in or any time before check-out, and changed after it */
    if (p && S.payOn && v.ph !== "booked" && v.ph !== "gone") quiet.push(["payEdit", PARTS.some(x => C.payOf(S.pays, v.key, x)) ? "Change today’s payment" : "Record today’s payment"]);
    const ish = p ? C.indexShort(p, S.A, v.a, S.cfg) : null;
    const bt = (cls, [a, l]) => `<button class="btn ${cls}" data-act="${a}" data-k="${k}" data-m="${esc(v.a.mrn)}">${esc(l)}</button>`;
    return `<aside class="card detail" aria-label="${esc(c)}"><div class="dh"><div><div class="name">${esc(c)}${loc ? `<span class="loc">${esc(loc)}</span>` : ""}</div><div class="meta">${esc(meta)}</div></div>
        ${p ? `<button class="btn mute sm" data-act="openPt" data-m="${esc(v.a.mrn)}">Patient view</button>` : ""}</div>
      ${tags.length ? `<div class="tags">${tags.map(([t, l]) => `<span class="tag${t ? ` ${t}` : ""}">${esc(l)}</span>`).join("")}</div>` : ""}
      <p class="lead${vs.warn ? " soon" : ""}">${esc(vs.lead)}</p>${vs.body ? `<p class="body">${esc(vs.body)}</p>` : ""}
      ${ish ? `<div class="msg warn"><b>Index not fully booked.</b> ${ish.booked} of the ${ish.need} visits after this start are on the list, and ${esc(c)} is in Induction at ${esc(p.freq || "the index protocol")}. <button class="btn ghost sm" data-act="pvPlan" data-m="${esc(v.a.mrn)}">Plan index visits</button></div>` : ""}
      ${xs && !xs.due ? `<p class="note">Transfer to ${esc(xs.to)}${xs.date != null ? ` on ${C.fmtD(xs.date)}` : ""}; confirm it at their last visit here${xs.last ? ` (${C.fmtD(xs.last.date)})` : ""}.${thereTxt(xs.there, xs.to)}</p>` : ""}
      <ol class="steps">${steps}</ol>
      ${p ? factsHTML(p, v) : `<p class="note warn">This patient isn’t on the Master Patient Tracker yet, so there’s no protocol, auth or payment to check.</p><button class="btn ghost" data-act="profile" data-m="${esc(v.a.mrn)}">Create their profile</button>`}
      ${vs.pri ? `<button class="btn primary big" data-act="${vs.pri[0]}" data-k="${k}" data-m="${esc(v.a.mrn)}"${S.busy ? " disabled" : ""}>${esc(vs.pri[1])}</button>` : ""}
      ${sec.length ? `<div class="actions mid">${sec.map(g => bt("mute", g)).join("")}</div>` : ""}
      ${quiet.length ? `<div class="actions mid">${quiet.map(g => bt("mute sm", g)).join("")}</div>` : ""}</aside>`;
  }
  /* Protocol, Last confirmed, Booked through, Owes today; auth and balance only when they need attention */
  function factsHTML(p, v) {
    const a = v ? v.ask : C.asks(p, S.x, null), T = todayS(), d = C.dueOf(p);
    const dd = (t, cls) => `<dd${cls ? ` class="${cls}"` : ""}>${t}</dd>`;
    const paid = v && S.payOn ? PARTS.map(k => [k, C.payOf(S.pays, v.key, k)]).filter(x => x[1]) : [];
    const owes = paid.length ? paid.map(([k, x]) => `${PART_WORD[k]} ${C.money(x.amount)} ${x.kind === "collected" ? "collected" : "not collected"}`).join(" · ")
      : [d.copay > 0 ? `Provider visit ${C.money(d.copay)}` : "", d.infusion > 0 ? `Infusion ${C.money(d.infusion)}` : ""].filter(Boolean).join(" · ") || "Nothing";
    const authBad = p.authExp != null && p.authExp < T, authSoon = p.authExp != null && !authBad && p.authExp <= T + +S.cfg.AuthWarnDays, encLow = p.encLeft != null && p.encLeft <= +S.cfg.EncountersWarn;
    const L = v && S.payOn ? C.ledger(p.mrn, S.pays) : null, t = ptOf(p.mrn);
    return `<dl class="facts vfacts">
        <dt>Protocol</dt>${dd(esc(p.freq || "No frequency"), p.freq ? "" : "flag")}
        <dt>Last confirmed</dt>${dd(p.protoAt != null ? C.fmtD(p.protoAt) : "Not yet", a.proto ? "flag" : "")}
        <dt>Booked through</dt>${dd(a.thru != null ? C.fmtD(a.thru) : "Nothing after today", a.book ? (a.thru == null && C.isActive(p) && !/prn/i.test(p.freq) ? "bad" : "flag") : "")}
        <dt>Owes today</dt>${dd(esc(owes), !paid.length && d.total > 0 && v && v.ph !== "booked" ? "flag" : "")}
        ${authBad || authSoon ? `<dt>Auth</dt>${dd(`Ends ${C.fmtD(p.authExp)}${p.authNo ? `, ${esc(p.authNo)}` : ""}`, authBad ? "bad" : "flag")}` : ""}
        ${encLow ? `<dt>Visits on auth</dt>${dd(`${p.encLeft}${p.encAuth != null ? ` of ${p.encAuth}` : ""} left`, p.encLeft <= 0 ? "bad" : "flag")}` : ""}
        ${L && L.balance > 0 ? `<dt>Balance</dt>${dd(`${C.money(L.balance)}${t?.card ? `, card on file${t.okCollect ? ", okay to charge" : ""}` : ""}`, "flag")}` : ""}
      </dl>`;
  }

  /* ---------- Schedule ---------- */
  /* what kind of IVK visit: a treatment start (its own tab colour), an index visit (treatments 1 to N, holds, Induction; lavender),
     or a follow-up */
  /* an index visit is any visit of an Induction patient */
  const visitKind = a => { if (/start/i.test(a.event || "")) return "st"; const p = S.PM.get(a.mrn); return p && /induction/i.test(p.status) ? "ix" : ""; };
  const KIND_WORD = { st: "treatment start", ix: "index visit" };
  /* the small (i) on a visit booked off the template: hover or press it for the bubble */
  const MISS_TXT = "This visit does not match the templated slot-type";
  const missTag = () => `<span class="miss" role="button" tabindex="0" data-act="missTip" aria-label="${MISS_TXT}" title="${MISS_TXT}">i<span class="tipb" role="tooltip">${MISS_TXT}</span></span>`;
  function scheduleHTML() {
    const T = todayS(), start = C.weekStart(T) + 7 * S.wk, F = floor();
    /* open template starts from today on, holds NextGen doesn't have yet, and where each visit's post check lands */
    const DS = new Map([1, 2, 3, 4, 5].map(i => [start + i, C.daySlots(start + i, S.X)]));
    const openIn = (d, t) => d < T ? [] : DS.get(d).slots.filter(x => x.open && x.time >= t && x.time < t + 30);
    const holdIn = (d, t) => DS.get(d).bookings.filter(b => b.kind === "hold" && b.time >= t && b.time < t + 30);
    const postIn = (d, t) => DS.get(d).bookings.filter(b => b.post && b.postAt != null && b.postAt >= t && b.postAt < t + 30);
    const bookOf = a => DS.get(a.date)?.bookings.find(b => b.kind === "appt" && b.mrn === a.mrn && b.time === a.time) || null;
    const W = C.week(S.A, start, S.PM, [...DS.values()].flatMap(x => x.slots.map(y => y.time).concat(x.bookings.map(b => b.time), x.bookings.filter(b => b.postAt != null).map(b => b.postAt))));
    const elig = eligibleSet(start + 1, start + 6);
    const nOpen = [...DS.entries()].reduce((n, [d, x]) => n + (d < T ? 0 : x.slots.filter(y => y.openIndex && elig.has(slotKey(y))).length), 0);
    const stateOf = a => {
      if (C.isCancelled(a)) return "cx";
      if (/no ?show/i.test(a.status)) return "noshow";
      if (a.date < T) return "past";
      if (a.date > T) return "";
      const v = S.list.find(x => x.a === a || (x.a.mrn === a.mrn && x.a.time === a.time)); if (!v) return "";
      if (v.ph === "done" || v.ph === "gone") return "past";
      if (v.ph === "tx") return C.askDue(v) ? "soon" : "track";
      return v.ph === "in" ? "track" : "";
    };
    const blk = a => { const s = stateOf(a), p = a.p, flag = p && a.date >= T && C.futureOf(S.A, p.mrn, a.date).length === 0 && C.isActive(p) && !/prn/i.test(p.freq);
      const k = visitKind(a), bk = bookOf(a), miss = !!bk?.miss, pc = !!p?.reqPost, pk = bk?.post ? `${a.date}|${bk.postAt}|${bk.pkey}` : "";
      const label = `${a.code || "No code"}, MRN ${a.mrn}, at ${C.fmtT(a.time)}, ${C.provName(a.provider)}${k ? `, ${KIND_WORD[k]}` : ""}${pc ? ", requires a post-check" : ""}, ${a.status}${miss ? `. ${MISS_TXT}` : ""}`;
      return `<div class="blkw"><button class="blk ${s}${k ? ` k-${k}` : ""}" data-act="schedPick" data-m="${esc(a.mrn)}" data-d="${a.date}" data-t="${a.time}"${pk ? ` data-pc="${esc(pk)}"` : ""} aria-label="${esc(label)}"><b>${esc(a.code || "No code")}</b> <span class="pm">${esc(a.mrn)}</span> ${C.fmtT(a.time)}${pc ? ` <span class="pcb" title="Requires a post-check">PC</span>` : ""}${S.SO.get(a.mrn)?.status === "waiting" && S.SO.get(a.mrn).booked === a.date ? ` <span class="pcb so" title="On the sooner waitlist">Sooner</span>` : ""}<small>${esc(C.provName(a.provider))}${k === "st" ? ", start" : k === "ix" ? `, index${(() => { const n = p ? C.txNo(p, S.A, a.date, a.time) : null; return n ? ` ${n}` : ""; })()}` : ""}${s === "noshow" ? ", no-show" : flag ? ", last booked" : ""}</small></button>${miss ? missTag() : ""}</div>`; };
    const hblk = b => { const late = C.holdState(b.h, S.A, Date.now(), S.cfg) === "late", pk = `${b.h.date}|${b.postAt}|${b.pkey}`;
      return `<button class="blk hold k-ix${late ? " late" : ""}" data-act="ipPick" data-m="${esc(b.mrn)}" data-pc="${esc(pk)}" aria-label="${esc(`Hold for ${b.code} at ${C.fmtT(b.time)}${late ? ", not confirmed booked" : ""}`)}">Hold <b>${esc(b.code)}</b> ${C.fmtT(b.time)}<small>${esc(provLabel(b.provider))}${late ? ", not confirmed" : b.h.by ? `, by ${esc(b.h.by)}` : ""}</small></button>`; };
    /* an open slot: pressed, it's held for a patient (or, for an index start, planned from) */
    const oblk = (d, x) => { const ix = x.openIndex && elig.has(slotKey(x));
      return `<button class="blk open" data-act="schedOpen" data-d="${d}" data-t="${x.time}" data-p="${esc(x.provider)}" aria-label="${esc(`${ix ? "Open index start" : "Open follow-up"} ${C.fmtT(x.time)}, ${provLabel(x.provider)}: hold it for a patient`)}">${ix ? "Open index" : "Open follow-up"} ${C.fmtT(x.time)}<small>${esc(provLabel(x.provider))}</small></button>`; };
    /* where a post check lands, shaded */
    const pblk = (d, b) => `<div class="blk pcs" data-pck="${esc(`${d}|${b.postAt}|${b.pkey}`)}" title="${esc(`Post check for ${b.code}`)}">Post check <b>${esc(b.code)}</b> ${C.fmtT(b.postAt)}<small>${esc(provLabel(b.provider))}</small></div>`;
    const cells = W.rows.map(t => `<div class="tm">${C.fmtT(t)}</div>${W.days.map(d => `<div class="sc${d === T ? " td" : ""}">${W.cell(d, t).map(blk).join("")}${holdIn(d, t).map(hblk).join("")}${postIn(d, t).map(b => pblk(d, b)).join("")}${openIn(d, t).map(x => oblk(d, x)).join("")}</div>`).join("")}`).join("");
    const label = `${C.fmtD(start)} to ${C.fmtD(start + 6)}`;
    const nPc = new Set(S.A.filter(a => a.date > start && a.date < start + 6 && !C.isCancelled(a) && S.PM.get(a.mrn)?.reqPost).map(a => a.mrn)).size;
    return `${head("Schedule", `Week of ${label}: ${plural(W.count, "IVK visit")}${W.count ? `, ${W.days.map((d, i) => W.byDay[i] ? `${["Mon", "Tue", "Wed", "Thu", "Fri"][i]} ${W.byDay[i]}` : "").filter(Boolean).join(", ")}` : ""}${S.grids.length ? `, ${plural(nOpen, "index-eligible start")}` : ""}${nPc ? `, ${plural(nPc, "patient")} needing a post check` : ""}`, true)}
      <section class="card sched"><div class="wkbar"><span class="lbl">${label}</span>
        <button class="btn ghost sm" data-act="wk" data-v="-1">Previous week</button><button class="btn ghost sm" data-act="wk" data-v="0"${S.wk === 0 ? " disabled" : ""}>This week</button><button class="btn ghost sm" data-act="wk" data-v="1">Next week</button>
        <button class="btn ghost sm" data-act="printWeek" style="margin-left:auto">Print this week</button></div>
      <ul class="key skey"><li><span class="ktab st"></span>Treatment start</li><li><span class="ktab ix"></span>Index visit</li><li><span class="pcb">PC</span>Requires a post-check</li><li><span class="tsw pcs"></span>Where a post check lands</li><li><span class="miss sm">i</span>Not on its template slot</li><li class="muted">“last booked” = nothing booked after it</li></ul>
      <div class="sgrid"><div></div>${W.days.map(d => `<div class="ch${d === T ? " today" : ""}">${["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"][C.weekday(d) - 1]}<small>${C.fmtD(d)}</small></div>`).join("")}${cells}</div></section>${soonerHTML()}`;
  }

  /* ---------- sooner waitlist ---------- */
  /* the visit a waitlisted patient is booked for (a start needs an index start slot) */
  const soonerVisit = r => S.A.find(a => a.mrn === r.mrn && a.date === r.booked && (r.bookedTime == null || a.time === r.bookedTime) && !C.isCancelled(a)) || null;
  /* each waiting patient and the openings that would move them sooner: open template slots, and (after an upload) times a
     cancellation freed. extra = more openings to look through. */
  function soonerHits(extra) {
    const W = [...S.SO.values()].filter(r => r.status === "waiting"), T = todayS(); if (!W.length) return [];
    const until = Math.max(...W.map(r => r.booked)) - 1;
    const slots = S.grids.length ? memo(`soOpen|${until}`, () => C.openSlots(C.memoX(S.X), T, until)) : [];
    return W.map(r => { const v = soonerVisit(r); return { r, v, moved: !v, hits: C.soonerMatch(r, slots.concat(extra || []), T, !!v && /start/i.test(v.event)) }; });
  }
  const soonerTxt = r => [`booked ${dshort(r.booked)}${r.bookedTime != null ? ` ${C.fmtT(r.bookedTime)}` : ""}`, r.from != null ? `from ${C.fmtD(r.from)}` : "",
    r.days.size ? [...r.days].sort().map(d => `${C.DAYS5[+d[0] - 1]} ${d.slice(1).toUpperCase()}`).join(", ") : "any day", r.sameProv ? `with ${C.provName(r.provider)} only` : "any provider"].filter(Boolean).join(", ");
  /* an open slot can be held; a time freed by a cancellation off the template is booked straight in NextGen */
  const hitBtn = (r, c) => c.src === "freed" && !C.daySlots(c.date, S.X).slots.some(s => s.open && s.time === c.time && s.pkey === TM.provKey(c.provider))
    ? `<span class="pcb so">${dshort(c.date)} ${C.fmtT(c.time)}, ${esc(C.provName(c.provider))}: cancelled, book in NextGen</span>` : `<button class="btn ghost sm" data-act="schedOpen" data-d="${c.date}" data-t="${c.time}" data-p="${esc(c.provider)}" data-w="${esc(r.mrn)}">${dshort(c.date)} ${C.fmtT(c.time)}, ${esc(C.provName(c.provider))}${c.src === "freed" ? " (cancelled)" : ""}</button>`;
  function soonerHTML() {
    const L = soonerHits(); if (!L.length && !S.soonerOn) return "";
    const row = x => `<div class="row"><span class="p">${esc(x.r.code)}${mrnTag(x.r.mrn)}</span><span class="i">${esc(soonerTxt(x.r))}${x.r.note ? `. ${esc(x.r.note)}` : ""}${x.moved ? `<br><b>That visit isn’t on the list any more:</b> moved or cancelled?` : ""}
      ${x.hits.length ? `<br>Sooner: ${x.hits.slice(0, 4).map(c => hitBtn(x.r, c)).join(" ")}` : S.grids.length ? "<br><span class=\"muted\">No opening yet</span>" : ""}</span>
      <span class="w"><button class="btn mute sm" data-act="soEdit" data-m="${esc(x.r.mrn)}">Edit</button><button class="btn mute sm" data-act="soDone" data-m="${esc(x.r.mrn)}">Done</button></span></div>`;
    return `<section class="card list"><div class="lh"><h2>Sooner waitlist</h2><span>${L.length ? `${plural(L.length, "patient")}, ${L.filter(x => x.hits.length).length} with an opening` : "nobody waiting"}</span></div>
      ${needTable("tblIvkSooner", "The sooner waitlist")}
      ${L.map(row).join("") || `<p class="empty">Click a visit on the schedule to waitlist the patient for a sooner date.</p>`}
      <p class="note">${S.grids.length ? "Openings are open template slots before their visit, on the days they said" : "No template yet, so openings are only the times a cancellation frees on a new upload"}. Each upload alerts any new opening.</p></section>`;
  }

  /* ---------- Index planner ---------- */
  const DAYW =["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
  const dshort = d => `${["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"][C.weekday(d)]} ${C.fmtD(d)}`;
  /* a provider as the templates name them ("Ana Rivera, MD"), else as NextGen does ("Dr. Rivera") */
  const provLabel = n => { const k = TM.provKey(n), g = S.grids.find(x => x.key === k); return g ? TM.nameOf(g) : C.provName(n) || "any provider"; };
  const liveHolds = () => (S.holds || []).map(h => Object.assign({}, h, { state: C.holdState(h, S.A, Date.now(), S.cfg) })).filter(h => h.state === "held" || h.state === "late");
  const holdsOf = mrn => (S.holds || []).filter(h => h.mrn === mrn).map(h => Object.assign({}, h, { state: C.holdState(h, S.A, Date.now(), S.cfg), conf: C.holdConfirm(h, S.A, S.raw.apptUpdated) })).filter(h => h.state !== "released");
  /* holds marked booked that the uploaded NextGen report hasn't confirmed yet (pending), or that a newer report doesn't have (missing) */
  const unconfirmed = () => (S.holds || []).filter(h => h.status === "booked").map(h => Object.assign({}, h, { conf: C.holdConfirm(h, S.A, S.raw.apptUpdated) })).filter(h => h.conf === "pending" || h.conf === "missing");
  /* patients already in index, whose index visits can be changed here: Induction, or started within the index weeks */
  const inIndex = p => /induction/i.test(p.status);
  const listTime = () => S.raw.apptUpdated != null ? `${C.fmtD(Math.floor(S.raw.apptUpdated))} at ${C.fmtT(C.toMin(S.raw.apptUpdated))}` : "never";
  const uploadBtn = (l, cls) => `<label class="btn ${cls || "primary"}" style="cursor:pointer">${l}<input type="file" id="apptFile" accept=".xls,.xlsx,.xlsm,.xlsb,.csv,.txt,.xml,.htm,.html,.ods" hidden></label>`;
  const minsAgo = ms => Math.max(0, Math.floor((Date.now() - ms) / 60000));
  /* who is on the list: plans saved in the workbook, and consults waiting for their treatment start */
  function ipList() {
    const out = new Map();
    for (const c of S.CS) if (c.state === "act" && c.mrn && /treatment start/i.test(c.flag)) out.set(c.mrn, { mrn: c.mrn, code: c.code, c, from: `Consult ${c.consult != null ? C.fmtD(c.consult) : ""}`.trim() });
    for (const [mrn, pl] of S.plans) { if (/^done$/i.test(pl.note)) { out.delete(mrn); continue; } const o = out.get(mrn) || { mrn, code: pl.code || S.PM.get(mrn)?.code || "(no code)", from: "Added here" }; o.pl = pl; out.set(mrn, o); }
    return [...out.values()].map(o => Object.assign(o, { p: S.PM.get(o.mrn) || null, holds: holdsOf(o.mrn) }));
  }
  /* the plan for one patient: protocol, start, availability, and the proposal (or their holds) */
  function ipModel(o) {
    const cfg = S.cfg, pl = o.pl || {}, T = todayS();
    const perWeek = pl.perWeek || +cfg.IndexPerWeek, weeks = pl.weeks || +cfg.IndexWeeks;
    const start = pl.start ?? (o.c?.start != null && o.c.start > T ? o.c.start : T + 1), avail = C.availOf(pl, o.p);
    const pins = {}; for (const [n, p] of Object.entries(S.ipPins[o.mrn] || {})) if (pinOk(p)) pins[n] = p;
    const fixed = S.A.filter(a => a.mrn === o.mrn && a.date >= start && !C.isCancelled(a)).map(a => ({ date: a.date, time: a.time, provider: a.provider }));
    const res = C.planIndex({ start, today: T, perWeek, weeks, minGap: +cfg.MinGapDays, maxGap: +cfg.MaxGapDays, avail, X: S.X, fixed, pins });
    return { pl, perWeek, weeks, start, avail, res, custom: perWeek !== +cfg.IndexPerWeek || weeks !== +cfg.IndexWeeks, pins };
  }
  /* a hand-picked time is kept only while it is still an open index start */
  const pinOk = p => { const s = C.daySlots(p.date, S.X).slots.find(x => x.time === p.time && x.pkey === TM.provKey(p.provider)); return !!s && s.openIndex; };
  const asPin = s => ({ date: s.date, time: s.time, provider: s.provider, pkey: s.pkey, type: s.type, postUse: s.postUse, postType: s.postType });
  const holdLead = hs => { const held = hs.filter(h => h.state === "held" || h.state === "late"), late = held.filter(h => h.state === "late"), first = held[0];
    if (!held.length) return null;
    const ago = minsAgo(Math.min(...held.map(h => h.at ?? Date.now()))), left = +S.cfg.HoldMin - ago;
    return late.length ? { cls: "late", t: `Held ${ago} min ago${first.by ? ` by ${first.by}` : ""}, not confirmed booked in NextGen` }
      : { cls: "soon", t: `Held ${ago < 1 ? "just now" : `${ago} min ago`}${first.by ? ` by ${first.by}` : ""}. Book in NextGen and confirm within ${left} min` }; };
  /* the planner's options without a patient, on a copy of the templates that caches each day (worked out once per read) */
  const ipOpts = () => { const c = S.cfg; return { today: todayS(), perWeek: +c.IndexPerWeek, weeks: +c.IndexWeeks, minGap: +c.MinGapDays, maxGap: +c.MaxGapDays, X: memo("X", () => C.memoX(S.X)) }; };
  /* how many more patients could start index on the templates */
  const capacityNow = () => memo("cap", () => S.grids.length ? C.capacity(ipOpts(), +S.cfg.CapacityWeeks) : null);
  const slotKey = s => `${s.date}|${s.time}|${s.pkey}`;
  /* index-eligible starts: an open index start counts only when all 6 index visits still fit in their windows from it */
  const eligibleSet = (from, to) => memo(`el${from}|${to}`, () => { const T = todayS(), f = Math.max(from, T + 1);
    return new Set(S.grids.length && to >= f ? C.eligibleStarts(ipOpts(), f, 500, to - f + 1).map(e => slotKey(e.slot)) : []); });
  const nextEligible = () => memo("next", () => S.grids.length ? C.eligibleStarts(ipOpts(), todayS() + 1, 8, 7 * Math.max(8, +S.cfg.CapacityWeeks)) : []);

  function indexHTML() {
    const L = ipList(), H = liveHolds(), late = H.filter(h => h.state === "late"), cap = capacityNow();
    const byPt = new Map(); for (const h of H) byPt.set(h.mrn, (byPt.get(h.mrn) || []).concat(h));
    const sub = [`${plural(L.filter(o => !o.holds.length).length, "patient")} to plan`, byPt.size ? `${plural(byPt.size, "patient")} with holds to book` : "",
      late.length ? `<strong>${plural(new Set(late.map(h => h.mrn)).size, "patient")}’s holds not confirmed after ${S.cfg.HoldMin} min</strong>` : "",
      cap ? `room to start ${plural(cap.starts, "more index patient")} in the next ${S.cfg.CapacityWeeks} weeks` : "no template yet"].filter(Boolean).join(", ");
    const chips = accessChips();
    /* nobody picked: the planner works on its own (a draft plan from any start), with the next index-eligible starts beside it */
    if (S.ipSel == null) S.ipSel = (L.find(o => o.holds.some(h => h.state === "late")) || {}).mrn || "";
    if (S.ipSel && !L.find(o => o.mrn === S.ipSel)) S.ipSel = "";
    const cur = S.ipSel ? L.find(o => o.mrn === S.ipSel) : null, model = cur ? (cur.holds.length ? null : ipModel(cur)) : S.ipPins[""] ? ipModel(DRAFT()) : null;
    const row = o => { const hl = holdLead(o.holds), booked = o.holds.length && o.holds.every(h => h.state === "booked");
      const miss = o.holds.some(h => h.conf === "missing"), pend = o.holds.some(h => h.conf === "pending");
      const st = hl ? `<span class="f${hl.cls === "late" ? " bad" : ""}">${hl.cls === "late" ? "Holds not confirmed" : `${plural(o.holds.filter(h => h.state !== "booked").length, "hold")} to book`}</span>`
        : booked ? (miss ? `<span class="f bad">Not on the NextGen list</span>` : pend ? `<span class="f">Booked: upload to confirm</span>` : `<span class="ok">On the appointment list</span>`) : "Not held yet";
      return `<button class="crow" data-act="ipPick" data-m="${esc(o.mrn)}" aria-pressed="${S.ipSel === o.mrn}"${S.ipSel === o.mrn ? ' title="Press again to plan without a patient"' : ""}><span class="p">${esc(o.code)}${mrnTag(o.mrn)}</span><span class="proto">${esc(o.from)}${o.p ? `<br>${esc(o.p.status || "")}` : ""}</span><span class="thru">${o.pl?.start != null ? `Start ${C.fmtD(o.pl.start)}` : o.c?.start != null && o.c.start > todayS() ? `Start ${C.fmtD(o.c.start)}` : "As soon as possible"}</span><span class="flags">${st}</span></button>`; };
    /* new index patients aren't on the Master Patient Tracker yet: they're planned by LASFIR and MRN. The tracker list is only for
       changing the index visits of a patient already in index. */
    const opts = S.P.filter(p => inIndex(p) && !L.find(o => o.mrn === p.mrn)).sort((a, b) => a.code.localeCompare(b.code)).map(p => `<option value="${esc(p.mrn)}">${esc(p.code)}${p.status ? `, ${esc(p.status)}` : ""}${p.start != null ? `, started ${C.fmtD(p.start)}` : ""}</option>`).join("");
    const U = unconfirmed(), uMiss = U.filter(h => h.conf === "missing"), uPts = new Set(U.map(h => h.mrn)).size;
    const confirmBanner = U.length ? `<div class="msg ${uMiss.length ? "bad" : "warn"}"><b>${uMiss.length ? `${plural(new Set(uMiss.map(h => h.mrn)).size, "patient")}’s booked visits aren’t on the NextGen list uploaded ${listTime()}.` : `${plural(uPts, "patient")}’s index visits are marked booked, not confirmed yet.`}</b>
      ${uMiss.length ? "Check them in NextGen, then upload the report again." : `Run the NextGen appointment report and upload it to confirm the visits are booked (last upload ${listTime()}).`} <div class="actions">${uploadBtn("Upload the NextGen report", "primary inl")}</div></div>` : "";
    return `${setupBanner()}${head("Index planner", sub, false)}${chips}${confirmBanner}
      ${S.grids.length ? "" : `<div class="msg warn"><b>No provider template yet.</b> Plans are made from each provider’s start/index slots and the post checks that line up with them. <button class="btn ghost sm" data-act="nav" data-v="template">Set up the template</button></div>`}
      <div class="cols"><div class="stack"><section class="card plist"><div class="chead"><span>Patient</span><span>From</span><span>Start</span><span>Holds</span></div>
        ${L.map(row).join("") || `<p class="empty" style="padding:10px 14px">Nobody to plan. Add a new patient by LASFIR and MRN below. Approved consults waiting for a treatment start show here too.</p>`}
        <div class="addpt new"><span class="lbl">New patient</span><input id="ipNewCode" placeholder="LASFIR" aria-label="New patient’s LASFIR code" maxlength="8" autocomplete="off" spellcheck="false"><input id="ipNewMrn" placeholder="MRN" aria-label="New patient’s MRN" inputmode="numeric" autocomplete="off"><button class="btn ghost sm line" data-act="ipNewPt">Plan</button></div>
        <div class="addpt"><select id="ipAdd" aria-label="Change a patient in index"><option value="">Change the visits of a patient in index…</option>${opts}</select><button class="btn ghost sm line" data-act="ipAddPt">Change</button></div></section>
        ${holdsHTML(byPt)}${accessHTML(model)}</div>
      <div class="stack">${cur ? ipCardHTML(cur, model) : model ? draftCardHTML(model) : ""}${eligibleHTML(cur)}</div></div>`;
  }
  /* a plan with no patient yet: explored from any start, then held for whoever it's for */
  const DRAFT = () => ({ mrn: "", code: "", pl: null, p: null, c: null, holds: [], from: "", draft: true });
  /* the patient (or the draft) the planner is working on */
  const planTarget = () => S.ipSel ? ipList().find(x => x.mrn === S.ipSel) || null : DRAFT();
  /* The next index-eligible starts, standing on the right: each fits all of the index (2 a week for 3 weeks) on what's open now.
     Pressed: it becomes treatment 1 for the patient picked, or for a draft plan */
  function eligibleHTML(cur) {
    if (!S.grids.length) return "";
    const E = nextEligible(), c = S.cfg, busy = cur && cur.holds.length, N = C.indexN(c);
    const row = (e, i) => { const it = e.plan.items, last = it[it.length - 1], s0 = e.slot;
      return `<button class="erow" data-act="ipElig" data-i="${i}"${busy ? " disabled" : ""}><span class="t">${dshort(s0.date)}, ${C.fmtT(s0.time)}</span><span class="m">${esc(provLabel(s0.provider))}${s0.postUse === "flexpost" ? ", flex post check" : ""}. ${plural(it.length, "visit")} through ${dshort(last.date)}</span></button>`; };
    return `<section class="card elig"><div class="lh"><h2>Next index-eligible starts</h2><span>${E.length ? `from ${dshort(E[0].slot.date)}` : "none"}</span></div>
      ${E.length ? E.map(row).join("") : `<p class="note warn">No open index start in the next ${Math.max(8, +c.CapacityWeeks)} weeks has room for all ${N} index visits.</p>`}
      <p class="note">${busy ? `${esc(cur.code)} has holds: book or release them first.` : `A start counts only when all ${N} index visits (${c.IndexPerWeek} a week for ${c.IndexWeeks} weeks, ${c.MinGapDays} to ${c.MaxGapDays} days apart) are open after it. ${cur ? `Pressing one makes it treatment 1 for ${esc(cur.code)}.` : "Pressing one plans from it; pick who it’s for when you hold."}`}</p></section>`;
  }
  /* the draft's card: the visits, then who to hold them for (someone waiting, someone being planned, or a new LASFIR and MRN) */
  function draftCardHTML(M) {
    const r = M.res, it = r.items, cfg = S.cfg, N = r.N, ok = C.fullPlan(r);
    const lead = !it.length ? { cls: "soon", t: "No times fit yet" } : ok ? { cls: "ready", t: `${plural(it.length, "visit")}, ${dshort(it[0].date)} to ${dshort(it[it.length - 1].date)}` }
      : { cls: "soon", t: `${plural(it.length, "visit")} of ${N}. Check the notes` };
    const step = x => `<li class="step ${x.long ? "soon" : "now"}"><span class="k">${x.n}</span><span><span class="t">${dshort(x.date)}, ${C.fmtT(x.time)}</span>
        <span class="m">Week ${x.week}, ${esc(provLabel(x.provider))}, post check ${C.fmtT(x.post)}${x.postUse === "flexpost" ? " (flex)" : ""}${x.gap != null ? `, ${plural(x.gap, "day")} after the last` : ""}</span>
        ${S.ipAlt === x.n - 1 ? altHTML(r, x.n - 1) : ""}</span><span class="w"><button class="btn mute sm" data-act="ipAlt" data-n="${x.n - 1}">${S.ipAlt === x.n - 1 ? "Keep" : "Change"}</button></span></li>`;
    const planned = new Set(ipList().filter(o => o.holds.length).map(o => o.mrn)), seen = new Set();
    const opt = (mrn, code, why) => { if (!mrn || seen.has(mrn) || planned.has(mrn)) return ""; seen.add(mrn); return `<option value="${esc(mrn)}"${S.ipFor === mrn ? " selected" : ""}>${esc(code || mrn)}, ${esc(why)}</option>`; };
    const opts = [...S.WL.list.map(w => opt(w.mrn, w.code, w.next ? "starts next" : `#${w.pos} on the waitlist`)), ...ipList().map(o => opt(o.mrn, o.code, "on the planner")),
      ...S.gaps.filter(g => g.start).map(g => opt(g.mrn, g.code, "treatment start on the list"))].join("");
    return `<aside class="card detail" aria-label="Plan without a patient"><div class="dh"><div><div class="name">Index plan<span class="loc">no patient yet</span></div><div class="meta">${esc(`${cfg.IndexPerWeek} a week for ${cfg.IndexWeeks} weeks`)}</div></div></div>
      <p class="lead ${lead.cls}">${esc(lead.t)}</p>
      ${it.length ? `<ul class="steps">${it.map(step).join("")}</ul>` : ""}${r.notes.map(n => `<p class="note warn">${esc(n)}</p>`).join("")}
      ${it.length ? `<label class="fld">Hold these times for<select id="ipFor"><option value="">Choose the patient…</option>${opts}<option value="__new"${S.ipFor === "__new" ? " selected" : ""}>A new patient (LASFIR and MRN)</option></select></label>
        ${S.ipFor === "__new" ? `<div class="two"><label class="fld">LASFIR<input id="ipForCode" maxlength="8" autocomplete="off" spellcheck="false" placeholder="SMIJAN"></label><label class="fld">MRN<input id="ipForMrn" inputmode="numeric" autocomplete="off"></label></div>` : ""}
        <button class="btn primary" data-act="ipHoldDraft"${S.busy ? " disabled" : ""}>Hold ${plural(it.filter(x => !x.fixed).length, "time")}</button>` : ""}
      <div class="actions"><button class="btn ghost" data-act="ipReset" data-m="">Start over</button></div>
      <p class="note">Nothing is saved until you hold. Book the times in NextGen and confirm here within ${cfg.HoldMin} min.</p></aside>`;
  }
  const altHTML = (r, n) => { const A = C.altSlots(r, n, +S.cfg.MinGapDays).slice(0, 12);
    return `<div class="alts">${A.length ? A.map((s, i) => `<button data-act="ipUse" data-n="${n}" data-i="${i}">${dshort(s.date)}<small>${C.fmtT(s.time)}, ${esc(provLabel(s.provider))}${s.postUse === "flexpost" ? ", flex post check" : ""}</small></button>`).join("") : `<span class="muted small">No other open index start keeps a day off on each side.</span>`}</div>`; };
  /* every hold waiting to be booked, oldest first; past HoldMin they turn red */
  function holdsHTML(byPt) {
    if (!byPt.size) return "";
    const rows = [...byPt.values()].sort((a, b) => Math.min(...a.map(h => h.at || 0)) - Math.min(...b.map(h => h.at || 0))).map(hs => { const hl = holdLead(hs), h0 = hs[0];
      return `<div class="row" role="button" tabindex="0" data-act="ipPick" data-m="${esc(h0.mrn)}" data-state="${hl.cls === "late" ? "late" : ""}"><span class="p">${esc(h0.code)}</span><span class="w">${plural(hs.length, "time")}</span><span class="i"><span class="f${hl.cls === "late" ? " bad" : ""}">${esc(hl.t)}</span></span><span></span></div>`; }).join("");
    return `<section class="card list"><div class="lh"><h2>Holds to book in NextGen</h2><span>${byPt.size}</span></div>${rows}</section>`;
  }
  /* The week on the provider templates, as faint tiles. IVK slots (start/index, follow-up, post check, flex post check) say what's
     open, booked or held, tinted by their slot type; every other slot type shows plainly. Open index starts are buttons: pressed,
     one becomes treatment 1, or the treatment being changed. The selected patient's proposed visits are outlined and numbered. */
  const IVK_USES = ["start", "followup", "post", "flexpost"];
  function accessHTML(model) {
    if (!S.grids.length) return "";
    const T = todayS(), start = C.weekStart(T) + 7 * S.ipWk, days = [1, 2, 3, 4, 5].map(i => start + i), DS = days.map(d => C.daySlots(d, S.X));
    const times = [...new Set(DS.flatMap(x => x.cells.map(c => c.time)))].sort((a, b) => a - b);
    const mine = new Map((model?.res.items || []).map(x => [`${x.date}|${x.time}|${x.pkey}`, x])), many = S.X.providers > 1, elig = eligibleSet(start + 1, start + 5);
    /* Each start pairs with the post check at its end: hovering (or focusing) either one shades the other. Treatment starts carry
       the start tab, index visits and holds the lavender index tab; a visit off its template slot gets the (i). */
    const tile = (d, c, ds) => {
      const s = ds.slots.find(x => x.time === c.time && x.pkey === c.pkey), b = c.by, me = mine.get(`${d}|${c.time}|${c.pkey}`), ivk = IVK_USES.includes(c.use);
      let t, cls = ivk ? `tile ${c.use}` : "tile gen", pick = false, pc = "", miss = false;
      if (!ivk) t = b ? (b.kind === "other" ? b.event : b.code) : c.type;
      else if (c.use === "start" || c.use === "followup") {
        if (b && b.kind === "hold") { t = `Hold ${b.code}`; cls += " held k-ix"; }
        else if (b) { t = b.kind === "other" ? b.event : b.code; cls += " booked"; if (b.kind === "appt") { const k = visitKind({ event: b.event, mrn: b.mrn, date: d, time: b.time }); if (k) cls += ` k-${k}`; miss = !!b.miss; } }
        else if (s && s.full) { t = "Chairs full"; cls += " booked"; }
        else if (s && s.openIndex) { const ok = elig.has(slotKey(s)); t = ok ? "Index open" : "Open, index won’t fit"; pick = d > T; if (!ok) cls += " noidx"; }
        else t = c.use === "start" ? "Follow-up only" : "Follow-up";
        const at = b && b.postAt != null ? b.postAt : s && s.postUse ? s.post : null; if (at != null) pc = `${d}|${at}|${c.pkey}`;
      } else if (b) { t = b.kind === "other" ? b.event : `Post check, ${b.code}`; cls += " booked"; }
      else t = c.use === "post" ? "Post check" : "Flex post check";
      if (me) cls += " mine";
      const pck = c.use === "post" || c.use === "flexpost" ? ` data-pck="${esc(`${d}|${c.time}|${c.pkey}`)}"` : "", pca = pc ? ` data-pc="${esc(pc)}"` : "";
      const label = `${DAYW[C.weekday(d) - 1]} ${C.fmtT(c.time)}, ${c.who}: ${t}${me ? `, treatment ${me.n}` : ""}${miss ? `. ${MISS_TXT}` : ""}`;
      const body = `${me ? `<b class="nb">${me.n}</b>` : ""}<span>${esc(t)}</span>${many && ivk ? `<small>${esc(c.provider)}</small>` : ""}`;
      const st = ivk ? ` style="--c:${c.fill}"` : "";
      const el = pick ? `<button class="${cls} pick" data-act="ipTile" data-d="${d}" data-t="${c.time}" data-p="${esc(c.provider)}"${pca}${st} aria-label="${esc(label)}">${body}</button>` : `<div class="${cls}"${pca}${pck}${st} title="${esc(label)}"${pc ? ' tabindex="0"' : ""}>${body}</div>`;
      return miss ? `<div class="blkw">${el}${missTag()}</div>` : el;
    };
    const rows = times.map(t => `<div class="tm">${C.fmtT(t)}</div>${days.map((d, i) => `<div class="ac${d === T ? " td" : ""}${d < T ? " past" : ""}">${DS[i].cells.filter(c => c.time === t).map(c => tile(d, c, DS[i])).join("")}</div>`).join("")}`).join("");
    const nIdx = DS.reduce((n, x, i) => n + (days[i] > T ? x.slots.filter(s => s.openIndex && elig.has(slotKey(s))).length : 0), 0);
    const sw = (use, l) => { const t = S.types.find(x => x.use === use); return `<li><span class="tsw" style="--c:${t ? t.fill : "#999"}"></span>${l}</li>`; };
    return `<section class="card access"><div class="lh"><h2>The week on the templates</h2><span>${plural(nIdx, "index-eligible start")}</span></div>
      <div class="wkbar"><span class="lbl">${C.fmtD(start)} to ${C.fmtD(start + 6)}</span><button class="btn ghost sm" data-act="ipWk" data-v="-1"${S.ipWk <= 0 ? " disabled" : ""}>Previous week</button><button class="btn ghost sm" data-act="ipWk" data-v="0"${S.ipWk === 0 ? " disabled" : ""}>This week</button><button class="btn ghost sm" data-act="ipWk" data-v="1">Next week</button></div>
      <div class="agrid"><div></div>${days.map(d => `<div class="ch${d === T ? " today" : ""}">${DAYW[C.weekday(d) - 1]}<small>${C.fmtD(d)}</small></div>`).join("")}${rows}</div>
      ${(() => { const off = S.EX.filter(e => e.from <= start + 6 && e.to >= start); return off.length ? `<p class="note warn">Time off this week: ${off.map(e => `${esc(exLabel(e.provider))} ${exDates(e)}${e.cov ? ` (covered by ${esc(exLabel(e.cov))}${e.useTpl ? ", same template" : ""})` : ""}`).join("; ")}.</p>` : ""; })()}
      <ul class="key">${sw("start", "Start/index")}${sw("followup", "Follow-up")}${sw("post", "IVK post check")}${sw("flexpost", "Flex post check")}<li><span class="tsw gen"></span>Not IVK</li><li><span class="ktab st"></span>Treatment start</li><li><span class="ktab ix"></span>Index visit</li><li><span class="miss sm">i</span>Not on its template slot</li></ul>
      <p class="note">${model ? (S.ipAlt != null ? `Pick an open index start for treatment ${S.ipAlt + 1}.` : "Pick an open index start to make it treatment 1.") : "Pick an open index start to plan from it, with or without a patient. Hover a start to see its post check."} “Index open” means all ${C.indexN(S.cfg)} index visits still fit from it.</p></section>`;
  }
  function ipCardHTML(o, model) {
    const hs = o.holds, hl = holdLead(hs), p = o.p, cfg = S.cfg;
    const meta = [o.from, o.c?.provider, p?.status, p?.freq].filter(Boolean).join(", ");
    if (hs.length) {
      const all = hs.slice().sort((a, b) => a.date - b.date || a.time - b.time), open = all.filter(h => h.state !== "booked");
      const miss = all.filter(h => h.conf === "missing"), pend = all.filter(h => h.conf === "pending");
      const cls = h => h.conf === "onlist" ? "done" : h.conf === "missing" || h.state === "late" ? "late" : "soon";
      const how = h => h.conf === "onlist" ? ", on the appointment list" : h.conf === "missing" ? `, not on the list uploaded ${listTime()}` : h.conf === "pending" ? ", booked: upload the report to confirm" : "";
      const st = h => `<li class="step ${cls(h)}"><span class="k">${h.conf === "onlist" ? "" : h.n || ""}</span><span><span class="t">${dshort(h.date)}, ${C.fmtT(h.time)}</span><span class="m">${esc(provLabel(h.provider))}, post check ${C.fmtT(h.time + +cfg.PostCheckMin)}${how(h)}</span></span><span class="w">Treatment ${h.n || ""}</span></li>`;
      const lead = hl ? hl : miss.length ? { cls: "late", t: `${plural(miss.length, "booked visit")} not on the NextGen list uploaded ${listTime()}. Check NextGen, then upload again` }
        : pend.length ? { cls: "soon", t: `Marked booked in NextGen. Upload a fresh appointment list to confirm ${pend.length === all.length ? "them" : `the other ${pend.length}`}` } : { cls: "ready", t: `All ${plural(all.length, "visit")} on the appointment list` };
      return `<aside class="card detail" aria-label="${esc(o.code)}"><div class="dh"><div><div class="name">${esc(o.code)}${o.mrn ? `<span class="mrn">MRN ${esc(o.mrn)}</span>` : ""}</div><div class="meta">${esc(meta)}</div></div></div>
        <p class="lead ${lead.cls}">${esc(lead.t)}</p>
        <ul class="steps">${all.map(st).join("")}</ul>
        ${open.length ? `<button class="btn primary" data-act="ipBooked" data-m="${esc(o.mrn)}">Booked in NextGen</button><div class="actions"><button class="btn ghost" data-act="ipRelease" data-m="${esc(o.mrn)}">Release ${open.length === 1 ? "the hold" : "the holds"}</button></div>`
          : miss.length || pend.length ? `${uploadBtn("Upload the NextGen report")}${miss.length ? `<div class="actions"><button class="btn ghost" data-act="ipRelease" data-m="${esc(o.mrn)}">Release the missing ${miss.length === 1 ? "visit" : "visits"}</button></div>` : ""}`
          : `<button class="btn ghost" data-act="ipDone" data-m="${esc(o.mrn)}">Take off the list</button>`}
        <div class="actions"><button class="btn ghost sm line" data-act="printSched" data-m="${esc(o.mrn)}">Print schedule</button><button class="btn mute sm" data-act="startForms" data-m="${esc(o.mrn)}">Start paperwork</button><button class="btn mute sm" data-act="startEmail" data-m="${esc(o.mrn)}">Start email</button></div>
        <p class="note">Holds keep the start and its post check for everyone, on the Schedule, the Access week and the Index Planner tab. NextGen isn’t changed by the add-in.</p></aside>`;
    }
    const M = model || ipModel(o), r = M.res;
    const avTxt = M.pl.avail ? (/^any/i.test(M.pl.avail) ? "any time, set here" : "set here") : p && C.availOf({}, p) ? "from the tracker" : "any time";
    const av = M.avail || C.DAYS5.map(() => ({ am: true, pm: true }));
    const avail = `<div class="prefs av">${C.DAYS5.map(d => `<b>${d}</b>`).join("")}${["am", "pm"].map(h => av.map((x, i) => `<button data-act="ipAv" data-d="${i}" data-h="${h}" aria-pressed="${!!x[h]}" aria-label="${DAYW[i]} ${h.toUpperCase()}">${h.toUpperCase()}</button>`).join("")).join("")}</div>`;
    const per = `<select id="ipPer" aria-label="Visits a week">${[1, 2, 3].map(n => `<option${n === M.perWeek ? " selected" : ""}>${n}</option>`).join("")}</select>`;
    const wks = `<select id="ipWeeks" aria-label="Weeks">${[1, 2, 3, 4, 5, 6].map(n => `<option${n === M.weeks ? " selected" : ""}>${n}</option>`).join("")}</select>`;
    const inputs = `<label class="fld">Preferred start<input id="ipStart" type="date" value="${C.isoOf(M.start)}"></label>
      <div class="fld">Index protocol${M.custom ? " (custom)" : ""}<div class="proto-in">${per}<span>a week for</span>${wks}<span>weeks</span></div></div>
      <div class="section"><div class="h"><h2>Can come</h2><span>${avTxt}</span></div>${avail}
        <div class="actions">${M.pl.avail && p && C.availOf({}, p) ? `<button class="btn mute sm" data-act="ipAvReset">Use the tracker’s days</button>` : ""}${M.avail ? `<button class="btn mute sm" data-act="ipAvAny">Any time</button>` : ""}</div></div>`;
    const N = r.N, it = r.items, toHold = it.filter(x => !x.fixed);
    const lead = !S.grids.length ? { cls: "soon", t: "Set up the template first" } : !it.length ? { cls: "soon", t: "No times fit yet" }
      : it.length < N || it.some(x => x.long) || r.notes.length ? { cls: "soon", t: `${plural(it.length, "visit")} planned, ${dshort(it[0].date)} to ${dshort(it[it.length - 1].date)}. Check the notes` }
      : { cls: "ready", t: `${plural(it.length, "visit")}, ${dshort(it[0].date)} to ${dshort(it[it.length - 1].date)}` };
    const alt = n => altHTML(r, n);
    const step = x => `<li class="step ${x.fixed ? "done" : x.long ? "soon" : "now"}"><span class="k">${x.fixed ? "" : x.n}</span><span><span class="t">${dshort(x.date)}, ${C.fmtT(x.time)}</span>
        <span class="m">Week ${x.week}, ${esc(provLabel(x.provider))}, ${x.fixed ? "already booked, " : ""}post check ${C.fmtT(x.post)}${x.postUse === "flexpost" ? " (flex)" : ""}${x.gap != null ? `, ${plural(x.gap, "day")} after the last` : ""}${x.long ? ` (over ${cfg.MaxGapDays})` : ""}</span>
        ${S.ipAlt === x.n - 1 ? alt(x.n - 1) : ""}</span><span class="w">${x.fixed ? "" : `<button class="btn mute sm" data-act="ipAlt" data-n="${x.n - 1}">${S.ipAlt === x.n - 1 ? "Keep" : "Change"}</button>`}</span></li>`;
    const last = it[it.length - 1], afterD = last ? last.date + (C.freqDays(cfg.AfterIndex) || 7) : null;
    return `<aside class="card detail" aria-label="${esc(o.code)}"><div class="dh"><div><div class="name">${esc(o.code)}${o.mrn ? `<span class="mrn">MRN ${esc(o.mrn)}</span>` : ""}</div><div class="meta">${esc(meta)}</div></div></div>
      <p class="lead ${lead.cls}">${esc(lead.t)}</p>${inputs}
      ${it.length ? `<div class="section"><div class="h"><h2>Index visits</h2><span>week 1 starts ${dshort(it[0].date)}</span></div><ul class="steps">${it.map(step).join("")}</ul>
        ${r.notes.map(n => `<p class="note warn">${esc(n)}</p>`).join("")}${afterD ? `<p class="note">Then ${esc(cfg.AfterIndex)}, from around ${dshort(afterD)}.</p>` : ""}</div>` : r.notes.map(n => `<p class="note warn">${esc(n)}</p>`).join("")}
      ${toHold.length ? `<button class="btn primary" data-act="ipHold" data-m="${esc(o.mrn)}"${S.busy ? " disabled" : ""}>Hold ${toHold.length === 1 ? "this time" : `these ${toHold.length} times`}</button>` : ""}
      <div class="actions">${Object.keys(M.pins).length ? `<button class="btn ghost" data-act="ipReset" data-m="${esc(o.mrn)}">Undo my changes</button>` : ""}${o.pl ? `<button class="btn mute" data-act="ipDone" data-m="${esc(o.mrn)}">Take off the list</button>` : ""}</div>
      <p class="note">Holding saves the times at once, so the team sees them. Book them in NextGen and confirm here within ${cfg.HoldMin} min.</p></aside>`;
  }

  /* ---------- template configurator ---------- */
  /* The provider grids on the IVK Template tab, edited by dragging across half hours and picking a slot type; the tally, the
     checks and the slot types (colour, IVK use, eligible events) beside them. Every change is written to the tab at once. */
  const inkFor = h => { const n = parseInt(String(h).slice(1), 16); return ((n >> 16 & 255) * 0.299 + (n >> 8 & 255) * 0.587 + (n & 255) * 0.114) / 255 > 0.6 ? "#1B211F" : "#FFFFFF"; };
  const curGrid = () => S.grids[Math.min(S.tg || 0, S.grids.length - 1)] || null;
  const inRect = (q, d, t) => !!q && d >= Math.min(q.d0, q.d1) && d <= Math.max(q.d0, q.d1) && t >= Math.min(q.t0, q.t1) && t <= Math.max(q.t0, q.t1);
  /* "Lark MD, Ellis" -> { label: "Lark", title: "Ellis Lark, MD" } */
  function nameParts(n) {
    const s = C.txt(n), cred = (s.match(/\b(MD|DO|NP|PA|PMHNP|FNP)\b/i) || [""])[0].toUpperCase(), clean = x => x.replace(/\b(MD|DO|NP|PA|PMHNP|FNP|RN|PhD)\b\.?/gi, "").replace(/\s+/g, " ").trim();
    let first = "", last = "";
    if (s.includes(",")) { last = clean(s.split(",")[0]); first = clean(s.split(",").slice(1).join(" ")).split(" ")[0] || ""; }
    else { const w = clean(s).split(" "); last = w.pop() || ""; first = w[0] || ""; }
    return { label: last, title: `${[first, last].filter(Boolean).join(" ")}${cred ? `, ${cred}` : ""}` };
  }
  function templateHTML() {
    const r = S.raw, cfg = S.cfg;
    if (r.template == null || r.slotTypes == null) return `<section class="card set"><h3>Set up the template</h3><p>Adds two tabs in the Clinic Template Tool’s format: <b>${X.TPL_SHEET}</b>, one weekly grid per provider with its Event Types tally, and <b>${X.TYPES_SHEET}</b>, the slot types with their colours and the NextGen events each one takes. A Slot Types tab you already keep is used as it is.</p>
      <div class="actions"><button class="btn primary inl" data-act="tplSetup">Set up the template</button></div></section>`;
    const g = curGrid();
    /* the providers, toggled at the top right of the template card */
    const chips = `<div class="pchips" role="group" aria-label="Provider">${S.grids.map((x, i) => `<button class="chip sm" data-act="tplPick" data-v="${i}" aria-pressed="${x === g}">${esc(TM.nameOf(x))}${x.effFrom != null ? `<span class="n">from ${C.fmtD(x.effFrom)}</span>` : ""}</button>`).join("")}<button class="btn ghost sm line" data-act="tplAdd">Add a provider</button></div>`;
    if (!g) return `<section class="card tedit"><div class="lh"><h2>Provider templates</h2>${chips}</div><p class="empty">No provider grids yet. Add one for each provider who sees IVK patients; where two providers work the same hours, each gets a grid and the planner shows both.</p></section>`;
    const T = TM.tally(g, S.types), warn = TM.checks(g, S.types, +cfg.PostCheckMin);
    const idx = [1, 2, 3, 4, 5].map(d => TM.startsOn(g, d, S.types, +cfg.PostCheckMin).filter(s => s.index).length);
    const sel = S.tsel && S.tsel.g === g.top ? S.tsel : null, nSel = sel ? g.dayCols.filter(x => inRect({ ...sel, t0: 0, t1: 1e4 }, x.d, 0)).length * g.rows.filter(x => inRect({ ...sel, d0: 0, d1: 9 }, 0, x.t)).length : 0;
    const cell = (d, t) => { const x = g.cells.get(d + "|" + t), ty = x ? TM.typeOf(S.types, x.type) : null;
      return `<button class="tc${inRect(sel, d, t) ? " sel" : ""}" data-d="${d}" data-t="${t}"${ty ? ` style="background:${ty.fill};color:${ty.font}"` : ""} aria-label="${TM.DAYS[d]} ${C.fmtT(t)}: ${esc(x ? x.type : "empty")}">${x ? esc(x.type) : ""}</button>`; };
    const editor = `<section class="card tedit"><div class="lh"><div><h2>${esc(TM.nameOf(g))}</h2><span class="muted small">${esc(g.location || cfg.Clinic)}${g.modified != null ? `, last modified ${C.fmtD(g.modified)}` : ""}</span></div>${chips}</div>
      ${sel ? `<div class="tbar" role="toolbar" aria-label="Set the selected half hours"><b>${plural(nSel, "half hour")}</b><span class="muted small">set to</span>${S.types.map(t => `<button class="tchip" data-act="tSet" data-v="${esc(t.type)}" style="background:${t.fill};color:${t.font}">${esc(t.type)}</button>`).join("")}<button class="btn mute sm" data-act="tSet" data-v="">Clear</button><button class="btn mute sm" data-act="tCancel">Cancel</button></div>`
        : `<p class="note">Drag across the half hours (or Shift-click) to select them, then pick a slot type. It’s written to the ${X.TPL_SHEET} tab straight away.</p>`}
      <div class="tg" style="grid-template-columns:76px repeat(${g.dayCols.length},minmax(0,1fr))"><div class="tgh">Time</div>${g.dayCols.map(x => `<div class="tgh">${TM.DAYS[x.d]}</div>`).join("")}
        ${g.rows.map(x => `<div class="tgt">${C.fmtT(x.t)}</div>${g.dayCols.map(c => cell(c.d, x.t)).join("")}`).join("")}</div></section>`;
    const trow = (t, x, cls) => `<tr class="${cls || ""}"><td>${t ? `<span class="tsw" style="--c:${t.fill}"></span>` : ""}${esc(t ? t.type : x.label)}</td><td>${x.slots}</td><td>${x.hours}</td>${x.byDay.map(n => `<td>${n || ""}</td>`).join("")}</tr>`;
    const tally = `<section class="card ttally"><div class="lh"><h2>Event types</h2><span>${plural(idx.reduce((a, b) => a + b, 0), "index start")} a week</span></div>
      <table class="tt"><thead><tr><th>Event type</th><th>Slots</th><th>Hours</th>${TM.TALLY_DAYS.map(d => `<th>${d}</th>`).join("")}</tr></thead>
      <tbody>${T.rows.filter(x => x.slots).map(x => trow(S.types.find(t => t.type === x.type), x)).join("")}${trow(null, Object.assign({ label: "Clinical" }, T.clinical), "sum")}${trow(null, Object.assign({ label: "Total" }, T.total), "sum")}</tbody></table>
      <p class="note">Index starts by day: ${idx.map((n, i) => `${C.DAYS5[i]} ${n}`).join(", ")}. A start counts when a post check or flex post check is ${cfg.PostCheckMin} min after it.</p>
      ${warn.map(w => `<p class="note warn">${esc(w)}</p>`).join("")}</section>`;
    return `<div class="cols"><div class="stack">${editor}${ptoHTML()}</div><div class="stack">${tally}${typesHTML()}</div></div>`;
  }
  /* ---------- Access › Demand ---------- */
  /* What the schedule has to hold at everyone's current protocol, against how they're booked and what the templates offer, week by
     week; and when index patients stepping down to once a week (after treatment 6) open access. */
  const n1 = v => { const r = Math.round(v * 10) / 10; return Number.isInteger(r) ? String(r) : r.toFixed(1); };
  function demandHTML() {
    const o = S.dm || (S.dm = { paused: false, sched: "all" }), c = S.cfg;
    const D = memo(`dm|${o.paused}|${o.sched}`, () => C.demand({ patients: S.P, appts: S.A, holds: S.holds, listAt: S.raw.apptUpdated, today: todayS(), cfg: c, weeks: 12, paused: o.paused, sched: o.sched, X: S.grids.length ? ipOpts().X : null }));
    const W = D.weeks, sumW = (k, n) => W.slice(0, n).reduce((a, w) => a + w[k], 0), unsched = D.rows.filter(r => r.unsched);
    const sub = [`${plural(D.rows.length, "patient")}${o.paused ? " (with paused)" : ""}`, `${n1(D.perWeek)} visits a week at their protocols`, D.rows.length ? `${n1(D.avg)} a week each on average` : "",
      unsched.length && o.sched !== "exclude" ? `<strong>${plural(unsched.length, "patient")} with nothing booked</strong>` : ""].filter(Boolean).join(", ");
    const seg = (k, l) => `<button data-act="dmSched" data-v="${k}" aria-pressed="${o.sched === k}">${l}</button>`;
    const bar = `<div class="dmbar"><button class="chip" data-act="dmPaused" aria-pressed="${o.paused}">Include paused</button>
      <div class="pick seg" role="group" aria-label="Unscheduled patients">${seg("all", "Everyone")}${seg("exclude", "Exclude unscheduled")}${seg("only", "Unscheduled only")}</div></div>`;
    const tile = (l, v, s, cls) => `<div class="stat${cls ? ` ${cls}` : ""}"><small>${l}</small><b>${v}</b>${s ? `<span>${s}</span>` : ""}</div>`;
    const tb4 = D.toBook4, tb8 = D.toBook8, open4 = sumW("open", 4), open8 = sumW("open", 8);
    const stats = `<div class="stats">${tile("Visits a week needed", n1(D.perWeek), `${n1(D.avg)} per patient`)}${tile("Next 4 weeks", n1(sumW("expect", 4)), `${sumW("booked", 4)} booked, ${n1(tb4)} to book`, tb4 > open4 && S.grids.length ? "warn" : "")}
      ${tile("Next 8 weeks", n1(sumW("expect", 8)), `${sumW("booked", 8)} booked, ${n1(tb8)} to book`, tb8 > open8 && S.grids.length ? "warn" : "")}${S.grids.length ? tile("Open IVK slots, 4 weeks", String(open4), `${open8} in 8 weeks`) : tile("Open IVK slots", "—", "no template yet")}</div>`;
    const G = D.groups, gRow = g => `<tr${g.freq === "Index" ? ' class="ix"' : ""}><td>${esc(g.freq)}${g.freq === "Index" ? `<small>${c.IndexPerWeek} a week to treatment ${D.N}</small>` : ""}</td><td>${g.n}</td><td>${n1(g.rate)}</td><td>${n1(g.perWeek)}</td>
      <td>${n1(g.d4)}</td><td>${g.b4}</td><td${g.t4 > 0.5 ? ' class="flag"' : ""}>${n1(g.t4)}</td><td>${n1(g.d8)}</td><td>${g.b8}</td><td${g.t8 > 0.5 ? ' class="flag"' : ""}>${n1(g.t8)}</td><td>${g.unsched || ""}</td></tr>`;
    const tot = G.reduce((a, g) => ({ n: a.n + g.n, perWeek: a.perWeek + g.perWeek, d4: a.d4 + g.d4, b4: a.b4 + g.b4, t4: a.t4 + g.t4, d8: a.d8 + g.d8, b8: a.b8 + g.b8, t8: a.t8 + g.t8, unsched: a.unsched + g.unsched }), { n: 0, perWeek: 0, d4: 0, b4: 0, t4: 0, d8: 0, b8: 0, t8: 0, unsched: 0 });
    const byProto = `<section class="card dmt"><div class="lh"><h2>By protocol</h2><span>needed at the protocol, booked, still to book</span></div><div class="tscroll"><table class="tt dm">
      <thead><tr><th rowspan="2">Protocol</th><th rowspan="2">Patients</th><th rowspan="2">A week each</th><th rowspan="2">A week, all</th><th colspan="3">Next 4 weeks</th><th colspan="3">Next 8 weeks</th><th rowspan="2">Nothing booked</th></tr>
      <tr><th>Needed</th><th>Booked</th><th>To book</th><th>Needed</th><th>Booked</th><th>To book</th></tr></thead>
      <tbody>${G.map(gRow).join("") || `<tr><td colspan="11" class="muted">No patients in this view.</td></tr>`}${G.length ? `<tr class="sum"><td>Total</td><td>${tot.n}</td><td>${n1(D.avg)}</td><td>${n1(tot.perWeek)}</td><td>${n1(tot.d4)}</td><td>${tot.b4}</td><td>${n1(tot.t4)}</td><td>${n1(tot.d8)}</td><td>${tot.b8}</td><td>${n1(tot.t8)}</td><td>${tot.unsched || ""}</td></tr>` : ""}</tbody></table></div>
      <p class="note">Needed is each patient’s protocol (the tracker’s Frequency) over the weeks; PRN and blank frequencies need none. To book is each patient’s need less what they have booked in the window, so one patient booked far ahead doesn’t cover another. Index patients count at ${c.IndexPerWeek} a week until treatment ${D.N}, then ${esc(c.AfterIndex)}. Booked counts NextGen visits and holds.</p></section>`;
    const wRow = w => `<tr><td>${w.w === 0 ? "This week" : `Week of ${C.fmtD(w.start)}`}</td><td>${n1(w.expect)}</td><td>${w.booked}</td><td${w.toBook > 0.5 ? ' class="flag"' : ""}>${n1(w.toBook)}</td>${S.grids.length ? `<td>${w.slots}</td><td${w.open < w.toBook ? ' class="bad"' : ""}>${w.open}</td>` : ""}<td>${w.freed > 0.05 ? `+${n1(w.freed)}` : ""}${w.steps.length ? `<small>${esc(w.steps.map(r => r.code).join(", "))}</small>` : ""}</td></tr>`;
    const weeks = `<section class="card dmt"><div class="lh"><h2>Week by week</h2><span>Sunday to Saturday</span></div><div class="tscroll"><table class="tt dm">
      <thead><tr><th>Week</th><th>Needed</th><th>Booked</th><th>To book</th>${S.grids.length ? "<th>IVK slots</th><th>Open</th>" : ""}<th>Freed by index step-down</th></tr></thead><tbody>${W.map(wRow).join("")}</tbody></table></div>
      <p class="note">${S.grids.length ? "IVK slots are the template’s start/index and follow-up times, less time off; open ones have no visit or hold and a free chair. " : ""}Freed is the visits a week index patients no longer need once they step down to ${esc(c.AfterIndex)} after treatment ${D.N}.</p></section>`;
    const T = D.transitions, tRow = r => { const proj = !(r.p && C.treatments(r.p, S.A)[D.N - 1]); const fr = Math.max(0, D.perIdx - r.afterRate);
      return `<li class="step ${proj ? "soon" : "now"}"><span class="k"></span><span><span class="t">${esc(r.code)}: treatment ${D.N} ${proj ? "around" : "on"} ${dshort(r.idxEnd)}</span><span class="m">Then ${esc(c.AfterIndex)}${fr ? `, frees about ${n1(fr)} visit${fr === 1 ? "" : "s"} a week from ${dshort(r.idxEnd + 1)}` : ""}${proj ? " (not booked yet: projected)" : ""}</span></span><span class="w"><button class="btn mute sm" data-act="openPt" data-m="${esc(r.mrn)}">Open</button></span></li>`; };
    const firstFree = W.find(w => w.freed > 0.05);
    const steps = `<section class="card dmt"><div class="lh"><h2>When index creates access</h2><span>${T.length ? plural(T.length, "patient") + " in index" : "nobody in index"}</span></div>
      ${T.length ? `<p class="note">${firstFree ? `Access starts opening ${firstFree.w === 0 ? "this week" : `the week of ${C.fmtD(firstFree.start)}`}; by the week of ${C.fmtD(W[W.length - 1].start)}, about ${n1(W[W.length - 1].freed)} visits a week are free from index step-downs.` : ""}</p><ul class="steps">${T.map(tRow).join("")}</ul>` : `<p class="note">Index patients (Induction, or started within the index weeks) show here with the week they step down.</p>`}</section>`;
    const uRow = r => `<div class="row"><span class="p">${esc(r.code)}${mrnTag(r.mrn)}</span><span class="w">${esc(r.proto || "—")}</span><span class="i">${r.paused ? "Paused, " : ""}${n1(r.idx ? D.perIdx : r.rate)} a week needed, nothing booked</span><button class="btn ghost sm" data-act="openPt" data-m="${esc(r.mrn)}">Open</button></div>`;
    const ul = o.sched === "only" ? `<section class="card list"><div class="lh"><h2>Unscheduled patients</h2><span>${unsched.length}</span></div>${unsched.map(uRow).join("") || `<p class="empty">Everyone in this view has a visit booked.</p>`}</section>` : "";
    return `${setupBanner()}${head("Access", sub, false)}${accessChips()}${bar}${stats}${ul}${byProto}<div class="cols"><div class="stack">${weeks}</div><div class="stack">${steps}</div></div>`;
  }
  /* Access › Templates: the provider grids, their tally, time off and the slot types */
  function templatePageHTML() {
    const n = S.grids.length, idx = S.grids.reduce((k, g) => k + [1, 2, 3, 4, 5].reduce((a, d) => a + TM.startsOn(g, d, S.types, +S.cfg.PostCheckMin).filter(s => s.index).length, 0), 0);
    return `${setupBanner()}${head("Templates", n ? `${plural(n, "provider")}, ${plural(idx, "index start")} a week across all their templates` : "No provider template yet", false)}${accessChips()}${templateHTML()}`;
  }
  /* the four Access pages */
  const accessChips = () => { const B = badges(); return `<div class="chips achips" role="group" aria-label="Access">${ACCESS.map(([k, l]) => { const b = B[k];
    return `<button class="chip" data-act="nav" data-v="${k}" aria-pressed="${S.view === k}">${l}${k === "template" ? `<span class="n">${S.grids.length}</span>` : b && k !== "access" ? `<span class="n">${b.n}</span>` : ""}</button>`; }).join("")}</div>`; };
  /* Time off and coverage: a provider's PTO takes their template out on those days (an exception by date); a covering provider can
     work the absent provider's template (pre-filled) or their own */
  const exLabel = n => { const g = S.grids.find(x => x.key === TM.provKey(n)); return g ? TM.nameOf(g) : C.provFull(n) || n; };
  const exDates = e => `${dshort(e.from)}${e.to !== e.from ? ` to ${dshort(e.to)}` : ""}`;
  function ptoHTML() {
    const T = todayS(), L = S.EX.filter(e => e.to >= T).sort((a, b) => a.from - b.from), ok = !!S.raw.tables?.own?.tblIvkExceptions;
    const row = e => { const cl = C.exClashes(e, S.A, S.holds);
      return `<li class="step ${cl.length ? "late" : e.cov ? "done" : "soon"}"><span class="k"></span><span><span class="t">${esc(exLabel(e.provider))}: ${exDates(e)}</span>
        <span class="m">${e.cov ? `Covered by ${esc(exLabel(e.cov))}, ${e.useTpl ? `working ${esc(exLabel(e.provider))}’s template` : "on their own template"}` : "Not covered: their slots are closed"}${e.note ? `, ${esc(e.note)}` : ""}${cl.length ? `. <b>${plural(cl.length, "visit or hold", "visits or holds")} still with them${e.useTpl ? `: move them to ${esc(exLabel(e.cov))} in NextGen` : ": move or cancel them"}</b> (${cl.slice(0, 4).map(a => `${esc(a.code)} ${dshort(a.date)}${a.hold ? " hold" : ""}`).join(", ")})` : ""}</span></span>
        <span class="w"><button class="btn mute sm" data-act="ptoEdit" data-v="${esc(e.id)}">Edit</button></span></li>`; };
    return `<section class="card tpto"><div class="lh"><h2>Time off and coverage</h2><span>${L.length ? plural(L.length, "upcoming") : "none upcoming"}</span></div>
      ${needTable("tblIvkExceptions", "Time off")}
      ${L.length ? `<ul class="steps">${L.map(row).join("")}</ul>` : `<p class="note">When a provider is off, add it here: their template’s slots close on those days, on the Access week, the Schedule and in every plan.</p>`}
      <div class="actions"><button class="btn ghost sm line" data-act="ptoAdd"${ok && S.grids.length ? "" : " disabled"}>Add time off</button></div></section>`;
  }
  /* Slot types: colour, IVK use, Clinical or Blocked, and the NextGen events each takes (a multi-select of the events on the
     appointment list and the wider report). Drafted here; Save writes the Slot Types tab and repaints every grid. */
  function typesHTML() {
    const te = S.te || S.types, used = new Set(S.grids.flatMap(g => [...g.cells.values()].map(x => x.type.toLowerCase())));
    /* the events on offer: every NextGen event seen, plus any a slot type already names, spelled as written; NextGen holds aren't visits */
    const named = te.flatMap(t => String(t.accepts || "").split(/[,;\n]/).map(a => a.trim()).filter(a => a && a !== "*" && !a.includes("+")));
    const events = [...new Map(S.events.concat(named).filter(e => !/\bhold\b/i.test(e)).map(e => [e.toLowerCase(), S.events.find(v => v.toLowerCase() === e.toLowerCase()) || e])).values()].sort();
    const on = (t, ev) => TM.acceptTerms(t.accepts).some(x => x.length === 1 && x[0] === ev.toLowerCase());
    const row = (t, i) => { const any = /(^|,)\s*\*\s*(,|$)/.test(t.accepts), L = String(t.accepts || "").split(/[,;\n]/).map(a => a.trim()).filter(Boolean);
      return `<div class="tline"><div class="tl1"><input type="color" class="tcol" data-ti="${i}" data-k="fill" value="${esc(t.fill)}" aria-label="Colour of ${esc(t.type)}">
        <input class="tname" data-ti="${i}" data-k="type" value="${esc(t.type)}" aria-label="Slot type name"${used.has(t.type.toLowerCase()) ? ` readonly title="In use on the template; add a new type instead of renaming"` : ""}>
        ${used.has(t.type.toLowerCase()) ? "<span></span>" : `<button class="x" data-act="teDel" data-i="${i}" aria-label="Remove ${esc(t.type)}">×</button>`}</div>
        <div class="tl2"><select data-ti="${i}" data-k="use" aria-label="IVK use of ${esc(t.type)}">${TM.USES.map(u => `<option value="${u}"${t.use === u ? " selected" : ""}>${TM.USE_LABEL[u]}</option>`).join("")}</select>
        <select data-ti="${i}" data-k="kind" aria-label="Kind of ${esc(t.type)}">${["Clinical", "Blocked"].map(k => `<option${t.kind === k ? " selected" : ""}>${k}</option>`).join("")}</select>
        <details class="ms"><summary title="${esc(L.join(", "))}">${any ? "Any visit" : L.length ? esc(L.length === 1 ? L[0] : plural(L.length, "event")) : "No events"}</summary><div class="msl">
          <label class="opt"><input type="checkbox" data-ti="${i}" data-ev="*"${any ? " checked" : ""}> Any visit</label>
          ${events.map(ev => `<label class="opt"><input type="checkbox" data-ti="${i}" data-ev="${esc(ev)}"${on(t, ev) ? " checked" : ""}> ${esc(ev)}</label>`).join("") || `<span class="muted small">Upload the NextGen report to list its events.</span>`}</div></details></div></div>`; };
    return `<section class="card ttypes"><div class="lh"><h2>Slot types</h2><span>${S.te ? "not saved yet" : `on the ${X.TYPES_SHEET} tab`}</span></div>
      <p class="note">Eligible events are the NextGen visits a slot takes. Start/index takes treatment starts and, second, follow-ups; a flex post check can take a TMS follow-up or tele visit when no index patient needs it.</p>
      ${te.map(row).join("")}
      <div class="actions"><button class="btn ghost sm line" data-act="teAdd">Add a slot type</button>${S.te ? `<button class="btn primary inl" data-act="teSave"${S.busy ? " disabled" : ""}>Save slot types</button><button class="btn mute sm" data-act="teUndo">Undo</button>` : ""}</div></section>`;
  }
  /* where each grid's tally ends (its Total: row) on the tab, so saving the slot types rewrites only that block */
  /* placed from the grid's Monday header row (the layout puts it 4 rows under the title), not from the labels found above it */
  const gridInfo = () => S.grids.map(g => { const V = S.raw.template || [], r0 = g.anchor[0] - 4, hdr = C.txt((V[r0 + 2] || [])[8]) === "Event Types";
    let end = null; if (hdr) for (let r = r0 + 3; r < Math.min(V.length, r0 + 60); r++) if (C.txt((V[r] || [])[8]) === "Total:") { end = r; break; }
    const c0 = TM.colL(g.dayCols[0].c), c1 = TM.colL(g.dayCols[g.dayCols.length - 1].c);
    return { r0, slots: `${c0}${g.rows[0].r + 1}:${c1}${g.rows[g.rows.length - 1].r + 1}`, step: g.step, start: g.rows[0].t, end: g.rows[g.rows.length - 1].t + g.step, tallyEnd: end }; });

  /* ---------- Patients ---------- */
  const CHIPS = [["need", "Needs attention"], ["profile", "New patients"], ["active", "Active"], ["paused", "Paused"], ["disc", "Discontinued"], ["all", "All"]];
  function chipOf(k, p, n) {
    if (k === "need") return n.length > 0;
    if (k === "profile") return false;   // new patients aren't on the tracker yet: they're listed above the table
    if (k === "active") return C.isActive(p);
    if (k === "paused") return /paused/i.test(p.status);
    if (k === "disc") return /discontinued/i.test(p.status);
    return true;
  }
  /* the table sorts by any column, ascending or descending (press the header again to flip it) */
  const SORTS = [["code", "Patient"], ["freq", "Protocol"], ["status", "Status"], ["conf", "Protocol confirmed"], ["thru", "Booked through"], ["need", "Needs"]];
  const sortVal = (k, r) => k === "code" ? r.p.code : k === "freq" ? (C.freqDays(r.p.freq) ?? (r.p.freq ? 1e4 : 1e5)) : k === "status" ? (r.p.status || "~").toLowerCase()
    : k === "conf" ? r.p.protoAt ?? -1 : k === "thru" ? r.thru ?? -1 : (r.loud.some(y => y.bad) ? 1000 : 0) + r.loud.length;
  function patientsHTML() {
    const T = todayS(), L = S.P.map(p => { const n = C.needs(p, S.x), fut = C.futureOf(S.A, p.mrn, T - 1); return { p, n, loud: C.loud(n), thru: fut.length ? fut[fut.length - 1].date : null }; });
    const off = S.gaps.filter(g => !g.onMpt);
    const counts = Object.fromEntries(CHIPS.map(([k]) => [k, L.filter(r => chipOf(k, r.p, r.loud)).length + (k === "profile" ? off.length : 0)]));
    const srt = S.sort || { k: "need", dir: -1 }, q = S.q.trim().toUpperCase();
    const shown = L.filter(r => chipOf(S.chip, r.p, r.loud) && (!q || r.p.code.includes(q))).sort((a, b) => {
      const x = sortVal(srt.k, a), y = sortVal(srt.k, b), c = typeof x === "string" ? x.localeCompare(y) : x - y;
      return c * srt.dir || a.p.code.localeCompare(b.p.code); });
    const red = L.filter(r => r.loud.some(y => y.bad)).length, act = L.filter(r => C.isActive(r.p)).length;
    const sub = [`${act} active`, red ? `<strong>${red} with something blocking the next infusion</strong>` : "", `${counts.need} need attention`, S.gaps.length ? `${plural(S.gaps.length, "profile")} to create` : ""].filter(Boolean).join(", ");
    const th = ([k, l]) => { const on = srt.k === k; return `<button class="sorth" data-act="pSort" data-v="${k}" aria-sort="${on ? (srt.dir > 0 ? "ascending" : "descending") : "none"}" title="Sort by ${l.toLowerCase()}">${l}<span class="arr" aria-hidden="true">${on ? (srt.dir > 0 ? "▲" : "▼") : ""}</span></button>`; };
    const row = r => { const p = r.p, conf = p.protoAt != null ? C.fmtD(p.protoAt) : "Not yet", stale = C.isActive(p) && (p.protoAt == null || T - p.protoAt > +S.cfg.ProtocolDays);
      const fl = r.loud.length ? r.loud.slice(0, 2).map(y => `<span class="f${y.bad ? " bad" : ""}">${esc(y.t)}</span>`).join(", ") + (r.loud.length > 2 ? ` <span class="muted">and ${r.loud.length - 2} more</span>` : "") : `<span class="ok">${esc(r.n[0]?.t || "On track")}</span>`;
      return `<button class="prow" data-act="pOpen" data-m="${esc(p.mrn)}"><span class="p">${esc(p.code)}${mrnTag(p.mrn)}${p.reqPost ? ` <span class="pcb" title="Requires a post-check">PC</span>` : ""}</span><span>${esc(p.freq || "No frequency")}</span><span class="muted">${esc(p.status || "No status")}</span>
        <span${stale ? ' class="flagt"' : ""}>${conf}</span><span>${r.thru != null ? C.fmtD(r.thru) : "Nothing booked"}</span><span class="flags">${fl}</span></button>`; };
    const offRows = (S.chip === "profile" || S.chip === "need") && off.length ? `<section class="card list"><div class="lh"><h2>On the appointment list, not on the tracker</h2><span>${off.length}</span></div>
      ${off.map(g => `<div class="row grow" data-state="${g.start ? "soon" : ""}"><span class="p">${esc(g.code || "—")}</span><span class="w">${dshort(g.first.date)}</span><span class="i">${g.start ? `<span class="f">Treatment start</span>, ` : ""}MRN ${esc(g.mrn)}, ${C.fmtT(g.first.time)} with ${esc(C.provName(g.first.provider))}</span><button class="btn ghost sm" data-act="profile" data-m="${esc(g.mrn)}">Create profile</button></div>`).join("")}</section>` : "";
    return `${setupBanner()}${head("Patients", sub, false)}
      <div class="chips">${CHIPS.map(([k, l]) => `<button class="chip" data-act="chip" data-v="${k}" aria-pressed="${S.chip === k}">${l}<span class="n">${counts[k]}</span></button>`).join("")}
        <input class="search" id="ptSearch" type="search" placeholder="Find a code" aria-label="Find a patient by code" value="${esc(S.q)}" autocomplete="off"></div>
      ${offRows}
      <section class="card ptable"><div class="phead">${SORTS.map(th).join("")}</div>
        ${shown.map(row).join("") || `<p class="empty" style="padding:10px 14px">${S.chip === "need" ? "Nothing here. Every patient in this group is on track." : "No patients in this group."}</p>`}</section>`;
  }
  /* a provider can ask for a post check on every visit (acuity): the patient is then booked only in starts that end on a post check */
  function postCheckHTML(p) {
    const L = p.reqPost && S.grids.length ? C.postCheckSlots(S.X, todayS() + 1, 5) : [];
    return `<div class="section"><div class="h"><h2>Post check</h2><span>${p.reqPost ? "required" : "not required"}</span></div>
      <label class="opt"><input type="checkbox" id="reqPost" data-m="${esc(p.mrn)}"${p.reqPost ? " checked" : ""}> Requires a post-check<small>Tick when the provider asks for one. Book them in a start whose last 30 min is a post check or flex post check.</small></label>
      ${p.reqPost ? `<p class="note">${L.length ? `Next open with a post check: ${L.map(s => `${dshort(s.date)} ${C.fmtT(s.time)}`).join(", ")}.` : S.grids.length ? "No open start with a free post check in the next 6 weeks." : "Set up the template to list open times with a post check."}</p>` : ""}</div>`;
  }

  /* ---------- the patient view ---------- */
  /* Laid out as the TMS Tracker's: the patient's lead need and its action at the top, today's visit if they're here, then one
     section at a time from the side: Schedule, Protocol, Benefits (with the auth), Financial. The section is remembered on this
     computer. */
  const PV_SECT = [["sched", "Schedule"], ["proto", "Protocol"], ["ben", "Benefits"], ["fin", "Financial"]];
  const pvSect = () => { const v = S.pvSect || LS.get("ivkPtSect", "sched"); return PV_SECT.some(s => s[0] === v) ? v : "sched"; };
  const PH_WORD = { booked: "booked", in: "checked in", tx: "in treatment", done: "checked out", noshow: "no-show", gone: "not coming today" };
  function ptViewHTML(p) {
    const n = C.needs(p, S.x), L = C.loud(n), first = L[0], tv = S.list.find(v => v.a.mrn === p.mrn), on = pvSect(), m = esc(p.mrn);
    const meta = [p.status || "No status", p.freq || "no frequency", p.loc, p.start != null ? `started ${C.fmtD(p.start)}` : "", p.reqPost ? "requires post-check" : ""].filter(Boolean).join(" · ");
    const A = y => `data-act="${y.act[1]}" data-m="${m}"`;
    const also = L.slice(1, 5).map(y => `<div class="row"><span class="f ${y.bad ? "bad" : "flag"} needline">${esc(y.t)}</span>${y.act ? `<button class="btn ghost sm" ${A(y)}>${esc(y.act[0])}</button>` : ""}</div>`).join("");
    const quiet = n.filter(y => y.quiet).map(y => `<span class="tag">${esc(y.t)}</span>`).join("");
    const sect = { sched: pvSched, proto: pvProto, ben: pvBen, fin: pvFin };
    const back = { patients: "Patients", schedule: "Schedule", today: "Today" }[S.pvFrom] || "the list";
    return `<div class="top"><div><h1>${esc(p.code)} <span class="mrn">MRN ${esc(p.mrn)}</span></h1><p class="sub">${esc(meta)}</p></div><div class="topr">${tools()}<button class="btn ghost sm line" data-act="pvClose">Back to ${esc(back)}</button></div></div>
      <div class="pview">
      ${tv && ["in", "tx"].includes(tv.ph) ? `<div class="pvtoday"><span><b>Here today</b> · ${S.plan && tv.st.chair ? `${esc(chairTxt(tv.st.chair))} · ` : ""}${tv.ph === "tx" ? `In treatment since ${C.fmtT(tv.st.seat)}` : `Checked in ${C.fmtT(tv.st.arr)}`}</span><button class="btn ghost sm" data-act="toToday" data-k="${esc(tv.key)}">Open visit</button></div>`
        : tv ? `<div class="pvtoday quiet"><span><b>Today</b> ${C.fmtT(tv.a.time)}, ${esc(PH_WORD[tv.ph] || "")}</span><button class="btn ghost sm" data-act="toToday" data-k="${esc(tv.key)}">Today’s visit</button></div>` : ""}
      <p class="lead ${first ? (first.bad ? "late" : "soon") : ""}">${esc(first ? first.t : "Nothing needed right now")}</p>${quiet ? `<div class="tags">${quiet}</div>` : ""}
      ${first && first.act ? `<button class="btn primary" ${A(first)}>${esc(first.act[0])}</button>` : ""}
      ${also ? `<div class="section needs"><div class="h"><h2>Also needs</h2><span>${L.length - 1} more</span></div>${also}</div>` : ""}
      <div class="pvgrid"><nav class="pvnav" aria-label="Sections">${PV_SECT.map(([k, l]) => `<button data-act="pvShow" data-v="${k}" aria-current="${on === k}">${l}</button>`).join("")}</nav>
        <div class="pvs card" data-sect="${on}">${sect[on](p)}</div></div></div>`;
  }
  /* Schedule: what's booked (start and index visits marked, post checks, anything off the template), holds, recent visits, the
     printouts and the patient's availability */
  function pvSched(p) {
    const T = todayS(), m = esc(p.mrn), fut = C.scheduleOf(p, S.A, T), past = S.A.filter(a => a.mrn === p.mrn && a.date < T && !C.isCancelled(a)).sort((a, b) => b.date - a.date).slice(0, 5);
    const miss = (d, t) => !!C.daySlots(d, S.X).bookings.find(b => b.kind === "appt" && b.mrn === p.mrn && b.time === t)?.miss;
    const row = x => { const a = { event: x.type === "Treatment start" ? "start" : "", mrn: p.mrn, date: x.date, time: x.time }, k = visitKind(a), mi = S.grids.length && miss(x.date, x.time);
      return `<li class="step now${k ? ` k-${k}` : ""}"><span class="k">${x.n || ""}</span><span><span class="t">${esc(x.day)}, ${esc(x.at)}${mi ? missTag() : ""}</span><span class="m">${x.loc ? `${esc(x.loc)}, ` : ""}${esc(x.provider)}, ${k === "st" ? "treatment start" : k === "ix" ? "index visit" : "infusion"}${p.reqPost || k ? `, post check ${C.fmtT(x.time + +S.cfg.PostCheckMin)}` : ""}${x.status && !/expected/i.test(x.status) ? `, ${esc(x.status)}` : ""}</span></span><span class="w">${x.n ? `Treatment ${x.n}` : ""}</span></li>`; };
    const hs = holdsOf(p.mrn).filter(h => h.state !== "booked" || h.conf !== "onlist");
    /* start paperwork and the start email only for a treatment start on the list (the email only before its day), never follow-ups */
    const st = C.nextStart(p, S.A, T), vc = C.visitCounts(p.mrn, S.A, S.x.E, T);
    const away = (S.x.E || []).filter(a => a.mrn === p.mrn && a.date >= T && !C.isCancelled(a)).sort((a, b) => a.date - b.date || (a.time ?? 0) - (b.time ?? 0));
    return `<div class="h"><h2>Schedule</h2><span>${fut.length ? `${plural(fut.length, "visit")} from today` : "nothing booked"}</span></div>
      <dl class="facts"><dt>At ${esc(S.cfg.Clinic)}</dt><dd>${plural(vc.kept + vc.booked, "appointment")}: ${vc.kept} kept, ${vc.booked} booked</dd>${vc.away.map(x => `<dt>${esc(x.loc || "Another clinic")}</dt><dd>${plural(x.n, "appointment")} booked</dd>`).join("")}</dl>
      ${fut.length ? `<ul class="steps">${fut.slice(0, 12).map(row).join("")}</ul>${fut.length > 12 ? `<p class="note">And ${plural(fut.length - 12, "more")}.</p>` : ""}` : `<p class="note${C.isActive(p) && !/prn/i.test(p.freq) ? " warn" : ""}">No IVK visit booked from today.</p>`}
      ${away.length ? `<div class="section"><div class="h"><h2>At other clinics</h2><span>as of ${C.fmtD(S.elseAt)}</span></div><ul class="steps">${away.slice(0, 8).map(a => `<li class="step soon"><span class="k"></span><span><span class="t">${dshort(a.date)}, ${C.fmtT(a.time)}</span><span class="m">${esc(a.loc)}, ${esc(C.provName(a.provider))}${a.status && !/expected/i.test(a.status) ? `, ${esc(a.status)}` : ""}</span></span><span class="w"></span></li>`).join("")}</ul></div>` : ""}
      ${hs.length ? `<div class="section"><div class="h"><h2>Holds</h2><span>${hs.length}</span></div><ul class="steps">${hs.map(h => `<li class="step soon k-ix"><span class="k">${h.n || ""}</span><span><span class="t">${dshort(h.date)}, ${C.fmtT(h.time)}</span><span class="m">${esc(provLabel(h.provider))}, ${h.status === "booked" ? "booked: upload to confirm" : h.state === "late" ? "not confirmed booked" : "held"}</span></span><span class="w"></span></li>`).join("")}</ul><div class="actions"><button class="btn ghost sm line" data-act="ipPick" data-m="${m}">Open in the Index planner</button></div></div>` : ""}
      <div class="actions"><button class="btn ghost sm line" data-act="printSched" data-m="${m}">Print schedule</button>${st ? `<button class="btn mute sm" data-act="startForms" data-m="${m}">Start paperwork</button>` : ""}${st && st.date > T ? `<button class="btn mute sm" data-act="startEmail" data-m="${m}">Start email</button>` : ""}<button class="btn mute sm" data-act="pvPlan" data-m="${m}">Plan index visits</button></div>
      ${past.length ? `<div class="section"><div class="h"><h2>Recent visits</h2></div><dl class="facts">${past.map(a => `<dt>${dshort(a.date)}</dt><dd>${C.fmtT(a.time)}, ${esc(C.provName(a.provider))}${/no ?show/i.test(a.status) ? ", no-show" : ""}</dd>`).join("")}</dl></div>` : ""}
      ${availHTML(p)}`;
  }
  /* Protocol: the frequency, when it was last confirmed (and every confirmation logged), index progress, post check, Oakland */
  function pvProto(p) {
    const m = esc(p.mrn), a = C.asks(p, S.x, null), T = todayS(), N = C.indexN(S.cfg), L = C.treatments(p, S.A), ok = C.txNo(p, S.A, L[0]?.date, L[0]?.time) != null;
    const hist = S.log.filter(l => l.ev === "confirmed" && l.code === p.code).slice(-6).reverse(), lastAsk = lastSent("protocol", p.code);
    const done = L.filter(x => x.date < T).length, idxDay = L[N - 1] ? L[N - 1].date : null;
    return `<div class="h"><h2>Protocol</h2><span>${esc(p.freq || "no frequency")}</span></div>
      <dl class="facts"><dt>Frequency</dt><dd${p.freq ? "" : ' class="flag"'}>${esc(p.freq || "Not on the tracker")}</dd>
        <dt>Last confirmed</dt><dd${a.proto ? ' class="flag"' : ""}>${p.protoAt != null ? `${C.fmtD(p.protoAt)}, ${C.dateOf(p.protoAt).getFullYear()} (${plural(T - p.protoAt, "day")} ago)` : "Not yet"}</dd>
        <dt>Confirm again by</dt><dd>${p.protoAt != null ? C.fmtD(p.protoAt + +S.cfg.ProtocolDays) : "Now"}</dd>
        ${lastAsk ? `<dt>Last asked</dt><dd>${dateOfIso(lastAsk.at)} at ${timeOf(lastAsk.at)}${lastAsk.by ? `, ${esc(lastAsk.by)}` : ""}</dd>` : ""}
        ${ok ? `<dt>Index</dt><dd>${done >= N ? `Done${idxDay ? ` ${C.fmtD(idxDay)}` : ""}, then ${esc(S.cfg.AfterIndex)}` : `${done} of ${N} done${idxDay ? `, treatment ${N} on ${C.fmtD(idxDay)}` : ""}`}</dd>` : ""}</dl>
      ${hist.length ? `<p class="note">Confirmed: ${hist.map(l => `${dateOfIso(l.at)}${l.msg ? ` (${esc(l.msg)})` : ""}${l.by ? `, ${esc(l.by)}` : ""}`).join("; ")}.</p>` : ""}
      <div class="actions"><button class="btn ghost sm line" data-act="proto" data-m="${m}">Confirm protocol</button>${C.isActive(p) ? `<button class="btn mute sm" data-act="ask" data-m="${m}">Ask the group chat</button>` : ""}</div>
      ${postCheckHTML(p)}${oakHTML(p)}${xferHTML(p)}`;
  }
  /* Transfer: a patient whose home clinic isn't this one (or an Oakland patient) can be discontinued here as transferred now, or have
     a transfer scheduled for a future date that's confirmed at their last visit here */
  /* the hand-off from the other clinics' visits (a wider report): booked there, or nothing there yet; nothing said without one */
  const thereTxt = (e, to) => S.x.E == null ? "" : e ? ` Booked at ${esc(e.loc || to)} ${C.fmtD(e.date)}.` : ` Nothing booked at ${esc(to)} yet (other clinics’ visits as of ${C.fmtD(S.elseAt)}).`;
  function xferHTML(p) {
    const xf = S.XF.get(p.mrn), xs = C.xferState(xf, S.A, todayS(), S.x.E), m = esc(p.mrn);
    if (!xf && !C.elsewhere(p, S.cfg) && !C.oakState(p, ptOf(p.mrn), S.A, todayS(), S.cfg)) return "";
    const st = xs ? `to ${esc(xs.to)}${xs.date != null ? ` ${C.fmtD(xs.date)}` : ""}` : xf?.status === "done" ? `done${xf.confirmed != null ? ` ${C.fmtD(xf.confirmed)}` : ""}` : "none scheduled";
    return `<div class="section"><div class="h"><h2>Transfer</h2><span>${st}</span></div>
      ${needTable("tblIvkTransfers", "Transfers")}
      ${xs ? `<p class="note${xs.due || xs.after.length ? " warn" : ""}">${xs.due ? `<b>Confirm the transfer</b>${xs.last ? `: their last visit here is ${C.fmtD(xs.last.date)}` : ""}.` : `Confirm it at their last visit here${xs.last ? `, ${C.fmtD(xs.last.date)}` : ""}.`}${xs.after.length ? ` ${plural(xs.after.length, "visit")} still booked here after ${C.fmtD(xs.date)}: cancel ${xs.after.length === 1 ? "it" : "them"} in NextGen.` : ""}${thereTxt(xs.there, xs.to)}${xf.note ? ` ${esc(xf.note)}` : ""}</p>`
        : xf?.status === "done" ? `<p class="note">Transferred to ${esc(xf.to)}${xf.note ? `. ${esc(xf.note)}` : ""}.${S.x.E ? thereTxt(C.bookedAt(S.x.E, p.mrn, xf.to, todayS()), xf.to) : ""}</p>` : `<p class="note">Home clinic ${esc(p.loc || "not set")}. Schedule a transfer for a future date, or discontinue here now as transferred.</p>`}
      <div class="actions"><button class="btn ghost sm line" data-act="xfer" data-m="${m}"${S.xferOn ? "" : " disabled"}>${xs ? (xs.due ? "Confirm transfer" : "Change transfer") : "Transfer"}</button></div></div>`;
  }
  /* Benefits: who pays each part of the visit, the latest Sohar benefits, the authorization, the BIDF */
  function pvBen(p) {
    const m = esc(p.mrn), pf = S.PF.get(p.mrn) || null, cv = C.coverage(p, pf), sh = pf?.sohar || {}, T = todayS();
    const modeTxt = cv.mode ? (C.COVER_MODES.find(x => x[0] === cv.mode) || [])[1] : !cv.prov.cover && !cv.inf.cover ? "Coverage not set" : !cv.inf.cover ? "Infusion coverage not set" : "Provider visit coverage not set";
    const part = (l, x) => `<dt>${l}</dt><dd>${x.cover === "selfpay" ? "Self pay" : esc(x.payer || (x.cover ? "Insurance, payer not named" : "—"))}${x.due ? `, patient pays ${C.money(x.due)}` : x.cover ? ", nothing due" : ""}</dd>`;
    const v = (l, x) => x ? `<dt>${l}</dt><dd>${esc(x)}</dd>` : "";
    const both = (a, b) => [a, b].filter(Boolean).join(", ");
    const auth = p.authExp == null ? `<dd${C.isActive(p) ? ' class="flag"' : ""}>No end date</dd>` : `<dd class="${p.authExp < T ? "bad" : p.authExp <= T + +S.cfg.AuthWarnDays ? "flag" : ""}">Ends ${C.fmtD(p.authExp)}</dd>`;
    return `<div class="h"><h2>Benefits</h2><span>${esc(modeTxt)}</span></div>
      ${needTable("tblIvkProfile", "Profiles")}
      <dl class="facts">${part("Provider visit (99214/99215)", cv.prov)}${part("Infusion", cv.inf)}</dl>
      <div class="section"><div class="h"><h2>Sohar benefits</h2><span${sh.date != null && T - sh.date > 90 ? ' class="flagt"' : ""}>${sh.date != null ? `verified ${C.fmtD(sh.date)}` : "none entered"}</span></div>
        ${sh.payer || sh.status ? `<dl class="facts">${v("Payer", sh.payer)}${v("Plan", sh.plan)}${v("Network", sh.network)}${sh.covStart != null || sh.covEnd != null ? `<dt>Coverage</dt><dd>${sh.covStart != null ? C.fmtD(sh.covStart) : "—"} to ${sh.covEnd != null ? C.fmtD(sh.covEnd) : "—"}</dd>` : ""}
          ${v("Copay", sh.copay)}${v("Coinsurance", sh.coins)}${v("Deductible", both(sh.ded && `${sh.ded} total`, sh.dedRem && `${sh.dedRem} left`))}${v("OOP max", both(sh.oop && `${sh.oop} total`, sh.oopRem && `${sh.oopRem} left`))}${v("Status", sh.status)}${v("Sohar ref", sh.ref)}</dl>` : `<p class="note">Enter the latest Sohar verification on the profile, or fill it from the Sohar Eligibility add-in’s benefits file.</p>`}</div>
      <div class="section"><div class="h"><h2>Authorization</h2><span>${esc(p.authNo || "")}</span></div><dl class="facts">
        <dt>Auth number</dt><dd>${esc(p.authNo || "—")}</dd><dt>Starts</dt><dd>${p.authStart != null ? C.fmtD(p.authStart) : "—"}</dd><dt>Ends</dt>${auth}
        <dt>Visits left</dt><dd${p.encLeft != null && p.encLeft <= +S.cfg.EncountersWarn ? ` class="${p.encLeft <= 0 ? "bad" : "flag"}"` : ""}>${p.encLeft != null ? `${p.encLeft}${p.encAuth != null ? ` of ${p.encAuth}` : ""}` : esc(p.encLeftTxt || p.encAuthTxt || "Not tracked")}</dd>
        ${p.pcn ? `<dt>PCN</dt><dd>${esc(p.pcn)}${p.task ? `, ${esc(p.task)}` : ""}</dd>` : ""}</dl></div>
      <div class="section"><div class="h"><h2>BIDF</h2></div><dl class="facts"><dt>2026 BIDF</dt><dd${!p.bidf && C.isActive(p) ? ' class="flag"' : ""}>${esc(p.bidf || "Not recorded")}</dd><dt>Printed at the desk</dt><dd>${pf?.bidfPrinted != null ? C.fmtD(pf.bidfPrinted) : "—"}</dd></dl>
        <p class="note">The BIDF is completed by the benefits team and printed from NextGen.</p></div>
      <div class="actions"><button class="btn ghost sm line" data-act="profile" data-m="${m}">Edit coverage and benefits</button><button class="btn mute sm" data-act="edit" data-m="${m}">Edit tracker fields</button></div>
      ${p.notes ? `<p class="note">Clinic notes: ${esc(p.notes)}</p>` : ""}`;
  }
  const pvFin = p => payHTML(p);
  const startVisit = p => C.futureOf(S.A, p.mrn, todayS() - 1).find(a => /start/i.test(a.event)) || (/induction/i.test(p.status) ? C.futureOf(S.A, p.mrn, todayS() - 1)[0] : null) || null;
  const ptOf = mrn => S.PT.get(mrn) || null;
  const needTable = (n, what) => S.raw.tables?.own?.[n] ? "" : `<p class="note warn">${what} need the newer tables. Use <b>Update the workbook</b> at the top of the page or in Settings.</p>`;
  /* Payments: the balance from missed payments, what each visit owes (the provider visit and the infusion, collected separately),
     Luma card on file and okay to charge at the time of service, reminders, and visits nobody has said were paid */
  function payHTML(p) {
    const t = ptOf(p.mrn), L = C.ledger(p.mrn, S.pays), due = C.dueOf(p), m = esc(p.mrn), T = todayS(), conf = S.payOn ? C.toConfirm(p, S.A, S.pays, T, S.payStart) : [];
    const last = L.lastReminder, remDue = L.balance > 0 && (!last || T - (last.date ?? 0) >= +S.cfg.ReminderDays), cv = C.coverage(p, S.PF.get(p.mrn));
    const chk = (k, on, l, s) => `<label class="opt"><input type="checkbox" data-pt="${k}" data-m="${m}"${on ? " checked" : ""}${S.raw.tables?.own?.tblIvkPatient ? "" : " disabled"}> ${l}${s ? `<small>${s}</small>` : ""}</label>`;
    const partOf = k => k.endsWith("|prov") ? "provider visit" : k.endsWith("|inf") ? "infusion" : "";
    const miss = L.open.map(r => `<li class="step soon"><span class="k"></span><span><span class="t">${r.date != null ? dshort(r.date) : "Visit"}: ${C.money(r.left)} ${partOf(r.key) ? `${partOf(r.key)} ` : ""}not collected</span><span class="m">${r.left < r.amount ? `${C.money(r.amount - r.left)} of ${C.money(r.amount)} paid since` : esc(r.by ? `Marked by ${r.by}` : "")}</span></span>
      <span class="w"><button class="btn mute sm" data-act="payMissed" data-m="${m}" data-id="${esc(r.id)}">Collected</button> <button class="btn mute sm" data-act="payWaive" data-m="${m}" data-id="${esc(r.id)}">Waive</button></span></li>`).join("");
    const cf = conf.slice(0, 8).map(a => `<li class="step now"><span class="k"></span><span><span class="t">${dshort(a.date)} ${C.fmtT(a.time)}: ${C.money(due.total)} due</span><span class="m">Nobody said whether it was collected</span></span>
      <span class="w"><button class="btn mute sm" data-act="payConf" data-v="collected" data-m="${m}" data-d="${a.date}" data-t="${a.time}">Collected</button> <button class="btn mute sm" data-act="payConf" data-v="missed" data-m="${m}" data-d="${a.date}" data-t="${a.time}">Not collected</button></span></li>`).join("");
    const who = x => x.cover === "selfpay" ? "self pay" : x.payer ? esc(x.payer) : "";
    return `<div class="h"><h2>Financial</h2><span${L.balance > 0 ? ' class="flagt"' : ""}>${L.balance > 0 ? `Balance ${C.money(L.balance)}` : "No balance"}</span></div>
      ${needTable("tblIvkPay", "Payments")}
      <dl class="facts"><dt>Provider visit</dt><dd>${due.copay ? C.money(due.copay) : "Nothing"}${who(cv.prov) ? `, ${who(cv.prov)}` : ""}</dd><dt>Infusion</dt><dd>${due.infusion ? C.money(due.infusion) : "Nothing"}${who(cv.inf) ? `, ${who(cv.inf)}` : ""}</dd>
        <dt>Due each visit</dt><dd>${due.total ? C.money(due.total) : "Nothing"}</dd>
        <dt>Collected</dt><dd>${C.money(L.collected) || "$0"}</dd>
        <dt>Last reminder</dt><dd${remDue ? ' class="flag"' : ""}>${last ? `${C.fmtD(last.date)}${last.method ? `, ${esc(last.method)}` : ""}${last.by ? `, ${esc(last.by)}` : ""}` : L.balance > 0 ? "None sent" : "—"}</dd></dl>
      ${chk("CardOnFile", t?.card, "Luma card on file")}${chk("OkToCollect", t?.okCollect, "Okay to charge at time of service", "The patient agreed we can charge the card on file at each visit")}
      ${miss || cf ? `<ul class="steps">${miss}${cf}</ul>${conf.length > 8 ? `<p class="note">And ${plural(conf.length - 8, "older visit")} to confirm.</p>` : ""}` : ""}
      <div class="actions"><button class="btn ghost sm line" data-act="pay" data-m="${m}"${S.payOn ? "" : " disabled"}>Record a payment or reminder</button><button class="btn mute sm" data-act="profile" data-m="${m}">Change what's due</button></div>`;
  }
  const AV_CLS = { P: "pref", M: "on", S: "some", N: "off" };
  function availGridHTML(a) {
    const g = a || C.DAYS5.map(() => ({ am: "", pm: "" }));
    return `<div class="avgrid"><span></span><b>AM</b><b>PM</b>${C.DAYS5.map((d, i) => `<b>${d}</b>${["am", "pm"].map(h => `<span class="lv ${AV_CLS[g[i][h]] || ""}">${esc(g[i][h] ? C.avName(g[i][h]).replace(" available", "") : "—")}</span>`).join("")}`).join("")}</div>`;
  }
  function availHTML(p) {
    const t = ptOf(p.mrn), m = esc(p.mrn), T = todayS(), a = t?.avail || null, away = (t?.away || []).filter(w => w.to >= T);
    const fromTracker = !a && p.prefs.some(x => x.am || x.pm) ? p.prefs.map(x => ({ am: x.am ? "M" : "", pm: x.pm ? "M" : "" })) : null;
    const st = a ? (t.availExp != null && t.availExp < T ? `<span class="flagt">expired ${C.fmtD(t.availExp)}</span>` : `${t.availEff != null ? `from ${C.fmtD(t.availEff)}` : "on file"}${t.availExp != null ? ` to ${C.fmtD(t.availExp)}` : ""}`) : fromTracker ? "AM/PM from the tracker" : "not on file";
    return `<div class="section"><div class="h"><h2>Availability</h2><span>${st}</span></div>
      ${availGridHTML(a || fromTracker)}
      ${t?.availNote ? `<p class="note">${esc(t.availNote)}</p>` : ""}
      ${away.length ? `<p class="note">Away: ${away.map(w => `${C.fmtD(w.from)}${w.to !== w.from ? ` to ${C.fmtD(w.to)}` : ""}${w.note ? ` (${esc(w.note)})` : ""}`).join(", ")}</p>` : ""}
      <div class="actions"><button class="btn ghost sm line" data-act="avail" data-m="${m}">Edit availability</button><button class="btn mute sm" data-act="printAvail" data-m="${m}">Print availability form</button></div></div>`;
  }
  /* ---------- Availability ---------- */
  /* Everyone's availability as the IVK Patient Availability Form asks it: each weekday morning (about 9 to 10 AM) and afternoon
     (about 12:30 to 3 PM) is Preferred, Mostly available, Sometimes available or Not available, with effective and expiry dates.
     The form asks for at least three days a week. Above the table, how many patients can come each half day. */
  const AVCH = [["active", "Active"], ["update", "Needs updating"], ["all", "All"]];
  const availOfPt = p => { const t = ptOf(p.mrn); return t?.avail ? { g: t.avail, src: "form", t } : p.prefs.some(x => x.am || x.pm) ? { g: p.prefs.map(x => ({ am: x.am ? "M" : "", pm: x.pm ? "M" : "" })), src: "tracker", t } : { g: null, src: "", t }; };
  function availPageHTML() {
    const T = todayS(), k = S.avChip || "active";
    const L = S.P.filter(p => !/discontinued/i.test(p.status)).map(p => { const a = availOfPt(p), t = a.t, exp = t?.availExp != null && t.availExp < T, days = a.g ? a.g.filter(x => C.canCome(x.am) || C.canCome(x.pm)).length : 0;
      return { p, a, exp, days, upd: !a.g || exp || a.src === "tracker" || days < 3 }; });
    const pick = r => k === "all" || (k === "active" ? C.isActive(r.p) : r.upd && C.isActive(r.p)), shown = L.filter(pick).sort((x, y) => (y.upd - x.upd) || x.p.code.localeCompare(y.p.code));
    const n = Object.fromEntries(AVCH.map(([c]) => [c, L.filter(r => c === "all" || (c === "active" ? C.isActive(r.p) : r.upd && C.isActive(r.p))).length]));
    const half = [].concat(...C.DAYS5.map((d, i) => [[i, "am"], [i, "pm"]]));
    const can = half.map(([i, h]) => shown.filter(r => r.a.g && C.canCome(r.a.g[i][h])).length), pref = half.map(([i, h]) => shown.filter(r => r.a.g && r.a.g[i][h] === "P").length);
    const lv = v => `<span class="lv ${AV_CLS[v] || ""}" title="${esc(v ? C.avName(v) : "Not set")}">${v ? esc(v) : "—"}</span>`;
    const row = r => { const t = r.a.t, away = (t?.away || []).filter(w => w.to >= T);
      return `<button class="avrow" data-act="avail" data-m="${esc(r.p.mrn)}"><span class="p">${esc(r.p.code)}${mrnTag(r.p.mrn)}</span>${half.map(([i, h]) => lv(r.a.g ? r.a.g[i][h] : "")).join("")}
        <span class="small${r.exp ? " flagt" : ""}">${t?.availEff != null ? C.fmtD(t.availEff) : r.a.src === "tracker" ? "Tracker" : "—"}${t?.availExp != null ? ` to ${C.fmtD(t.availExp)}` : ""}</span>
        <span class="small">${!r.a.g ? `<span class="f">No form on file</span>` : r.exp ? `<span class="f">Expired</span>` : r.days < 3 ? `<span class="f">${plural(r.days, "day")}, form asks for 3</span>` : r.a.src === "tracker" ? "AM/PM from the tracker" : ""}${away.length ? ` Away ${away.map(w => C.fmtD(w.from)).join(", ")}` : ""}</span></button>`; };
    return `${setupBanner()}${head("Availability", n.active ? `${n.update} of ${plural(n.active, "active patient")} need${n.update === 1 ? "s" : ""} their availability updated` : "No active patients", false)}
      <div class="chips">${AVCH.map(([c, l]) => `<button class="chip" data-act="avChip" data-v="${c}" aria-pressed="${k === c}">${l}<span class="n">${n[c]}</span></button>`).join("")}</div>
      ${needTable("tblIvkPatient", "Availability forms")}
      <section class="card avtable"><div class="avrow head"><span>Patient</span>${C.DAYS5.map(d => `<span class="dh2">${d}<small>AM</small></span><span class="dh2">${d}<small>PM</small></span>`).join("")}<span>Effective</span><span>Notes</span></div>
        <div class="avrow sum" aria-label="Patients who can come each half day"><span class="small muted">Can come<br>(preferred)</span>${can.map((c, j) => `<span class="cnt">${c}<small>${pref[j]}</small></span>`).join("")}<span></span><span></span></div>
        ${shown.map(row).join("") || `<p class="empty" style="padding:10px 14px">Nobody in this group.</p>`}</section>
      <ul class="key"><li>${lv("P")}Preferred</li><li>${lv("M")}Mostly available</li><li>${lv("S")}Sometimes available</li><li>${lv("N")}Not available</li><li class="muted">Preferred and Mostly go to the tracker’s AM/PM columns and the Index planner. AM is about 9 to 10 AM, PM about 12:30 to 3 PM.</li></ul>`;
  }
  /* Oakland patients: their index here, then maintenance back at Oakland; OAK is tasked to book it at treatment OakRemindAt */
  function oakHTML(p) {
    const t = ptOf(p.mrn), o = C.oakState(p, t, S.A, todayS(), S.cfg), m = esc(p.mrn), auto = t?.oak == null && /oak/i.test(p.loc);
    const note = !o ? "" : o.unknown && o.tasked == null ? "This course started before the appointment list begins, so treatments can't be counted here. Mark OAK tasked once maintenance is booked at Oakland." : o.tasked != null ? `OAK tasked to book maintenance ${C.fmtD(o.tasked)}.${thereTxt(C.bookedAt(S.x.E, p.mrn, "Oakland", todayS()), "Oakland")}` : o.done >= +S.cfg.OakRemindAt ? `<b>Task OAK to book maintenance now.</b>` : `At treatment ${S.cfg.OakRemindAt}, task OAK to book maintenance.`;
    return `<div class="section"><div class="h"><h2>Oakland</h2><span>${o ? (o.unknown ? "treatments not countable" : `${o.done} of ${o.idx} index treatments done`) : "Castro patient"}</span></div>
      <label class="opt"><input type="checkbox" data-pt="Oakland" data-m="${m}"${o ? " checked" : ""}${S.raw.tables?.own?.tblIvkPatient ? "" : " disabled"}> Oakland patient<small>Index (the first ${esc(S.cfg.OakIndex)}) at ${esc(S.cfg.Clinic)}, then back to Oakland for maintenance.${auto ? " Ticked because the home clinic is Oakland." : ""}</small></label>
      ${o ? `<p class="note${o.task || o.back ? " warn" : ""}">${note}${o.back ? ` Index is done and ${plural(o.after.length, "visit")} ${o.after.length === 1 ? "is" : "are"} still booked here.` : ""}</p>
        <div class="actions">${o.tasked == null ? `<button class="btn ghost sm line" data-act="oakTask" data-m="${m}">Mark OAK tasked</button>` : `<button class="btn mute sm" data-act="oakUntask" data-m="${m}">Undo OAK tasked</button>`}</div>` : ""}</div>`;
  }

  /* ---------- Consults ---------- */
  /* Waitlist: everyone waiting for a treatment start, in line by consult date or the order set here; the Consults tab beneath it */
  const WAIT_TXT = { ready: "Ready to start", referral: "Waiting on the referral", auth: "Waiting on the authorization" };
  function consultsHTML() {
    const ORD = { cm: 0, act: 1, wait: 2, done: 3, started: 4, closed: 5 };
    const L = S.CS.slice().sort((a, b) => ORD[a.state] - ORD[b.state] || (b.days ?? 0) - (a.days ?? 0));
    const n = s => L.filter(c => c.state === s).length, W = S.WL, ok = !!S.raw.tables?.own?.tblIvkPatient, planned = new Set(ipList().filter(o => o.pl || o.holds.length).map(o => o.mrn));
    const sub = [plural(W.list.length, "patient") + " on the waitlist", W.next ? `${W.next.code} starts next` : W.list.length ? "nobody ready to start yet" : "",
      n("done") ? `${n("done")} scheduled` : "", n("cm") ? `${plural(n("cm"), "start")} to update in Case Management` : ""].filter(Boolean).join(", ");
    if (!S.wSel || !W.list.find(w => w.mrn === S.wSel)) S.wSel = W.list[0]?.mrn || null;
    const wrow = (w, i) => { const c = w.c, why = w.next ? `<span class="ok">Starts next</span>` : w.on === "referral" ? `<span class="f">Referral: keeps their place</span>`
        : w.on === "auth" ? `<span class="f">${w.appeal ? "Appeal: the next ready patient starts first" : "Authorization pending"}</span>` : w.ahead ? "" : "Ready";
      return `<div class="wrow${S.wSel === w.mrn ? " cur" : ""}" role="button" tabindex="0" data-act="wSel" data-m="${esc(w.mrn)}" aria-pressed="${S.wSel === w.mrn}"><span class="pos">${w.pos}</span><span class="p">${esc(w.code || "—")}${mrnTag(w.mrn)}</span>
        <span class="proto">${c.consult != null ? `Consult ${C.fmtD(c.consult)}` : "No consult date"}<br><span class="muted small">${esc(c.provider || "")}</span></span>
        <span class="thru">${esc(WAIT_TXT[w.on])}<br><span class="muted small">${esc(c.stage || "")}</span></span><span class="flags">${why}${planned.has(w.mrn) ? ` <span class="muted small">planning</span>` : ""}</span>
        <span class="mv"><button class="btn mute sm" data-act="wMove" data-m="${esc(w.mrn)}" data-v="-1" aria-label="Move ${esc(w.code)} up"${i === 0 || !ok ? " disabled" : ""}>↑</button><button class="btn mute sm" data-act="wMove" data-m="${esc(w.mrn)}" data-v="1" aria-label="Move ${esc(w.code)} down"${i === W.list.length - 1 || !ok ? " disabled" : ""}>↓</button></span></div>`; };
    const cur = W.list.find(w => w.mrn === S.wSel);
    const row = c => `<div class="crow" role="button" tabindex="0" data-act="editConsult" data-m="${esc(c.mrn)}"${c.mrn ? "" : ' style="cursor:default"'}><span class="p">${esc(c.code || "—")}</span><span class="proto">${c.consult != null ? `Consult ${C.fmtD(c.consult)}` : "No consult date"}<br>${esc(c.provider || "")}</span><span class="thru">${esc(c.stage || "No stage")}${c.days != null ? `<br><span class="muted small">${plural(c.days, "day")} in stage</span>` : ""}</span>
      <span class="flags">${c.flag ? `<span class="f">${esc(c.flag)}</span>` : c.state === "started" ? `<span class="ok">Started ${C.fmtD(c.kept)}</span>` : c.state === "done" ? `<span class="ok">Scheduled</span>` : c.state === "closed" ? "Not proceeding" : "Waiting"}${c.state === "cm" ? ` <button class="btn ghost sm" data-act="cmDone" data-m="${esc(c.mrn)}">Case Management updated</button>` : ""}</span></div>`;
    return `${setupBanner()}${head("Waitlist", sub, false)}${accessChips()}${!S.raw.tables?.consults ? `<div class="msg warn">There’s no table on a Consults tab in this workbook.</div>` : ""}
      <div class="cols"><section class="card plist wl"><div class="chead"><span>#</span><span>Patient</span><span>Consult</span><span>Waiting on</span><span></span><span></span></div>
        ${W.list.map(wrow).join("") || `<p class="empty" style="padding:10px 14px">Nobody is waiting for a treatment start.</p>`}
        <p class="note" style="padding:0 14px 12px">In line by consult date until someone is moved. Waiting on a referral keeps their place. Waiting on an authorization that needs an appeal moves them behind the next ready patient.</p></section>
      <div>${cur ? waitCardHTML(cur) : ""}</div></div>
      <h2 class="sech">Consults</h2>
      <section class="card plist"><div class="chead"><span>Patient</span><span>Consult</span><span>Stage</span><span>Next step</span></div>${L.map(row).join("") || `<p class="empty" style="padding:10px 14px">Nothing on the Consults tab.</p>`}</section>`;
  }
  function waitCardHTML(w) {
    const c = w.c, m = esc(w.mrn), ok = !!S.raw.tables?.own?.tblIvkPatient, p = S.PM.get(w.mrn);
    const lead = w.next ? { cls: "ready", t: "Next to start. Plan their index visits." } : w.on === "referral" ? { cls: "soon", t: "Waiting on the referral. They keep their place in line." }
      : w.on === "auth" ? { cls: "soon", t: w.appeal ? "The authorization needs an appeal, so the next ready patient starts first." : "Waiting on the authorization." } : { cls: "", t: `Ready, number ${w.pos} in line` };
    return `<aside class="card detail" aria-label="${esc(w.code)}"><div class="dh"><div><div class="name">${esc(w.code)}<span class="mrn">MRN ${esc(w.mrn)}</span><span class="loc">#${w.pos}</span></div><div class="meta">${esc([c.loc, c.provider, c.consult != null ? `consult ${C.fmtD(c.consult)}` : ""].filter(Boolean).join(", "))}</div></div></div>
      <p class="lead ${lead.cls}">${esc(lead.t)}</p>
      <div class="fld">Waiting on<div class="pick">${C.WAIT_ON.map(([k, l]) => `<button data-act="wOn" data-m="${m}" data-v="${k}" aria-pressed="${w.on === k}"${ok ? "" : " disabled"}>${l}</button>`).join("")}</div></div>
      ${w.on === "auth" ? `<label class="opt"><input type="checkbox" data-wappeal="1" data-m="${m}"${w.appeal ? " checked" : ""}${ok ? "" : " disabled"}> An appeal is needed<small>Moves them behind the next ready patient, who usually starts first</small></label>` : ""}
      <label class="fld">Note<textarea id="wNote" rows="2" data-m="${m}">${esc(w.t?.waitNote || "")}</textarea></label>
      <button class="btn primary" data-act="wPlan" data-m="${m}">Plan index visits</button>
      <div class="actions"><button class="btn ghost" data-act="wMove" data-m="${m}" data-v="top"${w.pos === 1 || !ok ? " disabled" : ""}>Move to the top</button><button class="btn ghost" data-act="editConsult" data-m="${m}">Edit consult</button>${p ? `<button class="btn ghost" data-act="openPt" data-m="${m}">Patient</button>` : ""}</div>
      ${needTable("tblIvkPatient", "Reordering and waiting-on")}
      <dl class="facts"><dt>Referred</dt><dd>${c.ref != null ? C.fmtD(c.ref) : "—"}</dd><dt>Stage</dt><dd>${esc(c.stage || "—")}</dd><dt>In stage</dt><dd>${c.days != null ? plural(c.days, "day") : "—"}</dd>${c.start != null ? `<dt>Start date</dt><dd>${C.fmtD(c.start)}</dd>` : ""}</dl></aside>`;
  }

  /* ---------- Settings ---------- */
  function settingsHTML() {
    const r = S.raw, miss = r.missing || [], c = S.cfg, u = r.apptUpdated;
    const route = (ev, title, hint) => { const x = S.routes[ev] || { to: "", url: "", on: true }, ok = N.flowUrl(x.url);
      return `<div class="section"><div class="h"><h2>${title}</h2><span>${ok ? (x.on ? "on" : "off") : "not set up"}</span></div><p>${hint}</p>
        <div class="two"><label class="fld">Sends to<input id="rt_${ev}_to" value="${esc(x.to)}" autocomplete="off"></label>
        <label class="fld">Workflow URL<input id="rt_${ev}_url" type="password" value="${esc(x.url)}" autocomplete="off" spellcheck="false"><span class="hint${x.url && !ok ? " warn" : ""}">${x.url && !ok ? "That isn’t a Power Automate workflow URL." : "Kept in the workbook. Treat it like a password."}</span></label></div>
        <label class="opt"><input type="checkbox" id="rt_${ev}_on"${x.on ? " checked" : ""}> Post ${ev === "arrival" ? "arrivals" : "asks"} to this chat</label>
        <div class="actions"><button class="btn ghost sm line" data-act="testRoute" data-v="${ev}">Send test</button></div></div>`; };
    const tplF = (k, l) => `<label class="fld">${l}<textarea id="tp_${k}" rows="2">${esc(S.tpl[k])}</textarea></label>`;
    const P = S.plan;
    return `${head("Settings", "", false)}${setupBanner()}<div class="sgrid2">
      <section class="card set"><h3>You</h3><p>Your first name goes on Teams posts and in the log. It’s kept on this computer.</p>
        <label class="fld">First name<input id="myName" value="${esc(S.name)}" autocomplete="given-name"></label><div class="actions"><button class="btn ghost sm line" data-act="saveName">Save name</button></div></section>

      <section class="card set"><h3>Appointment list</h3><p>${u != null ? `Last updated ${C.fmtD(Math.floor(u))} at ${C.fmtT(C.toMin(u))}.` : "No record of when it was last updated."} ${plural(S.A.length, "appointment")} on the Appt List tab.</p>
        <div class="actions"><label class="btn primary inl" style="cursor:pointer">Upload a NextGen export<input type="file" id="apptFile" accept=".xls,.xlsx,.xlsm,.xlsb,.csv,.txt,.xml,.htm,.html,.ods" hidden></label></div>
        <details class="fold"><summary>How to run the report</summary><ol class="how">
          <li>In NextGen PM, open Reports, then Memorized.</li><li>Set Report Type to Appointment Reports and pick the IVK appointment list (Castro Street, IV Ketamine visits).</li>
          <li>Select OK to run it, then export with the x icon beside the bar graph <img src="assets/export-icon.png" alt="Export icon" height="16">.</li>
          <li>Choose that file here. It replaces MRN to Location on the Appt List tab; the Name and Lookup Key formulas are left alone.</li>
          <li>The Index Planner tab is then filled with every IVK visit and hold, as plain values.</li></ol></details></section>

      <section class="card set"><h3>Teams</h3><p>Each chat gets its own workflow made from Teams’ <b>Send webhook alerts to a chat</b> template. Patients are named by code only.</p>
        <label class="opt"><input type="checkbox" id="nfyOn"${N.truthy(c.Enabled) ? " checked" : ""}> Post to Teams from this workbook</label>
        ${route("arrival", "Arrivals", "When the desk checks a patient in, and when they’re seated. The RN takes vitals, so this goes to the RN’s chat.")}
        ${route("protocol", "Protocol and scheduling", "Asks the providers to confirm the protocol and send the patient to the desk at discharge.")}
        <div class="actions"><button class="btn primary inl" data-act="saveRoutes">Save Teams settings</button></div>
        <details class="fold"><summary>Set up a workflow</summary><ol class="how">
          <li>Open the chat in Teams (the Castro IVK group chat, or the RN’s chat).</li><li>Choose More options (…), then Workflows, then <b>Send webhook alerts to a chat</b>.</li>
          <li>Name it “IVK Tracker - arrivals” or “IVK Tracker - group chat”, confirm the chat and finish.</li><li>Copy the URL Teams shows and paste it above. Don’t paste it anywhere else.</li>
          <li>Save, then Send test. “Accepted (202)” means the card is on its way to the chat.</li></ol></details></section>

      <section class="card set"><h3>Messages</h3><p>{patient} is the code, {time} the visit time, {provider} the provider, {chair} the chair, {protocol} the frequency (“every 2 week”).</p>
        ${tplF("arrival", "Arrival")}${tplF("seated", "Chair picked at the start of treatment")}${tplF("protocol", "Confirm protocol and book")}${tplF("protocolOnly", "Confirm protocol only")}${tplF("bookOnly", "Book only")}
        <label class="fld">Start email subject<input id="tp_startSubject" value="${esc(S.tpl.startSubject)}"></label>
        <label class="fld">Start email (HTML)<textarea id="tp_startEmail" rows="9" spellcheck="false">${esc(S.tpl.startEmail)}</textarea><span class="hint">The Castro desk’s “Upcoming Appointment(s)” email. {mrn}, {first} (first name), {clinic}, {firstVisit} (“Wednesday, November 4th, 2026, at 10:00 AM”), {address}, and {schedule} where the appointment table goes.</span></label>
        <label class="fld">Clinic address in the start email<input id="cf_ClinicAddress" value="${esc(S.cfg.ClinicAddress)}"></label>
        <div class="two"><label class="fld">CC on every start email<input id="cf_StartCc" value="${esc(S.cfg.StartCc)}" autocomplete="off"></label><label class="fld">BCC on every start email<input id="cf_StartBcc" value="${esc(S.cfg.StartBcc)}" autocomplete="off"></label></div>
        <div class="actions"><button class="btn ghost sm line" data-act="saveTpl">Save messages</button><button class="btn mute sm" data-act="resetTpl">Use the standard wording</button></div></section>

      <section class="card set"><h3>Chair map</h3><p>A file from the <a href="${BUILDER}" target="_blank" rel="noopener">Bay Map Builder</a>. Until a map is loaded, visits are tracked without chairs and Today shows no chair map. Bay numbers become chair numbers.</p>
        <div class="status${P ? " ok" : ""}"><span class="dot"></span>${P ? `${esc(P.name || "Map")} · ${plural(P.bays.length, "chair")} (${P.bays.map(b => esc(b.id)).join(", ")})` : "No map loaded. Visits are tracked without chairs."}</div>
        <div class="actions"><label class="btn ghost sm line" style="cursor:pointer">Load a map file<input type="file" id="mapFile" accept=".json,application/json" hidden></label>${P ? `<button class="btn ghost sm line" data-act="mapDown">Download this map</button><button class="btn mute sm" data-act="mapRemove">Remove the map</button>` : ""}</div></section>

      <section class="card set"><h3>Timing</h3>
        <div class="two"><label class="fld">Visit length, min<input id="cf_VisitMin" inputmode="numeric" value="${esc(c.VisitMin)}"></label>
        <label class="fld">Chairs in the room<input id="cf_Chairs" inputmode="numeric" value="${esc(C.chairCount(c, P))}"${P ? " disabled" : ""}><span class="hint">${P ? "From the chair map" : "For open slots on the Schedule"}</span></label>
        <label class="fld">Ask to book when booked less than, weeks<input id="cf_AheadWeeks" inputmode="numeric" value="${esc(c.AheadWeeks)}"></label><label class="fld">Confirm protocol every, days<input id="cf_ProtocolDays" inputmode="numeric" value="${esc(c.ProtocolDays)}"></label>
        <label class="fld">Auth ending within, days<input id="cf_AuthWarnDays" inputmode="numeric" value="${esc(c.AuthWarnDays)}"></label><label class="fld">Visits left on auth, at most<input id="cf_EncountersWarn" inputmode="numeric" value="${esc(c.EncountersWarn)}"></label></div>
        <label class="fld">Clinic<input id="cf_Clinic" value="${esc(c.Clinic)}"></label>
        <div class="actions"><button class="btn ghost sm line" data-act="saveCfg">Save timing</button></div></section>

      <section class="card set"><h3>Index protocol</h3><p>The standard for new starts. Week 1 begins on the day of treatment 1. A patient’s own protocol can be changed in the Index planner.</p>
        <div class="two"><label class="fld">Visits a week<input id="cf_IndexPerWeek" inputmode="numeric" value="${esc(c.IndexPerWeek)}"></label><label class="fld">Weeks<input id="cf_IndexWeeks" inputmode="numeric" value="${esc(c.IndexWeeks)}"></label>
        <label class="fld">Days apart, at least<input id="cf_MinGapDays" inputmode="numeric" value="${esc(c.MinGapDays)}"><span class="hint">2 = a day off between visits</span></label><label class="fld">Days apart, at most<input id="cf_MaxGapDays" inputmode="numeric" value="${esc(c.MaxGapDays)}"><span class="hint">Broken only when no slot fits</span></label>
        <label class="fld">Post check after the start, min<input id="cf_PostCheckMin" inputmode="numeric" value="${esc(c.PostCheckMin)}"></label><label class="fld">Confirm a hold within, min<input id="cf_HoldMin" inputmode="numeric" value="${esc(c.HoldMin)}"></label>
        <label class="fld">Count room to start within, weeks<input id="cf_CapacityWeeks" inputmode="numeric" value="${esc(c.CapacityWeeks)}"></label></div>
        <label class="fld">After index<select id="cf_AfterIndex">${[...new Set([c.AfterIndex, ...S.freqs])].map(f => `<option${f === c.AfterIndex ? " selected" : ""}>${esc(f)}</option>`).join("")}</select></label>
        <div class="actions"><button class="btn ghost sm line" data-act="saveIdx">Save index protocol</button></div></section>

      <section class="card set"><h3>Payments and Oakland</h3><p>Each IVK visit owes the tracker’s Copay plus Infusion. Check-out asks whether it was collected; what isn’t becomes the balance.</p>
        <div class="two"><label class="fld">Balance reminder every, days<input id="cf_ReminderDays" inputmode="numeric" value="${esc(c.ReminderDays)}"></label>
        <label class="fld">Ask about payments from<input id="cf_PayStart" type="date" value="${esc(C.isoOf(S.payStart))}"><span class="hint">Visits before this day are never asked about</span></label>
        <label class="fld">Oakland index, treatments here<input id="cf_OakIndex" inputmode="numeric" value="${esc(c.OakIndex)}"></label><label class="fld">Task OAK at treatment<input id="cf_OakRemindAt" inputmode="numeric" value="${esc(c.OakRemindAt)}"></label></div>
        <div class="actions"><button class="btn ghost sm line" data-act="savePayCfg">Save</button></div></section>

      <section class="card set"><h3>Appearance</h3><label class="opt"><input type="checkbox" id="darkOn"${S.theme === "dark" ? " checked" : ""}> Dark mode<small>On this computer only</small></label>
        <label class="opt"><input type="checkbox" id="followOn"${LS.get("ivkFollow", "1") === "1" ? " checked" : ""}> Follow along in the workbook<small>Picking a patient or a visit selects their row on the Master Patient Tracker behind the add-in. On this computer only</small></label></section>

      <section class="card set"><h3>Workbook</h3>
        ${["mpt", "appt", "consults"].map(k => `<div class="status ${r.tables?.[k] ? "ok" : "warn"}"><span class="dot"></span>${{ mpt: "Master Patient Tracker", appt: "Appt List", consults: "Consults" }[k]} ${r.tables?.[k] ? "found" : "not found"}</div>`).join("")}
        <div class="status ${miss.length ? "warn" : "ok"}"><span class="dot"></span>${miss.length ? `Add-in tables missing: ${miss.join(", ")}` : `Add-in tables on the hidden ${esc(X.DATA_SHEET)} tab`}</div>
        ${miss.length ? `<div class="actions"><button class="btn primary inl" data-act="setup">Set up this workbook</button></div>` : ""}</section>
      </div>`;
  }

  /* ---------- modals ---------- */
  const chairPick = (cur, allowNone) => {
    const F = floor();
    return `<div class="pick" role="group" aria-label="Chair">${F.chairs.map(c => { const busy = !!c.v && c.v.key !== S.m.key;
      return `<button data-act="mPick" data-b="${esc(c.id)}" aria-pressed="${cur === c.id}"${busy ? " disabled" : ""}>${esc(chairTxt(c.id))}<small>${busy ? esc(C.code(c.v)) : "Open"}</small></button>`; }).join("")}${allowNone ? `<button data-act="mPick" data-b="" aria-pressed="${!cur}">No chair yet<small>Pick at start</small></button>` : ""}</div>`;
  };
  /* the patient's own frequency first; with none on file, a blank choice is shown first so nothing is written by default */
  const freqOpts = (p, cur) => (p.freq ? "" : `<option value=""${cur ? "" : " selected"}>Choose the protocol</option>`)
    + [...new Set([p.freq, ...S.freqs].filter(Boolean))].map(f => `<option${f === cur ? " selected" : ""}>${esc(f)}</option>`).join("");
  /* today's payment, the provider visit and the infusion each on its own: what the patient owes today (the tracker's amount unless
     changed) and whether it was collected. Asked at check-in (optional), at check-out (each part that owes something and has no
     answer yet), and changed any time after, check-out included: a new answer is a new row and the latest one counts. */
  const PART_WORD = { prov: "Provider visit", inf: "Infusion" }, PARTS = ["prov", "inf"];
  const trackerDue = (p, k) => k === "prov" ? C.dueOf(p).copay : C.dueOf(p).infusion;
  const amtTxt = s => String(s ?? "").replace(/[$,\s]/g, "");
  const amtN = (m, k) => { const n = C.num(amtTxt(m.amt?.[k])); return n == null ? 0 : n; };
  const amtBad = m => PARTS.some(k => amtTxt(m.amt?.[k]) !== "" && C.num(amtTxt(m.amt[k])) == null);
  /* the modal's starting point: each part's last answer for this visit, else the tracker's amount and no answer */
  function payInit(v) {
    if (!v.p || !S.payOn) return {};
    const amt = {}, pay = {};
    for (const k of PARTS) { const r = C.payOf(S.pays, v.key, k), own = r && r.key === C.partKey(v.key, k); amt[k] = String(own ? r.amount : trackerDue(v.p, k)); if (r) pay[k] = r.kind; }
    return { amt, pay };
  }
  /* check-out can't finish until every part that owes something today has an answer */
  const payAnswered = (v, m) => !v.p || !S.payOn || (!amtBad(m) && PARTS.every(k => amtN(m, k) <= 0 || !!(m.pay && m.pay[k])));
  /* rows to add: only parts whose answer or amount changed. Owing nothing today counts as collected ($0), so it's answered. */
  function payRowsOf(v, m) {
    if (!v.p || !S.payOn) return [];
    return PARTS.map(k => {
      const n = amtN(m, k), r = C.payOf(S.pays, v.key, k), due = trackerDue(v.p, k), kind = n > 0 ? (m.pay?.[k] || "") : (r || due > 0 ? "collected" : "");
      if (!kind || (r && r.kind === kind && r.amount === n)) return null;
      return payRow(v.p, { Date: C.isoOf(v.a.date), Key: C.partKey(v.key, k), Kind: kind, Amount: n, Note: `${PART_WORD[k]}${n !== due ? `, ${C.money(n)} today (tracker ${C.money(due)})` : ""}${r ? ", changed" : ""}` });
    }).filter(Boolean);
  }
  /* mode: "in" (check-in: Collected or Not yet), "out" (check-out) or "edit" (any time): Collected or Not collected */
  function payPickHTML(v, m, mode) {
    if (!v.p || !S.payOn) return "";
    const t = ptOf(v.p.mrn), L = C.ledger(v.p.mrn, S.pays), cur = k => m.pay?.[k] || "";
    const pk = (k, kind, l, s) => `<button data-act="mPay" data-p="${k}" data-v="${kind}" aria-pressed="${cur(k) === kind}">${l}<small>${s}</small></button>`;
    const part = k => { const n = amtN(m, k), due = trackerDue(v.p, k), r = C.payOf(S.pays, v.key, k);
      return `<div class="fld payp">${PART_WORD[k]} today${r && mode !== "edit" ? ` <span class="muted">(${r.kind === "collected" ? "collected" : "not collected"}, ${C.money(r.amount)})</span>` : ""}
        <div class="payr"><label class="amt">Patient owes<input id="mAmt_${k}" inputmode="decimal" value="${esc(m.amt?.[k] ?? "")}" autocomplete="off"></label>
          ${n > 0 ? `<div class="pick">${pk(k, "collected", "Collected", C.money(n))}${mode === "in" ? pk(k, "", "Not yet", "ask at check-out") : pk(k, "missed", "Not collected", "adds to the balance")}</div>` : `<span class="hint">Nothing due</span>`}</div>
        ${n !== due ? `<span class="hint">The tracker has ${due ? C.money(due) : "nothing"} each visit</span>` : ""}</div>`; };
    const changed = PARTS.some(k => amtN(m, k) !== trackerDue(v.p, k)), anyMissed = PARTS.some(k => cur(k) === "missed" && amtN(m, k) > 0);
    return `<div class="section"><div class="h"><h2>Payment</h2>${mode === "in" ? "<span>optional at check-in</span>" : ""}</div>${PARTS.map(part).join("")}
      ${amtBad(m) ? `<p class="note warn">What the patient owes has to be a number, like 40.</p>` : ""}
      ${changed ? `<label class="opt"><input type="checkbox" id="mAmtSave"${m.amtSave ? " checked" : ""}> Make this what they owe each visit<small>Writes Copay and Infusion on the Master Patient Tracker. Leave it unticked to change today only.</small></label>` : ""}
      ${L.balance > 0 ? `<p class="note warn">Balance from earlier visits: ${C.money(L.balance)}.</p>` : ""}
      ${anyMissed ? `<p class="note">${t?.card && t?.okCollect ? "Luma card on file and okay to charge at time of service: charge it in Luma, then record the payment on the patient’s Financial page." : t?.card ? "Luma card on file, but not okay to charge at time of service yet. Ask the patient." : "No Luma card on file. Ask for one, or send a balance reminder."}</p>` : ""}</div>`;
  }
  /* the modal's inputs back into S.m before a re-render, so nothing typed is lost */
  function keepPay() {
    const m = S.m; if (!m) return;
    for (const k of PARTS) { const el = $(`#mAmt_${k}`); if (el) m.amt = Object.assign({}, m.amt, { [k]: el.value }); }
    const as = $("#mAmtSave"); if (as) m.amtSave = as.checked;
    const pr = $("#mProto"); if (pr) m.proto = pr.checked;
    const fq = $("#mFreq"); if (fq) m.freq = fq.value;
    const rn = $("#mRn"); if (rn) m.rn = rn.checked;
    const ak = $("#mAsk"); if (ak) m.ask = ak.checked;
    const xc = $("#mXfer"); if (xc) m.xferOn = xc.checked;
  }
  /* save today's payment answers (and, when ticked, the new amounts on the tracker). label for the toast; false when nothing saved */
  async function savePay(v, m, what) {
    const rows = payRowsOf(v, m), track = m.amtSave && PARTS.some(k => amtN(m, k) !== trackerDue(v.p, k));
    if (!rows.length && !track) return true;
    const ok = await write("Saving the payment…", async () => {
      if (rows.length) await Store.addPays(rows);
      if (track) { const H = mptHead(); return Store.saveRow("mpt", v.p.mrn, Object.fromEntries([[H("copay") || "Copay", amtN(m, "prov")], [H("infusion") || "Infusion", amtN(m, "inf")]])); }
      return true;
    });
    if (!ok && what) toast(`${what}, but the payment didn’t save. Use Record today’s payment on the visit card.`, true);
    return !!ok;
  }
  /* the edit form's fields from the row: label, current value as the input shows it, type, and suggestions from the column */
  /* the locations in the appointment list's scope, as NextGen names them: this clinic's and the other clinics' (wider report) */
  const locOpts = () => [...new Set(S.A.map(a => a.loc).concat(S.E.map(a => a.loc)).map(C.txt).filter(Boolean))].sort();
  /* the AM/PM columns are edited under Availability (found by position, C.prefCols) */
  const SKIP_EDIT = { mpt: ["mrn", "patientname"], consults: ["patientmrn", "patient"] };
  function editFields(m) {
    const R = m.row, src = m.kind === "consults" ? S.raw.consults : S.raw.mpt, hdr = (src && src[0] || []).map(C.norm);
    const prefAt = m.kind === "mpt" ? C.prefCols(R.head) : [];
    return R.head.map((h, i) => {
      const k = C.norm(h); if (!k || SKIP_EDIT[m.kind].includes(k) || prefAt.includes(i)) return null;
      const v = R.values[i], f = R.formulas[i], fmt = String(R.formats[i] || ""), dateLike = /[dmy]{1,4}[\/\-.][dmy]/i.test(fmt) && !/h/i.test(fmt) || (/date|exp\b|start|updated|thru|dob|birth/i.test(h) && (typeof v === "number" || v === ""));
      const type = dateLike ? "date" : typeof v === "number" ? "num" : "text";
      const cur = type === "date" ? (C.toSerial(v) != null ? C.isoOf(C.toSerial(v)) : "") : C.txt(v);
      const col = hdr.indexOf(k), seen = col < 0 || !src ? [] : [...new Set(src.slice(1).map(r => C.txt(r[col])).filter(x => x && !/^\d+(\.\d+)?$/.test(x)))];
      const opts = k === "frequency" ? S.freqs : k === "primarylocation" ? locOpts() : seen.length <= 40 && type === "text" ? seen.sort() : [];
      return { i, h, label: h.replace(/\s+/g, " ").trim(), type, cur, formula: !!f, shown: type === "date" && C.toSerial(v) != null ? C.fmtD(C.toSerial(v)) : C.txt(v), opts, long: /notes?\b/i.test(h) };
    }).filter(Boolean);
  }
  const MODALS = {
    checkin() {
      const v = visitBy(S.m.key), m = S.m, c = C.code(v), aOk = teamsReady("arrival"), pOk = teamsReady("protocol");
      const arrTxt = C.arrivalText(v, S.plan ? m.chair : "", S.tpl.arrival), askTxt = C.askText(v.p, v.ask, S.tpl, v.p?.freq);
      return `<h2 id="mh">Check in ${esc(c)}${mrnH(v.a.mrn)}</h2><p class="sub muted">${C.fmtT(v.a.time)} visit with ${esc(C.provName(v.a.provider) || "the provider")}</p>
        ${S.plan ? `<div class="fld">Chair${chairPick(m.chair, true)}</div>` : ""}
        <label class="preview opt"><input type="checkbox" id="mRn"${m.rn && aOk ? " checked" : ""}${aOk ? "" : " disabled"}><span><small>To ${esc(routeTo("arrival"))}</small>${aOk ? esc(arrTxt) : "Arrivals aren’t set up yet (Settings › Teams)."}</span></label>
        ${v.p && v.ask.any ? `<label class="preview opt"><input type="checkbox" id="mAsk"${m.ask && pOk ? " checked" : ""}${pOk ? "" : " disabled"}><span><small>To ${esc(routeTo("protocol"))}, at the same time</small>${pOk ? esc(askTxt) : "The group chat isn’t set up yet (Settings › Teams)."}</span></label>` : ""}
        ${payPickHTML(v, m, "in")}
        <div class="foot"><button class="btn mute" data-act="close">Cancel</button><button class="btn primary inl" data-act="doCheckin"${S.busy || amtBad(m) ? " disabled" : ""}>Check in</button></div>`;
    },
    /* start treatment in a chair (only with a chair map), or move to another one */
    seat() {
      const v = visitBy(S.m.key), c = C.code(v), aOk = teamsReady("arrival"), moving = v.ph === "tx";
      return `<h2 id="mh">${moving ? `Move ${esc(c)} to another chair` : `Start treatment for ${esc(c)}`}${mrnH(v.a.mrn)}</h2><p class="sub muted">${moving && v.st.chair ? `Now in ${esc(chairTxt(v.st.chair))}` : "Pick the chair they are in."}</p>
        <div class="fld">Chair${chairPick(S.m.chair, false)}</div>
        <label class="preview opt"><input type="checkbox" id="mRn"${S.m.rn && aOk ? " checked" : ""}${aOk ? "" : " disabled"}><span><small>Tell ${esc(routeTo("arrival"))} which chair</small>${aOk ? esc(C.fill(S.tpl.seated, { patient: c, chair: S.m.chair || "…" })) : "Arrivals aren’t set up yet (Settings › Teams)."}</span></label>
        <div class="foot"><button class="btn mute" data-act="close">Cancel</button><button class="btn primary inl" data-act="doSeat"${S.m.chair && !S.busy ? "" : " disabled"}>${!S.m.chair ? "Choose a chair" : moving ? "Move" : "Start treatment"}</button></div>`;
    },
    ask() {
      const p = S.PM.get(S.m.mrn), ok = teamsReady("protocol");
      return `<h2 id="mh">Ask ${esc(routeTo("protocol"))}</h2><p class="sub muted">${esc(p.code)}, ${esc(p.freq || "no frequency on the tracker")}</p>
        <label class="fld">Protocol to confirm<select id="mFreq">${["", ...S.freqs].map(f => `<option value="${esc(f)}"${f === S.m.freq ? " selected" : ""}>${esc(f || "Ask for the protocol")}</option>`).join("")}</select></label>
        <label class="fld">Message<textarea id="mText" rows="3">${esc(S.m.text)}</textarea></label>
        ${ok ? "" : `<p class="note warn">The group chat isn’t set up yet. Add its workflow in Settings › Teams, or copy the message and post it yourself.</p>`}
        <div class="foot"><button class="btn mute" data-act="close">Cancel</button>${ok ? "" : `<button class="btn ghost" data-act="copyAsk">Copy message</button>`}<button class="btn primary inl" data-act="doAsk"${ok && !S.busy ? "" : " disabled"}>Send to ${esc(routeTo("protocol"))}</button></div>`;
    },
    checkout() {
      const v = visitBy(S.m.key), c = C.code(v), m = S.m, p = v.p;
      const bk = (k, t, s) => `<button data-act="mBook" data-v="${k}" aria-pressed="${m.booked === k}">${t}<small>${s}</small></button>`;
      return `<h2 id="mh">Check out ${esc(c)}${mrnH(v.a.mrn)}</h2><p class="sub muted">${C.fmtT(v.a.time)} visit · ${esc(C.provName(v.a.provider) || "the provider")}${v.st.seat != null ? ` · in treatment since ${C.fmtT(v.st.seat)}` : ""}</p>
        <div class="fld">Next visit${v.ask.next ? ` <span class="muted">(on the list: ${C.fmtD(v.ask.next.date)})</span>` : ""}<div class="pick">${bk("yes", "Booked", "in NextGen")}${bk("later", "Book later", "keep reminding")}${bk("no", "Not needed", "PRN or ending")}</div></div>
        ${p ? `<label class="opt"><input type="checkbox" id="mProto"${m.proto ? " checked" : ""}> The provider confirmed the protocol<small>Writes Protocol Updated ${C.fmtD(todayS())}${m.freq && m.freq !== p.freq ? ` and Frequency “${esc(m.freq)}”` : ""} on the Master Patient Tracker</small></label>
        <label class="fld">Protocol<select id="mFreq">${freqOpts(p, m.freq)}</select></label>` : ""}
        ${m.idx ? `<div class="msg warn bidfp"><b>Treatment ${m.idx.N} kept: index is done.</b> Move ${esc(c)} to Ongoing and set the protocol the provider chose (usually 1x per week; it varies with acuity).
          <label class="opt"><input type="checkbox" id="mIdx"${m.idxOn ? " checked" : ""}> Move to Ongoing<small>Writes Status Ongoing, Frequency and Protocol Updated ${C.fmtD(todayS())} on the Master Patient Tracker. Pick the protocol above.</small></label></div>` : ""}
        ${m.xs ? `<div class="msg warn bidfp"><b>${m.xs.due ? `Their last visit here: transfer to ${esc(m.xs.to)}.` : `Transfer to ${esc(m.xs.to)} scheduled${m.xs.date != null ? ` for ${C.fmtD(m.xs.date)}` : ""}.`}</b>${thereTxt(m.xs.there, m.xs.to)}
          <label class="opt"><input type="checkbox" id="mXfer"${m.xferOn ? " checked" : ""}> Confirm the transfer<small>Writes Status Discontinued and a transfer note on the Master Patient Tracker${m.xs.after.length ? `. ${plural(m.xs.after.length, "visit")} still booked here after it: cancel in NextGen` : ""}</small></label></div>` : ""}
        ${payPickHTML(v, m, "out")}
        <div class="foot"><button class="btn mute" data-act="close">Cancel</button><button class="btn primary inl" data-act="doCheckout"${m.booked && payAnswered(v, m) && !S.busy ? "" : " disabled"}>${!m.booked ? "Choose the next visit" : !payAnswered(v, m) ? "Say whether payment was collected" : `Check out`}</button></div>`;
    },
    /* today's payment, recorded or changed outside check-in and check-out (after check-out included) */
    payEdit() {
      const v = visitBy(S.m.key), m = S.m, c = C.code(v);
      return `<h2 id="mh">${esc(c)}${mrnH(v.a.mrn)}: today’s payment</h2><p class="sub muted">${C.fmtT(v.a.time)} visit${v.st.out != null ? `, checked out at ${C.fmtT(v.st.out)}` : ""}. A change is saved as a new answer with your name; the latest one counts.</p>
        ${payPickHTML(v, m, "edit")}
        <div class="foot"><button class="btn mute" data-act="close">Cancel</button><button class="btn primary inl" data-act="doPayEdit"${S.busy || amtBad(m) ? " disabled" : ""}>Save payment</button></div>`;
    },
    /* a visit on the Schedule: open the patient, or waitlist them for a sooner date (from when, which days and halves, which provider) */
    sooner() {
      const m = S.m, r = S.SO.get(m.mrn), on = r && r.status === "waiting" && r.booked === m.date;
      const day = (i, h) => `<button data-act="soDay" data-v="${i + 1}${h}" aria-pressed="${m.days.includes(`${i + 1}${h}`)}" aria-label="${DAYW[i]} ${h.toUpperCase()}">${h.toUpperCase()}</button>`;
      return `<h2 id="mh">${esc(m.code || "No code")} <span class="mrn">MRN ${esc(m.mrn)}</span></h2><p class="sub muted">Booked ${dshort(m.date)} at ${C.fmtT(m.time)} with ${esc(provLabel(m.provider))}.${on ? " On the sooner waitlist." : ""}</p>
        <div class="actions"><button class="btn ghost sm line" data-act="soPv">${S.PM.get(m.mrn) ? "Open the patient" : "Make a profile"}</button></div>
        <div class="section"><div class="h"><h2>Wants a sooner date</h2><span>${on ? "waitlisted" : "not waitlisted"}</span></div>
        ${needTable("tblIvkSooner", "The sooner waitlist")}
        <label class="fld">Earliest date<input id="soFrom" type="date" value="${esc(m.from)}" min="${C.isoOf(todayS() + 1)}" max="${C.isoOf(m.date - 1)}"><span class="hint">Openings from this day up to the day before their visit</span></label>
        <div class="fld">Days they can come<div class="prefs av">${C.DAYS5.map(d => `<b>${d}</b>`).join("")}${["am", "pm"].map(h => C.DAYS5.map((d, i) => day(i, h)).join("")).join("")}</div><span class="hint">${m.days.length ? "Only these" : "None picked: any day"}. AM is before noon.</span></div>
        <label class="opt"><input type="checkbox" id="soSame"${m.same ? " checked" : ""}> Only with ${esc(C.provName(m.provider) || "this provider")}</label>
        <label class="fld">Note<input id="soNote" value="${esc(m.note)}" autocomplete="off" placeholder="Can come at short notice"></label></div>
        <div class="foot"><button class="btn mute" data-act="close">Cancel</button>${on ? `<button class="btn ghost" data-act="soDone" data-m="${esc(m.mrn)}">Take off the waitlist</button>` : ""}<button class="btn primary inl" data-act="soSave"${S.soonerOn && !S.busy ? "" : " disabled"}>${on ? "Save" : "Waitlist for sooner"}</button></div>`;
    },
    /* after an upload: openings that would move waitlisted patients sooner, new since the last time anyone was shown them */
    soonerFound() {
      const L = S.m.list;
      return `<h2 id="mh">Sooner openings</h2><p class="sub muted">The new appointment list has ${plural(L.reduce((n, x) => n + x.hits.length, 0), "opening")} for ${plural(L.length, "waitlisted patient")}. Hold one, then book it in NextGen.</p>
        ${L.map(x => `<div class="section"><div class="h"><h2>${esc(x.r.code)}${mrnH(x.r.mrn)}</h2><span>${esc(soonerTxt(x.r))}</span></div><div class="actions">${x.hits.slice(0, 6).map(c => hitBtn(x.r, c)).join("")}</div></div>`).join("")}
        <div class="foot"><button class="btn primary inl" data-act="close">Close</button></div>`;
    },
    /* a patient going to another clinic: discontinue here now as transferred, or schedule it for a future date (confirmed at the
       last visit here). For Oakland maintenance, OAK is tasked to book it. */
    xfer() {
      const m = S.m, p = S.PM.get(m.mrn), xs = C.xferState(S.XF.get(p.mrn), S.A, todayS(), S.x.E), d = C.toSerial(m.date), T = todayS();
      const pk = (k, l, s) => `<button data-act="xfMode" data-v="${k}" aria-pressed="${m.mode === k}">${l}<small>${s}</small></button>`;
      const L = d != null ? S.A.filter(a => a.mrn === p.mrn && !C.isCancelled(a) && a.date < d && a.date >= T) : [], after = d != null ? S.A.filter(a => a.mrn === p.mrn && !C.isCancelled(a) && a.date >= d) : [];
      const oak = /oak/i.test(m.to), tasked = ptOf(p.mrn)?.oakTasked;
      return `<h2 id="mh">Transfer ${esc(p.code)} <span class="mrn">MRN ${esc(p.mrn)}</span></h2><p class="sub muted">Home clinic ${esc(p.loc || "not set")}. ${xs ? `Scheduled to ${esc(xs.to)}${xs.date != null ? ` ${C.fmtD(xs.date)}` : ""}.` : ""}</p>
        <div class="fld">What<div class="pick">${pk("now", "Discontinue now", "transferred, from today")}${pk("later", "Schedule a transfer", "a future date, confirmed at the last visit here")}</div></div>
        <div class="two"><label class="fld">Transferring to<input id="xfTo" value="${esc(m.to)}" autocomplete="off" list="xfClin"><datalist id="xfClin">${locOpts().filter(l => !C.isOurs(l, S.cfg)).map(x => `<option value="${esc(x)}">`).join("")}</datalist></label>
          ${m.mode === "later" ? `<label class="fld">Transfer date<input id="xfDate" type="date" value="${esc(m.date)}" min="${C.isoOf(T + 1)}"><span class="hint">Their first visit at the new clinic, or when it takes over</span></label>` : ""}</div>
        <label class="fld">Note<input id="xfNote" value="${esc(m.note)}" autocomplete="off" placeholder="Index here, maintenance at Oakland"></label>
        ${m.mode === "later" && d != null ? `<p class="note">${L.length ? `Last visit here before then: ${dshort(L[L.length - 1].date)} at ${C.fmtT(L[L.length - 1].time)}. Check-out of it asks to confirm the transfer.` : "No visits booked here before then: confirm the transfer from the patient view when it happens."}${after.length ? ` <b>${plural(after.length, "visit")} booked here on or after ${C.fmtD(d)}</b>: cancel in NextGen.` : ""}</p>` : ""}
        ${oak ? `<label class="opt"><input type="checkbox" id="xfOak"${m.oak ? " checked" : ""}${tasked != null ? " disabled" : ""}> Task OAK to schedule them<small>${tasked != null ? `OAK tasked ${C.fmtD(tasked)}` : "Marks OAK tasked today (the Oakland task in NextGen is still yours to send)"}</small></label>` : ""}
        ${m.mode === "now" ? `<p class="note">Writes Status Discontinued and “${esc(xferNote(m))}” to Clinic Notes on the Master Patient Tracker.</p>` : ""}
        <div class="foot">${xs ? `<button class="btn mute" data-act="xfCancel">Cancel the transfer</button>` : ""}<button class="btn mute" data-act="close">Close</button><button class="btn primary inl" data-act="doXfer"${S.busy ? " disabled" : ""}>${m.mode === "now" ? (xs ? "Confirm the transfer" : "Discontinue as transferred") : "Schedule the transfer"}</button></div>`;
    },
    /* every field on the patient's Master Patient Tracker row (or Consults row) that isn't a formula */
    edit() {
      const m = S.m;
      if (!m.row) return `<h2 id="mh">${esc(m.title)}</h2><p class="sub muted">Reading the row…</p><div class="foot"><button class="btn mute" data-act="close">Cancel</button></div>`;
      const F = editFields(m), fld = f => {
        const lab = esc(f.label), id = `ef_${f.i}`;
        if (f.formula) return `<div class="fld ro">${lab}<span class="rov">${esc(f.shown || "—")}</span><span class="hint">Formula, not editable</span></div>`;
        const dl = f.opts.length ? ` list="dl_${f.i}"` : "", list = f.opts.length ? `<datalist id="dl_${f.i}">${f.opts.map(o => `<option value="${esc(o)}">`).join("")}</datalist>` : "";
        if (f.type === "date") return `<label class="fld">${lab}<input id="${id}" type="date" value="${esc(f.cur)}"></label>`;
        if (f.long) return `<label class="fld wide">${lab}<textarea id="${id}" rows="2">${esc(f.cur)}</textarea></label>`;
        return `<label class="fld">${lab}<input id="${id}"${f.type === "num" ? ' inputmode="decimal"' : ""} value="${esc(f.cur)}"${dl} autocomplete="off">${list}</label>`;
      };
      return `<h2 id="mh">${esc(m.title)}${mrnH(m.mrn)}</h2><p class="sub muted">${m.kind === "consults" ? "The Consults tab" : "The Master Patient Tracker"}, saved straight to the row. Formula cells are left alone.${m.kind === "mpt" ? " Morning and afternoon days are under Availability." : ""}</p>
        <div class="egrid">${F.map(fld).join("")}</div>
        <div class="foot"><button class="btn mute" data-act="close">Cancel</button><button class="btn primary inl" data-act="doEdit"${S.busy ? " disabled" : ""}>Save changes</button></div>`;
    },
    /* a payment on account (comes off the oldest balance) or a balance reminder */
    pay() {
      const p = S.PM.get(S.m.mrn), L = C.ledger(p.mrn, S.pays), t = ptOf(p.mrn), today = C.isoOf(todayS());
      return `<h2 id="mh">${esc(p.code)}${mrnH(p.mrn)}: payments</h2><p class="sub muted">${L.balance > 0 ? `Balance ${C.money(L.balance)}` : "No balance"}${t?.card ? `. Luma card on file${t.okCollect ? ", okay to charge at time of service" : ", not okay to charge at time of service yet"}` : ". No card on file in Luma"}.</p>
        <div class="section"><div class="h"><h2>Payment</h2></div><div class="two">
          <label class="fld">Amount<input id="pyAmt" inputmode="decimal" value="${L.balance > 0 ? L.balance : ""}"></label>
          <label class="fld">How<select id="pyHow">${["Luma card on file", "Card in person", "Cash or check", "Other"].map(h => `<option${t?.card && t.okCollect && h === "Luma card on file" ? " selected" : ""}>${h}</option>`).join("")}</select></label>
          <label class="fld">Date<input id="pyDate" type="date" value="${today}"></label><label class="fld">Note<input id="pyNote" autocomplete="off"></label></div>
          <div class="actions"><button class="btn primary inl" data-act="doPay"${S.busy ? " disabled" : ""}>Record payment</button></div></div>
        <div class="section"><div class="h"><h2>Balance reminder sent</h2><span>${L.lastReminder ? `last ${C.fmtD(L.lastReminder.date)}` : "none yet"}</span></div><div class="two">
          <label class="fld">Sent<input id="rmDate" type="date" value="${today}"></label>
          <label class="fld">Method<select id="rmHow">${C.REMIND_METHODS.map(h => `<option>${h}</option>`).join("")}</select></label></div>
          <div class="actions"><button class="btn ghost inl line" data-act="doRemind"${S.busy ? " disabled" : ""}>Record reminder</button></div>
          ${L.reminders.length ? `<p class="note">${L.reminders.slice(-5).reverse().map(r => `${C.fmtD(r.date)}, ${esc(r.method || "—")}${r.by ? `, ${esc(r.by)}` : ""}`).join("; ")}</p>` : ""}</div>
        <div class="foot"><button class="btn mute" data-act="close">Close</button></div>`;
    },
    /* availability as the patient availability form asks it, effective and expiry dates, and dates away */
    avail() {
      const m = S.m, p = S.PM.get(m.mrn);
      const cell = (i, h) => `<div class="avpick" role="group" aria-label="${DAYW[i]} ${h.toUpperCase()}">${C.AV_LEVELS.map(([k, l]) => `<button data-act="avSet" data-d="${i}" data-h="${h}" data-v="${k}" aria-pressed="${m.grid[i][h] === k}" class="${AV_CLS[k]}" title="${l}">${l.replace(" available", "").replace("Not", "Not")}</button>`).join("")}</div>`;
      const away = m.away.map((w, j) => `<div class="awrow"><input type="date" id="awF${j}" value="${esc(w.from)}" aria-label="Away from"><input type="date" id="awT${j}" value="${esc(w.to)}" aria-label="Away to"><input id="awN${j}" value="${esc(w.note)}" placeholder="Note" aria-label="Note" autocomplete="off"><button class="btn mute sm" data-act="avDelAway" data-v="${j}" aria-label="Remove these dates">Remove</button></div>`).join("");
      return `<h2 id="mh">${esc(p.code)}${mrnH(p.mrn)}: availability</h2><p class="sub muted">From the patient availability form. Preferred and Mostly available go to the tracker’s AM/PM columns and the Index planner.</p>
        <div class="avedit"><span></span><b>Morning (AM)</b><b>Afternoon (PM)</b>${C.DAYS5.map((d, i) => `<b>${DAYW[i]}</b>${cell(i, "am")}${cell(i, "pm")}`).join("")}</div>
        <div class="two"><label class="fld">Effective<input id="avEff" type="date" value="${esc(m.eff)}"></label><label class="fld">Expires<input id="avExp" type="date" value="${esc(m.exp)}"></label></div>
        <label class="fld">Note<input id="avNote" value="${esc(m.note)}" autocomplete="off"></label>
        <div class="section"><div class="h"><h2>Dates away</h2></div>${away || `<p class="note">None.</p>`}<div class="actions"><button class="btn ghost sm line" data-act="avAddAway">Add dates away</button></div></div>
        <div class="foot"><button class="btn mute" data-act="close">Cancel</button><button class="btn ghost" data-act="printAvail" data-m="${esc(m.mrn)}">Print the form</button><button class="btn primary inl" data-act="doAvail"${S.busy ? " disabled" : ""}>Save availability</button></div>`;
    },
    /* a provider's time off: dates, who covers, and whether the covering provider works the absent provider's template (pre-fill) */
    pto() {
      const m = S.m, absent = S.grids.find(g => g.key === TM.provKey(m.provider)), aName = absent ? TM.nameOf(absent) : m.provider;
      const covs = [...new Map(S.grids.map(g => [g.key, TM.nameOf(g)]).concat(S.providers.map(p => [TM.provKey(p), C.provFull(p)]))).entries()].filter(([k]) => k !== TM.provKey(m.provider));
      const cKey = m.cov ? TM.provKey(m.cov) : "", cName = m.cov ? exLabel(m.cov) : "", from = C.toSerial(m.from), to = C.toSerial(m.to) ?? from;
      /* does the covering provider have a template of their own on any of those weekdays? */
      const cg = S.grids.find(g => g.key === cKey), ownDays = cg && from != null ? [...Array(Math.min(31, (to ?? from) - from + 1)).keys()].filter(i => cg.dayCols?.some(x => x.d === C.weekday(from + i))).length : 0;
      const pk = (v, l, s) => `<button data-act="ptoTpl" data-v="${v}" aria-pressed="${m.useTpl === v}">${l}<small>${s}</small></button>`;
      return `<h2 id="mh">${m.id && !m.isNew ? "Change time off" : "Add time off"}</h2><p class="sub muted">An exception on the template: the provider’s slots close on these days. Nothing changes in NextGen.</p>
        <label class="fld">Provider<select id="ptProv">${S.grids.map(g => `<option value="${esc(TM.nameOf(g))}"${g.key === TM.provKey(m.provider) ? " selected" : ""}>${esc(TM.nameOf(g))}</option>`).join("")}</select></label>
        <div class="two"><label class="fld">First day off<input id="ptFrom" type="date" value="${esc(m.from)}"></label><label class="fld">Last day off<input id="ptTo" type="date" value="${esc(m.to)}"><span class="hint">Blank: one day</span></label></div>
        <label class="fld">Covered by<select id="ptCov"><option value="">Nobody: close their slots</option>${covs.map(([k, n]) => `<option value="${esc(n)}"${k === cKey ? " selected" : ""}>${esc(n)}</option>`).join("")}</select></label>
        ${m.cov ? `<div class="fld">Use ${esc(aName)}’s template for ${esc(cName)} on these days?<div class="pick">${pk("yes", `Yes, pre-fill it`, `${esc(cName)} takes ${esc(aName)}’s slots`)}${pk("no", "No", `${esc(cName)}’s own template`)}</div>
          ${m.useTpl === "no" && !ownDays ? `<p class="note warn">${esc(cName)} has no template on those days, so no slots will be offered. Pre-fill to keep ${esc(aName)}’s slots open.</p>` : ""}</div>` : ""}
        <label class="fld">Note<input id="ptNote" value="${esc(m.note || "")}" autocomplete="off" placeholder="PTO, conference…"></label>
        <div class="foot">${m.id && !m.isNew ? `<button class="btn mute" data-act="ptoDel">Remove</button>` : ""}<button class="btn mute" data-act="close">Cancel</button><button class="btn primary inl" data-act="doPto"${m.cov && !m.useTpl || S.busy ? " disabled" : ""}>${m.cov && !m.useTpl ? "Answer the template question" : "Save time off"}</button></div>`;
    },
    /* the forms printed (or emailed) for a patient: their schedule, the IVK consent, the availability form */
    forms() {
      const m = S.m, p = formsPatient(m.mrn), email = m.mode === "email";
      const provs = [...new Set([m.prov, ...S.providers.map(C.provFull)].filter(Boolean))];
      const part = (k, l, s) => `<label class="opt"><input type="checkbox" id="fp_${k}"${m.parts[k] ? " checked" : ""}> ${l}${s ? `<small>${s}</small>` : ""}</label>`;
      return `<h2 id="mh">${email ? `Start email for ${esc(p.code)}` : `Print for ${esc(p.code)}`}${mrnH(p.mrn)}</h2>
        <p class="sub muted">${p.onTracker ? "The patient’s name comes from the Master Patient Tracker. " : ""}It’s only put on the forms, never saved or posted.${p.onTracker && !p.name ? " <b>This row has no Patient Name, so the forms will have a blank name.</b>" : ""}</p>
        ${p.onTracker ? "" : `<label class="fld">Patient’s name<input id="fmName" value="${esc(m.name || "")}" autocomplete="off" spellcheck="false"><span class="hint">Not on the Master Patient Tracker yet; type it for the forms (not saved). The schedule only needs the MRN.</span></label>`}
        <div class="two"><label class="fld">Date of birth<input id="fmDob" type="date" value="${esc(m.dob)}"><span class="hint">${p.dob != null ? "From the tracker" : "Not on the tracker; type it for the forms (not saved)"}</span></label>
          <label class="fld">Clinic location<input id="fmLoc" value="${esc(m.loc)}" autocomplete="off"></label></div>
        <label class="fld">Clinician<select id="fmProv">${provs.map(x => `<option${x === m.prov ? " selected" : ""}>${esc(x)}</option>`).join("")}</select></label>
        ${email ? `<p class="note">The email gets the start template (Settings › Messages) with the schedule in the body, and the schedule and the filled-in IVK consent attached. In the next window, upload the BIDF as a PDF (save it from NextGen first).</p>`
          : `<div class="fld">Print${part("sched", "Schedule")}${part("consent", "IV ketamine consent", "Name, date of birth, clinic location and clinician filled in")}${part("avail", "Availability form", "Name and date of birth filled in")}</div>
            ${m.parts.consent ? `<div class="msg warn bidfp"><b>Print the BIDF from NextGen too.</b> The benefits team completes it; it isn’t made here.${p.bidf ? ` The tracker has it: ${esc(p.bidf)}.` : " The tracker doesn’t show one yet: check with the benefits team before the start."}
              <label class="opt"><input type="checkbox" id="fp_bidf"${m.bidf ? " checked" : ""}> I printed the BIDF from NextGen<small>Saves today’s date on the patient’s profile</small></label></div>` : ""}`}
        ${email || m.parts.sched ? schedRangeHTML(p, m) : ""}
        <div class="foot"><button class="btn mute" data-act="close">Cancel</button><button class="btn primary inl" data-act="doForms">${email ? "Make the email" : "Open to print"}</button></div>`;
    },
    /* an Induction patient who has kept their 6th visit: Ongoing, on the protocol the provider set */
    idxDone() {
      const p = S.PM.get(S.m.mrn), ix = C.indexState(p, S.A, todayS(), S.cfg);
      return `<h2 id="mh">${esc(p.code)}${mrnH(p.mrn)}: index done</h2><p class="sub muted">${ix ? `${ix.done} index visits kept${ix.sixth ? `, treatment ${ix.N} on ${C.fmtD(ix.sixth.date)}` : ""}.` : ""} The status moves from Induction to Ongoing.</p>
        <label class="fld">Protocol after index<select id="mFreq">${[...new Set([S.m.freq, ...S.freqs].filter(Boolean))].map(f => `<option${f === S.m.freq ? " selected" : ""}>${esc(f)}</option>`).join("")}</select><span class="hint">Usually ${esc(S.cfg.AfterIndex)}; the provider sets it by acuity. Ask the group chat if it isn’t confirmed.</span></label>
        <p class="note">Writes Status Ongoing, Frequency and Protocol Updated ${C.fmtD(todayS())} on the Master Patient Tracker.</p>
        <div class="foot"><button class="btn mute" data-act="close">Cancel</button><button class="btn ghost" data-act="ask" data-m="${esc(p.mrn)}">Ask the group chat</button><button class="btn primary inl" data-act="doIdxDone"${S.busy ? " disabled" : ""}>Move to Ongoing</button></div>`;
    },
    proto() {
      const p = S.PM.get(S.m.mrn);
      return `<h2 id="mh">Confirm ${esc(p.code)}’s protocol${mrnH(p.mrn)}</h2><p class="sub muted">${p.protoAt != null ? `Last confirmed ${C.fmtD(p.protoAt)}` : "Not confirmed on the tracker yet"}</p>
        <label class="fld">Protocol the provider confirmed<select id="mFreq">${freqOpts(p, S.m.freq)}</select></label>
        <p class="note">Writes Protocol Updated ${C.fmtD(todayS())} on the Master Patient Tracker${S.m.freq !== p.freq ? `, and Frequency “${esc(S.m.freq)}”` : ""}.</p>
        <div class="foot"><button class="btn mute" data-act="close">Cancel</button><button class="btn primary inl" data-act="doProto"${S.busy ? " disabled" : ""}>Save protocol</button></div>`;
    },
    addProv() {
      const m = S.m, have = new Set(S.grids.map(g => g.key)), free = S.providers.filter(p => !have.has(TM.provKey(p)));
      const times = []; for (let t = 420; t <= 1200; t += 30) times.push(t);
      const tsel = (id, v) => `<select id="${id}">${times.map(t => `<option value="${t}"${t === v ? " selected" : ""}>${C.fmtT(t)}</option>`).join("")}</select>`;
      return `<h2 id="mh">Add a provider</h2><p class="sub muted">A weekly grid on the ${X.TPL_SHEET} tab in the clinic’s template layout, with its Event Types tally.</p>
        <label class="fld">Provider<select id="apName">${free.map(p => `<option value="${esc(p)}"${p === m.name && !m.other ? " selected" : ""}>${esc(nameParts(p).title)}</option>`).join("")}<option value="__other"${m.other ? " selected" : ""}>Someone not on the report…</option></select></label>
        ${m.other ? `<label class="fld">Name<input id="apTyped" value="${esc(m.typed || "")}" placeholder="Last, First MD" autocomplete="off"></label>` : ""}
        <div class="fld">Days in clinic<div class="pick">${C.DAYS5.map((d, i) => `<button data-act="mDay" data-d="${i + 1}" aria-pressed="${m.days.includes(i + 1)}">${d}</button>`).join("")}</div></div>
        <div class="two"><label class="fld">From${tsel("apStart", m.start)}</label><label class="fld">Until${tsel("apEnd", m.end)}</label>
          <label class="fld">Effective from<input id="apFrom" type="date" value="${esc(m.from)}"><span class="hint">Blank: every week</span></label>
          <label class="fld">Start from<select id="apCopy"><option value="">An empty grid</option>${S.grids.map((g, i) => `<option value="${i}"${String(i) === String(m.copy) ? " selected" : ""}>A copy of ${esc(TM.nameOf(g))}</option>`).join("")}</select></label></div>
        <div class="foot"><button class="btn mute" data-act="close">Cancel</button><button class="btn primary inl" data-act="doAddProv"${S.busy ? " disabled" : ""}>Add the provider</button></div>`;
    },
    /* the patient profile: on the tracker (status, protocol, what's due), who covers the provider visit and the infusion, and the
       latest Sohar benefits, typed or filled from the Sohar Eligibility add-in's benefits file */
    profile() {
      const m = S.m, sh = m.sohar, ins = [...new Set(S.P.flatMap(p => [p.ins1, p.ins2]).concat([...S.PF.values()].map(x => x.sohar.payer)).filter(x => x && !C.isSelf(x)))].sort();
      const mp = C.modeParts(m.mode || "both"), f = (id, l, v, o) => `<label class="fld${o?.wide ? " wide" : ""}">${l}<input id="${id}" value="${esc(v ?? "")}"${o?.num ? ' inputmode="decimal"' : ""}${o?.list ? ' list="pfIns"' : ""}${o?.date ? ' type="date"' : ""} autocomplete="off"${o?.ro ? " readonly" : ""}>${o?.hint ? `<span class="hint">${o.hint}</span>` : ""}</label>`;
      const md = (k, l) => `<button data-act="pfMode" data-v="${k}" aria-pressed="${(m.mode || "both") === k}">${esc(l)}</button>`;
      const sts = ["Induction", "Ongoing", "Paused", "PRN"], freqs = [...new Set([m.freq, ...S.freqs].filter(Boolean))];
      return `<h2 id="mh">${m.onMpt ? `${esc(m.code)}’s profile${mrnH(m.mrn)}` : `New patient profile`}</h2><p class="sub muted">${m.onMpt ? "Saved to the Master Patient Tracker row and the add-in’s profile table." : "Adds the patient to the Master Patient Tracker by LASFIR code, with the profile beside it."}${m.first ? ` First visit on the list: ${dshort(m.first.date)} at ${C.fmtT(m.first.time)}${m.start ? ", a treatment start" : ""}.` : ""}</p>
        <datalist id="pfIns">${ins.map(x => `<option value="${esc(x)}">`).join("")}</datalist>
        <div class="two">${f("pfCode", "LASFIR", m.code, { ro: m.onMpt, hint: m.onMpt ? "" : "Last three and first three letters, never the name" })}${f("pfMrn", "MRN", m.mrn, { ro: true })}
          <label class="fld">Status<select id="pfStatus">${sts.map(x => `<option${x === m.status ? " selected" : ""}>${x}</option>`).join("")}${m.status && !sts.includes(m.status) ? `<option selected>${esc(m.status)}</option>` : ""}</select></label>
          <label class="fld">Protocol<select id="pfFreq"><option value="">Not set yet</option>${freqs.map(x => `<option${x === m.freq ? " selected" : ""}>${esc(x)}</option>`).join("")}</select></label>
          ${m.onMpt ? "" : f("pfStart", "Tx start date", m.startIso, { date: true })}</div>
        <div class="fld">Who covers what<div class="pick col">${C.COVER_MODES.map(([k, l]) => md(k, l)).join("")}</div></div>
        <div class="two"><div class="pfpart"><b>Provider visit (99214/99215)</b>${mp.prov === "selfpay" ? `<p class="note">Self pay</p>` : f("pfProvPayer", "Payer", m.provPayer, { list: true })}${f("pfProvDue", "Patient pays each visit", m.provDue, { num: true, hint: mp.prov === "selfpay" ? "The self-pay rate" : "Their copay for the provider visit" })}</div>
          <div class="pfpart"><b>Infusion</b>${mp.inf === "selfpay" ? `<p class="note">Self pay</p>` : f("pfInfPayer", "Payer", m.infPayer, { list: true })}${f("pfInfDue", "Patient pays each visit", m.infDue, { num: true, hint: mp.inf === "selfpay" ? "The self-pay infusion rate" : "Their share of the infusion" })}</div></div>
        <div class="section"><div class="h"><h2>Latest Sohar benefits</h2><label class="btn ghost sm line" style="cursor:pointer">Fill from a Sohar file<input type="file" id="sohFile" accept=".csv,.xlsx,.xls" hidden></label></div>
          ${m.fromFile ? `<p class="note">${esc(m.fromFile)}</p>` : ""}
          <div class="two">${f("pfSPayer", "Sohar payer", sh.payer, { list: true })}${f("pfSPlan", "Plan", sh.plan)}${f("pfSNet", "Network", sh.network)}${f("pfSStatus", "Sohar status", sh.status)}
            ${f("pfSCovS", "Coverage start", sh.covStart, { date: true })}${f("pfSCovE", "Coverage end", sh.covEnd, { date: true })}${f("pfSCopay", "Copay", sh.copay)}${f("pfSCoins", "Coinsurance", sh.coins)}
            ${f("pfSDed", "Deductible", sh.ded)}${f("pfSDedR", "Deductible remaining", sh.dedRem)}${f("pfSOop", "OOP max", sh.oop)}${f("pfSOopR", "OOP remaining", sh.oopRem)}
            ${f("pfSDate", "Verified", sh.date, { date: true })}${f("pfSRef", "Sohar ref", sh.ref)}</div></div>
        ${S.profOn ? "" : `<p class="note warn">The profile table isn’t in this workbook yet: the tracker fields save, the rest needs <b>Update the workbook</b> in Settings.</p>`}
        <div class="foot"><button class="btn mute" data-act="close">Cancel</button><button class="btn primary inl" data-act="doProfile"${S.busy ? " disabled" : ""}>${m.onMpt ? "Save the profile" : "Add to the tracker"}</button></div>`;
    },
    /* an open slot on the Schedule, held for one patient (a new start goes to the Index planner instead) */
    hold() {
      const m = S.m, s = m.slot, seen = new Set(), o = (mrn, code, why) => { if (!mrn || seen.has(mrn)) return ""; seen.add(mrn); return `<option value="${esc(mrn)}"${m.who === mrn ? " selected" : ""}>${esc(code)}${why ? `, ${esc(why)}` : ""}</option>`; };
      const opts = [...ipList().map(x => o(x.mrn, x.code, "on the planner")), ...S.WL.list.map(w => o(w.mrn, w.code, "waitlist")), ...S.P.filter(p => C.isActive(p)).sort((a, b) => a.code.localeCompare(b.code)).map(p => o(p.mrn, p.code, p.freq))].join("");
      return `<h2 id="mh">Hold ${dshort(s.date)} at ${C.fmtT(s.time)}</h2><p class="sub muted">${esc(provLabel(s.provider))}, ${s.openIndex ? "an open index start" : "an open follow-up"}${s.postUse ? `, post check ${C.fmtT(s.post)}` : ""}. The hold shows for everyone until it’s booked in NextGen.</p>
        <label class="fld">For<select id="hdWho"><option value="">Choose the patient…</option>${opts}<option value="__new"${m.who === "__new" ? " selected" : ""}>A new patient (LASFIR and MRN)</option></select></label>
        ${m.who === "__new" ? `<div class="two"><label class="fld">LASFIR<input id="hdCode" maxlength="8" autocomplete="off" spellcheck="false"></label><label class="fld">MRN<input id="hdMrn" inputmode="numeric" autocomplete="off"></label></div>` : ""}
        ${s.openIndex && elIs(s) ? `<p class="note">A new start? <button class="btn ghost sm" data-act="schedPlan">Plan all ${C.indexN(S.cfg)} index visits from here</button></p>` : ""}
        <div class="foot"><button class="btn mute" data-act="close">Cancel</button><button class="btn primary inl" data-act="doHold"${S.busy ? " disabled" : ""}>Hold this time</button></div>`;
    },
    confirm() {
      return `<h2 id="mh">${esc(S.m.title)}</h2><p class="sub">${esc(S.m.body)}</p><div class="foot"><button class="btn mute" data-act="close">Keep it</button><button class="btn primary inl" data-act="doConfirm">${esc(S.m.yes)}</button></div>`;
    },
    upload() {
      const u = S.m, sh = u.sh, ds = u.rows.map(r => r[2]), sts = {};
      u.rows.forEach(r => { const k = r[6] || "blank"; sts[k] = (sts[k] || 0) + 1; });
      const why = [sh.few ? `only ${plural(u.rows.length, "appointment")}` : "", sh.half ? `fewer than half of the ${sh.had} on file` : "", sh.short ? `it ends ${C.fmtD(sh.nMax)}; the list on file runs to ${C.fmtD(sh.cMax)}` : ""].filter(Boolean);
      const many = u.locs.length > 1, away = u.away.length;
      const locPick = many ? `<div class="fld">Which locations are ${esc(S.cfg.Clinic)}?${u.locs.map((l, i) => `<label class="opt"><input type="checkbox" id="upLoc_${i}"${l.ours ? " checked" : ""}> ${esc(l.loc || "No location")}<small>${plural(l.n, "IVK visit")}</small></label>`).join("")}
          <span class="hint">Ticked locations go on the Appt List tab. The rest are kept apart${S.elseOn ? "" : " (once Update the workbook adds their table)"}, only to check transfers and Oakland hand-offs are booked there; they don’t show on Today or the Schedule.</span></div>` : "";
      return `<h2 id="mh">Replace the appointment list?</h2><p class="sub muted">${esc(u.file)}</p>
        ${locPick}
        ${u.rows.length ? `<div class="preview">${plural(u.rows.length, "appointment")} at ${esc(S.cfg.Clinic)}, ${C.fmtD(Math.min(...ds))} to ${C.fmtD(Math.max(...ds))}. ${Object.entries(sts).map(([k, n]) => `${esc(k)} ${n}`).join(", ")}${away ? `. ${plural(away, "visit")} at other clinics.` : ""}</div>`
          : `<div class="msg warn">No visits at ${esc(S.cfg.Clinic)} in this report. Tick this clinic’s location${many ? " above" : ", or check the NextGen report’s filters"}.</div>`}
        ${u.other && u.other.length ? `<p class="note">${plural(u.other.length, "other visit")} (telehealth, TMS and the rest) are kept with no patient on them, so flex post checks those visits use aren’t offered for index.</p>` : `<p class="note">No other visits in this report, so every flex post check shows as free. Run the wider report to include the providers’ telehealth and TMS visits.</p>`}
        <p class="note">The ${sh.had} rows on the Appt List tab are replaced (MRN to Location). The Name and Lookup Key formulas stay. Then the Index Planner tab’s grid is rewritten with every IVK visit and hold, as values; anything typed in that grid is replaced.</p>
        ${sh.flag ? `<div class="msg warn" style="margin-top:12px"><b>This list is much smaller than the one on file:</b> ${why.join("; ")}. The date range in NextGen may have been narrowed.<label class="opt"><input type="checkbox" id="mShrink"${u.ok ? " checked" : ""}> The smaller list is right; replace anyway</label></div>` : ""}
        <div class="foot"><button class="btn mute" data-act="close">Cancel</button><button class="btn primary inl" data-act="doUpload"${sh.flag && !u.ok || !u.rows.length || S.busy ? " disabled" : ""}>Replace list</button></div>`;
    }
  };
  function openModal(name, m) { S.modal = name; S.m = m || {}; render(); }
  function closeModal() { S.modal = null; S.m = {}; render(); }

  /* ---------- actions ---------- */
  const visitBase = v => ({ Date: C.isoOf(v.a.date), MRN: v.a.mrn, Patient: C.code(v), Appt: C.fmtT(v.a.time) });
  const stamp = () => ({ By: S.name, Updated: new Date().toISOString() });
  const needName = () => { if (S.name) return false; toast("Add your first name in Settings first, so Teams posts say who sent them."); S.view = "settings"; render(); setTimeout(() => $("#myName")?.focus(), 50); return true; };
  async function sendArrival(v, chair, seatedOnly) {
    const c = C.code(v), text = seatedOnly ? C.fill(S.tpl.seated, { patient: c, chair }) : C.arrivalText(v, chair, S.tpl.arrival);
    return teams("arrival", text, [["Visit", C.fmtT(v.a.time)], ["Provider", C.provName(v.a.provider)], ["Chair", chair ? chairTxt(chair) : ""], ["Desk", S.name]], c, chair);
  }
  async function sendAsk(p, text, freq, key) {
    const r = await teams("protocol", text, [["Patient", p.code], ["Protocol on file", p.freq || "none"], freq && freq !== p.freq ? ["Asking about", freq] : ["", ""], ["Desk", S.name]], p.code, "");
    if (r.ok && key) { const v = visitBy(key); if (v) await Store.saveVisit(key, visitBase(v), Object.assign({ Asked: hm() }, stamp())); }
    return r;
  }
  /* follow along: picking a patient or a visit selects their row on the Master Patient Tracker behind the add-in (Settings turns it off) */
  let followed = "";
  function follow(kind, mrn) {
    if (!mrn || LS.get("ivkFollow", "1") !== "1" || followed === kind + mrn) return;
    followed = kind + mrn; Store.openRow(kind, mrn).catch(() => { followed = ""; });
  }
  const payRow = (p, o) => Object.assign({ MRN: p.mrn, Patient: p.code, By: S.name }, o);
  const ptPatch = (p, o) => Object.assign({ Patient: p.code, By: S.name, Updated: new Date().toISOString() }, o);
  async function openEdit(kind, mrn) {
    const code = kind === "consults" ? (S.CS.find(c => c.mrn === mrn)?.code || "") : S.PM.get(mrn)?.code || "";
    openModal("edit", { kind, mrn, title: `Edit ${code || "patient"}${kind === "consults" ? "’s consult" : ""}` });
    try { const row = await Store.readRow(kind, mrn); if (S.modal === "edit" && S.m.mrn === mrn) { S.m.row = row; render(); } }
    catch (e) { closeModal(); toast(String(e.message || e), true); }
  }
  /* the time-off form's fields, kept across re-renders */
  const keepPto = () => { if (S.modal !== "pto") return; const m = S.m, v = id => $(id)?.value;
    if ($("#ptProv")) m.provider = v("#ptProv"); if ($("#ptFrom")) m.from = v("#ptFrom"); if ($("#ptTo")) m.to = v("#ptTo"); if ($("#ptNote")) m.note = v("#ptNote"); };
  /* the availability form's state, kept across re-renders */
  /* the sooner form's fields, kept across re-renders */
  const keepSooner = () => { if (S.modal !== "sooner") return; const m = S.m; if ($("#soFrom")) m.from = $("#soFrom").value; if ($("#soSame")) m.same = $("#soSame").checked; if ($("#soNote")) m.note = $("#soNote").value; };
  const keepAvail = () => { if (S.modal !== "avail") return; const m = S.m, v = id => $(id)?.value ?? "";
    m.eff = v("#avEff"); m.exp = v("#avExp"); m.note = v("#avNote"); m.away = m.away.map((w, j) => ({ from: v(`#awF${j}`), to: v(`#awT${j}`), note: v(`#awN${j}`) })); };
  /* the forms window: schedule, consent, availability form, start email. It runs in a frame over the add-in (so it works in full
     screen too); the patient's name and DOB go to it by message and live only in its memory. */
  function openForms(data) {
    closeForms();
    const wrap = document.createElement("div"); wrap.className = "formwrap"; wrap.id = "formwrap";
    wrap.innerHTML = `<iframe title="Forms" src="forms.html?v=6"></iframe>`; document.body.appendChild(wrap);
    const fr = wrap.querySelector("iframe");
    const onMsg = ev => { if (ev.source !== fr.contentWindow || ev.origin !== location.origin) return; const d = ev.data || {};
      if (d.forms === "ready") fr.contentWindow.postMessage({ forms: "data", data }, location.origin);
      else if (d.forms === "close") closeForms();
      else if (d.forms === "saved") toast(d.text || "Saved."); };
    window.addEventListener("message", onMsg); wrap._off = () => window.removeEventListener("message", onMsg);
  }
  function closeForms() { const w = $("#formwrap"); if (w) { w._off && w._off(); w.remove(); } }
  /* what the forms need for one patient */
  /* The patient schedule, as the TMS Tracker makes it: the NextGen appointment list's columns for the visits in a date range
     (Next 14 days, Next 30 days, Rest of course, or From/To), "As of" the list's run time, named "{code} schedule m.d.pdf". */
  const pad2 = n => String(n).padStart(2, "0");
  const ymd = s => { const d = C.dateOf(s); return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`; };
  const schedVisits = (p, from, to) => S.A.filter(a => a.mrn === p.mrn && a.date >= Math.max(from ?? todayS(), todayS()) && a.date <= (to ?? Infinity) && !C.isCancelled(a) && !/no ?show/i.test(a.status))
    .sort((a, b) => a.date - b.date || (a.time ?? 0) - (b.time ?? 0));
  function asOfText() {
    const u = S.raw.apptUpdated, d = new Date(), s = u != null ? Math.floor(u) : C.serialOf(d), mins = u != null ? C.toMin(u) : d.getHours() * 60 + d.getMinutes(), x = C.dateOf(s);
    return `${pad2(x.getMonth() + 1)}/${pad2(x.getDate())}/${x.getFullYear()}  ${C.fmtT(mins)}`;
  }
  function schedObj(p, from, to) {
    const d = C.dateOf(todayS());
    return { fname: `${p.code} schedule ${d.getMonth() + 1}.${d.getDate()}.pdf`, asOf: `As of: ${asOfText()}`,
      rows: schedVisits(p, from, to).map(a => [a.mrn, a.loc || S.cfg.Clinic, ymd(a.date), C.fmtT(a.time), a.provider, a.event, a.status || "Expected"]) };
  }
  function schedRangeHTML(p, m) {
    const t = todayS(), all = schedVisits(p, t, null), last = all.length ? all[all.length - 1].date : t;
    const quick = [["Next 14 days", t + 14], ["Next 30 days", t + 30], ["Rest of course", last]], n = schedVisits(p, C.toSerial(m.from), C.toSerial(m.to)).length;
    return `<div class="fld">Schedule dates<div class="chips" role="group" aria-label="Range" style="margin:8px 0 0">${quick.map(([l, d]) => `<button class="chip" data-act="fmRange" data-v="${C.isoOf(d)}" aria-pressed="${m.to === C.isoOf(d) && m.from === C.isoOf(t)}">${l}</button>`).join("")}</div></div>
      <div class="two"><label class="fld">From<input type="date" id="fmFrom" value="${esc(m.from)}"></label><label class="fld">To<input type="date" id="fmTo" value="${esc(m.to)}"></label></div>
      <p class="note">${plural(n, "visit")} in this range.</p>`;
  }
  /* the forms modal's typed fields, kept across re-renders */
  const keepForms = () => { if (S.modal !== "forms") return; const m = S.m, v = id => $(id)?.value;
    if ($("#fmName")) m.name = v("#fmName"); if ($("#fmDob")) m.dob = v("#fmDob"); if ($("#fmLoc")) m.loc = v("#fmLoc"); if ($("#fmProv")) m.prov = v("#fmProv"); if ($("#fmFrom")) m.from = v("#fmFrom"); if ($("#fmTo")) m.to = v("#fmTo");
    for (const k of ["sched", "consent", "avail"]) if ($(`#fp_${k}`)) m.parts[k] = $(`#fp_${k}`).checked;
    if ($("#fp_bidf")) m.bidf = $("#fp_bidf").checked; };
  function formsData(p, o) {
    const sched = C.scheduleOf(p, S.A, todayS()), c = S.cfg;
    const L = C.treatments(p, S.A), d = o.dob ? C.toSerial(o.dob) : null;
    /* the template is HTML (the Castro desk's email), so what goes in is escaped */
    const first = sched[0], html = /<[a-z]/i.test(S.tpl.startEmail), enc = v => html ? esc(v) : String(v);
    const nm = o.name || p.name, vals = { first: C.firstName(nm) || "there", clinic: o.loc || c.Clinic, patient: p.code, mrn: p.mrn, address: c.ClinicAddress, firstVisit: first ? C.longWhen(first.date, first.time) : "(no visit booked yet)" };
    const pt = ptOf(p.mrn), body = String(S.tpl.startEmail).replace(/\{(\w+)\}/g, (x, k) => vals[k] != null ? enc(vals[k]) : x);
    return { parts: o.parts, code: p.code, mrn: p.mrn, name: C.fullName(nm), dob: d != null ? `${C.dateOf(d).getMonth() + 1}/${C.dateOf(d).getDate()}/${C.dateOf(d).getFullYear()}` : "",
      clinic: o.loc || c.Clinic, provider: o.prov || "", asOf: `As of ${C.fmtDay(todayS())}, ${C.dateOf(todayS()).getFullYear()}`, sch: schedObj(p, o.from, o.to), index: +c.OakIndex, total: L.length,
      avail: pt ? { grid: pt.avail, eff: pt.availEff != null ? C.fmtD(pt.availEff) : "", exp: pt.availExp != null ? C.fmtD(pt.availExp) : "", note: pt.availNote } : null,
      email: o.email ? { subject: C.fill(S.tpl.startSubject, { clinic: c.Clinic, patient: p.code }), body, html, cc: C.txt(c.StartCc), bcc: C.txt(c.StartBcc) } : null };
  }
  /* the patient the forms are for: their tracker row, or a new index patient planned by LASFIR and MRN (no name or DOB on file) */
  function formsPatient(mrn) {
    const p = S.PM.get(mrn); if (p) return Object.assign({ onTracker: true }, p);
    const o = ipList().find(x => x.mrn === mrn) || S.WL?.list.find(x => x.mrn === mrn);
    return o ? { onTracker: false, mrn, code: o.code, name: "", dob: null, status: "", start: null, prefs: [], loc: "", freq: "" } : null;
  }
  function formsModal(mrn, mode, parts) {
    const p = formsPatient(mrn); if (!p) return;
    const sv = startVisit(p) || C.futureOf(S.A, p.mrn, todayS() - 1)[0];
    /* a start packet covers the rest of the course; a schedule on its own starts at the next 14 days, as in the TMS Tracker */
    const t = todayS(), all = schedVisits(p, t, null), to = mode === "email" || parts.consent ? (all.length ? all[all.length - 1].date : t) : t + 14;
    openModal("forms", { mrn, mode, parts, dob: p.dob != null ? C.isoOf(p.dob) : "", loc: S.cfg.Clinic, prov: sv ? C.provFull(sv.provider) : C.provFull(S.providers[0] || ""), from: C.isoOf(t), to: C.isoOf(to) });
  }
  const ACT2 = {
    edit: e => openEdit("mpt", e.dataset.m),
    editConsult: e => { if (e.dataset.m) openEdit("consults", e.dataset.m); },
    open: e => openEdit("mpt", e.dataset.m),
    openConsult: e => openEdit("consults", e.dataset.m),
    async doEdit() {
      const m = S.m; if (!m.row) return;
      const patch = {};
      for (const f of editFields(m)) { if (f.formula) continue; const el = $(`#ef_${f.i}`); if (!el) continue; const nv = el.value.trim(); if (nv === f.cur) continue;
        if (f.type === "date") patch[f.h] = { v: nv ? C.toSerial(nv) : "", date: true };
        else if (f.type === "num" || (/^-?\d+(\.\d+)?$/.test(nv) && f.cur === "")) patch[f.h] = { v: nv === "" ? "" : C.num(nv) ?? nv };
        else patch[f.h] = { v: nv }; }
      if (!Object.keys(patch).length) return closeModal();
      const r = await write("Saving…", () => Store.saveRow(m.kind, m.mrn, patch)); if (!r) return;
      closeModal(); toast(`Saved ${plural(r.wrote.length, "field")}${r.skipped.length ? `; left ${r.skipped.join(", ")} alone (a formula or a missing column)` : ""}.`, !!r.skipped.length);
    },
    mPay: e => { keepPay(); S.m.pay = Object.assign({}, S.m.pay, { [e.dataset.p || "prov"]: e.dataset.v }); render(); },
    payEdit: e => { const v = visitBy(e.dataset.k); if (!v || !v.p) return; if (!S.payOn) return toast("Payments need the newer tables. Use Update the workbook in Settings.", true);
      openModal("payEdit", Object.assign({ key: v.key }, payInit(v))); },
    async doPayEdit() {
      keepPay(); const m = S.m, v = visitBy(m.key); if (!v) return closeModal();
      if (!payRowsOf(v, m).length && !(m.amtSave && PARTS.some(k => amtN(m, k) !== trackerDue(v.p, k)))) return closeModal();
      closeModal(); if (await savePay(v, m, null)) toast(`${C.code(v)}: today’s payment saved.`);
    },
    /* transfers */
    xfer: e => { const p = S.PM.get(e.dataset.m); if (!p) return; if (!S.xferOn) return toast("Transfers need the newer table. Use Update the workbook in Settings.", true);
      const xf = S.XF.get(p.mrn), xs = C.xferState(xf, S.A, todayS(), S.x.E), oak = C.oakState(p, ptOf(p.mrn), S.A, todayS(), S.cfg);
      const to = xs?.to || (C.elsewhere(p, S.cfg) ? p.loc : oak ? "Oakland" : "");
      openModal("xfer", { mrn: p.mrn, mode: xs ? (xs.due ? "now" : "later") : "later", to, date: xs?.date != null ? C.isoOf(xs.date) : "", note: xf?.status === "scheduled" ? xf.note : "", oak: !!oak && ptOf(p.mrn)?.oakTasked == null }); },
    xfMode: e => { keepXfer(); S.m.mode = e.dataset.v; render(); },
    async doXfer() {
      keepXfer(); const m = S.m, p = S.PM.get(m.mrn), to = C.txt(m.to), d = C.toSerial(m.date);
      if (!to) return toast("Say which clinic they’re transferring to.", true);
      if (m.mode === "later" && (d == null || d <= todayS())) return toast("Pick a transfer date after today, or choose Discontinue now.", true);
      const oakTask = /oak/i.test(to) && m.oak && ptOf(p.mrn)?.oakTasked == null; closeModal();
      if (oakTask) await write("Saving…", () => Store.savePt(p.mrn, ptPatch(p, { OakTasked: C.isoOf(todayS()), Oakland: "TRUE" })));
      if (m.mode === "now") return confirmXfer(p, to, C.txt(m.note));
      await write("Scheduling the transfer…", () => Store.saveXfer(p.mrn, { Patient: p.code, To: to, Date: C.isoOf(d), Status: "scheduled", Note: C.txt(m.note), Confirmed: "", By: S.name, Updated: new Date().toISOString() }),
        `${p.code}: transfer to ${to} scheduled for ${C.fmtD(d)}${oakTask ? ", OAK tasked" : ""}. Confirm it at their last visit here.`);
    },
    xfCancel: () => { const p = S.PM.get(S.m.mrn); openModal("confirm", { title: `Cancel ${p.code}’s transfer?`, body: "They stay a patient here. Nothing changes on the Master Patient Tracker.", yes: "Cancel the transfer",
      run: () => Store.saveXfer(p.mrn, { Status: "cancelled", By: S.name, Updated: new Date().toISOString() }) }); },
    pay: e => { const p = S.PM.get(e.dataset.m); if (!p) return; if (!S.payOn) return toast("Payments need the newer tables. Use Update the workbook in Settings.", true); openModal("pay", { mrn: p.mrn }); },
    async doPay() {
      const p = S.PM.get(S.m.mrn), n = C.num($("#pyAmt").value.replace(/[$,\s]/g, "")), d = C.toSerial($("#pyDate").value);
      if (!(n > 0)) return toast("Enter the amount paid.", true);
      if (await write("Saving…", () => Store.addPays([payRow(p, { Date: C.isoOf(d ?? todayS()), Key: "", Kind: "paid", Amount: n, Method: $("#pyHow").value, Note: $("#pyNote").value.trim() })]), `${C.money(n)} recorded for ${p.code}.`)) closeModal();
    },
    async doRemind() {
      const p = S.PM.get(S.m.mrn), d = C.toSerial($("#rmDate").value) ?? todayS(), how = $("#rmHow").value;
      if (await write("Saving…", () => Store.addPays([payRow(p, { Date: C.isoOf(d), Kind: "reminder", Amount: C.ledger(p.mrn, S.pays).balance, Method: how })]), `Reminder recorded: ${C.fmtD(d)}, ${how}.`)) closeModal();
    },
    /* a missed visit settled later, or waived */
    payMissed: async e => { const p = S.PM.get(e.dataset.m), r = C.ledger(p.mrn, S.pays).open.find(x => x.id === e.dataset.id); if (!r) return;
      await write("Saving…", () => Store.addPays([payRow(p, { Date: C.isoOf(todayS()), Key: r.key, Kind: "paid", Amount: r.left, Method: "At the desk" })]), `${C.money(r.left)} collected for ${dshort(r.date)}.`); },
    payWaive: e => { const p = S.PM.get(e.dataset.m), r = C.ledger(p.mrn, S.pays).open.find(x => x.id === e.dataset.id); if (!r) return;
      openModal("confirm", { title: `Waive ${C.money(r.left)} for ${p.code}?`, body: `The ${dshort(r.date)} payment comes off their balance. The waiver is logged with your name.`, yes: "Waive it",
        run: () => Store.addPays([payRow(p, { Date: C.isoOf(todayS()), Key: r.key, Kind: "waived", Amount: r.left })]) }); },
    /* a past visit nobody answered: collected or not */
    payConf: async e => { const p = S.PM.get(e.dataset.m), d = +e.dataset.d, t = +e.dataset.t, k = C.visitKey(d, p.mrn, t), kind = e.dataset.v;
      const rows = C.dueParts(p).filter(([part]) => !C.payOf(S.pays, k, part)).map(([part, n]) => payRow(p, { Date: C.isoOf(d), Key: C.partKey(k, part), Kind: kind, Amount: n, Note: `${PART_WORD[part]}, confirmed later` }));
      if (!rows.length) return;
      await write("Saving…", () => Store.addPays(rows), `${dshort(d)}: ${kind === "collected" ? "collected" : "not collected, added to the balance"}.`); },
    cmDone: e => cmSet(e.dataset.m, e.dataset.s || "", true),
    oakTask: async e => { const p = S.PM.get(e.dataset.m); if (!p) return; await write("Saving…", () => Store.savePt(p.mrn, ptPatch(p, { OakTasked: C.isoOf(todayS()), Oakland: "TRUE" })), `${p.code}: OAK tasked to book maintenance, ${C.fmtD(todayS())}.`); },
    oakUntask: async e => { const p = S.PM.get(e.dataset.m); if (!p) return; await write("Saving…", () => Store.savePt(p.mrn, ptPatch(p, { OakTasked: "" }))); },
    avail: e => {
      const p = S.PM.get(e.dataset.m); if (!p) return; if (!S.raw.tables?.own?.tblIvkPatient) return toast("Availability needs the newer tables. Use Update the workbook in Settings.", true);
      const t = ptOf(p.mrn), g = t?.avail || p.prefs.map(x => ({ am: x.am ? "M" : "", pm: x.pm ? "M" : "" }));
      openModal("avail", { mrn: p.mrn, grid: g.map(x => Object.assign({}, x)), eff: t?.availEff != null ? C.isoOf(t.availEff) : "", exp: t?.availExp != null ? C.isoOf(t.availExp) : "", note: t?.availNote || "",
        away: (t?.away || []).map(w => ({ from: C.isoOf(w.from), to: C.isoOf(w.to), note: w.note })) });
    },
    /* provider time off */
    ptoAdd: () => { const g = S.grids[Math.min(S.tg || 0, S.grids.length - 1)]; openModal("pto", { id: `x${Date.now().toString(36)}`, isNew: true, provider: g ? TM.nameOf(g) : "", from: C.isoOf(todayS() + 1), to: "", cov: "", useTpl: "", note: "" }); },
    ptoEdit: e => { const x = S.EX.find(y => y.id === e.dataset.v); if (!x) return;
      openModal("pto", { id: x.id, provider: x.provider, from: C.isoOf(x.from), to: x.to !== x.from ? C.isoOf(x.to) : "", cov: x.cov, useTpl: x.cov ? (x.useTpl ? "yes" : "no") : "", note: x.note }); },
    ptoTpl: e => { keepPto(); S.m.useTpl = e.dataset.v; render(); },
    async doPto() {
      keepPto(); const m = S.m, from = C.toSerial(m.from), to = m.to ? C.toSerial(m.to) : from;
      if (from == null) return toast("Pick the first day off.", true); if (to == null || to < from) return toast("The last day off is before the first.", true);
      if (m.cov && TM.provKey(m.cov) === TM.provKey(m.provider)) return toast("Pick someone else to cover.", true);
      if (m.cov && !m.useTpl) return toast(`Say whether ${exLabel(m.cov)} uses ${exLabel(m.provider)}’s template.`, true);
      const row = { Provider: m.provider, From: C.isoOf(from), To: C.isoOf(to), Coverage: m.cov || "", UseTemplate: m.cov && m.useTpl === "yes" ? "TRUE" : "FALSE", Note: C.txt(m.note), Status: "", By: S.name, Updated: new Date().toISOString() };
      const span = `${dshort(from)}${to !== from ? ` to ${dshort(to)}` : ""}`;
      if (await write("Saving time off…", () => Store.saveException(m.id, row), `${exLabel(m.provider)} off ${span}${m.cov ? `, covered by ${exLabel(m.cov)}${m.useTpl === "yes" ? " on the same template" : ""}` : ": slots closed"}.`)) closeModal();
    },
    ptoDel: () => { const m = S.m; openModal("confirm", { title: `Remove ${exLabel(m.provider)}’s time off?`, body: "Their template’s slots open again on those days.", yes: "Remove",
      run: () => Store.saveException(m.id, { Status: "removed", By: S.name, Updated: new Date().toISOString() }) }); },
    avSet: e => { keepAvail(); const g = S.m.grid[+e.dataset.d], h = e.dataset.h; g[h] = g[h] === e.dataset.v ? "" : e.dataset.v; render(); },
    avAddAway: () => { keepAvail(); S.m.away.push({ from: C.isoOf(todayS()), to: "", note: "" }); render(); },
    avDelAway: e => { keepAvail(); S.m.away.splice(+e.dataset.v, 1); render(); },
    async doAvail() {
      keepAvail(); const m = S.m, p = S.PM.get(m.mrn), eff = C.toSerial(m.eff), exp = C.toSerial(m.exp);
      if (eff != null && exp != null && exp < eff) return toast("The expiry date is before the effective date.", true);
      const away = m.away.map(w => ({ from: C.toSerial(w.from), to: C.toSerial(w.to) ?? C.toSerial(w.from), note: w.note })).filter(w => w.from != null);
      if (away.some(w => w.to < w.from)) return toast("One of the dates away ends before it starts.", true);
      const ok = await write("Saving availability…", async () => {
        await Store.savePt(p.mrn, ptPatch(p, { Avail: C.availStr(m.grid), AvailEffective: eff != null ? C.isoOf(eff) : "", AvailExpires: exp != null ? C.isoOf(exp) : "", AvailNote: m.note.trim(), Away: C.awayStr(away) }));
        return Store.saveRow("mpt", p.mrn, C.prefCells(m.grid, S.raw.mpt && S.raw.mpt[0]));
      });
      if (ok) { closeModal(); toast(`${p.code}’s availability saved${ok.skipped?.length ? `; the tracker’s ${ok.skipped.join(", ")} weren’t changed` : ", and the tracker’s AM/PM columns updated"}.`, !!ok.skipped?.length); }
    },
    printAvail: e => { if (S.modal === "avail") keepAvail(); formsModal(e.dataset.m || S.m.mrn, "print", { avail: true }); },
    printSched: e => formsModal(e.dataset.m, "print", { sched: true }),
    startForms: e => formsModal(e.dataset.m, "print", { sched: true, consent: true }),
    startEmail: e => formsModal(e.dataset.m, "email", { sched: true, consent: true }),
    doForms() {
      const m = S.m, p = formsPatient(m.mrn), parts = m.mode === "email" ? { sched: true, consent: true } : { sched: !!$("#fp_sched")?.checked, consent: !!$("#fp_consent")?.checked, avail: !!$("#fp_avail")?.checked };
      if (!Object.values(parts).some(Boolean)) return toast("Pick at least one form.", true);
      keepForms(); const from = C.toSerial(m.from), to = C.toSerial(m.to);
      if (parts.sched && (from == null || to == null || to < from)) return toast("Pick the schedule’s From and To dates.", true);
      const d = formsData(p, { parts, name: C.txt(m.name), dob: m.dob, loc: C.txt(m.loc), prov: m.prov, email: m.mode === "email", from, to });
      if (parts.sched && !d.sch.rows.length) return toast(`No scheduled visits in that date range${p.onTracker ? "" : ". If they're booked in NextGen, upload the NextGen report first"}.`, true);
      const bidf = m.mode !== "email" && parts.consent ? !!m.bidf : null;
      closeModal(); openForms(d);
      /* start paperwork: the BIDF is printed from NextGen; ticking it saves the day on the profile, leaving it reminds */
      if (bidf === false) toast("Remember the BIDF: print it from NextGen with the start paperwork.", true);
      else if (bidf && S.profOn) write("Saving…", () => Store.saveProfile(p.mrn, { Patient: p.code, BidfPrinted: C.isoOf(todayS()), By: S.name, Updated: new Date().toISOString() }));
    },
    fmRange: e => { keepForms(); S.m.from = C.isoOf(todayS()); S.m.to = e.dataset.v; render(); },
    printWeek: () => {
      const T = todayS(), start = C.weekStart(T) + 7 * S.wk, A = S.A.filter(a => a.date > start && a.date < start + 6 && !C.isCancelled(a)).sort((a, b) => a.date - b.date || (a.time ?? 0) - (b.time ?? 0));
      openForms({ parts: { week: true }, asOf: `Week of ${C.fmtD(start + 1)} to ${C.fmtD(start + 5)}, ${C.dateOf(start).getFullYear()}`, clinic: S.cfg.Clinic,
        week: A.map(a => ({ day: C.fmtDay(a.date), at: C.fmtT(a.time), code: a.code, provider: C.provName(a.provider), type: /start/i.test(a.event) ? "Start" : "Infusion", n: S.PM.get(a.mrn) ? C.txNo(S.PM.get(a.mrn), S.A, a.date, a.time) : null, status: a.status })) });
    },
    /* waitlist */
    wSel: e => { S.wSel = e.dataset.m; render(); follow("consults", S.wSel); },
    async wMove(e) {
      const mrn = e.dataset.m, v = e.dataset.v, L = S.WL.list, order = C.waitMove(L, mrn, v === "top" ? "top" : +v);
      await write("Saving the order…", () => Store.saveRanks(order.map(x => ({ mrn: x, code: L.find(w => w.mrn === x).code })), S.name));
    },
    async wOn(e) {
      const w = S.WL.list.find(x => x.mrn === e.dataset.m); if (!w) return;
      await write("Saving…", () => Store.savePt(w.mrn, Object.assign(ptPatch({ code: w.code }, { WaitingOn: e.dataset.v }), e.dataset.v === "auth" ? {} : { Appeal: "" })));
    },
    async wPlan(e) {
      const w = S.WL.list.find(x => x.mrn === e.dataset.m); if (!w) return;
      if (!ipList().find(o => o.mrn === w.mrn)) await write("Saving…", () => Store.savePlan(w.mrn, { Patient: w.code, Note: "", By: S.name, Updated: new Date().toISOString() }));
      S.ipSel = w.mrn; S.ipView = "plan"; S.view = "index"; render();
    },
  };
  function keepXfer() { const m = S.m; for (const [id, k] of [["xfTo", "to"], ["xfDate", "date"], ["xfNote", "note"]]) { const el = $(`#${id}`); if (el) m[k] = el.value; } const o = $("#xfOak"); if (o) m.oak = o.checked; }
  const xferNote = m => `${C.fmtD(todayS())}, ${C.dateOf(todayS()).getFullYear()}: transferred to ${C.txt(m.to) || "another clinic"}${C.txt(m.note) ? ` (${C.txt(m.note)})` : ""}${S.name ? `, ${S.name}` : ""}`;
  /* the transfer done: Status Discontinued and the note added to Clinic Notes on the tracker, the transfer row done, a log line */
  async function confirmXfer(p, to, note) {
    const H = mptHead(), line = xferNote({ to, note }), nc = H("clinicnotes");
    const ok = await write("Saving the transfer…", async () => {
      const r = await Store.saveRow("mpt", p.mrn, Object.assign({ [H("status") || "Status"]: "Discontinued" }, nc ? { [nc]: p.notes ? `${p.notes}\n${line}` : line } : {}));
      if (S.xferOn) await Store.saveXfer(p.mrn, { Patient: p.code, To: to, Status: "done", Note: note, Confirmed: C.isoOf(todayS()), By: S.name, Updated: new Date().toISOString() });
      if (S.raw.tables?.own?.tblIvkLog) { try { await Store.appendLog([`x${Date.now().toString(36)}`, "transfer", p.code, "", line, S.name, new Date().toISOString(), "saved"]); } catch (e) { console.error(e); } }
      return r;
    }, r => `${p.code}: discontinued here, transferred to ${to}.${r?.skipped?.length ? ` The tracker’s ${r.skipped.join(", ")} wasn’t changed (a formula or no such column).` : ""}`);
    return !!ok;
  }
  /* waitlist: an appeal moves them behind the next ready patient; turning it off leaves them where they are */
  async function wAppeal(t) {
    const w = S.WL.list.find(x => x.mrn === t.dataset.m); if (!w) return;
    await write("Saving…", async () => {
      await Store.savePt(w.mrn, ptPatch({ code: w.code }, { Appeal: t.checked ? "TRUE" : "", WaitingOn: "auth" }));
      if (t.checked) { const L = S.WL.list, order = C.waitMove(L, w.mrn, "appeal"); await Store.saveRanks(order.map(x => ({ mrn: x, code: L.find(y => y.mrn === x).code })), S.name); }
    }, t.checked ? `${w.code} moved behind the next ready patient.` : null);
  }
  /* the start checklist: Case Management updated in NextGen (or undone). The start is found by key, else the patient's open one,
     else their kept start from the Consults tab. Saved on tblIvkStarts, and as CM Updated on the Consults row when there is one. */
  async function cmSet(mrn, key, on) {
    const T = todayS(), c = S.CS.find(x => x.mrn === mrn) || null;
    const s = (key && S.SL.find(x => x.key === key)) || S.SL.find(x => x.mrn === mrn && (on ? x.cmDue : x.cmDone)) || null;
    const date = s ? s.date : c?.kept ?? null; if (date == null) { render(); return toast("No treatment start found for that patient.", true); }
    const k = s ? s.key : C.startKey(date, mrn), code = s?.code || c?.code || S.PM.get(mrn)?.code || "";
    if (!S.stOn && !c) { render(); return toast("Ticking Case Management needs the newer table. Use Update the workbook in Settings.", true); }
    const ok = await write("Saving…", async () => {
      if (S.stOn) await Store.saveStart(k, { Date: C.isoOf(date), MRN: mrn, Patient: code, CmUpdated: on ? C.isoOf(T) : "", By: S.name, Updated: new Date().toISOString() });
      return c ? Store.saveRow("consults", mrn, { "CM Updated": on ? { v: T, date: true } : { v: "" } }) : true;
    });
    if (ok) toast(on ? `${code}: Case Management marked updated (${cmStage()}), ${C.fmtD(T)}.` : `${code}: Case Management tick taken off.`);
  }
  async function ptFlag(t) {
    const p = S.PM.get(t.dataset.m); if (!p) return; const k = t.dataset.pt, on = t.checked;
    const words = { CardOnFile: ["Luma card on file", "No Luma card on file"], OkToCollect: ["Okay to charge at time of service", "Not okay to charge at time of service"], Oakland: ["Oakland patient", "Not an Oakland patient"] }[k];
    await write("Saving…", () => Store.savePt(p.mrn, ptPatch(p, { [k]: on ? "TRUE" : "FALSE" })), `${p.code}: ${words[on ? 0 : 1].toLowerCase()}.`);
  }
  const ACT = {
    nav: e => { S.view = e.dataset.v === "access" && inAccess(S.view) ? "access" : e.dataset.v === "access" ? (S.acLast || "access") : e.dataset.v; if (inAccess(S.view)) S.acLast = S.view; S.pv = null; render(); },
    async reload() { S.reading = true; render(); try { await refresh(true); toast(S.err ? S.err : `Up to date. Read ${plural(S.list.length, "visit")} today and ${plural(S.P.length, "patient")} at ${hm()}.`, !!S.err); } finally { S.reading = false; render(); } },
    theme: () => { S.theme = S.theme === "dark" ? "light" : "dark"; LS.set("ivkTheme", S.theme); document.documentElement.dataset.theme = S.theme; render(); },
    full: () => openFull(),
    sel: e => { S.sel = e.dataset.k; render(); const v = visitBy(S.sel); if (v) follow("mpt", v.a.mrn); },
    bSel: e => { const c = floor().chairs.find(x => x.id === e.dataset.b); if (c && c.v) { S.sel = c.v.key; render(); follow("mpt", c.v.a.mrn); } else toast(`${chairTxt(e.dataset.b)} is open.`); },
    pSel: e => { S.pSel = e.dataset.m; render(); follow("mpt", S.pSel); },
    chip: e => { S.chip = e.dataset.v; render(); },
    wk: e => { const d = +e.dataset.v; S.wk = d === 0 ? 0 : S.wk + d; render(); },
    /* a booked visit on the Schedule opens the patient view; someone not on the tracker yet gets the profile form */
    schedPick: e => { const mrn = e.dataset.m, d = +e.dataset.d, t = e.dataset.t === "" ? null : +e.dataset.t, a = S.A.find(x => x.mrn === mrn && x.date === d && (t == null || x.time === t)); if (!a) return;
      const r = S.SO.get(mrn), on = r && r.status === "waiting" && r.booked === d, p = S.PM.get(mrn);
      /* days: what they said before, else the availability they gave (Preferred or Mostly), else any */
      const av = ptOf(mrn)?.avail || (p ? p.prefs.map(x => ({ am: x.am ? "M" : "", pm: x.pm ? "M" : "" })) : []);
      const days = on ? [...r.days] : av.flatMap((x, i) => ["am", "pm"].filter(h => C.canCome(x[h])).map(h => `${i + 1}${h}`));
      openModal("sooner", { mrn, code: a.code, date: d, time: a.time, provider: a.provider, from: on && r.from != null ? C.isoOf(r.from) : C.isoOf(todayS() + 1), days, same: on ? r.sameProv : false, note: on ? r.note : "" }); },
    soPv: () => { const m = S.m.mrn; closeModal(); if (S.PM.get(m)) openPv(m); else ACT.profile({ dataset: { m } }); },
    soDay: e => { keepSooner(); const v = e.dataset.v, L = S.m.days; S.m.days = L.includes(v) ? L.filter(x => x !== v) : L.concat(v); render(); },
    soEdit: e => { const r = S.SO.get(e.dataset.m), v = r && soonerVisit(r); if (!v) return toast("That visit isn’t on the appointment list any more. Take them off the waitlist, or waitlist the new visit.", true);
      ACT.schedPick({ dataset: { m: r.mrn, d: String(v.date), t: String(v.time ?? "") } }); },
    async soSave() {
      keepSooner(); const m = S.m, from = C.toSerial(m.from); if (needName()) return;
      if (from != null && from >= m.date) return toast("The earliest date has to be before their visit.", true);
      const ok = await write("Saving…", () => Store.saveSooner(m.mrn, { Patient: m.code, Booked: C.isoOf(m.date), BookedTime: C.fmtT(m.time), Provider: m.provider, From: from != null ? C.isoOf(from) : "",
        Days: m.days.slice().sort().join(","), SameProv: m.same ? "TRUE" : "FALSE", Status: "waiting", Note: C.txt(m.note), Seen: "", By: S.name, Updated: new Date().toISOString() }));
      if (ok) { closeModal(); const x = soonerHits().find(y => y.r.mrn === m.mrn); toast(`${m.code}: on the sooner waitlist${x && x.hits.length ? `; ${plural(x.hits.length, "opening")} already, see Sooner waitlist under the schedule` : ""}.`); }
    },
    async soDone(e) {
      const mrn = e.dataset.m, r = S.SO.get(mrn); if (!r) return;
      if (await write("Saving…", () => Store.saveSooner(mrn, { Status: "done", By: S.name, Updated: new Date().toISOString() }), `${r.code}: off the sooner waitlist.`) && S.modal === "sooner") closeModal();
    },
    toToday: e => { S.sel = e.dataset.k; S.view = "today"; S.pv = null; render(); },
    openPt: e => openPv(e.dataset.m),
    pOpen: e => openPv(e.dataset.m),
    pvClose: () => { const m = S.pv; S.pv = null; render(); const b = document.querySelector(`[data-act="pOpen"][data-m="${CSS.escape(m || "")}"]`); if (b) b.focus(); },
    pvShow: e => { S.pvSect = e.dataset.v; LS.set("ivkPtSect", S.pvSect); render(); },
    pSort: e => { const k = e.dataset.v, s = S.sort || { k: "need", dir: -1 }; S.sort = s.k === k ? { k, dir: -s.dir } : { k, dir: k === "need" || k === "conf" || k === "thru" ? -1 : 1 }; render(); },
    avChip: e => { S.avChip = e.dataset.v; render(); },
    dmPaused: () => { S.dm = Object.assign({}, S.dm, { paused: !S.dm?.paused }); render(); },
    dmSched: e => { S.dm = Object.assign({}, S.dm, { sched: e.dataset.v }); render(); },
    missTip: e => { e.classList.toggle("show"); },
    async pvPlan(e) {
      const p = S.PM.get(e.dataset.m); if (!p) return;
      if (!ipList().find(o => o.mrn === p.mrn) && !(await write("Saving…", () => Store.savePlan(p.mrn, { Patient: p.code, Note: "", By: S.name, Updated: new Date().toISOString() })))) return;
      S.pv = null; S.ipSel = p.mrn; S.ipAlt = null; S.view = "index"; S.acLast = "index"; render();
    },
    /* the profile form: from the tracker row and the profile table, or the appointment list for someone not on the tracker */
    profile: e => {
      const mrn = e.dataset.m, p = S.PM.get(mrn), pf = S.PF.get(mrn) || null, g = S.gaps.find(x => x.mrn === mrn), a = g?.first || C.futureOf(S.A, mrn, todayS() - 1)[0] || null;
      const cv = p ? C.coverage(p, pf) : { prov: { cover: pf?.prov.cover || "", payer: pf?.prov.payer || "" }, inf: { cover: pf?.inf.cover || "", payer: pf?.inf.payer || "" }, mode: "" };
      const sh = pf?.sohar || {}, iso = d => d != null ? C.isoOf(d) : "";
      openModal("profile", { mrn, onMpt: !!p, code: p?.code || g?.code || a?.code || "", first: a, start: !!g?.start || /start/i.test(a?.event || ""),
        status: p?.status || (g?.start ? "Induction" : "Ongoing"), freq: p?.freq || "", startIso: iso(p?.start ?? (g?.start ? a?.date : null)),
        mode: cv.mode || "both", provPayer: cv.prov.payer, provDue: p ? C.txt(p.copay) : "", infPayer: cv.inf.payer, infDue: p ? C.txt(p.infusion) : "",
        sohar: { payer: sh.payer || "", plan: sh.plan || "", network: sh.network || "", status: sh.status || "", covStart: iso(sh.covStart), covEnd: iso(sh.covEnd), copay: sh.copay || "", coins: sh.coins || "",
          ded: sh.ded || "", dedRem: sh.dedRem || "", oop: sh.oop || "", oopRem: sh.oopRem || "", date: iso(sh.date), ref: sh.ref || "" } });
    },
    pfMode: e => { keepProfile(); S.m.mode = e.dataset.v; render(); },
    async doProfile() {
      keepProfile(); const m = S.m, code = C.txt(m.code).toUpperCase(), mp = C.modeParts(m.mode || "both");
      if (!/^[A-Z]{4,8}$/.test(code)) return toast("Type the LASFIR code: last three and first three letters, for example SMIJAN.", true);
      const amtOk = v => C.txt(v) === "" || C.num(C.txt(v).replace(/[$,\s]/g, "")) != null; if (!amtOk(m.provDue) || !amtOk(m.infDue)) return toast("What the patient pays has to be a number, like 40.", true);
      const money = v => C.txt(v) === "" ? "" : C.num(C.txt(v).replace(/[$,\s]/g, ""));
      const provIns = mp.prov === "selfpay" ? "Self Pay" : C.txt(m.provPayer), infIns = mp.inf === "selfpay" ? "Self Pay" : C.txt(m.infPayer);
      const H = mptHead(), mpt = { [H("copay")]: money(m.provDue), [H("infusion")]: money(m.infDue), [H("primaryencounterinsurance", "primaryinsurance")]: provIns, [H("infusioninsurance", "secondaryinsurance")]: infIns, [H("status")]: m.status, [H("frequency", "txfrequency")]: m.freq };
      const sh = m.sohar, prof = { Patient: code, ProvCover: mp.prov === "selfpay" ? "Self pay" : "Insurance", ProvPayer: provIns, InfCover: mp.inf === "selfpay" ? "Self pay" : "Insurance", InfPayer: infIns,
        SoharPayer: sh.payer, Plan: sh.plan, Network: sh.network, CovStart: sh.covStart, CovEnd: sh.covEnd, Copay: sh.copay, Coinsurance: sh.coins, Deductible: sh.ded, DedRemaining: sh.dedRem,
        OopMax: sh.oop, OopRemaining: sh.oopRem, SoharStatus: sh.status, SoharDate: sh.date, SoharRef: sh.ref, By: S.name, Updated: new Date().toISOString() };
      const ok = await write(m.onMpt ? "Saving the profile…" : "Adding to the tracker…", async () => {
        let r;
        if (m.onMpt) r = await Store.saveRow("mpt", m.mrn, Object.fromEntries(Object.entries(mpt).filter(([k]) => k)));
        else { const st = C.toSerial(m.startIso); r = await Store.addMpt(Object.assign({ "MRN": /^\d+$/.test(m.mrn) ? +m.mrn : m.mrn, "Patient Name": code, "Primary Location": S.cfg.Clinic },
          Object.fromEntries(Object.entries(mpt).filter(([k]) => k)), st != null ? { "Tx Start Date": { v: st, date: true } } : {})); }
        if (S.profOn) await Store.saveProfile(m.mrn, prof);
        return r;
      });
      if (!ok) return; closeModal();
      toast(`${code}: ${m.onMpt ? "profile saved" : "added to the Master Patient Tracker"}${ok.skipped?.length ? `; the tracker has no ${ok.skipped.join(", ")} column or it's a formula, so that was left alone` : ""}${S.profOn ? "" : ". Update the workbook to keep the Sohar benefits"}.`, !!ok.skipped?.length);
    },
    /* Schedule: an open slot held for one patient */
    schedOpen: e => { const d = +e.dataset.d, t = +e.dataset.t, pk = TM.provKey(e.dataset.p), s = C.daySlots(d, S.X).slots.find(x => x.time === t && x.pkey === pk); if (!s || !s.open) return toast("That time was just taken.", true);
      if (needName()) return; openModal("hold", { slot: s, who: e.dataset.w || "" }); },
    schedPlan: () => { const s = S.m.slot; closeModal(); S.ipSel = ""; S.ipPins[""] = { 0: asPin(s) }; S.ipAlt = null; S.view = "index"; S.acLast = "index"; render(); },
    async doHold() {
      const s = S.m.slot, w = $("#hdWho")?.value || S.m.who; let mrn = w, code = "";
      if (!w) return toast("Choose who the time is for.", true);
      if (w === "__new") { const r = newPt($("#hdCode")?.value, $("#hdMrn")?.value); if (!r) return; mrn = r.mrn; code = r.code; }
      else code = S.PM.get(w)?.code || ipList().find(o => o.mrn === w)?.code || S.WL.list.find(x => x.mrn === w)?.code || "";
      const fresh = C.daySlots(s.date, S.X).slots.find(x => x.time === s.time && x.pkey === s.pkey); if (!fresh || !fresh.open) { closeModal(); return toast("That time was just taken.", true); }
      /* on the planner too, so the hold can be marked booked in NextGen there */
      const ok = await write("Holding the time…", async () => { if (!ipList().find(o => o.mrn === mrn)) await Store.savePlan(mrn, { Patient: code, Note: "", By: S.name, Updated: new Date().toISOString() });
        return Store.placeHolds([{ MRN: mrn, Patient: code, Date: C.isoOf(s.date), Time: C.fmtT(s.time), Provider: s.provider, Treatment: "" }], S.name); });
      if (!ok) return; closeModal(); toast(`${dshort(s.date)} ${C.fmtT(s.time)} held for ${code}. Book it in NextGen, then press Booked in NextGen (Index planner) within ${S.cfg.HoldMin} min.`);
      await write("Filling the Index Planner tab…", () => Store.fillPlanner());
    },
    setup: async () => { const r = await write("Setting up…", () => Store.setup()); if (r && r.blocked?.length) toast(`Couldn’t add ${r.blocked.join(", ")}: cells on the ${X.DATA_SHEET} tab already hold something.`, true); else if (r) toast("Workbook set up. Check-ins and Teams settings now save here."); },
    /* a chair is only asked for when a chair map is loaded: the first open one, or none yet (picked at the start of treatment) */
    checkin: e => { if (needName()) return; const v = visitBy(e.dataset.k); if (!v) return; openModal("checkin", Object.assign({ key: v.key, chair: S.plan ? freeChair() : "", rn: true, ask: true }, payInit(v))); },
    /* Start treatment: straight away with no map or a chair already picked; else which chair */
    start: async e => {
      const v = visitBy(e.dataset.k); if (!v) return;
      if (S.plan && !v.st.chair) return openModal("seat", { key: v.key, chair: freeChair(), rn: true });
      const t = hm();
      await write("Saving…", () => Store.saveVisit(v.key, visitBase(v), Object.assign({ Seated: t, Arrived: v.st.arr != null ? C.fmtT(v.st.arr) : t }, stamp())), `${C.code(v)} is in treatment${S.plan && v.st.chair ? ` in ${chairTxt(v.st.chair)}` : ""}.`);
    },
    seat: e => { const v = visitBy(e.dataset.k); if (!v) return; openModal("seat", { key: v.key, chair: v.ph === "tx" ? "" : freeChair(), rn: true }); },
    mPick: e => { keepPay(); S.m.chair = e.dataset.b; render(); },
    mBook: e => { keepPay(); S.m.booked = e.dataset.v; render(); },
    close: () => closeModal(),
    scrim: (e, ev) => { if (ev.target === e) closeModal(); },
    async doCheckin() {
      keepPay(); const m = S.m, v = visitBy(m.key), chair = S.plan ? m.chair : "", rn = $("#mRn")?.checked, ask = $("#mAsk")?.checked, t = hm();
      if (!v) return closeModal();
      const ok = await write("Saving check-in…", () => Store.saveVisit(v.key, visitBase(v), Object.assign({ Arrived: t, Chair: chair, Seated: "", Out: "", Status: "" }, stamp()), "Arrived"));
      if (!ok) return;
      closeModal(); S.sel = v.key;
      await savePay(v, m, `${C.code(v)} is checked in`);
      const msgs = [`${C.code(v)} checked in${chair ? ` to ${chairTxt(chair)}` : ""}.`];
      if (rn) { const r = await sendArrival(v, chair); msgs.push(teamsWord(r, routeTo("arrival"))); }
      if (ask && v.p) { const r = await sendAsk(v.p, C.askText(v.p, v.ask, S.tpl, v.p.freq), v.p.freq, v.key); msgs.push(r.ok ? `Asked ${routeTo("protocol")}.` : r.reason); }
      toast(msgs.join(" "), msgs.some(x => /answered|couldn’t|No Teams|off/i.test(x)));
    },
    /* the chair picked at the start of treatment (it starts now), or a move to another chair (the start time stays) */
    async doSeat() {
      const v = visitBy(S.m.key), chair = S.m.chair, rn = $("#mRn")?.checked, t = hm(); if (!v || !chair) return;
      const moving = v.ph === "tx";
      const ok = await write("Saving…", () => Store.saveVisit(v.key, visitBase(v), Object.assign({ Chair: chair, Seated: moving ? (C.fmtT(v.st.seat) || t) : t, Arrived: v.st.arr != null ? C.fmtT(v.st.arr) : t }, stamp())));
      if (!ok) return; closeModal(); S.sel = v.key;
      let msg = moving ? `${C.code(v)} is in ${chairTxt(chair)}.` : `${C.code(v)} is in treatment in ${chairTxt(chair)}.`;
      if (rn) { const r = await sendArrival(v, chair, true); msg += " " + teamsWord(r, routeTo("arrival")); }
      toast(msg);
    },
    async remind(e) {
      if (needName()) return; const v = visitBy(e.dataset.k); if (!v) return;
      const r = await sendArrival(v, S.plan ? v.st.chair || "" : ""); toast(teamsWord(r, routeTo("arrival")), !r.ok); render();
    },
    ask: e => {
      if (needName()) return;
      const v = e.dataset.k ? visitBy(e.dataset.k) : null, p = v ? v.p : S.PM.get(e.dataset.m); if (!p) return;
      const a = v ? v.ask : C.asks(p, S.x, null), use = a.any ? a : { book: true, proto: true };
      openModal("ask", { mrn: p.mrn, key: v ? v.key : null, freq: p.freq, ask: use, text: C.askText(p, use, S.tpl, p.freq) });
    },
    async doAsk() {
      const p = S.PM.get(S.m.mrn), text = $("#mText").value.trim(), freq = $("#mFreq").value; if (!p || !text) return;
      S.busy = "Posting to Teams…"; render();
      let r;
      try { r = await sendAsk(p, text, freq, S.m.key); }
      catch (e) { console.error(e); r = { ok: true, saveErr: e.message || String(e) }; }
      finally { S.busy = null; }
      if (r.ok) closeModal(); else render();
      toast(r.saveErr ? `Posted to ${routeTo("protocol")}, but the ask time didn’t save (${r.saveErr}). Don’t post again; it shows as asked from the Teams log.` : teamsWord(r, routeTo("protocol")), !r.ok || !!r.saveErr);
    },
    copyAsk: async () => { try { await navigator.clipboard.writeText($("#mText").value); toast("Message copied. Paste it in the group chat."); } catch (e) { $("#mText").select(); toast("Select the message and copy it with Ctrl+C."); } },
    checkout: e => {
      const v = visitBy(e.dataset.k); if (!v) return;
      /* the visit that completes index (an Induction patient's 6th kept visit): offer the move to Ongoing with the after-index protocol */
      const ix = v.p ? C.indexState(v.p, S.A, todayS(), S.cfg, a => a === v.a || (a.mrn === v.a.mrn && a.date === v.a.date && a.time === v.a.time)) : null;
      const idx = ix && ix.complete ? ix : null;
      /* a scheduled transfer is confirmed at the last visit here: ticked when this is it (or the date has come) */
      const xs = v.p ? C.xferState(S.XF.get(v.p.mrn), S.A, todayS(), S.x.E) : null, lastHere = !!xs && (xs.due || (xs.last && xs.last.date === v.a.date && xs.last.time === v.a.time));
      openModal("checkout", Object.assign({ key: v.key, booked: v.ask.book ? "" : "yes", proto: !!idx, freq: idx ? S.cfg.AfterIndex : v.p?.freq || "", idx, idxOn: !!idx, xs, xferOn: lastHere }, payInit(v)));
    },
    async doCheckout() {
      keepPay(); const m = S.m, v = visitBy(m.key); if (!v || !m.booked) return;
      const toOngoing = !!m.idx && !!$("#mIdx")?.checked, proto = (!!$("#mProto")?.checked || toOngoing) && !!v.p, freq = $("#mFreq")?.value || "", t = hm();
      if (proto && !freq) return toast("Pick the protocol the provider confirmed, or untick the box.", true);
      const ok = await write("Checking out…", () => Store.saveVisit(v.key, visitBase(v), Object.assign({ Out: t, Booked: S.m.booked, Protocol: proto ? freq : v.st.proto, Arrived: v.st.arr != null ? C.fmtT(v.st.arr) : t }, stamp())),
        proto ? null : `${C.code(v)} checked out.`);
      if (!ok) return;
      closeModal();
      await savePay(v, m, `${C.code(v)} is checked out`);
      if (m.xs && m.xferOn) await confirmXfer(v.p, m.xs.to, S.XF.get(v.p.mrn)?.note || "");
      if (proto) (await write("Saving the protocol…", () => Store.confirmProtocol(v.a.mrn, freq, todayS()), `${C.code(v)} checked out; protocol ${freq} saved on the Master Patient Tracker.`) && logConfirm(v.p, freq))
        || toast(`${C.code(v)} is checked out, but the protocol didn’t save to the Master Patient Tracker. Use Patients › Confirm protocol.`, true);
      if (toOngoing && proto) await write("Moving to Ongoing…", () => Store.saveRow("mpt", v.a.mrn, { [mptHead()("status") || "Status"]: "Ongoing" }), `${C.code(v)}: index done, now Ongoing at ${freq}.`)
        || toast(`${C.code(v)} is checked out, but the status didn’t change to Ongoing. Use Move to Ongoing on the patient.`, true);
    },
    async markBooked(e) { const v = visitBy(e.dataset.k); if (!v) return; await write("Saving…", () => Store.saveVisit(v.key, visitBase(v), Object.assign({ Booked: "yes" }, stamp())), `${C.code(v)}’s next visit marked booked.`); },
    gone: e => { const v = visitBy(e.dataset.k); openModal("confirm", { title: `Take ${C.code(v)} off today’s list?`, body: "NextGen isn’t changed. Cancel or reschedule the visit there too.", yes: "Take off today’s list", run: () => Store.saveVisit(v.key, visitBase(v), Object.assign({ Status: "cancelled" }, stamp())) }); },
    undoGone: async e => { const v = visitBy(e.dataset.k); await write("Saving…", () => Store.saveVisit(v.key, visitBase(v), Object.assign({ Status: "" }, stamp()))); },
    undoIn: e => { const v = visitBy(e.dataset.k); openModal("confirm", { title: `Undo ${C.code(v)}’s check-in?`, body: "They go back to Still to come. A Teams post already sent stays in the chat.", yes: "Undo check-in", run: () => Store.saveVisit(v.key, visitBase(v), Object.assign({ Arrived: "", Chair: "", Seated: "" }, stamp())) }); },
    undoSeat: async e => { const v = visitBy(e.dataset.k); await write("Saving…", () => Store.saveVisit(v.key, visitBase(v), Object.assign({ Seated: "" }, stamp())), `${C.code(v)} is back to checked in.`); },
    undoOut: async e => { const v = visitBy(e.dataset.k); await write("Saving…", () => Store.saveVisit(v.key, visitBase(v), Object.assign({ Out: "", Booked: "" }, stamp())), `${C.code(v)} is back ${v.st.seat != null ? "in treatment" : "to checked in"}.`); },
    async doConfirm() { const run = S.m.run; closeModal(); await write("Saving…", run); },
    proto: e => { const p = S.PM.get(e.dataset.m); if (p) openModal("proto", { mrn: p.mrn, freq: p.freq || "" }); },
    idxDone: e => { const p = S.PM.get(e.dataset.m); if (p) openModal("idxDone", { mrn: p.mrn, freq: S.cfg.AfterIndex }); },
    async doIdxDone() {
      const p = S.PM.get(S.m.mrn), f = $("#mFreq").value; if (!p || !f) return toast("Pick the protocol after index.", true);
      const ok = await write("Moving to Ongoing…", async () => { await Store.confirmProtocol(p.mrn, f, todayS()); return Store.saveRow("mpt", p.mrn, { [mptHead()("status") || "Status"]: "Ongoing" }); }, `${p.code}: index done, now Ongoing at ${f}.`);
      if (ok) { closeModal(); await logConfirm(p, f); }
    },
    async doProto() { const p = S.PM.get(S.m.mrn), f = $("#mFreq").value; if (!f) return toast("Pick the protocol the provider confirmed.", true); const ok = await write("Saving protocol…", () => Store.confirmProtocol(p.mrn, f, todayS()), `${p.code}’s protocol saved: ${f}, confirmed ${C.fmtD(todayS())}.`); if (ok) { closeModal(); await logConfirm(p, f); } },
    ...ACT2,
    saveName: () => { S.name = $("#myName").value.trim().split(/\s+/)[0] || ""; LS.set("ivkName", S.name); toast(S.name ? `Saved. Posts will say ${S.name}.` : "Name cleared."); render(); },
    async saveRoutes() {
      const r = {}; for (const ev of ["arrival", "protocol"]) r[ev] = { to: $(`#rt_${ev}_to`).value.trim(), url: $(`#rt_${ev}_url`).value.trim(), on: $(`#rt_${ev}_on`).checked };
      const bad = Object.entries(r).filter(([, x]) => x.url && !N.flowUrl(x.url)); if (bad.length) return toast("One of the URLs isn’t a Power Automate workflow URL. Copy it again from Teams.", true);
      await write("Saving Teams settings…", async () => { await Store.saveRoutes(r); await Store.saveKeyed("tblIvkConfig", { Enabled: $("#nfyOn").checked ? "TRUE" : "FALSE" }); }, "Teams settings saved.");
    },
    async testRoute(e) {
      const ev = e.dataset.v, url = $(`#rt_${ev}_url`).value.trim(); if (!N.flowUrl(url)) return toast("Paste the workflow URL first.", true);
      const r = await N.test(ev, { [ev]: { url, on: true } }, S.name);
      toast(r.ok ? (r.unconfirmed ? "Sent. The browser couldn’t read the workflow’s answer; check the chat." : `The workflow accepted the test (${r.status}). It should show in the chat within a minute.`) : r.reason, !r.ok);
    },
    async saveTpl() { const o = {}; const D = Object.assign({}, C.TEMPLATES, C.START_TPL); for (const k of Object.keys(D)) { const el = $(`#tp_${k}`); if (el) o[k] = el.value.trim() || D[k]; }
      const cf = {}; for (const k of ["ClinicAddress", "StartCc", "StartBcc"]) { const el = $(`#cf_${k}`); if (el) cf[k] = el.value.trim(); }
      const bad = ["StartCc", "StartBcc"].find(k => cf[k] && !cf[k].split(/[;,]\s*/).every(a => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(a))); if (bad) return toast(`${bad === "StartCc" ? "CC" : "BCC"} needs email addresses, separated by commas.`, true);
      await write("Saving messages…", async () => { await Store.saveKeyed("tblIvkTemplates", o); await Store.saveKeyed("tblIvkConfig", cf); }, "Messages saved."); },
    async resetTpl() { await write("Saving messages…", () => Store.saveKeyed("tblIvkTemplates", Object.assign({}, C.TEMPLATES, C.START_TPL)), "Standard wording restored."); },
    async saveCfg() {
      const o = {}; for (const k of ["VisitMin", "AheadWeeks", "ProtocolDays", "AuthWarnDays", "EncountersWarn"]) { const n = C.num($(`#cf_${k}`).value); if (n == null || n < 0) return toast(`${k} needs a number.`, true); o[k] = n; }
      o.Clinic = $("#cf_Clinic").value.trim() || C.DEFAULTS.Clinic;
      if (!S.plan) { const n = C.num($("#cf_Chairs").value); if (n == null || n < 1) return toast("Chairs in the room needs a number.", true); o.Chairs = String(n); }
      await write("Saving…", () => Store.saveKeyed("tblIvkConfig", o), "Timing saved.");
    },
    mapDown: () => { const P = S.raw.plan, a = document.createElement("a"); a.href = URL.createObjectURL(new Blob([JSON.stringify(P, null, 1)], { type: "application/json" })); a.download = `Chair map - ${S.cfg.Clinic} IVK.json`; a.click(); },
    mapRemove: () => openModal("confirm", { title: "Remove the chair map?", body: "Today shows chair tiles instead. You can load the file again any time.", yes: "Remove the map", run: () => Store.setDoc(MAP_KEY, null) }),
    async doUpload() {
      const u = S.m; if (u.sh.flag && !$("#mShrink")?.checked || !u.rows.length) return; closeModal();
      const wasPending = new Set(unconfirmed().map(h => h.id)), oldA = S.A.slice();   // before the new list is read
      const ok = await write("Writing the appointment list…", () => Store.replaceAppts(u.rows, C.serialOf(new Date()) + nowMin() / 1440)); if (!ok) return;
      if (S.raw.tables?.own?.tblIvkOther && !(await write("Saving the providers’ other visits…", () => Store.replaceOther(u.other))))
        toast("The appointment list is updated, but the providers’ other visits didn’t save, so flex post checks may show as free. Upload again.", true);
      /* a report with other clinics: their IVK visits replace tblIvkElsewhere, and the locations picked as this clinic's are kept */
      if (u.locs.length > 1) {
        const L = u.locs.filter(l => l.ours && l.loc).map(l => l.loc).join("; ");
        const kept = S.elseOn ? await write("Saving the other clinics’ visits…", () => Store.replaceElse(u.away)) : null;
        await write("Saving…", () => Store.saveKeyed("tblIvkConfig", Object.assign({ Locations: L }, kept ? { ElsewhereAt: C.isoOf(todayS()) } : {})));
        if (!S.elseOn) toast("Visits at the other clinics weren’t kept: use Update the workbook in Settings, then upload again.", true);
      }
      await write("Filling the Index Planner tab…", () => Store.fillPlanner(), r => `Appointment list updated: ${plural(u.rows.length, "appointment")}.${r && r.skipped ? ` Index Planner tab not filled (${r.skipped}).` : r ? ` Index Planner tab filled, ${C.fmtD(r.from)} to ${C.fmtD(r.to)}.` : ""}`)
        || toast("The appointment list is updated, but the Index Planner tab wasn’t filled. Upload again, or place or release a hold, to retry.", true);
      /* the booked index visits this report confirms, and any it doesn't have */
      if (wasPending.size) {
        const H = (S.holds || []).filter(h => wasPending.has(h.id)).map(h => C.holdConfirm(h, S.A, S.raw.apptUpdated)), ok = H.filter(c => c === "onlist").length, miss = H.filter(c => c === "missing").length;
        toast(miss ? `${plural(ok, "booked index visit")} confirmed; ${plural(miss, "visit")} marked booked ${miss === 1 ? "isn’t" : "aren’t"} on this report. Check them in NextGen (Index planner).` : `${ok === 1 ? "The booked index visit is" : `All ${ok} booked index visits are`} confirmed on the NextGen report.`, !!miss);
      }
      /* the sooner waitlist: openings on the new list (open slots, times cancellations freed) nobody has been shown yet */
      const fresh = soonerHits(C.freedTimes(oldA, S.A, todayS())).map(x => Object.assign({}, x, { hits: x.hits.filter(c => !x.r.seen.has(c.key)) })).filter(x => x.hits.length);
      if (fresh.length) {
        openModal("soonerFound", { list: fresh });
        for (const x of fresh) await write("Saving…", () => Store.saveSooner(x.r.mrn, { Seen: [...x.r.seen, ...x.hits.map(c => c.key)].slice(-60).join(";"), Updated: new Date().toISOString() }));
      }
    },
    /* index planner */
    ipView: e => { S.ipAlt = null; S.view = e.dataset.v === "template" ? "template" : "index"; S.acLast = S.view; render(); },
    /* pressing the patient already picked plans without one */
    ipPick: e => { const m = e.dataset.m; S.ipSel = S.view === "index" && S.ipSel === m ? "" : m; S.ipAlt = null; S.pv = null; S.view = "index"; S.acLast = "index"; render(); if (S.ipSel) follow("mpt", S.ipSel); },
    ipElig: e => { const x = nextEligible()[+e.dataset.i]; if (!x) return; const o = planTarget(); if (o && o.holds.length) return toast(`${o.code} already has holds. Book or release them first.`, true);
      S.ipPins[o ? o.mrn : ""] = { 0: asPin(x.slot) }; S.ipAlt = null; render(); },
    async ipHoldDraft() {
      if (needName()) return; const w = $("#ipFor")?.value || S.ipFor; let mrn = w, code = "";
      if (!w) return toast("Choose who these times are for.", true);
      if (w === "__new") { const r = newPt($("#ipForCode")?.value, $("#ipForMrn")?.value); if (!r) return; mrn = r.mrn; code = r.code; }
      else code = S.PM.get(w)?.code || ipList().find(o => o.mrn === w)?.code || S.WL.list.find(x => x.mrn === w)?.code || S.gaps.find(g => g.mrn === w)?.code || "";
      if (holdsOf(mrn).some(h => h.state === "held" || h.state === "late")) return toast(`${code} already has holds waiting. Book or release them first.`, true);
      const M = ipModel(DRAFT()), it = M.res.items.filter(x => !x.fixed), gone = it.filter(x => !pinOk(x));
      if (gone.length) { delete S.ipPins[""]; render(); return toast(`${dshort(gone[0].date)} at ${C.fmtT(gone[0].time)} was just taken. Pick a start again.`, true); }
      const rows = it.map(x => ({ MRN: mrn, Patient: code, Date: C.isoOf(x.date), Time: C.fmtT(x.time), Provider: x.provider, Treatment: x.n }));
      const ok = await write("Holding the times…", async () => { await Store.savePlan(mrn, { Patient: code, Start: C.isoOf(it[0].date), Note: "", By: S.name, Updated: new Date().toISOString() }); return Store.placeHolds(rows, S.name); });
      if (!ok) return; delete S.ipPins[""]; S.ipFor = ""; S.ipSel = mrn;
      toast(`${plural(rows.length, "time")} held for ${code}. Book them in NextGen, then press Booked in NextGen within ${S.cfg.HoldMin} min.`);
      await write("Filling the Index Planner tab…", () => Store.fillPlanner());
    },
    ipAlt: e => { const n = +e.dataset.n; S.ipAlt = S.ipAlt === n ? null : n; render(); },
    ipUse: e => {
      const o = planTarget(); if (!o) return;
      const M = ipModel(o), n = +e.dataset.n, s = C.altSlots(M.res, n, +S.cfg.MinGapDays)[+e.dataset.i]; if (!s) return;
      /* the visits before this one stay where they are; the ones after are planned again around it */
      const pins = {}; M.res.items.slice(0, n).forEach((x, k) => { if (!x.fixed) pins[k] = asPin(x); });
      pins[n] = asPin(s); S.ipPins[o.mrn] = pins; S.ipAlt = null; render();
    },
    /* an open index start pressed on the Access week: treatment 1, or the treatment being changed */
    ipTile: e => {
      const o = planTarget(); if (!o) return;
      if (o.holds.length) return toast(`${o.code} already has holds. Book or release them first.`, true);
      const d = +e.dataset.d, t = +e.dataset.t, pk = TM.provKey(e.dataset.p), s = C.daySlots(d, S.X).slots.find(x => x.time === t && x.pkey === pk);
      if (!s || !s.openIndex) return;
      const n = S.ipAlt ?? 0, M = ipModel(o);
      if (n > 0) {
        if (!C.altSlots(M.res, n, +S.cfg.MinGapDays).some(x => x.date === d && x.time === t && x.pkey === pk)) return toast(`That time isn’t ${S.cfg.MinGapDays}+ days from the visits either side of treatment ${n + 1}.`, true);
        const pins = {}; M.res.items.slice(0, n).forEach((x, k) => { if (!x.fixed) pins[k] = asPin(x); }); pins[n] = asPin(s); S.ipPins[o.mrn] = pins;
      } else S.ipPins[o.mrn] = { 0: asPin(s) };
      S.ipAlt = null; render();
    },
    ipWk: e => { const v = +e.dataset.v; S.ipWk = v === 0 ? 0 : Math.max(0, S.ipWk + v); render(); },
    ipReset: e => { delete S.ipPins[e.dataset.m || ""]; S.ipAlt = null; render(); },
    /* a new index patient, by LASFIR and MRN (they aren't on the Master Patient Tracker yet) */
    async ipNewPt() {
      const code = C.txt($("#ipNewCode")?.value).toUpperCase(), mrn = C.mrnKey($("#ipNewMrn")?.value);
      if (!/^[A-Z]{4,8}$/.test(code)) return toast("Type the LASFIR code: last three and first three letters, for example SMIJAN.", true);
      if (!/^\d{3,12}$/.test(mrn)) return toast("Type the MRN (numbers only).", true);
      const p = S.PM.get(mrn), have = ipList().find(o => o.mrn === mrn);
      if (p && p.code !== code) return toast(`MRN ${mrn} is ${p.code} on the Master Patient Tracker. Check the MRN or the code.`, true);
      const other = ipList().find(o => o.code === code && o.mrn !== mrn); if (other) return toast(`${code} is already being planned with MRN ${other.mrn}. Check the MRN.`, true);
      if (have) { S.ipSel = mrn; S.ipAlt = null; render(); return toast(`${code} is already on the planner.`); }
      if (await write("Saving…", () => Store.savePlan(mrn, { Patient: code, Note: "", By: S.name, Updated: new Date().toISOString() }), `${code} added. Pick their treatment 1, then hold the times.`)) { S.ipSel = mrn; S.ipAlt = null; render(); }
    },
    async ipAddPt() {
      const m = $("#ipAdd")?.value, p = S.PM.get(m); if (!p) return toast("Pick the patient in index first.", true);
      if (await write("Saving…", () => Store.savePlan(m, { Patient: p.code, Note: "", By: S.name, Updated: new Date().toISOString() }))) { S.ipSel = m; render(); }
    },
    async ipAv(e) {
      const o = ipList().find(x => x.mrn === S.ipSel); if (!o) return;
      const a = (C.availOf(o.pl || {}, o.p) || C.DAYS5.map(() => ({ am: true, pm: true }))).map(x => Object.assign({}, x)), d = +e.dataset.d, h = e.dataset.h;
      a[d][h] = !a[d][h];
      if (!a.some(x => x.am || x.pm)) return toast("Leave at least one morning or afternoon.", true);
      await savePlanOf(o, { Avail: a.every(x => x.am && x.pm) ? "any" : C.availText(a) });
    },
    ipAvAny: async () => { const o = ipList().find(x => x.mrn === S.ipSel); if (o) await savePlanOf(o, { Avail: "any" }); },
    ipAvReset: async () => { const o = ipList().find(x => x.mrn === S.ipSel); if (o) await savePlanOf(o, { Avail: "" }); },
    async ipHold(e) {
      if (needName()) return; const o = ipList().find(x => x.mrn === e.dataset.m); if (!o) return;
      const M = ipModel(o), it = M.res.items.filter(x => !x.fixed);
      const gone = it.filter(x => !pinOk(x));
      if (gone.length) { delete S.ipPins[o.mrn]; render(); return toast(`${dshort(gone[0].date)} at ${C.fmtT(gone[0].time)} was just taken. The plan is made again; check it and hold.`, true); }
      const rows = it.map(x => ({ MRN: o.mrn, Patient: o.code, Date: C.isoOf(x.date), Time: C.fmtT(x.time), Provider: x.provider, Treatment: x.n }));
      const ok = await write("Holding the times…", async () => { await Store.savePlan(o.mrn, { Patient: o.code, Start: C.isoOf(M.start), By: S.name, Updated: new Date().toISOString() }); return Store.placeHolds(rows, S.name); });
      if (!ok) return; delete S.ipPins[o.mrn];
      toast(`${plural(rows.length, "time")} held for ${o.code}. Book them in NextGen, then press Booked in NextGen within ${S.cfg.HoldMin} min.`);
      await write("Filling the Index Planner tab…", () => Store.fillPlanner());
    },
    async ipBooked(e) {
      const hs = holdsOf(e.dataset.m).filter(h => h.state !== "booked"); if (!hs.length) return;
      if (await write("Saving…", () => Store.setHolds(hs.map(h => h.id), "booked", S.name), `${hs[0].code}: ${plural(hs.length, "visit")} marked booked. Run the NextGen appointment report and upload it to confirm.`)) await write("Filling the Index Planner tab…", () => Store.fillPlanner());
    },
    ipRelease: e => {
      const hs = holdsOf(e.dataset.m).filter(h => h.state !== "booked" || h.conf === "missing"); if (!hs.length) return;
      openModal("confirm", { title: `Release ${hs[0].code}’s ${plural(hs.length, "hold")}?`, body: "The times open up for other patients. Anything already booked in NextGen stays booked there.", yes: "Release",
        run: async () => { await Store.setHolds(hs.map(h => h.id), "released", S.name); await Store.fillPlanner(); } });
    },
    ipDone: e => {
      const o = ipList().find(x => x.mrn === e.dataset.m); if (!o) return;
      openModal("confirm", { title: `Take ${o.code} off the Index planner?`, body: "Their holds and bookings stay. Add them again from Plan another patient.", yes: "Take off the list",
        run: () => Store.savePlan(o.mrn, { Patient: o.code, Note: "done", By: S.name, Updated: new Date().toISOString() }) });
    },
    /* template configurator */
    async tplSetup() { const r = await write("Setting up the template…", () => Store.setupTemplate()); if (r) toast(r.made.length ? `Added ${r.made.join(" and ")}.` : "The template tabs were already there."); },
    tplPick: e => { S.tg = +e.dataset.v; S.tsel = null; render(); },
    tplAdd: () => { const have = new Set(S.grids.map(g => g.key)), free = S.providers.filter(p => !have.has(TM.provKey(p)));
      openModal("addProv", { name: free[0] || "", other: !free.length, days: [1, 2, 3, 4, 5], start: 480, end: 1080, from: "", copy: "" }); },
    mDay: e => { const d = +e.dataset.d; keepAddProv(); S.m.days = S.m.days.includes(d) ? S.m.days.filter(x => x !== d) : S.m.days.concat(d).sort(); render(); },
    async doAddProv() {
      keepAddProv(); const m = S.m, name = C.txt(m.other ? m.typed : m.name);
      if (!name) return toast("Choose or type the provider.", true); if (!m.days.length) return toast("Pick the days they work.", true);
      if (m.end <= m.start) return toast("The last half hour has to be after the first.", true);
      const np = nameParts(name), src = m.copy !== "" ? S.grids[+m.copy] : null;
      const o = { title: np.title, label: np.label, location: S.cfg.Clinic, from: m.from ? C.toSerial(m.from) : null, modified: todayS(), days: m.days, start: m.start, end: m.end, step: 30,
        copy: src ? [...src.cells].filter(([k]) => m.days.includes(+k.split("|")[0])).map(([k, x]) => [k, x.type]) : [], r0: TM.nextRow(S.grids, S.types.length) };
      closeModal();
      const res = await write("Adding the provider…", () => Store.addGrid(o, S.types), `${np.title} added to the ${X.TPL_SHEET} tab.`);
      if (res) { S.tg = S.grids.findIndex(g => g.anchor[0] - 4 === res.r0); if (S.tg < 0) S.tg = S.grids.length - 1; render(); }
    },
    async tSet(e) {
      const g = curGrid(), q = S.tsel; if (!g || !q) return;
      const ds = g.dayCols.filter(x => x.d >= Math.min(q.d0, q.d1) && x.d <= Math.max(q.d0, q.d1)), rs = g.rows.filter(x => x.t >= Math.min(q.t0, q.t1) && x.t <= Math.max(q.t0, q.t1));
      if (!ds.length || !rs.length) return;
      const addr = `${TM.colL(ds[0].c)}${rs[0].r + 1}:${TM.colL(ds[ds.length - 1].c)}${rs[rs.length - 1].r + 1}`, v = e.dataset.v, mod = g.at.modified ? TM.A1(g.at.modified[0], g.at.modified[1]) : null;
      if (await write("Saving the template…", () => Store.setCells(addr, rs.length, ds.length, v, mod, todayS()), `${plural(rs.length * ds.length, "half hour")} ${v ? `set to ${v}` : "cleared"}.`)) { S.tsel = null; render(); }
    },
    tCancel: () => { S.tsel = null; render(); },
    teAdd: () => { S.te = (S.te || S.types.map(t => Object.assign({}, t))).concat({ type: "New slot type", kind: "Clinical", accepts: "", notes: "", use: "other", fill: "#E9EBE6", font: "#1B211F" }); render(); },
    teDel: e => { S.te = (S.te || S.types.map(t => Object.assign({}, t))).filter((t, i) => i !== +e.dataset.i); render(); },
    teUndo: () => { S.te = null; render(); },
    async teSave() {
      const te = S.te; if (!te) return; if (!te.length) return toast("Keep at least one slot type.", true);
      const names = te.map(t => C.txt(t.type).toLowerCase()); if (names.some(n => !n)) return toast("Every slot type needs a name.", true);
      if (new Set(names).size !== names.length) return toast("Two slot types have the same name.", true);
      const types = te.map(t => Object.assign({}, t, { type: C.txt(t.type), fill: String(t.fill).toUpperCase(), font: inkFor(t.fill) }));
      if (await write("Saving slot types…", () => Store.saveTypes(types, gridInfo()), `Slot types saved; ${plural(S.grids.length, "grid")} repainted.`)) { S.te = null; render(); }
    },
    async flagPost(e) {
      const p = S.PM.get(e.dataset.m); if (!p) return; const on = !p.reqPost;
      await write("Saving…", () => Store.saveFlag(p.mrn, { Patient: p.code, RequiresPostCheck: on ? "TRUE" : "FALSE", By: S.name, Updated: new Date().toISOString() }), on ? `${p.code} requires a post-check on every visit.` : `${p.code} no longer requires a post-check.`);
    },
    async savePayCfg() {
      const o = {}; for (const k of ["ReminderDays", "OakIndex", "OakRemindAt"]) { const n = C.num($(`#cf_${k}`).value); if (n == null || n < 1) return toast(`${k} needs a number of 1 or more.`, true); o[k] = n; }
      if (o.OakRemindAt > o.OakIndex) return toast("Task OAK at a treatment within the index.", true);
      const d = C.toSerial($("#cf_PayStart").value); o.PayStart = d != null ? C.isoOf(d) : "";
      await write("Saving…", () => Store.saveKeyed("tblIvkConfig", o), "Saved.");
    },
    async saveIdx() {
      const o = {}; for (const k of ["IndexPerWeek", "IndexWeeks", "MinGapDays", "MaxGapDays", "PostCheckMin", "HoldMin", "CapacityWeeks"]) { const n = C.num($(`#cf_${k}`).value); if (n == null || n < 0) return toast(`${k} needs a number.`, true); o[k] = n; }
      if (o.IndexPerWeek < 1 || o.IndexWeeks < 1) return toast("At least 1 visit a week for 1 week.", true);
      if (o.MinGapDays < 1 || o.MaxGapDays < o.MinGapDays) return toast("Days apart: at least 1, and the most can’t be less than the least.", true);
      o.AfterIndex = $("#cf_AfterIndex").value;
      await write("Saving…", () => Store.saveKeyed("tblIvkConfig", o), "Index protocol saved.");
    }
  };
  /* the add-a-provider form's fields, kept across re-renders */
  const keepAddProv = () => { if (S.modal !== "addProv") return; const v = id => $(id)?.value; const m = S.m;
    if ($("#apName")) m.name = v("#apName"); if ($("#apTyped")) m.typed = v("#apTyped"); if ($("#apStart")) m.start = +v("#apStart"); if ($("#apEnd")) m.end = +v("#apEnd");
    if ($("#apFrom")) m.from = v("#apFrom"); if ($("#apCopy")) m.copy = v("#apCopy"); };
  const savePlanOf = (o, patch) => write("Saving…", () => Store.savePlan(o.mrn, Object.assign({ Patient: o.code, By: S.name, Updated: new Date().toISOString() }, patch)));
  function openPv(mrn) { if (!S.PM.get(mrn)) return; if (!S.pv) S.pvFrom = S.view; S.pv = mrn; render(); $(".page") && ($(".page").scrollTop = 0); follow("mpt", mrn); }
  /* a new patient by LASFIR and MRN, checked against the tracker and the planner; null (with a toast) when it doesn't add up */
  function newPt(codeIn, mrnIn) {
    const code = C.txt(codeIn).toUpperCase(), mrn = C.mrnKey(mrnIn);
    if (!/^[A-Z]{4,8}$/.test(code)) { toast("Type the LASFIR code: last three and first three letters, for example SMIJAN.", true); return null; }
    if (!/^\d{3,12}$/.test(mrn)) { toast("Type the MRN (numbers only).", true); return null; }
    const p = S.PM.get(mrn); if (p && p.code !== code) { toast(`MRN ${mrn} is ${p.code} on the Master Patient Tracker. Check the MRN or the code.`, true); return null; }
    const other = ipList().find(o => o.code === code && o.mrn !== mrn); if (other) { toast(`${code} is already being planned with MRN ${other.mrn}. Check the MRN.`, true); return null; }
    return { code, mrn };
  }
  /* the tracker's own header for a field (the first of the spellings it has), or "" when it has none */
  const mptHead = () => { const H = (S.raw.mpt && S.raw.mpt[0] || []).map(h => [C.norm(h), h]); return (...keys) => { for (const k of keys) { const x = H.find(h => h[0] === k); if (x) return x[1]; } return ""; }; };
  /* every protocol confirmation goes in the log too, so the patient view can show when each one was */
  const logConfirm = async (p, freq) => { if (!S.raw.tables?.own?.tblIvkLog) return true; try { await Store.appendLog([`c${Date.now().toString(36)}`, "confirmed", p.code, "", freq, S.name, new Date().toISOString(), "saved"]); } catch (e) { console.error(e); } return true; };
  /* the profile form's fields, kept across re-renders */
  const keepProfile = () => { if (S.modal !== "profile") return; const m = S.m, v = id => $(id) ? $(id).value : null, set = (k, id) => { const x = v(id); if (x != null) m[k] = x; }, sh = (k, id) => { const x = v(id); if (x != null) m.sohar[k] = x; };
    set("code", "#pfCode"); set("status", "#pfStatus"); set("freq", "#pfFreq"); set("startIso", "#pfStart"); set("provPayer", "#pfProvPayer"); set("provDue", "#pfProvDue"); set("infPayer", "#pfInfPayer"); set("infDue", "#pfInfDue");
    sh("payer", "#pfSPayer"); sh("plan", "#pfSPlan"); sh("network", "#pfSNet"); sh("status", "#pfSStatus"); sh("covStart", "#pfSCovS"); sh("covEnd", "#pfSCovE"); sh("copay", "#pfSCopay"); sh("coins", "#pfSCoins");
    sh("ded", "#pfSDed"); sh("dedRem", "#pfSDedR"); sh("oop", "#pfSOop"); sh("oopRem", "#pfSOopR"); sh("date", "#pfSDate"); sh("ref", "#pfSRef"); };
  const elIs = s => eligibleSet(s.date, s.date).has(slotKey(s));
  /* the Sohar Eligibility add-in's benefits file: the row for this patient's MRN fills the Sohar fields (nothing is saved until Save) */
  async function readSohar(f) {
    keepProfile(); let rows = null;
    try {
      if (/\.csv$/i.test(f.name)) rows = csvRows(await f.text());
      else { if (!(await loadXLSX())) return toast("The spreadsheet reader didn’t load. Check the connection and try again.", true); const wb = XLSX.read(await f.arrayBuffer(), { type: "array" }); rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: false, defval: "" }); }
    } catch (e) { console.error(e); return toast("That file couldn’t be read.", true); }
    const M = C.parseSohar(rows); if (!M) return toast("That isn’t a Sohar benefits file. Use the benefits CSV from the Sohar Eligibility add-in (it has MRN and Deductible Remaining columns).", true);
    const r = M.get(C.mrnKey(S.m.mrn)); if (!r) return toast(`MRN ${S.m.mrn} isn’t in that file (${plural(M.size, "patient")} in it).`, true);
    const iso = d => d != null ? C.isoOf(d) : "";
    S.m.sohar = { payer: r.payer, plan: r.plan, network: r.network, status: r.status, covStart: iso(r.covStart), covEnd: iso(r.covEnd), copay: r.copay, coins: r.coins, ded: r.ded, dedRem: r.dedRem, oop: r.oop, oopRem: r.oopRem, date: iso(r.date), ref: r.ref };
    S.m.fromFile = `Filled from ${f.name}${r.date != null ? `, verified ${C.fmtD(r.date)}` : ""}. Check it, then save.`; render();
  }
  /* a small CSV reader: quoted fields, doubled quotes, commas and line breaks inside quotes */
  function csvRows(t) {
    const out = []; let row = [], cur = "", q = false;
    for (let i = 0; i < t.length; i++) { const ch = t[i];
      if (q) { if (ch === '"') { if (t[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += ch; }
      else if (ch === '"') q = true; else if (ch === ",") { row.push(cur); cur = ""; } else if (ch === "\n" || ch === "\r") { if (ch === "\r" && t[i + 1] === "\n") i++; row.push(cur); out.push(row); row = []; cur = ""; } else cur += ch; }
    if (cur !== "" || row.length) { row.push(cur); out.push(row); }
    return out;
  }
  function freeChair() { const F = floor(); return (F.chairs.find(c => !c.v) || {}).id || ""; }

  /* ---------- file inputs and fields ---------- */
  let xlsxP = null;
  const loadXLSX = () => window.XLSX ? Promise.resolve(true) : xlsxP || (xlsxP = new Promise(res => { const s = document.createElement("script"); s.src = "assets/xlsx.full.min.js"; s.onload = () => res(!!window.XLSX); s.onerror = () => { xlsxP = null; s.remove(); res(false); }; document.head.appendChild(s); }));
  async function readAppt(f) {
    if (!(await loadXLSX())) return toast("The spreadsheet reader didn’t load. Check the connection and reopen the add-in.", true);
    let rows;
    /* older exports: .xls (Excel 97-2003, or a web page saved as .xls), Excel 2003 XML, text; the report on whichever sheet has it */
    try { const wb = XLSX.read(await f.arrayBuffer(), { type: "array", cellDates: false });
      for (const n of wb.SheetNames) { rows = C.parseReport(XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: "" })); if (rows && rows.length) break; } }
    catch (e) { console.error(e); return toast("That file couldn’t be read. Save it from Excel as .xlsx and upload that, or upload the NextGen export unchanged.", true); }
    if (!rows || !rows.length) return toast("This doesn’t look like a NextGen appointment listing (no MRN, Appointment Date and Status columns).", true);
    /* a wider report (the providers' whole schedule) is split: IVK visits for the Appt List tab, the others kept with no patient on them */
    const ivk = rows.filter(r => C.isIvkVisit(r[5])), other = rows.filter(r => !C.isIvkVisit(r[5])).map(r => [C.isoOf(r[2]), r[3] === "" ? "" : C.fmtT(Math.round(r[3] * 1440)), r[4], r[5], r[6]]);
    if (!ivk.length) return toast("This report has no IV ketamine visits. Check the NextGen report's filters.", true);
    /* a report with the neighbouring clinics in it: each location is this clinic's (the Locations setting, else the ones naming the
       clinic) or another clinic's. A report with one location is all this clinic's unless Locations says otherwise. */
    const names = [...new Set(ivk.map(r => r[7]))].sort(), one = names.length === 1 && !C.txt(S.cfg.Locations);
    const locs = names.map(l => ({ loc: l, n: ivk.filter(r => r[7] === l).length, ours: one || C.isOurs(l, S.cfg) }));
    openModal("upload", upSplit({ file: f.name, all: ivk, other, locs, ok: false }));
  }
  /* the upload's rows for the Appt List tab (this clinic's locations) and for tblIvkElsewhere (the rest, LASFIR only) */
  function upSplit(u) {
    const ours = new Set(u.locs.filter(l => l.ours).map(l => l.loc)), tm = r => r[3] === "" ? "" : C.fmtT(Math.round(r[3] * 1440));
    u.rows = u.all.filter(r => ours.has(r[7]));
    u.away = u.all.filter(r => !ours.has(r[7])).map(r => { const mrn = C.mrnKey(r[0]), p = S.PM.get(mrn);
      return [mrn, p && p.code !== "(no code)" ? p.code : C.lasfir(r[1]), C.isoOf(r[2]), tm(r), r[4], r[5], r[6], r[7]]; });
    u.sh = C.shrink(u.rows, S.A.map(a => a.date)); u.sh.had = S.A.length;
    return u;
  }
  async function readMap(f) {
    let x; try { x = JSON.parse(await f.text()); } catch (e) { return toast("That file isn’t a map file from the Bay Map Builder.", true); }
    const c = bpCheck(x); if (c.err) return toast(c.err, true);
    await write("Saving the chair map…", () => Store.setDoc(MAP_KEY, c.plan), `Chair map loaded: ${plural(c.plan.bays.length, "chair")}.`);
  }
  document.addEventListener("change", e => {
    const t = e.target;
    if (t.id === "apptFile" && t.files[0]) readAppt(t.files[0]);
    else if (t.id === "mapFile" && t.files[0]) readMap(t.files[0]);
    else if (t.id === "sohFile" && t.files[0]) readSohar(t.files[0]);
    else if (t.id === "ipFor") { S.ipFor = t.value; render(); }
    else if (t.id === "hdWho" && S.modal === "hold") { S.m.who = t.value; render(); }
    else if (S.modal === "profile" && t.id === "pfStatus") keepProfile();
    else if (t.dataset && t.dataset.cm != null) cmSet(t.dataset.m, t.dataset.cm, t.checked);
    else if (t.dataset && t.dataset.prof != null) { t.checked = false; ACT.profile({ dataset: { m: t.dataset.prof } }); }
    else if (t.id === "darkOn") { S.theme = t.checked ? "dark" : "light"; LS.set("ivkTheme", S.theme); document.documentElement.dataset.theme = S.theme; }
    else if (t.id === "mFreq" && S.modal === "ask") { S.m.freq = t.value; S.m.text = C.askText(S.PM.get(S.m.mrn), S.m.ask, S.tpl, t.value); $("#mText").value = S.m.text; }
    else if (t.id === "mFreq" && (S.modal === "checkout" || S.modal === "proto")) { S.m.freq = t.value; S.m.proto = $("#mProto")?.checked ?? S.m.proto; render(); }
    else if (t.id === "mProto") S.m.proto = t.checked;
    else if (t.id === "mIdx") S.m.idxOn = t.checked;
    else if (t.id === "mShrink") { S.m.ok = t.checked; render(); }
    else if (S.modal === "upload" && /^upLoc_\d+$/.test(t.id)) { S.m.locs[+t.id.slice(6)].ours = t.checked; upSplit(S.m); render(); }
    else if (t.id === "apName") { keepAddProv(); S.m.other = t.value === "__other"; render(); }
    else if (t.dataset && t.dataset.ti != null) {
      /* the slot types draft: colour, name, IVK use, kind, and the eligible events checkboxes */
      const te = S.te || (S.te = S.types.map(x => Object.assign({}, x))), x = te[+t.dataset.ti]; if (!x) return;
      if (t.dataset.ev != null) {
        /* the other events keep the spelling they were written with */
        const ev = t.dataset.ev, keep = String(x.accepts || "").split(/[,;\n]/).map(a => a.trim()).filter(a => a && a.toLowerCase() !== ev.toLowerCase());
        x.accepts = (t.checked ? keep.concat(ev) : keep).join(", ");
      } else x[t.dataset.k] = t.value;
      if (t.dataset.k === "fill") x.font = inkFor(t.value);
      const open = t.closest("details") ? +t.dataset.ti : null; render();
      if (open != null) { const dd = document.querySelectorAll(".ttypes details.ms")[open]; if (dd) dd.open = true; }
    }
    else if (t.id === "reqPost") ACT.flagPost(t);
    else if (t.dataset && t.dataset.pt) ptFlag(t);
    else if (t.dataset && t.dataset.wappeal) wAppeal(t);
    else if (t.id === "wNote") { const w = S.WL.list.find(x => x.mrn === t.dataset.m); if (w) write("Saving…", () => Store.savePt(w.mrn, ptPatch({ code: w.code }, { WaitNote: t.value.trim() })), "Note saved."); }
    else if (S.modal === "forms" && ["fmFrom", "fmTo", "fp_sched", "fp_consent", "fp_avail", "fp_bidf"].includes(t.id)) { keepForms(); render(); }
    else if (S.modal === "pto" && (t.id === "ptCov" || t.id === "ptProv" || t.id === "ptFrom" || t.id === "ptTo")) { const was = S.m.cov; keepPto(); if (t.id === "ptCov") { S.m.cov = t.value; if (t.value !== was) S.m.useTpl = ""; } render(); }
    else if (t.id === "followOn") { LS.set("ivkFollow", t.checked ? "1" : "0"); followed = ""; }
    else if (["ipStart", "ipPer", "ipWeeks"].includes(t.id)) {
      const o = ipList().find(x => x.mrn === S.ipSel); if (!o) return; delete S.ipPins[o.mrn];
      if (t.id === "ipStart") { const d = C.toSerial(t.value); if (d != null) savePlanOf(o, { Start: C.isoOf(d) }); }
      else savePlanOf(o, t.id === "ipPer" ? { PerWeek: t.value } : { Weeks: t.value });
    }
    else if (t.id === "mRn") S.m.rn = t.checked;
    else if (t.id === "mAsk") S.m.ask = t.checked;
    else if (/^mAmt_/.test(t.id) || t.id === "mAmtSave" || t.id === "mXfer") { keepPay(); render(); }
    else if (S.modal === "xfer" && ["xfTo", "xfDate"].includes(t.id)) { keepXfer(); render(); }
  });
  document.addEventListener("input", e => { if (e.target.id === "ptSearch") { S.q = e.target.value; const pos = e.target.selectionStart; render(); const s = $("#ptSearch"); if (s) { s.focus(); s.setSelectionRange(pos, pos); } } });
  /* a start and its post check shade together while either is hovered or focused (Schedule and the Access week) */
  const pairOf = el => el && el.closest ? el.closest("[data-pc],[data-pck]") : null;
  const shade = (el, on) => { const k = el.dataset.pc || el.dataset.pck; if (!k) return;
    document.querySelectorAll(`[data-pc="${CSS.escape(k)}"],[data-pck="${CSS.escape(k)}"]`).forEach(x => x.classList.toggle("pair", on)); };
  for (const [ev, on] of [["mouseover", true], ["focusin", true], ["mouseout", false], ["focusout", false]]) document.addEventListener(ev, e => { const el = pairOf(e.target); if (el) shade(el, on); });
  /* the folded chair tiles stay as they were left across re-renders */
  /* the template grid: press and drag across half hours (mouse, pen or touch) to select a rectangle; Shift-click or Enter extends it */
  let drag = null;
  const paintSel = () => document.querySelectorAll(".tc").forEach(b => b.classList.toggle("sel", inRect(S.tsel, +b.dataset.d, +b.dataset.t)));
  const cellAt = (x, y) => document.elementFromPoint(x, y)?.closest?.(".tc") || null;
  document.addEventListener("pointerdown", e => {
    const c = e.target.closest?.(".tc"), g = curGrid(); if (!c || !g || e.button > 0) return;
    e.preventDefault(); try { c.releasePointerCapture(e.pointerId); } catch (x) { /* not captured */ }
    const d = +c.dataset.d, t = +c.dataset.t;
    if (e.shiftKey && S.tsel && S.tsel.g === g.top) { S.tsel = Object.assign({}, S.tsel, { d1: d, t1: t }); render(); return; }
    drag = { g: g.top, d0: d, t0: t, d1: d, t1: t }; S.tsel = drag; paintSel();
  });
  document.addEventListener("pointermove", e => { if (!drag) return; const c = cellAt(e.clientX, e.clientY); if (!c) return; const d = +c.dataset.d, t = +c.dataset.t; if (d !== drag.d1 || t !== drag.t1) { drag.d1 = d; drag.t1 = t; paintSel(); } });
  document.addEventListener("pointerup", () => { if (!drag) return; drag = null; render(); });
  document.addEventListener("pointercancel", () => { if (!drag) return; drag = null; render(); });
  document.addEventListener("click", e => {
    /* a keyboard press on a half hour selects it, or extends the selection with Shift */
    const tc = e.target.closest?.(".tc");
    if (tc && e.detail === 0) { const g = curGrid(), d = +tc.dataset.d, t = +tc.dataset.t; S.tsel = e.shiftKey && S.tsel ? Object.assign({}, S.tsel, { d1: d, t1: t }) : { g: g.top, d0: d, t0: t, d1: d, t1: t }; render(); $(`.tc[data-d="${d}"][data-t="${t}"]`)?.focus(); return; }
    const el = e.target.closest("[data-act]"); if (!el || el.disabled) return;
    if (el.tagName === "LABEL" || el.tagName === "INPUT" || (el.dataset.act === "scrim" && e.target !== el)) return;
    const f = ACT[el.dataset.act]; if (!f) return;
    e.preventDefault(); Promise.resolve(f(el, e)).catch(err => { console.error(err); toast(String(err.message || err), true); });
  });
  document.addEventListener("keydown", e => {
    if (e.key === "Escape" && S.modal) return closeModal();
    const el = e.target.closest?.('[role="button"][data-act]');
    if (el && (e.key === "Enter" || e.key === " ")) { e.preventDefault(); el.click(); }
  });

  /* ---------- start ---------- */
  const typing = () => { const a = document.activeElement; return a && /^(INPUT|TEXTAREA|SELECT)$/.test(a.tagName) && a.type !== "checkbox"; };
  function landing() {
    $("#app").innerHTML = `<div class="landing"><h1>IVK Tracker</h1><p class="muted">An Excel add-in for the IV ketamine front desk: who’s here, which chair, and what to ask before they leave. It opens inside the IVK Patient Tracker workbook in Excel on the web.</p>
      <div class="card"><h2>Install it (once)</h2><ol class="how"><li>Download the <a href="manifest.xml" download>add-in manifest</a>.</li><li>In Excel on the web, open the IVK Patient Tracker workbook.</li><li>Home › Add-ins › More add-ins › My add-ins › Upload my add-in, then choose the file.</li><li>Open it from the IVK button on the Home tab.</li></ol></div></div>`;
  }
  function start(info) {
    if (!info || !info.host) return landing();
    if (FULL) {
      document.body.classList.add("fs");
      const hello = () => { if (S.raw) return; try { toPane({ k: "ready" }); } catch (e) { /* pane not listening yet */ } setTimeout(hello, 2000); };
      Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, a => reassemble(a.message, onPane), hello);
    } else {
      refresh(true);
      setInterval(() => { if (pop.dlg || (!S.modal && !typing())) refresh(); }, POLL_MS);
    }
    setInterval(() => { if (S.raw && !S.modal && !typing()) { derive(); render(); } }, TICK_MS);
    render();
    /* Dev Update notices: the bell top right; full screen passes what it dismisses back to the pane */
    if (window.DevUpdates) DevUpdates.init({ app: "ivk", send: FULL ? m => toPane(m) : null });
  }
  window.addEventListener("error", e => console.error(e.message));
  if (window.Office && Office.onReady) Office.onReady(start); else landing();
  window.IvkApp = { S, render, derive, refresh, ACT, floor, ipList, ipModel };   // for the test harness
})();
