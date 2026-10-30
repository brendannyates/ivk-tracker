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

  const S = { view: "today", raw: null, sig: "", live: false, err: "", sel: null, bSel: null, pSel: null, chip: "need", q: "", wk: 0, ipView: "plan", ipSel: null, ipPins: {}, ipAlt: null, ipWk: 0, tg: 0, tsel: null, te: null,
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
    S.log = C.byHeader(r.tblIvkLog).rows.map(x => ({ ev: C.txt(x.get("event")), code: C.txt(x.get("patient")), chair: C.txt(x.get("chair")), at: C.txt(x.get("sentat")), by: C.txt(x.get("sentby")), result: C.txt(x.get("result")) })).filter(l => l.ev);
    S.P = C.patients(r.mpt); S.PM = new Map(S.P.map(p => [p.mrn, p]));
    S.A = C.appts(r.appt);
    /* a tracker row whose Patient Name isn't a code takes the code from the Appt List's Name column, never the name itself */
    const codeOf = new Map(S.A.filter(a => a.code).map(a => [a.mrn, a.code]));
    for (const p of S.P) if (!p.code) p.code = codeOf.get(p.mrn) || "(no code)"; S.V = C.visitRows(r.visits); S.CS = C.consults(r.consults, T);
    let plan = null; try { plan = r.plan ? bpCheck(r.plan).plan || null : null; } catch (e) { plan = null; }
    S.plan = plan; S.planErr = r.plan && !plan ? "The saved chair map can’t be read. Load the map file again." : "";
    S.x = { today: T, cfg, appts: S.A, patients: S.P, visits: S.V, clinic: cfg.Clinic };
    S.list = C.today(S.x); S.chairs = C.chairList(cfg, plan);
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
    S.WL = C.waitlist(S.CS, S.PT);
    /* a visit takes its post check when it's an index visit (a treatment start, an Induction patient) or the patient requires one */
    const needsPost = (mrn, ev) => { const p = S.PM.get(mrn); return /start/i.test(ev) || !!p?.reqPost || /induction/i.test(p?.status || ""); };
    /* provider time off: their template slots go on those days, or become the covering provider's */
    S.EX = C.exRows(r.tblIvkExceptions);
    S.X = Object.assign({ appts: S.A, holds: S.holds, others: S.others, cfg, chairs: S.chairs.length, listAt: r.apptUpdated, needsPost }, C.withExceptions(TM.onDays(S.grids, S.types, +cfg.PostCheckMin), S.EX));
    const pv = new Map(), addPv = n => { const k = TM.provKey(n); if (n && k && !pv.has(k)) pv.set(k, n); };
    S.A.forEach(a => addPv(a.provider)); S.others.forEach(o => addPv(o.provider));
    S.providers = [...pv.values()].sort((a, b) => TM.provKey(a).localeCompare(TM.provKey(b)));
    S.events = [...new Set(S.A.map(a => a.event).concat(S.others.map(o => o.event)).filter(Boolean))].sort();
    S.freqs = (r.freqs && r.freqs.length ? r.freqs : ["2x per week", "1x per week", "every 2 weeks", "every 3 weeks", "every 4 weeks", "every 6 weeks", "every 2 months", "every 3 months", "PRN"]);
  }
  /* last Teams post of an event for a patient today (from tblIvkLog) */
  const sentToday = (ev, code) => { const iso = C.isoOf(todayS()); return S.log.filter(l => l.ev === ev && l.code === code && new Date(l.at).toString() !== "Invalid Date" && C.isoOf(C.serialOf(new Date(l.at))) === iso && !/^failed/.test(l.result)).pop() || null; };
  const lastSent = (ev, code) => S.log.filter(l => l.ev === ev && l.code === code && !/^failed/.test(l.result)).pop() || null;
  const timeOf = iso => { const d = new Date(iso); return isNaN(d) ? "" : C.fmtT(d.getHours() * 60 + d.getMinutes()); };
  const dateOfIso = iso => { const d = new Date(iso); return isNaN(d) ? "" : C.fmtD(C.serialOf(d)); };
  const teamsReady = ev => !!(S.routes[ev] && S.routes[ev].on && N.flowUrl(S.routes[ev].url) && N.truthy(S.cfg.Enabled));
  const routeTo = ev => S.routes[ev]?.to || (ev === "arrival" ? "the RN" : "the group chat");

  function floor() {
    const now = nowMin(), L = S.list, cfg = S.cfg;
    const chairs = S.chairs.map(id => Object.assign({ id }, C.chairState(id, L, now, cfg)));
    const by = ph => L.filter(v => v.ph === ph);
    const waiting = by("waiting").sort((a, b) => a.st.arr - b.st.arr), late = waiting.filter(v => C.isLate(v, now, cfg));
    const seated = by("chair"), off = seated.filter(v => !S.chairs.includes(v.st.chair));
    const asks = L.filter(v => !["gone", "noshow", "done"].includes(v.ph) && v.ask.any && v.st.asked == null);
    const openBook = by("done").filter(v => v.st.booked === "later" && v.ask.book);
    return { now, chairs, waiting, late, seated, off, coming: by("booked"), done: by("done"), noshow: by("noshow"), gone: by("gone"), asks, openBook,
      ready: chairs.filter(c => c.state === "ready"), soon: chairs.filter(c => c.state === "soon") };
  }
  /* the visit to show when nothing is picked: waiting too long, ready to check out, out soon, waiting, in a chair, next to arrive */
  function pickVisit(F) {
    return F.late[0] || F.ready[0]?.v || F.soon[0]?.v || F.waiting[0] || F.seated[0] || F.coming.find(v => (v.a.time ?? 0) >= F.now - 30) || F.coming[0] || F.done[F.done.length - 1] || null;
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
    saveTypes: (t, g) => X.saveTypes(t, g), replaceOther: r => X.replaceOther(r),
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
    if (m.k === "snap") { S.raw = m.raw; S.live = true; derive(); if (!S.modal && !typing()) render(); }
    else if (m.k === "wait") rpc.wait.get(m.id)?.alive();
    else if (m.k === "ack") { const w = rpc.wait.get(m.id); if (!w) return; rpc.wait.delete(m.id); if (m.ok) w.res(m.result); else w.rej(new Error(m.err || "the task pane couldn’t do that")); }
  }
  function openFull() {
    if (FULL) { try { toPane({ k: "close" }); } catch (e) { /* pane gone */ } return; }
    if (pop.dlg) return toast("Full screen is already open.");
    if (!Office.context.requirements.isSetSupported("DialogApi", "1.2")) return toast("This version of Excel can’t open full screen.", true);
    const url = new URL(location.href); url.search = "?full=1"; url.hash = "";
    Office.context.ui.displayDialogAsync(url.href, { height: 95, width: 95, displayInIframe: true }, res => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) return toast(res.error.code === 12007 ? "Close the other add-in window first." : "Full screen couldn’t open. Try again.", true);
      pop.dlg = res.value; render();
      pop.dlg.addEventHandler(Office.EventType.DialogMessageReceived, onChild);
      pop.dlg.addEventHandler(Office.EventType.DialogEventReceived, () => { pop.dlg = null; render(); });
    });
  }
  function closeFull() { if (pop.dlg) { try { pop.dlg.close(); } catch (e) { /* closed */ } pop.dlg = null; } render(); }

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
      const raw = await X.load(); raw.plan = X.getDoc(MAP_KEY) || null;
      const sig = hash(JSON.stringify(raw));
      S.live = true; S.err = "";
      if (force || sig !== S.sig) { S.sig = sig; S.raw = raw; derive(); toChild({ k: "snap", raw }); if (!S.modal && !typing()) render(); }
    } catch (e) { console.error(e); S.err = "Couldn’t read the workbook. Check you’re still signed in; the tracker tries again in a few seconds."; render(); }
  }

  /* ---------- frame ---------- */
  const NAV = [["today", "Today"], ["schedule", "Schedule"], ["index", "Index planner"], ["patients", "Patients"], ["consults", "Waitlist"]];
  function badges() {
    const F = floor(), b = {};
    const t = F.late.length + F.ready.length; if (t) b.today = { n: t, bad: F.late.length > 0 };
    /* Patients counts only what blocks the next infusion; the amber needs live under the Needs attention chip */
    const pn = S.P.filter(p => C.needs(p, S.x).some(y => y.bad)).length; if (pn) b.patients = { n: pn, bad: true };
    const ca = S.CS.filter(c => c.state === "act").length; if (ca) b.consults = { n: ca };
    /* holds waiting to be booked; red once one is past HoldMin without being confirmed */
    /* and booked holds the uploaded report hasn't confirmed; red when a newer report doesn't have them */
    const hs = liveHolds(), U = unconfirmed(), np = new Set(hs.map(h => h.mrn).concat(U.map(h => h.mrn))).size;
    if (np) b.index = { n: np, bad: hs.some(h => h.state === "late") || U.some(h => h.conf === "missing") };
    const set = (S.raw?.missing || []).length + (teamsReady("arrival") && teamsReady("protocol") ? 0 : 1); if (set) b.settings = { n: set };
    return b;
  }
  function render() {
    const page = $(".page"), top = page ? page.scrollTop : 0;
    if (!S.raw) { $("#app").innerHTML = `<div class="landing"><h1>IVK Tracker</h1><p class="muted">${S.err ? esc(S.err) : "Reading the workbook…"}</p></div>`; return; }
    const B = badges(), item = ([k, label]) => `<button class="item" data-act="nav" data-v="${k}"${S.view === k ? ' aria-current="page"' : ""}>${label}${B[k] ? `<span class="count${B[k].bad ? " bad" : ""}" aria-label="${B[k].n} need attention">${B[k].n}</span>` : ""}</button>`;
    const body = { today: todayHTML, schedule: scheduleHTML, index: indexHTML, patients: patientsHTML, consults: consultsHTML, settings: settingsHTML }[S.view]();
    $("#app").innerHTML = `<div class="app">
      <nav class="rail" aria-label="IVK Tracker">
        <div class="brand">IVK Tracker<small>${esc(S.cfg.Clinic)}</small></div>
        ${NAV.map(item).join("")}
        <div class="gap"></div>${item(["settings", "Settings"])}
        <div class="spacer"></div>
        <button class="item" data-act="full">${FULL ? "Exit full screen" : pop.dlg ? "Full screen is open" : "Full screen"}</button>
        <div class="conn${S.err ? " off" : ""}"><span class="cdot"></span>${S.err ? "Not connected" : S.busy ? esc(S.busy) : FULL ? "Saving through the task pane" : "Connected to the workbook"}</div>
      </nav>
      <main class="page" id="main">${S.err ? `<div class="msg bad">${esc(S.err)}</div>` : ""}${body}</main></div>
      ${S.modal ? `<div class="scrim" data-act="scrim"><div class="modal" role="dialog" aria-modal="true" aria-labelledby="mh">${MODALS[S.modal]()}</div></div>` : ""}`;
    const p2 = $(".page"); if (p2) p2.scrollTop = top;
    if (S.modal) { const f = $(".modal [autofocus]") || $(".modal button.primary"); if (f && !$(".modal").contains(document.activeElement)) f.focus(); }
  }
  let toastT = null;
  function toast(msg, bad) {
    const t = $("#toast"); if (!t) return; t.textContent = msg; t.classList.toggle("bad", !!bad); t.classList.add("show");
    clearTimeout(toastT); toastT = setTimeout(() => t.classList.remove("show"), bad ? 6000 : 3200);
  }
  /* top right of every screen: read the workbook again (circular arrow), and light or dark (the moon switches to dark, the sun back) */
  const ICON = {
    refresh: `<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 12a8 8 0 1 1-2.34-5.66"/><path d="M20 4v5h-5"/></svg>`,
    moon: `<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 14.5A8 8 0 0 1 9.5 4a8 8 0 1 0 10.5 10.5z"/></svg>`,
    sun: `<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/></svg>`
  };
  const tools = () => { const dark = S.theme === "dark", busy = !!S.reading;
    return `<div class="tools"><button class="ibtn${busy ? " spin" : ""}" data-act="reload" title="Refresh: read the workbook again" aria-label="Refresh from the workbook"${busy ? " disabled" : ""}>${ICON.refresh}</button>
      <button class="ibtn" data-act="theme" title="${dark ? "Light mode" : "Dark mode"}" aria-label="${dark ? "Switch to light mode" : "Switch to dark mode"}" aria-pressed="${dark}">${dark ? ICON.sun : ICON.moon}</button></div>`; };
  const head = (title, sub, clock) => `<div class="top"><div><h1>${title}</h1>${sub ? `<p class="sub">${sub}</p>` : ""}</div><div class="topr">${tools()}${clock ? `<div class="clock"><div class="t">${hm()}</div><div class="d">${esc(C.fmtDay(todayS()))}</div></div>` : ""}</div></div>`;
  const TABLE_WHAT = { tblIvkPlans: "index plans", tblIvkHolds: "holds", tblIvkOther: "the providers’ other visits", tblIvkFlags: "post-check flags", tblIvkPatient: "payment flags, Oakland, availability and the waitlist order", tblIvkPay: "payments and balance reminders", tblIvkExceptions: "provider time off and coverage" };
  function setupBanner() {
    const miss = S.raw.missing || [], noTabs = ["mpt", "appt"].filter(k => !S.raw.tables?.[k]);
    if (noTabs.length) return `<div class="msg warn"><b>This workbook has no ${noTabs.map(k => k === "mpt" ? "Master Patient Tracker" : "Appt List").join(" or ")} table.</b> Open the IVK Patient Tracker workbook, then the add-in.</div>`;
    const NEW = ["tblIvkPlans", "tblIvkHolds", "tblIvkOther", "tblIvkFlags", "tblIvkPatient", "tblIvkPay", "tblIvkExceptions"];
    if (miss.length && miss.every(n => NEW.includes(n))) return `<div class="msg warn"><b>This version needs ${miss.length === 1 ? "one more table" : `${miss.length} more tables`}.</b> Updating adds them to the hidden ${esc(X.DATA_SHEET)} tab (${miss.map(n => TABLE_WHAT[n]).join(", ")}); nothing else changes. <div class="actions"><button class="btn primary inl" data-act="setup">Update the workbook</button></div></div>`;
    if (miss.length) return `<div class="msg warn"><b>This workbook isn’t set up for the add-in yet.</b> Setting up adds a hidden IVK Tracker Data tab for check-ins and Teams settings; nothing else changes. <div class="actions"><button class="btn primary inl" data-act="setup">Set up this workbook</button></div></div>`;
    return "";
  }

  /* ---------- Today ---------- */
  function todayHTML() {
    const F = floor(), sub = [];
    if (F.seated.length) sub.push(`${F.seated.length} in ${F.seated.length === 1 ? "a chair" : "chairs"}`);
    if (F.late.length) sub.push(`<strong>${F.late.length} waiting ${C.waitMin(F.late[0], F.now)}+ min for a chair</strong>`);
    else if (F.waiting.length) sub.push(`${F.waiting.length} checked in, waiting`);
    if (F.ready.length) sub.push(`${F.ready.length} ready to check out`);
    if (F.coming.length) sub.push(`${F.coming.length} still to come`);
    if (F.asks.length) sub.push(`${F.asks.length} to ask about protocol or booking`);
    if (!sub.length) sub.push(S.list.length ? "Everyone has been seen today" : "No IVK visits on the list today");
    if (!S.sel || !visitBy(S.sel)) { const v = pickVisit(F); S.sel = v ? v.key : null; }
    return `${setupBanner()}${apptStale()}${head("Today", sub.join(", "), true)}
      <div class="cols"><div class="stack">${chairsHTML(F)}${waitingHTML(F)}${comingHTML(F)}${doneHTML(F)}</div>
      <div>${S.bSel && !F.chairs.find(c => c.id === S.bSel)?.v ? openChairHTML(F, S.bSel) : visitCardHTML(visitBy(S.sel), F)}</div></div>`;
  }
  function apptStale() {
    const u = S.raw.apptUpdated, T = todayS();
    if (u != null && Math.floor(u) >= T - 1) return "";
    return `<div class="msg warn"><b>${u == null ? "No record of when the appointment list was last updated." : `The appointment list was last updated ${C.fmtD(Math.floor(u))}.`}</b> Upload a fresh NextGen export so today’s list is right. <button class="btn ghost sm" data-act="nav" data-v="settings">Upload in Settings</button></div>`;
  }
  const chairTxt = id => `Chair ${id}`;
  function chairsHTML(F) {
    const P = S.plan, used = F.chairs.filter(c => c.v).length;
    const hd = `<div class="hd"><h2>Chairs</h2><span>${used} of ${F.chairs.length} in use${F.ready.length ? `, ${F.ready.length} ready to check out` : ""}</span></div>`;
    const key = `<ul class="key" aria-label="Key"><li><span class="sw open"></span>Open</li><li><span class="sw track"></span>In chair</li><li><span class="sw soon"></span>Out within ${S.cfg.SoonMin} min</li><li><span class="sw ready"></span>Ready to check out</li></ul>`;
    const label = c => `${chairTxt(c.id)}, ${c.v ? `${C.code(c.v)}, ${C.chairStatus(c).replace(/min\b/, "minutes")}` : "open"}`;
    const cur = c => S.bSel ? S.bSel === c.id : !!(c.v && c.v.key === S.sel);
    const chairOf = id => F.chairs.find(c => c.id === id) || { id, state: "open", v: null };
    const off = F.off.length ? `<p class="boff">In a chair that isn’t on the map: ${F.off.map(v => `${esc(C.code(v))} (${esc(chairTxt(v.st.chair))})`).join(", ")}.</p>` : "";
    if (!P) {
      const tiles = F.chairs.map(c => `<button class="ctile" data-act="bSel" data-b="${esc(c.id)}" data-state="${c.state}" aria-pressed="${cur(c)}" aria-label="${esc(label(c))}">
        <span class="n"><span class="num">${esc(c.id)}</span>${c.v ? esc(C.code(c.v)) : "Open"}</span><span class="s"><span class="dot"></span>${c.v ? esc(C.chairStatus(c)) : F.waiting.length ? `${esc(C.code(F.waiting[0]))} waiting` : "Nobody waiting"}</span></button>`).join("");
      return `<section class="card mapcard">${hd}<div class="strip">${tiles}</div>${key}${off}${S.planErr ? `<p class="boff">${esc(S.planErr)}</p>` : ""}</section>`;
    }
    const W = P.w, H = P.h;
    const bays = P.bays.map(b => { const c = chairOf(b.id); return Object.assign({}, b, { c, st: c.state, title: c.v ? C.code(c.v) : "Open", s: c.v ? C.chairStatus(c) : "" }); });
    const lit = b => !!b.c.v, usedB = id => lit(bays.find(b => b.id === id) || { c: {} });
    const pos = b => `left:${bpPct(b.cx - b.r, W)};top:${bpPct(b.cy - b.r, H)};width:${bpPct(2 * b.r, W)}`;
    const tag = b => `<div class="btag" data-state="${b.st}" style="left:${bpPct(b.tag[0], W)};top:${bpPct(b.tag[1], H)}" aria-hidden="true"><div class="n"><span class="num">${esc(b.id)}</span>${esc(b.title)}</div>${b.s ? `<div class="s"><span class="dot"></span><span class="tx">${esc(b.s)}</span></div>` : ""}</div>`;
    const bay = b => `<button class="bay" data-act="bSel" data-b="${esc(b.id)}" data-state="${b.st}" aria-pressed="${cur(b.c)}" aria-current="${cur(b.c)}" aria-label="${esc(label(b.c))}" style="${pos(b)}"></button>`;
    return `<section class="card mapcard">${hd}<div class="bplan" style="aspect-ratio:${W}/${H}">
      <svg viewBox="0 0 ${W} ${H}" aria-hidden="true">${bpUnder(P, usedB)}</svg>
      ${bpRooms(P)}${bays.map(bay).join("")}
      <svg viewBox="0 0 ${W} ${H}" aria-hidden="true">${bpOver(P, bays, lit)}</svg>
      ${bays.map(tag).join("")}</div>${key}${off}</section>`;
  }
  /* what the row says it needs, in amber; red only for a too-long wait */
  function askFlag(v) {
    if (!v.p) return `<span class="f">Not on the Master Patient Tracker</span>`;
    if (v.st.asked != null) return `<span class="ok">Asked the group chat at ${C.fmtT(v.st.asked)}</span>`;
    const a = v.ask, t = a.book && a.proto ? "Confirm protocol, book at discharge" : a.book ? (a.thru == null ? "Nothing booked after today" : `Booked through ${C.fmtD(a.thru)}`) : a.proto ? "Confirm protocol" : "";
    return t ? `<span class="f">${t}</span>` : `<span>Booked through ${C.fmtD(a.thru)}</span>`;
  }
  const rowBtn = (v, act, label) => `<button class="btn ghost sm" data-act="${act}" data-k="${esc(v.key)}">${label}</button>`;
  function waitingHTML(F) {
    if (!F.waiting.length) return "";
    const rows = F.waiting.map(v => { const w = C.waitMin(v, F.now), late = C.isLate(v, F.now, S.cfg);
      return `<div class="row" role="button" tabindex="0" data-act="sel" data-k="${esc(v.key)}" data-state="${late ? "late" : "track"}" aria-pressed="${S.sel === v.key && !S.bSel}">
        <span class="p">${esc(C.code(v))}</span><span class="w">${C.fmtT(v.a.time)}</span><span class="i">${late ? `<span class="f bad">Waiting ${w} min for a chair</span>` : `Checked in ${C.fmtT(v.st.arr)}, ${askFlag(v)}`}</span>${rowBtn(v, "seat", "Seat")}</div>`; }).join("");
    return `<section class="card list"><div class="lh"><h2>Checked in, waiting for a chair</h2><span>${F.waiting.length}</span></div>${rows}</section>`;
  }
  function comingHTML(F) {
    const rows = F.coming.map(v => { const late = v.a.time != null && F.now - v.a.time > 15;
      return `<div class="row" role="button" tabindex="0" data-act="sel" data-k="${esc(v.key)}" aria-pressed="${S.sel === v.key && !S.bSel}">
        <span class="p">${esc(C.code(v))}</span><span class="w">${C.fmtT(v.a.time)}</span><span class="i">${late ? `<span class="f">Not here yet</span>, ` : ""}${esc(C.provName(v.a.provider))}${v.a.event && /start/i.test(v.a.event) ? ", treatment start" : ""}, ${askFlag(v)}</span>${rowBtn(v, "checkin", "Check in")}</div>`; }).join("");
    return `<section class="card list"><div class="lh"><h2>Still to come</h2><span>${F.coming.length ? F.coming.length : ""}</span></div>${rows || `<p class="empty">Nobody else is booked today.</p>`}</section>`;
  }
  function doneHTML(F) {
    const L = [...F.done, ...F.noshow, ...F.gone];
    if (!L.length) return "";
    const what = v => v.ph === "noshow" ? "No-show in NextGen" : v.ph === "gone" ? "Not coming today" : v.st.booked === "later" && v.ask.book ? `<span class="f">Out ${C.fmtT(v.st.out)}, next visit not booked</span>` : `Out ${C.fmtT(v.st.out)}${v.st.booked === "yes" || !v.ask.book ? ", next visit booked" : ""}`;
    const rows = L.map(v => `<div class="row" role="button" tabindex="0" data-act="sel" data-k="${esc(v.key)}" aria-pressed="${S.sel === v.key && !S.bSel}"><span class="p">${esc(C.code(v))}</span><span class="w">${C.fmtT(v.a.time)}</span><span class="i">${what(v)}</span><span></span></div>`).join("");
    return `<section class="card list"><details class="fold"${F.openBook.length ? " open" : ""}><summary>Done for today <span>${L.length}${F.openBook.length ? `, ${F.openBook.length} still to book` : ""}</span></summary>${rows}</details></section>`;
  }
  function openChairHTML(F, id) {
    const w = F.waiting[0];
    return `<aside class="card detail" aria-label="${esc(chairTxt(id))}"><div class="dh"><div><div class="name">${esc(chairTxt(id))}<span class="loc">open</span></div></div></div>
      <p class="body">${w ? `${esc(C.code(w))} checked in at ${C.fmtT(w.st.arr)} and is waiting.` : F.coming[0] ? `Nobody is waiting. Next to arrive: ${esc(C.code(F.coming[0]))} at ${C.fmtT(F.coming[0].a.time)}.` : "Nobody is waiting."}</p>
      ${w ? `<button class="btn primary" data-act="seatHere" data-k="${esc(w.key)}" data-b="${esc(id)}">Seat ${esc(C.code(w))} in ${esc(chairTxt(id))}</button>` : ""}
      ${F.waiting.length > 1 ? `<div class="actions">${F.waiting.slice(1).map(v => `<button class="btn ghost" data-act="seatHere" data-k="${esc(v.key)}" data-b="${esc(id)}">Seat ${esc(C.code(v))} here</button>`).join("")}</div>` : ""}</aside>`;
  }

  /* the visit card: lead, timeline, steps, the one next step */
  function visitState(v, F) {
    const cs = v.ph === "chair" ? F.chairs.find(c => c.id === v.st.chair) || C.chairState(v.st.chair, S.list, F.now, S.cfg) : null;
    const askDue = v.ask.any && v.st.asked == null && !!v.p;
    if (v.ph === "booked") { const late = v.a.time != null && F.now - v.a.time > 15;
      return { cls: late ? "soon" : "", lead: late ? `${C.fmtT(v.a.time)} visit, not checked in yet` : `Booked at ${C.fmtT(v.a.time)}`, pri: ["checkin", `Check in ${C.code(v)}`] }; }
    if (v.ph === "waiting") { const w = C.waitMin(v, F.now), late = C.isLate(v, F.now, S.cfg);
      return { cls: late ? "late" : "", lead: late ? `Waiting ${w} min for a chair` : `Checked in at ${C.fmtT(v.st.arr)}, waiting for a chair`, pri: ["seat", `Seat ${C.code(v)} in a chair`] }; }
    if (v.ph === "chair") {
      const lead = cs.state === "ready" ? `Due out at ${C.fmtT(cs.end)}. Check out at the desk` : cs.state === "soon" ? `Out in ${cs.left} min, around ${C.fmtT(cs.end)}` : `In ${chairTxt(v.st.chair)} since ${C.fmtT(v.st.seat)}, out around ${C.fmtT(cs.end)}`;
      return { cls: cs.state === "track" ? "" : cs.state, lead, cs, pri: askDue && cs.state === "track" ? ["ask", "Ask the group chat"] : ["checkout", `Check out ${C.code(v)}`] };
    }
    if (v.ph === "done") { const open = v.st.booked === "later" && v.ask.book;
      return { cls: open ? "soon" : "ready", lead: open ? `Checked out at ${C.fmtT(v.st.out)}. Next visit not booked yet` : `Checked out at ${C.fmtT(v.st.out)}`, pri: open ? ["markBooked", "Mark next visit booked"] : null }; }
    if (v.ph === "noshow") return { cls: "", lead: "NextGen has this visit as a no-show", pri: ["checkin", `Check in ${C.code(v)} anyway`] };
    return { cls: "", lead: "Taken off today’s list", pri: ["undoGone", "Put back on today’s list"] };
  }
  function visitCardHTML(v, F) {
    if (!v) return `<aside class="card detail"><div class="name">No visit picked</div><p class="body">${S.list.length ? "Pick a chair or a row to see the visit." : "Nobody is on today’s IVK list. Upload a fresh appointment list in Settings if that looks wrong."}</p></aside>`;
    const vs = visitState(v, F), k = esc(v.key), c = C.code(v), p = v.p;
    const meta = [C.fmtT(v.a.time), v.a.event ? C.eventShort(v.a.event).toLowerCase() : "", C.provName(v.a.provider), p?.freq, p?.reqPost ? "requires post-check" : ""].filter(Boolean).join(", ");
    const loc = v.ph === "chair" ? chairTxt(v.st.chair) : v.ph === "waiting" ? "lobby" : "";
    const arr = sentToday("arrival", c), ask = sentToday("protocol", c);
    /* timeline while in a chair: in chair, provider check (SOP: within the first 30 min), expected out */
    let tl = "";
    if (vs.cs) {
      const s0 = v.st.seat, e0 = vs.cs.end, span = Math.max(1, e0 - s0), at = t => Math.min(100, Math.max(0, (t - s0) / span * 100));
      tl = `<div class="tl" aria-hidden="true"><div class="bar"></div><div class="fill" style="width:${at(F.now)}%"></div>
        <div class="mark done" style="left:0"></div><div class="mark${F.now >= vs.cs.check ? " done" : ""}" style="left:${at(vs.cs.check)}%"></div><div class="mark${F.now >= e0 ? " done" : ""}" style="left:100%"></div>
        <div class="now" style="left:${at(F.now)}%"></div>
        <span class="lbl" style="left:${at(vs.cs.check)}%">Provider check ${C.fmtT(vs.cs.check)}</span><span class="lbl end" style="left:100%">Out ${C.fmtT(e0)}</span></div>`;
    }
    const st = (cls, t, m, w) => `<li class="step ${cls}"><span class="k">${cls === "done" ? "✓" : ""}</span><span><span class="t">${t}</span>${m ? `<span class="m">${m}</span>` : ""}</span><span class="w">${w || ""}</span></li>`;
    const steps = [
      st(v.st.arr != null ? "done" : v.ph === "booked" || v.ph === "noshow" ? "now" : "next", "Checked in", arr ? `${esc(routeTo("arrival"))} told in Teams at ${timeOf(arr.at)}` : v.st.arr != null ? (teamsReady("arrival") ? `${esc(routeTo("arrival"))} not told in Teams` : "Teams not set up for arrivals") : "", v.st.arr != null ? C.fmtT(v.st.arr) : ""),
      st(v.st.seat != null ? "done" : v.ph === "waiting" ? (C.isLate(v, F.now, S.cfg) ? "late" : "now") : "next", v.st.chair ? `In ${chairTxt(v.st.chair)}` : "In a chair", v.ph === "waiting" && C.isLate(v, F.now, S.cfg) ? `Waiting ${C.waitMin(v, F.now)} min` : "", v.st.seat != null ? C.fmtT(v.st.seat) : ""),
      v.ask.any || v.st.asked != null ? st(v.st.asked != null ? "done" : v.ph === "chair" || v.ph === "waiting" ? "now" : "next", v.ask.book && v.ask.proto ? "Ask to confirm protocol and book at discharge" : v.ask.book ? "Ask to send to the desk for booking" : "Ask to confirm protocol",
        v.st.asked != null ? (ask ? `Sent to ${esc(routeTo("protocol"))}` : "Marked asked") : teamsReady("protocol") ? `Goes to ${esc(routeTo("protocol"))}` : "Teams not set up for the group chat", v.st.asked != null ? C.fmtT(v.st.asked) : "") : "",
      st(v.st.out != null ? (vs.cls === "soon" ? "soon" : "done") : vs.cs && vs.cs.state !== "track" ? (vs.cs.state === "soon" ? "soon" : "now") : "next", "Checked out",
        v.st.out != null ? (v.st.booked === "yes" ? "Next visit booked" : v.st.booked === "no" ? "No next visit needed" : v.ask.book ? "Next visit not booked yet" : "Next visit on the list") : vs.cs ? `Expected around ${C.fmtT(vs.cs.end)}` : "", v.st.out != null ? C.fmtT(v.st.out) : "")
    ].join("");
    /* secondary steps as ghost buttons; corrections and lookups as quiet ones underneath */
    const ghosts = [], quiet = [];
    if (v.ph === "chair" && vs.pri[0] === "ask") ghosts.push(["checkout", "Check out"]);
    if ((v.ph === "chair" || v.ph === "waiting") && vs.pri[0] !== "ask" && v.ask.any && v.st.asked == null && v.p) ghosts.push(["ask", "Ask the group chat"]);
    if (v.ph === "waiting") ghosts.push(["remind", "Remind the RN"]);
    if (v.ph === "chair") ghosts.push(["seat", "Move to another chair"]);
    if (v.ph === "booked") quiet.push(["gone", "Not coming today"]);
    if (v.ph === "waiting") quiet.push(["undoIn", "Undo check-in"]);
    if (v.ph === "chair") quiet.push(["undoSeat", "Back to waiting"]);
    if (v.ph === "done") quiet.push(["undoOut", "Undo check-out"]);
    if (v.p) quiet.push(["openPt", "Patient details"]);
    if (v.p) { (v.ph === "chair" || v.ph === "done" ? ghosts : quiet).push(["printSched", "Print schedule"]); if (/start/i.test(v.a.event) || C.txNo(v.p, S.A, v.a.date, v.a.time) === 1) quiet.push(["startForms", "Print start forms"], ["startEmail", "Start email"]); }
    const oak = v.p ? C.oakState(v.p, ptOf(v.p.mrn), S.A, todayS(), S.cfg) : null;
    if (oak && oak.task && v.ph !== "booked") ghosts.push(["oakTask", "Mark OAK tasked"]);
    const bt = (cls, [a, l]) => `<button class="btn ${cls}" data-act="${a}" data-k="${k}" data-m="${esc(v.a.mrn)}">${l}</button>`;
    return `<aside class="card detail" aria-label="${esc(c)}"><div class="dh"><div><div class="name">${esc(c)}${loc ? `<span class="loc">${esc(loc)}</span>` : ""}</div><div class="meta">${esc(meta)}</div></div></div>
      <p class="lead ${vs.cls}">${esc(vs.lead)}</p>${tl}
      <ol class="steps">${steps}</ol>
      ${vs.pri ? `<button class="btn primary" data-act="${vs.pri[0]}" data-k="${k}"${S.busy ? " disabled" : ""}>${esc(vs.pri[1])}</button>` : ""}
      ${ghosts.length ? `<div class="actions">${ghosts.map(g => bt("ghost", g)).join("")}</div>` : ""}
      ${quiet.length ? `<div class="actions">${quiet.map(g => bt("mute sm", g)).join("")}</div>` : ""}
      ${p ? factsHTML(p, v) : `<p class="note warn">This patient isn’t on the Master Patient Tracker, so there’s no protocol or auth to check. Add them there.</p>`}</aside>`;
  }
  /* before they leave: booking, protocol, auth and cost, read from the Master Patient Tracker and the appointment list */
  function factsHTML(p, v) {
    const a = v ? v.ask : C.asks(p, S.x, null), fut = C.futureOf(S.A, p.mrn, todayS()), T = todayS();
    const dd = (t, cls) => `<dd${cls ? ` class="${cls}"` : ""}>${t}</dd>`;
    const auth = p.authExp == null ? dd("No end date", C.isActive(p) ? "flag" : "") : dd(`Ends ${C.fmtD(p.authExp)}${p.authNo ? `, ${esc(p.authNo)}` : ""}`, p.authExp < T ? "bad" : p.authExp <= T + +S.cfg.AuthWarnDays ? "flag" : "");
    const enc = p.encLeft != null ? dd(`${p.encLeft}${p.encAuth != null ? ` of ${p.encAuth}` : ""} left`, p.encLeft <= 0 ? "bad" : p.encLeft <= +S.cfg.EncountersWarn ? "flag" : "") : dd(esc(p.encLeftTxt || p.encAuthTxt || "Not tracked"));
    const lastAsk = lastSent("protocol", p.code);
    return `<div class="section"><div class="h"><h2>${v ? "Before they leave" : "Visits"}</h2><span>${fut.length ? `${plural(fut.length, "visit")} booked after today` : "nothing booked after today"}</span></div>
      <dl class="facts">
        <dt>Protocol</dt>${dd(esc(p.freq || "No frequency"), p.freq ? "" : "flag")}
        <dt>Confirmed</dt>${dd(p.protoAt != null ? C.fmtD(p.protoAt) : "Not yet", a.proto ? "flag" : "")}
        <dt>Next visit</dt>${dd(fut[0] ? `${C.fmtD(fut[0].date)} at ${C.fmtT(fut[0].time)}` : "None booked", fut.length ? "" : C.isActive(p) && !/prn/i.test(p.freq) ? "bad" : "")}
        <dt>Booked through</dt>${dd(a.thru != null ? C.fmtD(a.thru) : "—", a.book && a.thru != null ? "flag" : "")}
        ${lastAsk ? `<dt>Last asked</dt>${dd(`${dateOfIso(lastAsk.at)} at ${timeOf(lastAsk.at)}${lastAsk.by ? `, ${esc(lastAsk.by)}` : ""}`)}` : ""}
        <dt>Auth</dt>${auth}<dt>Visits on auth</dt>${enc}
        <dt>Infusion</dt>${dd(esc([p.infusion && p.infusion !== "0" ? `$${p.infusion}` : "", p.ins2].filter(Boolean).join(", ") || "—"))}
        <dt>Copay</dt>${dd(esc(p.copay && p.copay !== "0" ? `$${p.copay}` : p.copay === "0" ? "None" : "—"))}
        ${v && S.payOn ? (() => { const L = C.ledger(p.mrn, S.pays), t = ptOf(p.mrn); return `<dt>Balance</dt>${dd(L.balance > 0 ? `${C.money(L.balance)}${t?.card ? `, card on file${t.okCollect ? ", okay to collect" : ""}` : ""}` : "None", L.balance > 0 ? "flag" : "")}`; })() : ""}
      </dl></div>`;
  }

  /* ---------- Schedule ---------- */
  function scheduleHTML() {
    const T = todayS(), start = C.weekStart(T) + 7 * S.wk, F = floor();
    /* open template starts from today on, and holds NextGen doesn't have yet */
    const DS = new Map([1, 2, 3, 4, 5].map(i => [start + i, C.daySlots(start + i, S.X)]));
    const openIn = (d, t) => d < T ? [] : DS.get(d).slots.filter(x => x.open && x.time >= t && x.time < t + 30);
    const holdIn = (d, t) => DS.get(d).bookings.filter(b => b.kind === "hold" && b.time >= t && b.time < t + 30);
    const W = C.week(S.A, start, S.PM, [...DS.values()].flatMap(x => x.slots.map(y => y.time).concat(x.bookings.map(b => b.time))));
    const nOpen = [...DS.entries()].reduce((n, [d, x]) => n + (d < T ? 0 : x.slots.filter(y => y.openIndex).length), 0);
    const stateOf = a => {
      if (C.isCancelled(a)) return "cx";
      if (/no ?show/i.test(a.status)) return "noshow";
      if (a.date < T) return "past";
      if (a.date > T) return "";
      const v = S.list.find(x => x.a === a || (x.a.mrn === a.mrn && x.a.time === a.time)); if (!v) return "";
      if (v.ph === "done" || v.ph === "gone") return "past";
      if (v.ph === "chair") { const c = F.chairs.find(c => c.id === v.st.chair); return c ? (c.state === "track" ? "track" : c.state) : "track"; }
      return v.ph === "waiting" ? "track" : "";
    };
    const blk = a => { const s = stateOf(a), p = a.p, flag = p && a.date >= T && C.futureOf(S.A, p.mrn, a.date).length === 0 && C.isActive(p) && !/prn/i.test(p.freq);
      return `<button class="blk ${s}" data-act="schedPick" data-m="${esc(a.mrn)}" data-d="${a.date}" aria-label="${esc(`${a.code} at ${C.fmtT(a.time)}, ${C.provName(a.provider)}, ${a.status}`)}"><b>${esc(a.code)}</b> ${C.fmtT(a.time)}<small>${esc(C.provName(a.provider))}${/start/i.test(a.event) ? ", start" : ""}${s === "noshow" ? ", no-show" : flag ? ", last booked" : ""}</small></button>`; };
    const hblk = b => { const late = C.holdState(b.h, S.A, Date.now(), S.cfg) === "late";
      return `<button class="blk hold${late ? " late" : ""}" data-act="ipPick" data-m="${esc(b.mrn)}" aria-label="${esc(`Hold for ${b.code} at ${C.fmtT(b.time)}${late ? ", not confirmed booked" : ""}`)}">Hold <b>${esc(b.code)}</b> ${C.fmtT(b.time)}<small>${esc(provLabel(b.provider))}${late ? ", not confirmed" : b.h.by ? `, by ${esc(b.h.by)}` : ""}</small></button>`; };
    const oblk = x => `<div class="blk open">${x.openIndex ? "Open index" : "Open follow-up"} ${C.fmtT(x.time)}<small>${esc(provLabel(x.provider))}</small></div>`;
    const cells = W.rows.map(t => `<div class="tm">${C.fmtT(t)}</div>${W.days.map(d => `<div class="sc${d === T ? " td" : ""}">${W.cell(d, t).map(blk).join("")}${holdIn(d, t).map(hblk).join("")}${openIn(d, t).map(oblk).join("")}</div>`).join("")}`).join("");
    const label = `${C.fmtD(start)} to ${C.fmtD(start + 6)}`;
    return `${head("Schedule", `Week of ${label}: ${plural(W.count, "IVK visit")}${W.count ? `, ${W.days.map((d, i) => W.byDay[i] ? `${["Mon", "Tue", "Wed", "Thu", "Fri"][i]} ${W.byDay[i]}` : "").filter(Boolean).join(", ")}` : ""}${S.grids.length ? `, ${plural(nOpen, "open index start")}` : ""}`, true)}
      <section class="card sched"><div class="wkbar"><span class="lbl">${label}</span>
        <button class="btn ghost sm" data-act="wk" data-v="-1">Previous week</button><button class="btn ghost sm" data-act="wk" data-v="0"${S.wk === 0 ? " disabled" : ""}>This week</button><button class="btn ghost sm" data-act="wk" data-v="1">Next week</button>
        <span class="muted small" style="margin-left:auto">“last booked” = no visit booked after this one</span><button class="btn ghost sm" data-act="printWeek">Print this week</button></div>
      <div class="sgrid"><div></div>${W.days.map(d => `<div class="ch${d === T ? " today" : ""}">${["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"][C.weekday(d) - 1]}<small>${C.fmtD(d)}</small></div>`).join("")}${cells}</div></section>`;
  }

  /* ---------- Index planner ---------- */
  const DAYW = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
  const dshort = d => `${["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"][C.weekday(d)]} ${C.fmtD(d)}`;
  /* a provider as the templates name them ("Ana Rivera, MD"), else as NextGen does ("Dr. Rivera") */
  const provLabel = n => { const k = TM.provKey(n), g = S.grids.find(x => x.key === k); return g ? TM.nameOf(g) : C.provName(n) || "any provider"; };
  const liveHolds = () => (S.holds || []).map(h => Object.assign({}, h, { state: C.holdState(h, S.A, Date.now(), S.cfg) })).filter(h => h.state === "held" || h.state === "late");
  const holdsOf = mrn => (S.holds || []).filter(h => h.mrn === mrn).map(h => Object.assign({}, h, { state: C.holdState(h, S.A, Date.now(), S.cfg), conf: C.holdConfirm(h, S.A, S.raw.apptUpdated) })).filter(h => h.state !== "released");
  /* holds marked booked that the uploaded NextGen report hasn't confirmed yet (pending), or that a newer report doesn't have (missing) */
  const unconfirmed = () => (S.holds || []).filter(h => h.status === "booked").map(h => Object.assign({}, h, { conf: C.holdConfirm(h, S.A, S.raw.apptUpdated) })).filter(h => h.conf === "pending" || h.conf === "missing");
  /* patients already in index, whose index visits can be changed here: Induction, or started within the index weeks */
  const inIndex = p => /induction/i.test(p.status) || (p.start != null && todayS() - p.start <= 7 * (+S.cfg.IndexWeeks + 1));
  const listTime = () => S.raw.apptUpdated != null ? `${C.fmtD(Math.floor(S.raw.apptUpdated))} at ${C.fmtT(C.toMin(S.raw.apptUpdated))}` : "never";
  const uploadBtn = (l, cls) => `<label class="btn ${cls || "primary"}" style="cursor:pointer">${l}<input type="file" id="apptFile" accept=".xls,.xlsx,.csv" hidden></label>`;
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
  /* how many more patients could start index on the templates; worked out once per workbook read */
  function capacityNow() {
    if (S.cap && S.cap.raw === S.raw) return S.cap.v;
    const c = S.cfg, v = S.grids.length ? C.capacity({ today: todayS(), perWeek: +c.IndexPerWeek, weeks: +c.IndexWeeks, minGap: +c.MinGapDays, maxGap: +c.MaxGapDays, X: S.X }, +c.CapacityWeeks) : null;
    S.cap = { raw: S.raw, v }; return v;
  }

  function indexHTML() {
    const L = ipList(), H = liveHolds(), late = H.filter(h => h.state === "late"), cap = capacityNow();
    const byPt = new Map(); for (const h of H) byPt.set(h.mrn, (byPt.get(h.mrn) || []).concat(h));
    const sub = [`${plural(L.filter(o => !o.holds.length).length, "patient")} to plan`, byPt.size ? `${plural(byPt.size, "patient")} with holds to book` : "",
      late.length ? `<strong>${plural(new Set(late.map(h => h.mrn)).size, "patient")}’s holds not confirmed after ${S.cfg.HoldMin} min</strong>` : "",
      cap ? `room to start ${plural(cap.starts, "more index patient")} in the next ${S.cfg.CapacityWeeks} weeks` : "no template yet"].filter(Boolean).join(", ");
    const chips = `<div class="chips">${[["plan", "Plan"], ["template", "Template"]].map(([k, l]) => `<button class="chip" data-act="ipView" data-v="${k}" aria-pressed="${S.ipView === k}">${l}${k === "template" ? `<span class="n">${S.grids.length}</span>` : ""}</button>`).join("")}</div>`;
    if (S.ipView === "template") return `${setupBanner()}${head("Index planner", sub, false)}${chips}${templateHTML()}`;
    if (!S.ipSel || !L.find(o => o.mrn === S.ipSel)) S.ipSel = (L.find(o => o.holds.some(h => h.state === "late")) || L.find(o => !o.holds.length) || L[0] || {}).mrn || null;
    const cur = L.find(o => o.mrn === S.ipSel), model = cur && !cur.holds.length ? ipModel(cur) : null;
    const row = o => { const hl = holdLead(o.holds), booked = o.holds.length && o.holds.every(h => h.state === "booked");
      const miss = o.holds.some(h => h.conf === "missing"), pend = o.holds.some(h => h.conf === "pending");
      const st = hl ? `<span class="f${hl.cls === "late" ? " bad" : ""}">${hl.cls === "late" ? "Holds not confirmed" : `${plural(o.holds.filter(h => h.state !== "booked").length, "hold")} to book`}</span>`
        : booked ? (miss ? `<span class="f bad">Not on the NextGen list</span>` : pend ? `<span class="f">Booked: upload to confirm</span>` : `<span class="ok">On the appointment list</span>`) : "Not held yet";
      return `<button class="crow" data-act="ipPick" data-m="${esc(o.mrn)}" aria-pressed="${S.ipSel === o.mrn}"><span class="p">${esc(o.code)}</span><span class="proto">${esc(o.from)}${o.p ? `<br>${esc(o.p.status || "")}` : ""}</span><span class="thru">${o.pl?.start != null ? `Start ${C.fmtD(o.pl.start)}` : o.c?.start != null && o.c.start > todayS() ? `Start ${C.fmtD(o.c.start)}` : "As soon as possible"}</span><span class="flags">${st}</span></button>`; };
    /* new index patients aren't on the Master Patient Tracker yet: they're planned by LASFIR and MRN. The tracker list is only for
       changing the index visits of a patient already in index. */
    const opts = S.P.filter(p => inIndex(p) && !L.find(o => o.mrn === p.mrn)).sort((a, b) => a.code.localeCompare(b.code)).map(p => `<option value="${esc(p.mrn)}">${esc(p.code)}${p.status ? `, ${esc(p.status)}` : ""}${p.start != null ? `, started ${C.fmtD(p.start)}` : ""}</option>`).join("");
    const U = unconfirmed(), uMiss = U.filter(h => h.conf === "missing"), uPts = new Set(U.map(h => h.mrn)).size;
    const confirmBanner = U.length ? `<div class="msg ${uMiss.length ? "bad" : "warn"}"><b>${uMiss.length ? `${plural(new Set(uMiss.map(h => h.mrn)).size, "patient")}’s booked visits aren’t on the NextGen list uploaded ${listTime()}.` : `${plural(uPts, "patient")}’s index visits are marked booked, not confirmed yet.`}</b>
      ${uMiss.length ? "Check them in NextGen, then upload the report again." : `Run the NextGen appointment report and upload it to confirm the visits are booked (last upload ${listTime()}).`} <div class="actions">${uploadBtn("Upload the NextGen report", "primary inl")}</div></div>` : "";
    return `${setupBanner()}${head("Index planner", sub, false)}${chips}${confirmBanner}
      ${S.grids.length ? "" : `<div class="msg warn"><b>No provider template yet.</b> Plans are made from each provider’s start/index slots and the post checks that line up with them. <button class="btn ghost sm" data-act="ipView" data-v="template">Set up the template</button></div>`}
      <div class="cols"><div class="stack"><section class="card plist"><div class="chead"><span>Patient</span><span>From</span><span>Start</span><span>Holds</span></div>
        ${L.map(row).join("") || `<p class="empty" style="padding:10px 14px">Nobody to plan. Add a new patient by LASFIR and MRN below. Approved consults waiting for a treatment start show here too.</p>`}
        <div class="addpt new"><span class="lbl">New patient</span><input id="ipNewCode" placeholder="LASFIR" aria-label="New patient’s LASFIR code" maxlength="8" autocomplete="off" spellcheck="false"><input id="ipNewMrn" placeholder="MRN" aria-label="New patient’s MRN" inputmode="numeric" autocomplete="off"><button class="btn ghost sm line" data-act="ipNewPt">Plan</button></div>
        <div class="addpt"><select id="ipAdd" aria-label="Change a patient in index"><option value="">Change the visits of a patient in index…</option>${opts}</select><button class="btn ghost sm line" data-act="ipAddPt">Change</button></div></section>
        ${holdsHTML(byPt)}${accessHTML(model)}</div>
      <div>${cur ? ipCardHTML(cur, model) : ""}</div></div>`;
  }
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
    const mine = new Map((model?.res.items || []).map(x => [`${x.date}|${x.time}|${x.pkey}`, x])), many = S.X.providers > 1;
    const tile = (d, c, ds) => {
      const s = ds.slots.find(x => x.time === c.time && x.pkey === c.pkey), b = c.by, me = mine.get(`${d}|${c.time}|${c.pkey}`), ivk = IVK_USES.includes(c.use);
      let t, cls = ivk ? `tile ${c.use}` : "tile gen", pick = false;
      if (!ivk) t = b ? (b.kind === "other" ? b.event : b.code) : c.type;
      else if (c.use === "start" || c.use === "followup") {
        if (b && b.kind === "hold") { t = `Hold ${b.code}`; cls += " held"; }
        else if (b) { t = b.kind === "other" ? b.event : b.code; cls += " booked"; }
        else if (s && s.full) { t = "Chairs full"; cls += " booked"; }
        else if (s && s.openIndex) { t = "Index open"; pick = d > T; }
        else t = c.use === "start" ? "Follow-up only" : "Follow-up";
      } else if (b) { t = b.kind === "other" ? b.event : `Post check, ${b.code}`; cls += " booked"; }
      else t = c.use === "post" ? "Post check" : "Flex post check";
      if (me) cls += " mine";
      const label = `${DAYW[C.weekday(d) - 1]} ${C.fmtT(c.time)}, ${c.who}: ${t}${me ? `, treatment ${me.n}` : ""}`;
      const body = `${me ? `<b class="nb">${me.n}</b>` : ""}<span>${esc(t)}</span>${many && ivk ? `<small>${esc(c.provider)}</small>` : ""}`;
      const st = ivk ? ` style="--c:${c.fill}"` : "";
      return pick ? `<button class="${cls} pick" data-act="ipTile" data-d="${d}" data-t="${c.time}" data-p="${esc(c.provider)}"${st} aria-label="${esc(label)}">${body}</button>` : `<div class="${cls}"${st} title="${esc(label)}">${body}</div>`;
    };
    const rows = times.map(t => `<div class="tm">${C.fmtT(t)}</div>${days.map((d, i) => `<div class="ac${d === T ? " td" : ""}${d < T ? " past" : ""}">${DS[i].cells.filter(c => c.time === t).map(c => tile(d, c, DS[i])).join("")}</div>`).join("")}`).join("");
    const nIdx = DS.reduce((n, x, i) => n + (days[i] > T ? x.slots.filter(s => s.openIndex).length : 0), 0);
    const sw = (use, l) => { const t = S.types.find(x => x.use === use); return `<li><span class="tsw" style="--c:${t ? t.fill : "#999"}"></span>${l}</li>`; };
    return `<section class="card access"><div class="lh"><h2>Access</h2><span>${plural(nIdx, "open index start")}</span></div>
      <div class="wkbar"><span class="lbl">${C.fmtD(start)} to ${C.fmtD(start + 6)}</span><button class="btn ghost sm" data-act="ipWk" data-v="-1"${S.ipWk <= 0 ? " disabled" : ""}>Previous week</button><button class="btn ghost sm" data-act="ipWk" data-v="0"${S.ipWk === 0 ? " disabled" : ""}>This week</button><button class="btn ghost sm" data-act="ipWk" data-v="1">Next week</button></div>
      <div class="agrid"><div></div>${days.map(d => `<div class="ch${d === T ? " today" : ""}">${DAYW[C.weekday(d) - 1]}<small>${C.fmtD(d)}</small></div>`).join("")}${rows}</div>
      ${(() => { const off = S.EX.filter(e => e.from <= start + 6 && e.to >= start); return off.length ? `<p class="note warn">Time off this week: ${off.map(e => `${esc(exLabel(e.provider))} ${exDates(e)}${e.cov ? ` (covered by ${esc(exLabel(e.cov))}${e.useTpl ? ", same template" : ""})` : ""}`).join("; ")}.</p>` : ""; })()}
      <ul class="key">${sw("start", "Start/index")}${sw("followup", "Follow-up")}${sw("post", "IVK post check")}${sw("flexpost", "Flex post check")}<li><span class="tsw gen"></span>Not IVK</li></ul>
      <p class="note">${model ? (S.ipAlt != null ? `Pick an open index start for treatment ${S.ipAlt + 1}.` : "Pick an open index start to make it treatment 1.") : "Open index starts have a free post check at their end and a chair for the 2 hours."}</p></section>`;
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
      return `<aside class="card detail" aria-label="${esc(o.code)}"><div class="dh"><div><div class="name">${esc(o.code)}</div><div class="meta">${esc(meta)}</div></div></div>
        <p class="lead ${lead.cls}">${esc(lead.t)}</p>
        <ul class="steps">${all.map(st).join("")}</ul>
        ${open.length ? `<button class="btn primary" data-act="ipBooked" data-m="${esc(o.mrn)}">Booked in NextGen</button><div class="actions"><button class="btn ghost" data-act="ipRelease" data-m="${esc(o.mrn)}">Release ${open.length === 1 ? "the hold" : "the holds"}</button></div>`
          : miss.length || pend.length ? `${uploadBtn("Upload the NextGen report")}${miss.length ? `<div class="actions"><button class="btn ghost" data-act="ipRelease" data-m="${esc(o.mrn)}">Release the missing ${miss.length === 1 ? "visit" : "visits"}</button></div>` : ""}`
          : `<button class="btn ghost" data-act="ipDone" data-m="${esc(o.mrn)}">Take off the list</button>`}
        <div class="actions"><button class="btn ghost sm line" data-act="printSched" data-m="${esc(o.mrn)}">Print schedule</button><button class="btn mute sm" data-act="startForms" data-m="${esc(o.mrn)}">Print start forms</button><button class="btn mute sm" data-act="startEmail" data-m="${esc(o.mrn)}">Start email</button></div>
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
    const alt = n => { const A = C.altSlots(r, n, +cfg.MinGapDays).slice(0, 12);
      return `<div class="alts">${A.length ? A.map((s, i) => `<button data-act="ipUse" data-n="${n}" data-i="${i}">${dshort(s.date)}<small>${C.fmtT(s.time)}, ${esc(provLabel(s.provider))}${s.postUse === "flexpost" ? ", flex post check" : ""}</small></button>`).join("") : `<span class="muted small">No other open index start keeps a day off on each side.</span>`}</div>`; };
    const step = x => `<li class="step ${x.fixed ? "done" : x.long ? "soon" : "now"}"><span class="k">${x.fixed ? "" : x.n}</span><span><span class="t">${dshort(x.date)}, ${C.fmtT(x.time)}</span>
        <span class="m">Week ${x.week}, ${esc(provLabel(x.provider))}, ${x.fixed ? "already booked, " : ""}post check ${C.fmtT(x.post)}${x.postUse === "flexpost" ? " (flex)" : ""}${x.gap != null ? `, ${plural(x.gap, "day")} after the last` : ""}${x.long ? ` (over ${cfg.MaxGapDays})` : ""}</span>
        ${S.ipAlt === x.n - 1 ? alt(x.n - 1) : ""}</span><span class="w">${x.fixed ? "" : `<button class="btn mute sm" data-act="ipAlt" data-n="${x.n - 1}">${S.ipAlt === x.n - 1 ? "Keep" : "Change"}</button>`}</span></li>`;
    const last = it[it.length - 1], afterD = last ? last.date + (C.freqDays(cfg.AfterIndex) || 7) : null;
    return `<aside class="card detail" aria-label="${esc(o.code)}"><div class="dh"><div><div class="name">${esc(o.code)}</div><div class="meta">${esc(meta)}</div></div></div>
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
    const chips = `<div class="chips">${S.grids.map((x, i) => `<button class="chip" data-act="tplPick" data-v="${i}" aria-pressed="${x === g}">${esc(TM.nameOf(x))}${x.effFrom != null ? `<span class="n">from ${C.fmtD(x.effFrom)}</span>` : ""}</button>`).join("")}<button class="btn ghost sm line" data-act="tplAdd">Add a provider</button></div>`;
    if (!g) return `${chips}<section class="card"><p class="empty">No provider grids yet. Add one for each provider who sees IVK patients; where two providers work the same hours, each gets a grid and the planner shows both.</p></section>`;
    const T = TM.tally(g, S.types), warn = TM.checks(g, S.types, +cfg.PostCheckMin);
    const idx = [1, 2, 3, 4, 5].map(d => TM.startsOn(g, d, S.types, +cfg.PostCheckMin).filter(s => s.index).length);
    const sel = S.tsel && S.tsel.g === g.top ? S.tsel : null, nSel = sel ? g.dayCols.filter(x => inRect({ ...sel, t0: 0, t1: 1e4 }, x.d, 0)).length * g.rows.filter(x => inRect({ ...sel, d0: 0, d1: 9 }, 0, x.t)).length : 0;
    const cell = (d, t) => { const x = g.cells.get(d + "|" + t), ty = x ? TM.typeOf(S.types, x.type) : null;
      return `<button class="tc${inRect(sel, d, t) ? " sel" : ""}" data-d="${d}" data-t="${t}"${ty ? ` style="background:${ty.fill};color:${ty.font}"` : ""} aria-label="${TM.DAYS[d]} ${C.fmtT(t)}: ${esc(x ? x.type : "empty")}">${x ? esc(x.type) : ""}</button>`; };
    const editor = `<section class="card tedit"><div class="lh"><h2>${esc(TM.nameOf(g))}</h2><span>${esc(g.location || cfg.Clinic)}${g.modified != null ? `, last modified ${C.fmtD(g.modified)}` : ""}</span></div>
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
    return `${chips}<div class="cols"><div class="stack">${editor}${ptoHTML()}</div><div class="stack">${tally}${typesHTML()}</div></div>`;
  }
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
  const CHIPS = [["need", "Needs attention"], ["active", "Active"], ["paused", "Paused"], ["disc", "Discontinued"], ["all", "All"]];
  function chipOf(k, p, n) {
    if (k === "need") return n.length > 0;
    if (k === "active") return C.isActive(p);
    if (k === "paused") return /paused/i.test(p.status);
    if (k === "disc") return /discontinued/i.test(p.status);
    return true;
  }
  function patientsHTML() {
    const rows = S.P.map(p => ({ p, n: C.needs(p, S.x) })), L = rows.map(r => ({ ...r, loud: C.loud(r.n) }));
    const counts = Object.fromEntries(CHIPS.map(([k]) => [k, L.filter(r => chipOf(k, r.p, r.loud)).length]));
    const q = S.q.trim().toUpperCase(), shown = L.filter(r => chipOf(S.chip, r.p, r.loud) && (!q || r.p.code.includes(q)))
      .sort((a, b) => (b.loud.some(y => y.bad) - a.loud.some(y => y.bad)) || (b.loud.length - a.loud.length) || a.p.code.localeCompare(b.p.code));
    if (!S.pSel || !shown.find(r => r.p.mrn === S.pSel)) S.pSel = shown[0]?.p.mrn || null;
    const red = L.filter(r => r.loud.some(y => y.bad)).length, act = L.filter(r => C.isActive(r.p)).length;
    const sub = [`${act} active`, red ? `<strong>${red} with something blocking the next infusion</strong>` : "", `${counts.need} need attention`].filter(Boolean).join(", ");
    const row = r => { const fut = C.futureOf(S.A, r.p.mrn, todayS() - 1);
      const thru = fut.length ? `Through ${C.fmtD(fut[fut.length - 1].date)}` : "Nothing booked";
      const fl = r.loud.length ? r.loud.map(y => `<span class="f${y.bad ? " bad" : ""}">${esc(y.t)}</span>`).join(", ") : `<span class="ok">${esc(r.n[0]?.t || "On track")}</span>`;
      return `<button class="crow" data-act="pSel" data-m="${esc(r.p.mrn)}" aria-pressed="${S.pSel === r.p.mrn}"><span class="p">${esc(r.p.code)}</span><span class="proto">${esc(r.p.freq || "No frequency")}<br>${esc(r.p.status || "No status")}</span><span class="thru">${thru}</span><span class="flags">${fl}</span></button>`; };
    const cur = L.find(r => r.p.mrn === S.pSel);
    return `${setupBanner()}${head("Patients", sub, false)}
      <div class="chips">${CHIPS.map(([k, l]) => `<button class="chip" data-act="chip" data-v="${k}" aria-pressed="${S.chip === k}">${l}<span class="n">${counts[k]}</span></button>`).join("")}
        <input class="search" id="ptSearch" type="search" placeholder="Find a code" aria-label="Find a patient by code" value="${esc(S.q)}" autocomplete="off"></div>
      <div class="cols"><section class="card plist"><div class="chead"><span>Patient</span><span>Protocol</span><span>Booked</span><span>Needs</span></div>
        ${shown.map(row).join("") || `<p class="empty" style="padding:10px 14px">${S.chip === "need" ? "Nothing here. Every patient in this group is on track." : "No patients in this group."}</p>`}</section>
      <div>${cur ? patientCardHTML(cur.p, cur.n) : ""}</div></div>`;
  }
  /* a provider can ask for a post check on every visit (acuity): the patient is then booked only in starts that end on a post check */
  function postCheckHTML(p) {
    const L = p.reqPost && S.grids.length ? C.postCheckSlots(S.X, todayS() + 1, 5) : [];
    return `<div class="section"><div class="h"><h2>Post check</h2><span>${p.reqPost ? "required" : "not required"}</span></div>
      <label class="opt"><input type="checkbox" id="reqPost" data-m="${esc(p.mrn)}"${p.reqPost ? " checked" : ""}> Requires a post-check<small>Tick when the provider asks for one. Book them in a start whose last 30 min is a post check or flex post check.</small></label>
      ${p.reqPost ? `<p class="note">${L.length ? `Next open with a post check: ${L.map(s => `${dshort(s.date)} ${C.fmtT(s.time)}`).join(", ")}.` : S.grids.length ? "No open start with a free post check in the next 6 weeks." : "Set up the template to list open times with a post check."}</p>` : ""}</div>`;
  }
  function patientCardHTML(p, n) {
    const top = C.loud(n)[0], tv = S.list.find(v => v.a.mrn === p.mrn);
    const meta = [p.status || "No status", p.loc, p.start != null ? `started ${C.fmtD(p.start)}` : ""].filter(Boolean).join(", ");
    const pri = top?.act ? top.act : ["edit", "Edit tracker fields"], m = esc(p.mrn), hasStart = !!startVisit(p);
    const b = (cls, act, l) => `<button class="btn ${cls}" data-act="${act}" data-m="${m}">${l}</button>`;
    return `<aside class="card detail" aria-label="${esc(p.code)}"><div class="dh"><div><div class="name">${esc(p.code)}</div><div class="meta">${esc(meta)}</div></div></div>
      <p class="lead ${top ? (top.bad ? "late" : "soon") : ""}">${esc(top ? top.t : tv ? `Here today at ${C.fmtT(tv.a.time)}` : "Nothing needed right now")}</p>
      ${C.loud(n).length > 1 ? `<ul class="steps">${C.loud(n).slice(1).map(y => `<li class="step ${y.bad ? "late" : "soon"}"><span class="k"></span><span><span class="t">${esc(y.t)}</span></span><span class="w">${y.act && y.act[1] !== pri[1] ? `<button class="btn mute sm" data-act="${y.act[1]}" data-m="${m}">${esc(y.act[0])}</button>` : ""}</span></li>`).join("")}</ul>` : ""}
      ${b("primary", pri[1], esc(pri[0]))}
      <div class="actions">${pri[1] !== "edit" ? b("ghost", "edit", "Edit tracker fields") : ""}${pri[1] !== "proto" ? b("ghost", "proto", "Confirm protocol") : ""}${pri[1] !== "ask" && C.isActive(p) ? b("ghost", "ask", "Ask the group chat") : ""}${tv ? `<button class="btn ghost" data-act="toToday" data-k="${esc(tv.key)}">Today’s visit</button>` : ""}</div>
      <div class="actions">${b("ghost sm line", "printSched", "Print schedule")}${b("mute sm", "startForms", hasStart ? "Print start forms" : "Start forms")}${b("mute sm", "startEmail", "Start email")}</div>
      ${factsHTML(p, null)}
      ${payHTML(p)}${availHTML(p)}${oakHTML(p)}
      ${postCheckHTML(p)}
      <div class="section"><div class="h"><h2>Coverage</h2><button class="btn mute sm" data-act="edit" data-m="${m}">Edit</button></div><dl class="facts">
        <dt>Primary</dt><dd>${esc(p.ins1 || "—")}</dd><dt>2026 BIDF</dt><dd${!p.bidf && C.isActive(p) ? ' class="flag"' : ""}>${esc(p.bidf || "Not recorded")}</dd>
        <dt>Auth start</dt><dd>${p.authStart != null ? C.fmtD(p.authStart) : "—"}</dd>${p.pcn ? `<dt>PCN</dt><dd>${esc(p.pcn)}${p.task ? `, ${esc(p.task)}` : ""}</dd>` : ""}
        ${p.notes ? `<dt>Clinic notes</dt><dd>${esc(p.notes)}</dd>` : ""}${p.pcnNotes ? `<dt>PCN notes</dt><dd>${esc(p.pcnNotes)}</dd>` : ""}</dl></div></aside>`;
  }
  /* the treatment start on the list (today or later), else the first visit of an Induction patient */
  const startVisit = p => C.futureOf(S.A, p.mrn, todayS() - 1).find(a => /start/i.test(a.event)) || (/induction/i.test(p.status) ? C.futureOf(S.A, p.mrn, todayS() - 1)[0] : null) || null;
  const ptOf = mrn => S.PT.get(mrn) || null;
  const needTable = (n, what) => S.raw.tables?.own?.[n] ? "" : `<p class="note warn">${what} need the newer tables. Use <b>Update the workbook</b> at the top of the page or in Settings.</p>`;
  /* Payments: the balance from missed payments, what each visit owes, card on file and okay to collect, reminders, and visits
     nobody has said were paid */
  function payHTML(p) {
    const t = ptOf(p.mrn), L = C.ledger(p.mrn, S.pays), due = C.dueOf(p), m = esc(p.mrn), T = todayS(), conf = S.payOn ? C.toConfirm(p, S.A, S.pays, T, S.payStart) : [];
    const last = L.lastReminder, remDue = L.balance > 0 && (!last || T - (last.date ?? 0) >= +S.cfg.ReminderDays);
    const chk = (k, on, l, s) => `<label class="opt"><input type="checkbox" data-pt="${k}" data-m="${m}"${on ? " checked" : ""}${S.raw.tables?.own?.tblIvkPatient ? "" : " disabled"}> ${l}${s ? `<small>${s}</small>` : ""}</label>`;
    const miss = L.open.map(r => `<li class="step soon"><span class="k"></span><span><span class="t">${r.date != null ? dshort(r.date) : "Visit"}: ${C.money(r.left)} not collected</span><span class="m">${r.left < r.amount ? `${C.money(r.amount - r.left)} of ${C.money(r.amount)} paid since` : esc(r.by ? `Marked by ${r.by}` : "")}</span></span>
      <span class="w"><button class="btn mute sm" data-act="payMissed" data-m="${m}" data-id="${esc(r.id)}">Collected</button> <button class="btn mute sm" data-act="payWaive" data-m="${m}" data-id="${esc(r.id)}">Waive</button></span></li>`).join("");
    const cf = conf.slice(0, 8).map(a => `<li class="step now"><span class="k"></span><span><span class="t">${dshort(a.date)} ${C.fmtT(a.time)}: ${C.money(due.total)} due</span><span class="m">Nobody said whether it was collected</span></span>
      <span class="w"><button class="btn mute sm" data-act="payConf" data-v="collected" data-m="${m}" data-d="${a.date}" data-t="${a.time}">Collected</button> <button class="btn mute sm" data-act="payConf" data-v="missed" data-m="${m}" data-d="${a.date}" data-t="${a.time}">Not collected</button></span></li>`).join("");
    return `<div class="section"><div class="h"><h2>Payments</h2><span${L.balance > 0 ? ' class="flagt"' : ""}>${L.balance > 0 ? `Balance ${C.money(L.balance)}` : "No balance"}</span></div>
      ${needTable("tblIvkPay", "Payments")}
      <dl class="facts"><dt>Due each visit</dt><dd>${due.total ? `${C.money(due.total)}${due.copay && due.infusion ? ` (${C.money(due.copay)} copay, ${C.money(due.infusion)} infusion)` : due.infusion ? " infusion" : " copay"}` : "Nothing"}</dd>
        <dt>Collected</dt><dd>${C.money(L.collected) || "$0"}</dd>
        <dt>Last reminder</dt><dd${remDue ? ' class="flag"' : ""}>${last ? `${C.fmtD(last.date)}${last.method ? `, ${esc(last.method)}` : ""}${last.by ? `, ${esc(last.by)}` : ""}` : L.balance > 0 ? "None sent" : "—"}</dd></dl>
      ${chk("CardOnFile", t?.card, "Luma card on file")}${chk("OkToCollect", t?.okCollect, "Okay to collect", "The patient agreed we can charge the card on file for a balance")}
      ${miss || cf ? `<ul class="steps">${miss}${cf}</ul>${conf.length > 8 ? `<p class="note">And ${plural(conf.length - 8, "older visit")} to confirm.</p>` : ""}` : ""}
      <div class="actions"><button class="btn ghost sm line" data-act="pay" data-m="${m}"${S.payOn ? "" : " disabled"}>Record a payment or reminder</button></div></div>`;
  }
  /* Availability, as the patient availability form has it: each weekday morning and afternoon, effective and expiry dates, dates away */
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
  /* Oakland patients: their index here, then maintenance back at Oakland; OAK is tasked to book it at treatment OakRemindAt */
  function oakHTML(p) {
    const t = ptOf(p.mrn), o = C.oakState(p, t, S.A, todayS(), S.cfg), m = esc(p.mrn), auto = t?.oak == null && /oak/i.test(p.loc);
    const note = !o ? "" : o.unknown && o.tasked == null ? "This course started before the appointment list begins, so treatments can't be counted here. Mark OAK tasked once maintenance is booked at Oakland." : o.tasked != null ? `OAK tasked to book maintenance ${C.fmtD(o.tasked)}.` : o.done >= +S.cfg.OakRemindAt ? `<b>Task OAK to book maintenance now.</b>` : `At treatment ${S.cfg.OakRemindAt}, task OAK to book maintenance.`;
    return `<div class="section"><div class="h"><h2>Oakland</h2><span>${o ? (o.unknown ? "treatments not countable" : `${o.done} of ${o.idx} index treatments done`) : "Castro patient"}</span></div>
      <label class="opt"><input type="checkbox" data-pt="Oakland" data-m="${m}"${o ? " checked" : ""}${S.raw.tables?.own?.tblIvkPatient ? "" : " disabled"}> Oakland patient<small>Index (the first ${esc(S.cfg.OakIndex)}) at ${esc(S.cfg.Clinic)}, then back to Oakland for maintenance.${auto ? " Ticked because the home clinic is Oakland." : ""}</small></label>
      ${o ? `<p class="note${o.task || o.back ? " warn" : ""}">${note}${o.back ? ` Index is done and ${plural(o.after.length, "visit")} ${o.after.length === 1 ? "is" : "are"} still booked here.` : ""}</p>
        <div class="actions">${o.tasked == null ? `<button class="btn ghost sm line" data-act="oakTask" data-m="${m}">Mark OAK tasked</button>` : `<button class="btn mute sm" data-act="oakUntask" data-m="${m}">Undo OAK tasked</button>`}</div>` : ""}</div>`;
  }

  /* ---------- Consults ---------- */
  /* Waitlist: everyone waiting for a treatment start, in line by consult date or the order set here; the Consults tab beneath it */
  const WAIT_TXT = { ready: "Ready to start", referral: "Waiting on the referral", auth: "Waiting on the authorization" };
  function consultsHTML() {
    const L = S.CS.slice().sort((a, b) => ({ act: 0, wait: 1, done: 2, closed: 3 }[a.state] - { act: 0, wait: 1, done: 2, closed: 3 }[b.state]) || (b.days ?? 0) - (a.days ?? 0));
    const n = s => L.filter(c => c.state === s).length, W = S.WL, ok = !!S.raw.tables?.own?.tblIvkPatient, planned = new Set(ipList().filter(o => o.pl || o.holds.length).map(o => o.mrn));
    const sub = [plural(W.list.length, "patient") + " on the waitlist", W.next ? `${W.next.code} starts next` : W.list.length ? "nobody ready to start yet" : "",
      n("done") ? `${n("done")} scheduled` : ""].filter(Boolean).join(", ");
    if (!S.wSel || !W.list.find(w => w.mrn === S.wSel)) S.wSel = W.list[0]?.mrn || null;
    const wrow = (w, i) => { const c = w.c, why = w.next ? `<span class="ok">Starts next</span>` : w.on === "referral" ? `<span class="f">Referral: keeps their place</span>`
        : w.on === "auth" ? `<span class="f">${w.appeal ? "Appeal: the next ready patient starts first" : "Authorization pending"}</span>` : w.ahead ? "" : "Ready";
      return `<div class="wrow${S.wSel === w.mrn ? " cur" : ""}" role="button" tabindex="0" data-act="wSel" data-m="${esc(w.mrn)}" aria-pressed="${S.wSel === w.mrn}"><span class="pos">${w.pos}</span><span class="p">${esc(w.code || "—")}</span>
        <span class="proto">${c.consult != null ? `Consult ${C.fmtD(c.consult)}` : "No consult date"}<br><span class="muted small">${esc(c.provider || "")}</span></span>
        <span class="thru">${esc(WAIT_TXT[w.on])}<br><span class="muted small">${esc(c.stage || "")}</span></span><span class="flags">${why}${planned.has(w.mrn) ? ` <span class="muted small">planning</span>` : ""}</span>
        <span class="mv"><button class="btn mute sm" data-act="wMove" data-m="${esc(w.mrn)}" data-v="-1" aria-label="Move ${esc(w.code)} up"${i === 0 || !ok ? " disabled" : ""}>↑</button><button class="btn mute sm" data-act="wMove" data-m="${esc(w.mrn)}" data-v="1" aria-label="Move ${esc(w.code)} down"${i === W.list.length - 1 || !ok ? " disabled" : ""}>↓</button></span></div>`; };
    const cur = W.list.find(w => w.mrn === S.wSel);
    const row = c => `<div class="crow" role="button" tabindex="0" data-act="editConsult" data-m="${esc(c.mrn)}"${c.mrn ? "" : ' style="cursor:default"'}><span class="p">${esc(c.code || "—")}</span><span class="proto">${c.consult != null ? `Consult ${C.fmtD(c.consult)}` : "No consult date"}<br>${esc(c.provider || "")}</span><span class="thru">${esc(c.stage || "No stage")}${c.days != null ? `<br><span class="muted small">${plural(c.days, "day")} in stage</span>` : ""}</span>
      <span class="flags">${c.flag ? `<span class="f">${esc(c.flag)}</span>` : c.state === "done" ? `<span class="ok">Scheduled</span>` : c.state === "closed" ? "Not proceeding" : "Waiting"}</span></div>`;
    return `${setupBanner()}${head("Waitlist", sub, false)}${!S.raw.tables?.consults ? `<div class="msg warn">There’s no table on a Consults tab in this workbook.</div>` : ""}
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
    return `<aside class="card detail" aria-label="${esc(w.code)}"><div class="dh"><div><div class="name">${esc(w.code)}<span class="loc">#${w.pos}</span></div><div class="meta">${esc([c.loc, c.provider, c.consult != null ? `consult ${C.fmtD(c.consult)}` : ""].filter(Boolean).join(", "))}</div></div></div>
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
        <div class="actions"><label class="btn primary inl" style="cursor:pointer">Upload a NextGen export<input type="file" id="apptFile" accept=".xls,.xlsx,.csv" hidden></label></div>
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
        ${tplF("arrival", "Arrival")}${tplF("seated", "Seated")}${tplF("protocol", "Confirm protocol and book")}${tplF("protocolOnly", "Confirm protocol only")}${tplF("bookOnly", "Book only")}
        <label class="fld">Start email subject<input id="tp_startSubject" value="${esc(S.tpl.startSubject)}"></label>
        <label class="fld">Start email (HTML)<textarea id="tp_startEmail" rows="9" spellcheck="false">${esc(S.tpl.startEmail)}</textarea><span class="hint">The Castro desk’s “Upcoming Appointment(s)” email. {mrn}, {first} (first name), {clinic}, {firstVisit} (“Wednesday, November 4th, 2026, at 10:00 AM”), {address}, and {schedule} where the appointment table goes.</span></label>
        <label class="fld">Clinic address in the start email<input id="cf_ClinicAddress" value="${esc(S.cfg.ClinicAddress)}"></label>
        <div class="two"><label class="fld">CC on every start email<input id="cf_StartCc" value="${esc(S.cfg.StartCc)}" autocomplete="off"></label><label class="fld">BCC on every start email<input id="cf_StartBcc" value="${esc(S.cfg.StartBcc)}" autocomplete="off"></label></div>
        <div class="actions"><button class="btn ghost sm line" data-act="saveTpl">Save messages</button><button class="btn mute sm" data-act="resetTpl">Use the standard wording</button></div></section>

      <section class="card set"><h3>Chair map</h3><p>${P ? `${esc(P.name || "Map")} with ${plural(P.bays.length, "chair")}: ${P.bays.map(b => esc(b.id)).join(", ")}.` : "No chair map in this workbook. Today shows chair tiles instead."} Draw one in the <a href="${BUILDER}" target="_blank" rel="noopener">Bay Map Builder</a>, download it, and load the file here. Bay numbers become chair numbers.</p>
        <div class="actions"><label class="btn ghost sm line" style="cursor:pointer">Load a map file<input type="file" id="mapFile" accept=".json,application/json" hidden></label>${P ? `<button class="btn ghost sm line" data-act="mapDown">Download this map</button><button class="btn mute sm" data-act="mapRemove">Remove the map</button>` : ""}</div>
        ${P ? "" : `<label class="fld">Chairs without a map<input id="cfgChairs" value="${esc(c.Chairs)}"><span class="hint">Chair numbers, separated by commas</span></label><div class="actions"><button class="btn ghost sm line" data-act="saveChairs">Save chairs</button></div>`}</section>

      <section class="card set"><h3>Visit timing and flags</h3>
        <div class="two"><label class="fld">Visit length, min<input id="cf_VisitMin" inputmode="numeric" value="${esc(c.VisitMin)}"></label><label class="fld">Provider check within, min<input id="cf_ProviderCheckMin" inputmode="numeric" value="${esc(c.ProviderCheckMin)}"></label>
        <label class="fld">Waiting too long after, min<input id="cf_LateMin" inputmode="numeric" value="${esc(c.LateMin)}"></label><label class="fld">“Out soon” within, min<input id="cf_SoonMin" inputmode="numeric" value="${esc(c.SoonMin)}"></label>
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
      return `<button data-act="mPick" data-b="${esc(c.id)}" aria-pressed="${cur === c.id}"${busy ? " disabled" : ""}>${esc(chairTxt(c.id))}<small>${busy ? esc(C.code(c.v)) : "open"}</small></button>`; }).join("")}${allowNone ? `<button data-act="mPick" data-b="" aria-pressed="${!cur}">Not yet<small>lobby</small></button>` : ""}</div>`;
  };
  /* the patient's own frequency first; with none on file, a blank choice is shown first so nothing is written by default */
  const freqOpts = (p, cur) => (p.freq ? "" : `<option value=""${cur ? "" : " selected"}>Choose the protocol</option>`)
    + [...new Set([p.freq, ...S.freqs].filter(Boolean))].map(f => `<option${f === cur ? " selected" : ""}>${esc(f)}</option>`).join("");
  /* check-out asks about today's payment when the visit owes something (copay + infusion) and it hasn't been answered */
  const payNeeded = v => !!(v.p && S.payOn && C.dueOf(v.p).total > 0 && !C.payOf(S.pays, v.key));
  function payPickHTML(v, m) {
    if (!v.p || !S.payOn) return "";
    const due = C.dueOf(v.p), had = C.payOf(S.pays, v.key), t = ptOf(v.p.mrn), L = C.ledger(v.p.mrn, S.pays);
    if (had) return `<p class="note">Payment today: ${C.money(had.amount)} ${had.kind === "collected" ? "collected" : "not collected"}${had.by ? ` (${esc(had.by)})` : ""}.</p>`;
    if (due.total <= 0) return L.balance > 0 ? `<p class="note warn">Nothing due today, but there’s a balance of ${C.money(L.balance)}.</p>` : "";
    const pk = (k, l, s) => `<button data-act="mPay" data-v="${k}" aria-pressed="${m.pay === k}">${l}<small>${s}</small></button>`;
    return `<div class="fld">Payment today: ${C.money(due.total)}${due.copay && due.infusion ? ` <span class="muted">(${C.money(due.copay)} copay, ${C.money(due.infusion)} infusion)</span>` : ""}
      <div class="pick">${pk("collected", "Collected", C.money(due.total))}${pk("missed", "Not collected", "adds to the balance")}</div></div>
      ${L.balance > 0 ? `<p class="note warn">Balance from earlier visits: ${C.money(L.balance)}.</p>` : ""}
      ${m.pay === "missed" ? `<p class="note">${t?.card && t?.okCollect ? "Luma card on file and okay to collect: charge it in Luma, then record the payment on the patient’s card." : t?.card ? "Luma card on file, but not okay to collect yet. Ask the patient." : "No Luma card on file. Ask for one, or send a balance reminder."}</p>` : ""}`;
  }
  /* the edit form's fields from the row: label, current value as the input shows it, type, and suggestions from the column */
  const SKIP_EDIT = { mpt: ["mrn", "patientname", ...PREFS_KEYS()], consults: ["patientmrn", "patient"] };
  function PREFS_KEYS() { return ["am", "pm", "am2", "pm3", "am4", "pm5", "am6", "pm7", "am8", "pm9"]; }
  function editFields(m) {
    const R = m.row, src = m.kind === "consults" ? S.raw.consults : S.raw.mpt, hdr = (src && src[0] || []).map(C.norm);
    return R.head.map((h, i) => {
      const k = C.norm(h); if (!k || SKIP_EDIT[m.kind].includes(k)) return null;
      const v = R.values[i], f = R.formulas[i], fmt = String(R.formats[i] || ""), dateLike = /[dmy]{1,4}[\/\-.][dmy]/i.test(fmt) && !/h/i.test(fmt) || (/date|exp\b|start|updated|thru|dob|birth/i.test(h) && (typeof v === "number" || v === ""));
      const type = dateLike ? "date" : typeof v === "number" ? "num" : "text";
      const cur = type === "date" ? (C.toSerial(v) != null ? C.isoOf(C.toSerial(v)) : "") : C.txt(v);
      const col = hdr.indexOf(k), seen = col < 0 || !src ? [] : [...new Set(src.slice(1).map(r => C.txt(r[col])).filter(x => x && !/^\d+(\.\d+)?$/.test(x)))];
      const opts = k === "frequency" ? S.freqs : seen.length <= 40 && type === "text" ? seen.sort() : [];
      return { i, h, label: h.replace(/\s+/g, " ").trim(), type, cur, formula: !!f, shown: type === "date" && C.toSerial(v) != null ? C.fmtD(C.toSerial(v)) : C.txt(v), opts, long: /notes?\b/i.test(h) };
    }).filter(Boolean);
  }
  const MODALS = {
    checkin() {
      const v = visitBy(S.m.key), m = S.m, c = C.code(v), aOk = teamsReady("arrival"), pOk = teamsReady("protocol");
      const arrTxt = C.arrivalText(v, m.chair, S.tpl.arrival), askTxt = C.askText(v.p, v.ask, S.tpl, v.p?.freq);
      return `<h2 id="mh">Check in ${esc(c)}</h2><p class="sub muted">${C.fmtT(v.a.time)} with ${esc(C.provName(v.a.provider) || "the provider")}</p>
        <div class="fld">Chair${chairPick(m.chair, true)}</div>
        <label class="opt"><input type="checkbox" id="mRn"${m.rn && aOk ? " checked" : ""}${aOk ? "" : " disabled"}> Tell ${esc(routeTo("arrival"))} in Teams<small>${aOk ? `“${esc(arrTxt)}”` : "Arrivals aren’t set up yet (Settings › Teams)."}</small></label>
        ${v.p && v.ask.any ? `<label class="opt"><input type="checkbox" id="mAsk"${m.ask && pOk ? " checked" : ""}${pOk ? "" : " disabled"}> Ask ${esc(routeTo("protocol"))}<small>${pOk ? `“${esc(askTxt)}”` : "The group chat isn’t set up yet (Settings › Teams)."}</small></label>` : ""}
        <div class="foot"><button class="btn mute" data-act="close">Cancel</button><button class="btn primary inl" data-act="doCheckin"${S.busy ? " disabled" : ""}>Check in ${esc(c)}${m.chair ? ` to ${esc(chairTxt(m.chair))}` : ""}</button></div>`;
    },
    seat() {
      const v = visitBy(S.m.key), c = C.code(v), aOk = teamsReady("arrival"), moving = v.ph === "chair";
      return `<h2 id="mh">${moving ? `Move ${esc(c)} to another chair` : `Seat ${esc(c)}`}</h2><p class="sub muted">${moving ? `Now in ${esc(chairTxt(v.st.chair))}` : `Checked in at ${C.fmtT(v.st.arr)}`}</p>
        <div class="fld">Chair${chairPick(S.m.chair, false)}</div>
        <label class="opt"><input type="checkbox" id="mRn"${S.m.rn && aOk ? " checked" : ""}${aOk ? "" : " disabled"}> Tell ${esc(routeTo("arrival"))} which chair<small>${aOk ? `“${esc(C.fill(S.tpl.seated, { patient: c, chair: S.m.chair || "…" }))}”` : "Arrivals aren’t set up yet (Settings › Teams)."}</small></label>
        <div class="foot"><button class="btn mute" data-act="close">Cancel</button><button class="btn primary inl" data-act="doSeat"${S.m.chair && !S.busy ? "" : " disabled"}>${S.m.chair ? `Seat ${esc(c)} in ${esc(chairTxt(S.m.chair))}` : "Choose a chair"}</button></div>`;
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
      return `<h2 id="mh">Check out ${esc(c)}</h2><p class="sub muted">${v.st.seat != null ? `In ${esc(chairTxt(v.st.chair))} since ${C.fmtT(v.st.seat)}` : `Checked in at ${C.fmtT(v.st.arr)}`}</p>
        <div class="fld">Next visit${v.ask.next ? ` <span class="muted">(on the list: ${C.fmtD(v.ask.next.date)})</span>` : ""}<div class="pick">${bk("yes", "Booked", "in NextGen")}${bk("later", "Book later", "keep reminding")}${bk("no", "Not needed", "PRN or ending")}</div></div>
        ${p ? `<label class="opt"><input type="checkbox" id="mProto"${m.proto ? " checked" : ""}> The provider confirmed the protocol<small>Writes Protocol Updated ${C.fmtD(todayS())}${m.freq && m.freq !== p.freq ? ` and Frequency “${esc(m.freq)}”` : ""} on the Master Patient Tracker</small></label>
        <label class="fld">Protocol<select id="mFreq">${freqOpts(p, m.freq)}</select></label>` : ""}
        ${payPickHTML(v, m)}
        <div class="foot"><button class="btn mute" data-act="close">Cancel</button><button class="btn primary inl" data-act="doCheckout"${m.booked && (!payNeeded(v) || m.pay) && !S.busy ? "" : " disabled"}>${!m.booked ? "Choose the next visit" : payNeeded(v) && !m.pay ? "Say whether payment was collected" : `Check out ${esc(c)}`}</button></div>`;
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
      return `<h2 id="mh">${esc(m.title)}</h2><p class="sub muted">${m.kind === "consults" ? "The Consults tab" : "The Master Patient Tracker"}, saved straight to the row. Formula cells are left alone.${m.kind === "mpt" ? " Morning and afternoon days are under Availability." : ""}</p>
        <div class="egrid">${F.map(fld).join("")}</div>
        <div class="foot"><button class="btn mute" data-act="close">Cancel</button><button class="btn primary inl" data-act="doEdit"${S.busy ? " disabled" : ""}>Save changes</button></div>`;
    },
    /* a payment on account (comes off the oldest balance) or a balance reminder */
    pay() {
      const p = S.PM.get(S.m.mrn), L = C.ledger(p.mrn, S.pays), t = ptOf(p.mrn), today = C.isoOf(todayS());
      return `<h2 id="mh">${esc(p.code)}: payments</h2><p class="sub muted">${L.balance > 0 ? `Balance ${C.money(L.balance)}` : "No balance"}${t?.card ? `. Luma card on file${t.okCollect ? ", okay to collect" : ", not okay to collect yet"}` : ". No card on file in Luma"}.</p>
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
      return `<h2 id="mh">${esc(p.code)}: availability</h2><p class="sub muted">From the patient availability form. Preferred and Mostly available go to the tracker’s AM/PM columns and the Index planner.</p>
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
      return `<h2 id="mh">${email ? `Start email for ${esc(p.code)}` : `Print for ${esc(p.code)}`}</h2>
        <p class="sub muted">${p.onTracker ? "The patient’s name comes from the Master Patient Tracker. " : ""}It’s only put on the forms, never saved or posted.${p.onTracker && !p.name ? " <b>This row has no Patient Name, so the forms will have a blank name.</b>" : ""}</p>
        ${p.onTracker ? "" : `<label class="fld">Patient’s name<input id="fmName" value="${esc(m.name || "")}" autocomplete="off" spellcheck="false"><span class="hint">Not on the Master Patient Tracker yet; type it for the forms (not saved). The schedule only needs the MRN.</span></label>`}
        <div class="two"><label class="fld">Date of birth<input id="fmDob" type="date" value="${esc(m.dob)}"><span class="hint">${p.dob != null ? "From the tracker" : "Not on the tracker; type it for the forms (not saved)"}</span></label>
          <label class="fld">Clinic location<input id="fmLoc" value="${esc(m.loc)}" autocomplete="off"></label></div>
        <label class="fld">Clinician<select id="fmProv">${provs.map(x => `<option${x === m.prov ? " selected" : ""}>${esc(x)}</option>`).join("")}</select></label>
        ${email ? `<p class="note">The email gets the start template (Settings › Messages) with the schedule in the body, and the schedule and the filled-in IVK consent attached. The BIDF is attached in the next window.</p>`
          : `<div class="fld">Print${part("sched", "Schedule")}${part("consent", "IV ketamine consent", "Name, date of birth, clinic location and clinician filled in")}${part("avail", "Availability form", "Name and date of birth filled in")}</div>`}
        ${email || m.parts.sched ? schedRangeHTML(p, m) : ""}
        <div class="foot"><button class="btn mute" data-act="close">Cancel</button><button class="btn primary inl" data-act="doForms">${email ? "Make the email" : "Open to print"}</button></div>`;
    },
    proto() {
      const p = S.PM.get(S.m.mrn);
      return `<h2 id="mh">Confirm ${esc(p.code)}’s protocol</h2><p class="sub muted">${p.protoAt != null ? `Last confirmed ${C.fmtD(p.protoAt)}` : "Not confirmed on the tracker yet"}</p>
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
    confirm() {
      return `<h2 id="mh">${esc(S.m.title)}</h2><p class="sub">${esc(S.m.body)}</p><div class="foot"><button class="btn mute" data-act="close">Keep it</button><button class="btn primary inl" data-act="doConfirm">${esc(S.m.yes)}</button></div>`;
    },
    upload() {
      const u = S.m, sh = u.sh, ds = u.rows.map(r => r[2]), sts = {};
      u.rows.forEach(r => { const k = r[6] || "blank"; sts[k] = (sts[k] || 0) + 1; });
      const why = [sh.few ? `only ${plural(u.rows.length, "appointment")}` : "", sh.half ? `fewer than half of the ${sh.had} on file` : "", sh.short ? `it ends ${C.fmtD(sh.nMax)}; the list on file runs to ${C.fmtD(sh.cMax)}` : ""].filter(Boolean);
      return `<h2 id="mh">Replace the appointment list?</h2><p class="sub muted">${esc(u.file)}</p>
        <div class="preview">${plural(u.rows.length, "appointment")}, ${C.fmtD(Math.min(...ds))} to ${C.fmtD(Math.max(...ds))}. ${Object.entries(sts).map(([k, n]) => `${esc(k)} ${n}`).join(", ")}</div>
        ${u.other && u.other.length ? `<p class="note">${plural(u.other.length, "other visit")} (telehealth, TMS and the rest) are kept with no patient on them, so flex post checks those visits use aren’t offered for index.</p>` : `<p class="note">No other visits in this report, so every flex post check shows as free. Run the wider report to include the providers’ telehealth and TMS visits.</p>`}
        <p class="note">The ${sh.had} rows on the Appt List tab are replaced (MRN to Location). The Name and Lookup Key formulas stay. Then the Index Planner tab’s grid is rewritten with every IVK visit and hold, as values; anything typed in that grid is replaced.</p>
        ${sh.flag ? `<div class="msg warn" style="margin-top:12px"><b>This list is much smaller than the one on file:</b> ${why.join("; ")}. The date range in NextGen may have been narrowed.<label class="opt"><input type="checkbox" id="mShrink"${u.ok ? " checked" : ""}> The smaller list is right; replace anyway</label></div>` : ""}
        <div class="foot"><button class="btn mute" data-act="close">Cancel</button><button class="btn primary inl" data-act="doUpload"${sh.flag && !u.ok || S.busy ? " disabled" : ""}>Replace list</button></div>`;
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
    for (const k of ["sched", "consent", "avail"]) if ($(`#fp_${k}`)) m.parts[k] = $(`#fp_${k}`).checked; };
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
    mPay: e => { S.m.pay = e.dataset.v; S.m.proto = $("#mProto")?.checked ?? S.m.proto; S.m.freq = $("#mFreq")?.value ?? S.m.freq; render(); },
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
      await write("Saving…", () => Store.addPays([payRow(p, { Date: C.isoOf(d), Key: k, Kind: kind, Amount: C.dueOf(p).total, Note: "confirmed later" })]), `${dshort(d)}: ${kind === "collected" ? "collected" : "not collected, added to the balance"}.`); },
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
        return Store.saveRow("mpt", p.mrn, C.prefCells(m.grid));
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
      closeModal(); openForms(d);
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
  /* waitlist: an appeal moves them behind the next ready patient; turning it off leaves them where they are */
  async function wAppeal(t) {
    const w = S.WL.list.find(x => x.mrn === t.dataset.m); if (!w) return;
    await write("Saving…", async () => {
      await Store.savePt(w.mrn, ptPatch({ code: w.code }, { Appeal: t.checked ? "TRUE" : "", WaitingOn: "auth" }));
      if (t.checked) { const L = S.WL.list, order = C.waitMove(L, w.mrn, "appeal"); await Store.saveRanks(order.map(x => ({ mrn: x, code: L.find(y => y.mrn === x).code })), S.name); }
    }, t.checked ? `${w.code} moved behind the next ready patient.` : null);
  }
  async function ptFlag(t) {
    const p = S.PM.get(t.dataset.m); if (!p) return; const k = t.dataset.pt, on = t.checked;
    const words = { CardOnFile: ["Luma card on file", "No Luma card on file"], OkToCollect: ["Okay to collect", "Not okay to collect"], Oakland: ["Oakland patient", "Not an Oakland patient"] }[k];
    await write("Saving…", () => Store.savePt(p.mrn, ptPatch(p, { [k]: on ? "TRUE" : "FALSE" })), `${p.code}: ${words[on ? 0 : 1].toLowerCase()}.`);
  }
  const ACT = {
    nav: e => { S.view = e.dataset.v; S.bSel = null; render(); },
    async reload() { S.reading = true; render(); try { await refresh(true); toast("Read the workbook again: up to date."); } finally { S.reading = false; render(); } },
    theme: () => { S.theme = S.theme === "dark" ? "light" : "dark"; LS.set("ivkTheme", S.theme); document.documentElement.dataset.theme = S.theme; render(); },
    full: () => openFull(),
    sel: e => { S.sel = e.dataset.k; S.bSel = null; render(); const v = visitBy(S.sel); if (v) follow("mpt", v.a.mrn); },
    bSel: e => { const F = floor(), c = F.chairs.find(x => x.id === e.dataset.b); if (c && c.v) { S.sel = c.v.key; S.bSel = null; follow("mpt", c.v.a.mrn); } else S.bSel = e.dataset.b; render(); },
    pSel: e => { S.pSel = e.dataset.m; render(); follow("mpt", S.pSel); },
    chip: e => { S.chip = e.dataset.v; render(); },
    wk: e => { const d = +e.dataset.v; S.wk = d === 0 ? 0 : S.wk + d; render(); },
    schedPick: e => { const m = e.dataset.m, d = +e.dataset.d, tv = S.list.find(v => v.a.mrn === m); if (d === todayS() && tv) { S.sel = tv.key; S.view = "today"; } else { S.pSel = m; S.chip = "all"; S.view = "patients"; } render(); follow("mpt", m); },
    toToday: e => { S.sel = e.dataset.k; S.bSel = null; S.view = "today"; render(); },
    openPt: e => { S.pSel = e.dataset.m; S.chip = "all"; S.view = "patients"; render(); follow("mpt", S.pSel); },
    setup: async () => { const r = await write("Setting up…", () => Store.setup()); if (r && r.blocked?.length) toast(`Couldn’t add ${r.blocked.join(", ")}: cells on the ${X.DATA_SHEET} tab already hold something.`, true); else if (r) toast("Workbook set up. Check-ins and Teams settings now save here."); },
    checkin: e => { if (needName()) return; const v = visitBy(e.dataset.k); if (!v) return; openModal("checkin", { key: v.key, chair: freeChair(), rn: true, ask: true }); },
    seat: e => { const v = visitBy(e.dataset.k); if (!v) return; openModal("seat", { key: v.key, chair: v.ph === "chair" ? "" : freeChair(), rn: true }); },
    seatHere: e => { const v = visitBy(e.dataset.k); if (!v) return; openModal("seat", { key: v.key, chair: e.dataset.b, rn: true }); },
    mPick: e => { S.m.chair = e.dataset.b; S.m.rn = $("#mRn")?.checked ?? S.m.rn; S.m.ask = $("#mAsk")?.checked ?? S.m.ask; render(); },
    mBook: e => { S.m.booked = e.dataset.v; S.m.proto = $("#mProto")?.checked ?? S.m.proto; S.m.freq = $("#mFreq")?.value ?? S.m.freq; render(); },
    close: () => closeModal(),
    scrim: (e, ev) => { if (ev.target === e) closeModal(); },
    async doCheckin() {
      const v = visitBy(S.m.key), chair = S.m.chair, rn = $("#mRn")?.checked, ask = $("#mAsk")?.checked, t = hm();
      if (!v) return closeModal();
      const ok = await write("Saving check-in…", () => Store.saveVisit(v.key, visitBase(v), Object.assign({ Arrived: t, Chair: chair, Seated: chair ? t : "", Out: "", Status: "" }, stamp()), "Arrived"));
      if (!ok) return;
      closeModal(); S.sel = v.key; S.bSel = null;
      const msgs = [`${C.code(v)} checked in${chair ? ` to ${chairTxt(chair)}` : ""}.`];
      if (rn) { const r = await sendArrival(v, chair); msgs.push(teamsWord(r, routeTo("arrival"))); }
      if (ask && v.p) { const r = await sendAsk(v.p, C.askText(v.p, v.ask, S.tpl, v.p.freq), v.p.freq, v.key); msgs.push(r.ok ? `Asked ${routeTo("protocol")}.` : r.reason); }
      toast(msgs.join(" "), msgs.some(x => /answered|couldn’t|No Teams|off/i.test(x)));
    },
    async doSeat() {
      const v = visitBy(S.m.key), chair = S.m.chair, rn = $("#mRn")?.checked, t = hm(); if (!v || !chair) return;
      const ok = await write("Seating…", () => Store.saveVisit(v.key, visitBase(v), Object.assign({ Chair: chair, Seated: v.ph === "chair" ? (C.fmtT(v.st.seat) || t) : t, Arrived: v.st.arr != null ? C.fmtT(v.st.arr) : t }, stamp())));
      if (!ok) return; closeModal(); S.sel = v.key; S.bSel = null;
      let msg = `${C.code(v)} is in ${chairTxt(chair)}.`;
      if (rn) { const r = await sendArrival(v, chair, true); msg += " " + teamsWord(r, routeTo("arrival")); }
      toast(msg);
    },
    async remind(e) {
      if (needName()) return; const v = visitBy(e.dataset.k); if (!v) return;
      const r = await sendArrival(v, v.st.chair || "", v.ph === "chair"); toast(teamsWord(r, routeTo("arrival")), !r.ok); render();
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
      openModal("checkout", { key: v.key, booked: v.ask.book ? "" : "yes", proto: false, freq: v.p?.freq || "" });
    },
    async doCheckout() {
      const v = visitBy(S.m.key); if (!v || !S.m.booked) return;
      const proto = !!$("#mProto")?.checked && !!v.p, freq = $("#mFreq")?.value || "", t = hm();
      if (proto && !freq) return toast("Pick the protocol the provider confirmed, or untick the box.", true);
      const ok = await write("Checking out…", () => Store.saveVisit(v.key, visitBase(v), Object.assign({ Out: t, Booked: S.m.booked, Protocol: proto ? freq : v.st.proto, Arrived: v.st.arr != null ? C.fmtT(v.st.arr) : t }, stamp())),
        proto ? null : `${C.code(v)} checked out.`);
      if (!ok) return;
      const pay = S.m.pay; closeModal();
      if (pay && payNeeded(v)) await write("Saving the payment…", () => Store.addPays([payRow(v.p, { Date: C.isoOf(v.a.date), Key: v.key, Kind: pay, Amount: C.dueOf(v.p).total })]))
        || toast(`${C.code(v)} is checked out, but the payment answer didn’t save. Record it on the patient’s card.`, true);
      if (proto) await write("Saving the protocol…", () => Store.confirmProtocol(v.a.mrn, freq, todayS()), `${C.code(v)} checked out; protocol ${freq} saved on the Master Patient Tracker.`)
        || toast(`${C.code(v)} is checked out, but the protocol didn’t save to the Master Patient Tracker. Use Patients › Confirm protocol.`, true);
    },
    async markBooked(e) { const v = visitBy(e.dataset.k); if (!v) return; await write("Saving…", () => Store.saveVisit(v.key, visitBase(v), Object.assign({ Booked: "yes" }, stamp())), `${C.code(v)}’s next visit marked booked.`); },
    gone: e => { const v = visitBy(e.dataset.k); openModal("confirm", { title: `Take ${C.code(v)} off today’s list?`, body: "NextGen isn’t changed. Cancel or reschedule the visit there too.", yes: "Take off today’s list", run: () => Store.saveVisit(v.key, visitBase(v), Object.assign({ Status: "cancelled" }, stamp())) }); },
    undoGone: async e => { const v = visitBy(e.dataset.k); await write("Saving…", () => Store.saveVisit(v.key, visitBase(v), Object.assign({ Status: "" }, stamp()))); },
    undoIn: e => { const v = visitBy(e.dataset.k); openModal("confirm", { title: `Undo ${C.code(v)}’s check-in?`, body: "They go back to Still to come. A Teams post already sent stays in the chat.", yes: "Undo check-in", run: () => Store.saveVisit(v.key, visitBase(v), Object.assign({ Arrived: "", Chair: "", Seated: "" }, stamp())) }); },
    undoSeat: async e => { const v = visitBy(e.dataset.k); await write("Saving…", () => Store.saveVisit(v.key, visitBase(v), Object.assign({ Chair: "", Seated: "" }, stamp())), `${C.code(v)} is back to waiting.`); },
    undoOut: async e => { const v = visitBy(e.dataset.k); await write("Saving…", () => Store.saveVisit(v.key, visitBase(v), Object.assign({ Out: "", Booked: "" }, stamp())), `${C.code(v)} is back in ${v.st.chair ? chairTxt(v.st.chair) : "the lobby"}.`); },
    async doConfirm() { const run = S.m.run; closeModal(); await write("Saving…", run); },
    proto: e => { const p = S.PM.get(e.dataset.m); if (p) openModal("proto", { mrn: p.mrn, freq: p.freq || "" }); },
    async doProto() { const p = S.PM.get(S.m.mrn), f = $("#mFreq").value; if (!f) return toast("Pick the protocol the provider confirmed.", true); const ok = await write("Saving protocol…", () => Store.confirmProtocol(p.mrn, f, todayS()), `${p.code}’s protocol saved: ${f}, confirmed ${C.fmtD(todayS())}.`); if (ok) closeModal(); },
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
    async saveChairs() { await write("Saving chairs…", () => Store.saveKeyed("tblIvkConfig", { Chairs: $("#cfgChairs").value.trim() || C.DEFAULTS.Chairs }), "Chairs saved."); },
    async saveCfg() {
      const o = {}; for (const k of ["VisitMin", "ProviderCheckMin", "LateMin", "SoonMin", "AheadWeeks", "ProtocolDays", "AuthWarnDays", "EncountersWarn"]) { const n = C.num($(`#cf_${k}`).value); if (n == null || n < 0) return toast(`${k} needs a number.`, true); o[k] = n; }
      o.Clinic = $("#cf_Clinic").value.trim() || C.DEFAULTS.Clinic;
      await write("Saving…", () => Store.saveKeyed("tblIvkConfig", o), "Timing saved.");
    },
    mapDown: () => { const P = S.raw.plan, a = document.createElement("a"); a.href = URL.createObjectURL(new Blob([JSON.stringify(P, null, 1)], { type: "application/json" })); a.download = `Chair map - ${S.cfg.Clinic} IVK.json`; a.click(); },
    mapRemove: () => openModal("confirm", { title: "Remove the chair map?", body: "Today shows chair tiles instead. You can load the file again any time.", yes: "Remove the map", run: () => Store.setDoc(MAP_KEY, null) }),
    async doUpload() {
      const u = S.m; if (u.sh.flag && !$("#mShrink")?.checked) return; closeModal();
      const wasPending = new Set(unconfirmed().map(h => h.id));   // before the new list is read
      const ok = await write("Writing the appointment list…", () => Store.replaceAppts(u.rows, C.serialOf(new Date()) + nowMin() / 1440)); if (!ok) return;
      if (S.raw.tables?.own?.tblIvkOther && !(await write("Saving the providers’ other visits…", () => Store.replaceOther(u.other))))
        toast("The appointment list is updated, but the providers’ other visits didn’t save, so flex post checks may show as free. Upload again.", true);
      await write("Filling the Index Planner tab…", () => Store.fillPlanner(), r => `Appointment list updated: ${plural(u.rows.length, "appointment")}.${r && r.skipped ? ` Index Planner tab not filled (${r.skipped}).` : r ? ` Index Planner tab filled, ${C.fmtD(r.from)} to ${C.fmtD(r.to)}.` : ""}`)
        || toast("The appointment list is updated, but the Index Planner tab wasn’t filled. Upload again, or place or release a hold, to retry.", true);
      /* the booked index visits this report confirms, and any it doesn't have */
      if (wasPending.size) {
        const H = (S.holds || []).filter(h => wasPending.has(h.id)).map(h => C.holdConfirm(h, S.A, S.raw.apptUpdated)), ok = H.filter(c => c === "onlist").length, miss = H.filter(c => c === "missing").length;
        toast(miss ? `${plural(ok, "booked index visit")} confirmed; ${plural(miss, "visit")} marked booked ${miss === 1 ? "isn’t" : "aren’t"} on this report. Check them in NextGen (Index planner).` : `${ok === 1 ? "The booked index visit is" : `All ${ok} booked index visits are`} confirmed on the NextGen report.`, !!miss);
      }
    },
    /* index planner */
    ipView: e => { S.ipView = e.dataset.v; S.ipAlt = null; render(); },
    ipPick: e => { S.ipSel = e.dataset.m; S.ipAlt = null; S.ipView = "plan"; S.view = "index"; render(); follow("mpt", S.ipSel); },
    ipAlt: e => { const n = +e.dataset.n; S.ipAlt = S.ipAlt === n ? null : n; render(); },
    ipUse: e => {
      const o = ipList().find(x => x.mrn === S.ipSel); if (!o) return;
      const M = ipModel(o), n = +e.dataset.n, s = C.altSlots(M.res, n, +S.cfg.MinGapDays)[+e.dataset.i]; if (!s) return;
      /* the visits before this one stay where they are; the ones after are planned again around it */
      const pins = {}; M.res.items.slice(0, n).forEach((x, k) => { if (!x.fixed) pins[k] = asPin(x); });
      pins[n] = asPin(s); S.ipPins[o.mrn] = pins; S.ipAlt = null; render();
    },
    /* an open index start pressed on the Access week: treatment 1, or the treatment being changed */
    ipTile: e => {
      const o = ipList().find(x => x.mrn === S.ipSel); if (!o) return toast("Pick a patient to plan first.", true);
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
    ipReset: e => { delete S.ipPins[e.dataset.m]; S.ipAlt = null; render(); },
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
  function freeChair() { const F = floor(); return (F.chairs.find(c => !c.v) || {}).id || ""; }

  /* ---------- file inputs and fields ---------- */
  let xlsxP = null;
  const loadXLSX = () => window.XLSX ? Promise.resolve(true) : xlsxP || (xlsxP = new Promise(res => { const s = document.createElement("script"); s.src = "assets/xlsx.full.min.js"; s.onload = () => res(!!window.XLSX); s.onerror = () => { xlsxP = null; s.remove(); res(false); }; document.head.appendChild(s); }));
  async function readAppt(f) {
    if (!(await loadXLSX())) return toast("The spreadsheet reader didn’t load. Check the connection and reopen the add-in.", true);
    let rows;
    try { const wb = XLSX.read(await f.arrayBuffer(), { type: "array" }); rows = C.parseReport(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: "" })); }
    catch (e) { console.error(e); return toast("That file couldn’t be read. Upload the NextGen export unchanged.", true); }
    if (!rows || !rows.length) return toast("This doesn’t look like a NextGen appointment listing (no MRN, Appointment Date and Status columns).", true);
    /* a wider report (the providers' whole schedule) is split: IVK visits for the Appt List tab, the others kept with no patient on them */
    const ivk = rows.filter(r => C.isIvkVisit(r[5])), other = rows.filter(r => !C.isIvkVisit(r[5])).map(r => [C.isoOf(r[2]), r[3] === "" ? "" : C.fmtT(Math.round(r[3] * 1440)), r[4], r[5], r[6]]);
    if (!ivk.length) return toast("This report has no IV ketamine visits. Check the NextGen report's filters.", true);
    const sh = C.shrink(ivk, S.A.map(a => a.date)); sh.had = S.A.length;
    openModal("upload", { file: f.name, rows: ivk, other, sh, ok: false });
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
    else if (t.id === "darkOn") { S.theme = t.checked ? "dark" : "light"; LS.set("ivkTheme", S.theme); document.documentElement.dataset.theme = S.theme; }
    else if (t.id === "mFreq" && S.modal === "ask") { S.m.freq = t.value; S.m.text = C.askText(S.PM.get(S.m.mrn), S.m.ask, S.tpl, t.value); $("#mText").value = S.m.text; }
    else if (t.id === "mFreq" && (S.modal === "checkout" || S.modal === "proto")) { S.m.freq = t.value; S.m.proto = $("#mProto")?.checked ?? S.m.proto; render(); }
    else if (t.id === "mProto") S.m.proto = t.checked;
    else if (t.id === "mShrink") { S.m.ok = t.checked; render(); }
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
    else if (S.modal === "forms" && ["fmFrom", "fmTo", "fp_sched", "fp_consent", "fp_avail"].includes(t.id)) { keepForms(); render(); }
    else if (S.modal === "pto" && (t.id === "ptCov" || t.id === "ptProv" || t.id === "ptFrom" || t.id === "ptTo")) { const was = S.m.cov; keepPto(); if (t.id === "ptCov") { S.m.cov = t.value; if (t.value !== was) S.m.useTpl = ""; } render(); }
    else if (t.id === "followOn") { LS.set("ivkFollow", t.checked ? "1" : "0"); followed = ""; }
    else if (["ipStart", "ipPer", "ipWeeks"].includes(t.id)) {
      const o = ipList().find(x => x.mrn === S.ipSel); if (!o) return; delete S.ipPins[o.mrn];
      if (t.id === "ipStart") { const d = C.toSerial(t.value); if (d != null) savePlanOf(o, { Start: C.isoOf(d) }); }
      else savePlanOf(o, t.id === "ipPer" ? { PerWeek: t.value } : { Weeks: t.value });
    }
    else if (t.id === "mRn") S.m.rn = t.checked;
    else if (t.id === "mAsk") S.m.ask = t.checked;
  });
  document.addEventListener("input", e => { if (e.target.id === "ptSearch") { S.q = e.target.value; const pos = e.target.selectionStart; render(); const s = $("#ptSearch"); if (s) { s.focus(); s.setSelectionRange(pos, pos); } } });
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
  }
  window.addEventListener("error", e => console.error(e.message));
  if (window.Office && Office.onReady) Office.onReady(start); else landing();
  window.IvkApp = { S, render, derive, refresh, ACT, floor, ipList, ipModel };   // for the test harness
})();
