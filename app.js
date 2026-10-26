/* app.js: the IVK Tracker screens. Used at the Castro front desk; the RN never opens it and gets Teams posts instead.
   Today:     chair map (a Bay Map Builder file, else chair tiles), who's waiting and still to come, and one visit card
   Schedule:  the week by day and time (replaces the Index Planner tab)
   Patients:  everyone on the Master Patient Tracker, sorted by what they need (auth, booking, protocol)
   Consults:  the Consults tab, approved-not-scheduled first
   Settings:  your name, the NextGen upload, Teams workflows and messages, chair map, timing, appearance, workbook setup
   The task pane reads and writes the workbook (excel.js). Full screen is an Office dialog floating over the Excel page; it holds a
   copy of what the pane read and asks the pane for every write (callPane). Teams posts go straight from either one. */
(() => {
  "use strict";
  const C = window.IvkCore, N = window.IvkNotify, X = window.IvkExcel;
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

  const S = { view: "today", raw: null, sig: "", live: false, err: "", sel: null, bSel: null, pSel: null, chip: "need", q: "", wk: 0,
    modal: null, m: {}, loading: false, name: LS.get("ivkName", ""), theme: LS.get("ivkTheme", "light") };
  document.documentElement.dataset.theme = S.theme;

  /* ---------- derived model ---------- */
  const kv = vals => { const o = {}; (vals || []).slice(1).forEach(r => { const k = C.txt(r[0]); if (k) o[k] = r[1]; }); return o; };
  function derive() {
    const r = S.raw || {}, T = todayS(), cfg = Object.assign({}, C.DEFAULTS);
    for (const [k, v] of Object.entries(kv(r.tblIvkConfig))) if (C.txt(v) !== "") cfg[k] = typeof C.DEFAULTS[k] === "number" && C.num(v) != null ? C.num(v) : C.txt(v);
    S.cfg = cfg;
    S.tpl = Object.assign({}, C.TEMPLATES); for (const [k, v] of Object.entries(kv(r.tblIvkTemplates))) if (C.txt(v)) S.tpl[k] = C.txt(v);
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
    setup: () => X.setup(), setDoc: (k, v) => X.setDoc(k, v)
  };
  const Store = {};
  for (const k of Object.keys(Local)) Store[k] = FULL ? (...a) => callPane(k, a) : async (...a) => { const r = await Local[k](...a); if (k !== "openRow") await refresh(true); return r; };
  async function write(label, fn, done) {
    S.busy = label; render();
    try { const r = await fn(); if (done) toast(typeof done === "function" ? done(r) : done); return r || true; }
    catch (e) { console.error(e); toast(`Couldn’t save to the workbook: ${e.message || e}. Try again.`, true); return null; }
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
  const PANE = Object.fromEntries(Object.keys(Local).map(k => [k, async (...a) => { const r = await Local[k](...a); if (k !== "openRow") await refresh(true); return r; }]));
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
  const NAV = [["today", "Today"], ["schedule", "Schedule"], ["patients", "Patients"], ["consults", "Consults"]];
  function badges() {
    const F = floor(), b = {};
    const t = F.late.length + F.ready.length; if (t) b.today = { n: t, bad: F.late.length > 0 };
    /* Patients counts only what blocks the next infusion; the amber needs live under the Needs attention chip */
    const pn = S.P.filter(p => C.needs(p, S.x).some(y => y.bad)).length; if (pn) b.patients = { n: pn, bad: true };
    const ca = S.CS.filter(c => c.state === "act").length; if (ca) b.consults = { n: ca };
    const set = (S.raw?.missing || []).length + (teamsReady("arrival") && teamsReady("protocol") ? 0 : 1); if (set) b.settings = { n: set };
    return b;
  }
  function render() {
    const page = $(".page"), top = page ? page.scrollTop : 0;
    if (!S.raw) { $("#app").innerHTML = `<div class="landing"><h1>IVK Tracker</h1><p class="muted">${S.err ? esc(S.err) : "Reading the workbook…"}</p></div>`; return; }
    const B = badges(), item = ([k, label]) => `<button class="item" data-act="nav" data-v="${k}"${S.view === k ? ' aria-current="page"' : ""}>${label}${B[k] ? `<span class="count${B[k].bad ? " bad" : ""}" aria-label="${B[k].n} need attention">${B[k].n}</span>` : ""}</button>`;
    const body = { today: todayHTML, schedule: scheduleHTML, patients: patientsHTML, consults: consultsHTML, settings: settingsHTML }[S.view]();
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
  const head = (title, sub, clock) => `<div class="top"><div><h1>${title}</h1>${sub ? `<p class="sub">${sub}</p>` : ""}</div>${clock ? `<div class="clock"><div class="t">${hm()}</div><div class="d">${esc(C.fmtDay(todayS()))}</div></div>` : ""}</div>`;
  function setupBanner() {
    const miss = S.raw.missing || [], noTabs = ["mpt", "appt"].filter(k => !S.raw.tables?.[k]);
    if (noTabs.length) return `<div class="msg warn"><b>This workbook has no ${noTabs.map(k => k === "mpt" ? "Master Patient Tracker" : "Appt List").join(" or ")} table.</b> Open the IVK Patient Tracker workbook, then the add-in.</div>`;
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
    const meta = [C.fmtT(v.a.time), v.a.event ? C.eventShort(v.a.event).toLowerCase() : "", C.provName(v.a.provider), p?.freq].filter(Boolean).join(", ");
    const loc = v.ph === "chair" ? chairTxt(v.st.chair) : v.ph === "waiting" ? "lobby" : "";
    const arr = sentToday("arrival", c), ask = sentToday("protocol", c);
    /* timeline while in a chair: in chair, provider check (SOP: within the first 30 min), expected out */
    let tl = "";
    if (vs.cs) {
      const s0 = v.st.seat, e0 = vs.cs.end, span = Math.max(1, e0 - s0), at = t => Math.min(100, Math.max(0, (t - s0) / span * 100));
      tl = `<div class="tl" aria-hidden="true"><div class="bar"></div><div class="fill" style="width:${at(F.now)}%"></div>
        <div class="mark done" style="left:0"></div><div class="mark${F.now >= vs.cs.check ? " done" : ""}" style="left:${at(vs.cs.check)}%"></div><div class="mark${F.now >= e0 ? " done" : ""}" style="left:100%"></div>
        <div class="now" style="left:${at(F.now)}%"></div>
        <span class="lbl start" style="left:0">${C.fmtT(s0)}</span><span class="lbl" style="left:${at(vs.cs.check)}%">Provider check ${C.fmtT(vs.cs.check)}</span><span class="lbl end" style="left:100%">Out ${C.fmtT(e0)}</span></div>`;
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
    const bt = (cls, [a, l]) => `<button class="btn ${cls}" data-act="${a}" data-k="${k}"${a === "openPt" ? ` data-m="${esc(v.a.mrn)}"` : ""}>${l}</button>`;
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
      </dl></div>`;
  }

  /* ---------- Schedule ---------- */
  function scheduleHTML() {
    const T = todayS(), start = C.weekStart(T) + 7 * S.wk, W = C.week(S.A, start, S.PM), F = floor();
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
    const cells = W.rows.map(t => `<div class="tm">${C.fmtT(t)}</div>${W.days.map(d => `<div class="sc${d === T ? " td" : ""}">${W.cell(d, t).map(blk).join("")}</div>`).join("")}`).join("");
    const label = `${C.fmtD(start)} to ${C.fmtD(start + 6)}`;
    return `${head("Schedule", `Week of ${label}: ${plural(W.count, "IVK visit")}${W.count ? `, ${W.days.map((d, i) => W.byDay[i] ? `${["Mon", "Tue", "Wed", "Thu", "Fri"][i]} ${W.byDay[i]}` : "").filter(Boolean).join(", ")}` : ""}`, true)}
      <section class="card sched"><div class="wkbar"><span class="lbl">${label}</span>
        <button class="btn ghost sm" data-act="wk" data-v="-1">Previous week</button><button class="btn ghost sm" data-act="wk" data-v="0"${S.wk === 0 ? " disabled" : ""}>This week</button><button class="btn ghost sm" data-act="wk" data-v="1">Next week</button>
        <span class="muted small" style="margin-left:auto">“last booked” = no visit booked after this one</span></div>
      <div class="sgrid"><div></div>${W.days.map(d => `<div class="ch${d === T ? " today" : ""}">${["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"][C.weekday(d) - 1]}<small>${C.fmtD(d)}</small></div>`).join("")}${cells}</div></section>`;
  }

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
  function patientCardHTML(p, n) {
    const top = C.loud(n)[0], tv = S.list.find(v => v.a.mrn === p.mrn);
    const meta = [p.status || "No status", p.loc, p.start != null ? `started ${C.fmtD(p.start)}` : ""].filter(Boolean).join(", ");
    const pri = top?.act ? top.act : ["open", "Open in the tracker"];
    const prefs = `<div class="prefs">${p.prefs.map(d => `<b>${d.d}</b>`).join("")}${p.prefs.map(d => `<span class="${d.am ? "on" : ""}">AM</span>`).join("")}${p.prefs.map(d => `<span class="${d.pm ? "on" : ""}">PM</span>`).join("")}</div>`;
    return `<aside class="card detail" aria-label="${esc(p.code)}"><div class="dh"><div><div class="name">${esc(p.code)}</div><div class="meta">${esc(meta)}</div></div></div>
      <p class="lead ${top ? (top.bad ? "late" : "soon") : ""}">${esc(top ? top.t : tv ? `Here today at ${C.fmtT(tv.a.time)}` : "Nothing needed right now")}</p>
      ${C.loud(n).length > 1 ? `<ul class="steps">${C.loud(n).slice(1).map(y => `<li class="step ${y.bad ? "late" : "soon"}"><span class="k"></span><span><span class="t">${esc(y.t)}</span></span><span class="w"></span></li>`).join("")}</ul>` : ""}
      <button class="btn primary" data-act="${pri[1]}" data-m="${esc(p.mrn)}">${esc(pri[0])}</button>
      <div class="actions">${pri[1] !== "open" ? `<button class="btn ghost" data-act="open" data-m="${esc(p.mrn)}">Open in the tracker</button>` : ""}${pri[1] !== "proto" ? `<button class="btn ghost" data-act="proto" data-m="${esc(p.mrn)}">Confirm protocol</button>` : ""}${pri[1] !== "ask" && C.isActive(p) ? `<button class="btn ghost" data-act="ask" data-m="${esc(p.mrn)}">Ask the group chat</button>` : ""}${tv ? `<button class="btn ghost" data-act="toToday" data-k="${esc(tv.key)}">Today’s visit</button>` : ""}</div>
      ${factsHTML(p, null)}
      <div class="section"><div class="h"><h2>Scheduling preferences</h2><span>from the tracker</span></div>${prefs}</div>
      <div class="section"><div class="h"><h2>Coverage</h2></div><dl class="facts">
        <dt>Primary</dt><dd>${esc(p.ins1 || "—")}</dd><dt>2026 BIDF</dt><dd${!p.bidf && C.isActive(p) ? ' class="flag"' : ""}>${esc(p.bidf || "Not recorded")}</dd>
        <dt>Auth start</dt><dd>${p.authStart != null ? C.fmtD(p.authStart) : "—"}</dd>${p.pcn ? `<dt>PCN</dt><dd>${esc(p.pcn)}${p.task ? `, ${esc(p.task)}` : ""}</dd>` : ""}
        ${p.notes ? `<dt>Clinic notes</dt><dd>${esc(p.notes)}</dd>` : ""}${p.pcnNotes ? `<dt>PCN notes</dt><dd>${esc(p.pcnNotes)}</dd>` : ""}</dl></div></aside>`;
  }

  /* ---------- Consults ---------- */
  function consultsHTML() {
    const L = S.CS.slice().sort((a, b) => ({ act: 0, wait: 1, done: 2, closed: 3 }[a.state] - { act: 0, wait: 1, done: 2, closed: 3 }[b.state]) || (b.days ?? 0) - (a.days ?? 0));
    const n = s => L.filter(c => c.state === s).length;
    const sub = [n("act") ? `${n("act")} to act on` : "", n("wait") ? `${n("wait")} waiting on the payer or referral` : "", n("done") ? `${n("done")} scheduled` : ""].filter(Boolean).join(", ") || "No consults on the Consults tab";
    const row = c => `<div class="crow" style="cursor:default"><span class="p">${esc(c.code || "—")}</span><span class="proto">${c.consult != null ? `Consult ${C.fmtD(c.consult)}` : "No consult date"}<br>${esc(c.provider || "")}</span><span class="thru">${esc(c.stage || "No stage")}${c.days != null ? `<br><span class="muted small">${plural(c.days, "day")} in stage</span>` : ""}</span>
      <span class="flags">${c.flag ? `<span class="f">${esc(c.flag)}</span>` : c.state === "done" ? `<span class="ok">Scheduled</span>` : c.state === "closed" ? "Not proceeding" : "Waiting"} ${c.mrn ? `<button class="btn ghost sm" data-act="openConsult" data-m="${esc(c.mrn)}">Open</button>` : ""}</span></div>`;
    return `${head("Consults", sub, false)}${!S.raw.tables?.consults ? `<div class="msg warn">There’s no table on a Consults tab in this workbook.</div>` : ""}
      <section class="card plist"><div class="chead"><span>Patient</span><span>Consult</span><span>Stage</span><span>Next step</span></div>${L.map(row).join("") || `<p class="empty" style="padding:10px 14px">Nothing on the Consults tab.</p>`}</section>`;
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
          <li>Choose that file here. It replaces MRN to Location on the Appt List tab; the Name and Lookup Key formulas are left alone.</li></ol></details></section>

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

      <section class="card set"><h3>Appearance</h3><label class="opt"><input type="checkbox" id="darkOn"${S.theme === "dark" ? " checked" : ""}> Dark mode<small>On this computer only</small></label></section>

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
        <div class="foot"><button class="btn mute" data-act="close">Cancel</button><button class="btn primary inl" data-act="doCheckout"${m.booked && !S.busy ? "" : " disabled"}>${m.booked ? `Check out ${esc(c)}` : "Choose the next visit"}</button></div>`;
    },
    proto() {
      const p = S.PM.get(S.m.mrn);
      return `<h2 id="mh">Confirm ${esc(p.code)}’s protocol</h2><p class="sub muted">${p.protoAt != null ? `Last confirmed ${C.fmtD(p.protoAt)}` : "Not confirmed on the tracker yet"}</p>
        <label class="fld">Protocol the provider confirmed<select id="mFreq">${freqOpts(p, S.m.freq)}</select></label>
        <p class="note">Writes Protocol Updated ${C.fmtD(todayS())} on the Master Patient Tracker${S.m.freq !== p.freq ? `, and Frequency “${esc(S.m.freq)}”` : ""}.</p>
        <div class="foot"><button class="btn mute" data-act="close">Cancel</button><button class="btn primary inl" data-act="doProto"${S.busy ? " disabled" : ""}>Save protocol</button></div>`;
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
        <p class="note">The ${sh.had} rows on the Appt List tab are replaced (MRN to Location). The Name and Lookup Key formulas stay.</p>
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
  const ACT = {
    nav: e => { S.view = e.dataset.v; S.bSel = null; render(); },
    full: () => openFull(),
    sel: e => { S.sel = e.dataset.k; S.bSel = null; render(); },
    bSel: e => { const F = floor(), c = F.chairs.find(x => x.id === e.dataset.b); if (c && c.v) { S.sel = c.v.key; S.bSel = null; } else S.bSel = e.dataset.b; render(); },
    pSel: e => { S.pSel = e.dataset.m; render(); },
    chip: e => { S.chip = e.dataset.v; render(); },
    wk: e => { const d = +e.dataset.v; S.wk = d === 0 ? 0 : S.wk + d; render(); },
    schedPick: e => { const m = e.dataset.m, d = +e.dataset.d, tv = S.list.find(v => v.a.mrn === m); if (d === todayS() && tv) { S.sel = tv.key; S.view = "today"; } else { S.pSel = m; S.chip = "all"; S.view = "patients"; } render(); },
    toToday: e => { S.sel = e.dataset.k; S.bSel = null; S.view = "today"; render(); },
    openPt: e => { S.pSel = e.dataset.m; S.chip = "all"; S.view = "patients"; render(); },
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
      closeModal();
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
    open: e => Store.openRow("mpt", e.dataset.m).then(ok => toast(ok ? "Opened on the Master Patient Tracker." : "Couldn’t find that row.", !ok)).catch(err => toast(String(err.message || err), true)),
    openConsult: e => Store.openRow("consults", e.dataset.m).then(ok => toast(ok ? "Opened on the Consults tab." : "Couldn’t find that row.", !ok)).catch(err => toast(String(err.message || err), true)),
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
    async saveTpl() { const o = {}; for (const k of Object.keys(C.TEMPLATES)) o[k] = $(`#tp_${k}`).value.trim() || C.TEMPLATES[k]; await write("Saving messages…", () => Store.saveKeyed("tblIvkTemplates", o), "Messages saved."); },
    async resetTpl() { await write("Saving messages…", () => Store.saveKeyed("tblIvkTemplates", C.TEMPLATES), "Standard wording restored."); },
    async saveChairs() { await write("Saving chairs…", () => Store.saveKeyed("tblIvkConfig", { Chairs: $("#cfgChairs").value.trim() || C.DEFAULTS.Chairs }), "Chairs saved."); },
    async saveCfg() {
      const o = {}; for (const k of ["VisitMin", "ProviderCheckMin", "LateMin", "SoonMin", "AheadWeeks", "ProtocolDays", "AuthWarnDays", "EncountersWarn"]) { const n = C.num($(`#cf_${k}`).value); if (n == null || n < 0) return toast(`${k} needs a number.`, true); o[k] = n; }
      o.Clinic = $("#cf_Clinic").value.trim() || C.DEFAULTS.Clinic;
      await write("Saving…", () => Store.saveKeyed("tblIvkConfig", o), "Timing saved.");
    },
    mapDown: () => { const P = S.raw.plan, a = document.createElement("a"); a.href = URL.createObjectURL(new Blob([JSON.stringify(P, null, 1)], { type: "application/json" })); a.download = `Chair map - ${S.cfg.Clinic} IVK.json`; a.click(); },
    mapRemove: () => openModal("confirm", { title: "Remove the chair map?", body: "Today shows chair tiles instead. You can load the file again any time.", yes: "Remove the map", run: () => Store.setDoc(MAP_KEY, null) }),
    async doUpload() { const u = S.m; if (u.sh.flag && !$("#mShrink")?.checked) return; closeModal(); await write("Writing the appointment list…", () => Store.replaceAppts(u.rows, C.serialOf(new Date()) + nowMin() / 1440), `Appointment list updated: ${plural(u.rows.length, "appointment")}.`); }
  };
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
    const sh = C.shrink(rows, S.A.map(a => a.date)); sh.had = S.A.length;
    openModal("upload", { file: f.name, rows, sh, ok: false });
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
    else if (t.id === "mRn") S.m.rn = t.checked;
    else if (t.id === "mAsk") S.m.ask = t.checked;
  });
  document.addEventListener("input", e => { if (e.target.id === "ptSearch") { S.q = e.target.value; const pos = e.target.selectionStart; render(); const s = $("#ptSearch"); if (s) { s.focus(); s.setSelectionRange(pos, pos); } } });
  document.addEventListener("click", e => {
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
  window.IvkApp = { S, render, derive, refresh, ACT, floor };   // for the test harness
})();
