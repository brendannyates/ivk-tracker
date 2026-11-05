/* core.js: the IVK Tracker floor rules. Pure functions (no Office.js, no DOM), shared with tests/.
   Reads the rows excel.js hands over (Master Patient Tracker, Appt List, Consults, the add-in's own IVK Visits table) and
   derives everything the screens show: today's visits and their phase, chair states, what each patient still needs, and the
   Teams messages, and the index planner on top of the provider templates (template.js). Patients are only ever named by LASFIR code. */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory(require("./template.js"));
  else root.IvkCore = factory(root.IvkTemplate);
})(typeof self !== "undefined" ? self : this, function (TM) {
  "use strict";

  /* Defaults for tblIvkConfig: 4 chairs, each IVK visit holds a chair 2 hours. The SOP has the clinician check in within
     the first 30 min and again in the last 5-10 min. Nothing clinic-specific is built in: Clinic comes from the workbook's saved
     setting, else the workbook name (clinicFromName); the start email's address, phone, CC and BCC are entered once per workbook. */
  const DEFAULTS = { Enabled: "TRUE", Clinic: "", Chairs: "4", VisitMin: 120,
    AheadWeeks: 6, ProtocolDays: 90, AuthWarnDays: 14, EncountersWarn: 3,
    /* index: 2 visits in each 7 days from treatment 1, for 3 weeks; at least a day off between visits, no more than 5 days apart
       when the slots allow; the provider's post check is in the last 30 min of the 2 hours; a hold is booked in NextGen within 10 min */
    IndexPerWeek: 2, IndexWeeks: 3, MinGapDays: 2, MaxGapDays: 5, PostCheckMin: 90, HoldMin: 10, AfterIndex: "1x per week", CapacityWeeks: 4,
    /* Oakland patients: index (the first 6) at Castro, OAK tasked to book maintenance at treatment 3. Balance reminders every 14 days.
       PayStart: the day the payment log started (visits before it are never asked about) */
    OakIndex: 6, OakRemindAt: 3, ReminderDays: 14, PayStart: "",
    /* Locations: this clinic's NextGen location names ("SF Castro; Castro IVK"), so a wider report's other clinics are kept apart.
       Blank = any location naming the clinic. ElsewhereAt: when the other clinics' visits were last uploaded */
    Locations: "", ElsewhereAt: "",
    /* the start email: the clinic's address for "your first treatment at", its phone, and who is copied on every one.
       Blank until the first person at the clinic enters them (Settings › Messages); after that only an admin changes them */
    ClinicAddress: "", ClinicPhone: "", StartCc: "", StartBcc: "",
    /* the start checklist: the Case Management stage a kept start moves the patient to in NextGen, and how many days a kept start
       stays on the list while Case Management isn't marked updated */
    CmStage: "6. Started", StartBackDays: 14,
    /* cancellations: a visit cancelled on its own day is always a late cancel; one cancelled up to LateCancelDays before it goes to
       the desk to say whether it was late. MissWeeks: how many past weeks the no-show and late-cancel rate is taken from */
    LateCancelDays: 1, MissWeeks: 8 };
  const ACTIVE = ["ongoing", "induction"];
  const DAY_NAMES = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

  /* ---------- small helpers ---------- */
  const norm = s => String(s ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");
  const txt = v => String(v ?? "").trim();
  const num = v => { if (typeof v === "number" && isFinite(v)) return v; const s = txt(v); return s !== "" && isFinite(+s) ? +s : null; };
  const pad = n => String(n).padStart(2, "0");
  /* Excel serial days <-> calendar dates, both in local time */
  const EPOCH = Date.UTC(1899, 11, 30);
  const serialOf = d => Math.round((Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - EPOCH) / 864e5);
  const dateOf = s => { const u = new Date(EPOCH + Math.floor(s) * 864e5); return new Date(u.getUTCFullYear(), u.getUTCMonth(), u.getUTCDate()); };
  const isoOf = s => { const d = dateOf(s); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
  const fromIso = s => { const m = txt(s).match(/^(\d{4})-(\d\d)-(\d\d)/); return m ? serialOf(new Date(+m[1], +m[2] - 1, +m[3])) : null; };
  /* a date cell: a serial, a Date, "10/21/2026", "2026-10-21"; anything else is null */
  function toSerial(v) {
    if (v instanceof Date && !isNaN(v)) return serialOf(v);
    const n = num(v); if (n != null) return n > 1000 && n < 200000 ? Math.floor(n) : null;
    const s = txt(v); let m;
    if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/))) return serialOf(new Date(+m[3] < 100 ? 2000 + +m[3] : +m[3], +m[1] - 1, +m[2]));
    const iso = fromIso(s); if (iso != null) return iso;
    /* older exports write the date out ("Nov 4, 2026", "Wednesday, November 04, 2026") */
    if (/[a-z]{3}/i.test(s) && /\d{4}/.test(s)) { const d = new Date(s.replace(/^[a-z]+day,?\s*/i, "")); if (!isNaN(d)) return serialOf(d); }
    return null;
  }
  /* a time cell: a day fraction (or a full serial), "1:30 PM", "13:30"; minutes after midnight, else null */
  function toMin(v) {
    const n = num(v); if (n != null) { const f = n % 1; return Math.round(f * 1440) % 1440; }
    /* also a time inside a date-time ("11/04/2026 10:00 AM"), as older exports write it */
    const m = txt(v).match(/(?:^|\s)(\d{1,2}):(\d{2})(?::\d{2})?\s*([AaPp])?/); if (!m) return null;
    let h = +m[1]; if (m[3]) { if (/p/i.test(m[3]) && h < 12) h += 12; if (/a/i.test(m[3]) && h === 12) h = 0; }
    return h * 60 + +m[2];
  }
  /* shown, saved and posted as "1:30 PM"; row keys keep "13:30" so rows saved before the switch still match */
  const fmtT = m => m == null ? "" : `${(Math.floor(m / 60) + 11) % 12 + 1}:${pad(m % 60)} ${m < 720 ? "AM" : "PM"}`;
  const keyT = m => m == null ? "" : `${Math.floor(m / 60)}:${pad(m % 60)}`;
  const fmtD = s => { if (s == null) return ""; const d = dateOf(s); return `${MONTHS[d.getMonth()].slice(0, 3)} ${d.getDate()}`; };
  const fmtDay = s => { const d = dateOf(s); return `${DAY_NAMES[d.getDay()]}, ${MONTHS[d.getMonth()]} ${d.getDate()}`; };
  const weekday = s => dateOf(s).getDay();
  /* weeks run Sunday to Saturday */
  const weekStart = s => s - weekday(s);

  /* "Smith, Jane" or "Jane Smith" -> SMIJAN. A value that is already a code (one upper-case word of 4-8 letters) is kept; any other
     single word ("Johnson") could be a surname, so it gives "" rather than being shown or posted. */
  function lasfir(name) {
    const s = txt(name), L = w => String(w || "").replace(/[^A-Za-z]/g, "").slice(0, 3); if (!s) return "";
    if (s.includes(",")) { const [last, rest] = s.split(","); return (L(last) + L(rest.trim().split(/\s+/)[0])).toUpperCase(); }
    const w = s.split(/\s+/);
    return w.length === 1 ? (/^[A-Z]{4,8}$/.test(s) ? s : "") : (L(w[w.length - 1]) + L(w[0])).toUpperCase();
  }
  /* "Rivera MD, Ana Maria", "Chen, MD Paul", "Chen, Paul MD" -> "Dr. Rivera" / "Dr. Chen" */
  function provName(p) {
    const s = txt(p); if (!s) return "";
    const last = s.split(",")[0].replace(/\b(MD|DO|NP|PA|PMHNP|FNP|RN|PhD)\b\.?/gi, "").trim();
    const md = /\b(MD|DO)\b/i.test(s);
    return last ? (md ? `Dr. ${last}` : last) : s;
  }
  const mrnKey = v => txt(v).replace(/\.0+$/, "").replace(/^0+(?=\d)/, "");
  const eventShort = e => /start/i.test(e) ? "Treatment start" : /follow/i.test(e) ? "Infusion" : txt(e) || "Visit";

  /* frequency on the tracker -> days between visits (PRN and blanks have none) */
  const FREQ_DAYS = [[/2x per week/i, 3.5], [/1x per week|weekly/i, 7], [/alt eow\/e3w/i, 17.5], [/every 2 weeks|eow/i, 14], [/every 3 weeks/i, 21],
    [/every 4 weeks/i, 28], [/every 6 weeks/i, 42], [/every 2 months/i, 61], [/every 3 months/i, 91]];
  const freqDays = f => { const s = txt(f); for (const [re, d] of FREQ_DAYS) if (re.test(s)) return d; return null; };

  /* ---------- table rows -> objects ---------- */
  /* values = a table's range including the header row. Columns are found by header name, so a moved column still reads. */
  function byHeader(values) {
    if (!values || !values.length) return { rows: [], cols: {} };
    const cols = {};
    values[0].forEach((h, i) => { const k = norm(h); if (k && cols[k] == null) cols[k] = i; });
    const rows = values.slice(1).map((r, i) => ({ r, i, get: k => { const keys = Array.isArray(k) ? k : [k]; for (const x of keys) if (cols[x] != null) return r[cols[x]]; return ""; } }));
    return { rows, cols };
  }
  const PREF_DAYS = [["Mon", "am", "pm"], ["Tue", "am2", "pm3"], ["Wed", "am4", "pm5"], ["Thu", "am6", "pm7"], ["Fri", "am8", "pm9"]];
  /* the tracker's AM/PM columns by position, not name: the first ten headers that read AM or PM (any number after it), in order,
     Mon AM to Fri PM. So a mistyped header ("PM2" where "PM9" belongs) still reads as Friday afternoon. */
  const prefCols = head => (head || []).map((h, i) => /^(am|pm)\d*$/.test(norm(h)) ? i : -1).filter(i => i >= 0).slice(0, 10);
  /* A patient listed twice: the active row (Ongoing, Induction) wins, else the later row. excel.js picks the same row to write. */
  function pickRow(rows, mrnOf, statusOf, mrn) {
    let best = -1;
    rows.forEach((r, i) => { if (mrnKey(mrnOf(r)) !== mrnKey(mrn)) return; if (best < 0 || ACTIVE.includes(txt(statusOf(r)).toLowerCase()) || !ACTIVE.includes(txt(statusOf(rows[best])).toLowerCase())) best = i; });
    return best;
  }
  /* Master Patient Tracker rows, one per MRN (pickRow). i = the row's index in the table body (for writes back). */
  function patients(values) {
    const { rows } = byHeader(values), keep = new Set(), pc = prefCols(values && values[0]), on = (r, j) => pc[j] != null && num(r[pc[j]]) === 1;
    const live = rows.filter(x => txt(x.get("mrn")));
    live.forEach(x => keep.add(pickRow(live, y => y.get("mrn"), y => y.get("status"), x.get("mrn"))));
    return live.filter((x, j) => keep.has(j)).map(x => {
      const g = x.get, el = g("encountersremaining"), ea = g("encountersauthorized");
      return {
        i: x.i, mrn: mrnKey(g("mrn")), code: lasfir(g("patientname")), status: txt(g("status")),
        /* the name and DOB are only for the forms printed for the patient; never shown on screen or posted */
        name: txt(g("patientname")), dob: toSerial(g(["dob", "dateofbirth", "patientdob", "birthdate"])),
        authExp: toSerial(g(["authexpdate", "authexp"])), encLeft: num(el), encLeftTxt: txt(el), encAuth: num(ea), encAuthTxt: txt(ea),
        copay: txt(g("copay")), infusion: txt(g("infusion")), protoAt: toSerial(g(["protocolupdated", "protocolexpdate"])),
        freq: txt(g(["frequency", "txfrequency"])), loc: txt(g("primarylocation")),
        ins1: txt(g(["primaryencounterinsurance", "primaryinsurance"])), ins2: txt(g(["infusioninsurance", "secondaryinsurance"])),
        prefs: PREF_DAYS.map(([d], j) => ({ d, am: on(x.r, 2 * j), pm: on(x.r, 2 * j + 1) })),
        start: toSerial(g("txstartdate")), bidf: txt(g("2026bidf")), authNo: txt(g("authnumber")), authStart: toSerial(g("authstart")),
        pcn: txt(g("pcn")), task: txt(g("taskstatus")), pcnNotes: txt(g("pcnnotes")), notes: txt(g("clinicnotes")),
        reqPost: /^(y|yes|true|1|x)$/i.test(txt(g(["requirespostcheck", "postcheckrequired", "requirespostcheckyn"])))
      };
    });
  }
  /* Appt List rows: MRN | Patient | Date | Time | Provider | Event | Status | Location | Name (LASFIR formula) | Lookup Key */
  function appts(values) {
    const { rows } = byHeader(values);
    return rows.map(x => {
      const g = x.get, date = toSerial(g("date")), time = toMin(g("time")), mrn = mrnKey(g("mrn"));
      const code = /^[A-Z]{4,8}$/.test(txt(g("name"))) ? txt(g("name")) : lasfir(g("patient"));
      return { mrn, code, date, time, provider: txt(g("provider")), event: txt(g("event")), status: txt(g("status")), loc: txt(g("location")) };
    }).filter(a => a.mrn && a.date != null);
  }
  const isCancelled = a => /cancel|resched/i.test(a.status);
  const visitKey = (date, mrn, time) => `${isoOf(date)}|${mrn}|${keyT(time)}`;

  /* IVK Visits rows (the add-in's own table): what the desk did today. Times are "1:22 PM" text. */
  const VISIT_COLS = ["Key", "Date", "MRN", "Patient", "Appt", "Arrived", "Chair", "Seated", "Out", "Booked", "Asked", "Protocol", "Status", "By", "Updated"];
  function visitRows(values) {
    const { rows } = byHeader(values), m = new Map();
    rows.forEach(x => {
      const g = x.get, key = txt(g("key")); if (!key) return;
      m.set(key, { key, i: x.i, arr: toMin(g("arrived")), chair: txt(g("chair")), seat: toMin(g("seated")), out: toMin(g("out")),
        booked: txt(g("booked")).toLowerCase(), asked: toMin(g("asked")), proto: txt(g("protocol")), status: txt(g("status")).toLowerCase(), by: txt(g("by")) });
    });
    return m;
  }

  /* ---------- today's floor ---------- */
  /* phase of one visit: gone (not coming), noshow (NextGen says so), done (checked out), tx (in treatment: the Seated column
     holds when it started), in (checked in), booked. A chair is only kept when a chair map is loaded. */
  function phase(v) {
    if (v.st.status === "cancelled") return "gone";
    if (v.st.out != null) return "done";
    if (v.st.seat != null) return "tx";
    if (v.st.arr != null) return "in";
    if (/no ?show/i.test(v.a.status)) return "noshow";
    return "booked";
  }
  const isActive = p => !!p && ACTIVE.includes(p.status.toLowerCase());
  /* future visits for a patient after a day, from the appointment list (cancelled ones don't count) */
  const futureOf = (A, mrn, after) => A.filter(a => a.mrn === mrn && a.date > after && !isCancelled(a)).sort((x, y) => x.date - y.date || (x.time ?? 0) - (y.time ?? 0));

  /* the patient's last visit before a day that wasn't cancelled or a no-show (null when the list has none) */
  function lastSeen(A, mrn, before) {
    let best = null;
    for (const a of A) if (a.mrn === mrn && a.date < before && !isCancelled(a) && !/no ?show/i.test(a.status) && (!best || a.date > best.date)) best = a;
    return best;
  }
  /* Does this patient need booking out, and is the protocol due a confirmation? x = { today, cfg, appts }.
     book: nothing booked after today, or booked through less than AheadWeeks out.
     proto (the standard ask): Protocol Updated is filled in and older than ProtocolDays, or the patient has gone two of their
     intervals without a visit (every 2 weeks, last seen over 28 days ago). A blank Protocol Updated alone doesn't send the ask
     (unconfirmed shows it on Patients instead). why: the reasons, for the Teams card. */
  function asks(p, x, st) {
    const c = x.cfg, T = x.today, fut = futureOf(x.appts, p?.mrn ?? "", T), thru = fut.length ? fut[fut.length - 1].date : null;
    const ahead = +c.AheadWeeks * 7, A = !!p && isActive(p), prn = /prn/i.test(p?.freq ?? "");
    const expired = A && p.protoAt != null && T - p.protoAt > +c.ProtocolDays;
    const fd = A && !prn ? freqDays(p.freq) : null, last = A ? lastSeen(x.appts, p.mrn, T) : null;
    const overdue = !!fd && !!last && T - last.date > 2 * fd;
    const book = A && !prn && (thru == null || thru < T + ahead) && !(st && /^(yes|no)$/.test(st.booked));
    const done = !!(st && st.proto), proto = (expired || overdue) && !done, why = [];
    if (proto && expired) why.push(`Protocol last confirmed ${fmtD(p.protoAt)}, ${T - p.protoAt} days ago`);
    if (proto && overdue) why.push(`Last seen ${fmtD(last.date)}, ${T - last.date} days ago (${p.freq})`);
    return { book, proto, expired, overdue, last, why, unconfirmed: A && p.protoAt == null && !done, next: fut[0] || null, thru, any: book || proto };
  }

  /* Today's visits, each with its phase and what the desk still has to do. */
  function today(x) {
    const T = x.today, st = x.visits || new Map(), P = new Map((x.patients || []).map(p => [p.mrn, p]));
    const list = x.appts.filter(a => a.date === T && !isCancelled(a)).map(a => {
      const key = visitKey(T, a.mrn, a.time), s = st.get(key) || { key, arr: null, chair: "", seat: null, out: null, booked: "", asked: null, proto: "", status: "", by: "" };
      const v = { key, a, st: s, p: P.get(a.mrn) || null }; v.ph = phase(v); v.ask = asks(v.p, x, s);
      return v;
    });
    /* a desk entry for a visit no longer on the list (moved in NextGen) still shows while the patient is here */
    st.forEach((s, key) => {
      if (!key.startsWith(isoOf(T) + "|") || list.some(v => v.key === key) || s.out != null || s.status === "cancelled" || s.arr == null) return;
      const [, mrn, hm] = key.split("|"), p = P.get(mrn) || null;
      const v = { key, a: { mrn, code: p?.code || "", date: T, time: toMin(hm), provider: "", event: "", status: "Not on the list", loc: "" }, st: s, p, off: true };
      v.ph = phase(v); v.ask = asks(p, x, s); list.push(v);
    });
    return list.sort((a, b) => (a.a.time ?? 1e4) - (b.a.time ?? 1e4));
  }
  const code = v => v.p?.code || v.a.code || "Patient";

  /* in treatment with a protocol or booking ask nobody has sent yet: the providers need asking before discharge */
  const askDue = v => v.ph === "tx" && !!v.p && v.ask.any && v.st.asked == null;
  /* Chair state on the map. open: empty; track: someone checked in or in treatment there; soon: in treatment and the providers
     still need asking. Nothing on a chair is red: waiting time isn't tracked. */
  function chairState(chair, list) {
    const v = list.find(y => (y.ph === "tx" || y.ph === "in") && y.st.chair === chair);
    if (!v) return { state: "open", v: null };
    return { state: askDue(v) ? "soon" : "track", v };
  }
  const chairStatus = cs => !cs.v ? "" : cs.state === "soon" ? "Ask providers" : cs.v.ph === "tx" ? "In treatment" : "Checked in";

  /* chairs only come from a loaded chair map; with none, visits are tracked without chairs */
  const chairList = plan => plan ? plan.bays.map(b => b.id) : [];
  /* how many visits can run at once, for open slots: the map's chairs, else the Chairs setting (a number, or the old list "1, 2, 3") */
  const chairCount = (cfg, plan) => { if (plan) return plan.bays.length; const t = txt(cfg.Chairs); return /^\d+$/.test(t) ? +t : t.split(/[,;\s]+/).filter(Boolean).length || 4; };

  /* ---------- needs (Patients) ---------- */
  /* What one patient needs, sorted red, amber, then quiet. Red is kept for what blocks the next infusion: an expired
     authorization, no visits left on it, or an active patient with nothing booked. Each need carries the action that resolves it. */
  function needs(p, x) {
    const n = [], add = (k, t, act, o) => n.push(Object.assign({ k, t, act }, o)), c = x.cfg, T = x.today;
    if (/discontinued/i.test(p.status)) return n;
    const A = isActive(p), a = asks(p, x, null), fut = futureOf(x.appts, p.mrn, T - 1), ED = ["Update the auth", "edit"];
    if (A && p.authExp != null && p.authExp < T) add("auth", `Auth ended ${fmtD(p.authExp)}`, ED, { bad: true });
    else if (A && p.authExp != null && p.authExp <= T + +c.AuthWarnDays) add("auth", `Auth ends ${fmtD(p.authExp)}`, ED);
    if (A && p.encLeft != null && p.encLeft <= 0) add("enc", "No visits left on the auth", ED, { bad: true });
    else if (A && p.encLeft != null && p.encLeft <= +c.EncountersWarn) add("enc", `${p.encLeft} visit${p.encLeft === 1 ? "" : "s"} left on the auth`, ED);
    /* money: a balance from missed payments, its reminder, and past visits nobody said were paid */
    const pt = x.PT ? x.PT.get(p.mrn) || null : null, pays = x.pays || [];
    if (x.pays) {
      const L = ledger(p.mrn, pays), last = L.lastReminder;
      if (L.balance > 0) add("bal", `Balance ${money(L.balance)}${pt?.card && pt?.okCollect ? ", card on file" : ""}`, ["Record a payment", "pay"]);
      if (L.balance > 0 && (!last || T - (last.date ?? 0) >= +c.ReminderDays)) add("remind", last ? `Balance reminder last sent ${fmtD(last.date)}` : "No balance reminder sent", ["Record a reminder", "pay"]);
      const tc = toConfirm(p, x.appts, pays, T, x.payStart); if (tc.length) add("payconf", `${tc.length} payment${tc.length === 1 ? "" : "s"} to confirm`, ["Confirm payments", "pay"]);
    }
    /* Oakland: OAK books maintenance from treatment OakRemindAt; after the index they go back */
    const oak = oakState(p, pt, x.appts, T, c);
    if (oak && A && oak.task) add("oak", `Treatment ${oak.done}: task OAK to book maintenance`, ["Mark OAK tasked", "oakTask"]);
    if (oak && A && oak.back) add("oakback", `Index done: back to Oakland for maintenance (${oak.after.length} booked here)`, ["Mark OAK tasked", "oakTask"]);
    /* availability: expired, or a visit booked on a day they said they're away */
    if (pt && A && pt.availExp != null && pt.availExp < T) add("avail", `Availability expired ${fmtD(pt.availExp)}`, ["Update availability", "avail"]);
    const clash = pt ? fut.find(f => awayOn(pt.away, f.date)) : null;
    if (clash && A) add("away", `Booked ${fmtD(clash.date)} while away`, ["Update availability", "avail"]);
    if (A && !/prn/i.test(p.freq) && !fut.length) add("book", "Nothing booked", ["Ask the group chat", "ask"], { bad: true });
    else if (A && a.book) add("book", a.thru == null ? "Nothing booked after today" : `Booked through ${fmtD(a.thru)}`, ["Ask the group chat", "ask"]);
    if (A && !p.freq) add("freq", "No frequency on the tracker", ["Ask the group chat", "ask"]);
    if (A && (a.proto || a.unconfirmed)) add("proto", a.overdue ? `Not seen since ${fmtD(a.last.date)} (${p.freq}): confirm the protocol`
      : p.protoAt == null ? "Protocol not confirmed" : `Protocol last confirmed ${fmtD(p.protoAt)}`, ["Confirm protocol", "proto"]);
    if (A && p.encAuth != null && p.authStart == null) add("authstart", "No auth start date", ED);
    /* a treatment start without the rest of index booked behind it */
    const sv = A ? nextStart(p, x.appts, T) : null, short = sv ? indexShort(p, x.appts, sv, c) : null;
    if (short) add("idxbook", `Index not fully booked: ${short.booked} of ${short.need} visits after the ${fmtD(sv.date)} start`, ["Plan index visits", "pvPlan"]);
    /* the start checklist: Case Management in NextGen on the day of the start (x.SL, from startList) */
    const cms = (x.SL || []).find(s => s.mrn === p.mrn && s.cmDue);
    if (cms) add("cm", cms.today ? `Treatment start today: update Case Management to ${c.CmStage || DEFAULTS.CmStage}` : `Start kept ${fmtD(cms.date)}: update Case Management to ${c.CmStage || DEFAULTS.CmStage}`, ["Case Management updated", "cmDone"]);
    /* a transfer to another clinic: confirmed at the last visit here; visits booked here after it need moving */
    const xs = xferState(x.XF ? x.XF.get(p.mrn) : null, x.appts, T, x.E);
    if (xs && xs.after.length) add("xferafter", `Transfer to ${xs.to} ${fmtD(xs.date)}: ${xs.after.length} visit${xs.after.length === 1 ? "" : "s"} still booked here after it`, ["Transfer", "xfer"]);
    if (xs && xs.due) add("xferdue", `Confirm the transfer to ${xs.to}${xs.last ? ` (last visit here ${fmtD(xs.last.date)})` : ""}`, ["Confirm transfer", "xfer"]);
    else if (xs) add("xferq", `Transfer to ${xs.to}${xs.date != null ? ` ${fmtD(xs.date)}` : ""}`, null, { quiet: true });
    /* the hand-off, from the wider report: nothing booked at the new clinic yet; once booked, a quiet tag */
    if (xs && xs.there === null) add("xfernone", `Transfer to ${xs.to}: nothing booked there yet`, ["Transfer", "xfer"]);
    else if (xs && xs.there) add("xferthere", `Booked at ${xs.there.loc || xs.to} ${fmtD(xs.there.date)}`, null, { quiet: true });
    /* Oakland: OAK tasked, but no visit at Oakland on the wider report once the index here is done or one visit from it */
    const oakThere = oak && x.E ? bookedAt(x.E, p.mrn, "Oakland", T) : undefined;
    if (oak && A && oak.tasked != null && oakThere === null && oak.done != null && oak.done >= oak.idx - 1) add("oakbook", `OAK tasked ${fmtD(oak.tasked)}: nothing booked at Oakland yet`, null);
    else if (oakThere) add("oakthere", `Booked at ${oakThere.loc || "Oakland"} ${fmtD(oakThere.date)}`, null, { quiet: true });
    else if (!oak && !xs && x.E && A) { const e = bookedAt(x.E, p.mrn, "", T); if (e) add("elsewhere", `Also booked at ${e.loc || "another clinic"} ${fmtD(e.date)}`, null, { quiet: true }); }
    /* index done: an Induction patient who has kept their 6th visit moves to Ongoing on the protocol the provider sets */
    const ix = indexState(p, x.appts, T, c);
    if (ix && ix.complete) add("indexdone", `Index done (${ix.done} kept): move to Ongoing and confirm the protocol`, ["Move to Ongoing", "idxDone"]);
    const sd = A ? stepDue(p, x.appts, T, c) : null;
    if (sd) add("stepdown", `Kept all ${sd.N} at ${sd.freq}, protocol not updated: change to ${sd.to}?`, ["Change protocol", "stepDown"]);
    /* a visit or hold on a day the clinic is closed */
    const shut = A && x.EX ? fut.find(f => closureOn(x.EX, f.date)) : null;
    if (shut) add("closed", `Booked ${fmtD(shut.date)}: clinic closed (${closureOn(x.EX, shut.date).note || "closure"})`, ["Open the schedule", "schedAt"], { bad: true });
    /* no-shows and late cancels, counted (x.missBy: missByPatient) */
    const mb = x.missBy ? x.missBy.get(p.mrn) : null;
    if (mb && mb.noshow + mb.late) add("missed", `${[mb.noshow ? `${mb.noshow} no-show${mb.noshow === 1 ? "" : "s"}` : "", mb.late ? `${mb.late} late cancel${mb.late === 1 ? "" : "s"}` : ""].filter(Boolean).join(", ")}${mb.exempt === "kpva" ? " (Kaiser/VA: no fee)" : ""}`, null, { quiet: true });
    if (/induction/i.test(p.status)) add("induction", "Induction", null, { quiet: true });
    if (/paused/i.test(p.status)) add("paused", "Paused", null, { quiet: true });
    if (oak) add("oakq", oak.unknown ? "Oakland patient" : `Oakland: index here, ${oak.done} of ${oak.idx} done`, null, { quiet: true });
    else if (elsewhere(p, c)) add("home", `Home clinic ${p.loc}`, null, { quiet: true });
    if (/self ?pay/i.test(p.ins2)) add("selfpay", "Self-pay infusion", null, { quiet: true });
    if (p.reqPost) add("post", "Requires post-check", null, { quiet: true });
    const rank = y => y.quiet ? 2 : y.bad ? 0 : 1;
    return n.map((y, i) => [y, i]).sort((u, w) => rank(u[0]) - rank(w[0]) || u[1] - w[1]).map(u => u[0]);
  }
  const loud = list => list.filter(y => !y.quiet);

  /* ---------- Consults ---------- */
  /* Consults rows: Location | Referral Date | Start Date | Patient | Patient MRN | Consult Date | Provider | Patient Journey Stage |
     Individual PJ Stage | Days in current Stage | Case Management Latest Status Note | Kept Start | CM Updated | CM Flag
     (the last three: the kept treatment start since the consult, a formula on the Appt List; the day Case Management was updated
     in NextGen; and the formula's flag while it wasn't) */
  function consults(values, today) {
    const { rows } = byHeader(values);
    return rows.filter(x => txt(x.get("patient")) || txt(x.get("patientmrn"))).map(x => {
      const g = x.get, stage = txt(g("individualpjstage")) || txt(g("patientjourneystage")), days = num(g("daysincurrentstage"));
      const c = { i: x.i, loc: txt(g("location")), code: lasfir(g("patient")), mrn: mrnKey(g("patientmrn")), ref: toSerial(g("referraldate")),
        start: toSerial(g("startdate")), consult: toSerial(g("consultdate")), provider: provName(g("provider")), stage, days,
        keptSheet: toSerial(g("keptstart")), cmAt: toSerial(g("cmupdated")) };
      const st = stage.toLowerCase();
      c.state = /not proceeding/.test(st) ? "closed" : /scheduled/.test(st) && !/not yet scheduled/.test(st) ? "done"
        : /not yet scheduled/.test(st) ? "act" : /unclear/.test(st) ? "act" : "wait";
      c.flag = c.state === "act" ? (/unclear/.test(st) ? "Ask the provider whether to refer" : "Book the treatment start") : "";
      return c;
    });
  }

  /* ---------- treatment starts: kept, Case Management, the start checklist (tblIvkStarts) ---------- */
  /* Case Management's stage once a start is kept: "6. Started" (the Individual PJ Stage; the combined "5. ... / 6. Started" isn't) */
  const isStartedStage = s => /^\s*6\b/.test(txt(s));
  const isStartEv = e => /start/i.test(txt(e));
  const keptSt = a => /\b(kept|completed?|checked ?out)\b/i.test(txt(a.status));
  /* the latest kept treatment start on the appointment list on or after a day (from null = any) */
  function keptStartOf(mrn, A, from) {
    const L = A.filter(a => a.mrn === mrn && isStartEv(a.event) && keptSt(a) && (from == null || a.date >= from)).sort((x, y) => x.date - y.date);
    return L.length ? L[L.length - 1] : null;
  }
  /* Key | Date | MRN | Patient | CmUpdated | By | Updated. One row per treatment start (start day|MRN, so a new time keeps the tick):
     the day the technician updated Case Management in NextGen for it. */
  const START_COLS = ["Key", "Date", "MRN", "Patient", "CmUpdated", "By", "Updated"];
  const startKey = (date, mrn) => `${isoOf(date)}|${mrnKey(mrn)}`;
  function startRows(values) {
    const m = new Map();
    byHeader(values).rows.forEach(x => { const g = x.get, mrn = mrnKey(g("mrn")), date = toSerial(g("date")); if (!mrn || date == null) return;
      m.set(startKey(date, mrn), { key: startKey(date, mrn), date, mrn, code: txt(g("patient")), cm: toSerial(g("cmupdated")), by: txt(g("by")) }); });
    return m;
  }
  /* Consults rows against the appointment list: kept (the start kept since the consult, else the tab's Kept Start), and once kept the
     row leaves the waitlist: state "cm" while Case Management still needs updating (not at stage 6, no CM Updated day on the tab or
     the checklist), else "started". Changes the rows in place. */
  function consultStarts(CS, A, ST, cfg) {
    const stage = txt((cfg || {}).CmStage) || DEFAULTS.CmStage;
    for (const c of CS) {
      if (!c.mrn || c.state === "closed") continue;
      const a = keptStartOf(c.mrn, A, c.consult), kept = a ? Math.max(a.date, c.keptSheet ?? 0) : c.keptSheet;
      c.kept = kept ?? null; if (c.kept == null) continue;
      const rec = ST ? ST.get(startKey(c.kept, c.mrn)) : null;
      c.cmDone = isStartedStage(c.stage) || c.cmAt != null || (rec?.cm != null);
      c.cmAt = c.cmAt ?? rec?.cm ?? null;
      c.state = c.cmDone ? "started" : "cm";
      c.flag = c.cmDone ? "" : `Started ${fmtD(c.kept)}: update Case Management to ${stage}`;
    }
    return CS;
  }
  /* The start checklist. o = { today, appts, PM (tracker), PF (profiles), ST (tblIvkStarts), CS (Consults), visits (desk Map), cfg }.
     Every treatment start from StartBackDays ago to `ahead` days out, with two items:
       profile  a newly booked start: the patient needs a profile (done once they're on the tracker or have a profile row)
       cm       on the day of the start (and after it, once kept): update Case Management in NextGen to CmStage
     cmDue: the day has come, Case Management isn't marked updated (or at stage 6 on the Consults tab), and a past start was kept.
     open: something still to do. Today's starts are always listed; others only while open. */
  function startList(o, ahead = 28) {
    const T = o.today, back = +o.cfg?.StartBackDays || DEFAULTS.StartBackDays, by = new Map();
    for (const a of o.appts) {
      if (!isStartEv(a.event) || isCancelled(a) || a.date < T - back || a.date > T + ahead) continue;
      const k = startKey(a.date, a.mrn), was = by.get(k); if (!was || (a.time ?? 0) < (was.time ?? 0)) by.set(k, a);
    }
    const out = [];
    for (const [key, a] of by) {
      const rec = o.ST ? o.ST.get(key) || null : null, c = (o.CS || []).find(x => x.mrn === a.mrn) || null;
      const out1 = o.visits ? o.visits.get(visitKey(a.date, a.mrn, a.time)) : null, noshow = /no ?show/i.test(a.status);
      const kept = keptSt(a) || (out1 && out1.out != null);
      const profDone = !!(o.PM && o.PM.has(a.mrn)) || !!(o.PF && o.PF.has(a.mrn));
      const cmDone = rec?.cm != null || (c && (isStartedStage(c.stage) || (c.cmAt != null && c.cmAt >= a.date)));
      const cmDue = a.date <= T && !cmDone && !noshow && (a.date === T || kept);
      const s = { key, a, mrn: a.mrn, code: a.code || o.PM?.get(a.mrn)?.code || "", date: a.date, kept: !!kept, noshow, rec, c,
        profDone, profDue: !profDone && a.date >= T, cmDone: !!cmDone, cmDue, cmAt: rec?.cm ?? c?.cmAt ?? null, today: a.date === T };
      s.open = s.profDue || s.cmDue;
      if (s.open || s.today) out.push(s);
    }
    return out.sort((x, y) => x.date - y.date || (x.a.time ?? 0) - (y.a.time ?? 0));
  }

  /* ---------- Teams messages ---------- */
  const TEMPLATES = {
    arrival: "{patient} is here for their {time} IVK visit with {provider}.{chair}",
    seated: "{patient} is in Chair {chair}.",
    protocol: "Can we please confirm {protocol} protocol for {patient} and send to the desk at discharge for scheduling?",
    protocolOnly: "Can we please confirm {protocol} protocol for {patient}?",
    bookOnly: "Can we please send {patient} to the desk at discharge for scheduling?"
  };
  /* {key} -> value; <b>x</b> -> **x** for the Adaptive Card; stray spaces before punctuation go */
  function fill(tpl, d) {
    return String(tpl).replace(/\{(\w+)\}/g, (_, k) => d[k] == null ? "" : String(d[k]))
      .replace(/<\/?b>/gi, "**").replace(/<[^>]*>/g, "").replace(/\s{2,}/g, " ").replace(/\s+([,.?)])/g, "$1").trim();
  }
  /* the frequency read the way the team says it: "every 2 weeks" -> "every 2 week"; blank or PRN -> "the" */
  function protoWords(freq) {
    const f = txt(freq); if (!f || /^prn$/i.test(f)) return "the";
    return f.replace(/\bevery (\d+) (week|month)s\b/i, "every $1 $2");
  }
  function arrivalText(v, chair, tpl) {
    return fill(tpl || TEMPLATES.arrival, { patient: code(v), time: fmtT(v.a.time), provider: provName(v.a.provider) || "the provider",
      chair: chair ? ` Chair ${chair}.` : " Waiting for a chair." });
  }
  /* which ask: confirm + book (the team's usual message), or just one of them */
  function askText(p, ask, tpls, freq) {
    const t = Object.assign({}, TEMPLATES, tpls || {}), d = { patient: p?.code || "", protocol: protoWords(freq ?? p?.freq) };
    const which = ask.book && ask.proto ? "protocol" : ask.book ? "bookOnly" : ask.proto ? "protocolOnly" : "protocol";
    return fill(t[which], d);
  }

  /* ---------- Teams @mentions and card replies ---------- */
  /* Who a post tags, kept in tblIvkConfig as "Name <email>": TagArrival (the RN) and "Tag Dr. Lark" for each provider (provName).
     Names and emails stay in the workbook, never in the code. */
  const tagKey = prov => `Tag ${prov}`;
  function parseTag(v) {
    const m = txt(v).match(/^(.*?)\s*<\s*([^<>\s@]+@[^<>\s@]+\.[^<>\s@]+)\s*>$/);
    return m && m[1].trim() ? { name: m[1].trim(), id: m[2].toLowerCase() } : null;
  }
  const tagStr = (name, email) => txt(name) && txt(email) ? `${txt(name)} <${txt(email)}>` : "";
  /* { arrival: tag|null, prov: Map(provName -> tag) } from the config */
  function tagsOf(cfg) {
    const prov = new Map();
    for (const [k, v] of Object.entries(cfg || {})) { const m = /^tag (.+)$/i.exec(k); if (m && !/^arrival$/i.test(m[1])) { const t = parseTag(v); if (t) prov.set(m[1].trim(), t); } }
    return { arrival: parseTag(cfg && cfg.TagArrival), prov };
  }
  /* the provider a protocol ask goes to: today's visit, else the next booked one, else the last one seen */
  const askProvider = (v, a) => provName(v?.a?.provider || a?.next?.provider || a?.last?.provider || "");

  /* tblIvkReplies: one row per card that takes an answer in Teams. The add-in writes Ref..SentAt before posting; the Teams
     workflow fills Reply (the card's answer as JSON), ReplyBy, ReplyEmail and RepliedAt; the add-in applies it and fills Applied.
     Kind chair: Value = the chair the card was sent with. Kind protocol: Value = the frequency on file. The card carries only Ref. */
  const REPLY_COLS = ["Ref", "Kind", "Key", "MRN", "Patient", "Value", "SentBy", "SentAt", "Reply", "ReplyBy", "ReplyEmail", "RepliedAt", "Applied"];
  function replyRows(values) {
    return byHeader(values).rows.map(x => {
      const g = x.get, raw = txt(g("reply")); let data = null;
      if (raw) { try { data = JSON.parse(raw); } catch (e) { data = { bad: raw }; } }
      return { i: x.i, ref: txt(g("ref")), kind: txt(g("kind")).toLowerCase(), key: txt(g("key")), mrn: mrnKey(g("mrn")), patient: txt(g("patient")),
        value: txt(g("value")), sentAt: txt(g("sentat")), data, by: txt(g("replyby")), email: txt(g("replyemail")).toLowerCase(), at: txt(g("repliedat")), applied: txt(g("applied")) };
    }).filter(r => r.ref);
  }
  /* what an answered card asks the tracker to do, or why it can't. chairs = the chair ids; freqs = the frequency list;
     providers = the tagged providers' emails (only they can confirm a protocol). */
  function replyAction(r, o) {
    const d = r.data || {};
    if (!r.data || d.bad) return { skip: "the answer couldn’t be read" };
    if (r.kind === "chair") {
      const ch = txt(d.chair);
      if (!ch) return { skip: "no chair picked" };
      if (o.chairs && o.chairs.length && !o.chairs.includes(ch)) return { skip: `Chair ${ch} isn’t on the chair map` };
      return { chair: ch };
    }
    if (r.kind === "protocol") {
      const f = txt(/^pick$/i.test(d.act) ? d.pick : d.freq);
      if (!f) return { skip: "no protocol picked" };
      if (o.freqs && o.freqs.length && !o.freqs.some(x => x.toLowerCase() === f.toLowerCase())) return { skip: `“${f}” isn’t on the frequency list` };
      if (!o.providers || !o.providers.includes(r.email)) return { skip: `${r.by || r.email || "Someone"} isn’t a tagged provider, so the protocol wasn’t saved` };
      return { freq: o.freqs && o.freqs.find(x => x.toLowerCase() === f.toLowerCase()) || f };
    }
    return { skip: "unknown card" };
  }
  const repliesDue = R => (R || []).filter(r => r.data && !r.applied);

  /* ---------- NextGen "Appointments Listing" export -> Appt List rows ---------- */
  /* Report headers (row 7 of the export): MRN | Location | Appointment Date | Time | Provider | Appointment Type | Patient | Status.
     Out: [MRN, Patient, Date serial, Time fraction, Provider, Event, Status, Location], the Appt List table's order. */
  const REPORT_HDR = { mrn: ["mrn", "mdrc", "medrecnbr", "medrecno", "medicalrecordnumber", "medrec"], patient: ["patient", "patientname", "patname"],
    date: ["appointmentdate", "apptdate", "apptdt", "date", "appointmentdatetime", "apptdatetime"],
    time: ["time", "begtm", "appointmenttime", "appttime", "starttime", "begintime", "begtime"], provider: ["provider", "renderingprovider", "resource", "providername", "resourcename"],
    event: ["appointmenttype", "event", "eventtype", "apptevent", "eventname", "appttype"], status: ["status", "sts", "appointmentstatus", "apptstatus"],
    loc: ["location", "loc", "locationname", "apptlocation"] };
  const REPORT_ORDER = ["mrn", "patient", "date", "time", "provider", "event", "status", "loc"];
  function parseReport(aoa) {
    const hr = aoa.slice(0, 40).findIndex(r => r.some(v => REPORT_HDR.mrn.includes(norm(v))) && r.some(v => REPORT_HDR.date.includes(norm(v))));
    if (hr < 0) return null;
    const hn = aoa[hr].map(norm), idx = REPORT_ORDER.map(k => hn.findIndex(h => REPORT_HDR[k].includes(h)));
    if (["mrn", "patient", "date", "status"].some(k => idx[REPORT_ORDER.indexOf(k)] < 0)) return null;
    const rows = [];
    for (const r of aoa.slice(hr + 1)) {
      /* no time column (older exports): the time inside the date-time */
      const v = idx.map(i => i < 0 ? "" : r[i] ?? ""), d = toSerial(v[2]), t = idx[3] >= 0 ? toMin(v[3]) : num(v[2]) != null ? (num(v[2]) % 1 ? toMin(v[2]) : null) : toMin(v[2]);
      if (txt(v[0]) === "" || d == null) continue;   // blank and footer lines
      v[0] = txt(v[0]); v[1] = txt(v[1]); v[2] = d; v[3] = t == null ? "" : t / 1440; for (const j of [4, 5, 6, 7]) v[j] = txt(v[j]);
      rows.push(v);
    }
    return rows;
  }
  /* flag a list that is much smaller than the one on file: under 20 rows, under half, or ends more than a week earlier */
  function shrink(rows, curDates) {
    const n = rows.length, had = curDates.length, ds = rows.map(r => r[2]), cd = curDates.filter(d => d != null);
    const nMax = ds.length ? Math.max(...ds) : null, cMax = cd.length ? Math.max(...cd) : null;
    const few = n < 20, half = had >= 20 && n < had * 0.5, short = nMax != null && cMax != null && nMax < cMax - 7;
    return { flag: few || half || short, few, half, short, n, had, nMax, cMax };
  }

  /* ---------- Index planner ---------- */
  const DAYS5 = ["Mon", "Tue", "Wed", "Thu", "Fri"];
  const dayNum = s => { const i = ["mo", "tu", "we", "th", "fr"].indexOf(norm(s).slice(0, 2)); return i < 0 ? null : i + 1; };
  const sameProv = (a, b) => TM.provKey(a) === TM.provKey(b);
  /* tblIvkOther: the providers' other NextGen visits (telehealth, TMS follow-ups, consults) from the wider report, with no patient
     on them: Date | Time | Provider | Event | Status. They take the template cell they start in, so a flex post check booked for a
     tele visit can't be an index post check. */
  const OTHER_COLS = ["Date", "Time", "Provider", "Event", "Status"];
  const otherRows = values => byHeader(values).rows.map(x => ({ date: toSerial(x.get("date")), time: toMin(x.get("time")), provider: txt(x.get("provider")), event: txt(x.get("event")), status: txt(x.get("status")) }))
    .filter(o => o.date != null && o.time != null);
  /* tblIvkExceptions: a provider's time off (PTO) on their template, by date. Id | Provider | From | To | Coverage | UseTemplate |
     Note | Status | By | Updated. On those days the provider's template slots aren't offered. With Coverage, the covering provider
     either works the absent provider's template (UseTemplate, a pre-fill that replaces the covering provider's own template that day)
     or keeps their own template. Status "removed" takes the entry back out.
     A clinic closure (a holiday) is a row with Provider "Clinic closed": every provider's slots close those days, nobody covers. */
  const EX_COLS = ["Id", "Provider", "From", "To", "Coverage", "UseTemplate", "Note", "Status", "By", "Updated"];
  const CLOSED = "Clinic closed", isClosedRow = v => /^clinic( closed)?$|^all providers$/i.test(txt(v));
  function exRows(values) {
    return byHeader(values).rows.map(x => { const g = x.get, from = toSerial(g("from")), to = toSerial(g("to")) ?? from, closed = isClosedRow(g("provider")), cov = closed ? "" : txt(g("coverage"));
      return { id: txt(g("id")), provider: closed ? CLOSED : txt(g("provider")), pkey: closed ? "*" : TM.provKey(g("provider")), closed, from, to: to == null ? null : Math.max(from, to), cov, covKey: cov ? TM.provKey(cov) : "",
        useTpl: !!cov && /^(true|yes|1)$/i.test(txt(g("usetemplate"))), note: txt(g("note")), status: txt(g("status")).toLowerCase(), by: txt(g("by")) }; })
      .filter(e => e.id && e.pkey && e.from != null && e.status !== "removed");
  }
  /* a closure wins over a provider's own time off: nobody covers on a day the clinic is shut */
  const exFor = (EX, pkey, d) => EX.find(e => e.closed && e.from <= d && d <= e.to) || EX.find(e => e.pkey === pkey && e.from <= d && d <= e.to) || null;
  const closureOn = (EX, d) => (EX || []).find(e => e.closed && e.from <= d && d <= e.to) || null;
  /* The templates' days (template.onDays) with the time off applied: an absent provider's slots and cells go, or become the covering
     provider's when they work that template; a covering provider who works the absent one's template doesn't also keep their own. */
  function withExceptions(D, EX) {
    if (!EX || !EX.length) return D;
    const memo = new Map(), on = d => {
      if (memo.has(d)) return memo.get(d);
      const conv = list => list.flatMap(s => {
        if (EX.some(e => e.useTpl && e.covKey === s.pkey && e.from <= d && d <= e.to && !s.covering)) return [];   // replaced by the pre-fill
        const e = exFor(EX, s.pkey, d); if (!e) return [s];
        return e.useTpl ? [Object.assign({}, s, { provider: e.cov, pkey: e.covKey, who: provName(e.cov) || e.cov, covering: e.provider })] : [];
      });
      const v = { slots: conv(D.slotsOn(d)), cells: conv(D.cellsOn(d)) }; memo.set(d, v); return v;
    };
    /* the absent provider's visits and holds those days sit in the covering provider's pre-filled slots */
    const remap = (pk, d) => { const e = exFor(EX, pk, d); return e && e.useTpl ? e.covKey : pk; };
    return Object.assign({}, D, { slotsOn: d => on(d).slots, cellsOn: d => on(d).cells, remap });
  }
  /* IVK visits still booked with a provider on a day they're off (cancelled ones don't count) */
  /* (a clinic closure: every IVK visit and hold those days) */
  const exClashes = (e, A, H) => { const who = n => e.closed || TM.provKey(n) === e.pkey;
    return A.filter(a => who(a.provider) && a.date >= e.from && a.date <= e.to && !isCancelled(a))
      .concat((H || []).filter(h => h.status !== "released" && who(h.provider) && h.date >= e.from && h.date <= e.to && !onList(h, A)).map(h => Object.assign({ hold: true }, h))); };
  /* every visit or hold on a closure day still to come, closure by closure */
  const closureClashes = (EX, A, H, today) => (EX || []).filter(e => e.closed && e.to >= today).map(e => ({ e, list: exClashes(e, A, H).filter(a => a.date >= today) })).filter(x => x.list.length);
  /* the IVK visits the Appt List tab keeps; everything else in the report goes to tblIvkOther */
  const isIvkVisit = ev => /\b(iv ketamine|ivk)\b/i.test(ev) && !/consult/i.test(ev);
  /* tblIvkFlags: MRN | Patient | RequiresPostCheck | By | Updated. A provider asks for a post check on a patient's follow-ups. */
  const FLAG_COLS = ["MRN", "Patient", "RequiresPostCheck", "By", "Updated"];
  const flagRows = values => { const m = new Map(); byHeader(values).rows.forEach(x => { const mrn = mrnKey(x.get("mrn")); if (mrn) m.set(mrn, /^(true|yes|1)$/i.test(txt(x.get("requirespostcheck")))); }); return m; };
  /* tblIvkPlans: MRN | Patient | Start | PerWeek | Weeks | Avail | Note | By | Updated, one row per patient being planned.
     PerWeek/Weeks blank = the clinic's index protocol; Avail blank = the tracker's AM/PM preferences, "any" = any time. */
  const PLAN_COLS = ["MRN", "Patient", "Start", "PerWeek", "Weeks", "Avail", "Note", "By", "Updated"];
  function planRows(values) {
    const m = new Map();
    byHeader(values).rows.forEach(x => { const mrn = mrnKey(x.get("mrn")); if (!mrn) return;
      m.set(mrn, { mrn, code: txt(x.get("patient")), start: toSerial(x.get("start")), perWeek: num(x.get("perweek")), weeks: num(x.get("weeks")),
        avail: txt(x.get("avail")), note: txt(x.get("note")), by: txt(x.get("by")) }); });
    return m;
  }
  /* tblIvkHolds: a time kept for a patient until the desk books it in NextGen. Status held, booked or released. */
  const HOLD_COLS = ["Id", "MRN", "Patient", "Date", "Time", "Provider", "Treatment", "Status", "PlacedBy", "PlacedAt", "DoneBy", "DoneAt"];
  function holdRows(values) {
    return byHeader(values).rows.map(x => { const g = x.get, at = Date.parse(txt(g("placedat"))), done = Date.parse(txt(g("doneat")));
      return { id: txt(g("id")), mrn: mrnKey(g("mrn")), code: txt(g("patient")), date: toSerial(g("date")), time: toMin(g("time")), provider: txt(g("provider")),
        n: num(g("treatment")), status: txt(g("status")).toLowerCase() || "held", by: txt(g("placedby")), at: isNaN(at) ? null : at, doneBy: txt(g("doneby")),
        doneAt: isNaN(done) ? null : done }; })
      .filter(h => h.id && h.date != null && h.time != null);
  }
  /* NextGen has the visit (same patient, day and time), so the hold is done whatever its status says */
  const onList = (h, A) => A.some(a => a.mrn === h.mrn && a.date === h.date && a.time === h.time && !isCancelled(a));
  /* Does the hold still keep its time? Not once released or on the list; and a hold marked booked stops once an appointment list
     uploaded after the booking is in (listAt, the Appt List's Last Updated serial), since NextGen is then the record, even if the
     visit went in at another time. */
  const msSerial = ms => { const d = new Date(ms); return serialOf(d) + (d.getHours() * 60 + d.getMinutes()) / 1440; };
  const holding = (h, A, listAt) => h.status !== "released" && !onList(h, A) && !(h.status === "booked" && listAt != null && h.doneAt != null && msSerial(h.doneAt) < listAt);
  /* held: waiting to be booked; late: not confirmed booked within HoldMin, so someone has to act (red); booked; released */
  /* A hold marked booked is only confirmed by NextGen: onlist when the appointment list has the visit, missing when a list uploaded
     after the booking (listAt) still doesn't, pending until then (the desk uploads a fresh report) */
  const holdConfirm = (h, A, listAt) => onList(h, A) ? "onlist" : h.status !== "booked" ? null
    : listAt != null && h.doneAt != null && msSerial(h.doneAt) < listAt ? "missing" : "pending";
  function holdState(h, A, nowMs, cfg) {
    if (h.status === "released") return "released";
    if (h.status === "booked" || onList(h, A)) return "booked";
    return h.at != null && nowMs - h.at > +cfg.HoldMin * 60000 ? "late" : "held";
  }
  /* One day on the templates. X = { slotsOn(d), cellsOn(d), appts, others, holds, cfg, chairs, listAt, needsPost(mrn, event) }:
     slotsOn gives each provider's infusion starts that day (template.startsOn), cellsOn every template cell.
     An IVK visit or hold takes the start at its time; an index visit (a hold, a treatment start, an Induction patient) or a patient who
     requires a post check also takes the post check PostCheckMin later. A visit from tblIvkOther takes the cell it starts in.
     slots: the starts with by (who has it), postBy (who has its post check), full (all chairs in use during its 2 hours).
     cells: every template cell with by, for the tiles. */
  function daySlots(d, X) {
    if (X._memo && X._memo.has(d)) return X._memo.get(d);
    const out = daySlots0(d, X); if (X._memo) X._memo.set(d, out); return out;
  }
  /* X with a per-day cache, for the long scans (eligible starts, capacity); holds must not change while it's used */
  const memoX = X => Object.assign({}, X, { _memo: new Map() });
  function daySlots0(d, X) {
    const vm = +X.cfg.VisitMin, pm = +X.cfg.PostCheckMin, key = X.remap ? n => X.remap(TM.provKey(n), d) : TM.provKey;   // remap: a covering provider works an absent one's template
    const cells = (X.cellsOn ? X.cellsOn(d) : []).map(c => ({ ...c, date: d, by: null }));
    /* the provider's cell a time falls in (a 9:15 visit is in the 9:00 half hour) */
    const cellAt = (pk, t) => cells.find(c => c.pkey === pk && c.time <= t && t < c.time + (c.len || 30));
    const B = X.appts.filter(a => a.date === d && a.time != null && !isCancelled(a) && !/no ?show/i.test(a.status))
      /* a booked visit that was one of the patient's holds is an index visit, whatever NextGen calls it */
      .map(a => ({ mrn: a.mrn, code: a.code, time: a.time, provider: a.provider, pkey: key(a.provider), kind: "appt", event: a.event,
        post: !!(X.needsPost && X.needsPost(a.mrn, a.event)) || (X.holds || []).some(h => h.mrn === a.mrn && h.date === d && h.time === a.time && h.status !== "released") }))
      .concat((X.holds || []).filter(h => h.date === d && holding(h, X.appts, X.listAt)).map(h => ({ mrn: h.mrn, code: h.code, time: h.time, provider: h.provider, pkey: key(h.provider), kind: "hold", h, post: true })))
      /* a late cancel keeps its time: the schedule shows it cancelled, not open (X.late, the late-cancel visit keys) */
      .concat(X.late && X.late.size ? X.appts.filter(a => a.date === d && a.time != null && isCancelled(a) && X.late.has(visitKey(a.date, a.mrn, a.time)))
        .map(a => ({ mrn: a.mrn, code: a.code, time: a.time, provider: a.provider, pkey: key(a.provider), kind: "late", event: a.event, post: false })) : []);
    const slots = (X.slotsOn ? X.slotsOn(d) : []).map(s => ({ ...s, date: d, by: null, postBy: null, full: false }));
    for (const b of B) {
      const s = slots.find(x => !x.by && x.pkey === b.pkey && x.time <= b.time && b.time < x.time + (x.len || 30)); if (s) s.by = b;
      const c = cellAt(b.pkey, b.time); if (c && !c.by) c.by = b;
      /* where its post check lands (the shaded tile on the Schedule and the Access week) */
      if (b.post) { b.postAt = (s ? s.time : b.time) + pm; const pc = cellAt(b.pkey, b.postAt); if (pc && !pc.by && (pc.use === "post" || pc.use === "flexpost")) pc.by = Object.assign({}, b, { kind: b.kind === "hold" ? "holdpost" : "post" }); }
      /* booked off the template: the provider has a template but no IVK slot at that time, or a treatment start sits in a
         follow-up-only slot (follow-ups may take a start/index slot) */
      if (b.kind === "appt" && X.hasGrid && X.hasGrid(b.pkey)) b.miss = !c || !(c.use === "start" || c.use === "followup") || (/start/i.test(b.event) && c.use === "followup");
    }
    for (const o of (X.others || []).filter(o => o.date === d && !isCancelled(o) && !/no ?show/i.test(o.status) && !/\bhold\b/i.test(o.event))) {
      const pk = key(o.provider), c = cells.find(x => x.pkey === pk && !x.by && o.time >= x.time && o.time < x.time + (x.len || 30));
      if (c) c.by = { kind: "other", event: o.event, time: o.time, provider: o.provider };
    }
    const inUse = t => Math.max(...[t, ...B.map(b => b.time).filter(x => x > t && x < t + vm)].map(p => B.filter(b => b.time <= p && p < b.time + vm).length));
    for (const s of slots) {
      if (!s.by) { const c = cells.find(x => x.pkey === s.pkey && x.time === s.time); if (c && c.by && c.by.kind === "other") s.by = c.by; }
      if (s.postUse) s.postBy = cells.find(x => x.pkey === s.pkey && x.time === s.post)?.by || null;
      s.full = !s.by && inUse(s.time) >= X.chairs;
      s.open = !s.by && !s.full;
      s.openIndex = s.open && s.index && !s.postBy;
    }
    return { slots, cells, bookings: B };
  }
  /* open starts a patient who requires a post check could book: any IVK start (index or follow-up slot) whose end lines up with a
     free IVK or flex post check, from a day on, up to n */
  function postCheckSlots(X, from, n, days = 42) {
    const out = [];
    for (let d = from; d < from + days && out.length < n; d++) { const wd = weekday(d); if (wd < 1 || wd > 5) continue;
      for (const s of daySlots(d, X).slots) if (s.open && s.postUse && !s.postBy) out.push(s); }
    return out.slice(0, n);
  }
  /* availability Mon-Fri as [{ am, pm }], or null for any time. AM is a start before noon. */
  function availOf(plan, p) {
    const t = txt(plan?.avail);
    if (/^any/i.test(t)) return null;
    if (t) {
      const a = DAYS5.map(() => ({ am: false, pm: false }));
      t.split(/[,;]+/).forEach(w => { const d = dayNum(w.trim()), h = /pm/i.test(w) ? "pm" : /am/i.test(w) ? "am" : null; if (d && h) a[d - 1][h] = true; });
      return a.some(x => x.am || x.pm) ? a : null;
    }
    const pr = p?.prefs; return pr && pr.some(x => x.am || x.pm) ? pr.map(x => ({ am: x.am, pm: x.pm })) : null;
  }
  const availText = a => a ? a.flatMap((x, i) => [x.am ? `${DAYS5[i]} AM` : "", x.pm ? `${DAYS5[i]} PM` : ""]).filter(Boolean).join(", ") || "any" : "any";
  const fits = (a, wd, t) => !a || !!(a[wd - 1] && (t < 720 ? a[wd - 1].am : a[wd - 1].pm));

  /* Proposes a patient's index visits. o = { start, today, perWeek, weeks, minGap, maxGap, avail, X, fixed, pins }.
     Treatment 1 is the first open index start (a Start/Index slot whose post check is free) on or after the start date that fits,
     or a later one in that first week when starting there keeps every visit within maxGap (each day later costs a little, so it
     only moves when the first one would force a long gap). Week 1 is the 7 days from treatment 1 (Wednesday to Tuesday when
     treatment 1 is a Wednesday), and each week gets perWeek visits. Visits are at least minGap days apart (never back to back), no
     more than maxGap apart when the slots allow it, and as evenly spaced as they can be.
     fixed: the patient's visits already on the list, kept as they are. pins: { n: slot } picked by hand (n from 0). */
  function planIndex(o) {
    const N = Math.max(1, Math.round(o.perWeek * o.weeks)), ideal = 7 / o.perWeek, pins = o.pins || {};
    const from = Math.max(o.start ?? o.today + 1, o.today + 1), endD = from + 7 * (N + 1);
    const fixed = (o.fixed || []).filter(f => f.date >= from).sort((a, b) => a.date - b.date), fixedD = fixed.map(f => f.date);
    const cand = new Map();
    for (let d = from; d <= endD; d++) {
      const wd = weekday(d); if (wd < 1 || wd > 5) continue;
      if (o.away && awayOn(o.away, d)) continue;
      const f = fixed.find(x => x.date === d); if (f) { cand.set(d, [{ ...f, fixed: true }]); continue; }
      const L = daySlots(d, o.X).slots.filter(s => s.openIndex && fits(o.avail, wd, s.time));
      if (L.length) cand.set(d, L);
    }
    for (const p of Object.values(pins)) if (p && !cand.has(p.date)) cand.set(p.date, [p]);
    const days = [...cand.keys()].sort((a, b) => a - b), first = pins[0] ? pins[0].date : days[0];
    if (first == null) return { items: [], N, cand, notes: ["No open index start fits this patient in the weeks after the start date. Check the template, the post checks and their availability."] };
    const starts = pins[0] ? [first] : days.filter(d => d < first + 7 && !fixedD.some(f => f < d));
    let best = null, budget = 300000;
    for (const t0 of starts) {
      const wk = d => Math.floor((d - t0) / 7);
      const go = (seq, c) => {
        if (--budget < 0 || (best && best.seq.length === N && c >= best.c)) return;
        const n = seq.length, lastD = seq[n - 1], kids = [];
        if (n < N) for (const d of days) {
          if (d <= lastD) continue; const g = d - lastD;
          if (g > o.maxGap + 4 || fixedD.some(f => f > lastD && f < d)) break;
          if (g < o.minGap || (pins[n] && pins[n].date !== d)) continue;
          const w = wk(d), inW = seq.filter(x => wk(x) === w).length;
          kids.push([d, (g - ideal) ** 2 + (g > o.maxGap ? 50 : 0) + (inW >= o.perWeek ? 100 : 0) + (w >= o.weeks ? 100 : 0)]);
        }
        kids.sort((a, b) => a[1] - b[1]).forEach(([d, add]) => { seq.push(d); go(seq, c + add); seq.pop(); });
        if (!kids.length && (!best || n > best.seq.length || (n === best.seq.length && c < best.c))) best = { seq: seq.slice(), c };
      };
      go([t0], (t0 - first) * 3);
    }
    const seq = best.seq, t1 = seq[0], wk = d => Math.floor((d - t1) / 7);
    /* the slot on each day: the pinned one, a visit already booked, then a start ending on an IVK post check before one ending on a
       flex post check (flex post checks are meant for TMS follow-ups and tele), closest to treatment 1's time */
    const pick = (d, n, t) => pins[n] && pins[n].date === d ? pins[n] : cand.get(d).slice().sort((a, b) => (b.fixed ? 1 : 0) - (a.fixed ? 1 : 0)
      || (a.postUse === "flexpost") - (b.postUse === "flexpost") || (t == null ? 0 : Math.abs(a.time - t) - Math.abs(b.time - t)) || a.time - b.time)[0];
    const s1 = pick(t1, 0, null), pc = +o.X.cfg.PostCheckMin;
    const items = seq.map((d, n) => { const s = n === 0 ? s1 : pick(d, n, s1.time), g = n ? d - seq[n - 1] : null;
      return { n: n + 1, week: wk(d) + 1, date: d, time: s.time, provider: s.provider || "", pkey: s.pkey || TM.provKey(s.provider), type: s.fixed ? "Booked" : s.type,
        postUse: s.postUse || null, postType: s.postType || "", post: s.time + pc, gap: g,
        fixed: !!s.fixed, pinned: !!pins[n], long: g != null && g > o.maxGap }; });
    const notes = [];
    if (items.length < N) notes.push(`Only ${items.length} of ${N} visits fit: no open slot within ${o.maxGap + 4} days after ${fmtD(seq[seq.length - 1])} for this patient.`);
    for (let w = 1; w <= o.weeks; w++) { const k = items.filter(x => x.week === w).length; if (items.length === N && k !== o.perWeek) notes.push(`Week ${w} has ${k} visit${k === 1 ? "" : "s"} instead of ${o.perWeek}.`); }
    if (items.some(x => x.week > o.weeks)) notes.push(`Runs into week ${Math.max(...items.map(x => x.week))}.`);
    return { items, N, cand, notes, t1 };
  }
  /* other open slots treatment n (from 0) could move to, still minGap+ days from the visits either side */
  function altSlots(res, n, minGap) {
    const it = res.items, prev = it[n - 1], next = it[n + 1], cur = it[n], out = [];
    for (const [d, L] of [...res.cand.entries()].sort((a, b) => a[0] - b[0])) {
      if ((prev && d - prev.date < minGap) || (next && next.date - d < minGap) || (!prev && next && d >= next.date)) continue;
      for (const s of L) if (!s.fixed && !(cur && s.date === cur.date && s.time === cur.time && sameProv(s.provider, cur.provider))) out.push(s);
    }
    return out;
  }
  /* How many more patients could start index: plan one patient free any time, take their visits, and repeat while a full plan
     still starts within weeks. o = planIndex's options without the patient. Also counts the open index starts in that window. */
  /* A full index fits: all N visits, perWeek in each of the weeks, none past the protocol's weeks, no gap over maxGap */
  const fullPlan = r => r.items.length === r.N && !r.items.some(x => x.long) && !r.notes.length;
  /* Index-eligible starts: open index starts from a day on where a full plan (the first 6 visits, in the protocol's windows) can
     still be made from them on what's left. o = planIndex's options without the patient. Up to n, within days. */
  function eligibleStarts(o, from, n, days) {
    const X = o.X._memo ? o.X : memoX(o.X), out = [];
    for (let d = from; d < from + days && out.length < n; d++) {
      const wd = weekday(d); if (wd < 1 || wd > 5) continue;
      for (const s of daySlots(d, X).slots.filter(x => x.openIndex)) {
        const r = planIndex(Object.assign({}, o, { start: d, avail: null, fixed: [], pins: { 0: s }, X }));
        if (fullPlan(r)) { out.push({ slot: s, plan: r }); if (out.length >= n) break; }
      }
    }
    return out;
  }
  function capacity(o, weeks) {
    const until = o.today + 7 * weeks; let n = 0, holds = (o.X.holds || []).slice();
    /* open counts only starts a full index can still be planned from */
    const open = eligibleStarts(o, o.today + 1, 500, 7 * weeks).length;
    for (; n < 25; n++) {
      /* the first start a full index fits from, with the patients counted so far taking their visits */
      const e = eligibleStarts(Object.assign({}, o, { X: memoX(Object.assign({}, o.X, { holds, _memo: null })) }), o.today + 1, 1, 7 * weeks)[0], r = e && e.plan;
      if (!r || r.items[0].date > until) break;
      holds = holds.concat(r.items.map((x, i) => ({ id: `cap${n}.${i}`, mrn: "cap" + n, code: "", date: x.date, time: x.time, provider: x.provider, status: "held" })));
    }
    return { starts: n, open };
  }

  /* The Index Planner tab: dates across the "Date" row, half hours down column A. Each cell gets the codes booked in that half
     hour, then "HOLD CODE" for holds NextGen doesn't have yet. layout = { rows: [[row, minutes]], cols: [[col, serial]] } */
  function plannerLayout(values) {
    const V = values || [], r = V.findIndex(row => /^date$/i.test(txt(row[0]))); if (r < 0) return null;
    const isT = v => (typeof v === "number" && v >= 0 && v < 1) || /^\d{1,2}:\d\d/.test(txt(v)), rows = [], cols = [];
    for (let i = r + 1; i < V.length && isT(V[i][0]); i++) rows.push([i, toMin(V[i][0])]);
    for (let c = 1; c < V[r].length; c++) { const d = toSerial(V[r][c]); if (d != null) cols.push([c, d]); }
    return rows.length && cols.length ? { dateRow: r, rows, cols } : null;
  }
  function plannerGrid(L, A, holds, listAt) {
    return L.rows.map(([, t], i) => { const t2 = L.rows[i + 1] ? L.rows[i + 1][1] : t + 30;
      return L.cols.map(([, d]) => { const inCell = x => x.date === d && x.time != null && x.time >= t && x.time < t2;
        return A.filter(a => inCell(a) && !isCancelled(a)).map(a => a.code || "HOLD")
          .concat(holds.filter(h => inCell(h) && holding(h, A, listAt)).map(h => `HOLD ${h.code}`)).join(", "); }); });
  }
  /* ---------- Schedule ---------- */
  /* one Sunday-to-Saturday week of IVK visits by weekday and time. Rows are every 30 min from the first to the last visit (9:00 AM to 3:30 PM at least). */
  function week(A, start, P, extra) {
    const days = [1, 2, 3, 4, 5].map(i => start + i), inWeek = A.filter(a => a.date >= start && a.date < start + 7);
    const times = inWeek.map(a => a.time).concat(extra || []).filter(t => t != null);
    const lo = Math.min(540, ...times), hi = Math.max(930, ...times), rows = [];
    for (let t = Math.floor(lo / 30) * 30; t <= hi; t += 30) rows.push(t);
    const cell = (d, t) => inWeek.filter(a => a.date === d && a.time != null && a.time >= t && a.time < t + 30).map(a => ({ ...a, p: P?.get(a.mrn) || null }));
    return { start, days, rows, cell, count: inWeek.filter(a => !isCancelled(a)).length, byDay: days.map(d => inWeek.filter(a => a.date === d && !isCancelled(a)).length) };
  }

  /* ---------- per patient: payments flags, Oakland, availability, waitlist (tblIvkPatient) ---------- */
  /* MRN | Patient | CardOnFile | OkToCollect | Oakland | OakTasked | Avail | AvailEffective | AvailExpires | AvailNote | Away |
     WaitRank | WaitingOn | Appeal | WaitNote | By | Updated. Blank Oakland = decide from the home clinic. */
  const PT_COLS = ["MRN", "Patient", "CardOnFile", "OkToCollect", "Oakland", "OakTasked", "Avail", "AvailEffective", "AvailExpires", "AvailNote", "Away",
    "WaitRank", "WaitingOn", "Appeal", "WaitNote", "By", "Updated"];
  const yes = v => /^(y|yes|true|1|x)$/i.test(txt(v)), no = v => /^(n|no|false|0)$/i.test(txt(v));
  function ptRows(values) {
    const m = new Map();
    byHeader(values).rows.forEach(x => { const g = x.get, mrn = mrnKey(g("mrn")); if (!mrn) return;
      m.set(mrn, { mrn, card: yes(g("cardonfile")), okCollect: yes(g("oktocollect")), oak: yes(g("oakland")) ? true : no(g("oakland")) ? false : null,
        oakTasked: toSerial(g("oaktasked")), avail: availGrid(g("avail")), availEff: toSerial(g("availeffective")), availExp: toSerial(g("availexpires")),
        availNote: txt(g("availnote")), away: awayList(g("away")), rank: num(g("waitrank")), waitOn: txt(g("waitingon")).toLowerCase(), appeal: yes(g("appeal")),
        waitNote: txt(g("waitnote")), by: txt(g("by")) }); });
    return m;
  }

  /* Availability, as the patient availability form asks it: each weekday morning and afternoon is Preferred, Mostly available,
     Sometimes available or Not available. Saved as "Mon AM:P; Mon PM:S; ...". Preferred and Mostly count as available for planning
     and go to the tracker's AM/PM columns as 1. */
  const AV_LEVELS = [["P", "Preferred"], ["M", "Mostly available"], ["S", "Sometimes available"], ["N", "Not available"]];
  const avName = k => (AV_LEVELS.find(l => l[0] === k) || [])[1] || "Not set";
  function availGrid(v) {
    const a = DAYS5.map(() => ({ am: "", pm: "" })), s = txt(v); if (!s) return null;
    s.split(/[;,]+/).forEach(w => { const m = w.trim().match(/^(\w{3})\w*\s*(AM|PM)\s*:\s*([PMSN])$/i); if (!m) return; const d = dayNum(m[1]); if (d) a[d - 1][m[2].toLowerCase()] = m[3].toUpperCase(); });
    return a.some(x => x.am || x.pm) ? a : null;
  }
  const availStr = a => (a || []).flatMap((x, i) => ["am", "pm"].filter(h => x[h]).map(h => `${DAYS5[i]} ${h.toUpperCase()}:${x[h]}`)).join("; ");
  const canCome = k => k === "P" || k === "M";
  /* the tracker's preference columns (AM, PM, AM2, PM3, ...) from a grid: 1 for Preferred or Mostly, 0 otherwise */
  /* with the tracker's header row, the cells go by position ("@12" = the 13th column, see excel.js saveRow); without it, by name */
  const prefCells = (a, head) => { const v = (i, h) => canCome(a[i][h]) ? 1 : 0;
    if (head) return Object.fromEntries(prefCols(head).map((c, k) => [`@${c}`, v(k >> 1, k % 2 ? "pm" : "am")]));
    return Object.fromEntries(PREF_DAYS.flatMap(([, am, pm], i) => [[am, v(i, "am")], [pm, v(i, "pm")]])); };
  /* dates away: "2026-11-02|2026-11-06|Vacation; ..." */
  const awayList = v => txt(v).split(";").map(w => { const [f, t, ...n] = w.split("|"); const from = fromIso(f), to = fromIso(t) ?? from; return from == null ? null : { from, to: Math.max(from, to), note: n.join("|").trim() }; }).filter(Boolean).sort((a, b) => a.from - b.from);
  const awayStr = L => (L || []).filter(w => w.from != null).map(w => `${isoOf(w.from)}|${isoOf(w.to ?? w.from)}|${txt(w.note).replace(/[;|]/g, ",")}`).join("; ");
  const awayOn = (L, d) => (L || []).find(w => w.from <= d && d <= w.to) || null;

  /* ---------- payments (tblIvkPay) ---------- */
  /* Id | Date | MRN | Patient | Key | Kind | Amount | Method | Note | By | At, one row per event:
     collected / missed  the payment due at a visit (Key = the visit), collected at the desk or not
     paid / waived       a missed payment settled later, or waived (Key = the missed visit)
     reminder            a balance reminder sent (Method: In-person or Luma) */
  const PAY_COLS = ["Id", "Date", "MRN", "Patient", "Key", "Kind", "Amount", "Method", "Note", "By", "At"];
  const REMIND_METHODS = ["In-person", "Luma"];
  const money = n => n == null ? "" : `$${(+n).toFixed(2).replace(/\.00$/, "")}`;
  const amt = v => { const n = num(txt(v).replace(/[$,\s]/g, "")); return n == null ? 0 : n; };
  function payRows(values) {
    return byHeader(values).rows.map(x => { const g = x.get;
      return { id: txt(g("id")), date: toSerial(g("date")), mrn: mrnKey(g("mrn")), code: txt(g("patient")), key: txt(g("key")), kind: txt(g("kind")).toLowerCase(),
        amount: amt(g("amount")), method: txt(g("method")), note: txt(g("note")), by: txt(g("by")), at: txt(g("at")) }; })
      .filter(r => r.mrn && r.kind);
  }
  /* what an IVK visit costs the patient: the provider visit's copay (99214/99215, the tracker's Copay) plus the infusion
     responsibility (the tracker's Infusion), each collected on its own */
  const dueOf = p => p ? { copay: amt(p.copay), infusion: amt(p.infusion), total: amt(p.copay) + amt(p.infusion) } : { copay: 0, infusion: 0, total: 0 };
  const PARTS = [["prov", "Provider visit"], ["inf", "Infusion"]];
  const partKey = (key, part) => `${key}|${part}`;
  /* the parts of a visit with something due: [["prov", 40], ["inf", 300]] */
  const dueParts = p => { const d = dueOf(p); return [["prov", d.copay], ["inf", d.infusion]].filter(x => x[1] > 0); };
  /* the visit's own answer: the last collected or missed row for its key. With a part, that part's row, or a row for the whole
     visit (saved before the parts were split) */
  const payOf = (pays, key, part) => pays.filter(r => (r.key === key || (part && r.key === partKey(key, part))) && (r.kind === "collected" || r.kind === "missed")).pop() || null;
  /* every part due at this visit answered */
  const visitAnswered = (pays, key, p) => dueParts(p).every(([part]) => !!payOf(pays, key, part));
  /* A patient's money: every missed visit with what's been paid or waived against it, the balance, the last reminder. */
  function ledger(mrn, pays) {
    const mine = pays.filter(r => r.mrn === mrn), settled = k => mine.filter(r => r.key === k && (r.kind === "paid" || r.kind === "waived")).reduce((n, r) => n + r.amount, 0);
    const missed = mine.filter(r => r.kind === "missed" && payOf(mine, r.key) === r).map(r => Object.assign({}, r, { left: Math.max(0, r.amount - settled(r.key)) }));
    /* a payment with no visit (paid on account) comes off the oldest balance */
    let loose = mine.filter(r => r.kind === "paid" && !r.key).reduce((n, r) => n + r.amount, 0);
    for (const r of missed.sort((a, b) => (a.date ?? 0) - (b.date ?? 0))) { const x = Math.min(loose, r.left); r.left -= x; loose -= x; }
    const open = missed.filter(r => r.left > 0.004), reminders = mine.filter(r => r.kind === "reminder").sort((a, b) => (a.date ?? 0) - (b.date ?? 0));
    /* a visit's answer can be changed after check-out (a new row for the same key), so only its latest answer counts as collected */
    return { open, balance: Math.round(open.reduce((n, r) => n + r.left, 0) * 100) / 100, lastReminder: reminders[reminders.length - 1] || null, reminders,
      collected: mine.filter(r => (r.kind === "collected" && payOf(mine, r.key) === r) || r.kind === "paid").reduce((n, r) => n + r.amount, 0) };
  }
  /* past IVK visits with something due and no answer yet (since the payment log started, PayStart), newest first */
  function toConfirm(p, A, pays, today, from) {
    if (!p || dueOf(p).total <= 0) return [];
    return A.filter(a => a.mrn === p.mrn && a.date < today && a.date >= (from ?? today) && !isCancelled(a) && !/no ?show/i.test(a.status))
      .filter(a => !visitAnswered(pays, visitKey(a.date, a.mrn, a.time), p)).sort((a, b) => b.date - a.date);
  }

  /* ---------- no-shows and late cancels (tblIvkMissed) ---------- */
  /* Key | Date | Time | MRN | Patient | Provider | Kind | Exempt | Reason | Source | By | Updated, one row per visit the desk or an
     upload said something about. Kind: late (a late cancel), review (cancelled close to the visit: the desk says whether it was
     late), notlate (the desk said it wasn't), noshow (a NextGen no-show the desk exempted). Exempt: clinical (a clinical reason, in
     Reason). A visit cancelled on its own day is always late. Kaiser and VA patients are exempt from fees; that's worked out from
     the payer, never stored. No-shows come from NextGen's status. Nothing here writes a fee or a balance: every one is counted. */
  const MISS_COLS = ["Key", "Date", "Time", "MRN", "Patient", "Provider", "Kind", "Exempt", "Reason", "Source", "By", "Updated"];
  function missRows(values) {
    const m = new Map();
    byHeader(values).rows.forEach(x => { const g = x.get, key = txt(g("key")), date = toSerial(g("date")); if (!key || date == null) return;
      m.set(key, { key, date, time: toMin(g("time")), mrn: mrnKey(g("mrn")), code: txt(g("patient")), provider: txt(g("provider")), kind: txt(g("kind")).toLowerCase(),
        exempt: txt(g("exempt")).toLowerCase(), reason: txt(g("reason")), source: txt(g("source")).toLowerCase(), by: txt(g("by")) }); });
    return m;
  }
  /* Kaiser (KP) and VA (VA, TriWest, Community Care, CHAMPVA) patients pay no no-show or late-cancel fee */
  const KPVA = /kaiser|\bkp\b|\bva\b|veteran|triwest|community care|champva/i;
  const isKpVa = (p, pf) => [p?.ins1, p?.ins2, pf?.prov?.payer, pf?.inf?.payer, pf?.sohar?.payer].some(v => KPVA.test(txt(v)));
  /* every no-show and late cancel (and the cancellations waiting for the desk to say), newest first, each with why it's exempt:
     kpva, clinical or "" (no exemption). o = { appts, MR (missRows), PM, PF, from, to } */
  function missList(o) {
    const out = new Map(), MR = o.MR || new Map(), lo = o.from ?? -Infinity, hi = o.to ?? Infinity;
    for (const a of o.appts || []) if (/no ?show/i.test(a.status) && a.date >= lo && a.date <= hi) { const key = visitKey(a.date, a.mrn, a.time);
      out.set(key, { key, date: a.date, time: a.time, mrn: a.mrn, code: a.code, provider: a.provider, kind: "noshow", source: "nextgen" }); }
    for (const r of MR.values()) if ((r.kind === "late" || r.kind === "review") && !out.has(r.key) && r.date >= lo && r.date <= hi) out.set(r.key, Object.assign({}, r));
    for (const m of out.values()) { const p = o.PM?.get(m.mrn), r = MR.get(m.key);
      m.code = m.code || p?.code || r?.code || ""; m.exempt = isKpVa(p, o.PF?.get(m.mrn)) ? "kpva" : r?.exempt === "clinical" ? "clinical" : ""; m.reason = r?.reason || ""; }
    return [...out.values()].sort((a, b) => b.date - a.date || (b.time ?? 0) - (a.time ?? 0));
  }
  /* the visits a late cancel still holds on the schedule: shown as cancelled, never as an open slot or a gap */
  const lateKeys = MR => new Set([...(MR || new Map()).values()].filter(r => r.kind === "late").map(r => r.key));
  /* Cancellations a new appointment list shows that the old one didn't (cancelled, rescheduled, or gone from a list that still covers
     that day), for visits from LateCancelDays back to LateCancelDays ahead. sameDay: the visit is today and the old list was
     uploaded today, so it was cancelled on its own day (a late cancel, no question). The rest go to the desk. */
  function newCancels(oldA, newA, today, cfg, oldAt) {
    const days = Math.max(0, +cfg.LateCancelDays || 0), live = a => !isCancelled(a) && !/no ?show/i.test(a.status), from = newA.reduce((m, a) => a.date < m ? a.date : m, Infinity);
    const now = new Map(newA.map(a => [visitKey(a.date, a.mrn, a.time), a])), sameUpload = oldAt != null && Math.floor(oldAt) === today;
    return oldA.filter(a => live(a) && a.time != null && a.date >= today - days && a.date <= today + days).map(a => ({ a, key: visitKey(a.date, a.mrn, a.time), n: now.get(visitKey(a.date, a.mrn, a.time)) }))
      .filter(x => x.n ? isCancelled(x.n) : x.a.date >= from)
      .map(x => ({ key: x.key, a: x.a, sameDay: x.a.date === today && sameUpload }));
  }
  /* What no-shows and late cancels cost access: each past week (Sunday to Saturday) with visits on the list, the template's IVK
     slots, what was booked in them (kept + no-shows + late cancels), what was lost, and the slots that ended the week open.
     rate: the share of booked visits lost, over those weeks, to project how many of the weeks ahead's booked visits won't happen.
     o = { appts, X (templates with time off), missed (missList), today, weeks } */
  function missImpact(o) {
    const w0 = weekStart(o.today), W = Math.max(1, +o.weeks || 8), L = o.missed || [], late = new Set(L.filter(m => m.kind === "late").map(m => m.key)), rows = [];
    for (let i = W; i >= 1; i--) {
      const a = w0 - 7 * i, b = a + 7, inW = (o.appts || []).filter(x => x.date >= a && x.date < b);
      if (!inW.length) continue;
      const kept = inW.filter(x => !isCancelled(x) && !/no ?show/i.test(x.status) && !late.has(visitKey(x.date, x.mrn, x.time))).length;
      const ns = L.filter(m => m.kind === "noshow" && m.date >= a && m.date < b).length, lc = L.filter(m => m.kind === "late" && m.date >= a && m.date < b).length;
      let slots = null; if (o.X) { slots = 0; for (let d = a + 1; d < a + 6; d++) slots += (o.X.slotsOn ? o.X.slotsOn(d) : []).length; }
      const booked = kept + ns + lc;
      rows.push({ start: a, slots, booked, kept, noshow: ns, late: lc, lost: ns + lc, endOpen: slots == null ? null : Math.max(0, slots - kept), rate: booked ? (ns + lc) / booked : 0 });
    }
    const bk = rows.reduce((n, r) => n + r.booked, 0), lost = rows.reduce((n, r) => n + r.lost, 0);
    return { rows, booked: bk, lost, noshow: rows.reduce((n, r) => n + r.noshow, 0), late: rows.reduce((n, r) => n + r.late, 0), rate: bk ? lost / bk : 0 };
  }
  /* per patient over the missed list: how many no-shows and late cancels, and how many of them have no exemption */
  function missByPatient(L) {
    const m = new Map();
    for (const x of L) { if (x.kind === "review") continue; const o = m.get(x.mrn) || { mrn: x.mrn, code: x.code, noshow: 0, late: 0, open: 0, last: null, exempt: x.exempt === "kpva" ? "kpva" : "" };
      o[x.kind === "noshow" ? "noshow" : "late"]++; if (!x.exempt) o.open++; if (!o.last || x.date > o.last) o.last = x.date; m.set(x.mrn, o); }
    return [...m.values()].sort((a, b) => (b.noshow + b.late) - (a.noshow + a.late) || (b.last ?? 0) - (a.last ?? 0));
  }

  /* ---------- patient profile (tblIvkProfile) ---------- */
  /* One row per patient: who pays each part of the visit (insurance or self pay, and the payer) and the latest Sohar benefits.
     The amounts due and the payer names also go to the Master Patient Tracker (Copay, Infusion, Primary Encounter Insurance,
     Infusion Insurance), which stays the record the clinic reads. BidfPrinted: the day the desk printed the BIDF from NextGen. */
  const PROFILE_COLS = ["MRN", "Patient", "ProvCover", "ProvPayer", "InfCover", "InfPayer", "SoharPayer", "Plan", "Network", "CovStart", "CovEnd",
    "Copay", "Coinsurance", "Deductible", "DedRemaining", "OopMax", "OopRemaining", "SoharStatus", "SoharDate", "SoharRef", "BidfPrinted", "Note", "By", "Updated"];
  /* who covers what: insurance for both, self pay for both, or one of each */
  const COVER_MODES = [["both", "Insurance covers both"], ["self", "Self pay for both"], ["provins", "Insurance for the provider visit, self-pay infusion"], ["infins", "Self-pay provider visit, insurance for the infusion"]];
  const isSelf = v => /self/i.test(txt(v));
  const modeOf = (prov, inf) => prov === "selfpay" && inf === "selfpay" ? "self" : prov === "selfpay" ? "infins" : inf === "selfpay" ? "provins" : "both";
  const modeParts = m => ({ prov: m === "self" || m === "infins" ? "selfpay" : "insurance", inf: m === "self" || m === "provins" ? "selfpay" : "insurance" });
  const SOHAR_KEYS = ["payer", "plan", "network", "covStart", "covEnd", "copay", "coins", "ded", "dedRem", "oop", "oopRem", "status", "date", "ref"];
  function profileRows(values) {
    const m = new Map();
    byHeader(values).rows.forEach(x => { const g = x.get, mrn = mrnKey(g("mrn")); if (!mrn) return;
      const cov = v => txt(v) ? (isSelf(v) ? "selfpay" : "insurance") : "";
      m.set(mrn, { mrn, code: txt(g("patient")), prov: { cover: cov(g("provcover")), payer: txt(g("provpayer")) }, inf: { cover: cov(g("infcover")), payer: txt(g("infpayer")) },
        sohar: { payer: txt(g("soharpayer")), plan: txt(g("plan")), network: txt(g("network")), covStart: toSerial(g("covstart")), covEnd: toSerial(g("covend")),
          copay: txt(g("copay")), coins: txt(g("coinsurance")), ded: txt(g("deductible")), dedRem: txt(g("dedremaining")), oop: txt(g("oopmax")), oopRem: txt(g("oopremaining")),
          status: txt(g("soharstatus")), date: toSerial(g("sohardate")), ref: txt(g("soharref")) },
        bidfPrinted: toSerial(g("bidfprinted")), note: txt(g("note")), by: txt(g("by")), updated: txt(g("updated")) }); });
    return m;
  }
  /* each part of the visit for one patient: insurance or self pay, the payer, and what's due. The tracker's insurance columns
     and amounts win; the profile fills in what the tracker doesn't say. */
  function coverage(p, prof) {
    /* "N/A", "none" or a dash on the tracker means no payer for that part */
    const named = v => { const s = txt(v); return /^(n\/?a|none|-+|—)$/i.test(s) ? "" : s; };
    const part = (ins0, pf, due) => { const ins = named(ins0), cover = pf?.cover || (isSelf(ins) ? "selfpay" : ins ? "insurance" : ""), payer = cover === "selfpay" ? "" : (ins && !isSelf(ins) ? ins : pf?.payer || "");
      return { cover, payer, due }; };
    const d = dueOf(p), prov = part(p?.ins1, prof?.prov, d.copay), inf = part(p?.ins2, prof?.inf, d.infusion);
    return { prov, inf, mode: prov.cover && inf.cover ? modeOf(prov.cover, inf.cover) : "", total: d.total };
  }
  /* New patients: on the appointment list (today on, within days) with no row on the Master Patient Tracker. A patient on the
     tracker already has their profile in its columns. Treatment starts first, then by first visit. */
  function profileGaps(A, PM, PF, today, days) {
    const by = new Map();
    for (const a of A) { if (a.date < today || a.date > today + days || isCancelled(a)) continue; const o = by.get(a.mrn); if (!o || a.date < o.first.date || (a.date === o.first.date && (a.time ?? 0) < (o.first.time ?? 0))) by.set(a.mrn, { first: a, start: (o && o.start) || /start/i.test(a.event) }); else if (/start/i.test(a.event)) o.start = true; }
    return [...by.entries()].filter(([mrn]) => !PM.has(mrn)).map(([mrn, o]) => ({ mrn, code: o.first.code || "", first: o.first, start: o.start, onMpt: false }))
      .sort((a, b) => (b.start - a.start) || a.first.date - b.first.date || (a.first.time ?? 0) - (b.first.time ?? 0));
  }
  /* The Sohar Eligibility add-in's benefits CSV ("Per Nbr", "MRN", "LASFIR", "Sohar Payer", ... "Sohar Ref"; the tracker file adds
     Service and more at the end). The header row is the first with MRN and Deductible Remaining or OOP Remaining. Discovery rows
     (Sent As) are skipped. Out: Map MRN -> the latest verification, in the profile's Sohar fields. */
  const SOHAR_CSV = { mrn: "MRN", lasfir: "LASFIR", payer: "Sohar Payer", plan: "Plan", network: "Network", covStart: "Coverage Start", covEnd: "Coverage End",
    copay: "Copay", coins: "Coinsurance", ded: "Deductible", dedRem: "Deductible Remaining", oop: "OOP Max", oopRem: "OOP Remaining", status: "Sohar Status", date: "Sohar Date", ref: "Sohar Ref", sentAs: "Sent As" };
  function parseSohar(rows) {
    const hi = (rows || []).slice(0, 30).findIndex(r => { const h = (r || []).map(norm); return h.includes("mrn") && (h.includes("deductibleremaining") || h.includes("oopremaining")); });
    if (hi < 0) return null;
    const hn = rows[hi].map(norm), idx = Object.fromEntries(Object.entries(SOHAR_CSV).map(([k, h]) => [k, hn.indexOf(norm(h))]));
    const out = new Map();
    for (const r of rows.slice(hi + 1)) {
      const get = k => idx[k] < 0 ? "" : txt(r[idx[k]]), mrn = mrnKey(get("mrn")); if (!mrn) continue;
      if (get("sentAs") && !/^verification$/i.test(get("sentAs"))) continue;
      const rec = { mrn, code: get("lasfir"), payer: get("payer"), plan: get("plan"), network: get("network"), covStart: toSerial(get("covStart")), covEnd: toSerial(get("covEnd")),
        copay: get("copay"), coins: get("coins"), ded: get("ded"), dedRem: get("dedRem"), oop: get("oop"), oopRem: get("oopRem"), status: get("status"), date: toSerial(get("date")), ref: get("ref") };
      const prev = out.get(mrn); if (!prev || (rec.date ?? 0) >= (prev.date ?? 0)) out.set(mrn, rec);
    }
    return out;
  }

  /* The whole Sohar IVN export (the Sohar Eligibility add-in's file for the TMS, ESK and IVK Trackers) against this tracker.
     M: parseSohar's Map. o = { PM (Master Patient Tracker by MRN), PF (profiles by MRN), A (appointment list), by, now (iso) }.
     A patient counts when they're on the tracker, have a profile, or are on the appointment list. A verification older than the
     one already saved is left alone; the same date and ref is unchanged. Out: { items: [{ mrn, code, patch, isNew }], same,
     older, notHere, rows } where patch holds the profile's Sohar columns (dates as iso). */
  function planSoharImport(M, o) {
    const PM = o.PM || new Map(), PF = o.PF || new Map(), apptCode = new Map();
    for (const a of o.A || []) if (a.mrn && !apptCode.has(a.mrn)) apptCode.set(a.mrn, a.code || "");
    const iso = d => d != null ? isoOf(d) : "", out = { items: [], same: 0, older: 0, notHere: 0, rows: M ? M.size : 0 };
    for (const [mrn, r] of M || []) {
      const p = PM.get(mrn), pf = PF.get(mrn);
      if (!p && !pf && !apptCode.has(mrn)) { out.notHere++; continue; }
      const had = pf && pf.sohar && pf.sohar.date;
      if (had != null && r.date != null && r.date < had) { out.older++; continue; }
      if (had != null && r.date === had && txt(pf.sohar.ref) === txt(r.ref)) { out.same++; continue; }
      const code = p?.code || pf?.code || apptCode.get(mrn) || r.code || "";
      const patch = { SoharPayer: r.payer, Plan: r.plan, Network: r.network, CovStart: iso(r.covStart), CovEnd: iso(r.covEnd), Copay: r.copay, Coinsurance: r.coins,
        Deductible: r.ded, DedRemaining: r.dedRem, OopMax: r.oop, OopRemaining: r.oopRem, SoharStatus: r.status, SoharDate: iso(r.date), SoharRef: r.ref, By: o.by || "", Updated: o.now || "" };
      if (!pf && code) patch.Patient = code;
      out.items.push({ mrn, code, patch, isNew: !pf });
    }
    out.items.sort((a, b) => a.code.localeCompare(b.code));
    return out;
  }

  /* ---------- access: demand by protocol ---------- */
  /* What the IVK schedule has to hold, from each patient's protocol (the tracker's Frequency), against how they're booked.
     o = { patients, appts, holds, today, cfg, weeks, paused (include Paused patients), sched ("all" | "exclude" unscheduled | "only"
     unscheduled), X (the templates, for the slots each week offers) }.
     Index patients (Induction, or a course started within the index weeks with fewer than the index's visits done) come at
     IndexPerWeek until treatment N (IndexPerWeek x IndexWeeks, 6), then at least AfterIndex (once a week): that step down is the
     access they free. Weeks run Sunday to Saturday from this week. */
  function demand(o) {
    const c = o.cfg, T = o.today, W = o.weeks || 12, N = Math.max(1, Math.round(+c.IndexPerWeek * +c.IndexWeeks)), w0 = weekStart(T);
    const after = 7 / (freqDays(c.AfterIndex) || 7), perIdx = +c.IndexPerWeek;
    const held = (o.holds || []).filter(h => holding(h, o.appts, o.listAt));
    const rows = [];
    for (const p of o.patients || []) {
      const st = txt(p.status).toLowerCase(), paused = /paused/.test(st);
      if (!(isActive(p) || (o.paused && paused))) continue;
      const fut = futureOf(o.appts, p.mrn, T - 1).concat(held.filter(h => h.mrn === p.mrn && h.date >= T).map(h => ({ date: h.date, time: h.time, hold: true })));
      const unsched = !fut.length;
      if (o.sched === "exclude" && unsched) continue;
      if (o.sched === "only" && !unsched) continue;
      const fd = freqDays(p.freq), prn = /prn/i.test(p.freq || ""), rate = fd ? 7 / fd : 0;
      /* index: treatment N's day, booked, or projected at IndexPerWeek from the visits left */
      const L = countable(p, o.appts) ? treatments(p, o.appts) : null, done = L ? L.filter(a => a.date < T).length : null;
      /* an index patient is one whose status is Induction */
      const idx = /induction/.test(st);
      let idxEnd = null;
      if (idx) { const left = Math.max(0, N - (done ?? 0)), bookedN = L && L[N - 1] ? L[N - 1].date : null; idxEnd = bookedN ?? T + Math.ceil(left / perIdx) * 7; }
      const afterRate = Math.max(after, prn ? 0 : rate && rate < perIdx ? rate : after);
      const expect = Array.from({ length: W }, (_, w) => { const a = w0 + 7 * w, b = a + 7, from = Math.max(a, T), days = b - from; if (days <= 0) return 0;
        if (!idx) return rate * days / 7;
        if (idxEnd >= b) return perIdx * days / 7;
        if (idxEnd < from) return afterRate * days / 7;
        return (perIdx * (idxEnd - from + 1) + afterRate * (b - idxEnd - 1)) / 7; });
      const booked = Array.from({ length: W }, (_, w) => fut.filter(f => f.date >= w0 + 7 * w && f.date < w0 + 7 * w + 7).length);
      rows.push({ p, mrn: p.mrn, code: p.code, freq: idx ? "Index" : p.freq || "No frequency", proto: p.freq || "", rate: idx ? perIdx : rate, prn, unsched, paused, idx, idxEnd,
        afterRate: idx ? afterRate : null, expect, booked, next: fut.sort((x, y) => x.date - y.date)[0] || null });
    }
    const sum = (L, f) => L.reduce((n, x) => n + f(x), 0), win = (arr, n) => arr.slice(0, n).reduce((a, b) => a + b, 0);
    const weeks = Array.from({ length: W }, (_, w) => {
      const a = w0 + 7 * w, exp = sum(rows, r => r.expect[w]), bk = sum(rows, r => r.booked[w]), gap = sum(rows, r => Math.max(0, r.expect[w] - r.booked[w]));
      let slots = 0, open = 0;
      if (o.X) for (let d = Math.max(a, T); d < a + 7; d++) { const wd = weekday(d); if (wd < 1 || wd > 5) continue; const S = daySlots(d, o.X).slots; slots += S.length; open += S.filter(s => s.open).length; }
      /* the step down from index to after-index visits that lands in this week */
      const steps = rows.filter(r => r.idx && r.idxEnd >= a && r.idxEnd < a + 7);
      return { w, start: a, expect: exp, booked: bk, toBook: gap, slots, open, freed: sum(rows.filter(r => r.idx && r.idxEnd < a + 7), r => Math.max(0, perIdx - r.afterRate) * Math.min(7, a + 7 - Math.max(a, r.idxEnd + 1)) / 7), steps };
    });
    const groups = new Map();
    for (const r of rows) { const g = groups.get(r.freq) || { freq: r.freq, rate: r.rate, n: 0, unsched: 0, paused: 0, perWeek: 0, rows: [] }; g.n++; g.rows.push(r); if (r.unsched) g.unsched++; if (r.paused) g.paused++; g.perWeek += r.expect.length ? (r.idx ? perIdx : r.rate) : 0; groups.set(r.freq, g); }
    /* still to book over a window: each patient's need over it less what they have booked in it (never below 0), added up */
    const tb = (L, n) => sum(L, r => Math.max(0, win(r.expect, n) - win(r.booked, n)));
    const G = [...groups.values()].map(g => Object.assign(g, { d4: sum(g.rows, r => win(r.expect, 4)), b4: sum(g.rows, r => win(r.booked, 4)), t4: tb(g.rows, 4), d8: sum(g.rows, r => win(r.expect, 8)), b8: sum(g.rows, r => win(r.booked, 8)), t8: tb(g.rows, 8),
      d12: sum(g.rows, r => win(r.expect, 12)), b12: sum(g.rows, r => win(r.booked, 12)) }))
      .sort((a, b) => (a.freq === "Index" ? -1 : b.freq === "Index" ? 1 : 0) || (b.rate - a.rate) || a.freq.localeCompare(b.freq));
    const perWeek = sum(rows, r => r.idx ? perIdx : r.rate);
    return { rows, groups: G, weeks, N, perIdx, after, perWeek, avg: rows.length ? perWeek / rows.length : 0, toBook4: tb(rows, 4), toBook8: tb(rows, 8),
      transitions: rows.filter(r => r.idx).sort((a, b) => a.idxEnd - b.idxEnd) };
  }

  /* ---------- index: Induction patients ---------- */
  /* An Induction patient is in index. Their kept visits (not cancelled, not no-shows) are counted from the treatment start; once
     the Nth (IndexPerWeek x IndexWeeks, 6) is kept, index is complete: the desk moves them to Ongoing and confirms the protocol
     (usually 1x per week, set by the provider by acuity). Kept = before today, or today once checked out (keptToday). */
  function indexState(p, A, today, cfg, keptToday) {
    if (!p || !/induction/i.test(p.status || "")) return null;
    const N = Math.max(1, Math.round(+cfg.IndexPerWeek * +cfg.IndexWeeks)), L = treatments(p, A);
    const done = L.filter(a => a.date < today || (a.date === today && keptToday && keptToday(a))).length;
    return { N, done, complete: countable(p, A) && done >= N, sixth: L[N - 1] || null, countable: countable(p, A) };
  }
  /* A treatment start for an Induction patient still on the index protocol (IndexPerWeek x a week, 2x; or none set yet) should have
     the rest of index booked behind it: N - 1 (5) visits after the start. Null when it does; else how many are on the list. */
  function indexShort(p, A, a, cfg) {
    if (!p || !a || !/start/i.test(a.event || "") || !/induction/i.test(p.status || "")) return null;
    if (p.freq && !new RegExp(`^\\s*${+cfg.IndexPerWeek}x per week`, "i").test(p.freq)) return null;
    const need = Math.max(1, Math.round(+cfg.IndexPerWeek * +cfg.IndexWeeks)) - 1;
    const booked = A.filter(x => x.mrn === p.mrn && !isCancelled(x) && !/no ?show/i.test(x.status) && (x.date > a.date || (x.date === a.date && (x.time ?? 0) > (a.time ?? 0)))).length;
    return booked < need ? { booked, need, start: a } : null;
  }
  /* Kept all N index visits (6) and still on the index rate (IndexPerWeek x per week, 2x) with the protocol not updated since the
     Nth: time to step down to AfterIndex (1x per week). Induction patients get the Move to Ongoing ask (indexState) instead. */
  function stepDue(p, A, today, cfg) {
    if (!p || !isActive(p) || /induction/i.test(p.status) || !countable(p, A)) return null;
    if (!new RegExp(`^\\s*${+cfg.IndexPerWeek}\\s*x\\s*(per|a|/)?\\s*w(ee)?k`, "i").test(p.freq || "")) return null;
    const N = Math.max(1, Math.round(+cfg.IndexPerWeek * +cfg.IndexWeeks)), L = treatments(p, A).filter(a => a.date < today), nth = L[N - 1];
    if (!nth || (p.protoAt != null && p.protoAt >= nth.date)) return null;
    return { N, nth, done: L.length, freq: p.freq, to: cfg.AfterIndex };
  }
  /* the patient's treatment start from today on (today's first), for the index-booking flag */
  const nextStart = (p, A, today) => A.filter(a => a.mrn === p.mrn && a.date >= today && /start/i.test(a.event) && !isCancelled(a)).sort((x, y) => x.date - y.date || (x.time ?? 0) - (y.time ?? 0))[0] || null;

  /* ---------- transfers (tblIvkTransfers) ---------- */
  /* MRN | Patient | To | Date | Status | Note | Confirmed | By | Updated. One row per patient going to another clinic (Oakland
     maintenance after index here, or a move). Status: scheduled (a future date; confirmed at their last visit here), done (Status
     Discontinued and a note on the Master Patient Tracker), cancelled. */
  const XFER_COLS = ["MRN", "Patient", "To", "Date", "Status", "Note", "Confirmed", "By", "Updated"];
  function xferRows(values) {
    const m = new Map();
    byHeader(values).rows.forEach(x => { const g = x.get, mrn = mrnKey(g("mrn")); if (!mrn) return;
      m.set(mrn, { mrn, code: txt(g("patient")), to: txt(g("to")), date: toSerial(g("date")), status: txt(g("status")).toLowerCase(), note: txt(g("note")), confirmed: toSerial(g("confirmed")), by: txt(g("by")) }); });
    return m;
  }
  /* a scheduled transfer: the last visit here before the date (where it's confirmed), visits still booked here on or after it, and
     whether it's time to confirm (that last visit is today or past, or the date has come) */
  function xferState(xf, A, today, E) {
    if (!xf || xf.status !== "scheduled") return null;
    const L = A.filter(a => a.mrn === xf.mrn && !isCancelled(a) && !/no ?show/i.test(a.status)).sort((a, b) => a.date - b.date || (a.time ?? 0) - (b.time ?? 0));
    const before = xf.date == null ? L : L.filter(a => a.date < xf.date), last = before[before.length - 1] || null;
    const after = xf.date == null ? [] : L.filter(a => a.date >= xf.date && a.date >= today);
    const due = (xf.date != null && today >= xf.date) || !last || last.date <= today;
    /* there: their first visit at the new clinic from the wider report (null when none; E null = no wider report uploaded) */
    const there = E ? bookedAt(E, xf.mrn, xf.to, Math.max(today, xf.date ?? today)) : undefined;
    return { to: xf.to, date: xf.date, last, after, due, there };
  }
  /* ---------- other clinics (the wider report) ---------- */
  /* tblIvkElsewhere: IVK visits at the neighbouring clinics, from a wider NextGen report. LASFIR only, no names. */
  const ELSE_COLS = ["MRN", "Patient", "Date", "Time", "Provider", "Event", "Status", "Location"];
  const elseRows = values => byHeader(values).rows.map(x => ({ mrn: mrnKey(x.get("mrn")), code: txt(x.get("patient")), date: toSerial(x.get("date")), time: toMin(x.get("time")),
    provider: txt(x.get("provider")), event: txt(x.get("event")), status: txt(x.get("status")), loc: txt(x.get("location")) })).filter(a => a.mrn && a.date != null);
  /* one place names the other ("Oakland" and "MHS Oakland IVK") */
  const sameLoc = (a, b) => { const x = norm(a), y = norm(b); return !!x && !!y && (x.includes(y) || y.includes(x)); };
  /* is a report location this clinic's? the Locations setting when set, else blank or naming the clinic */
  function isOurs(loc, cfg) {
    if (!txt(loc)) return true;
    const L = txt(cfg.Locations).split(/[;\n]/).map(norm).filter(Boolean);
    return L.length ? L.includes(norm(loc)) : sameLoc(loc, cfg.Clinic);
  }
  /* their first visit at another clinic (one named like "to", or any when "to" is blank) on or after "from", not cancelled */
  const bookedAt = (E, mrn, to, from) => (E || []).filter(a => a.mrn === mrn && a.date >= from && !isCancelled(a) && !/no ?show/i.test(a.status) && (!txt(to) || sameLoc(a.loc, to)))
    .sort((a, b) => a.date - b.date || (a.time ?? 0) - (b.time ?? 0))[0] || null;
  /* ---------- sooner waitlist ---------- */
  /* tblIvkSooner: a patient booked on one date who'd take an earlier one, one row per patient. Days "1am,3pm" (weekday 1-5 and
     half), blank = any. SameProv TRUE = only with the provider they're booked with. Status waiting, done or removed. Seen: the
     openings already shown ("2026-11-10|600|rivera;..."), so a refreshed list only alerts what's new. */
  const SOONER_COLS = ["MRN", "Patient", "Booked", "BookedTime", "Provider", "From", "Days", "SameProv", "Status", "Note", "Seen", "By", "Updated"];
  function soonerRows(values) {
    const m = new Map();
    byHeader(values).rows.forEach(x => { const g = x.get, mrn = mrnKey(g("mrn")), booked = toSerial(g("booked")); if (!mrn || booked == null) return;
      m.set(mrn, { mrn, code: txt(g("patient")), booked, bookedTime: toMin(g("bookedtime")), provider: txt(g("provider")), from: toSerial(g("from")),
        days: new Set(txt(g("days")).split(/[,\s]+/).filter(d => /^[1-5](am|pm)$/.test(d))), sameProv: /^(true|yes|1)$/i.test(txt(g("sameprov"))),
        status: txt(g("status")).toLowerCase() || "waiting", note: txt(g("note")), seen: new Set(txt(g("seen")).split(";").map(txt).filter(Boolean)), by: txt(g("by")) }); });
    return m;
  }
  const soonerKey = c => `${isoOf(c.date)}|${c.time}|${TM.provKey(c.provider)}`;
  /* which of the openings (open slots, or times a cancellation freed) would move them sooner: from From (or tomorrow) to the day
     before the visit they have, on the days and halves they said, with their provider if asked; a start needs an index start slot */
  function soonerMatch(r, cands, today, isStart) {
    if (!r || r.status !== "waiting") return [];
    const from = Math.max(today + 1, r.from ?? 0), seen = new Set();
    return cands.filter(c => c.date >= from && c.date < r.booked && c.time != null && c.mrn !== r.mrn && (!isStart || c.index !== false)
      && (!r.days.size || r.days.has(`${weekday(c.date)}${c.time < 720 ? "am" : "pm"}`)) && (!r.sameProv || sameProv(c.provider, r.provider)))
      .map(c => Object.assign({}, c, { key: soonerKey(c) })).filter(c => !seen.has(c.key) && seen.add(c.key))
      .sort((a, b) => a.date - b.date || a.time - b.time);
  }
  /* open template slots from tomorrow to a date (index: false for follow-up-only slots) */
  function openSlots(X, today, until) {
    const out = [];
    for (let d = today + 1; d <= until; d++) { const wd = weekday(d); if (wd < 1 || wd > 5) continue;
      for (const s of daySlots(d, X).slots) if (s.open) out.push({ date: d, time: s.time, provider: s.provider, index: !!s.openIndex, src: "slot" }); }
    return out;
  }
  /* visits on the old list that the new one cancels or no longer has, from tomorrow on: times a cancellation freed */
  function freedTimes(oldA, newA, today) {
    const live = a => !isCancelled(a) && !/no ?show/i.test(a.status), L = newA.filter(live), now = new Set(L.map(a => `${a.mrn}|${a.date}|${a.time}`));
    const taken = new Set(L.map(a => `${a.date}|${a.time}|${TM.provKey(a.provider)}`));   // someone else is already in that time
    return oldA.filter(a => live(a) && a.date > today && a.time != null && !now.has(`${a.mrn}|${a.date}|${a.time}`) && !taken.has(`${a.date}|${a.time}|${TM.provKey(a.provider)}`))
      .map(a => ({ date: a.date, time: a.time, provider: a.provider, mrn: a.mrn, src: "freed" }));
  }

  /* home clinic isn't this one: Primary Location isn't one of this clinic's locations (cfg, or just the clinic's name) */
  const elsewhere = (p, cfg) => { const c = typeof cfg === "string" ? { Clinic: cfg } : cfg || {}; return !!(p && p.loc && c.Clinic && !isOurs(p.loc, c)); };
  /* a patient's IVK appointments: kept and booked here (the Appt List), and booked at each other clinic (the wider report) */
  function visitCounts(mrn, A, E, today) {
    const live = a => a.mrn === mrn && !isCancelled(a), L = A.filter(live), away = new Map();
    (E || []).filter(a => live(a) && a.date >= today && !/no ?show/i.test(a.status)).forEach(a => away.set(a.loc, (away.get(a.loc) || 0) + 1));
    return { kept: L.filter(a => a.date < today && !/no ?show/i.test(a.status)).length, booked: L.filter(a => a.date >= today).length, away: [...away].map(([loc, n]) => ({ loc, n })) };
  }

  /* ---------- treatment numbers, Oakland ---------- */
  /* the patient's IVK treatments in order, from the treatment start (Tx Start Date, else the first "start" visit on the list) */
  function treatments(p, A) {
    const L = A.filter(a => a.mrn === p.mrn && !isCancelled(a) && !/no ?show/i.test(a.status)).sort((a, b) => a.date - b.date || (a.time ?? 0) - (b.time ?? 0));
    const first = p.start ?? (L.find(a => /start/i.test(a.event)) || {}).date;
    return first == null ? [] : L.filter(a => a.date >= first);
  }
  /* the count is only right when the appointment list goes back to the treatment start; a course that began before the list's first
     day (the export's date range) can't be numbered from it */
  const listFrom = A => A.reduce((m, a) => a.date < m ? a.date : m, Infinity);
  const countable = (p, A) => p.start == null || p.start >= listFrom(A);
  const txNo = (p, A, date, time) => { if (!countable(p, A)) return null; const i = treatments(p, A).findIndex(a => a.date === date && a.time === time); return i < 0 ? null : i + 1; };
  /* Oakland patients have their index (the first OakIndex treatments) at Castro and go back to Oakland for maintenance. Blank on the
     add-in's table: an Oakland home clinic decides. */
  const isOak = (p, pt) => pt && pt.oak != null ? pt.oak : /oak/i.test(p?.loc || "");
  /* where an Oakland patient is: done (count of treatments up to today), task (treatment OakRemindAt reached, OAK not tasked yet),
     back (index finished, still booked here after it) */
  function oakState(p, pt, A, today, cfg) {
    if (!isOak(p, pt)) return null;
    if (!countable(p, A)) return { unknown: true, done: null, idx: +cfg.OakIndex, task: false, tasked: pt?.oakTasked ?? null, back: false, after: [], booked: 0 };
    const L = treatments(p, A), done = L.filter(a => a.date <= today).length, idx = +cfg.OakIndex, at = +cfg.OakRemindAt;
    const lastIdx = L[idx - 1], after = lastIdx ? L.filter(a => a.date > lastIdx.date && a.date > today) : [];
    return { done, idx, task: done >= at && !(pt && pt.oakTasked != null), tasked: pt?.oakTasked ?? null, back: done >= idx && after.length > 0, after, booked: L.length };
  }

  /* ---------- waitlist ---------- */
  /* Everyone on the Consults tab still waiting for a treatment start, in line by consult date (or the order set by hand, WaitRank).
     Waiting on: ready (approved, can start), referral (keeps their place until it comes), auth (an appeal can push their start back,
     so the next ready patient usually starts first). */
  const WAIT_ON = [["ready", "Ready to start"], ["referral", "Referral"], ["auth", "Authorization"]];
  const waitDefault = c => /unclear|referr/i.test(c.stage) ? "referral" : c.state === "wait" ? "auth" : "ready";
  function waitlist(CS, PT) {
    const L = CS.filter(c => c.mrn && (c.state === "act" || c.state === "wait")).map(c => { const t = PT.get(c.mrn) || null;
      return { c, mrn: c.mrn, code: c.code, t, rank: t?.rank ?? null, on: WAIT_ON.some(w => w[0] === t?.waitOn) ? t.waitOn : waitDefault(c), appeal: !!t?.appeal && (t?.waitOn || waitDefault(c)) === "auth" }; });
    /* hand-set places first (their order), then the rest by consult date, then referral date */
    L.sort((a, b) => (a.rank ?? 1e9) - (b.rank ?? 1e9) || (a.c.consult ?? 1e9) - (b.c.consult ?? 1e9) || (a.c.ref ?? 1e9) - (b.c.ref ?? 1e9) || a.code.localeCompare(b.code));
    const next = L.find(w => w.on === "ready") || null;
    L.forEach((w, i) => { w.pos = i + 1; w.next = w === next; w.ahead = !!next && i < L.indexOf(next); });
    return { list: L, next };
  }
  /* the line after moving one patient: dir -1 up, +1 down, "top"; or, for an appeal, behind the next ready patient below them */
  function waitMove(list, mrn, dir) {
    const L = list.map(w => w.mrn), i = L.indexOf(mrn); if (i < 0) return L;
    L.splice(i, 1);
    let j = dir === "top" ? 0 : dir === "appeal" ? (() => { const k = list.findIndex((w, x) => x > i && w.on === "ready"); return k < 0 ? L.length : k; })() : i + dir;
    L.splice(Math.max(0, Math.min(L.length, j)), 0, mrn);
    return L;
  }

  /* ---------- patient schedule (print, start packet, start email) ---------- */
  /* "Rivera MD, Ana Maria" -> "Ana Maria Rivera, MD"; "Chen, Paul MD" -> "Paul Chen, MD" */
  function provFull(p) {
    const s = txt(p); if (!s) return "";
    const cred = (s.match(/\b(MD|DO|NP|PA|PMHNP|FNP|RN|PhD)\b/i) || [""])[0].toUpperCase(), parts = s.split(",").map(x => x.replace(/\b(MD|DO|NP|PA|PMHNP|FNP|RN|PhD)\b\.?/gi, "").trim());
    const name = parts.length > 1 && parts[1] ? `${parts[1]} ${parts[0]}` : parts[0];
    return `${name.replace(/\s+/g, " ").trim()}${cred ? `, ${cred.replace("PHD", "PhD")}` : ""}`;
  }
  /* "Smith, Jane" -> "Jane Smith"; first name for a greeting */
  const fullName = n => { const s = txt(n); if (!s.includes(",")) return s; const [l, f] = s.split(","); return `${f.trim()} ${l.trim()}`.trim(); };
  const firstName = n => fullName(n).split(/\s+/)[0] || "";
  /* a patient's IVK visits from a day on, with treatment numbers: { date, time, day, when, provider, type, n, loc } */
  function scheduleOf(p, A, from) {
    const L = countable(p, A) ? treatments(p, A) : [], all = A.filter(a => a.mrn === p.mrn && a.date >= from && !isCancelled(a)).sort((a, b) => a.date - b.date || (a.time ?? 0) - (b.time ?? 0));
    return all.map(a => { const i = L.indexOf(a); return { date: a.date, time: a.time, day: fmtDay(a.date), when: `${fmtDay(a.date)}, ${dateOf(a.date).getFullYear()}`, at: fmtT(a.time),
      provider: provFull(a.provider), type: /start/i.test(a.event) ? "Treatment start" : "IV ketamine infusion", n: i < 0 ? null : i + 1, loc: a.loc, status: a.status }; });
  }
  /* "Monday, November 4th, 2026, at 10:00 AM", as the start email says it */
  const ordinal = n => n + (n % 100 >= 11 && n % 100 <= 13 ? "th" : ["th", "st", "nd", "rd"][n % 10] || "th");
  const longWhen = (d, t) => { const x = dateOf(d); return `${DAY_NAMES[x.getDay()]}, ${MONTHS[x.getMonth()]} ${ordinal(x.getDate())}, ${x.getFullYear()}${t != null ? `, at ${fmtT(t)}` : ""}`; };
  /* The desk's start email ("MHS: Upcoming Appointment(s)"), as HTML. {mrn} {first} {clinic} {firstVisit} {address} {phone} are
     filled in; {schedule} is where the appointment table goes and {signature} the sender's own (blank unless they added one in
     Settings › You: Outlook adds theirs to a new message). {map} is the clinic's location/parking map (Settings › Messages), with its
     own "Getting Here" heading; nothing shows there until a map is loaded. Edit it in Settings › Messages. */
  const LI = s => `<li style="margin:0">${s}</li>`, UL = a => `<ul style="margin-top:0;margin-bottom:0">${a.map(LI).join("")}</ul>`, H = (s, m) => `<p style="margin:${m || "12pt"} 0 0"><b>${s}</b></p>`;
  const START_TPL = {
    startSubject: "SECURE: Upcoming Visit Information | MHS",
    startEmail: [
      `<p style="margin:0"><b>MRN: {mrn}</b><br><br>Hi {first}!<br><br>Thank you for connecting with me via phone recently. As mentioned, please see the information below pertaining to your upcoming visit(s).<br><br>We’re looking forward to meeting you at our {clinic} clinic, <b>{firstVisit}</b> for your first treatment at <b>{address}</b>. Below is a brief overview of what to expect and how to prepare, so your first session goes as smoothly and comfortably as possible. Please note that you may want to budget extra time for your first appointment, though you typically can expect appointments to last at least 2 hours. Your Benefits Investigation Disclosure Form (BIDF) has also been included for your review.</p>`,
      H("Your Upcoming Appointments:", "6pt"), "{schedule}", "{map}",
      H("Before Your Appointment:", "6pt"), UL(["<b>Avoid food for at least 4 hours</b> prior to your infusion. Small sips of water are okay up to 1 hour before.",
        "<b>Hydrate well earlier in the day</b>, but limit liquids in the hour leading up to your appointment.", "<b>Take your usual medications</b> unless your provider has instructed otherwise.",
        "<b>Avoid alcohol or recreational substances</b> for at least 24 hours beforehand.", "<b>Arrange a ride home.</b> You will not be able to drive or take rideshare alone after your session.",
        "<b>Arrive Early</b> – You will have a short two documents to review prior to your visit, 5-10 minutes is the average time to review the documents!"]),
      H("What To Bring:"), UL(["Layers or a light sweater—rooms can feel cool while relaxing.", "Headphones, an eye mask, or a comfort item to help you feel grounded."]),
      H("What to Expect During the Session:"), UL(["You’ll be seated in a comfortable treatment chair throughout the infusion.",
        "Our RN will place the IV and monitor your vitals and response throughout treatment, while your provider will connect with you 1:1 and follow up with you during the end of your observation period, prior to discharge.",
        "Normal experiences may include floating sensations, mild dissociation, or visual distortions. These typically resolve shortly after the infusion ends.", "The infusion itself usually lasts 40–60 minutes, followed by a brief recovery period."]),
      H("After Your Treatment:"), UL(["Plan for quiet rest and avoid driving, working, major decisions, or strenuous activity until the next day.", "Continue hydrating and eat light, balanced meals once your stomach feels settled."]),
      H("How to Contact Us:"), UL(["<b>Email - </b>CC’d is our {clinic} help email, for direct access to our team regarding any needs you may have as it pertains to your appointments (best for non-urgent communication)",
        "<b>Phone -</b> For more immediate communication, any questions before or after your appointment, feel free to contact us directly at <b>{phone}</b>",
        "<b>Portal –</b> Patient Portal messaging is HIPAA secure, offers appointment visibility, and message history if you prefer to consolidate communication"]),
      `<p style="margin:12pt 0 0"><b>Please note:</b> We ask that patients confirm their appointments or request a reschedule at least 24 <i>business hours </i>prior to their appointment. With our patients’ cooperation, prompt cancellations and reschedules ensure we have appointment availability to accommodate reschedule requests, or make needed adjustments to your schedule. You will receive an electronic confirmation request at least 5 days prior to all appointments.</p>`,
      `<p style="margin:12pt 0 0">Whatever you may need, we’re here to ensure you feel supported and informed every step of the way.</p>`,
      "{signature}",
      `<p style="margin:12pt 0 0;font-size:10pt">This email communication and any attachments may contain confidential information for the use of the designated recipients named above.  Distribution, reproduction or any other use of this transmission by any party other than the intended recipient is prohibited and may be subject to protection under the law, including the Health Insurance Portability and Accountability Act (HIPAA).  If you receive this e-mail in error, please notify the sender immediately by reply e-mail and permanently delete this e-mail and any attachments.</p>`
    ].join("\n")
  };

  /* ---------- who's at the desk, the clinic, admin ---------- */
  /* the clinic in a tracker's name: "Castro IVK Patient Tracker" -> "Castro", "Walnut Creek IVK Tracker 2026" -> "Walnut Creek".
     "" when the name doesn't start with one (same rule as the TMS Tracker's) */
  function clinicFromName(name) {
    const s = String(name || "").replace(/\.(xlsx|xlsm|xlsb|xls)$/i, "").replace(/[_-]+/g, " ");
    const W = /^(ivk|iv\s?ketamine|ketamine|tms|esk|pt|patient|master|tracker|mhs|copy of)$/i;
    const m = s.match(/^(.*?)\s+(ivk|iv\s?ketamine|ketamine|tms|esk|pt|patient|master|tracker)\b/i), c = (m ? m[1] : "").replace(/\s{2,}/g, " ").trim();
    return W.test(c) ? "" : c.replace(/^copy of\s+/i, "");
  }
  /* tblIvkUsers: the people registered on this workbook. Only this list is offered when picking who's at the desk.
     Settings (JSON) holds their own settings (today only the optional email signature). */
  const USER_COLS = ["Name", "Added", "Updated", "Settings (JSON)"];
  const userKey = n => txt(n).toLowerCase();
  function userRows(values) {
    const seen = new Set(), out = [];
    for (const x of byHeader(values).rows) {
      const name = txt(x.get("name")), k = userKey(name); if (!name || seen.has(k)) continue; seen.add(k);
      let set = {}; try { const j = JSON.parse(txt(x.get("settingsjson")) || "{}"); if (j && typeof j === "object" && !Array.isArray(j)) set = j; } catch (e) { /* unreadable: no settings */ }
      out.push({ name, added: txt(x.get("added")), updated: txt(x.get("updated")), set });
    }
    return out;
  }
  /* the picker's order: the people who have used this computer on this workbook first (most recent first), then everyone else A-Z.
     here = names this computer remembers for this workbook, newest first. Nobody is picked for them. */
  function pickOrder(users, here) {
    const H = (here || []).map(userKey), rank = u => { const i = H.indexOf(userKey(u.name)); return i < 0 ? Infinity : i; };
    return (users || []).map(u => Object.assign({}, u, { here: rank(u) !== Infinity }))
      .sort((a, b) => rank(a) - rank(b) || a.name.localeCompare(b.name, undefined, { sensitivity: "base" }));
  }
  /* the name typed to register, as Teams posts show it: usually a first name (a last initial when two people share one) */
  const cleanName = s => txt(s).replace(/\s+/g, " ").slice(0, 30).trim();
  /* this computer's memory is kept per workbook: a short key from the workbook's URL (or the id saved in it) */
  function wbKeyOf(s) {
    let h = 0x811c9dc5; const t = String(s || "").trim().toLowerCase().replace(/[?#].*$/, "");
    for (let i = 0; i < t.length; i++) { h ^= t.charCodeAt(i); h = Math.imul(h, 0x01000193) >>> 0; }
    return t ? "w" + h.toString(36) : "";
  }
  /* what this computer remembers for one workbook: { day: "2026-11-05", name: "Sam", used: ["Sam", "Ana"] }. The pick lasts the day. */
  function whoRec(json) { try { const o = JSON.parse(json || "{}"); return { day: txt(o.day), name: txt(o.name), used: Array.isArray(o.used) ? o.used.map(txt).filter(Boolean) : [] }; } catch (e) { return { day: "", name: "", used: [] }; } }
  const whoToday = (rec, dayIso) => rec && rec.day === dayIso ? rec.name : "";
  const whoPick = (rec, name, dayIso) => ({ day: dayIso, name, used: [name].concat((rec?.used || []).filter(n => userKey(n) !== userKey(name))).slice(0, 20) });
  /* Admin: code 0582 unlocks on any workbook. Compared as a salted SHA-256 so the code isn't written here in plain text. The repo is
     public, so this keeps honest people from changing clinic-wide settings by accident; it is not a security control. */
  const ADMIN_SALT = "ivk-tracker-admin", ADMIN_HASH = "14068cdf14b0b225199c46811cd9623fd862c9bee291adc30e89c36649a1bf77", ADMIN_MIN = 15;
  async function adminOk(code) {
    const c = String(code ?? "").trim(); if (!c) return false;
    const sub = (typeof crypto !== "undefined" && crypto.subtle) || (typeof require === "function" ? require("crypto").webcrypto.subtle : null); if (!sub) return false;
    const b = await sub.digest("SHA-256", new TextEncoder().encode(`${ADMIN_SALT}|${c}`));
    return Array.from(new Uint8Array(b)).map(x => x.toString(16).padStart(2, "0")).join("") === ADMIN_HASH;
  }
  /* clinic-wide settings: whoever enters one first saves it for the workbook; once saved, changing it takes admin */
  const CLINIC_WIDE = ["Clinic", "ClinicAddress", "ClinicPhone", "StartCc", "StartBcc"];
  const lockedCfg = (saved, k, admin) => CLINIC_WIDE.includes(k) && !admin && txt((saved || {})[k]) !== "";
  /* An email signature from a file: Outlook keeps each signature as .htm and .txt in %APPDATA%\Microsoft\Signatures. The .htm's
     body is kept without scripts, styles, event handlers or local pictures (those can't travel in the email); text becomes lines. */
  const escH = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  function sigFromFile(text, isHtml) {
    let s = String(text || "").replace(/^\uFEFF/, "");
    if (!isHtml) return s.trim() ? s.trim().split(/\r?\n/).map(escH).join("<br>") : "";
    const b = s.match(/<body[^>]*>([\s\S]*)<\/body>/i); if (b) s = b[1];
    s = s.replace(/<!--[\s\S]*?-->/g, "").replace(/<(script|style|head|xml|object|iframe|embed)[\s\S]*?<\/\1>/gi, "").replace(/<\/?(script|style|object|iframe|embed|link|meta|base|form|input)[^>]*>/gi, "")
      .replace(/<\/?[a-z]+:[^>]*>/gi, "").replace(/\s+on[a-z]+\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/gi, "")
      .replace(/<img\b[^>]*>/gi, m => /\ssrc\s*=\s*["']?https:\/\//i.test(m) ? m : "")
      .replace(/(href|src)\s*=\s*(["'])\s*(javascript|vbscript|data):[^"']*\2/gi, "$1=$2#$2").replace(/(href|src)\s*=\s*(javascript|vbscript|data):[^\s>]*/gi, '$1="#"');
    return s.replace(/\s{2,}/g, " ").trim();
  }
  const sigText = h => String(h || "").replace(/<br\s*\/?>|<\/(p|div)>/gi, "\n").replace(/<[^>]+>/g, "").replace(/&nbsp;/g, " ").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, "&").replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
  /* the start email with the sender's signature: at {signature} when the template has it, else at the end; blank removes the token */
  function withSignature(body, sig, html) {
    const s = txt(sig), b = String(body || ""), blk = html ? (s ? `<div style="margin:12pt 0 0">${s}</div>` : "") : (s ? `\n\n${sigText(s)}` : "");
    return b.includes("{signature}") ? b.replace(/\{signature\}/g, blk) : b + blk;
  }
  /* The start email's directions map is kept in the workbook (document setting ivkStartMap), not in the template: a picture is far
     bigger than an Excel cell holds (32,767 characters). A template pasted with the map in it gives the picture up to be saved as
     the map and {map} takes its place. Pictures that can't travel in an email (cid:, blob:, file:, local paths) become {map} too, so
     the loaded map shows there; web pictures (https) stay as they are. Returns { html, img: { type, b64 } | null, lost }. */
  function tplImages(html) {
    let img = null, lost = 0, token = /\{map\}/.test(String(html || ""));
    const out = String(html || "").replace(/<img\b[^>]*>/gi, tag => {
      if (/\ssrc\s*=\s*["']?\s*https:\/\//i.test(tag)) return tag;
      const m = tag.match(/\ssrc\s*=\s*(["'])\s*data:(image\/(?:png|jpe?g|gif));base64,([A-Za-z0-9+/=\s]+)\1/i);
      if (m) { if (!img) img = { type: m[2].toLowerCase().replace("jpg", "jpeg"), b64: m[3].replace(/\s+/g, "") }; } else lost++;
      if (token) return ""; token = true; return "{map}";
    });
    return { html: out, img, lost };
  }
  /* one {map} in the body when a map is loaded (after the schedule when the template has no {map}, as saved templates don't); none without */
  function withMap(body, has) {
    const b = String(body || "");
    if (!has) return b.replace(/<p[^>]*>\s*\{map\}\s*<\/p>\s*/gi, "").replace(/\{map\}\n?/g, "");
    if (b.includes("{map}")) { let n = 0; return b.replace(/\{map\}/g, () => n++ ? "" : "{map}"); }
    if (b.includes("{schedule}")) return b.replace("{schedule}", "{schedule}\n{map}");
    return b.includes("{signature}") ? b.replace("{signature}", "{map}\n{signature}") : b + "\n{map}";
  }

  return { DEFAULTS, VISIT_COLS, TEMPLATES, norm, clinicFromName, USER_COLS, userKey, userRows, pickOrder, cleanName, wbKeyOf, whoRec, whoToday, whoPick, ADMIN_MIN, adminOk, CLINIC_WIDE, lockedCfg, sigFromFile, withSignature, tplImages, withMap, txt, num, serialOf, dateOf, isoOf, fromIso, toSerial, toMin, fmtT, fmtD, fmtDay, weekday, weekStart,
    PT_COLS, ptRows, AV_LEVELS, avName, availGrid, availStr, canCome, prefCells, awayList, awayStr, awayOn, PAY_COLS, REMIND_METHODS, money, amt, payRows, dueOf, payOf, ledger, toConfirm,
    PARTS, partKey, dueParts, visitAnswered, PROFILE_COLS, COVER_MODES, modeOf, modeParts, isSelf, SOHAR_KEYS, profileRows, coverage, profileGaps, parseSohar, planSoharImport, demand,
    fullPlan, eligibleStarts, memoX, indexState, indexShort, nextStart, XFER_COLS, xferRows, xferState, elsewhere, visitCounts, SOONER_COLS, soonerRows, soonerKey, soonerMatch, openSlots, freedTimes, ELSE_COLS, elseRows, sameLoc, isOurs, bookedAt, prefCols, indexN: cfg => Math.max(1, Math.round(+cfg.IndexPerWeek * +cfg.IndexWeeks)),
    treatments, txNo, isOak, oakState, WAIT_ON, waitlist, waitMove, provFull, fullName, firstName, scheduleOf, START_TPL, longWhen,
    lasfir, provName, mrnKey, pickRow, eventShort, freqDays, byHeader, patients, appts, isCancelled, visitKey, visitRows, phase, isActive, futureOf, asks,
    lastSeen, tagKey, parseTag, tagStr, tagsOf, askProvider, REPLY_COLS, replyRows, replyAction, repliesDue,
    today, code, chairState, chairStatus, askDue, chairList, chairCount, needs, loud, consults, isStartedStage, keptStartOf, START_COLS, startKey, startRows, consultStarts, startList, fill, protoWords, arrivalText, askText,
    parseReport, shrink, week, DAYS5, dayNum, sameProv, OTHER_COLS, otherRows, isIvkVisit, FLAG_COLS, flagRows, PLAN_COLS, planRows, HOLD_COLS, holdRows,
    onList, holding, holdState, holdConfirm, EX_COLS, CLOSED, exRows, exFor, closureOn, closureClashes, withExceptions, exClashes, MISS_COLS, missRows, isKpVa, missList, lateKeys, newCancels, missImpact, missByPatient, stepDue, daySlots, postCheckSlots, availOf, availText, fits, planIndex, altSlots, capacity, plannerLayout, plannerGrid };
});
