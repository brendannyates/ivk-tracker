/* core.js: the IVK Tracker floor rules. Pure functions (no Office.js, no DOM), shared with tests/.
   Reads the rows excel.js hands over (Master Patient Tracker, Appt List, Consults, the add-in's own IVK Visits table) and
   derives everything the screens show: today's visits and their phase, chair states, what each patient still needs, and the
   Teams messages, and the index planner on top of the provider templates (template.js). Patients are only ever named by LASFIR code. */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory(require("./template.js"));
  else root.IvkCore = factory(root.IvkTemplate);
})(typeof self !== "undefined" ? self : this, function (TM) {
  "use strict";

  /* Defaults for tblIvkConfig. Castro: 4 chairs, each IVK visit holds a chair 2 hours. The SOP has the clinician check in within
     the first 30 min and again in the last 5-10 min. */
  const DEFAULTS = { Enabled: "TRUE", Clinic: "Castro", Chairs: "1, 2, 3, 4", VisitMin: 120, ProviderCheckMin: 30, LateMin: 10, SoonMin: 10,
    AheadWeeks: 6, ProtocolDays: 90, AuthWarnDays: 14, EncountersWarn: 3,
    /* index: 2 visits in each 7 days from treatment 1, for 3 weeks; at least a day off between visits, no more than 5 days apart
       when the slots allow; the provider's post check is in the last 30 min of the 2 hours; a hold is booked in NextGen within 10 min */
    IndexPerWeek: 2, IndexWeeks: 3, MinGapDays: 2, MaxGapDays: 5, PostCheckMin: 90, HoldMin: 10, AfterIndex: "1x per week", CapacityWeeks: 4 };
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
    return fromIso(s);
  }
  /* a time cell: a day fraction (or a full serial), "1:30 PM", "13:30"; minutes after midnight, else null */
  function toMin(v) {
    const n = num(v); if (n != null) { const f = n % 1; return Math.round(f * 1440) % 1440; }
    const m = txt(v).match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*([AaPp])?/); if (!m) return null;
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
  /* A patient listed twice: the active row (Ongoing, Induction) wins, else the later row. excel.js picks the same row to write. */
  function pickRow(rows, mrnOf, statusOf, mrn) {
    let best = -1;
    rows.forEach((r, i) => { if (mrnKey(mrnOf(r)) !== mrnKey(mrn)) return; if (best < 0 || ACTIVE.includes(txt(statusOf(r)).toLowerCase()) || !ACTIVE.includes(txt(statusOf(rows[best])).toLowerCase())) best = i; });
    return best;
  }
  /* Master Patient Tracker rows, one per MRN (pickRow). i = the row's index in the table body (for writes back). */
  function patients(values) {
    const { rows } = byHeader(values), keep = new Set();
    const live = rows.filter(x => txt(x.get("mrn")));
    live.forEach(x => keep.add(pickRow(live, y => y.get("mrn"), y => y.get("status"), x.get("mrn"))));
    return live.filter((x, j) => keep.has(j)).map(x => {
      const g = x.get, el = g("encountersremaining"), ea = g("encountersauthorized");
      return {
        i: x.i, mrn: mrnKey(g("mrn")), code: lasfir(g("patientname")), status: txt(g("status")),
        authExp: toSerial(g(["authexpdate", "authexp"])), encLeft: num(el), encLeftTxt: txt(el), encAuth: num(ea), encAuthTxt: txt(ea),
        copay: txt(g("copay")), infusion: txt(g("infusion")), protoAt: toSerial(g(["protocolupdated", "protocolexpdate"])),
        freq: txt(g(["frequency", "txfrequency"])), loc: txt(g("primarylocation")),
        ins1: txt(g(["primaryencounterinsurance", "primaryinsurance"])), ins2: txt(g(["infusioninsurance", "secondaryinsurance"])),
        prefs: PREF_DAYS.map(([d, a, p]) => ({ d, am: num(g(a)) === 1, pm: num(g(p)) === 1 })),
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
  /* phase of one visit: gone (not coming), noshow (NextGen says so), done (checked out), chair, waiting (checked in), booked */
  function phase(v) {
    if (v.st.status === "cancelled") return "gone";
    if (v.st.out != null) return "done";
    if (v.st.seat != null && v.st.chair) return "chair";
    if (v.st.arr != null) return "waiting";
    if (/no ?show/i.test(v.a.status)) return "noshow";
    return "booked";
  }
  const isActive = p => !!p && ACTIVE.includes(p.status.toLowerCase());
  /* future visits for a patient after a day, from the appointment list (cancelled ones don't count) */
  const futureOf = (A, mrn, after) => A.filter(a => a.mrn === mrn && a.date > after && !isCancelled(a)).sort((x, y) => x.date - y.date || (x.time ?? 0) - (y.time ?? 0));

  /* Does this patient need booking out, and is the protocol due a confirmation? x = { today, cfg, appts }.
     book: nothing booked after today, or booked through less than AheadWeeks out. proto: never confirmed or older than ProtocolDays. */
  function asks(p, x, st) {
    const c = x.cfg, fut = futureOf(x.appts, p?.mrn ?? "", x.today), thru = fut.length ? fut[fut.length - 1].date : null;
    const ahead = +c.AheadWeeks * 7, protoDue = !p || p.protoAt == null || x.today - p.protoAt > +c.ProtocolDays;
    const prn = /prn/i.test(p?.freq ?? "");
    const book = !!p && isActive(p) && !prn && (thru == null || thru < x.today + ahead) && !(st && /^(yes|no)$/.test(st.booked));
    const proto = !!p && isActive(p) && protoDue && !(st && st.proto);
    return { book, proto, next: fut[0] || null, thru, any: book || proto };
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

  /* Chair state on the map. open: empty; track: in the chair, more than SoonMin left; soon: SoonMin or less left;
     ready: the expected end has passed, so check-out is the next step. Red is never on a chair: only a patient checked in and
     waiting LateMin+ for a chair is waiting on staff. */
  function chairState(chair, list, now, cfg) {
    const v = list.find(y => y.ph === "chair" && y.st.chair === chair);
    if (!v) return { state: "open", v: null };
    const end = v.st.seat + +cfg.VisitMin, left = end - now, check = v.st.seat + +cfg.ProviderCheckMin;
    if (left <= 0) return { state: "ready", v, end, left, check };
    return { state: left <= +cfg.SoonMin ? "soon" : "track", v, end, left, check };
  }
  function chairStatus(cs) {
    if (cs.state === "ready") return cs.left < -2 ? `Due out ${-cs.left} min ago` : "Ready to check out";
    if (cs.state === "soon") return `Out in ${cs.left} min`;
    if (cs.state === "track") return `Out around ${fmtT(cs.end)}`;
    return "";
  }
  const waitMin = (v, now) => v.st.arr == null ? 0 : now - v.st.arr;
  const isLate = (v, now, cfg) => v.ph === "waiting" && waitMin(v, now) >= +cfg.LateMin;

  /* chairs: the map's bays when a map is loaded, else the Chairs setting ("1, 2, 3") */
  const chairList = (cfg, plan) => plan ? plan.bays.map(b => b.id) : txt(cfg.Chairs).split(/[,;\s]+/).filter(Boolean);

  /* ---------- needs (Patients) ---------- */
  /* What one patient needs, sorted red, amber, then quiet. Red is kept for what blocks the next infusion: an expired
     authorization, no visits left on it, or an active patient with nothing booked. Each need carries the action that resolves it. */
  function needs(p, x) {
    const n = [], add = (k, t, act, o) => n.push(Object.assign({ k, t, act }, o)), c = x.cfg, T = x.today;
    if (/discontinued/i.test(p.status)) return n;
    const A = isActive(p), a = asks(p, x, null), fut = futureOf(x.appts, p.mrn, T - 1);
    if (A && p.authExp != null && p.authExp < T) add("auth", `Auth ended ${fmtD(p.authExp)}`, ["Open in the tracker", "open"], { bad: true });
    else if (A && p.authExp != null && p.authExp <= T + +c.AuthWarnDays) add("auth", `Auth ends ${fmtD(p.authExp)}`, ["Open in the tracker", "open"]);
    if (A && p.encLeft != null && p.encLeft <= 0) add("enc", "No visits left on the auth", ["Open in the tracker", "open"], { bad: true });
    else if (A && p.encLeft != null && p.encLeft <= +c.EncountersWarn) add("enc", `${p.encLeft} visit${p.encLeft === 1 ? "" : "s"} left on the auth`, ["Open in the tracker", "open"]);
    if (A && !/prn/i.test(p.freq) && !fut.length) add("book", "Nothing booked", ["Ask the group chat", "ask"], { bad: true });
    else if (A && a.book) add("book", a.thru == null ? "Nothing booked after today" : `Booked through ${fmtD(a.thru)}`, ["Ask the group chat", "ask"]);
    if (A && !p.freq) add("freq", "No frequency on the tracker", ["Ask the group chat", "ask"]);
    if (A && a.proto) add("proto", p.protoAt == null ? "Protocol not confirmed" : `Protocol last confirmed ${fmtD(p.protoAt)}`, ["Confirm protocol", "proto"]);
    if (A && p.encAuth != null && p.authStart == null) add("authstart", "No auth start date", ["Open in the tracker", "open"]);
    if (/induction/i.test(p.status)) add("induction", "Induction", null, { quiet: true });
    if (/paused/i.test(p.status)) add("paused", "Paused", null, { quiet: true });
    if (p.loc && x.clinic && norm(p.loc) !== norm(x.clinic)) add("home", `Home clinic ${p.loc}`, null, { quiet: true });
    if (/self ?pay/i.test(p.ins2)) add("selfpay", "Self-pay infusion", null, { quiet: true });
    if (p.reqPost) add("post", "Requires post-check", null, { quiet: true });
    const rank = y => y.quiet ? 2 : y.bad ? 0 : 1;
    return n.map((y, i) => [y, i]).sort((u, w) => rank(u[0]) - rank(w[0]) || u[1] - w[1]).map(u => u[0]);
  }
  const loud = list => list.filter(y => !y.quiet);

  /* ---------- Consults ---------- */
  /* Consults rows: Location | Referral Date | Start Date | Patient | Patient MRN | Consult Date | Provider | Patient Journey Stage |
     Individual PJ Stage | Days in current Stage | Case Management Latest Status Note */
  function consults(values, today) {
    const { rows } = byHeader(values);
    return rows.filter(x => txt(x.get("patient")) || txt(x.get("patientmrn"))).map(x => {
      const g = x.get, stage = txt(g("individualpjstage")) || txt(g("patientjourneystage")), days = num(g("daysincurrentstage"));
      const c = { i: x.i, loc: txt(g("location")), code: lasfir(g("patient")), mrn: mrnKey(g("patientmrn")), ref: toSerial(g("referraldate")),
        start: toSerial(g("startdate")), consult: toSerial(g("consultdate")), provider: provName(g("provider")), stage, days };
      const st = stage.toLowerCase();
      c.state = /not proceeding/.test(st) ? "closed" : /scheduled/.test(st) && !/not yet scheduled/.test(st) ? "done"
        : /not yet scheduled/.test(st) ? "act" : /unclear/.test(st) ? "act" : "wait";
      c.flag = c.state === "act" ? (/unclear/.test(st) ? "Ask the provider whether to refer" : "Book the treatment start") : "";
      return c;
    });
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

  /* ---------- NextGen "Appointments Listing" export -> Appt List rows ---------- */
  /* Report headers (row 7 of the export): MRN | Location | Appointment Date | Time | Provider | Appointment Type | Patient | Status.
     Out: [MRN, Patient, Date serial, Time fraction, Provider, Event, Status, Location], the Appt List table's order. */
  const REPORT_HDR = { mrn: ["mrn", "mdrc"], patient: ["patient", "patientname", "patname"], date: ["appointmentdate", "apptdate", "apptdt", "date"],
    time: ["time", "begtm", "appointmenttime"], provider: ["provider", "renderingprovider", "resource"], event: ["appointmenttype", "event", "eventtype"],
    status: ["status", "sts", "appointmentstatus"], loc: ["location", "loc"] };
  const REPORT_ORDER = ["mrn", "patient", "date", "time", "provider", "event", "status", "loc"];
  function parseReport(aoa) {
    const hr = aoa.slice(0, 40).findIndex(r => r.some(v => REPORT_HDR.mrn.includes(norm(v))) && r.some(v => REPORT_HDR.date.includes(norm(v))));
    if (hr < 0) return null;
    const hn = aoa[hr].map(norm), idx = REPORT_ORDER.map(k => hn.findIndex(h => REPORT_HDR[k].includes(h)));
    if (["mrn", "patient", "date", "time", "status"].some(k => idx[REPORT_ORDER.indexOf(k)] < 0)) return null;
    const rows = [];
    for (const r of aoa.slice(hr + 1)) {
      const v = idx.map(i => i < 0 ? "" : r[i] ?? ""), d = toSerial(v[2]), t = toMin(v[3]);
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
    const vm = +X.cfg.VisitMin, pm = +X.cfg.PostCheckMin, key = TM.provKey;
    const cells = (X.cellsOn ? X.cellsOn(d) : []).map(c => ({ ...c, date: d, by: null }));
    /* the provider's cell a time falls in (a 9:15 visit is in the 9:00 half hour) */
    const cellAt = (pk, t) => cells.find(c => c.pkey === pk && c.time <= t && t < c.time + (c.len || 30));
    const B = X.appts.filter(a => a.date === d && a.time != null && !isCancelled(a) && !/no ?show/i.test(a.status))
      /* a booked visit that was one of the patient's holds is an index visit, whatever NextGen calls it */
      .map(a => ({ mrn: a.mrn, code: a.code, time: a.time, provider: a.provider, pkey: key(a.provider), kind: "appt", event: a.event,
        post: !!(X.needsPost && X.needsPost(a.mrn, a.event)) || (X.holds || []).some(h => h.mrn === a.mrn && h.date === d && h.time === a.time && h.status !== "released") }))
      .concat((X.holds || []).filter(h => h.date === d && holding(h, X.appts, X.listAt)).map(h => ({ mrn: h.mrn, code: h.code, time: h.time, provider: h.provider, pkey: key(h.provider), kind: "hold", h, post: true })));
    const slots = (X.slotsOn ? X.slotsOn(d) : []).map(s => ({ ...s, date: d, by: null, postBy: null, full: false }));
    for (const b of B) {
      const s = slots.find(x => !x.by && x.pkey === b.pkey && x.time <= b.time && b.time < x.time + (x.len || 30)); if (s) s.by = b;
      const c = cellAt(b.pkey, b.time); if (c && !c.by) c.by = b;
      if (b.post) { const pc = cellAt(b.pkey, (s ? s.time : b.time) + pm); if (pc && !pc.by && (pc.use === "post" || pc.use === "flexpost")) pc.by = Object.assign({}, b, { kind: b.kind === "hold" ? "holdpost" : "post" }); }
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
  function capacity(o, weeks) {
    const until = o.today + 7 * weeks; let n = 0, holds = (o.X.holds || []).slice(), open = 0;
    for (let d = o.today + 1; d <= until; d++) { const wd = weekday(d); if (wd >= 1 && wd <= 5) open += daySlots(d, o.X).slots.filter(s => s.openIndex).length; }
    for (; n < 25; n++) {
      const r = planIndex(Object.assign({}, o, { start: o.today + 1, avail: null, fixed: [], pins: {}, X: Object.assign({}, o.X, { holds }) }));
      if (r.items.length < r.N || r.items[0].date > until) break;
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

  return { DEFAULTS, VISIT_COLS, TEMPLATES, norm, txt, num, serialOf, dateOf, isoOf, fromIso, toSerial, toMin, fmtT, fmtD, fmtDay, weekday, weekStart,
    lasfir, provName, mrnKey, pickRow, eventShort, freqDays, byHeader, patients, appts, isCancelled, visitKey, visitRows, phase, isActive, futureOf, asks,
    today, code, chairState, chairStatus, waitMin, isLate, chairList, needs, loud, consults, fill, protoWords, arrivalText, askText,
    parseReport, shrink, week, DAYS5, dayNum, sameProv, OTHER_COLS, otherRows, isIvkVisit, FLAG_COLS, flagRows, PLAN_COLS, planRows, HOLD_COLS, holdRows,
    onList, holding, holdState, daySlots, postCheckSlots, availOf, availText, fits, planIndex, altSlots, capacity, plannerLayout, plannerGrid };
});
