/* template.js: provider schedule templates in the Clinic Template Tool's format. Pure functions (no Office.js, no DOM).
   A template is a weekly grid on a visible tab: a Time column, Monday..Friday headers, and a slot type in each half hour, with
   Location:, Provider:, Last Modified: and Effective From: labels in the rows above (the layout of "a clinic's provider-configurable
   template"). Slot types come from a Slot Types tab: Slot type | Kind | Accepts | Notes (the Clinic Template Tool's
   columns) plus IVK use, Fill and Font. Written to be shared with the Clinic Template Tool later; nothing here is IVK-only except
   the IVK use column. */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.IvkTemplate = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  const txt = v => String(v ?? "").trim();
  const lc = v => txt(v).toLowerCase();
  const DAY_RE = [/^sun(day)?$/i, /^mon(day)?$/i, /^tue(s|sday)?$/i, /^wed(nesday)?$/i, /^thu(r|rs|rsday)?$/i, /^fri(day)?$/i, /^sat(urday)?$/i];
  const DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  const TALLY_DAYS = ["M", "T", "W", "TH", "F"];
  /* a time cell: a day fraction, "1:30 PM", "13:30"; minutes after midnight, else null */
  function toMin(v) {
    if (typeof v === "number" && isFinite(v)) return v >= 0 && v < 1 ? Math.round(v * 1440) % 1440 : null;
    const m = txt(v).match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*([AaPp])?\.?\s*[Mm]?\.?$/); if (!m) return null;
    let h = +m[1]; if (m[3]) { if (/p/i.test(m[3]) && h < 12) h += 12; if (/a/i.test(m[3]) && h === 12) h = 0; }
    return h < 24 && +m[2] < 60 ? h * 60 + +m[2] : null;
  }
  const fmtT = m => m == null ? "" : `${(Math.floor(m / 60) + 11) % 12 + 1}:${String(m % 60).padStart(2, "0")} ${m < 720 ? "AM" : "PM"}`;
  const toSerial = v => {
    if (typeof v === "number" && v > 20000 && v < 80000) return Math.floor(v);
    const s = txt(v); let m;
    if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/))) return Math.round((Date.UTC(+m[3] < 100 ? 2000 + +m[3] : +m[3], +m[1] - 1, +m[2]) - Date.UTC(1899, 11, 30)) / 864e5);
    if ((m = s.match(/^(\d{4})-(\d\d)-(\d\d)/))) return Math.round((Date.UTC(+m[1], +m[2] - 1, +m[3]) - Date.UTC(1899, 11, 30)) / 864e5);
    return null;
  };

  /* ---------- providers ---------- */
  /* the key a template and a NextGen row share: the last name. "Lark MD, Ellis", "Ellis Lark, MD", "Lark", "Dr. Lark" -> "lark" */
  /* credentials and suffixes; DO, PA and BC only in capitals, since Do and Pa are also surnames */
  const CRED = /\b(md|np|pmhnp|fnp|aprn|rn|phd|psyd|dr|jr|sr|ii|iii)\b/gi, CRED_UC = /\b(DO|PA|BC)\b/g;
  const strip = x => x.replace(/\./g, "").replace(CRED, " ").replace(CRED_UC, " ");
  function provKey(name) {
    const s = txt(name); if (!s) return "";
    const parts = s.split(","), head = strip(parts[0]).replace(/[^A-Za-z' -]/g, " ").trim().split(/\s+/).filter(Boolean);
    const rest = strip(parts.slice(1).join(" ")).replace(/[^A-Za-z' -]/g, " ").trim().split(/\s+/).filter(Boolean);
    if (!head.length) return lc(rest[rest.length - 1] || "");
    /* "Lark MD, Ellis" has the last name first; "Ellis Lark, MD" (nothing but credentials after the comma) has it last */
    return lc(rest.length ? head[0] : head[head.length - 1]);
  }

  /* ---------- slot types ---------- */
  /* IVK use: start (an infusion start; index when a post check lines up 90 min later), followup (an infusion start with no post
     check needed), post (the provider's last 30 min with an index patient), flexpost (the same, or a 30-min visit outside IVK when
     no index patient needs it), consult, other. Fill and Font are the colours on the tab and in the add-in. */
  const USES = ["start", "followup", "post", "flexpost", "consult", "other"];
  const USE_LABEL = { start: "IVK start/index", followup: "IVK follow-up", post: "IVK post check", flexpost: "Flex post check", consult: "Consult", other: "Not IVK" };
  /* the clinic template's types and colours; Accepts names NextGen events (comma = any of them) */
  const DEFAULT_TYPES = [
    ["Open Schedule", "Clinical", "*", "Any visit", "other", "#EBF1DE", "#1B211F"],
    ["IV Ketamine Start/Index", "Clinical", "IV Ketamine Treatment Start, IV Ketamine Follow Up", "Index visits first; follow-ups can book here too", "start", "#166569", "#FFFFFF"],
    ["IV Ketamine Follow Up", "Clinical", "IV Ketamine Follow Up", "", "followup", "#008080", "#FFFFFF"],
    ["IVK Post Check", "Clinical", "", "The last 30 min of an index visit", "post", "#808000", "#FFFFFF"],
    ["Flex Post Check", "Clinical", "TMS Follow Up, Telepsych", "A post check, or TMS F/U or tele when no index patient needs it", "flexpost", "#BD9557", "#FFFFFF"],
    ["IV Ketamine Consult", "Clinical", "IV Ketamine Consult", "", "consult", "#2F75B5", "#FFFFFF"],
    ["Interventional Consult", "Clinical", "Interventional Consult", "", "consult", "#6479F0", "#FFFFFF"],
    ["Telepsych", "Clinical", "Telepsych", "", "other", "#FF1493", "#FFFFFF"],
    ["New Patient", "Clinical", "New Patient", "", "other", "#FFFF00", "#1B211F"],
    ["TMS Follow Up", "Clinical", "TMS Follow Up", "", "other", "#DDEBF7", "#1B211F"],
    ["Meeting", "Blocked", "", "", "other", "#66FF33", "#1B211F"],
    ["Admin", "Blocked", "", "", "other", "#00FF00", "#1B211F"],
    ["Lunch", "Blocked", "", "Not counted in the total", "other", "#1E7B34", "#FFFFFF"]
  ].map(([type, kind, accepts, notes, use, fill, font]) => ({ type, kind, accepts, notes, use, fill, font }));
  const TYPE_COLS = ["Slot type", "Kind", "Accepts", "Notes", "IVK use", "Fill", "Font"];
  const HDR = { type: /^slot\s*type$|^type$/i, kind: /^kind$/i, accepts: /^accepts?/i, notes: /^notes?$/i, use: /^ivk\s*use$/i, fill: /^fill$/i, font: /^font$/i };
  const hex = v => /^#?[0-9a-f]{6}$/i.test(txt(v)) ? "#" + txt(v).replace("#", "").toUpperCase() : "";
  const guessUse = t => { const s = lc(t);
    if (/flex/.test(s) && /post|\bpc\b/.test(s)) return "flexpost"; if (/post\s*check|\bpc\b/.test(s)) return "post";
    if (/(ivk|ketamine)/.test(s) && /(start|index)/.test(s)) return "start"; if (/(ivk|ketamine)/.test(s) && /follow|f\/u/.test(s)) return "followup";
    if (/consult/.test(s)) return "consult"; return "other"; };
  /* Slot Types rows (header row first) -> types; a type missing from the tab still gets a guessed use and a neutral colour */
  function typeRows(values) {
    const V = values || []; if (!V.length) return [];
    const at = {}; V[0].forEach((h, i) => { for (const [k, re] of Object.entries(HDR)) if (at[k] == null && re.test(txt(h))) at[k] = i; });
    if (at.type == null) return [];
    return V.slice(1).map(r => { const g = k => at[k] == null ? "" : txt(r[at[k]]), type = g("type"); if (!type) return null;
      const use = USES.includes(lc(g("use"))) ? lc(g("use")) : guessUse(type), d = DEFAULT_TYPES.find(x => lc(x.type) === lc(type));
      return { type, kind: /^block|non|no$/i.test(g("kind")) ? "Blocked" : "Clinical", accepts: g("accepts"), notes: g("notes"), use,
        fill: hex(g("fill")) || d?.fill || "#E9EBE6", font: hex(g("font")) || d?.font || "#1B211F" }; }).filter(Boolean);
  }
  const typeOf = (types, name) => types.find(t => lc(t.type) === lc(name)) || { type: txt(name), kind: /\b(admin|lunch|meeting|break|blocked|pto)\b/i.test(name) ? "Blocked" : "Clinical", accepts: "", notes: "", use: guessUse(name), fill: "#E9EBE6", font: "#1B211F", unknown: true };
  /* Accepts: event names separated by commas; a term joined with + needs every part; * is any visit. Parts match the event name,
     or IVK/TMS/ESK as a word in it (the Clinic Template Tool's rule, on the one field the report gives). */
  const acceptTerms = a => txt(a).split(/[,;\n]/).map(x => x.split("+").map(lc).filter(Boolean)).filter(x => x.length);
  function accepts(t, event) {
    const e = lc(event), terms = acceptTerms(t.accepts); if (!e) return false;
    const line = /\b(iv ketamine|ivk)\b/.test(e) ? "ivk" : /\btms\b/.test(e) ? "tms" : /\b(esk|esketamine|spravato)\b/.test(e) ? "esk" : "";
    return terms.some(term => term.every(p => p === "*" || p === e || p === line));
  }
  const eventsOf = t => acceptTerms(t.accepts).map(x => x.join(" + "));

  /* ---------- grids ---------- */
  const LABELS = { location: /^location\s*:?$/i, provider: /^provider\s*:?$/i, modified: /^last\s*modified\s*:?$/i, from: /^effective\s*(from|start|date)?\s*:?$/i, to: /^effective\s*(to|through|thru|end)\s*:?$/i };
  /* Every weekly grid on a tab, as the Clinic Template Tool finds them: a row with Monday..Friday headers, the time column to its
     left, slot types in the cells, labels in the 5 rows above. cells: Map "day|minutes" -> { type, r, c } (0-based on the tab). */
  function parseGrids(sheet, values) {
    const V = values || [], out = [], used = new Set(), cell = (r, c) => (V[r] && V[r][c] != null ? V[r][c] : "");
    for (let r = 0; r < V.length; r++) for (let c = 1; c < (V[r] || []).length; c++) {
      if (used.has(r + ":" + c) || !DAY_RE[1].test(txt(cell(r, c)))) continue;
      const cols = []; for (let k = 0; k < 7; k++) { const d = DAY_RE.findIndex(re => re.test(txt(cell(r, c + k)))); if (d < 0) break; cols.push({ c: c + k, d }); }
      if (cols.length < 3 || cols[1].d !== 2) continue;
      cols.forEach(x => used.add(r + ":" + x.c));
      const tc = c - 1, rows = [];
      for (let rr = r + 1; rr < V.length && rr < r + 80; rr++) { const t = toMin(cell(rr, tc)); if (t == null) { if (rows.length) break; continue; } rows.push({ r: rr, t }); }
      if (rows.length < 2) continue;
      const steps = rows.slice(1).map((x, i) => x.t - rows[i].t).filter(x => x > 0).sort((a, b) => a - b), step = steps.length ? steps[Math.floor(steps.length / 2)] : 30;
      const cells = new Map();
      for (const x of rows) for (const col of cols) { const v = txt(cell(x.r, col.c)); if (v) cells.set(col.d + "|" + x.t, { type: v, r: x.r, c: col.c }); }
      const hdr = {}, at = {}; let title = "", top = r;
      for (let hr = Math.max(0, r - 5); hr <= r; hr++) for (let hc = Math.max(0, tc - 1); hc <= tc + cols.length + 2; hc++) {
        const v = txt(cell(hr, hc)); if (!v) continue;
        for (const [k, re] of Object.entries(LABELS)) if (re.test(v) && hdr[k] == null) {
          /* the label's value: the first filled cell to its right, inside the grid (and the spacer column after it), before the next label, so a
             blank Last Modified never reads (or later writes) a neighbouring label's value or the tally beside the grid */
          for (let j = hc + 1; j <= tc + cols.length + 1; j++) { const w = cell(hr, j); if (!txt(w)) continue; if (Object.values(LABELS).some(q => q.test(txt(w)))) break; hdr[k] = w; at[k] = [hr, j]; break; }
          if (hdr[k] == null) { hdr[k] = ""; at[k] = [hr, hc + 1]; } top = Math.min(top, hr);
        }
      }
      for (let hr = Math.max(0, r - 6); hr < r && !title; hr++) { const v = txt(cell(hr, tc)); if (v && /[a-z]{2,}/i.test(v) && !Object.values(LABELS).some(q => q.test(v)) && !/^(weekly\s+)?(clinic\s+)?(template|schedule)$/i.test(v)) { title = v; at.title = [hr, tc]; top = Math.min(top, hr); } }
      const label = txt(hdr.provider) || title;
      out.push({ sheet, anchor: [r, c], label, title, key: provKey(label || title), location: txt(hdr.location), at, top, bottom: rows[rows.length - 1].r,
        timeCol: tc, dayCols: cols, rows, step, effFrom: toSerial(hdr.from), effTo: toSerial(hdr.to), modified: toSerial(hdr.modified), cells });
    }
    return out;
  }
  /* the provider's grid in effect on a date: the latest Effective From on or before it (blank = always), within Effective To */
  function gridFor(grids, key, date) {
    const L = grids.filter(g => g.key === key && (g.effFrom == null || g.effFrom <= date) && (g.effTo == null || g.effTo >= date));
    return L.sort((a, b) => (b.effFrom ?? -1) - (a.effFrom ?? -1))[0] || null;
  }
  const nameOf = g => g.title || g.label;

  /* The template's infusion starts for one weekday, with the post check that lines up postMin later (same provider, same day).
     index: a start with an IVK or flex post check at its end. */
  function startsOn(g, wd, types, postMin) {
    const out = [];
    for (const [k, x] of g.cells) {
      const [d, t] = k.split("|").map(Number); if (d !== wd) continue;
      const ty = typeOf(types, x.type); if (ty.use !== "start" && ty.use !== "followup") continue;
      const pc = g.cells.get(d + "|" + (t + postMin)), pu = pc ? typeOf(types, pc.type).use : null, hasPost = pu === "post" || pu === "flexpost";
      out.push({ day: d, time: t, type: ty.type, use: ty.use, index: ty.use === "start" && hasPost, post: t + postMin, postUse: hasPost ? pu : null, postType: hasPost ? pc.type : "" });
    }
    return out.sort((a, b) => a.time - b.time);
  }
  /* What the template says needs fixing: a start with no post check at its end can't take an index patient */
  function checks(g, types, postMin) {
    const w = [];
    for (let d = 1; d <= 5; d++) for (const s of startsOn(g, d, types, postMin)) if (s.use === "start" && !s.index) w.push(`${DAYS[d]} ${fmtT(s.time)}: no post check at ${fmtT(s.post)}, so it can’t take an index patient`);
    for (const [, x] of g.cells) { const ty = typeOf(types, x.type); if (ty.unknown) { w.push(`“${x.type}” isn’t on the Slot Types tab`); break; } }
    return w;
  }
  /* the Event Types tally beside each grid: slots and hours by type, per weekday; Clinical; Total (everything except Lunch) */
  function tally(g, types) {
    const per = 60 / g.step, rows = types.map(t => { const byDay = [1, 2, 3, 4, 5].map(d => [...g.cells].filter(([k, x]) => +k.split("|")[0] === d && lc(x.type) === lc(t.type)).length);
      const slots = byDay.reduce((a, b) => a + b, 0); return { type: t.type, kind: t.kind, slots, hours: slots / per, byDay }; });
    const sum = L => { const byDay = [0, 1, 2, 3, 4].map(i => L.reduce((a, r) => a + r.byDay[i], 0)), slots = byDay.reduce((a, b) => a + b, 0); return { slots, hours: slots / per, byDay }; };
    return { rows, clinical: sum(rows.filter(r => r.kind === "Clinical")), total: sum(rows.filter(r => !/^lunch$/i.test(r.type))) };
  }

  /* ---------- writing a new grid ---------- */
  /* The clinic template layout at row r0 (0-based), columns A..P. Returns what excel.js writes: values (strings, numbers and formulas by cell),
     merges, the slot range, the tally range and fills. o = { title, label, location, from, modified (serial), days: [1..5 the
     provider works], start, end, step, types, copy: Map of "day|min" -> type } */
  const colL = i => { let s = ""; for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s; return s; };
  const A1 = (r, c) => colL(c) + (r + 1);
  function layout(o, r0) {
    const step = o.step || 30, times = []; for (let t = o.start ?? 480; t < (o.end ?? 1080); t += step) times.push(t);
    const put = [], hdrRow = r0 + 4, first = hdrRow + 1, last = first + times.length - 1, set = (r, c, v) => put.push([r, c, v]);
    set(r0, 0, o.title || o.label);
    set(r0 + 1, 4, "Effective From:"); if (o.from != null) set(r0 + 1, 5, o.from);
    set(r0 + 2, 0, "Location:"); set(r0 + 2, 1, o.location || ""); set(r0 + 2, 2, "Weekly Clinic Template"); set(r0 + 2, 4, "Last Modified:"); set(r0 + 2, 5, o.modified ?? "");
    set(r0 + 3, 0, "Provider:"); [1, 2, 3, 4, 5].forEach(d => set(r0 + 3, d, (o.days || [1, 2, 3, 4, 5]).includes(d) ? o.label : ""));
    set(hdrRow, 0, "Time"); [1, 2, 3, 4, 5].forEach(d => set(hdrRow, d, DAYS[d]));
    times.forEach((t, i) => { set(first + i, 0, t / 1440); [1, 2, 3, 4, 5].forEach(d => set(first + i, d, (o.copy && o.copy.get(d + "|" + t)) || "")); });
    /* tally: I..P, header on the Location row, one row per type, then Clinical and Total, as COUNTIF formulas over this grid */
    const T = o.types || DEFAULT_TYPES, g = `$B$${first + 1}:$F$${last + 1}`, tr0 = r0 + 2, per = 60 / step;
    ["Event Types", "Slots", "Hours by Event", ...TALLY_DAYS].forEach((h, i) => set(tr0, 8 + i, h));
    T.forEach((t, i) => { const r = tr0 + 1 + i, R = r + 1;
      set(r, 8, t.type); set(r, 9, `=COUNTIF(${g},$I${R})`); set(r, 10, `=J${R}/${per}`);
      [1, 2, 3, 4, 5].forEach((d, k) => set(r, 11 + k, `=COUNTIF(${colL(d)}$${first + 1}:${colL(d)}$${last + 1},$I${R})`)); });
    const rc = tr0 + 1 + T.length, rt = rc + 1, clin = T.map((t, i) => t.kind === "Clinical" ? tr0 + 2 + i : null).filter(x => x != null), tot = T.map((t, i) => /^lunch$/i.test(t.type) ? null : tr0 + 2 + i).filter(x => x != null);
    const sumOf = (L, col) => L.length ? `=${L.map(r => colL(col) + r).join("+")}` : "=0";
    set(rc, 8, "Clinical"); set(rt, 8, "Total:");
    for (const [r, L] of [[rc, clin], [rt, tot]]) { [9, 11, 12, 13, 14, 15].forEach(c => set(r, c, sumOf(L, c))); set(r, 10, `=J${r + 1}/${per}`); }
    return { put, merges: [`A${r0 + 1}:P${r0 + 1}`, `C${r0 + 3}:D${r0 + 3}`], slots: `B${first + 1}:F${last + 1}`, tally: `I${tr0 + 1}:P${rt + 1}`,
      types: T.map((t, i) => ({ type: t.type, at: `I${tr0 + 2 + i}` })), heads: [`A${r0 + 3}:A${hdrRow + 1}`, `B${hdrRow + 1}:F${hdrRow + 1}`, `A${first + 1}:A${last + 1}`, `I${tr0 + 1}:P${tr0 + 1}`, `I${rc + 1}:P${rt + 1}`],
      timeFmt: `A${first + 1}:A${last + 1}`, dateFmt: [A1(r0 + 1, 5), A1(r0 + 2, 5)], bottom: Math.max(last, rt), first, last, times };
  }
  /* where the next grid goes on a tab: 3 rows under the lowest one */
  const nextRow = (grids, nTypes = DEFAULT_TYPES.length) => grids.length ? Math.max(...grids.map(g => Math.max(g.bottom, g.anchor[0] - 2 + nTypes + 2))) + 4 : 1;
  /* a layout as one block of rows (A..P) from its top row, for a single write; formulas stay as "=..." strings */
  function block(lay) {
    const r0 = Math.min(...lay.put.map(p => p[0])), r1 = Math.max(...lay.put.map(p => p[0])), rows = Array.from({ length: r1 - r0 + 1 }, () => Array(16).fill(""));
    for (const [r, c, v] of lay.put) rows[r - r0][c] = v;
    return { r0, rows };
  }

  /* ---------- days ---------- */
  const dow = s => ((s - 1) % 7 + 7) % 7;   // Excel serial -> 0 = Sunday
  /* For core.daySlots: each provider's grid in effect on a day, as infusion starts (slotsOn) and as every cell (cellsOn). Cached by day. */
  function onDays(grids, types, postMin) {
    const keys = [...new Set(grids.map(g => g.key).filter(Boolean))], memo = new Map();
    const day = d => {
      if (memo.has(d)) return memo.get(d);
      const wd = dow(d), gs = keys.map(k => gridFor(grids, k, d)).filter(Boolean);
      const v = {
        slots: gs.flatMap(g => startsOn(g, wd, types, postMin).map(s => Object.assign(s, { len: g.step, provider: g.label, pkey: g.key, who: nameOf(g) }))),
        cells: gs.flatMap(g => [...g.cells].filter(([k]) => +k.split("|")[0] === wd).map(([k, x]) => { const t = typeOf(types, x.type);
          return { time: +k.split("|")[1], len: g.step, type: t.type, use: t.use, kind: t.kind, fill: t.fill, font: t.font, provider: g.label, pkey: g.key, who: nameOf(g) }; }))
      };
      memo.set(d, v); return v;
    };
    return { slotsOn: d => day(d).slots, cellsOn: d => day(d).cells, providers: keys.length };
  }

  return { DAYS, TALLY_DAYS, toMin, fmtT, provKey, USES, USE_LABEL, DEFAULT_TYPES, TYPE_COLS, typeRows, typeOf, guessUse, accepts, acceptTerms, eventsOf,
    parseGrids, gridFor, nameOf, startsOn, checks, tally, layout, block, nextRow, onDays, dow, colL, A1 };
});
