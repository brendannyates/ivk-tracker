// ==== bay plan: shared with bay-map-builder/render.js and esk-tracker src/taskpane.html. Keep both copies identical. ====
// A bay map file (schema "mhs-bay-map", v 1) is drawn in its own w × h box. Paths are SVG path data in that box; bay ids are
// the Bay # the tracker saves on a visit ("4", "2A"). bpCheck() is the only way a plan gets drawn or saved: it returns a clean
// copy holding just what the renderer reads (plus the builder's own edit model, untouched), or an error to show.
const BP_SCHEMA = 'mhs-bay-map';
const bpEsc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const bpPct = (n, of) => `${(n / of * 100).toFixed(3)}%`;
const BP_D = /^\s*[Mm][MLHVCSQTAZmlhvcsqtaz0-9eE.,\s+-]*$/;
const BP_MAX = { bays: 40, curtains: 80, walls: 400, doors: 120, windows: 120, tables: 400, chairs: 200, rooms: 60 };
function bpCheck(x){
  const fail = err => ({ err });
  if (!x || typeof x !== 'object' || x.schema !== BP_SCHEMA || x.v !== 1) return fail('This isn’t a bay map file.');
  let size = 0; try { size = JSON.stringify(x).length; } catch { return fail('The map file can’t be read.'); }
  if (size > 400000) return fail('The map file is too large.');
  const num = n => typeof n === 'number' && isFinite(n) && Math.abs(n) <= 20000;
  const pt = p => Array.isArray(p) && p.length === 2 && p.every(num);
  const okD = s => typeof s === 'string' && s.length <= 20000 && BP_D.test(s);
  const str = (s, n) => typeof s === 'string' ? s.slice(0, n) : '';
  if (!num(x.w) || !num(x.h) || x.w < 100 || x.h < 60 || x.w > 6000 || x.h > 6000) return fail('The map size is missing or out of range.');
  const L = {};
  for (const k in BP_MAX){ const a = x[k] == null ? [] : x[k];
    if (!Array.isArray(a)) return fail(`The ${k} list is malformed.`);
    if (a.length > BP_MAX[k]) return fail(`The map has more than ${BP_MAX[k]} ${k}.`); L[k] = a; }
  if (!L.bays.length) return fail('The map has no bays.');
  const ids = new Map(), bays = [];
  for (const b of L.bays){ const id = String(b?.id ?? '').trim();
    if (!/^[A-Za-z0-9][A-Za-z0-9-]{0,7}$/.test(id) || /^bay/i.test(id)) return fail(`Bay number “${id}” can’t be used. Use up to 8 letters or digits, like 4 or 2A.`);
    if (ids.has(id.toLowerCase())) return fail(`Bay ${id} is on the map twice.`);
    if (!num(b.cx) || !num(b.cy) || !num(b.r) || b.r < 10 || b.r > 400 || !pt(b.tag) || (b.lamp != null && !pt(b.lamp))) return fail(`Bay ${id} has a bad position or size.`);
    ids.set(id.toLowerCase(), id); bays.push({ id, cx: b.cx, cy: b.cy, r: b.r, tag: [b.tag[0], b.tag[1]], lamp: b.lamp ? [b.lamp[0], b.lamp[1]] : null }); }
  const curtains = [], walls = [], doors = [], windows = [], tables = [], chairs = [], rooms = [];
  for (const c of L.curtains){ if (!okD(c?.d) || !Array.isArray(c.bays)) return fail('A curtain track is malformed.');
    curtains.push({ d: c.d, bays: [...new Set(c.bays.map(b => ids.get(String(b).trim().toLowerCase())).filter(Boolean))] }); }
  for (const w of L.walls){ if (!okD(w?.d)) return fail('A wall is malformed.'); walls.push(w.thin ? { d: w.d, thin: true } : { d: w.d }); }
  for (const [k, out] of [['doors', doors], ['windows', windows]]) for (const o of L[k]){ if (!okD(o?.d)) return fail(`A ${k.slice(0, -1)} is malformed.`); out.push({ d: o.d }); }
  for (const t of L.tables){
    if (t?.c){ if (!Array.isArray(t.c) || t.c.length !== 3 || !t.c.every(num) || t.c[2] <= 0) return fail('A table is malformed.'); tables.push({ c: [...t.c] }); continue; }
    if (!Array.isArray(t?.rect) || t.rect.length < 4 || t.rect.length > 5 || !t.rect.every(num) || t.rect[2] <= 0 || t.rect[3] <= 0 || (t.a != null && !num(t.a))) return fail('A table or counter is malformed.');
    tables.push(t.a ? { rect: [...t.rect], a: t.a } : { rect: [...t.rect] }); }
  for (const c of L.chairs){ if (!Array.isArray(c) || c.length !== 3 || !c.every(num)) return fail('A chair is malformed.'); chairs.push([...c]); }
  for (const r of L.rooms){ const t = str(r?.t, 40).trim(); if (!t || !num(r.x) || !num(r.y)) return fail('A room label is malformed.'); rooms.push(r.small ? { t, x: r.x, y: r.y, small: true } : { t, x: r.x, y: r.y }); }
  const plan = { schema: BP_SCHEMA, v: 1, name: str(x.name, 60), clinic: str(x.clinic, 20), savedAt: str(x.savedAt, 40), savedBy: str(x.savedBy, 80),
    w: x.w, h: x.h, bays, curtains, walls, doors, windows, tables, chairs, rooms };
  if (x.edit && typeof x.edit === 'object' && !Array.isArray(x.edit)) plan.edit = x.edit;
  return { plan };
}
function bpChair(x, y, a){
  return `<g transform="translate(${x} ${y}) rotate(${a})"><rect class="seat" x="-26" y="-30" width="52" height="60" rx="10"/><rect class="back" x="-26" y="-30" width="52" height="18" rx="9"/><rect class="arm" x="-26" y="-9" width="9" height="34" rx="4"/><rect class="arm" x="17" y="-9" width="9" height="34" rx="4"/></g>`;
}
function bpTable(t){
  if (t.c) return `<circle cx="${t.c[0]}" cy="${t.c[1]}" r="${t.c[2]}"/>`;
  const [x, y, w, h, rx = 0] = t.rect;
  return `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="${rx}"${t.a ? ` transform="rotate(${t.a} ${x + w / 2} ${y + h / 2})"` : ''}/>`;
}
// Under the rugs: curtain tracks (drawn while used(bayId) is true for any bay they screen, else bunched at their first point), doors, walls, windows.
function bpUnder(P, used){
  const curtain = c => `<g class="curtain"><path class="crail" d="${c.d}"/>${c.bays.some(used) ? `<path class="fabric" d="${c.d}"/>` : `<path class="bunch" d="${c.d}" pathLength="100"/>`}</g>`;
  const paths = L => L.map(x => `<path d="${x.d}"/>`).join(''), thin = P.walls.filter(w => w.thin);
  return `${P.curtains.map(curtain).join('')}
        <g class="door">${paths(P.doors)}</g>
        <g class="wall">${paths(P.walls.filter(w => !w.thin))}</g>${thin.length ? `<g class="wall thin">${paths(thin)}</g>` : ''}${P.windows.length ? `<g class="window">${P.windows.map(x => `<path class="wf" d="${x.d}"/><path class="wg" d="${x.d}"/>`).join('')}</g>` : ''}`;
}
// Over the rugs: tables and fixtures, chairs, then each bay's lantern (bays = the plan's bays or a model copy of them; lit(bay) = someone is in it).
function bpOver(P, bays, lit){
  const lamp = b => !b.lamp ? '' : `${lit(b) ? `<circle class="glow" cx="${b.lamp[0]}" cy="${b.lamp[1]}" r="20" fill="url(#bmGlow)"/>` : ''}<g class="lantern${lit(b) ? ' on' : ''}" transform="translate(${b.lamp[0]} ${b.lamp[1]})"><circle class="body" r="7.5"/><circle class="lens" r="4.75"/><circle class="touch" r="1.4"/></g>`;
  return `<g class="table">${P.tables.map(bpTable).join('')}</g>
        ${P.chairs.map(c => bpChair(c[0], c[1], c[2])).join('')}
        <defs><radialGradient id="bmGlow"><stop offset="0" stop-color="var(--glow)" stop-opacity="var(--glow-a)"/><stop offset="1" stop-color="var(--glow)" stop-opacity="0"/></radialGradient></defs>
        ${bays.map(lamp).join('')}`;
}
const bpRooms = P => P.rooms.map(r => `<div class="room${r.small ? ' small' : ''}" style="left:${bpPct(r.x, P.w)};top:${bpPct(r.y, P.h)}">${bpEsc(r.t)}</div>`).join('');
// The built-in map: the two ESK treatment rooms traced from the clinic drawing (808 × 392), used until a clinic saves its own.
const BP_DEFAULT = { schema: BP_SCHEMA, v: 1, name: 'ESK rooms (built-in)', clinic: '', savedAt: '', savedBy: '', w: 808, h: 392,
  bays: [{ id: '1', cx: 118, cy: 298, r: 58, tag: [176, 300], lamp: [72, 253] }, { id: '2', cx: 118, cy: 128, r: 60, tag: [158, 180], lamp: [150, 84] },
    { id: '3', cx: 345, cy: 150, r: 55, tag: [332, 206], lamp: [322, 88] }, { id: '4', cx: 500, cy: 140, r: 58, tag: [572, 170], lamp: [527, 90] },
    { id: '5', cx: 695, cy: 150, r: 58, tag: [684, 230], lamp: [655, 90] }],
  curtains: [{ d: 'M30 207H158C202 207 226 242 226 292L226 364', bays: ['1', '2'] }, { d: 'M204 54C212 122 236 192 268 232C282 240 296 241 316 241L410 241', bays: ['2', '3'] },
    { d: 'M600 54C606 130 592 200 600 262C606 312 610 338 608 364', bays: ['4', '5'] }, { d: 'M424 244C470 246 516 236 554 208', bays: ['4'] }],
  walls: [{ d: 'M22 50H800V372H490M440 372H376M328 372H22V50' }, { d: 'M418 50V372' }],
  doors: [{ d: 'M376 372V324A48 48 0 0 0 328 372' }, { d: 'M440 372V324A50 50 0 0 1 490 372' }], windows: [],
  tables: [{ c: [72, 253, 24] }, { c: [150, 84, 26] }, { c: [62, 162, 22] }, { c: [322, 88, 26] }, { rect: [252, 80, 40, 40, 8] }, { c: [527, 90, 22] }, { c: [450, 185, 25] }, { c: [655, 90, 27] }, { c: [770, 272, 20] }],
  chairs: [[74, 300, -80], [96, 108, -40], [358, 122, 35], [472, 108, -35], [722, 126, 30]],
  rooms: [{ t: 'ESK Room 1', x: 219.776, y: 25.872 }, { t: 'ESK Room 2', x: 611.656, y: 25.872 }] };
// ==== end bay plan ====
if (typeof module !== 'undefined') module.exports = { BP_SCHEMA, BP_MAX, BP_DEFAULT, bpCheck, bpEsc, bpPct, bpChair, bpTable, bpUnder, bpOver, bpRooms };
