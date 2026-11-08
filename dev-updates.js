/* Dev Update bell, shared by the MHS add-ins (keep this file byte-identical across them). dev-updates.json (newest first):
   { "updates": [ { "id": "2026-10-29-a", "date": "2026-10-29", "bullets": ["..."] } ] }
   The bell only shows where the page has a .du-home: each add-in puts one right of its top-left name in full screen and none in
   the task pane. It floats there (draggable; a drop near home snaps back). A red dot means an update from the last 30 days that
   this computer hasn't opened; nothing pops up on its own, a click opens the history and clears the dot.
   Seen ids are kept per computer (localStorage) and also passed between the task pane and its full-screen dialog, because the
   dialog may not share the pane's storage: the pane adds them to the dialog URL (dialogUrl), and each side sends
   { k: "devupd", seen } through its own message channel (the send option) for the other to merge (receive). */
(function () {
  const FRESH_DAYS = 30, MAX_SEEN = 200;
  const S = { key: "", feed: "dev-updates.json", send: null, items: [], seen: new Set(), open: false, el: null };
  const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const fmt = d => { const t = new Date(d + "T12:00:00"); return isNaN(t) ? esc(d) : t.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" }); };
  const fresh = u => (Date.now() - new Date(u.date + "T12:00:00")) / 864e5 <= FRESH_DAYS;
  const unseen = () => S.items.filter(u => !S.seen.has(u.id) && fresh(u));

  function load() {
    try { JSON.parse(localStorage.getItem(S.key) || "[]").forEach(id => S.seen.add(id)); } catch { /* storage unavailable */ }
    const p = new URLSearchParams(location.search).get("devseen");
    if (p) p.split(",").filter(Boolean).forEach(id => S.seen.add(id));
  }
  function save() { try { localStorage.setItem(S.key, JSON.stringify([...S.seen].slice(-MAX_SEEN))); } catch { /* this session only */ } }
  function markSeen(ids) {
    let added = false; ids.forEach(id => { if (!S.seen.has(id)) { S.seen.add(id); added = true; } });
    if (!added) return;
    save(); render();
    if (S.send) try { S.send({ k: "devupd", seen: [...S.seen] }); } catch { /* other side gone */ }
  }

  const BELL = '<svg viewBox="0 0 24 24" width="20" height="20" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M6 16V11a6 6 0 0 1 12 0v5l1.5 2h-15z"/><path d="M10 20a2 2 0 0 0 4 0"/></svg>';
  const list = u => `<ul>${u.bullets.map(b => `<li>${esc(b)}</li>`).join("")}</ul>`;
  function build() {
    const el = document.createElement("div");
    el.className = "devupd"; el.hidden = true;
    el.innerHTML = `<button type="button" class="du-bell" data-tip="Dev updates" aria-haspopup="dialog" aria-expanded="false">${BELL}<span class="du-dot" hidden></span></button>
      <div class="du-panel" role="dialog" aria-label="Dev updates" hidden></div>`;
    document.body.appendChild(el);
    const bell = el.querySelector(".du-bell");
    bell.onclick = () => { if (P.justDragged) { P.justDragged = false; return; } toggle(!S.open); };
    dragOn(el, bell);
    el.addEventListener("click", e => { const b = e.target.closest("[data-du=close]"); if (b) toggle(false); });
    document.addEventListener("click", e => { if (S.open && !e.composedPath().includes(el)) toggle(false); });   // the path, not contains: a re-render may have removed the clicked button
    document.addEventListener("keydown", e => { if (e.key === "Escape" && S.open) { toggle(false); el.querySelector(".du-bell").focus(); } });
    return el;
  }
  function toggle(on) {
    S.open = on;
    if (on) markSeen(unseen().map(u => u.id));
    render();
    if (on) { const c = S.el.querySelector(".du-panel [data-du=close]"); if (c) c.focus(); }
  }
  function render() {
    if (!S.el) return;
    const n = unseen(), bell = S.el.querySelector(".du-bell"), panel = S.el.querySelector(".du-panel");
    bell.querySelector(".du-dot").hidden = !n.length;
    bell.setAttribute("aria-label", n.length ? `Dev updates, ${n.length} new` : "Dev updates");
    bell.setAttribute("aria-expanded", String(S.open));
    panel.hidden = !S.open;
    panel.innerHTML = S.open ? `<div class="du-head"><b>Dev updates</b><button type="button" data-du="close" aria-label="Close">Close</button></div>
      ${S.items.map(u => `<section${S.fresh.has(u.id) ? ' class="du-new"' : ""}><h3>Dev Update: ${fmt(u.date)}${S.fresh.has(u.id) ? ' <span class="du-tag">New</span>' : ""}</h3>${list(u)}</section>`).join("")}` : "";
    soon();
  }

  /* Home is the page's .du-home (the bell centres on it). It can be dragged anywhere and stays there (per computer, anchored
     to the nearest corner so it holds its place on resize); dropped within SNAP px of home it goes back home. Away from home,
     when a new screen or menu puts a button under it, it steps aside to the nearer side of that card and goes back once the
     way is clear. A press without moving opens it. */
  const SIZE = 44, EDGE = 4, GAP = 6, SNAP = 56;
  const P = { saved: null, drag: null, justDragged: false, raf: 0 };
  const posKey = () => `${S.key}_bellpos`;
  function loadPos() { try { const p = JSON.parse(localStorage.getItem(posKey()) || "null"); if (p && isFinite(p.dx) && isFinite(p.dy)) P.saved = p; } catch { /* none */ } }
  function savePos() { try { if (P.saved) localStorage.setItem(posKey(), JSON.stringify(P.saved)); else localStorage.removeItem(posKey()); } catch { /* this session only */ } }
  const clampX = x => Math.min(Math.max(EDGE, x), innerWidth - SIZE - EDGE), clampY = y => Math.min(Math.max(EDGE, y), innerHeight - SIZE - EDGE);
  const homeEl = () => document.querySelector(".du-home");
  function homeAt(h) { const r = h.getBoundingClientRect(); return { x: clampX(r.left + r.width / 2 - SIZE / 2), y: clampY(r.top + r.height / 2 - SIZE / 2) }; }
  function savedAt() { const p = P.saved; return { x: clampX(p.h === "l" ? p.dx : innerWidth - SIZE - p.dx), y: clampY(p.v === "t" ? p.dy : innerHeight - SIZE - p.dy) }; }
  function apply(x, y) {
    const el = S.el; if (!el) return;
    el.style.left = `${x}px`; el.style.top = `${y}px`;
    el.classList.toggle("du-left", x + SIZE / 2 < innerWidth / 2);   // the history opens towards the middle
    el.classList.toggle("du-up", y + SIZE / 2 > innerHeight / 2);
  }
  const INTERACTIVE = "button,a[href],input,select,textarea,summary,label,[data-act],[role=button],[tabindex]:not([tabindex=\"-1\"])";
  function blocker(x, y) {
    const pts = [[x + 6, y + 6], [x + SIZE - 6, y + 6], [x + 6, y + SIZE - 6], [x + SIZE - 6, y + SIZE - 6], [x + SIZE / 2, y + SIZE / 2]];
    for (const [px, py] of pts) {
      const hit = document.elementsFromPoint(px, py).find(n => !S.el.contains(n));
      const b = hit && hit.closest(INTERACTIVE);
      if (b && !b.matches(".scrim,[data-act=scrim]")) return b;
    }
    return null;
  }
  function avoid() {
    P.raf = 0; if (!S.el || P.drag) return;
    const h = homeEl();
    S.el.hidden = !h || !S.items.length;
    if (S.el.hidden) { if (S.open) { S.open = false; render(); } return; }
    if (!P.saved) { const at = homeAt(h); S.el.classList.add("du-athome"); apply(at.x, at.y); return; }
    S.el.classList.remove("du-athome");
    const s = savedAt(), fits = x => x >= EDGE && x <= innerWidth - SIZE - EDGE;
    let at = { x: s.x, y: s.y };
    for (let i = 0; i < 4; i++) {   // a step can land on the next card's button: step again, a few times at most
      const b = blocker(at.x, at.y); if (!b) break;
      const box = (b.closest(".modal,.card,.du-avoid,section,.foot,.tools,.top,.rail") || b).getBoundingClientRect();
      const opts = [box.left - SIZE - GAP, box.right + GAP].filter(fits).sort((a, c) => Math.abs(a - at.x) - Math.abs(c - at.x));
      const below = box.bottom + GAP;   // no room beside it (a full-width bar): step below it instead
      if (opts.length) at = { x: opts[0], y: at.y };
      else if (below <= innerHeight - SIZE - EDGE) at = { x: at.x, y: below };
      else break;
    }
    apply(at.x, at.y);
  }
  const soon = () => { if (!P.raf) P.raf = requestAnimationFrame(avoid); };
  function dragOn(el, bell) {
    bell.addEventListener("pointerdown", e => {
      if (e.button !== 0) return;
      const r = el.getBoundingClientRect();
      P.drag = { id: e.pointerId, sx: e.clientX, sy: e.clientY, x: r.left, y: r.top, moved: false };
      try { bell.setPointerCapture(e.pointerId); } catch { /* synthetic pointer */ }
    });
    bell.addEventListener("pointermove", e => {
      const d = P.drag; if (!d || d.id !== e.pointerId) return;
      const dx = e.clientX - d.sx, dy = e.clientY - d.sy;
      if (!d.moved && Math.hypot(dx, dy) < 6) return;
      if (!d.moved) { d.moved = true; el.classList.add("du-drag"); if (S.open) toggle(false); }
      apply(clampX(d.x + dx), clampY(d.y + dy));
    });
    const end = e => {
      const d = P.drag; if (!d || d.id !== e.pointerId) return;
      P.drag = null; el.classList.remove("du-drag");
      if (!d.moved) return;
      P.justDragged = true; setTimeout(() => { P.justDragged = false; }, 400);
      const r = el.getBoundingClientRect(), h = homeEl(), hm = h && homeAt(h);
      if (hm && Math.hypot(r.left - hm.x, r.top - hm.y) < SNAP) P.saved = null;
      else {
        const left = r.left + SIZE / 2 < innerWidth / 2, top = r.top + SIZE / 2 < innerHeight / 2;
        P.saved = { h: left ? "l" : "r", dx: Math.round(left ? r.left : innerWidth - SIZE - r.left), v: top ? "t" : "b", dy: Math.round(top ? r.top : innerHeight - SIZE - r.top) };
      }
      savePos(); soon();
    };
    bell.addEventListener("pointerup", end); bell.addEventListener("pointercancel", end);
    /* re-check after every screen change (the add-in re-renders its page), on scroll and on resize */
    new MutationObserver(ms => { if (ms.some(m => !el.contains(m.target))) soon(); }).observe(document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ["class", "hidden", "open"] });
    addEventListener("resize", soon); addEventListener("scroll", soon, true);
  }

  window.DevUpdates = {
    /* app names this add-in's storage key; send(msg) passes seen ids to the other window (pane or full screen), if there is one */
    init({ app, feed, send } = {}) {
      S.key = `${app || "mhs"}_devseen`; if (feed) S.feed = feed; S.send = send || null; S.fresh = new Set();
      load(); loadPos(); S.el = S.el || build(); render();
      return fetch(S.feed, { cache: "no-cache" }).then(r => r.ok ? r.json() : { updates: [] })
        .then(d => { S.items = (d.updates || []).filter(u => u && u.id && Array.isArray(u.bullets)).sort((a, b) => String(b.date).localeCompare(String(a.date))); S.fresh = new Set(unseen().map(u => u.id)); render(); })
        .catch(() => { /* no feed: the bell stays hidden */ });
    },
    setSend(fn) { S.send = fn || null; },
    /* a message from the other window: true when it was a Dev Update message (already merged) */
    receive(m) {
      if (!m || m.k !== "devupd") return false;
      let added = false; (m.seen || []).forEach(id => { if (!S.seen.has(id)) { S.seen.add(id); added = true; } });
      if (added) { save(); render(); }
      return true;
    },
    seen: () => [...S.seen],
    /* the full-screen dialog's URL, carrying what this side has seen so the dot clears on both sides */
    dialogUrl(url) { const u = new URL(url, location.href); if (S.seen.size) u.searchParams.set("devseen", [...S.seen].slice(-MAX_SEEN).join(",")); return u.href; }
  };
})();
