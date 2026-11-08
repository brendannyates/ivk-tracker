/* Dev Update notices, shared by the MHS add-ins. Docks in the toolbar when the page has a .du-slot; otherwise floats top right (draggable). dev-updates.json (newest first) lists what changed:
   { "updates": [ { "id": "2026-10-29-a", "date": "2026-10-29", "bullets": ["..."] } ] }
   Unseen updates from the last 30 days show as one rounded notice under a bell in the top right; the bell keeps the history.
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
    el.className = "devupd";
    el.innerHTML = `<button type="button" class="du-bell" data-tip="Dev updates" aria-haspopup="dialog" aria-expanded="false">${BELL}<span class="du-dot" hidden></span></button>
      <div class="du-card" role="status" aria-live="polite" hidden></div>
      <div class="du-panel" role="dialog" aria-label="Dev updates" hidden></div>`;
    document.body.appendChild(el);
    const bell = el.querySelector(".du-bell");
    bell.onclick = () => { if (P.justDragged) { P.justDragged = false; return; } toggle(!S.open); };
    dragOn(el, bell);
    el.addEventListener("click", e => {
      const b = e.target.closest("[data-du]"); if (!b) return;
      if (b.dataset.du === "dismiss") markSeen(unseen().map(u => u.id));
      if (b.dataset.du === "all") { markSeen(unseen().map(u => u.id)); toggle(true); }
      if (b.dataset.du === "close") toggle(false);
    });
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
    const n = unseen(), bell = S.el.querySelector(".du-bell"), card = S.el.querySelector(".du-card"), panel = S.el.querySelector(".du-panel");
    S.el.hidden = !S.items.length;
    if (!S.el.hidden) soon();
    bell.querySelector(".du-dot").hidden = !n.length;
    bell.setAttribute("aria-label", n.length ? `Dev updates, ${n.length} new` : "Dev updates");
    bell.setAttribute("aria-expanded", String(S.open));
    card.hidden = S.open || !n.length;
    card.innerHTML = n.length ? `<b>Dev Update: ${fmt(n[0].date)}</b>${list(n[0])}
      <div class="du-act">${n.length > 1 ? `<button type="button" data-du="all">See ${n.length - 1} more</button>` : ""}<button type="button" data-du="dismiss">Dismiss</button></div>` : "";
    panel.hidden = !S.open;
    panel.innerHTML = S.open ? `<div class="du-head"><b>Dev updates</b><button type="button" data-du="close" aria-label="Close">Close</button></div>
      ${S.items.map(u => `<section><h3>Dev Update: ${fmt(u.date)}</h3>${list(u)}</section>`).join("")}` : "";
  }

  /* The bell can be dragged anywhere and stays there (per computer, anchored to the nearest corner so it holds its place when
     the window is resized or goes full screen). When a new screen or menu puts a button under it, it steps aside to the left
     or right of that card, whichever is closer, GAP px off, and goes back once the way is clear. A press without moving opens it. */
  const SIZE = 44, EDGE = 4, GAP = 6, DEF = { r: 16, t: 14 };
  const P = { saved: null, nudge: null, drag: null, justDragged: false, raf: 0 };
  const posKey = () => `${S.key}_bell`;
  function loadPos() { try { const p = JSON.parse(localStorage.getItem(posKey()) || "null"); if (p && isFinite(p.dx) && isFinite(p.dy)) P.saved = p; } catch { /* none */ } }
  function savePos() { try { localStorage.setItem(posKey(), JSON.stringify(P.saved)); } catch { /* this session only */ } }
  const clampX = x => Math.min(Math.max(EDGE, x), innerWidth - SIZE - EDGE), clampY = y => Math.min(Math.max(EDGE, y), innerHeight - SIZE - EDGE);
  /* where the bell lives before any stepping aside: the saved spot, else the default top right */
  function home() {
    const p = P.saved;
    if (!p) return { x: innerWidth - SIZE - DEF.r, y: DEF.t };
    return { x: clampX(p.h === "l" ? p.dx : innerWidth - SIZE - p.dx), y: clampY(p.v === "t" ? p.dy : innerHeight - SIZE - p.dy) };
  }
  function apply(x, y) {
    const el = S.el; if (!el) return;
    el.style.left = `${x}px`; el.style.top = `${y}px`; el.style.right = "auto";
    el.classList.toggle("du-left", x + SIZE / 2 < innerWidth / 2);   // notice and history open towards the middle
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
    /* docked: the add-in's toolbar has a .du-slot (right of light/dark); the bell sits there and doesn't float or drag */
    const slot = document.querySelector(".du-slot");
    if (slot) {
      if (S.el.parentNode !== slot) slot.appendChild(S.el);
      S.el.classList.add("du-docked"); S.el.classList.remove("du-left", "du-up");
      S.el.style.left = S.el.style.top = S.el.style.right = "";
      return;
    }
    if (S.el.classList.contains("du-docked")) { S.el.classList.remove("du-docked"); document.body.appendChild(S.el); }
    if (S.el.hidden) return;
    const h = home(), fits = x => x >= EDGE && x <= innerWidth - SIZE - EDGE;
    let at = { x: h.x, y: h.y };
    for (let i = 0; i < 4; i++) {   // a step can land on the next card's button: step again, a few times at most
      const b = blocker(at.x, at.y); if (!b) break;
      const box = (b.closest(".modal,.card,.du-avoid,section,.foot,.tools,.top,.rail") || b).getBoundingClientRect();
      const opts = [box.left - SIZE - GAP, box.right + GAP].filter(fits).sort((a, c) => Math.abs(a - at.x) - Math.abs(c - at.x));
      /* no room beside it (a full-width bar): step below it instead */
      const below = box.bottom + GAP;
      if (opts.length) at = { x: opts[0], y: at.y };
      else if (below <= innerHeight - SIZE - EDGE) at = { x: at.x, y: below };
      else break;
    }
    P.nudge = at.x === h.x && at.y === h.y ? null : at;
    apply(at.x, at.y);
  }
  const soon = () => { if (!P.raf) P.raf = requestAnimationFrame(avoid); };
  function dragOn(el, bell) {
    bell.addEventListener("pointerdown", e => {
      if (e.button !== 0 || el.classList.contains("du-docked")) return;
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
      const r = el.getBoundingClientRect(), left = r.left + SIZE / 2 < innerWidth / 2, top = r.top + SIZE / 2 < innerHeight / 2;
      P.saved = { h: left ? "l" : "r", dx: Math.round(left ? r.left : innerWidth - SIZE - r.left), v: top ? "t" : "b", dy: Math.round(top ? r.top : innerHeight - SIZE - r.top) };
      P.nudge = null; savePos(); soon();
    };
    bell.addEventListener("pointerup", end); bell.addEventListener("pointercancel", end);
    /* re-check after every screen change (the add-in re-renders its page), and on resize */
    new MutationObserver(ms => { if (ms.some(m => !el.contains(m.target))) soon(); }).observe(document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ["class", "hidden", "open"] });
    addEventListener("resize", soon);
  }

  window.DevUpdates = {
    /* app names this add-in's storage key; send(msg) passes seen ids to the other window (pane or full screen), if there is one */
    init({ app, feed, send } = {}) {
      S.key = `${app || "mhs"}_devseen`; if (feed) S.feed = feed; S.send = send || null;
      load(); loadPos(); S.el = S.el || build(); render(); soon();
      return fetch(S.feed, { cache: "no-cache" }).then(r => r.ok ? r.json() : { updates: [] })
        .then(d => { S.items = (d.updates || []).filter(u => u && u.id && Array.isArray(u.bullets)).sort((a, b) => String(b.date).localeCompare(String(a.date))); render(); })
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
    /* the full-screen dialog's URL, carrying what this side has seen so its notice never shows twice */
    dialogUrl(url) { const u = new URL(url, location.href); if (S.seen.size) u.searchParams.set("devseen", [...S.seen].slice(-MAX_SEEN).join(",")); return u.href; }
  };
})();
