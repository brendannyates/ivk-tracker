/* Dev Update bell, shared by the MHS add-ins (keep this file byte-identical across them). dev-updates.json (newest first):
   { "updates": [ { "id": "2026-10-29-a", "date": "2026-10-29", "bullets": ["..."] } ] }
   The bell only shows where the page has a .du-home: each add-in puts one right of its top-left name in full screen and none in
   the task pane. It sits inside that spot and stays put (no dragging); the history opens below it, kept inside the window. A red
   dot means an update from the last 30 days that this computer hasn't opened; nothing pops up on its own, a click opens the
   history and clears the dot.
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
    bell.onclick = () => toggle(!S.open);
    watch(el);
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

  /* Home is the page's .du-home. The bell is moved inside it whenever the add-in re-renders (a new .du-home replaces the old
     one), and hidden when there is none. The history is fixed to the window under the bell, so a rail or header that clips its
     overflow can't cut it off. */
  const P = { raf: 0 };
  const homeEl = () => document.querySelector(".du-home");
  function place() {
    const el = S.el, panel = el.querySelector(".du-panel"); if (panel.hidden) return;
    const r = el.querySelector(".du-bell").getBoundingClientRect(), w = Math.min(360, innerWidth - 24);
    const x = Math.max(12, Math.min(r.left, innerWidth - w - 12)), y = r.bottom + 8;
    panel.style.left = `${x}px`; panel.style.top = `${y}px`;
    /* an ancestor with a transform, filter or backdrop-filter makes "fixed" relative to that ancestor, not the window: measure where
       it landed and shift it back by the difference */
    const p = panel.getBoundingClientRect(), dx = x - p.left, dy = y - p.top;
    if (Math.abs(dx) > 0.5 || Math.abs(dy) > 0.5) { panel.style.left = `${x + dx}px`; panel.style.top = `${y + dy}px`; }
  }
  function avoid() {
    P.raf = 0; if (!S.el) return;
    const h = homeEl();
    S.el.hidden = !h || !S.items.length;
    if (S.el.hidden) { if (S.open) { S.open = false; render(); } if (!h && S.el.parentNode !== document.body) document.body.appendChild(S.el); return; }
    if (S.el.parentNode !== h) h.appendChild(S.el);
    place();
  }
  /* a microtask, not a frame: the bell is back in the new .du-home before the page paints, so a re-render never flickers it */
  const soon = () => { if (!P.raf) { P.raf = 1; queueMicrotask(avoid); } };
  /* re-check after every screen change (the add-in re-renders its page), on scroll and on resize */
  function watch(el) {
    new MutationObserver(ms => { if (ms.some(m => !el.contains(m.target))) soon(); }).observe(document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ["class", "hidden", "open"] });
    addEventListener("resize", soon); addEventListener("scroll", soon, true);
  }

  window.DevUpdates = {
    /* app names this add-in's storage key; send(msg) passes seen ids to the other window (pane or full screen), if there is one */
    init({ app, feed, send } = {}) {
      S.key = `${app || "mhs"}_devseen`; if (feed) S.feed = feed; S.send = send || null; S.fresh = new Set();
      load(); S.el = S.el || build(); render();
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
