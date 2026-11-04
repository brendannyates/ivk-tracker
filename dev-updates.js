/* Dev Update notices, shared by the MHS add-ins. dev-updates.json (newest first) lists what changed:
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
    el.innerHTML = `<button type="button" class="du-bell" aria-haspopup="dialog" aria-expanded="false">${BELL}<span class="du-dot" hidden></span></button>
      <div class="du-card" role="status" aria-live="polite" hidden></div>
      <div class="du-panel" role="dialog" aria-label="Dev updates" hidden></div>`;
    document.body.appendChild(el);
    el.querySelector(".du-bell").onclick = () => toggle(!S.open);
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

  window.DevUpdates = {
    /* app names this add-in's storage key; send(msg) passes seen ids to the other window (pane or full screen), if there is one */
    init({ app, feed, send } = {}) {
      S.key = `${app || "mhs"}_devseen`; if (feed) S.feed = feed; S.send = send || null;
      load(); S.el = S.el || build(); render();
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
