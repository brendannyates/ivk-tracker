/* notify.js: Teams posts for the IVK Tracker, in the ESK Tracker's card layout.
   Each chat has its own "IVK Tracker - post to <chat>" flow, built from blank in Power Automate (a flow made from the Teams
   "Send webhook alerts to a chat" template stamps "<name> used a Workflow template" under every card). Trigger: When a Teams
   webhook request is received, who can trigger: Anyone. Its URL sits in tblIvkRoutes, one row per event:
     arrival   -> the RN's chat: a patient is here (the RN takes vitals and doesn't use the tracker)
     protocol  -> the Castro IVK group chat: confirm the protocol, send to the desk at discharge for scheduling
   When both routes are the same chat, check-in sends one card that carries both (app.js sendCheckin).
   Cards are plain (no buttons). A card for a visit also carries a one-line collapsed version and how many minutes to wait: the flow
   posts the card, waits, and swaps it for the collapsed line, COLLAPSE_MIN after check-in. The add-in decides what to say and when
   to collapse; the flow only posts, waits and updates. Nothing here touches the workbook: the caller passes the routes it read and
   a log function. Flow URLs are secrets (the sig= part). */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.IvkNotify = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  const EVENTS = { arrival: "Arrivals", protocol: "Protocol and scheduling" };
  /* a Power Automate trigger URL (new powerplatform.com form or the older logic.azure.com one), else "" */
  function flowUrl(v) {
    const s = String(v == null ? "" : v).trim();
    return /^https:\/\/[^\/\s?#]+\.(?:powerplatform\.com|logic\.azure\.com)(?::\d+)?\/\S+$/i.test(s) ? s : "";
  }
  const truthy = v => v === true || /^(true|yes|y|1)$/i.test(String(v == null ? "" : v).trim());

  /* minutes after check-in that a visit's card collapses to one line in Teams */
  const COLLAPSE_MIN = 120;

  /* What the flow expects: { type: 'message', attachments: [ Adaptive Card ] }. The layout is the ESK Tracker's: the message is the
     bold first line, then each further line under a separator. No FactSet, no footer, no buttons.
     o.mentions = [{ name, id: email }]: tagged at the start of the first line (<at>name</at> plus a msteams mention entity).
     o.lines = [{ text, mentions }]: more lines, each with its own tags (the combined check-in card: the RN on the first line, the
     provider on the ask). facts = [[title, value]]: one "Title: value" line each; blank values are left out. */
  const esc = s => String(s).replace(/[<>]/g, "");
  const tagged = (text, men) => { const lead = men.map(m => `<at>${esc(m.name)}</at>`).join(" "); return lead ? `${lead} ${text}` : text; };
  const okMen = list => (list || []).filter(m => m && m.name && m.id);
  const line = t => ({ type: "TextBlock", text: t, wrap: true, spacing: "Medium", separator: true });
  function card(text, facts, o) {
    const men = okMen(o && o.mentions), all = men.slice();
    const body = [{ type: "TextBlock", text: tagged(text, men), wrap: true, weight: "Bolder" }];
    for (const l of (o && o.lines) || []) {
      if (!l || !l.text) continue;
      const lm = okMen(l.mentions); all.push(...lm); body.push(line(tagged(l.text, lm)));
    }
    for (const x of (facts || []).filter(x => x && x[1])) body.push(line(x[0] ? `${x[0]}: ${x[1]}` : String(x[1])));
    const content = { $schema: "http://adaptivecards.io/schemas/adaptive-card.json", type: "AdaptiveCard", version: "1.4", body };
    const seen = new Set(), ents = all.filter(m => !seen.has(m.id) && seen.add(m.id));
    if (ents.length) content.msteams = { entities: ents.map(m => ({ type: "mention", text: `<at>${esc(m.name)}</at>`, mentioned: { id: m.id, name: m.name } })) };
    return { type: "message", attachments: [{ contentType: "application/vnd.microsoft.card.adaptive", contentUrl: null, content }] };
  }
  /* the collapsed card: one muted line, no tags */
  const collapsedCard = text => ({ $schema: "http://adaptivecards.io/schemas/adaptive-card.json", type: "AdaptiveCard", version: "1.4",
    body: [{ type: "TextBlock", text: String(text), wrap: true, isSubtle: true }] });
  /* minutes from now until the card should collapse: COLLAPSE_MIN after check-in (arr, minutes since midnight), at least 1 */
  const collapseIn = (arr, now) => arr == null || now == null ? COLLAPSE_MIN : Math.max(1, COLLAPSE_MIN - (now - arr));

  /* Real JSON with CORS: the trigger answers the preflight, so the add-in reads its answer (202 = accepted). Only a 2xx counts as
     sent. There is no no-cors retry: it could post a second card when the first got through, and its text body isn't read by the
     template, so a "sent" from it can't be trusted. */
  const HTTP_HINT = { 401: "the URL is incomplete or out of date (copy it again from the workflow), or the workflow isn’t set to Anyone",
    403: "the workflow isn’t set to Anyone, or the tenant blocks it", 404: "the URL is wrong or the workflow was deleted",
    409: "the workflow is turned off", 429: "too many posts at once, try again in a minute" };
  async function post(url, payload, fetchFn) {
    const F = fetchFn || fetch; let r;
    try { r = await F(url, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) }); }
    catch (e) { return { ok: false, reason: "Teams couldn’t be reached, or the browser blocked the answer. Check the chat before sending again." }; }
    if (r.ok) return { ok: true, status: r.status };
    let t = ""; try { t = await r.text(); } catch (e) { /* no body */ }
    const code = (t.match(/"code"\s*:\s*"([^"]+)"/) || [])[1], hint = HTTP_HINT[r.status] || (/disabled/i.test(t) ? "the workflow is turned off" : "");
    return { ok: false, status: r.status, reason: `The Teams workflow answered ${r.status}${code ? ` (${code})` : ""}${hint ? `: ${hint}` : ""}.` };
  }
  const newId = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 6);

  /* The eligibility feed: a Power Automate flow (When an HTTP request is received ... Response) that answers a POST of
     { tracker: "IVK" } with the text of Eligibility History.csv (raw, or JSON wrapping it: core.feedText unwraps it). Sent as
     text/plain so it's a simple CORS request with no preflight; the flow's Response action sets Access-Control-Allow-Origin.
     Same URL check as the Teams posts. Out: { ok, text } or { ok: false, reason } in plain words. */
  async function pullFeed(url, body, fetchFn) {
    const u = flowUrl(url); if (!u) return { ok: false, reason: "No eligibility feed URL yet. Add it in Settings › Eligibility feed." };
    const F = fetchFn || fetch; let r;
    try { r = await F(u, { method: "POST", headers: { "Content-Type": "text/plain;charset=UTF-8" }, body: JSON.stringify(body) }); }
    catch (e) { return { ok: false, reason: "The eligibility feed couldn’t be reached. Nothing changed; check the connection and try again." }; }
    let t = ""; try { t = await r.text(); } catch (e) { /* no body */ }
    if (r.ok) return { ok: true, status: r.status, text: t };
    const hint = HTTP_HINT[r.status] || (/disabled/i.test(t) ? "the workflow is turned off" : "");
    return { ok: false, status: r.status, reason: `The eligibility feed answered ${r.status}${hint ? `: ${hint}` : ""}. Nothing changed.` };
  }

  /* routes: { arrival: { to, url, on }, protocol: {...} } from tblIvkRoutes; enabled: tblIvkConfig Enabled.
     d = { event, events, text, facts, lines, patient, chair, sentBy, clinic, mentions, collapse }; log(row) appends to tblIvkLog.
     d.event picks the route; d.events (default [d.event]) are the events logged, one row each, so a combined check-in card counts as
     both the arrival and the ask. d.collapse = { text, min }: the flow swaps the card for that line after min minutes. */
  async function send(d, routes, enabled, log, fetchFn) {
    if (!truthy(enabled)) return { ok: false, reason: "Teams posts are off for this workbook (Settings › Teams)." };
    const r = routes && routes[d.event];
    if (!r || !truthy(r.on)) return { ok: false, reason: `${EVENTS[d.event] || d.event} posts are turned off (Settings › Teams).` };
    const url = flowUrl(r.url);
    if (!url) return { ok: false, reason: `No Teams workflow for ${EVENTS[d.event] || d.event} yet. Add it in Settings › Teams.` };
    const c = card(d.text, d.facts, { mentions: d.mentions, lines: d.lines });
    const msg = [d.text].concat((d.lines || []).filter(l => l && l.text).map(l => l.text)).join(" ").replace(/\*\*/g, "");
    const now = new Date(), payload = Object.assign(c, { event: d.event, message: msg,
      patient: d.patient || "", chair: String(d.chair || ""), clinic: d.clinic || "", sentBy: d.sentBy || "", sentAt: now.toISOString(), source: "IVK Tracker" });
    if (d.collapse && d.collapse.text) Object.assign(payload, { collapsed: collapsedCard(d.collapse.text), collapseMin: Math.max(1, Math.round(+d.collapse.min || COLLAPSE_MIN)) });
    const res = await post(url, payload, fetchFn);
    const result = res.ok ? `accepted ${res.status}` : `failed: ${res.reason}`;
    for (const ev of d.events && d.events.length ? d.events : [d.event]) {
      try { if (log) await log([newId(), ev, d.patient || "", String(d.chair || ""), payload.message, d.sentBy || "", payload.sentAt, result]); }
      catch (e) { console.warn("Teams log", e); }
    }
    return Object.assign({ message: payload.message, to: r.to }, res);
  }
  /* both events go to the same chat (same flow URL, both on): check-in sends one card */
  const sameChat = routes => { const a = routes && routes.arrival, p = routes && routes.protocol;
    return !!(a && p && truthy(a.on) && truthy(p.on) && flowUrl(a.url) && flowUrl(a.url) === flowUrl(p.url)); };
  /* Settings › Teams › Send test: one card with no patient, nothing logged */
  async function test(event, routes, sentBy, fetchFn) {
    const r = routes && routes[event], url = flowUrl(r && r.url);
    if (!url) return { ok: false, reason: "Paste the workflow URL first." };
    const now = new Date(), msg = `Test from IVK Tracker${sentBy ? ` (${sentBy})` : ""}: ${EVENTS[event] || event} posts will come here.`;
    return post(url, Object.assign(card(msg), { event: "test", message: msg, patient: "", sentBy: sentBy || "", sentAt: now.toISOString(), source: "IVK Tracker" }), fetchFn);
  }

  return { EVENTS, COLLAPSE_MIN, flowUrl, truthy, card, collapsedCard, collapseIn, sameChat, post, send, test, pullFeed, HTTP_HINT };
});
