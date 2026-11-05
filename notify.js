/* notify.js: Teams posts for the IVK Tracker, built on the ESK Tracker's notify.js (same workflow template, card and post).
   Each chat has its own Teams workflow made from the "Send webhook alerts to a chat" template (who can trigger: Anyone).
   Its URL sits in tblIvkRoutes, one row per event:
     arrival   -> the RN's chat: a patient is here (the RN takes vitals and doesn't use the tracker)
     protocol  -> the Castro IVK group chat: confirm the protocol, send to the desk at discharge for scheduling
   The add-in decides when to send and what to say; the workflow posts the card it is given (and, for a card with buttons, waits
   for the answer and writes it to tblIvkReplies under the payload's ref). Nothing here touches the
   workbook: the caller passes the routes it read and a log function. Workflow URLs are secrets (the sig= part). */
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

  /* What the Teams workflow template expects: { type: 'message', attachments: [ Adaptive Card ] }.
     o.mentions = [{ name, id: email }]: each is tagged at the start of the text (<at>name</at> plus a msteams mention entity).
     o.reply = a card that takes an answer in Teams (the workflow posts it with "wait for a response"):
       { kind: "chair", chair, chairs: ["1", ...] }        a chair picker and "Update chair"
       { kind: "protocol", freq, freqs: [...] }            "Confirm <freq>", or pick another from the list and "Confirm this one"
     Submit data carries act; the inputs add chair / pick. No patient details go in the data: the workflow sends back the payload's ref. */
  const esc = s => String(s).replace(/[<>]/g, "");
  function card(text, facts, o) {
    const men = ((o && o.mentions) || []).filter(m => m && m.name && m.id);
    const lead = men.map(m => `<at>${esc(m.name)}</at>`).join(" ");
    const body = [{ type: "TextBlock", text: lead ? `${lead} ${text}` : text, wrap: true }];
    const f = (facts || []).filter(x => x[1]).map(x => ({ title: x[0], value: String(x[1]) }));
    if (f.length) body.push({ type: "FactSet", facts: f });
    const content = { $schema: "http://adaptivecards.io/schemas/adaptive-card.json", type: "AdaptiveCard", version: "1.4", body };
    const r = o && o.reply;
    if (r && r.kind === "chair" && r.chairs && r.chairs.length) {
      body.push({ type: "Input.ChoiceSet", id: "chair", label: "Chair", style: "compact", value: String(r.chair || ""), choices: r.chairs.map(c => ({ title: `Chair ${c}`, value: String(c) })) });
      content.actions = [{ type: "Action.Submit", title: "Update chair", data: { act: "chair" } }];
    } else if (r && r.kind === "protocol") {
      const list = (r.freqs || []).filter(Boolean);
      content.actions = [];
      if (r.freq) content.actions.push({ type: "Action.Submit", title: `Confirm ${r.freq}`, data: { act: "confirm", freq: r.freq } });
      if (list.length) {
        body.push({ type: "Input.ChoiceSet", id: "pick", label: r.freq ? "Or a different protocol" : "Protocol", style: "compact", value: list.includes(r.freq) ? r.freq : list[0], choices: list.map(x => ({ title: x, value: x })) });
        content.actions.push({ type: "Action.Submit", title: r.freq ? "Confirm this one" : "Confirm", data: { act: "pick" } });
      }
      if (!content.actions.length) delete content.actions;
    }
    if (men.length) content.msteams = { entities: men.map(m => ({ type: "mention", text: `<at>${esc(m.name)}</at>`, mentioned: { id: m.id, name: m.name } })) };
    return { type: "message", attachments: [{ contentType: "application/vnd.microsoft.card.adaptive", contentUrl: null, content }] };
  }
  /* a card only waits for an answer when it has something to press */
  const takesReply = p => !!(p && p.attachments && p.attachments[0] && p.attachments[0].content.actions && p.attachments[0].content.actions.length);

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

  /* routes: { arrival: { to, url, on }, protocol: {...} } from tblIvkRoutes; enabled: tblIvkConfig Enabled.
     d = { event, text, facts: [[title, value]], patient, chair, sentBy, clinic, mentions, reply }; log(row) appends to tblIvkLog.
     A card that takes an answer gets a ref (the payload's ref; the workflow waits for the answer and writes it to tblIvkReplies
     under that ref). beforePost(ref) writes the tblIvkReplies row first, so the answer always has a row to land in; if it throws,
     the card goes without buttons. */
  async function send(d, routes, enabled, log, fetchFn, beforePost) {
    if (!truthy(enabled)) return { ok: false, reason: "Teams posts are off for this workbook (Settings › Teams)." };
    const r = routes && routes[d.event];
    if (!r || !truthy(r.on)) return { ok: false, reason: `${EVENTS[d.event] || d.event} posts are turned off (Settings › Teams).` };
    const url = flowUrl(r.url);
    if (!url) return { ok: false, reason: `No Teams workflow for ${EVENTS[d.event] || d.event} yet. Add it in Settings › Teams.` };
    let c = card(d.text, d.facts, { mentions: d.mentions, reply: d.reply }), ref = "";
    if (takesReply(c)) {
      ref = newId();
      try { if (beforePost) await beforePost(ref); else throw new Error("no reply table"); }
      catch (e) { console.warn("Teams reply row", e); ref = ""; c = card(d.text, d.facts, { mentions: d.mentions }); }
    }
    const now = new Date(), payload = Object.assign(c, { event: d.event, ref, message: d.text.replace(/\*\*/g, ""),
      patient: d.patient || "", chair: String(d.chair || ""), clinic: d.clinic || "", sentBy: d.sentBy || "", sentAt: now.toISOString(), source: "IVK Tracker" });
    const res = await post(url, payload, fetchFn);
    const result = (res.ok ? `accepted ${res.status}` : `failed: ${res.reason}`) + (ref ? ` (answer ${ref})` : "");
    try { if (log) await log([ref || newId(), d.event, d.patient || "", String(d.chair || ""), payload.message, d.sentBy || "", payload.sentAt, result]); }
    catch (e) { console.warn("Teams log", e); }
    return Object.assign({ message: payload.message, to: r.to, ref }, res);
  }
  /* Settings › Teams › Send test: one card with no patient, nothing logged */
  async function test(event, routes, sentBy, fetchFn) {
    const r = routes && routes[event], url = flowUrl(r && r.url);
    if (!url) return { ok: false, reason: "Paste the workflow URL first." };
    const now = new Date(), msg = `Test from IVK Tracker${sentBy ? ` (${sentBy})` : ""}: ${EVENTS[event] || event} posts will come here.`;
    return post(url, Object.assign(card(msg), { event: "test", message: msg, patient: "", sentBy: sentBy || "", sentAt: now.toISOString(), source: "IVK Tracker" }), fetchFn);
  }

  return { EVENTS, flowUrl, truthy, card, takesReply, post, send, test, HTTP_HINT };
});
