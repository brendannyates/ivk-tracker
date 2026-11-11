// node --test tests/   (no real workflow URLs: a fake fetch answers)
const test = require("node:test"), assert = require("node:assert/strict");
const N = require("../notify.js");

const URL1 = "https://prod-00.westus.logic.azure.com:443/workflows/abc/triggers/manual/paths/invoke?api-version=2016-06-01&sig=FAKE";
const URL2 = "https://default0000.00.environment.api.powerplatform.com/powerautomate/automations/direct/workflows/abc/triggers/manual/paths/invoke?sig=FAKE";
const routes = { arrival: { to: "RN", url: URL1, on: true }, protocol: { to: "Castro IVK group chat", url: URL2, on: true } };
const fake = (status, body = "") => { const calls = []; const f = async (url, o) => { calls.push({ url, o }); return { ok: status < 300, status, text: async () => body }; }; f.calls = calls; return f; };

test("flowUrl accepts workflow trigger URLs only", () => {
  assert.equal(N.flowUrl(URL1), URL1); assert.equal(N.flowUrl(" " + URL2 + " "), URL2);
  assert.equal(N.flowUrl("https://teams.microsoft.com/l/chat/19:abc"), "");
  assert.equal(N.flowUrl("http://prod-00.westus.logic.azure.com/x"), "");
});

test("send posts the Adaptive Card the Teams template expects, and logs it", async () => {
  const f = fake(202), logged = [];
  const r = await N.send({ event: "arrival", text: "TESANN is here for their 1:30 PM IVK visit with Dr. Rivera. Chair 2.", facts: [["Chair", "Chair 2"], ["Empty", ""]], patient: "TESANN", chair: "2", sentBy: "Sam" },
    routes, "TRUE", row => logged.push(row), f);
  assert.equal(r.ok, true); assert.equal(r.status, 202); assert.equal(r.to, "RN");
  assert.equal(f.calls[0].url, URL1);
  const body = JSON.parse(f.calls[0].o.body), card = body.attachments[0].content;
  assert.equal(body.type, "message"); assert.equal(card.type, "AdaptiveCard"); assert.equal(card.body[0].text.startsWith("TESANN is here"), true);
  assert.equal(card.body[0].weight, "Bolder", "the message is the bold first line");
  assert.equal(card.body.length, 2, "blank facts are dropped");
  assert.deepEqual(card.body[1], { type: "TextBlock", text: "Chair: Chair 2", wrap: true, spacing: "Medium", separator: true }, "a fact is its own line, no FactSet");
  assert.equal(card.actions, undefined, "no buttons"); assert.equal(body.collapsed, undefined, "no collapse asked for, none sent");
  assert.equal(body.patient, "TESANN"); assert.equal(body.source, "IVK Tracker");
  assert.equal(logged.length, 1); assert.deepEqual(logged[0].slice(1, 4), ["arrival", "TESANN", "2"]); assert.equal(logged[0][7], "accepted 202");
});

test("send refuses when off, unrouted or the URL is wrong; never posts then", async () => {
  const f = fake(202);
  assert.match((await N.send({ event: "arrival", text: "x" }, routes, "FALSE", null, f)).reason, /off for this workbook/);
  assert.match((await N.send({ event: "protocol", text: "x" }, { protocol: { url: URL2, on: false } }, "TRUE", null, f)).reason, /turned off/);
  assert.match((await N.send({ event: "protocol", text: "x" }, { protocol: { url: "https://example.com/x", on: true } }, "TRUE", null, f)).reason, /No Teams workflow/);
  assert.equal(f.calls.length, 0);
});

test("an error answer is explained and logged as failed", async () => {
  const f = fake(404, '{"error":{"code":"WorkflowNotFound"}}'), logged = [];
  const r = await N.send({ event: "protocol", text: "Can we please confirm every 2 week protocol for TESANN and send to the desk at discharge for scheduling?", patient: "TESANN" }, routes, true, row => logged.push(row), f);
  assert.equal(r.ok, false); assert.match(r.reason, /404 \(WorkflowNotFound\): the URL is wrong/);
  assert.match(logged[0][7], /^failed: /);
});

test("a blocked or failed request is not sent, and is not retried", async () => {
  const calls = [], logged = []; const f = async (url, o) => { calls.push(o); throw new TypeError("Failed to fetch"); };
  const r = await N.send({ event: "arrival", text: "x", patient: "TESANN" }, routes, "TRUE", row => logged.push(row), f);
  assert.equal(r.ok, false); assert.equal(calls.length, 1); assert.match(r.reason, /Check the chat before sending again/); assert.match(logged[0][7], /^failed/);
});

test("test card has no patient", async () => {
  const f = fake(202), r = await N.test("protocol", routes, "Sam", f), b = JSON.parse(f.calls[0].o.body);
  assert.equal(r.ok, true); assert.equal(b.patient, ""); assert.match(b.message, /^Test from IVK Tracker \(Sam\)/);
});

test("mentions tag at the start of the text, with a msteams entity per person", () => {
  const p = N.card("TESANN is here.", [], { mentions: [{ name: "Sam Nurse", id: "snurse@example.org" }, { name: "", id: "x@example.org" }] }), c = p.attachments[0].content;
  assert.equal(c.body[0].text, "<at>Sam Nurse</at> TESANN is here.");
  assert.deepEqual(c.msteams.entities, [{ type: "mention", text: "<at>Sam Nurse</at>", mentioned: { id: "snurse@example.org", name: "Sam Nurse" } }]);
  assert.equal(N.card("x").attachments[0].content.msteams, undefined);
});

test("combined check-in card: the RN on the arrival line, the provider on the ask line, protocol lines under them", () => {
  const rn = { name: "Sam Nurse", id: "snurse@example.org" }, md = { name: "Ana Rivera", id: "arivera@example.org" };
  const c = N.card("TESANN is here for their 9:30 AM IVK visit with Dr. Rivera.", [["Protocol on file", "every 2 weeks"], ["", ""], ["Why", ""]],
    { mentions: [rn], lines: [{ text: "Can we please send TESANN to the desk at discharge for scheduling?", mentions: [md] }] }).attachments[0].content;
  assert.deepEqual(c.body.map(b => b.text), ["<at>Sam Nurse</at> TESANN is here for their 9:30 AM IVK visit with Dr. Rivera.",
    "<at>Ana Rivera</at> Can we please send TESANN to the desk at discharge for scheduling?", "Protocol on file: every 2 weeks"]);
  assert.deepEqual(c.msteams.entities.map(e => e.mentioned.id), ["snurse@example.org", "arivera@example.org"]);
  const twice = N.card("x", [], { mentions: [rn], lines: [{ text: "y", mentions: [rn] }] }).attachments[0].content;
  assert.equal(twice.msteams.entities.length, 1, "one entity per person");
});

test("a combined card is one post, logged as both events, and carries its collapsed line", async () => {
  const f = fake(202), logged = [];
  const r = await N.send({ event: "arrival", events: ["arrival", "protocol"], text: "TESANN is here.", lines: [{ text: "Can we please send TESANN to the desk?" }], patient: "TESANN",
    collapse: { text: "TESANN · 9:30 AM IVK with Dr. Rivera · checked in 9:24 AM", min: 120 } }, routes, "TRUE", row => logged.push(row), f);
  const b = JSON.parse(f.calls[0].o.body);
  assert.equal(r.ok, true); assert.equal(f.calls.length, 1);
  assert.deepEqual(logged.map(l => l[1]), ["arrival", "protocol"]); assert.equal(logged[0][4], "TESANN is here. Can we please send TESANN to the desk?");
  assert.equal(b.collapseMin, 120); assert.deepEqual(b.collapsed.body, [{ type: "TextBlock", text: "TESANN · 9:30 AM IVK with Dr. Rivera · checked in 9:24 AM", wrap: true, isSubtle: true }]);
});

test("collapse timing counts from check-in; same chat means same URL with both routes on", () => {
  assert.equal(N.COLLAPSE_MIN, 120);
  assert.equal(N.collapseIn(540, 540), 120); assert.equal(N.collapseIn(540, 600), 60); assert.equal(N.collapseIn(540, 800), 1, "never less than a minute");
  assert.equal(N.collapseIn(null, 600), 120);
  assert.equal(N.sameChat({ arrival: { url: URL2, on: true }, protocol: { url: " " + URL2, on: "TRUE" } }), true);
  assert.equal(N.sameChat(routes), false); assert.equal(N.sameChat({ arrival: { url: URL2, on: true }, protocol: { url: URL2, on: false } }), false);
});
