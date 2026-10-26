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
  const r = await N.send({ event: "arrival", text: "TESANN is here for their 13:30 IVK visit with Dr. Rivera. Chair 2.", facts: [["Chair", "Chair 2"], ["Empty", ""]], patient: "TESANN", chair: "2", sentBy: "Sam" },
    routes, "TRUE", row => logged.push(row), f);
  assert.equal(r.ok, true); assert.equal(r.status, 202); assert.equal(r.to, "RN");
  assert.equal(f.calls[0].url, URL1);
  const body = JSON.parse(f.calls[0].o.body), card = body.attachments[0].content;
  assert.equal(body.type, "message"); assert.equal(card.type, "AdaptiveCard"); assert.equal(card.body[0].text.startsWith("TESANN is here"), true);
  assert.deepEqual(card.body[1].facts, [{ title: "Chair", value: "Chair 2" }], "blank facts are dropped");
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
