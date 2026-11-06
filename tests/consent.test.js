// The copay check at check-in: LumaPay card on file + verbal consent, saved as tblIvkPay consent rows (invented MRNs and codes)
const test = require("node:test"), assert = require("node:assert/strict");
const C = require("../core.js");

const T = C.serialOf(new Date(2026, 10, 5));
const payH = C.PAY_COLS;
const pay = o => payH.map(h => ({ Id: o.id, Date: o.date ?? C.isoOf(T), MRN: o.mrn, Patient: "TESANN", Key: o.key ?? "", Kind: o.kind, Amount: String(o.amt ?? 0), Method: o.method ?? "", By: o.by ?? "Sam", At: o.at ?? "" }[h] ?? ""));

test("consent method codes round-trip", () => {
  assert.equal(C.consentMethod(true, true), "card+consent");
  assert.equal(C.consentMethod(true, false), "card");
  assert.equal(C.consentMethod(false, true), "consent");
  assert.equal(C.consentMethod(false, false), "none");
  assert.deepEqual(C.CONSENT_METHODS, ["card+consent", "card", "consent", "none"]);
});

test("consentOf: the visit's latest consent row counts; none means nobody did the check", () => {
  const k = C.visitKey(T, "990301", 600), other = C.visitKey(T, "990302", 630);
  const P = C.payRows([payH, pay({ id: "a", mrn: "990301", key: k, kind: "consent", method: "card", by: "Sam", at: "2026-11-05T17:01:00Z" }),
    pay({ id: "b", mrn: "990301", key: k, kind: "consent", method: "card+consent", by: "Ana", at: "2026-11-05T17:03:00Z" }),
    pay({ id: "c", mrn: "990302", key: other, kind: "consent", method: "none" })]);
  const c = C.consentOf(P, k);
  assert.equal(c.done, true); assert.equal(c.card, true); assert.equal(c.consent, true); assert.equal(c.by, "Ana");
  const n = C.consentOf(P, other);
  assert.equal(n.done, false); assert.equal(n.card, false); assert.equal(n.consent, false);
  assert.equal(C.consentOf(P, C.visitKey(T, "990303", 700)), null);
  assert.equal(C.consentOf(null, k), null);
});

test("consent rows never touch balances, collected totals or the visit's payment answer", () => {
  const k = C.visitKey(T, "990301", 600);
  const P = C.payRows([payH, pay({ id: "a", mrn: "990301", key: C.partKey(k, "prov"), kind: "collected", amt: 40 }),
    pay({ id: "b", mrn: "990301", key: k, kind: "consent", method: "card+consent" })]);
  const L = C.ledger("990301", P);
  assert.equal(L.balance, 0); assert.equal(L.collected, 40);
  assert.equal(C.payOf(P, k, "prov").kind, "collected");
  assert.equal(C.payOf(P, k), null, "a consent row isn't a whole-visit payment answer");
});
