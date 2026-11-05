// node --test tests/   (invented names only): who's at the desk, the clinic from the workbook name, admin, signatures, no built-in clinic
const test = require("node:test"), assert = require("node:assert/strict");
const C = require("../core.js");
const TPL = require("../excel.js");

test("nothing clinic-specific is built in", () => {
  for (const k of ["Clinic", "ClinicAddress", "ClinicPhone", "StartCc", "StartBcc"]) assert.equal(C.DEFAULTS[k], "", k);
  const all = JSON.stringify(C.DEFAULTS) + C.START_TPL.startEmail + JSON.stringify(TPL.SPEC.map(s => s.rows));
  assert.ok(!/castro|helpcastro|amreed|3400|964-5419/i.test(all), "no Castro names, emails or numbers in defaults, the start email or new tables");
  assert.ok(C.START_TPL.startEmail.includes("{signature}") && C.START_TPL.startEmail.includes("{phone}") && C.START_TPL.startEmail.includes("{clinic} help email"));
  assert.ok(TPL.SPEC.some(s => s.name === "tblIvkUsers" && s.anchor === "GR"));
});

test("the clinic from the workbook name", () => {
  assert.equal(C.clinicFromName("Castro IVK Patient Tracker.xlsx"), "Castro");
  assert.equal(C.clinicFromName("Walnut Creek IVK Tracker 2026"), "Walnut Creek");
  assert.equal(C.clinicFromName("Sacramento IVK Tracker.xlsx"), "Sacramento");
  assert.equal(C.clinicFromName("Oak IVK Patient Tracker"), "Oak");
  assert.equal(C.clinicFromName("San_Rafael-Ketamine Tracker"), "San Rafael");
  assert.equal(C.clinicFromName("IVK Patient Tracker"), "");
  assert.equal(C.clinicFromName("Book1.xlsx"), "");
  assert.equal(C.clinicFromName(""), "");
});

test("saved config keeps its values; blank defaults fill only what's missing", () => {
  // the same merge app.js derive() does: defaults, then every non-blank saved value
  const merge = saved => { const cfg = Object.assign({}, C.DEFAULTS); for (const [k, v] of Object.entries(saved)) if (C.txt(v) !== "") cfg[k] = v; return cfg; };
  const a = merge({ Clinic: "Testville", StartCc: "help.test@example.com", Chairs: "3" });
  assert.equal(a.Clinic, "Testville"); assert.equal(a.StartCc, "help.test@example.com"); assert.equal(a.Chairs, "3"); assert.equal(a.StartBcc, ""); assert.equal(a.VisitMin, 120);
  assert.equal(merge({ Clinic: "" }).Clinic, "");
});

test("clinic-wide settings: anyone fills a blank one, only admin changes a saved one", () => {
  const saved = { Clinic: "Testville", StartCc: "", ClinicAddress: "1 Test St" };
  assert.equal(C.lockedCfg(saved, "Clinic", false), true);
  assert.equal(C.lockedCfg(saved, "Clinic", true), false);
  assert.equal(C.lockedCfg(saved, "StartCc", false), false);
  assert.equal(C.lockedCfg(saved, "StartBcc", false), false);
  assert.equal(C.lockedCfg(saved, "ClinicAddress", false), true);
  assert.equal(C.lockedCfg({ VisitMin: "90" }, "VisitMin", false), false);   // not clinic-wide
});

test("users: the workbook's list, the picker order, today's pick per workbook", () => {
  const U = C.userRows([C.USER_COLS, ["Sam", "2026-11-01T10:00:00Z", "", "{}"], ["ana", "", "", '{"sig":"<b>Ana</b>"}'], ["", "", "", ""], ["SAM", "", "", ""], ["Lee", "", "", "not json"]]);
  assert.deepEqual(U.map(u => u.name), ["Sam", "ana", "Lee"]);
  assert.equal(U[1].set.sig, "<b>Ana</b>"); assert.deepEqual(U[2].set, {});
  const L = C.pickOrder(U, ["Lee", "Someone Gone"]);
  assert.deepEqual(L.map(u => [u.name, u.here]), [["Lee", true], ["ana", false], ["Sam", false]]);
  assert.deepEqual(C.pickOrder(U, []).map(u => u.here), [false, false, false]);   // nobody pre-picked or flagged on a new computer
  const rec = C.whoRec(JSON.stringify({ day: "2026-11-04", name: "Sam", used: ["Sam"] }));
  assert.equal(C.whoToday(rec, "2026-11-05"), "");                                // yesterday's pick doesn't carry over
  assert.equal(C.whoToday(rec, "2026-11-04"), "Sam");
  const r2 = C.whoPick(rec, "Ana", "2026-11-05");
  assert.deepEqual(r2, { day: "2026-11-05", name: "Ana", used: ["Ana", "Sam"] });
  assert.deepEqual(C.whoPick(r2, "sam", "2026-11-05").used, ["sam", "Ana"]);
  assert.deepEqual(C.whoRec("garbage"), { day: "", name: "", used: [] });
  assert.equal(C.cleanName("  Sam   R.  "), "Sam R.");
});

test("workbook key: per file, stable, ignores query strings", () => {
  const a = C.wbKeyOf("https://example.sharepoint.com/sites/x/Castro IVK Patient Tracker.xlsx");
  assert.equal(a, C.wbKeyOf("https://example.sharepoint.com/sites/x/Castro IVK Patient Tracker.xlsx?web=1"));
  assert.notEqual(a, C.wbKeyOf("https://example.sharepoint.com/sites/x/Oak IVK Patient Tracker.xlsx"));
  assert.match(a, /^w[0-9a-z]+$/); assert.equal(C.wbKeyOf(""), "");
});

test("admin code", async () => {
  assert.equal(await C.adminOk("0582"), true);
  assert.equal(await C.adminOk(" 0582 "), true);
  assert.equal(await C.adminOk("0583"), false);
  assert.equal(await C.adminOk(""), false);
  assert.ok(!/0582/.test(require("fs").readFileSync(require.resolve("../core.js"), "utf8").replace(/\/\*[\s\S]*?\*\//g, "")), "the code isn't in core.js outside comments");
});

test("signatures: from Outlook's files, cleaned, added only when there is one", () => {
  const htm = `<html xmlns:o="urn:x"><head><style>p{}</style></head><body lang=EN-US><div class=WordSection1><p class=MsoNormal onclick="x()">Sam R.<o:p></o:p></p>
    <p>Front desk &amp; scheduling</p><img src="Sam_files/image001.png"><img src="https://example.com/logo.png"><script>alert(1)</script><a href="javascript:alert(1)">x</a></div></body></html>`;
  const s = C.sigFromFile(htm, true);
  assert.ok(s.includes("Sam R.") && s.includes("Front desk &amp; scheduling"));
  assert.ok(!/script|onclick|o:p|image001|<style|javascript:/i.test(s)); assert.ok(s.includes("https://example.com/logo.png"));
  assert.equal(C.sigFromFile("Sam R.\r\nFront desk <IVK>", false), "Sam R.<br>Front desk &lt;IVK&gt;");
  assert.equal(C.sigFromFile("   ", false), "");
  const body = "<p>Hi</p>{signature}<p>footer</p>";
  assert.equal(C.withSignature(body, "", true), "<p>Hi</p><p>footer</p>");
  assert.equal(C.withSignature(body, "<b>Sam</b>", true), '<p>Hi</p><div style="margin:12pt 0 0"><b>Sam</b></div><p>footer</p>');
  assert.equal(C.withSignature("<p>Hi</p>", "<b>Sam</b>", true), '<p>Hi</p><div style="margin:12pt 0 0"><b>Sam</b></div>');   // saved templates without the token
  assert.equal(C.withSignature("Hi", "Sam R.<br>Desk &amp; more", false), "Hi\n\nSam R.\nDesk & more");
  assert.equal(C.withSignature("Hi", "", false), "Hi");
});
