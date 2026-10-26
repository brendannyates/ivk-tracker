// node --test tests/
// The bay plan block in bayplan.js must stay byte-identical to bay-map-builder/render.js and esk-tracker/src/taskpane.html.
// Set BM_BUILDER=<path to bay-map-builder/render.js> and/or BM_ESK=<path to esk-tracker/src/taskpane.html> to check against them.
const test = require("node:test"), assert = require("node:assert/strict"), fs = require("node:fs"), path = require("node:path");
const B = require("../bayplan.js");
const block = s => { const a = s.indexOf("// ==== bay plan:"), b = s.indexOf("// ==== end bay plan ===="); return a < 0 || b < 0 ? null : s.slice(a, b).replace(/\r\n/g, "\n"); };
const mine = block(fs.readFileSync(path.join(__dirname, "..", "bayplan.js"), "utf8"));

test("shared block is present", () => assert.ok(mine && mine.length > 1000));
for (const env of ["BM_BUILDER", "BM_ESK"]) {
  test(`shared block matches ${env}`, { skip: !process.env[env] }, () => assert.equal(mine, block(fs.readFileSync(process.env[env], "utf8"))));
}
test("a map file is checked before it's drawn", () => {
  assert.ok(B.bpCheck(B.BP_DEFAULT).plan);
  assert.match(B.bpCheck({ schema: "nope" }).err, /isn’t a bay map/);
  const bad = JSON.parse(JSON.stringify(B.BP_DEFAULT)); bad.bays[0].id = "Bay 1"; assert.match(B.bpCheck(bad).err, /can’t be used/);
  const ok = B.bpCheck(B.BP_DEFAULT).plan; assert.deepEqual(ok.bays.map(b => b.id), ["1", "2", "3", "4", "5"]);
});
