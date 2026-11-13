// app.js is one script: a second function with the same name, or a repeated MODALS / ACT key, silently replaces the first.
// (Two startsHTML functions once broke Today: the Index planner's replaced Today's.)
const test = require("node:test"), assert = require("node:assert/strict");
const src = require("fs").readFileSync(require("path").join(__dirname, "..", "app.js"), "utf8");
const line = i => src.slice(0, i).split("\n").length;

test("app.js declares each function name once", () => {
  const seen = new Map(), dup = [];
  for (const m of src.matchAll(/^\s*(?:async\s+)?function\s+([A-Za-z_$][\w$]*)\s*\(/gm)) { if (seen.has(m[1])) dup.push(`${m[1]} (lines ${seen.get(m[1])} and ${line(m.index)})`); else seen.set(m[1], line(m.index)); }
  assert.deepEqual(dup, []);
});

test("MODALS, ACT and ACT2 have each key once", () => {
  for (const name of ["MODALS", "ACT", "ACT2"]) {
    const a = src.indexOf(`  const ${name} = {`); assert.ok(a >= 0, `${name} is defined`);
    let i = src.indexOf("{", a), d = 0, j = i; for (; j < src.length; j++) { const c = src[j]; if (c === "{") d++; else if (c === "}" && !--d) break; }
    const seen = new Set(), dup = [];
    for (const m of src.slice(i + 1, j).matchAll(/^    (?:async\s+)?([A-Za-z_$][\w$]*)\s*(?::|\()/gm)) { if (seen.has(m[1])) dup.push(m[1]); seen.add(m[1]); }
    assert.deepEqual(dup, [], `${name} repeats a key`);
  }
});
