// node --test tests/   the start email's directions map: pulled out of a pasted template, placed once, gone when there's no map
const test = require("node:test"), assert = require("node:assert/strict");
const C = require("../core.js");

const PNG = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==";

test("the standard start email has {map} right after the schedule", () => {
  const t = C.START_TPL.startEmail;
  assert.ok(t.includes("{schedule}\n{map}"));
  assert.equal(t.match(/\{map\}/g).length, 1);
});

test("a pasted data: picture becomes the map, and {map} takes its place", () => {
  const x = C.tplImages(`<p>Hi</p><img width="400" src="data:image/png;base64,${PNG.slice(0, 20)}\n${PNG.slice(20)}" alt="map"><p>Bye</p>`);
  assert.equal(x.html, "<p>Hi</p>{map}<p>Bye</p>");
  assert.deepEqual(x.img, { type: "image/png", b64: PNG });
  assert.equal(x.lost, 0);
});

test("jpg is saved as jpeg; a second picture is dropped, not a second {map}", () => {
  const x = C.tplImages(`<img src='data:image/jpg;base64,AAAA'><img src="data:image/png;base64,BBBB">`);
  assert.equal(x.html, "{map}");
  assert.deepEqual(x.img, { type: "image/jpeg", b64: "AAAA" });
});

test("pictures that can't travel (cid:, blob:, file:) become {map}; web pictures stay", () => {
  const x = C.tplImages(`<img src="cid:image001.png@01DB"><p>x</p><img src="https://example.org/a.png"><img src="blob:https://x/1">`);
  assert.equal(x.html, `{map}<p>x</p><img src="https://example.org/a.png">`);
  assert.equal(x.img, null); assert.equal(x.lost, 2);
});

test("a template that already has {map} keeps that one and loses the pasted picture", () => {
  const x = C.tplImages(`<p>{map}</p><img src="data:image/png;base64,${PNG}">`);
  assert.equal(x.html, "<p>{map}</p>");
  assert.equal(x.img.b64, PNG);
});

test("no pictures, nothing changes", () => {
  const t = C.START_TPL.startEmail, x = C.tplImages(t);
  assert.equal(x.html, t); assert.equal(x.img, null); assert.equal(x.lost, 0);
});

test("withMap: one {map} with a map, none without", () => {
  assert.equal(C.withMap("<p>a</p>{schedule}\n{map}\n<p>b</p>", false), "<p>a</p>{schedule}\n<p>b</p>");
  assert.equal(C.withMap("<p>a</p><p>{map}</p><p>b</p>", false), "<p>a</p><p>b</p>");
  assert.equal(C.withMap("x{map}y{map}", true), "x{map}y");
  /* saved templates from before {map}: after the schedule, else before the signature, else at the end */
  assert.equal(C.withMap("<p>a</p>{schedule}<p>b</p>", true), "<p>a</p>{schedule}\n{map}<p>b</p>");
  assert.equal(C.withMap("<p>a</p>{signature}", true), "<p>a</p>{map}\n{signature}");
  assert.equal(C.withMap("<p>a</p>", true), "<p>a</p>\n{map}");
});
