import { test } from "node:test";
import assert from "node:assert/strict";
import { leseSignatur, signaturHtml, mitSignatur, htmlZuText } from "../src/signatur.js";

const roh = "Bitte nicht löschen.\nName: Max Muster\nMotto: Test & Co\nSlogan: Bewegt\nTelefon: 0123 456\nE-Mail: max@example.org\nInternet: www.example.org\nLogo:\nHinweis: Vertraulich";

test("liest Felder, leere weg", () => {
  const f = leseSignatur(roh);
  assert.equal(f.Name, "Max Muster");
  assert.equal(f["E-Mail"], "max@example.org");
  assert.equal(f.Logo, undefined);
  assert.equal(leseSignatur("nix"), null);
});

test("HTML-Entwurf wird gelesen", () => {
  assert.equal(leseSignatur(htmlZuText("<p>Name: Max</p><p>Slogan: A &amp; B</p>")).Slogan, "A & B");
});

test("HTML escaped und verlinkt", () => {
  const h = signaturHtml(leseSignatur(roh));
  assert.match(h, /Test &amp; Co/);
  assert.match(h, /mailto:max@example.org/);
  assert.match(h, /https:\/\/www.example.org/);
  assert.match(h, /tel:0123456/);
  assert.doesNotMatch(h, /<img/);
});

test("mitSignatur", () => {
  assert.deepEqual(mitSignatur("a", null), { contentType: "text", content: "a" });
  const b = mitSignatur("a<b\nc", leseSignatur(roh));
  assert.equal(b.contentType, "html");
  assert.match(b.content, /a&lt;b<br>c/);
});
