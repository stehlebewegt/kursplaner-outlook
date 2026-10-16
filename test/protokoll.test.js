import { test } from "node:test";
import assert from "node:assert/strict";
import { leseProtokoll, schreibeProtokoll } from "../src/protokoll.js";

const AUSGEFUELLT =
  "Offizielle Erstvertretung: für Eva\r\n\r\nThema: Beweglichkeit mit Pilates Ball\r\n\r\nMaterial: Pilates Ball\r\n\r\n" +
  "Teilnehmer gesamt: 18\r\nAnwesend: 9\r\nEntschuldigt: 3\r\nUnentschuldigt: 6\r\n\r\nVorkommnis: nein\r\n\r\nBesonderheiten:";
const LEER = "Thema:\r\n\r\nMaterial:\r\n\r\nTeilnehmer:\r\nAnwesend:\r\nEntschuldigt:\r\nUnentschuldigt:\r\n\r\nBesonderheiten:";

test("liest ein ausgefülltes Protokoll von Michael", () => {
  const p = leseProtokoll(AUSGEFUELLT);
  assert.equal(p.erstvertretung, "für Eva");
  assert.equal(p.thema, "Beweglichkeit mit Pilates Ball");
  assert.equal(p.anwesend, "9");
  assert.equal(p.unentschuldigt, "6");
  assert.equal(p.vorkommnis, "nein");
  assert.equal(p.besonderheiten, "");
});

test("leere Vorlage gibt leere Felder", () => {
  const p = leseProtokoll(LEER);
  assert.ok(Object.values(p).every((v) => v === ""));
});

test("schreiben rechnet Gesamt aus und lässt sich wieder lesen", () => {
  const text = schreibeProtokoll({ ...leseProtokoll(LEER), thema: "Rücken", anwesend: "10", entschuldigt: "4", unentschuldigt: "1", vorkommnis: "nein", besonderheiten: "Zeile 1\nZeile 2" });
  assert.match(text, /Teilnehmer gesamt: 15/);
  const p = leseProtokoll(text);
  assert.equal(p.gesamt, "15");
  assert.equal(p.besonderheiten, "Zeile 1\nZeile 2");
});
