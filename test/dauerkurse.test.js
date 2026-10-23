import { test } from "node:test";
import assert from "node:assert/strict";
import { gleicheAb, istMeinKurs } from "../src/dauerkurse.js";
import { einstellungen } from "../src/einstellungen.js";

const PUR = ["01 PTM: PUR Kurs"];
const k = (titel, start, ende, extra = {}) => ({ titel, start, ende, kategorien: PUR, ort: "", ...extra });

test("istMeinKurs erkennt die Kursleitung am Titelende, auch mit Leerzeichen", () => {
  assert.ok(istMeinKurs("RehaSport / PUR (Michael) ", einstellungen));
  assert.ok(!istMeinKurs("RehaSport (Eva)", einstellungen));
  assert.ok(!istMeinKurs("Michael Stehle - Termin", einstellungen));
});

test("vorhandene Kopien (auch mit Kürzel) passen, fehlende werden angelegt", () => {
  const anbieter = [
    k("RehaSport / PUR (Michael)", "2026-10-27T14:15", "2026-10-27T15:00"),
    k("RehaSport / PUR (Michael)", "2026-10-27T15:00", "2026-10-27T15:45"),
    k("RehaSport / PUR (Michael) ", "2026-11-03T15:00", "2026-11-03T15:45"),
    k("RehaSport / PUR (Agne)", "2026-10-27T11:00", "2026-10-27T11:45"),
  ];
  const eigene = [
    k("(F) RehaSport / PUR (Michael)", "2026-10-27T14:15", "2026-10-27T15:00"),
    k("(F) RehaSport / PUR (Michael)", "2026-10-27T15:00", "2026-10-27T15:45"),
  ];
  const r = gleicheAb(anbieter, eigene, einstellungen);
  assert.deepEqual(r.anlegen.map((t) => t.start), ["2026-11-03T15:00"]);
  assert.equal(r.verschieben.length, 0);
  assert.equal(r.ohneGegenstueck.length, 0);
});

test("beim Anbieter verschoben: eigene Kopie zieht mit", () => {
  const anbieter = [k("RehaSport (Michael)", "2026-11-04T16:00", "2026-11-04T16:45", { originalStart: "2026-11-04T15:00" })];
  const eigene = [k("RehaSport (Michael)", "2026-11-04T15:00", "2026-11-04T15:45")];
  const r = gleicheAb(anbieter, eigene, einstellungen);
  assert.equal(r.anlegen.length, 0);
  assert.deepEqual(r.verschieben.map((v) => [v.termin.start, v.start]), [["2026-11-04T15:00", "2026-11-04T16:00"]]);
});

test("selbst verschobene Kopie bleibt, eigene ohne Gegenstück werden nur gezählt", () => {
  const anbieter = [k("RehaSport (Michael)", "2026-11-02T15:00", "2026-11-02T15:45")];
  const eigene = [
    k("RehaSport (Michael)", "2026-11-02T14:00", "2026-11-02T14:45"),
    k("RehaSport (Michael)", "2026-11-05T15:00", "2026-11-05T15:45"),
    { ...k("🚗 Jetzt losfahren - RehaSport (Michael)", "2026-11-02T13:45", "2026-11-02T14:00"), generiert: true },
  ];
  const r = gleicheAb(anbieter, eigene, einstellungen);
  assert.equal(r.anlegen.length, 0);
  assert.equal(r.verschieben.length, 0);
  assert.deepEqual(r.ohneGegenstueck.map((t) => t.start), ["2026-11-05T15:00"]);
});
