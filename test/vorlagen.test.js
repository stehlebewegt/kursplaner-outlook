import { test } from "node:test";
import assert from "node:assert/strict";
import { fuelleVorlage } from "../src/vorlagen.js";
import { einstellungen } from "../src/einstellungen.js";

test("Vorlage: Kurs, Datum und Uhrzeit werden eingesetzt", () => {
  const termin = { titel: "(V) RehaSport / PUR (Michael)", start: "2026-11-05T15:00", ende: "2026-11-05T15:45" };
  const v = einstellungen.vorlagen.find((x) => x.id === "vertretung");
  const m = fuelleVorlage(v, termin, einstellungen.anbieter[1]);
  assert.equal(m.betreff, "Vertretung gesucht: RehaSport / PUR am Donnerstag, 05.11.");
  assert.match(m.text, /am Donnerstag, 05\.11\.2026, von 15:00–15:45 Uhr \(Physio-Team Markdorf\)/);
});
