import { test } from "node:test";
import assert from "node:assert/strict";
import { pruefeKurs, kursEreignis, anzahlTermine, wochentagVon } from "../src/neuerKurs.js";
import { einstellungen } from "../src/einstellungen.js";

const ptm = einstellungen.anbieter.find((a) => a.kuerzel === "PTM");
const basis = { anbieter: ptm, kategorie: "02 PTM: FLEX Kurs", name: "RehaSport / FLEX", leitung: "Michael", datum: "2026-11-03", beginn: "17:00", ende: "17:45", zeitzone: "W. Europe Standard Time" };

test("Neuer Kurs: Einzeltermin im Anbieter-Format", () => {
  const e = kursEreignis(basis);
  assert.equal(e.subject, "RehaSport / FLEX (Michael)");
  assert.equal(e.start.dateTime, "2026-11-03T17:00:00");
  assert.equal(e.location.displayName, ptm.ort);
  assert.deepEqual(e.categories, ["02 PTM: FLEX Kurs"]);
  assert.match(e.body.content, /^Offizielle Erstvertretung:/);
  assert.equal(e.recurrence, undefined);
  assert.equal(anzahlTermine(basis), 1);
});

test("Neuer Kurs: wöchentliche Serie bis Datum oder ohne Ende", () => {
  const e = kursEreignis({ ...basis, serie: true, bis: "2026-12-15" });
  assert.deepEqual(e.recurrence.pattern.daysOfWeek, ["tuesday"]);
  assert.equal(e.recurrence.range.endDate, "2026-12-15");
  assert.equal(anzahlTermine({ ...basis, serie: true, bis: "2026-12-15" }), 7);
  assert.equal(kursEreignis({ ...basis, serie: true }).recurrence.range.type, "noEnd");
  assert.equal(wochentagVon("2026-10-18"), "sunday");
});

test("Neuer Kurs: Pflichtfelder und Zeiten werden geprüft", () => {
  assert.deepEqual(pruefeKurs(basis), []);
  assert.equal(pruefeKurs({ ...basis, name: " " }).length, 1);
  assert.deepEqual(pruefeKurs({ ...basis, ende: "16:00" }), ["Das Ende muss nach dem Beginn liegen."]);
  assert.equal(pruefeKurs({ ...basis, serie: true, bis: "2026-10-01" }).length, 1);
});
