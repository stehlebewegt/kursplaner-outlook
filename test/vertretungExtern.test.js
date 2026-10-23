// Vertretungen in anderen Vereinen (wechselnd, Michael 23.10.2026): Ort und
// Fahrzeit kommen aus dem Termin, sonst wie ein Kurs bei einem Anbieter.
import { test } from "node:test";
import assert from "node:assert/strict";
import { planeTag, kursAnbieter, ohneFahrzeit } from "../src/planer.js";
import { kursTitel, pruefeKurs } from "../src/neuerKurs.js";
import { streckenPaare } from "../automatik/strecken.js";
import { fahrtenJahr } from "../src/fahrtenbuch.js";
import { einstellungen } from "../src/einstellungen.js";

const KAT = "20 Vertretung andere Vereine";
const HALLE = "Turnhalle Beispielschule, Schulstraße 1, 88074 Meckenbeuren";
const SVE = "SVE Sport- & Gesundheitscenter\nEttenkircher Straße 17, 88048 Friedrichshafen, Deutschland";
const termin = (titel, datum, von, bis, kategorie, ort, extra = {}) => ({
  titel, start: `${datum}T${von}`, ende: `${datum}T${bis}`, kategorien: [kategorie], ort, generiert: false, ...extra,
});
const kurz = (bloecke) => bloecke.map((b) => `${b.start.slice(11)}-${b.ende.slice(11)} ${b.titel}`);

test("Vertretung im anderen Verein: Fahrten mit Fahrzeit aus dem Termin, Navigation zur Halle", () => {
  const d = "2026-11-03";
  const t = [termin("(V) RehaSport (TSV Beispiel)", d, "18:00", "18:45", KAT, HALLE, { fahrzeit: 25 })];
  const soll = planeTag(d, t, einstellungen);
  assert.deepEqual(kurz(soll), [
    "17:20-17:45 🚗 Jetzt losfahren - (V) RehaSport (TSV Beispiel)",
    "17:45-18:00 Vorbereitung - (V) RehaSport (TSV Beispiel)",
    "18:45-19:00 Nachbereitung - (V) RehaSport (TSV Beispiel)",
    "19:00-19:25 Rückfahrt - (V) RehaSport (TSV Beispiel)",
  ]);
  assert.match(soll[0].html, /Schulstra%C3%9Fe/);
  assert.equal(soll[0].ort, HALLE);
});

test("Ohne berechnete Fahrzeit gilt die Standard-Fahrzeit, und die Automatik soll sie berechnen", () => {
  const d = "2026-11-03";
  const t = [termin("(V) RehaSport (TSV Beispiel)", d, "18:00", "18:45", KAT, HALLE, { fahrzeit: null, adresse: HALLE })];
  assert.equal(kurz(planeTag(d, t, einstellungen))[0], "17:15-17:45 🚗 Jetzt losfahren - (V) RehaSport (TSV Beispiel)");
  assert.equal(ohneFahrzeit(t, einstellungen).length, 1);
  assert.equal(kursAnbieter(t[0], einstellungen).kuerzel, "VTR");
});

test("Vom SVE direkt weiter zum anderen Verein", () => {
  const d = "2026-11-04";
  const t = [
    termin("RehaSport (Michael)", d, "16:00", "16:45", "10 SVE: RehaSport", SVE),
    termin("(V) RehaSport (TSV Beispiel)", d, "17:30", "18:15", KAT, HALLE, { fahrzeit: 25 }),
  ];
  assert.ok(kurz(planeTag(d, t, einstellungen)).includes("16:50-17:15 Fahrt SVE → VTR - (V) RehaSport (TSV Beispiel)"));
});

test("Strecken je Halle, Formular prüft Verein und Adresse", () => {
  const t = [termin("(V) RehaSport (TSV Beispiel)", "2026-11-03", "18:00", "18:45", KAT, HALLE)];
  const schluessel = streckenPaare(einstellungen, t).map((p) => p[0]);
  assert.ok(schluessel.includes(`heim-VTR@${HALLE}`));
  assert.ok(schluessel.includes(`SVE-VTR@${HALLE}`));
  const f = { extern: true, anbieter: einstellungen.vertretungExtern, name: "RehaSport", leitung: "", ort: "", datum: "2026-11-03", beginn: "18:00", ende: "18:45" };
  assert.deepEqual(pruefeKurs(f), ["Verein fehlt.", "Adresse fehlt (für Fahrzeit und Navigation)."]);
  assert.equal(kursTitel({ ...f, leitung: "TSV Beispiel" }), "(V) RehaSport (TSV Beispiel)");
});

test("Fahrten fürs Finanzamt enthalten die Vertretung mit Ort", () => {
  const d = "2026-11-03";
  const kurs = termin("(V) RehaSport (TSV Beispiel)", d, "18:00", "18:45", KAT, HALLE, { fahrzeit: 25 });
  const bloecke = planeTag(d, [kurs], einstellungen).map((b) => ({ ...b, ort: b.ort }));
  const f = fahrtenJahr([kurs, ...bloecke], "2026", "2026-12-31T00:00", einstellungen, (v, n) =>
    v === "heim" || n === "heim" ? { km: 12.3, minuten: 21 } : null
  );
  assert.equal(f.reihen.length, 1);
  assert.equal(f.reihen[0].anbieter, "VTR");
  assert.equal(f.reihen[0].fahrten, 2);
  assert.match(f.reihen[0].strecke, /Turnhalle Beispielschule/);
});
