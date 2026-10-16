// Tests mit Michaels echten Kurstagen aus Oktober 2026.
// Erwartet wird jeweils das, was er bisher von Hand angelegt hat.
import { test } from "node:test";
import assert from "node:assert/strict";
import { planeTag, vergleiche } from "../src/planer.js";
import { einstellungen } from "../src/einstellungen.js";

const SVE = "SVE Sport- & Gesundheitscenter\nEttenkircher Straße 17, 88048 Friedrichshafen, Deutschland";
const PTM = "Physio-Team Markdorf\nRavensburgerstraße 36\n88677 Markdorf\nDeutschland";

const kurs = (titel, datum, von, bis, kategorie, ort) => ({
  titel, start: `${datum}T${von}`, ende: `${datum}T${bis}`, kategorien: [kategorie], ort, generiert: false,
});
const kurz = (bloecke) => bloecke.map((b) => `${b.start.slice(11)}-${b.ende.slice(11)} ${b.titel}`);

test("14.10. SVE: zwei Kurse mit 15 min Lücke wie bisher von Hand", () => {
  const d = "2026-10-14";
  const termine = [
    kurs("RehaSport (Michael)", d, "08:45", "09:30", "10 SVE: RehaSport", SVE),
    kurs("RehaSport (Michael)", d, "09:45", "10:30", "10 SVE: RehaSport", SVE),
  ];
  assert.deepEqual(kurz(planeTag(d, termine, einstellungen)), [
    "08:10-08:30 Anfahrt - RehaSport (Michael)",
    "08:30-08:45 Vorbereitung - RehaSport (Michael)",
    "09:30-09:45 Nach-/Vorbereitung - RehaSport (Michael)",
    "10:30-10:45 Nachbereitung - RehaSport (Michael)",
    "10:45-11:05 Rückfahrt - RehaSport (Michael)",
  ]);
});

test("15.10. PTM: zwei Kurse ohne Lücke, kein Puffer dazwischen", () => {
  const d = "2026-10-15";
  const termine = [
    kurs("RehaSport / PUR (Michael)", d, "11:00", "11:45", "01 PTM: PUR Kurs", PTM),
    kurs("RehaSport / PUR-Senioren (Michael)", d, "11:45", "12:30", "01 PTM: PUR Kurs", PTM),
  ];
  assert.deepEqual(kurz(planeTag(d, termine, einstellungen)), [
    "10:30-10:45 Anfahrt - RehaSport / PUR (Michael)",
    "10:45-11:00 Vorbereitung - RehaSport / PUR (Michael)",
    "12:30-12:45 Nachbereitung - RehaSport / PUR-Senioren (Michael)",
    "12:45-13:00 Rückfahrt - RehaSport / PUR-Senioren (Michael)",
  ]);
});

test("28.10. SVE: Vertretungen direkt nach eigenen Kursen werden ein Block", () => {
  const d = "2026-10-28";
  const termine = [
    kurs("RehaSport (Michael)", d, "08:45", "09:30", "10 SVE: RehaSport", SVE),
    kurs("RehaSport (Michael)", d, "09:45", "10:30", "10 SVE: RehaSport", SVE),
    kurs("(V) RehaSport (Eva)", d, "10:45", "11:30", "19 Vertretung Kurs SVE", SVE),
    kurs("(V) RehaSport (Eve)", d, "11:45", "12:30", "19 Vertretung Kurs SVE", SVE),
  ];
  assert.deepEqual(kurz(planeTag(d, termine, einstellungen)), [
    "08:10-08:30 Anfahrt - RehaSport (Michael)",
    "08:30-08:45 Vorbereitung - RehaSport (Michael)",
    "09:30-09:45 Nach-/Vorbereitung - RehaSport (Michael)",
    "10:30-10:45 Nach-/Vorbereitung - (V) RehaSport (Eva)",
    "11:30-11:45 Nach-/Vorbereitung - (V) RehaSport (Eve)",
    "12:30-12:45 Nachbereitung - (V) RehaSport (Eve)",
    "12:45-13:05 Rückfahrt - (V) RehaSport (Eve)",
  ]);
});

test("Urlaub, Ferien und Absage bekommen keine Blöcke", () => {
  const d = "2026-10-26";
  for (const k of ["U", "F", "A"]) {
    const termine = [kurs(`(${k}) RehaSport / PUR (Michael)`, d, "15:00", "15:45", "01 PTM: PUR Kurs", PTM)];
    assert.deepEqual(planeTag(d, termine, einstellungen), [], `Kürzel ${k}`);
  }
});

test("Kurs verschoben: alte Blöcke weg, neue an der richtigen Stelle", () => {
  const d = "2026-10-14";
  const alt = [kurs("RehaSport (Michael)", d, "08:45", "09:30", "10 SVE: RehaSport", SVE)];
  const vorhandeneBloecke = planeTag(d, alt, einstellungen);
  const verschoben = [kurs("RehaSport (Michael)", d, "10:00", "10:45", "10 SVE: RehaSport", SVE)];
  const { anlegen, loeschen } = vergleiche(planeTag(d, verschoben, einstellungen), [...verschoben, ...vorhandeneBloecke]);
  assert.equal(loeschen.length, 4);
  assert.deepEqual(kurz(anlegen), [
    "09:25-09:45 Anfahrt - RehaSport (Michael)",
    "09:45-10:00 Vorbereitung - RehaSport (Michael)",
    "10:45-11:00 Nachbereitung - RehaSport (Michael)",
    "11:00-11:20 Rückfahrt - RehaSport (Michael)",
  ]);
});

test("Kurs gelöscht: seine Blöcke werden entfernt, private Termine bleiben", () => {
  const d = "2026-10-14";
  const blöcke = planeTag(d, [kurs("RehaSport (Michael)", d, "08:45", "09:30", "10 SVE: RehaSport", SVE)], einstellungen);
  const privat = { titel: "Dr. Kirsner", start: `${d}T12:00`, ende: `${d}T12:30`, kategorien: [], generiert: false };
  const { anlegen, loeschen } = vergleiche(planeTag(d, [privat], einstellungen), [privat, ...blöcke]);
  assert.equal(anlegen.length, 0);
  assert.equal(loeschen.length, 4);
  assert.ok(!loeschen.includes(privat));
});

test("Zwei Anbieter kurz hintereinander: direkte Fahrt statt Heimfahrt", () => {
  const d = "2026-10-14";
  const termine = [
    kurs("RehaSport (Michael)", d, "08:45", "09:30", "10 SVE: RehaSport", SVE),
    kurs("RehaSport / PUR (Michael)", d, "10:30", "11:15", "01 PTM: PUR Kurs", PTM),
  ];
  assert.deepEqual(kurz(planeTag(d, termine, einstellungen)), [
    "08:10-08:30 Anfahrt - RehaSport (Michael)",
    "08:30-08:45 Vorbereitung - RehaSport (Michael)",
    "09:30-09:45 Nachbereitung - RehaSport (Michael)",
    "10:00-10:15 Fahrt SVE → PTM - RehaSport / PUR (Michael)",
    "10:15-10:30 Vorbereitung - RehaSport / PUR (Michael)",
    "11:15-11:30 Nachbereitung - RehaSport / PUR (Michael)",
    "11:30-11:45 Rückfahrt - RehaSport / PUR (Michael)",
  ]);
});
