// Tests mit Michaels echten Kurstagen aus Oktober 2026.
// Erwartet wird jeweils das, was er bisher von Hand angelegt hat.
import { test } from "node:test";
import assert from "node:assert/strict";
import { planeTag, vergleiche, markiereAlteBloecke } from "../src/planer.js";
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

test("Erster Lauf: von Hand angelegte Blöcke werden übernommen, nichts doppelt", () => {
  const d = "2026-10-14";
  const ORT = "SVE Sport- & Gesundheitscenter, Ettenkircher Straße 17, 88048 Friedrichshafen, Deutschland";
  const block = (titel, von, bis, kat) => ({ titel, start: `${d}T${von}`, ende: `${d}T${bis}`, kategorien: [kat], ort: ORT, generiert: false });
  const termine = [
    block("Anfahrt - RehaSport (Michael) ", "08:10", "08:30", "90 🚗 Fahrzeit"),
    block("Vorbereitung - RehaSport (Michael) ", "08:30", "08:45", "Vor-/Nachbereitung"),
    kurs("RehaSport (Michael) ", d, "08:45", "09:30", "10 SVE: RehaSport", SVE),
    block("Nach-/Vorbereitung - RehaSport (Michael)", "09:30", "09:45", "Vor-/Nachbereitung"),
    kurs("RehaSport (Michael)", d, "09:45", "10:30", "10 SVE: RehaSport", SVE),
    block("Nachbereitung - RehaSport (Michael)", "10:30", "10:45", "Vor-/Nachbereitung"),
    block("Rückfahrt - RehaSport (Michael)", "10:45", "11:05", "90 🚗 Fahrzeit"),
    { titel: "Anfahrt - Dr. Kirsner (Blutabnahme)", start: `${d}T12:00`, ende: `${d}T12:20`, kategorien: ["90 🚗 Fahrzeit"], ort: "Praxis Dr. Med. Andreas Kirsner", generiert: false },
  ];
  const markiert = markiereAlteBloecke(termine, einstellungen);
  const { anlegen, loeschen } = vergleiche(planeTag(d, markiert, einstellungen), markiert);
  assert.equal(anlegen.length, 0);
  assert.equal(loeschen.length, 0);
  assert.equal(markiert[7].generiert, false, "private Anfahrt bleibt unberührt");
});

test("10.11. PTM: 75 min Pause beim selben Anbieter, Michael bleibt vor Ort", () => {
  const d = "2026-11-10";
  const termine = [
    kurs("(V) RehaSport / PUR (Agne)", d, "11:00", "11:45", "01 PTM: PUR Kurs", PTM),
    kurs("(V) RehaSport / PUR-Senioren (Agne)", d, "11:45", "12:30", "01 PTM: PUR Kurs", PTM),
    kurs("RehaSport / PUR (Michael)", d, "14:15", "15:00", "01 PTM: PUR Kurs", PTM),
  ];
  assert.deepEqual(kurz(planeTag(d, termine, einstellungen)), [
    "10:30-10:45 Anfahrt - (V) RehaSport / PUR (Agne)",
    "10:45-11:00 Vorbereitung - (V) RehaSport / PUR (Agne)",
    "12:30-12:45 Nachbereitung - (V) RehaSport / PUR-Senioren (Agne)",
    "14:00-14:15 Vorbereitung - RehaSport / PUR (Michael)",
    "15:00-15:15 Nachbereitung - RehaSport / PUR (Michael)",
    "15:15-15:30 Rückfahrt - RehaSport / PUR (Michael)",
  ]);
});

test("Besprechung am Ort eines Anbieters, aber ohne Kurs-Kategorie: keine Blöcke", () => {
  const d = "2026-11-06";
  const termine = [{ titel: "Besprechung PTM", start: `${d}T12:00`, ende: `${d}T13:00`, kategorien: [], ort: PTM, generiert: false }];
  assert.deepEqual(planeTag(d, termine, einstellungen), []);
});

test("26.11.: alte Rückfahrt mit Heimadresse als Ort wird übernommen, nicht verdoppelt", () => {
  const d = "2026-11-26";
  const termine = markiereAlteBloecke(
    [
      kurs("RehaSport / PUR-Senioren (Michael)", d, "11:45", "12:30", "01 PTM: PUR Kurs", PTM),
      { titel: "Rückfahrt - RehaSport / PUR-Senioren (Michael)", start: `${d}T12:30`, ende: `${d}T13:00`, kategorien: ["90 🚗 Fahrzeit"], ort: "Tannenriedweg 14, 88048 Friedrichshafen", generiert: false },
      { titel: "Rückfahrt - Kontrolluntersuchung & Zahnstein", start: `${d}T15:30`, ende: `${d}T15:55`, kategorien: ["90 🚗 Fahrzeit"], ort: "Bahnhofstr. 1, 88682 Salem", generiert: false },
    ],
    einstellungen
  );
  const { loeschen } = vergleiche(planeTag(d, termine, einstellungen), termine);
  assert.deepEqual(kurz(loeschen), ["12:30-13:00 Rückfahrt - RehaSport / PUR-Senioren (Michael)"]);
});
