// Tests mit Michaels echten Kurstagen aus Oktober 2026.
// Erwartet wird jeweils das, was er bisher von Hand angelegt hat.
import { test } from "node:test";
import assert from "node:assert/strict";
import { planeTag, termineFuerTag, vergleiche, markiereAlteBloecke } from "../src/planer.js";
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
    "08:10-08:30 🚗 Jetzt losfahren - RehaSport (Michael)",
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
    "10:35-10:45 🚗 Jetzt losfahren - RehaSport / PUR (Michael)",
    "10:45-11:00 Vorbereitung - RehaSport / PUR (Michael)",
    "12:30-12:45 Nachbereitung - RehaSport / PUR-Senioren (Michael)",
    "12:45-12:55 Rückfahrt - RehaSport / PUR-Senioren (Michael)",
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
    "08:10-08:30 🚗 Jetzt losfahren - RehaSport (Michael)",
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
    "09:25-09:45 🚗 Jetzt losfahren - RehaSport (Michael)",
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
    "08:10-08:30 🚗 Jetzt losfahren - RehaSport (Michael)",
    "08:30-08:45 Vorbereitung - RehaSport (Michael)",
    "09:30-09:45 Nachbereitung - RehaSport (Michael)",
    "09:55-10:15 Fahrt SVE → PTM - RehaSport / PUR (Michael)",
    "10:15-10:30 Vorbereitung - RehaSport / PUR (Michael)",
    "11:15-11:30 Nachbereitung - RehaSport / PUR (Michael)",
    "11:30-11:40 Rückfahrt - RehaSport / PUR (Michael)",
  ]);
});

test("Erster Lauf: von Hand angelegte Blöcke werden übernommen, nichts doppelt", () => {
  const d = "2026-10-14";
  const ORT = "SVE Sport- & Gesundheitscenter, Ettenkircher Straße 17, 88048 Friedrichshafen, Deutschland";
  const block = (titel, von, bis, kat) => ({ titel, start: `${d}T${von}`, ende: `${d}T${bis}`, kategorien: [kat], ort: ORT, generiert: false });
  const termine = [
    block("🚗 Jetzt losfahren - RehaSport (Michael) ", "08:10", "08:30", "90 🚗 Fahrzeit"),
    block("Vorbereitung - RehaSport (Michael) ", "08:30", "08:45", "Vor-/Nachbereitung"),
    kurs("RehaSport (Michael) ", d, "08:45", "09:30", "10 SVE: RehaSport", SVE),
    block("Nach-/Vorbereitung - RehaSport (Michael)", "09:30", "09:45", "Vor-/Nachbereitung"),
    kurs("RehaSport (Michael)", d, "09:45", "10:30", "10 SVE: RehaSport", SVE),
    block("Nachbereitung - RehaSport (Michael)", "10:30", "10:45", "Vor-/Nachbereitung"),
    block("Rückfahrt - RehaSport (Michael)", "10:45", "11:05", "90 🚗 Fahrzeit"),
    { titel: "🚗 Jetzt losfahren - Dr. Kirsner (Blutabnahme)", start: `${d}T12:00`, ende: `${d}T12:20`, kategorien: ["90 🚗 Fahrzeit"], ort: "Praxis Dr. Med. Andreas Kirsner", generiert: false },
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
    "10:35-10:45 🚗 Jetzt losfahren - (V) RehaSport / PUR (Agne)",
    "10:45-11:00 Vorbereitung - (V) RehaSport / PUR (Agne)",
    "12:30-12:45 Nachbereitung - (V) RehaSport / PUR-Senioren (Agne)",
    "14:00-14:15 Vorbereitung - RehaSport / PUR (Michael)",
    "15:00-15:15 Nachbereitung - RehaSport / PUR (Michael)",
    "15:15-15:25 Rückfahrt - RehaSport / PUR (Michael)",
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

test("Privater Termin mit Fahrzeit bekommt Hin- und Rückfahrt", () => {
  const d = "2026-10-26";
  const termine = [{ titel: "Podologie Bernhardt ", start: `${d}T10:20`, ende: `${d}T11:20`, kategorien: [], ort: "Salem", generiert: false, fahrzeit: 20 }];
  assert.deepEqual(kurz(planeTag(d, termine, einstellungen)), [
    "10:00-10:20 🚗 Jetzt losfahren - Podologie Bernhardt",
    "11:20-11:40 Rückfahrt - Podologie Bernhardt",
  ]);
});

test("Von Hand angelegte private Fahrten: Fahrzeit wird erkannt und die Blöcke übernommen", () => {
  const d = "2026-10-26";
  const ort = "Podologie Bernhardt, Am Schlosssee 5, 88682 Salem";
  const fahrt = ["90 🚗 Fahrzeit"];
  const termine = [
    { id: "1", titel: "🚗 Jetzt losfahren - Podologie Bernhardt ", start: `${d}T10:00`, ende: `${d}T10:20`, kategorien: fahrt, ort, generiert: false },
    { id: "2", titel: "Podologie Bernhardt ", start: `${d}T10:20`, ende: `${d}T11:20`, kategorien: [], ort, generiert: false },
    { id: "3", titel: "Rückfahrt - Podologie Bernhardt ", start: `${d}T11:20`, ende: `${d}T11:40`, kategorien: fahrt, ort, generiert: false },
    { id: "4", titel: "Rate überweisen", start: `${d}T18:00`, ende: `${d}T18:15`, kategorien: [], ort: "", generiert: false },
  ];
  const markiert = markiereAlteBloecke(termine, einstellungen);
  assert.deepEqual(markiert.filter((t) => t.uebernommen).map((t) => t.id), ["1", "3"]);
  assert.deepEqual(markiert.filter((t) => t.fahrzeitMerken).map((t) => [t.id, t.fahrzeit]), [["2", 20]]);
  const { anlegen, loeschen } = vergleiche(planeTag(d, markiert, einstellungen), markiert);
  assert.equal(anlegen.length, 0);
  assert.equal(loeschen.length, 0);
});

test("Fahrzeit bewusst auf 0 gesetzt: keine Fahrten, auch wenn noch alte Blöcke da sind", () => {
  const d = "2026-10-26";
  const termine = [
    { id: "1", titel: "🚗 Jetzt losfahren - Arzt", start: `${d}T10:00`, ende: `${d}T10:20`, kategorien: ["90 🚗 Fahrzeit"], ort: "", generiert: true },
    { id: "2", titel: "Arzt", start: `${d}T10:20`, ende: `${d}T11:00`, kategorien: [], ort: "", generiert: false, fahrzeit: 0 },
  ];
  const markiert = markiereAlteBloecke(termine, einstellungen);
  const { anlegen, loeschen } = vergleiche(planeTag(d, markiert, einstellungen), markiert);
  assert.equal(anlegen.length, 0);
  assert.deepEqual(loeschen.map((t) => t.id), ["1"]);
});

test("ohneFahrzeit: nur private Termine ohne Fahrzeit und ohne eigene Anfahrt", async () => {
  const { ohneFahrzeit } = await import("../src/planer.js");
  const d = "2026-11-03";
  const t = (id, titel, extra = {}) => ({ id, titel, start: `${d}T10:00`, ende: `${d}T11:00`, kategorien: [], ort: "x", generiert: false, fahrzeit: null, ...extra });
  const termine = [
    t("1", "Zahnarzt"),
    t("2", "Arzt", { fahrzeit: 20 }),
    t("3", "Podologie"),
    t("4", "🚗 Jetzt losfahren - Podologie", { kategorien: ["90 🚗 Fahrzeit"] }),
    t("5", "RehaSport (Michael)", { kategorien: ["10 SVE: RehaSport"] }),
    t("6", "Friseur", { fahrzeit: 0 }),
  ];
  assert.deepEqual(ohneFahrzeit(termine, einstellungen).map((x) => x.id), ["1"]);
});

test("ohneFahrzeit: gleicher Titel an einem anderen Tag zählt nicht als Anfahrt", async () => {
  const { ohneFahrzeit } = await import("../src/planer.js");
  const termine = [
    { id: "1", titel: "🚗 Jetzt losfahren - Fortbildung", start: "2026-10-16T11:30", ende: "2026-10-16T13:00", kategorien: ["90 🚗 Fahrzeit"], ort: "x", generiert: true },
    { id: "2", titel: "Fortbildung", start: "2026-10-16T13:00", ende: "2026-10-16T16:00", kategorien: [], ort: "x", generiert: false, fahrzeit: 90 },
    { id: "3", titel: "Fortbildung", start: "2026-10-17T08:00", ende: "2026-10-17T14:30", kategorien: [], ort: "x", generiert: false, fahrzeit: null },
  ];
  assert.deepEqual(ohneFahrzeit(termine, einstellungen).map((x) => x.id), ["3"]);
});

test("Lehrgangstag: viele private Einheiten am Stück sind eine Fahrt", () => {
  const d = "2026-10-02";
  const p = (titel, von, bis) => ({ titel, start: `${d}T${von}`, ende: `${d}T${bis}`, kategorien: [], ort: "Albstadt", generiert: false, fahrzeit: 105 });
  const termine = [p("Eröffnung", "13:45", "14:30"), p("World Cafe", "14:15", "15:00"), p("Abendessen", "18:00", "19:00"), p("Spiele", "19:45", "21:00")];
  assert.deepEqual(kurz(planeTag(d, termine, einstellungen)), ["12:00-13:45 🚗 Jetzt losfahren - Eröffnung", "21:00-22:45 Rückfahrt - Spiele"]);
});

test("Mehrtägiger Termin: Anfahrt am ersten, Rückfahrt am letzten Tag", () => {
  const event = { titel: "Event", start: "2026-09-24T12:00", ende: "2026-09-26T10:00", kategorien: [], ort: "Künzelsau", generiert: false, fahrzeit: 195 };
  assert.deepEqual(kurz(planeTag("2026-09-24", [event], einstellungen)), ["08:45-12:00 🚗 Jetzt losfahren - Event"]);
  assert.deepEqual(kurz(planeTag("2026-09-25", [event], einstellungen)), []);
  assert.deepEqual(kurz(planeTag("2026-09-26", [event], einstellungen)), ["10:00-13:15 Rückfahrt - Event"]);
});

test("Übernachtung: Fortbildung über zwei Tage hat nur eine Anfahrt und eine Rückfahrt", () => {
  const p = (d, titel, von, bis) => ({ titel, start: `${d}T${von}`, ende: `${d}T${bis}`, kategorien: ["Übernachtung"], ort: "Albstadt", generiert: false, fahrzeit: 90 });
  const alle = [p("2026-10-16", "Eröffnung", "13:45", "14:15"), p("2026-10-16", "Praxis", "14:15", "20:30"), p("2026-10-17", "Faszien", "08:45", "13:00"), p("2026-10-17", "Abschluss", "14:00", "14:30")];
  const plane = (d) => kurz(planeTag(d, termineFuerTag(alle, d, einstellungen), einstellungen));
  assert.deepEqual(plane("2026-10-16"), ["12:15-13:45 🚗 Jetzt losfahren - Eröffnung"]);
  assert.deepEqual(plane("2026-10-17"), ["14:30-16:00 Rückfahrt - Abschluss"]);
  // Ohne die Kategorie fährt er heim: An- und Rückfahrt an beiden Tagen.
  const ohne = alle.map((t) => ({ ...t, kategorien: [] }));
  assert.equal(planeTag("2026-10-16", termineFuerTag(ohne, "2026-10-16", einstellungen), einstellungen).length, 2);
});

test("Fahrt-Blöcke haben einen Google-Maps-Link zum Ziel", () => {
  const d = "2026-10-20";
  const e = { ...einstellungen, heimadresse: "Musterweg 1, 88000 Musterstadt" };
  const privat = { titel: "Fortbildung", start: `${d}T10:00`, ende: `${d}T12:00`, kategorien: [], ort: "Halle", adresse: "Vogelsangstraße 21, 72461 Albstadt", generiert: false, fahrzeit: 60 };
  const [hin, rueck] = planeTag(d, [privat], e);
  assert.match(hin.html, /destination=Vogelsangstra%C3%9Fe%2021%2C%2072461%20Albstadt&amp;travelmode=driving">🧭 Route in Google Maps starten</);
  assert.match(rueck.html, /destination=Musterweg%201/);
  // Ohne Heimadresse (im Add-in) keine Rückfahrt-Link.
  assert.equal(planeTag(d, [privat], einstellungen)[1].html, undefined);
});

test("Nachbereitung erinnert bei Kursende ans Protokoll", () => {
  const d = "2026-10-21";
  const kurs = { titel: "RehaSport (Michael)", start: `${d}T08:45`, ende: `${d}T09:30`, kategorien: ["10 SVE: RehaSport"], generiert: false };
  const soll = planeTag(d, [kurs], einstellungen);
  const nach = soll.find((b) => b.titel.startsWith("Nachbereitung"));
  assert.equal(nach.erinnerung, 0);
  assert.equal(soll.find((b) => b.titel.startsWith("Vorbereitung")).erinnerung, undefined);
});

test("Kursfahrten: tatsächliche Strecke steht in der Notiz", () => {
  const d = "2026-10-21";
  const kurs = { titel: "RehaSport (Michael)", start: `${d}T08:45`, ende: `${d}T09:30`, kategorien: ["10 SVE: RehaSport"], generiert: false };
  const e = {
    ...einstellungen,
    heimadresse: "Musterweg 1, 88000 Musterstadt",
    strecken: { "heim-SVE": { km: 7.9, minuten: 9 }, "SVE-heim": { km: 8.1, minuten: 10 } },
  };
  const soll = planeTag(d, [kurs], e);
  const min = einstellungen.anbieter[0].fahrzeitMin;
  assert.match(soll[0].html, new RegExp(`Strecke laut OpenStreetMap: 7,9 km, ca\\. 9 Min\\. \\(eingeplant ${min} Min\\.\\)`));
  assert.match(soll.at(-1).html, /8,1 km, ca\. 10 Min\./);
  // Ohne Strecken (Add-in): nur der Link.
  assert.doesNotMatch(planeTag(d, [kurs], { ...e, strecken: undefined })[0].html, /Strecke/);
});

test("Vergangene Tage bekommen keine Erinnerungen", () => {
  const d = "2026-10-21";
  const kurs = { titel: "RehaSport (Michael)", start: `${d}T08:45`, ende: `${d}T09:30`, kategorien: ["10 SVE: RehaSport"], generiert: false };
  assert.ok(planeTag(d, [kurs], { ...einstellungen, heute: "2026-10-21" }).some((b) => b.erinnerung !== undefined));
  assert.ok(planeTag(d, [kurs], { ...einstellungen, heute: "2026-10-22" }).every((b) => b.erinnerung === undefined));
});
