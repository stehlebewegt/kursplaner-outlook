// Die Button-Aktionen gegen einen Kalender im Speicher.
import { test } from "node:test";
import assert from "node:assert/strict";
import { uebernehmeKurs, setzeKuerzel, planeTagNeu, vorschauTag } from "../src/aktionen.js";
import { einstellungen } from "../src/einstellungen.js";

function speicherKalender(start = []) {
  let naechsteId = 1;
  const termine = start.map((t) => ({ id: String(naechsteId++), generiert: false, ...t }));
  return {
    termine,
    async holeTag(datum) {
      return termine.filter((t) => t.start.startsWith(datum)).map((t) => ({ ...t }));
    },
    async legeAn(t, { generiert } = {}) {
      const neu = { ...t, id: String(naechsteId++), generiert: !!generiert };
      termine.push(neu);
      return neu;
    },
    async loesche(id) {
      termine.splice(termine.findIndex((t) => t.id === id), 1);
    },
    async aendere(id, felder) {
      const t = termine.find((x) => x.id === id);
      if (felder.subject) t.titel = felder.subject;
      if (felder.showAs) t.showAs = felder.showAs;
    },
    async markiere(id) {
      termine.find((x) => x.id === id).generiert = true;
    },
  };
}

const ORT = "SVE Sport- & Gesundheitscenter\nEttenkircher Straße 17, 88048 Friedrichshafen, Deutschland";
const titel = (api) => api.termine.map((t) => `${t.start.slice(11)} ${t.titel}`).sort();

test("Vertretung übernehmen: Kurs mit (V) und Vertretungs-Kategorie, Blöcke drumherum", async () => {
  const api = speicherKalender();
  const quelle = { titel: "RehaSport (Eva)", start: "2026-10-13T08:45", ende: "2026-10-13T09:30", kategorien: ["10 SVE: RehaSport"], ort: ORT };
  const { kurs } = await uebernehmeKurs(api, quelle, { vertretung: true }, einstellungen);
  assert.equal(kurs.titel, "(V) RehaSport (Eva)");
  assert.deepEqual(kurs.kategorien, ["19 Vertretung Kurs SVE", "10 SVE: RehaSport"]);
  assert.deepEqual(titel(api), [
    "08:10 Anfahrt - (V) RehaSport (Eva)",
    "08:30 Vorbereitung - (V) RehaSport (Eva)",
    "08:45 (V) RehaSport (Eva)",
    "09:30 Nachbereitung - (V) RehaSport (Eva)",
    "09:45 Rückfahrt - (V) RehaSport (Eva)",
  ]);
});

test("Vertretung direkt an eigenen Kurs: Fahrten werden zusammengelegt", async () => {
  const api = speicherKalender([
    { titel: "RehaSport (Michael)", start: "2026-10-28T09:45", ende: "2026-10-28T10:30", kategorien: ["10 SVE: RehaSport"], ort: ORT },
  ]);
  await planeTagNeu(api, "2026-10-28", einstellungen);
  const quelle = { titel: "RehaSport (Eva)", start: "2026-10-28T10:45", ende: "2026-10-28T11:30", kategorien: ["10 SVE: RehaSport"], ort: ORT };
  await uebernehmeKurs(api, quelle, { vertretung: true }, einstellungen);
  assert.deepEqual(titel(api), [
    "09:10 Anfahrt - RehaSport (Michael)",
    "09:30 Vorbereitung - RehaSport (Michael)",
    "09:45 RehaSport (Michael)",
    "10:30 Nach-/Vorbereitung - (V) RehaSport (Eva)",
    "10:45 (V) RehaSport (Eva)",
    "11:30 Nachbereitung - (V) RehaSport (Eva)",
    "11:45 Rückfahrt - (V) RehaSport (Eva)",
  ]);
});

test("Doppelt übernehmen legt nichts zweimal an", async () => {
  const api = speicherKalender();
  const quelle = { titel: "RehaSport (Eva)", start: "2026-10-13T08:45", ende: "2026-10-13T09:30", kategorien: ["10 SVE: RehaSport"], ort: ORT };
  await uebernehmeKurs(api, quelle, { vertretung: true }, einstellungen);
  const vorher = api.termine.length;
  const { doppelt } = await uebernehmeKurs(api, quelle, { vertretung: true }, einstellungen);
  assert.ok(doppelt);
  assert.equal(api.termine.length, vorher);
});

test("Urlaub setzen: (U) davor, Termin frei, Blöcke weg; Kürzel entfernen holt sie zurück", async () => {
  const api = speicherKalender([
    { titel: "RehaSport (Michael)", start: "2026-10-21T08:45", ende: "2026-10-21T09:30", kategorien: ["10 SVE: RehaSport"], ort: ORT },
  ]);
  await planeTagNeu(api, "2026-10-21", einstellungen);
  const kurs = api.termine[0];
  await setzeKuerzel(api, kurs, "U", einstellungen);
  assert.deepEqual(titel(api), ["08:45 (U) RehaSport (Michael)"]);
  assert.equal(kurs.showAs, "free");
  await setzeKuerzel(api, kurs, "", einstellungen);
  assert.equal(api.termine.length, 5);
});

test("Vorhandene Fahrt-Blöcke ohne Routen-Link werden erkannt", async () => {
  const d = "2026-10-20";
  const privat = { id: "p", titel: "Fortbildung", start: `${d}T10:00`, ende: `${d}T12:00`, kategorien: [], ort: "Albstadt", generiert: false, fahrzeit: 60 };
  const block = { id: "b", titel: "Anfahrt - Fortbildung", start: `${d}T09:00`, ende: `${d}T10:00`, kategorien: [einstellungen.kategorieFahrt], ort: "Albstadt", generiert: true, vorschau: "" };
  const api = { holeTag: async () => [privat, block] };
  const v = await vorschauTag(api, d, einstellungen);
  assert.deepEqual(v.nachbessern.map((l) => l.id), ["b"]);
  assert.deepEqual(v.anlegen.map((t) => t.titel), ["Rückfahrt - Fortbildung"]);
});

test("Anfahrt bekommt eine Erinnerung, vorhandene Blöcke werden nachgebessert", async () => {
  const d = "2026-10-20";
  const privat = { id: "p", titel: "Fortbildung", start: `${d}T10:00`, ende: `${d}T12:00`, kategorien: [], ort: "Albstadt", generiert: false, fahrzeit: 60 };
  const hin = { id: "h", titel: "Anfahrt - Fortbildung", start: `${d}T09:00`, ende: `${d}T10:00`, kategorien: [einstellungen.kategorieFahrt], ort: "Albstadt", generiert: true, vorschau: "🧭 Route in Google Maps starten", erinnerung: -1 };
  const rueck = { id: "r", titel: "Rückfahrt - Fortbildung", start: `${d}T12:00`, ende: `${d}T13:00`, kategorien: [einstellungen.kategorieFahrt], ort: "Albstadt", generiert: true, erinnerung: -1 };
  const v = await vorschauTag({ holeTag: async () => [privat, hin, rueck] }, d, einstellungen);
  assert.deepEqual(v.nachbessern, [{ id: "h", felder: { isReminderOn: true, reminderMinutesBeforeStart: 10 } }]);
});

test("Urlaub im Zeitraum: nur reguläre Kurse bekommen (U), Vertretungen bleiben", async () => {
  const { kurseImZeitraum } = await import("../src/aktionen.js");
  const k = (id, titel, kat = "10 SVE: RehaSport") => ({ id, titel, start: "2026-11-03T15:00", ende: "2026-11-03T15:45", kategorien: [kat], generiert: false });
  const termine = [k("a", "RehaSport (Michael)"), k("b", "(V) RehaSport (Eva)", "19 Vertretung Kurs SVE"), k("c", "(A) Yoga (Michael)"), k("d", "Zahnarzt", "privat")];
  const u = kurseImZeitraum(termine, "U", einstellungen);
  assert.deepEqual(u.aendern.map((t) => t.id), ["a"]);
  assert.deepEqual(u.bleiben.map((t) => t.id), ["b"]);
  assert.deepEqual(kurseImZeitraum(termine, "", einstellungen).aendern.map((t) => t.id), ["c"]);
});
