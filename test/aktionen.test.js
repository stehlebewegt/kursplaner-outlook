// Die Button-Aktionen gegen einen Kalender im Speicher.
import { test } from "node:test";
import assert from "node:assert/strict";
import { uebernehmeKurs, setzeKuerzel, planeTagNeu } from "../src/aktionen.js";
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
