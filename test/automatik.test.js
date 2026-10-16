// Die Hintergrund-Automatik gegen einen Kalender im Speicher.
import { test } from "node:test";
import assert from "node:assert/strict";
import { planeZeitraum, naechsteTage } from "../automatik/planen.js";
import { einstellungen } from "../src/einstellungen.js";

function speicherKalender(start) {
  let n = 1;
  const termine = start.map((t) => ({ id: String(n++), generiert: false, showAs: "busy", kategorien: [], ...t }));
  return {
    termine,
    async holeBereich() {
      return termine.map((t) => ({ ...t }));
    },
    async legeAn(t) {
      termine.push({ ...t, id: String(n++), generiert: true });
    },
    async loesche(id) {
      termine.splice(termine.findIndex((t) => t.id === id), 1);
    },
    async aendere(id, felder) {
      Object.assign(termine.find((t) => t.id === id), felder.showAs ? { showAs: felder.showAs } : {});
    },
    async markiere(id) {
      termine.find((t) => t.id === id).generiert = true;
    },
  };
}

const ORT = "SVE Sport- & Gesundheitscenter, Ettenkircher Straße 17, 88048 Friedrichshafen, Deutschland";
const kurs = (titel, start, ende) => ({ titel, start, ende, ort: ORT, kategorien: ["10 SVE: RehaSport"] });

test("naechsteTage zählt über den Monatswechsel", () => {
  assert.deepEqual(naechsteTage("2026-10-30", 4), ["2026-10-30", "2026-10-31", "2026-11-01", "2026-11-02"]);
});

test("Vorschau ändert nichts, Ausführen legt Blöcke an und ist danach stabil", async () => {
  const api = speicherKalender([kurs("RehaSport (Michael)", "2026-10-28T08:45", "2026-10-28T09:30")]);
  const tage = ["2026-10-27", "2026-10-28"];
  const vorschau = await planeZeitraum(api, tage, einstellungen, { schreiben: false });
  assert.equal(vorschau.length, 1);
  assert.equal(api.termine.length, 1);
  await planeZeitraum(api, tage, einstellungen, { schreiben: true });
  assert.equal(api.termine.length, 5);
  assert.deepEqual(await planeZeitraum(api, tage, einstellungen, { schreiben: true }), []);
});

test("Am iPhone verschobener Kurs: Blöcke ziehen mit", async () => {
  const api = speicherKalender([kurs("RehaSport (Michael)", "2026-10-28T08:45", "2026-10-28T09:30")]);
  await planeZeitraum(api, ["2026-10-28"], einstellungen, { schreiben: true });
  Object.assign(api.termine[0], { start: "2026-10-28T10:00", ende: "2026-10-28T10:45" });
  await planeZeitraum(api, ["2026-10-28"], einstellungen, { schreiben: true });
  const anfahrt = api.termine.find((t) => t.titel.startsWith("Anfahrt"));
  assert.equal(anfahrt.start, "2026-10-28T09:25");
  assert.equal(api.termine.length, 5);
});

test("Am iPhone (F) getippt: Kurs wird frei, Blöcke verschwinden", async () => {
  const api = speicherKalender([kurs("RehaSport (Michael)", "2026-10-28T08:45", "2026-10-28T09:30")]);
  await planeZeitraum(api, ["2026-10-28"], einstellungen, { schreiben: true });
  api.termine[0].titel = "(F) RehaSport (Michael)";
  await planeZeitraum(api, ["2026-10-28"], einstellungen, { schreiben: true });
  assert.equal(api.termine.length, 1);
  assert.equal(api.termine[0].showAs, "free");
});
