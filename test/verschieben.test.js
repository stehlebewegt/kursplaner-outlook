// Kurs verschieben: Kurs (und Anbieter-Kopie) auf die neue Zeit, beide Tage neu geplant.
import { test } from "node:test";
import assert from "node:assert/strict";
import { verschiebeKurs } from "../src/aktionen.js";
import { einstellungen } from "../src/einstellungen.js";

const ORT = "SVE Sport- & Gesundheitscenter, Ettenkircher Straße 17, 88048 Friedrichshafen";
function kalender(start) {
  let n = 1;
  const termine = start.map((t) => ({ id: String(n++), generiert: false, ...t }));
  return {
    termine,
    verschoben: [],
    async holeTag(d) {
      return termine.filter((t) => t.start.startsWith(d)).map((t) => ({ ...t }));
    },
    async legeAn(t, { generiert } = {}) {
      const neu = { ...t, id: String(n++), generiert: !!generiert };
      termine.push(neu);
      return neu;
    },
    async loesche(id) {
      termine.splice(termine.findIndex((t) => t.id === id), 1);
    },
    async aendere() {},
    async markiere() {},
    async verschiebe(id, start, ende) {
      this.verschoben.push(id);
      const t = termine.find((x) => x.id === id);
      if (t) Object.assign(t, { start, ende });
    },
  };
}

test("Kurs auf einen anderen Tag: alte Blöcke weg, neue am neuen Tag, Anbieter-Kopie wandert mit", async () => {
  const api = kalender([{ titel: "RehaSport (Michael)", start: "2026-11-04T08:45", ende: "2026-11-04T09:30", kategorien: ["10 SVE: RehaSport"], ort: ORT }]);
  const kurs = api.termine[0];
  // Erst den Tag planen, damit es alte Blöcke gibt.
  await verschiebeKurs(api, kurs, { datum: "2026-11-04", beginn: "08:45", ende: "09:30" }, einstellungen);
  assert.equal(api.termine.filter((t) => t.generiert && t.start.startsWith("2026-11-04")).length, 4);
  const v = await verschiebeKurs(api, { ...kurs }, { datum: "2026-11-05", beginn: "10:00", ende: "10:45" }, einstellungen, { anbieterKopie: { id: "99" } });
  assert.deepEqual(v.tage, ["2026-11-04", "2026-11-05"]);
  assert.deepEqual(api.verschoben.slice(-2), ["1", "99"]);
  assert.equal(api.termine.filter((t) => t.generiert && t.start.startsWith("2026-11-04")).length, 0);
  const neu = api.termine.filter((t) => t.generiert && t.start.startsWith("2026-11-05")).map((t) => `${t.start.slice(11)} ${t.titel}`).sort();
  assert.deepEqual(neu, [
    "09:25 🚗 Jetzt losfahren - RehaSport (Michael)",
    "09:45 Vorbereitung - RehaSport (Michael)",
    "10:45 Nachbereitung - RehaSport (Michael)",
    "11:00 Rückfahrt - RehaSport (Michael)",
  ]);
});
