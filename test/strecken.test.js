import { test } from "node:test";
import assert from "node:assert/strict";
import { abweichungen, ohneStrecke, streckenPaare } from "../automatik/strecken.js";
import { einstellungen } from "../src/einstellungen.js";

test("Paare: heim hin/zurück je Anbieter plus Anbieter untereinander", () => {
  assert.deepEqual(streckenPaare(einstellungen).map((p) => p[0]), ["heim-SVE", "SVE-heim", "heim-PTM", "PTM-heim", "SVE-PTM", "PTM-SVE"]);
});

test("Abweichung: aufgerundet auf Viertelstunde gegen eingestellt", () => {
  const e = { ...einstellungen, anbieter: einstellungen.anbieter.map((a) => ({ ...a, fahrzeitMin: 15 })), fahrzeitZwischen: {}, fahrzeitRundungMin: 15 };
  assert.deepEqual(abweichungen({ "heim-SVE": { km: 9, minuten: 8 }, "heim-PTM": { km: 20, minuten: 16 }, "SVE-PTM": { km: 15, minuten: 14 } }, e), ["PTM 15→30"]);
});

test("Nur Kursfahrten ohne Strecke zählen", () => {
  const d = "2026-11-06T";
  const f = (titel, vorschau) => ({ titel, start: `${d}10:00`, generiert: true, kategorien: [einstellungen.kategorieFahrt], ort: "Physio-Team", vorschau });
  const kurs = { titel: "PUR (Michael)", start: `${d}11:00`, generiert: false, kategorien: ["01 PTM: PUR Kurs"], ort: "Physio-Team" };
  const t = [kurs, f("🚗 Jetzt losfahren - PUR (Michael)", "🧭 Route"), f("Rückfahrt - PUR (Michael)", "📏 Strecke laut OpenStreetMap: 1 km"), f("🚗 Jetzt losfahren - Besprechung PTM", "🧭 Route")];
  assert.equal(ohneStrecke(t, einstellungen).length, 1);
});

test("Rundung auf 5 Minuten", () => {
  const e = { ...einstellungen, anbieter: einstellungen.anbieter.map((a) => ({ ...a, fahrzeitMin: 15 })), fahrzeitZwischen: {}, fahrzeitRundungMin: 5 };
  assert.deepEqual(abweichungen({ "heim-SVE": { km: 9, minuten: 17 }, "heim-PTM": { km: 9, minuten: 12 } }, e), ["SVE 15→20"]);
});
