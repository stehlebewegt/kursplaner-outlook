import { test } from "node:test";
import assert from "node:assert/strict";
import { ferienKurse } from "../src/ferien.js";
import { einstellungen } from "../src/einstellungen.js";

test("Ferien: nur reguläre PTM-Kurse in den Ferien bekommen (F)", () => {
  const k = (titel, d, kat) => ({ titel, start: `${d}T15:00`, ende: `${d}T15:45`, kategorien: [kat], generiert: false });
  const ferien = [{ von: "2026-10-26", bis: "2026-10-30" }];
  const termine = [
    k("RehaSport / PUR (Michael)", "2026-10-26", "01 PTM: PUR Kurs"),
    k("(V) RehaSport / FLEX (Eva)", "2026-10-27", "09 Vertretung Kurs PTM"),
    k("(U) RehaSport / PUR (Michael)", "2026-10-28", "01 PTM: PUR Kurs"),
    k("RehaSport (Michael)", "2026-10-28", "10 SVE: RehaSport"),
    k("RehaSport / PUR (Michael)", "2026-11-02", "01 PTM: PUR Kurs"),
  ];
  assert.deepEqual(ferienKurse(termine, ferien, einstellungen).map((t) => t.start), ["2026-10-26T15:00"]);
});

test("Feiertage: reguläre Kurse aller Anbieter mit feiertagFrei bekommen (F)", () => {
  const k = (titel, d, kat) => ({ titel, start: `${d}T15:00`, ende: `${d}T15:45`, kategorien: [kat], generiert: false });
  const feiertage = [{ von: "2026-12-25", bis: "2026-12-25" }];
  const termine = [
    k("RehaSport (Michael)", "2026-12-25", "10 SVE: RehaSport"),
    k("(V) RehaSport (Eva)", "2026-12-25", "19 Vertretung Kurs SVE"),
    k("RehaSport (Michael)", "2026-12-23", "10 SVE: RehaSport"),
  ];
  assert.deepEqual(ferienKurse(termine, [], einstellungen, feiertage).map((t) => t.start), ["2026-12-25T15:00"]);
});
