import { test } from "node:test";
import assert from "node:assert/strict";
import { fahrtenJahr, fahrtenPdf } from "../src/fahrtenbuch.js";
import { ladeJsPdf } from "../automatik/monatsbericht.js";
import { einstellungen } from "../src/einstellungen.js";

test("Fahrten im Jahr: nur gefahrene, eine Zeile pro Tag, km summiert", () => {
  const SVE = einstellungen.anbieter[0].ort;
  const t = (titel, s, e, kat, gen = false) => ({ titel, start: s, ende: e, kategorien: [kat], ort: SVE, generiert: gen });
  const termine = [
    t("RehaSport (Michael)", "2026-03-04T08:45", "2026-03-04T09:30", "10 SVE: RehaSport"),
    t("Anfahrt - RehaSport (Michael)", "2026-03-04T08:10", "2026-03-04T08:30", "90 🚗 Fahrzeit", true),
    t("Rückfahrt - RehaSport (Michael)", "2026-03-04T09:45", "2026-03-04T10:05", "90 🚗 Fahrzeit", true),
    t("(U) RehaSport (Michael)", "2026-03-11T08:45", "2026-03-11T09:30", "10 SVE: RehaSport"),
    t("RehaSport (Michael)", "2025-12-30T08:45", "2025-12-30T09:30", "10 SVE: RehaSport"),
  ];
  const f = fahrtenJahr(termine, "2026", "2026-10-19T08:00", einstellungen, () => ({ km: 12.4, minuten: 18 }));
  assert.equal(f.reihen.length, 1);
  assert.equal(f.reihen[0].strecke, "Zuhause – SVE – Zuhause");
  assert.equal(f.summe.fahrten, 2);
  assert.equal(f.summe.km, 24.8);
  assert.equal(f.monate[2].fahrten, 2);
  const pdf = Buffer.from(fahrtenPdf(f, "2026", einstellungen, ladeJsPdf(), "2026-10-19T08:00").output("arraybuffer"));
  assert.equal(pdf.subarray(0, 4).toString(), "%PDF");
});
