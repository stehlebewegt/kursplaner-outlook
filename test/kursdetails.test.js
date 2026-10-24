import { test } from "node:test";
import assert from "node:assert/strict";
import { kursDetails, ampel } from "../src/kursdetails.js";

const ohneKuerzel = (t) => t.replace(/^\([A-Z]\)\s*/, "");
const t = (d, titel = "RehaSport / PUR (Michael)", text = "") => ({ titel, start: `${d}T15:00`, ende: `${d}T15:45`, text });
const prot = (thema, a, e, u) => `Thema: ${thema}\nAnwesend: ${a}\nEntschuldigt: ${e}\nUnentschuldigt: ${u}\n`;

test("Kursdetails rechnen nur den gleichen Kurs (Name, Wochentag, Uhrzeit)", () => {
  const termine = [
    t("2026-10-05", undefined, prot("Kraft", 8, 2, 0)),
    t("2026-10-12", undefined, prot("Gleichgewicht", 6, 2, 2)),
    t("2026-10-19", "(U) RehaSport / PUR (Michael)"),
    t("2026-10-26"),
    t("2026-11-02"),
    t("2026-10-13", undefined, prot("Fremd", 1, 0, 0)), // Dienstag: anderer Kurs
    { ...t("2026-10-12"), start: "2026-10-12T16:00", ende: "2026-10-12T16:45" }, // andere Uhrzeit
  ];
  const d = kursDetails(termine[3], termine, {
    jetzt: "2026-10-23T18:00",
    ohneKuerzel,
    faelltAus: (x) => x.titel.startsWith("(U)"),
  });
  assert.equal(d.gehalten, 2);
  assert.equal(d.mitProtokoll, 2);
  assert.equal(d.schnittAnwesend, 7);
  assert.equal(d.quote, 14 / 20);
  assert.equal(d.ampel, ""); // erst ab 3 Protokollen
  assert.deepEqual(d.ausfaelle.map((x) => x.start.slice(0, 10)), ["2026-10-19"]);
  assert.equal(d.balken.length, 3);
  assert.deepEqual(d.verlauf.map((v) => v.p.thema), ["Gleichgewicht", "Kraft"]);
  assert.deepEqual(d.kommend.map((v) => v.t.start.slice(0, 10)), ["2026-11-02"]);
});

test("Ampel-Grenzen wie in der Kursauswertung", () => {
  assert.equal(ampel(0.95), "gut");
  assert.equal(ampel(0.65), "beobachten");
  assert.equal(ampel(0.5), "kritisch");
  assert.equal(ampel(null), "");
});

test("Ampel aus den letzten 8 Protokollen, Trend aus 4 gegen 4", () => {
  const tage = Array.from({ length: 10 }, (_, i) => new Date(Date.UTC(2026, 7, 3 + 7 * i)).toISOString().slice(0, 10));
  // erst 10 von 10 da, die letzten vier nur 5 von 10
  const termine = tage.map((d, i) => t(d, undefined, prot("X", i < 6 ? 10 : 5, i < 6 ? 0 : 5, 0)));
  const d = kursDetails(termine[9], termine, { jetzt: "2026-12-01T00:00", ohneKuerzel, faelltAus: () => false });
  assert.equal(d.quoteAmpel, (4 * 10 + 4 * 5) / 80); // nur die letzten 8
  assert.equal(d.ampel, "gut"); // 75 %
  assert.equal(d.trend.richtung, "sinkt");
  assert.equal(d.trend.jetzt, 5);
  assert.equal(d.trend.vorher, 10);
});
