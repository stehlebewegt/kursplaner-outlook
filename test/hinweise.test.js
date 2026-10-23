import { test } from "node:test";
import assert from "node:assert/strict";
import { fehlendeProtokolle, ueberschneidungen, hinweisHtml, wochenDaten, wochenHtml } from "../src/hinweise.js";
import { einstellungen } from "../src/einstellungen.js";

const kurs = (titel, start, ende, text = "") => ({ id: titel, titel, start, ende, kategorien: ["10 SVE: RehaSport"], generiert: false, text });

test("Fehlende Protokolle: nur vergangene aktive Kurse ohne Anwesende", () => {
  const termine = [
    kurs("RehaSport (Michael)", "2026-10-16T08:45", "2026-10-16T09:30"),
    kurs("RehaSport 2 (Michael)", "2026-10-16T10:00", "2026-10-16T10:45", "Thema: Rücken\nAnwesend: 9"),
    kurs("(A) RehaSport 3 (Michael)", "2026-10-16T11:00", "2026-10-16T11:45"),
    kurs("RehaSport 4 (Michael)", "2026-10-18T19:00", "2026-10-18T19:45"),
  ];
  assert.deepEqual(fehlendeProtokolle(termine, "2026-10-18T06:00", einstellungen).map((t) => t.id), ["RehaSport (Michael)"]);
});

test("Überschneidung: privater Termin mitten in Kurs oder Fahrt", () => {
  const d = "2026-10-20";
  const termine = [
    kurs("RehaSport (Michael)", `${d}T08:45`, `${d}T09:30`),
    { titel: "🚗 Jetzt losfahren - RehaSport (Michael)", start: `${d}T08:10`, ende: `${d}T08:30`, kategorien: [einstellungen.kategorieFahrt], generiert: true },
    { titel: "Zahnarzt", start: `${d}T08:00`, ende: `${d}T08:20`, kategorien: [], generiert: false, showAs: "busy" },
    { titel: "Mittag", start: `${d}T12:00`, ende: `${d}T13:00`, kategorien: [], generiert: false, showAs: "busy" },
    { titel: "Geburtstag Oma", start: `${d}T09:00`, ende: `${d}T10:00`, kategorien: [], generiert: false, showAs: "free" },
  ];
  const paare = ueberschneidungen(termine, einstellungen);
  assert.deepEqual(paare.map(([p, a]) => `${p.titel}|${a.titel}`), ["Zahnarzt|🚗 Jetzt losfahren - RehaSport (Michael)"]);
  const html = hinweisHtml({ protokolle: [], paare });
  assert.match(html, /1 Hinweis aus dem Kursplaner/);
  assert.match(html, /<b>Zahnarzt<\/b> 08:00–08:20<br><span[^>]*>mit 🚗 Jetzt losfahren - RehaSport \(Michael\) 08:10–08:30/);
});

test("Wochenvorschau listet Kurse, Termine und Fahrzeit", () => {
  const termine = [
    kurs("RehaSport (Michael)", "2026-10-19T08:45", "2026-10-19T09:30"),
    kurs("(U) RehaSport (Michael)", "2026-10-20T08:45", "2026-10-20T09:30"),
    { titel: "🚗 Jetzt losfahren - RehaSport (Michael)", start: "2026-10-19T08:10", ende: "2026-10-19T08:30", kategorien: [einstellungen.kategorieFahrt], generiert: true },
  ];
  const w = wochenDaten(termine, "2026-10-19", einstellungen);
  assert.equal(w.kurse, 1);
  assert.equal(w.fahrStd, 0.3);
  assert.deepEqual(w.tage[1].kurse.map((k) => [k.aktiv, k.zusatz]), [[false, "Urlaub"]]);
  const html = wochenHtml(termine, "2026-10-19", einstellungen);
  assert.match(html, /Deine Woche ab Mo 19\.10\./);
  assert.match(html, /1 Kurse · 0,3 Std\. Fahrzeit/);
  assert.match(html, /<s>RehaSport \(Michael\)<\/s>/);
  assert.match(html, /🚗 20 Min\. Fahrzeit \(1 Fahrt\)/);
});

test("Wochenvorschau als PDF entsteht", async () => {
  const { wochenPdf } = await import("../src/wochenPdf.js");
  const { ladeJsPdf } = await import("../automatik/monatsbericht.js");
  const { einstellungen } = await import("../src/einstellungen.js");
  const k = (titel, s, e, kat) => ({ titel, start: s, ende: e, kategorien: [kat], generiert: false, showAs: "busy" });
  const doc = wochenPdf(
    [k("RehaSport (Michael)", "2026-10-26T15:00", "2026-10-26T15:45", "10 SVE: RehaSport"), k("(U) RehaSport / PUR (Michael)", "2026-10-27T15:00", "2026-10-27T15:45", "01 PTM: PUR Kurs")],
    "2026-10-26", einstellungen, ladeJsPdf()
  );
  const bytes = Buffer.from(doc.output("arraybuffer"));
  assert.equal(bytes.subarray(0, 4).toString(), "%PDF");
});
