import { test } from "node:test";
import assert from "node:assert/strict";
import { fehlendeProtokolle, ueberschneidungen, hinweisText, wochenText } from "../src/hinweise.js";
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
    { titel: "Anfahrt - RehaSport (Michael)", start: `${d}T08:10`, ende: `${d}T08:30`, kategorien: [einstellungen.kategorieFahrt], generiert: true },
    { titel: "Zahnarzt", start: `${d}T08:00`, ende: `${d}T08:20`, kategorien: [], generiert: false, showAs: "busy" },
    { titel: "Mittag", start: `${d}T12:00`, ende: `${d}T13:00`, kategorien: [], generiert: false, showAs: "busy" },
    { titel: "Geburtstag Oma", start: `${d}T09:00`, ende: `${d}T10:00`, kategorien: [], generiert: false, showAs: "free" },
  ];
  const paare = ueberschneidungen(termine, einstellungen);
  assert.deepEqual(paare.map(([p, a]) => `${p.titel}|${a.titel}`), ["Zahnarzt|Anfahrt - RehaSport (Michael)"]);
  assert.match(hinweisText({ protokolle: [], paare }), /Di 20\.10\.: Zahnarzt \(08:00–08:20\) mit Anfahrt/);
});

test("Wochenvorschau listet Kurse, Termine und Fahrzeit", () => {
  const termine = [
    kurs("RehaSport (Michael)", "2026-10-19T08:45", "2026-10-19T09:30"),
    kurs("(U) RehaSport (Michael)", "2026-10-20T08:45", "2026-10-20T09:30"),
    { titel: "Anfahrt - RehaSport (Michael)", start: "2026-10-19T08:10", ende: "2026-10-19T08:30", kategorien: [einstellungen.kategorieFahrt], generiert: true },
  ];
  const text = wochenText(termine, "2026-10-19", einstellungen);
  assert.match(text, /^Diese Woche: 1 Kurse, 0.3 Std. Fahrzeit./);
  assert.match(text, /Mo 19\.10\.\n  08:45–09:30  RehaSport \(Michael\) \(SVE\)\n  Fahrzeit: 20 Min\. \(1 Fahrten\)/);
  assert.match(text, /Di 20\.10\.\n  08:45–09:30  RehaSport \(Michael\) \(SVE\) \[Urlaub\]/);
  assert.match(text, /So 25\.10\.\n  frei/);
});
