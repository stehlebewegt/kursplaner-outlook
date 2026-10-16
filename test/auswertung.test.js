import { test } from "node:test";
import assert from "node:assert/strict";
import { erstelleAuswertung, monatsZeitraum } from "../src/auswertung.js";
import { einstellungen } from "../src/einstellungen.js";

const SVE = einstellungen.anbieter.find((a) => a.kuerzel === "SVE");
const kurs = (titel, start, ende, kats, text = "") => ({ titel, start, ende, kategorien: kats, text, generiert: false });

test("Monatsauswertung SVE: geleistet, geplant, Vertretung, Ferien, U/A und andere Anbieter fallen raus", () => {
  const termine = [
    kurs("RehaSport (Michael)", "2026-10-07T08:45", "2026-10-07T09:30", ["10 SVE: RehaSport"], "Anwesend: 10\nEntschuldigt: 2"),
    kurs("(V) RehaSport (Eva)", "2026-10-13T10:45", "2026-10-13T11:30", ["19 Vertretung Kurs SVE", "10 SVE: RehaSport"], "Anwesend: 9"),
    kurs("RehaSport (Michael)", "2026-10-14T08:45", "2026-10-14T09:30", ["10 SVE: RehaSport"]),
    kurs("(U) RehaSport (Michael)", "2026-10-21T08:45", "2026-10-21T09:30", ["10 SVE: RehaSport"]),
    kurs("(F) RehaSport (Michael)", "2026-10-28T08:45", "2026-10-28T09:30", ["10 SVE: RehaSport"]),
    kurs("RehaSport (Michael)", "2026-10-30T08:45", "2026-10-30T09:30", ["10 SVE: RehaSport"]),
    kurs("RehaSport / PUR (Michael)", "2026-10-15T11:00", "2026-10-15T11:45", ["01 PTM: PUR Kurs"]),
    { ...kurs("Anfahrt - RehaSport (Michael)", "2026-10-07T08:10", "2026-10-07T08:30", ["90 🚗 Fahrzeit"]), generiert: true },
  ];
  const { zeilen, summe, fahrten } = erstelleAuswertung(termine, SVE, "2026-10-16T10:00", einstellungen);
  assert.equal(zeilen.length, 5);
  assert.deepEqual(fahrten, [
    { datum: "2026-10-07", von: "08:10", bis: "08:30", fahrt: "Hinfahrt", kurs: "RehaSport (Michael)", minuten: 20, status: "geleistet" },
  ]);
  assert.deepEqual(summe, {
    geleistet: 3, geleistetStunden: 2.25, davonVertretung: 1,
    geplant: 1, geplantStunden: 0.75, ferien: 1,
    teilnehmer: 19, teilnehmerSchnitt: 9.5, ohneProtokoll: 1,
    fahrten: 1, fahrtStunden: 20 / 60,
  });
});

test("Monatszeitraum", () => {
  assert.deepEqual(monatsZeitraum("2026-02"), { von: "2026-02-01", bis: "2026-02-28", name: "Februar 2026" });
});

test("Fahrten: nur zu Kursen des Anbieters am selben Tag, nicht zu abgesagten", () => {
  const fahrt = (titel, start, ende) => ({ ...kurs(titel, start, ende, ["90 🚗 Fahrzeit"]), generiert: true });
  const termine = [
    kurs("RehaSport (Michael)", "2026-10-07T08:45", "2026-10-07T09:30", ["10 SVE: RehaSport"]),
    fahrt("Rückfahrt - RehaSport (Michael)", "2026-10-07T09:45", "2026-10-07T10:05"),
    fahrt("Anfahrt - RehaSport (Michael)", "2026-10-08T08:10", "2026-10-08T08:30"), // anderer Tag
    kurs("(A) RehaSport (Michael)", "2026-10-09T08:45", "2026-10-09T09:30", ["10 SVE: RehaSport"]),
    fahrt("Anfahrt - (A) RehaSport (Michael)", "2026-10-09T08:10", "2026-10-09T08:30"),
    fahrt("Anfahrt - Zahnarzt", "2026-10-07T13:00", "2026-10-07T13:15"),
    // Titel nennt einen Kurs, der nicht im Kalender steht, Ort ist aber der Anbieter:
    { ...fahrt("Anfahrt - Faszien (Michael)", "2026-10-07T08:00", "2026-10-07T08:20"), ort: "SVE Sport- & Gesundheitscenter, Ettenkircher Straße 17" },
    { ...fahrt("Anfahrt - RehaSport (Michael)", "2026-10-07T07:00", "2026-10-07T07:20"), kategorien: ["90 🚗 Fahrzeit", "Fahrtkosten nicht abrechenbar"] },
  ];
  const { fahrten } = erstelleAuswertung(termine, SVE, "2026-10-16T10:00", einstellungen);
  assert.deepEqual(fahrten.map((f) => `${f.datum} ${f.fahrt}`), ["2026-10-07 Hinfahrt", "2026-10-07 Rückfahrt"]);
});
