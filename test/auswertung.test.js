import { test } from "node:test";
import assert from "node:assert/strict";
import { erstelleAuswertung, monatsZeitraum, fahrtenTabelle } from "../src/auswertung.js";
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
    { ...kurs("🚗 Jetzt losfahren - RehaSport (Michael)", "2026-10-07T08:10", "2026-10-07T08:30", ["90 🚗 Fahrzeit"]), generiert: true },
  ];
  const { zeilen, summe, fahrten } = erstelleAuswertung(termine, SVE, "2026-10-16T10:00", einstellungen);
  assert.equal(zeilen.length, 5);
  assert.deepEqual(fahrten, [
    { datum: "2026-10-07", uhrzeit: "08:10", fahrt: "Hinfahrt", von: "heim", nach: "SVE", kurs: "RehaSport (Michael)", minuten: 20, status: "geleistet", km: null, minutenBerechnet: null },
  ]);
  assert.deepEqual(summe, {
    geleistet: 3, geleistetStunden: 2.25, davonVertretung: 1,
    geplant: 1, geplantStunden: 0.75, ferien: 1,
    teilnehmer: 19, teilnehmerSchnitt: 9.5, ohneProtokoll: 1,
    fahrten: 1, fahrtKm: null,
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
    fahrt("🚗 Jetzt losfahren - RehaSport (Michael)", "2026-10-08T08:10", "2026-10-08T08:30"), // anderer Tag
    kurs("(A) RehaSport (Michael)", "2026-10-09T08:45", "2026-10-09T09:30", ["10 SVE: RehaSport"]),
    fahrt("🚗 Jetzt losfahren - (A) RehaSport (Michael)", "2026-10-09T08:10", "2026-10-09T08:30"),
    fahrt("🚗 Jetzt losfahren - Zahnarzt", "2026-10-07T13:00", "2026-10-07T13:15"),
    // Titel nennt einen Kurs, der nicht im Kalender steht, Ort ist aber der Anbieter:
    { ...fahrt("🚗 Jetzt losfahren - Faszien (Michael)", "2026-10-07T08:00", "2026-10-07T08:20"), ort: "SVE Sport- & Gesundheitscenter, Ettenkircher Straße 17" },
    { ...fahrt("🚗 Jetzt losfahren - RehaSport (Michael)", "2026-10-07T07:00", "2026-10-07T07:20"), kategorien: ["90 🚗 Fahrzeit", "Fahrtkosten nicht abrechenbar"] },
  ];
  const { fahrten } = erstelleAuswertung(termine, SVE, "2026-10-16T10:00", einstellungen);
  assert.deepEqual(fahrten.map((f) => `${f.datum} ${f.fahrt}`), ["2026-10-07 Hinfahrt", "2026-10-07 Rückfahrt"]);
});

test("Fahrten-Tabelle: eine Zeile pro Tag, km und Minuten aus dem Routenplaner", () => {
  const PTM = einstellungen.anbieter.find((a) => a.kuerzel === "PTM");
  const fahrt = (titel, start, ende) => ({ ...kurs(titel, start, ende, ["90 🚗 Fahrzeit"]), generiert: true });
  const termine = [
    kurs("(H) RehaSport (Dushko)", "2026-10-07T15:00", "2026-10-07T15:45", ["10 SVE: RehaSport"]),
    fahrt("Fahrt PTM → SVE - (H) RehaSport (Dushko)", "2026-10-07T14:25", "2026-10-07T14:45"),
    fahrt("Rückfahrt - (H) RehaSport (Dushko)", "2026-10-07T16:00", "2026-10-07T16:20"),
  ];
  const strecke = (von, nach) => ({ "PTM SVE": { km: 12.34, minuten: 17 }, "SVE heim": { km: 5.1, minuten: 9 } })[`${von} ${nach}`];
  const a = erstelleAuswertung(termine, SVE, "2026-10-16T10:00", einstellungen, strecke);
  assert.equal(a.zeilen[0].art, "Hospitation");
  assert.equal(a.summe.fahrtKm, 17.4);
  assert.deepEqual(fahrtenTabelle(a.fahrten), {
    kopf: ["Datum", "Strecke", "Fahrten", "km", "Min."],
    reihen: [["2026-10-07", "PTM – SVE – Zuhause", 2, "17,4", 26]],
  });
  assert.equal(erstelleAuswertung(termine, PTM, "2026-10-16T10:00", einstellungen).fahrten.length, 0);
});

test("Nachtrag Vormonat gegen gemerkten Stand", async () => {
  const { nachtrag, stand } = await import("../src/auswertung.js");
  const z = (datum, von, kurs, status, art = "regulär") => ({ datum, von, bis: "10:00", kurs, art, status, minuten: 45 });
  const alt = stand({ zeilen: [z("2026-09-10", "09:00", "PUR", "geleistet"), z("2026-09-20", "09:00", "PUR", "geplant"), z("2026-09-22", "09:00", "FLEX", "geplant")] }, "2026-09-15T05:00");
  const jetzt = { zeilen: [z("2026-09-10", "09:00", "PUR", "geleistet"), z("2026-09-20", "09:00", "PUR", "geleistet"), z("2026-09-25", "09:00", "PUR", "geleistet", "Vertretung")], fahrten: [] };
  const n = nachtrag(jetzt, alt, "2026-09-15");
  assert.deepEqual(n.zeilen.map((x) => [x.datum, x.kurs, x.status, x.aenderung]), [
    ["2026-09-20", "PUR", "geleistet", ""],
    ["2026-09-22", "FLEX", "entfallen", "entfallen"],
    ["2026-09-25", "PUR", "geleistet", "neu"],
  ]);
  assert.equal(n.summe.geleistet, 2);
  assert.equal(nachtrag(jetzt, null, "2026-09-15").zeilen.length, 2);
});
