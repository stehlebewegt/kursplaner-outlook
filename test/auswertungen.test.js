import { test } from "node:test";
import assert from "node:assert/strict";
import { einstellungen } from "../src/einstellungen.js";
import { kursbericht, wieOftWo, geld, pruefliste } from "../src/auswertungen.js";

const e = structuredClone(einstellungen);
e.anbieter = e.anbieter.map((a) => ({ ...a, satzJeKurs: a.kuerzel === "SVE" ? 25 : 20, grenzeJahr: a.kuerzel === "SVE" ? 3300 : 0, grenzeMonat: a.kuerzel === "PTM" ? 600 : 0 }));
e.ausbezahlt = { "2026-10-SVE": true };
const sve = e.anbieter.find((a) => a.kuerzel === "SVE");
const ptm = e.anbieter.find((a) => a.kuerzel === "PTM");
const t = (titel, d, von, kat, text = "", ort = "") => ({ titel, start: `${d}T${von}`, ende: `${d}T${von.slice(0, 2)}:45`, kategorien: [kat], ort, text });
const P = (a, en, u, th = "Kraft", m = "Matte") => `Thema: ${th}\nMaterial: ${m}\nAnwesend: ${a}\nEntschuldigt: ${en}\nUnentschuldigt: ${u}\n`;
const termine = [
  t("RehaSport (Michael)", "2026-10-07", "08:00", sve.kategorien[0], P(9, 1, 0)),
  t("RehaSport (Michael)", "2026-10-14", "08:00", sve.kategorien[0], P(5, 3, 2, "", "")),
  t("(U) RehaSport (Michael)", "2026-10-21", "08:00", sve.kategorien[0]),
  t("(V) RehaSport (Eva)", "2026-10-14", "10:00", sve.kategorieVertretung),
  t("RehaSport / PUR (Michael)", "2026-10-12", "15:00", ptm.kategorien[0], P(8, 2, 0, "Gleichgewicht", "Matte, Stab")),
  t("RehaSport / PUR (Michael)", "2026-10-12", "15:00", ptm.kategorien[0]),
  t("RehaSport (Michael)", "2026-10-28", "08:00", sve.kategorien[0]),
  t("Anfahrt - RehaSport (Michael)", "2026-10-07", "07:00", "90 🚗 Fahrzeit"),
];
const jetzt = "2026-10-23T12:00";

test("Kursbericht: Anwesenheit, Ampel, nur eigene", () => {
  const k = kursbericht(termine, { von: "2026-10-01", bis: "2026-10-31", jetzt }, e);
  assert.equal(k.gehalten, 5);
  assert.equal(k.mitProtokoll, 3);
  const mi = k.kurse.find((x) => x.anbieter === "SVE" && x.beginn === "08:00");
  assert.equal(mi.quote, 14 / 20);
  assert.equal(mi.ampel, "beobachten");
  assert.deepEqual(k.themen, [["Gleichgewicht", 1], ["Kraft", 1]]);
  assert.equal(kursbericht(termine, { von: "2026-10-01", bis: "2026-10-31", jetzt, nurEigene: true, anbieter: "SVE" }, e).gehalten, 2);
});

test("Wie oft wo: nur vergangene, nach Art", () => {
  const w = wieOftWo(termine, 2026, jetzt, e);
  const okt = w.monate[9].je.SVE;
  assert.equal(okt.Regulär, 2);
  assert.equal(okt.Vertretung, 1);
  assert.equal(okt.Urlaub, 1);
  assert.equal(okt.gehalten, 3);
  assert.equal(w.monate[9].gehalten, 5);
});

test("Geld: Betrag aus Satz, Urlaub zählt nicht, geplant dabei, ausbezahlt", () => {
  const g = geld(termine, 2026, jetzt, e, (k, m) => (k === "SVE" && m === 10 ? 52 : null));
  const s = g.find((x) => x.kuerzel === "SVE");
  assert.equal(s.monate[9].kurse, 4);
  assert.equal(s.monate[9].geplant, 1);
  assert.equal(s.monate[9].betrag, 100);
  assert.equal(s.monate[9].fahrtkosten, 52 * 0.3);
  assert.equal(s.monate[9].ausbezahlt, true);
  assert.equal(s.anteilJahr, 100 / 3300);
});

test("Prüfliste: ohne Teilnehmer, ohne Thema, doppelt", () => {
  const p = Object.fromEntries(pruefliste(termine, 2026, jetzt, e).punkte.map((x) => [x.id, x.termine.length]));
  assert.equal(p.teilnehmer, 2); // Vertretung 14.10. und das doppelte PTM 12.10.
  assert.equal(p.thema, 1);
  assert.equal(p.doppelt, 2);
});
