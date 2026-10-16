// Monatsauswertung pro Anbieter: welche Kurse geleistet, welche noch geplant,
// regulär oder Vertretung, mit Teilnehmern aus dem Kursprotokoll.
// (U) und (A) zählen nicht, (F) wird als "entfällt (Ferien)" nur aufgeführt.
import { leseKuerzel, ohneKuerzel, kursAnbieter } from "./planer.js";
import { leseProtokoll } from "./protokoll.js";

const minuten = (t) => (Date.parse(t.ende) - Date.parse(t.start)) / 60000;

export function erstelleAuswertung(termine, anbieter, jetzt, einstellungen) {
  const zeilen = termine
    .filter((t) => !t.generiert && kursAnbieter(t, einstellungen) === anbieter)
    .filter((t) => !["U", "A"].includes(leseKuerzel(t.titel)))
    .map((t) => {
      const k = leseKuerzel(t.titel);
      const p = leseProtokoll(t.text);
      return {
        datum: t.start.slice(0, 10),
        von: t.start.slice(11, 16),
        bis: t.ende.slice(11, 16),
        kurs: ohneKuerzel(t.titel).trim(),
        art: k === "V" ? "Vertretung" : "regulär",
        status: k === "F" ? "entfällt (Ferien)" : t.start <= jetzt ? "geleistet" : "geplant",
        minuten: minuten(t),
        anwesend: p.anwesend === "" ? null : parseInt(p.anwesend, 10) || 0,
        thema: p.thema,
      };
    })
    .sort((a, b) => (a.datum + a.von < b.datum + b.von ? -1 : 1));

  const zaehle = (f) => zeilen.filter(f);
  const geleistet = zaehle((z) => z.status === "geleistet");
  const geplant = zaehle((z) => z.status === "geplant");
  const mitTeilnehmern = geleistet.filter((z) => z.anwesend !== null);
  const summe = {
    geleistet: geleistet.length,
    geleistetStunden: geleistet.reduce((s, z) => s + z.minuten, 0) / 60,
    davonVertretung: geleistet.filter((z) => z.art === "Vertretung").length,
    geplant: geplant.length,
    geplantStunden: geplant.reduce((s, z) => s + z.minuten, 0) / 60,
    ferien: zaehle((z) => z.status === "entfällt (Ferien)").length,
    teilnehmer: mitTeilnehmern.reduce((s, z) => s + z.anwesend, 0),
    teilnehmerSchnitt: mitTeilnehmern.length
      ? Math.round((mitTeilnehmern.reduce((s, z) => s + z.anwesend, 0) / mitTeilnehmern.length) * 10) / 10
      : null,
    ohneProtokoll: geleistet.length - mitTeilnehmern.length,
  };
  return { zeilen, summe };
}

export const MONATE = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"];

// "2026-10" → { von: "2026-10-01", bis: "2026-10-31", name: "Oktober 2026" }
export function monatsZeitraum(monat) {
  const [j, m] = monat.split("-").map(Number);
  const letzter = new Date(Date.UTC(j, m, 0)).getUTCDate();
  return { von: `${monat}-01`, bis: `${monat}-${String(letzter).padStart(2, "0")}`, name: `${MONATE[m - 1]} ${j}` };
}
