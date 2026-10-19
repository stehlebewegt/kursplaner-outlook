// E-Mail-Vorlagen zu einem Kurs: Kurs, Datum, Uhrzeit und Anbieter werden
// eingesetzt, der Entwurf landet in den Entwürfen und öffnet sich. Gesendet
// wird nie automatisch.
import { ohneKuerzel } from "./planer.js";

const WOCHENTAGE = ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"];

// Platzhalter: {kurs} {wochentag} {datum} {zeit} {anbieter} {kurzdatum}
export function fuelleVorlage(vorlage, termin, anbieter) {
  const d = termin.start.slice(0, 10);
  const werte = {
    kurs: ohneKuerzel(termin.titel).replace(/\s*\([^)]*\)\s*$/, "").trim(),
    wochentag: WOCHENTAGE[new Date(`${d}T12:00:00Z`).getUTCDay()],
    datum: `${d.slice(8, 10)}.${d.slice(5, 7)}.${d.slice(0, 4)}`,
    kurzdatum: `${d.slice(8, 10)}.${d.slice(5, 7)}.`,
    zeit: `${termin.start.slice(11, 16)}–${termin.ende.slice(11, 16)} Uhr`,
    anbieter: anbieter?.name || "",
  };
  const ersetze = (text) => text.replace(/\{(\w+)\}/g, (m, k) => (k in werte ? werte[k] : m));
  return { betreff: ersetze(vorlage.betreff), text: ersetze(vorlage.text) };
}
