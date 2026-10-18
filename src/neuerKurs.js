// Neuer Kurs aus dem Formular im Kursplaner: wird im Anbieter-Kalender
// angelegt ("Kursname (Kursleitung)"), einmalig oder als wöchentliche Serie.
// Steht Michael als Kursleitung drin, holt die Automatik die Termine wie bei
// allen Dauerkursen in den eigenen Kalender. Reine Logik, ohne Outlook.
import { schreibeProtokoll } from "./protokoll.js";

const TAGE = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"];
export const wochentagVon = (datum) => TAGE[new Date(`${datum}T12:00:00Z`).getUTCDay()];

// f: { anbieter, kategorie, name, leitung, datum, beginn, ende, serie, bis, ort, zeitzone }
// Gibt Fehlertexte zurück (leer = alles ok).
export function pruefeKurs(f) {
  const fehler = [];
  if (!f.anbieter) fehler.push("Anbieter fehlt.");
  if (!f.name?.trim()) fehler.push("Kursname fehlt.");
  if (!f.leitung?.trim()) fehler.push("Kursleitung fehlt.");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(f.datum || "")) fehler.push("Datum fehlt.");
  if (!/^\d{2}:\d{2}$/.test(f.beginn || "") || !/^\d{2}:\d{2}$/.test(f.ende || "")) fehler.push("Uhrzeit fehlt.");
  else if (f.ende <= f.beginn) fehler.push("Das Ende muss nach dem Beginn liegen.");
  if (f.serie && f.bis && f.bis < f.datum) fehler.push("„Bis“ liegt vor dem ersten Termin.");
  return fehler;
}

export const kursTitel = (f) => `${f.name.trim()} (${f.leitung.trim()})`;

// Termin für Graph (POST /calendars/{id}/events).
export function kursEreignis(f) {
  const ereignis = {
    subject: kursTitel(f),
    start: { dateTime: `${f.datum}T${f.beginn}:00`, timeZone: f.zeitzone },
    end: { dateTime: `${f.datum}T${f.ende}:00`, timeZone: f.zeitzone },
    location: { displayName: (f.ort || f.anbieter.ort || "").trim() },
    categories: [f.kategorie || f.anbieter.kategorien[0]],
    showAs: "busy",
    body: { contentType: "text", content: schreibeProtokoll({}) },
  };
  if (f.serie) {
    ereignis.recurrence = {
      pattern: { type: "weekly", interval: 1, daysOfWeek: [wochentagVon(f.datum)], firstDayOfWeek: "monday" },
      range: f.bis
        ? { type: "endDate", startDate: f.datum, endDate: f.bis, recurrenceTimeZone: f.zeitzone }
        : { type: "noEnd", startDate: f.datum, recurrenceTimeZone: f.zeitzone },
    };
  }
  return ereignis;
}

// Wie viele Termine hat die Serie (bis einschließlich "bis")? null = ohne Ende.
export function anzahlTermine(f) {
  if (!f.serie) return 1;
  if (!f.bis) return null;
  return Math.floor((Date.parse(`${f.bis}T12:00:00Z`) - Date.parse(`${f.datum}T12:00:00Z`)) / (7 * 86400000)) + 1;
}
