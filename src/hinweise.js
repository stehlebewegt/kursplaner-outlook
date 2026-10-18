// Tägliche Hinweise und Wochenvorschau, als Entwurf in Michaels Postfach.
// Reine Logik ohne Outlook, damit sie sich testen lässt.
import { kursAnbieter, leseKuerzel, ohneKuerzel } from "./planer.js";
import { leseProtokoll } from "./protokoll.js";

const datumDe = (d) => `${d.slice(8, 10)}.${d.slice(5, 7)}.`;
const zeit = (t) => `${t.start.slice(11, 16)}–${t.ende.slice(11, 16)}`;
const WOCHENTAGE = ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"];
const wochentag = (d) => WOCHENTAGE[new Date(`${d}T12:00:00Z`).getUTCDay()];
// Kurs, der wirklich stattfindet (Anbieter-Kategorie, Kürzel aktiv).
const istKurs = (t, e) => !t.generiert && kursAnbieter(t, e) && e.kuerzel[leseKuerzel(t.titel)]?.aktiv;
export const plusTage = (d, n) => new Date(Date.parse(`${d}T12:00:00Z`) + n * 86400000).toISOString().slice(0, 10);

// Kurse, die schon vorbei sind und noch keine Anwesenden im Protokoll haben.
// termine brauchen dafür ihren Text (t.text).
export function fehlendeProtokolle(termine, jetzt, einstellungen) {
  return termine
    .filter((t) => istKurs(t, einstellungen) && t.ende <= jetzt)
    .filter((t) => leseProtokoll(t.text || "").anwesend === "");
}

// Private Termine (beschäftigt), die sich mit einem Kurs oder einem Fahrt-/
// Puffer-Block überschneiden. Liefert Paare [privat, anderer].
export function ueberschneidungen(termine, einstellungen) {
  const privat = termine.filter((t) => !t.generiert && !kursAnbieter(t, einstellungen) && t.showAs !== "free");
  const andere = termine.filter((t) => t.generiert || istKurs(t, einstellungen));
  const paare = [];
  for (const p of privat)
    for (const a of andere) if (p.start < a.ende && a.start < p.ende) paare.push([p, a]);
  return paare;
}

export function hinweisText({ protokolle, paare }) {
  const teile = [];
  if (protokolle.length)
    teile.push(
      "Kursprotokoll fehlt noch (Anwesende nicht eingetragen):\n" +
        protokolle.map((t) => `- ${wochentag(t.start.slice(0, 10))} ${datumDe(t.start)} ${zeit(t)}  ${t.titel.trim()}`).join("\n")
    );
  if (paare.length)
    teile.push(
      "Überschneidungen mit Kursen oder Fahrten:\n" +
        paare
          .map(([p, a]) => `- ${wochentag(p.start.slice(0, 10))} ${datumDe(p.start)}: ${p.titel.trim()} (${zeit(p)}) mit ${a.titel.trim()} (${zeit(a)})`)
          .join("\n")
    );
  return teile.join("\n\n");
}

// Überblick über eine Woche (montag = "JJJJ-MM-TT").
export function wochenText(termine, montag, einstellungen) {
  const fahrt = einstellungen.kategorieFahrt;
  const zeilen = [];
  let kurseSumme = 0;
  let fahrMin = 0;
  for (let i = 0; i < 7; i++) {
    const d = plusTage(montag, i);
    const amTag = termine.filter((t) => t.start.slice(0, 10) === d).sort((a, b) => (a.start < b.start ? -1 : 1));
    const kurse = amTag.filter((t) => !t.generiert && kursAnbieter(t, einstellungen));
    const privat = amTag.filter((t) => !t.generiert && !kursAnbieter(t, einstellungen) && t.showAs !== "free");
    const fahrten = amTag.filter((t) => t.generiert && (t.kategorien || []).includes(fahrt));
    const min = fahrten.reduce((s, t) => s + (Date.parse(t.ende) - Date.parse(t.start)) / 60000, 0);
    fahrMin += min;
    zeilen.push(`${wochentag(d)} ${datumDe(d)}`);
    if (!kurse.length && !privat.length) {
      zeilen.push("  frei");
      continue;
    }
    for (const k of kurse) {
      const kz = leseKuerzel(k.titel);
      const bed = einstellungen.kuerzel[kz];
      if (bed?.aktiv) kurseSumme++;
      const zusatz = kz && kz !== "" ? ` [${bed?.bedeutung || kz}]` : "";
      zeilen.push(`  ${zeit(k)}  ${ohneKuerzel(k.titel)} (${kursAnbieter(k, einstellungen).kuerzel})${zusatz}`);
    }
    for (const p of privat) zeilen.push(`  ${zeit(p)}  ${p.titel.trim()}`);
    if (fahrten.length) zeilen.push(`  Fahrzeit: ${Math.round(min)} Min. (${fahrten.length} Fahrten)`);
  }
  const kopf = `Diese Woche: ${kurseSumme} Kurse, ${Math.round((fahrMin / 60) * 10) / 10} Std. Fahrzeit.`;
  return `${kopf}\n\n${zeilen.join("\n")}`;
}
