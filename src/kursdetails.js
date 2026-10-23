// Kursdetails: alles zu einem Kurs aus seinen Terminen (gleicher Kurs = gleicher
// Name, gleicher Wochentag, gleiche Uhrzeit). Rechnet nur, lädt nichts.
import { leseProtokoll, gesamt } from "./protokoll.js";

// Ampel wie in Michaels Kursauswertung (Excel): ab 90 % gut, ab 65 % beobachten.
export const AMPEL = { gut: 0.9, beobachten: 0.65 };
export const ampel = (quote) => (quote == null ? "" : quote >= AMPEL.gut ? "gut" : quote >= AMPEL.beobachten ? "beobachten" : "kritisch");

const wochentag = (t) => new Date(`${t.start.slice(0, 10)}T12:00:00Z`).getUTCDay();
export const gleicherKurs = (a, b, ohneKuerzel) =>
  ohneKuerzel(a.titel).trim() === ohneKuerzel(b.titel).trim() && wochentag(a) === wochentag(b) && a.start.slice(11, 16) === b.start.slice(11, 16);

const zahl = (v) => (String(v ?? "").trim() === "" ? null : parseInt(v, 10) || 0);

export function kursDetails(kurs, termine, { jetzt, ohneKuerzel, faelltAus }) {
  const serie = termine.filter((t) => gleicherKurs(t, kurs, ohneKuerzel)).sort((a, b) => (a.start < b.start ? -1 : 1));
  const mit = (t) => {
    const p = leseProtokoll(t.text || "");
    return { t, p, anwesend: zahl(p.anwesend), gesamt: zahl(gesamt(p)), ausfall: faelltAus(t) };
  };
  const vergangen = serie.filter((t) => t.ende <= jetzt).map(mit);
  const kommend = serie.filter((t) => t.start > kurs.start).slice(0, 4).map(mit);
  const gehalten = vergangen.filter((v) => !v.ausfall);
  const gezaehlt = gehalten.filter((v) => v.anwesend != null);
  const summe = (k) => gezaehlt.reduce((s, v) => s + (v[k] || 0), 0);
  const quote = summe("gesamt") ? summe("anwesend") / summe("gesamt") : null;
  const themen = new Map();
  for (const v of gehalten) if (v.p.thema) themen.set(v.p.thema, (themen.get(v.p.thema) || 0) + 1);
  return {
    gehalten: gehalten.length,
    mitProtokoll: gezaehlt.length,
    schnittAnwesend: gezaehlt.length ? summe("anwesend") / gezaehlt.length : null,
    schnittGesamt: gezaehlt.length ? summe("gesamt") / gezaehlt.length : null,
    quote,
    ampel: ampel(quote),
    // Balken: die letzten 8 Termine, ausgefallene als leerer Platz.
    balken: vergangen.slice(-8),
    verlauf: gehalten.filter((v) => v.p.thema || v.anwesend != null).reverse().slice(0, 5),
    themen: [...themen].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], "de")).slice(0, 8),
    kommend,
    dieser: mit(kurs),
  };
}
