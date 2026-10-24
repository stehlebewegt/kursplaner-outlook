// Kursdetails: alles zu einem Kurs aus seinen Terminen (gleicher Kurs = gleicher
// Name, gleicher Wochentag, gleiche Uhrzeit). Rechnet nur, lädt nichts.
import { leseProtokoll, gesamt } from "./protokoll.js";

// Ampel wie in Michaels Kursbuch 2026 (Stammdaten): gut ab 70 %, kritisch unter 55 %.
// In den Kursdetails wie im Reha-Cockpit: nur die letzten 8 Protokolle, ab 3.
export const AMPEL = { gut: 0.7, beobachten: 0.55, letzte: 8, ab: 3 };
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
  const juengste = gezaehlt.slice(-AMPEL.letzte);
  const anteil = (liste, k) => liste.reduce((x, v) => x + (v[k] || 0), 0);
  const quoteAmpel = juengste.length >= AMPEL.ab && anteil(juengste, "gesamt") ? anteil(juengste, "anwesend") / anteil(juengste, "gesamt") : null;
  // Trend: Schnitt anwesend der letzten 4 Protokolle gegen die 4 davor (±10 %).
  const schnitt = (liste) => anteil(liste, "anwesend") / liste.length;
  const neu = gezaehlt.slice(-4);
  const davor = gezaehlt.slice(-8, -4);
  let trend = null;
  if (neu.length === 4 && davor.length >= 2) {
    const n = schnitt(neu);
    const d = schnitt(davor);
    trend = { jetzt: n, vorher: d, richtung: n > d * 1.1 ? "steigt" : n < d * 0.9 ? "sinkt" : "gleich" };
  }
  const themen = new Map();
  for (const v of gehalten) if (v.p.thema) themen.set(v.p.thema, (themen.get(v.p.thema) || 0) + 1);
  return {
    gehalten: gehalten.length,
    mitProtokoll: gezaehlt.length,
    schnittAnwesend: gezaehlt.length ? summe("anwesend") / gezaehlt.length : null,
    schnittGesamt: gezaehlt.length ? summe("gesamt") / gezaehlt.length : null,
    quote,
    quoteAmpel,
    ampel: ampel(quoteAmpel),
    trend,
    ausfaelle: vergangen.filter((v) => v.ausfall).map((v) => v.t).reverse(),
    // Balken: die letzten 8 Termine, ausgefallene als leerer Platz.
    balken: vergangen.slice(-8),
    verlauf: gehalten.filter((v) => v.p.thema || v.anwesend != null).reverse().slice(0, 5),
    themen: [...themen].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], "de")).slice(0, 8),
    kommend,
    dieser: mit(kurs),
  };
}
