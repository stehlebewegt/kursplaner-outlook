// Strecken der Kursfahrten (zu Hause ↔ Anbieter, Anbieter ↔ Anbieter) über
// OpenStreetMap, für die Notiz in den Fahrt-Blöcken. Die Werte verraten etwas
// über die Heimadresse, darum nie ausgeben oder ins Repo schreiben, nur in
// Michaels Kalender.
import { aufRunden } from "./route.js";
import { STRECKEN_BESCHRIFTUNG, kursAnbieter, istExtern, streckenSchluessel } from "../src/planer.js";

// termine: Vertretungen in anderen Vereinen darunter bekommen ihre eigenen
// Strecken (zu Hause und zu den festen Anbietern), Schlüssel "VTR@<Adresse>".
export function streckenPaare(e, termine = []) {
  const paare = [];
  for (const a of e.anbieter) paare.push([`heim-${a.kuerzel}`, null, a.ort], [`${a.kuerzel}-heim`, a.ort, null]);
  for (const a of e.anbieter) for (const b of e.anbieter) if (a !== b) paare.push([`${a.kuerzel}-${b.kuerzel}`, a.ort, b.ort]);
  const extern = new Map();
  for (const t of termine.filter((t) => !t.generiert && istExtern(t, e))) {
    const x = kursAnbieter(t, e);
    if (x.adresse) extern.set(streckenSchluessel(x), x.adresse);
  }
  for (const [k, ort] of extern) {
    paare.push([`heim-${k}`, null, ort], [`${k}-heim`, ort, null]);
    for (const a of e.anbieter) paare.push([`${a.kuerzel}-${k}`, a.ort, ort], [`${k}-${a.kuerzel}`, ort, a.ort]);
  }
  return paare;
}

export async function kursStrecken(rechne, e, termine = []) {
  const strecken = {};
  for (const [schluessel, von, nach] of streckenPaare(e, termine)) {
    const s = await rechne.strecke(von, nach).catch(() => null);
    if (s) strecken[schluessel] = s;
  }
  return strecken;
}

// Eingestellte Fahrzeiten, die nicht zur aufgerundeten Strecke passen
// (Rundung einstellungen.fahrzeitRundungMin, nur diese Minuten werden ausgegeben).
export function abweichungen(strecken, e) {
  const soll = (k) => (strecken[k] ? aufRunden(strecken[k].minuten, e.fahrzeitRundungMin ?? 15) : null);
  const liste = [];
  for (const a of e.anbieter) {
    const s = soll(`heim-${a.kuerzel}`);
    if (s != null && s !== a.fahrzeitMin) liste.push(`${a.kuerzel} ${a.fahrzeitMin}→${s}`);
  }
  for (const [k] of streckenPaare(e).filter(([k]) => !k.includes("heim"))) {
    const s = soll(k);
    const [von, nach] = k.split("-");
    const ist = e.fahrzeitZwischen[k] ?? e.anbieter.find((a) => a.kuerzel === nach).fahrzeitMin;
    if (s != null && s !== ist) liste.push(`${von}→${nach} ${ist}→${s}`);
  }
  return liste;
}

// Kursfahrten, deren Notiz noch keine Strecke hat. Fahrten zu privaten
// Terminen am Ort eines Anbieters (z. B. Besprechung) zählen nicht: der Titel
// muss einen Kurs des Tages nennen.
export function ohneStrecke(termine, e) {
  const kurse = new Set(termine.filter((t) => !t.generiert && kursAnbieter(t, e)).map((t) => `${t.start.slice(0, 10)} ${t.titel.trim()}`));
  const nenntKurs = (t) => {
    const m = t.titel.trim().match(/^(?:Anfahrt|Rückfahrt|Fahrt [A-Z]+ → [A-Z]+) - (.+)$/);
    return m && kurse.has(`${t.start.slice(0, 10)} ${m[1].trim()}`);
  };
  return termine.filter(
    (t) =>
      t.generiert &&
      t.kategorien.includes(e.kategorieFahrt) &&
      typeof t.vorschau === "string" &&
      !t.vorschau.includes(STRECKEN_BESCHRIFTUNG) &&
      nenntKurs(t)
  );
}
