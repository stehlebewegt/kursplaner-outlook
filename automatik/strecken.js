// Strecken der Kursfahrten (zu Hause ↔ Anbieter, Anbieter ↔ Anbieter) über
// OpenStreetMap, für die Notiz in den Fahrt-Blöcken. Die Werte verraten etwas
// über die Heimadresse, darum nie ausgeben oder ins Repo schreiben, nur in
// Michaels Kalender.
import { aufViertelstunde } from "./route.js";
import { STRECKEN_BESCHRIFTUNG } from "../src/planer.js";

export function streckenPaare(e) {
  const paare = [];
  for (const a of e.anbieter) paare.push([`heim-${a.kuerzel}`, null, a.ort], [`${a.kuerzel}-heim`, a.ort, null]);
  for (const a of e.anbieter) for (const b of e.anbieter) if (a !== b) paare.push([`${a.kuerzel}-${b.kuerzel}`, a.ort, b.ort]);
  return paare;
}

export async function kursStrecken(rechne, e) {
  const strecken = {};
  for (const [schluessel, von, nach] of streckenPaare(e)) {
    const s = await rechne.strecke(von, nach).catch(() => null);
    if (s) strecken[schluessel] = s;
  }
  return strecken;
}

// Eingestellte Fahrzeiten, die nicht zur aufgerundeten Strecke passen
// (nur Minuten auf die Viertelstunde, das ist unbedenklich für die Ausgabe).
export function abweichungen(strecken, e) {
  const soll = (k) => (strecken[k] ? aufViertelstunde(strecken[k].minuten) : null);
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

// Kursfahrten (Ort beim Anbieter), deren Notiz noch keine Strecke hat.
export function ohneStrecke(termine, e) {
  return termine.filter(
    (t) =>
      t.generiert &&
      t.kategorien.includes(e.kategorieFahrt) &&
      e.anbieter.some((a) => a.ortErkennung.some((x) => (t.ort || "").includes(x))) &&
      typeof t.vorschau === "string" &&
      !t.vorschau.includes(STRECKEN_BESCHRIFTUNG)
  );
}
