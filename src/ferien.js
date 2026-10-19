// Schulferien: Kurse eines Anbieters, der in den Ferien pausiert (ferienFrei),
// bekommen automatisch das Kürzel (F). Nur reguläre Kurse ohne Kürzel; was
// Michael selbst markiert hat (V, U, A, H, F), bleibt unangetastet.
import { kursAnbieter, leseKuerzel } from "./planer.js";

// ferien: [{ von: "JJJJ-MM-TT", bis: "JJJJ-MM-TT" }] (bis einschließlich)
export const inFerien = (datum, ferien) => ferien.some((f) => f.von <= datum && datum <= f.bis);

// feiertage: wie ferien; gilt für Anbieter mit feiertagFrei.
export function ferienKurse(termine, ferien, einstellungen, feiertage = []) {
  return termine.filter((t) => {
    const a = !t.generiert && kursAnbieter(t, einstellungen);
    if (!a || leseKuerzel(t.titel) !== "") return false;
    const d = t.start.slice(0, 10);
    return (a.ferienFrei && inFerien(d, ferien)) || (a.feiertagFrei && inFerien(d, feiertage));
  });
}
