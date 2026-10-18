// Schulferien: Kurse eines Anbieters, der in den Ferien pausiert (ferienFrei),
// bekommen automatisch das Kürzel (F). Nur reguläre Kurse ohne Kürzel; was
// Michael selbst markiert hat (V, U, A, H, F), bleibt unangetastet.
import { kursAnbieter, leseKuerzel } from "./planer.js";

// ferien: [{ von: "JJJJ-MM-TT", bis: "JJJJ-MM-TT" }] (bis einschließlich)
export const inFerien = (datum, ferien) => ferien.some((f) => f.von <= datum && datum <= f.bis);

export function ferienKurse(termine, ferien, einstellungen) {
  return termine.filter((t) => {
    const a = !t.generiert && kursAnbieter(t, einstellungen);
    return a && a.ferienFrei && leseKuerzel(t.titel) === "" && inFerien(t.start.slice(0, 10), ferien);
  });
}
