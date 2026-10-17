// Die Hintergrund-Automatik: plant alle Tage eines Zeitraums neu.
// Fängt auch Änderungen vom iPhone ab (Kurs verschoben, gelöscht, Kürzel
// im Titel getippt), weil sie jeden Tag einfach neu berechnet.
import { planeTagNeu, vorschauTag } from "../src/aktionen.js";
import { leseKuerzel, findeAnbieter } from "../src/planer.js";

// api.holeBereich(von, bis) liefert alle Termine; daraus wird ein api pro Tag.
export async function planeZeitraum(api, tage, einstellungen, { schreiben }) {
  const alle = await api.holeBereich(tage[0], tage[tage.length - 1]);
  const ergebnis = [];
  for (const datum of tage) {
    const tagApi = {
      ...api,
      // Auch mehrtägige Termine, die an diesem Tag noch laufen (für die Rückfahrt am letzten Tag).
      holeTag: async () => alle.filter((t) => t.start.slice(0, 10) <= datum && t.ende.slice(0, 10) >= datum),
    };
    const termine = await tagApi.holeTag();
    // Kurse mit (U), (A), (F) im Titel, die noch "beschäftigt" sind, freigeben.
    const freigeben = termine.filter(
      (t) =>
        !t.generiert &&
        t.showAs === "busy" &&
        findeAnbieter(t, einstellungen) &&
        einstellungen.kuerzel[leseKuerzel(t.titel)]?.aktiv === false
    );
    const v = schreiben ? await planeTagNeu(tagApi, datum, einstellungen) : await vorschauTag(tagApi, datum, einstellungen);
    if (schreiben) for (const t of freigeben) await api.aendere(t.id, { showAs: "free" });
    const fahrzeit = v.fahrzeitMerken || [];
    if (v.anlegen.length || v.loeschen.length || v.uebernehmen.length || freigeben.length || fahrzeit.length) {
      ergebnis.push({ datum, anlegen: v.anlegen, loeschen: v.loeschen, uebernehmen: v.uebernehmen, freigeben, fahrzeit });
    }
  }
  return ergebnis;
}

// Heute plus die nächsten n-1 Tage als "JJJJ-MM-TT".
export function naechsteTage(heute, n) {
  const d = new Date(`${heute}T12:00:00Z`);
  return Array.from({ length: n }, (_, i) => new Date(d.getTime() + i * 86400000).toISOString().slice(0, 10));
}
