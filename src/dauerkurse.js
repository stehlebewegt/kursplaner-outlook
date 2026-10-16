// Dauerkurse: Kurse, die in einem Anbieter-Kalender mit "(Michael)" stehen,
// gehören auch in Michaels eigenen Kalender. Reine Logik, ohne Outlook.
//
// Gleicht ab:
//  - Kurs im Anbieter-Kalender ohne Gegenstück im eigenen Kalender → anlegen.
//  - Anbieter hat einen Serientermin verschoben (originalStart ≠ start) und
//    die eigene Kopie steht noch auf der alten Zeit → mitverschieben.
//  - Hat Michael seine Kopie selbst verschoben, bleibt sie, wie sie ist.
//  - Eigene Kurse ohne Gegenstück werden nur gezählt, nie gelöscht.
import { ohneKuerzel, kursAnbieter } from "./planer.js";

const norm = (titel) => ohneKuerzel(titel).replace(/\s+/g, " ").trim();
const tag = (zeit) => zeit.slice(0, 10);

export function istMeinKurs(titel, einstellungen) {
  const name = einstellungen.dauerkurse.kursleitung;
  return new RegExp(`\\(${name}\\)\\s*$`).test(titel);
}

// anbieterTermine: Termine aus den Anbieter-Kalendern, optional mit
// originalStart ("JJJJ-MM-TTTHH:MM", Ortszeit) bei Serienterminen.
// eigeneTermine: alle Termine im eigenen Kalender im selben Zeitraum.
export function gleicheAb(anbieterTermine, eigeneTermine, einstellungen) {
  const meine = anbieterTermine
    .filter((p) => istMeinKurs(p.titel, einstellungen) && kursAnbieter(p, einstellungen))
    .sort((a, b) => a.start.localeCompare(b.start));
  const eigene = eigeneTermine.filter((t) => !t.generiert && kursAnbieter(t, einstellungen));
  const benutzt = new Set();
  const frei = (t) => !benutzt.has(t);

  // 1. Gleicher Titel, gleiche Zeit: passt.
  const offen = [];
  for (const p of meine) {
    const e = eigene.find((t) => frei(t) && norm(t.titel) === norm(p.titel) && t.start === p.start);
    if (e) benutzt.add(e);
    else offen.push(p);
  }

  const anlegen = [];
  const verschieben = [];
  for (const p of offen) {
    const gleich = (t) => frei(t) && norm(t.titel) === norm(p.titel);
    // 2. Beim Anbieter verschoben, eigene Kopie noch auf der alten Zeit.
    const alt =
      p.originalStart && p.originalStart !== p.start ? eigene.find((t) => gleich(t) && t.start === p.originalStart) : null;
    if (alt) {
      benutzt.add(alt);
      verschieben.push({ termin: alt, start: p.start, ende: p.ende });
      continue;
    }
    // 3. Am selben Tag gibt es den Kurs schon (von Michael verschoben): so lassen.
    const selberTag = eigene.find((t) => gleich(t) && tag(t.start) === tag(p.start));
    if (selberTag) {
      benutzt.add(selberTag);
      continue;
    }
    anlegen.push(p);
  }

  const ohneGegenstueck = eigene.filter((t) => frei(t) && istMeinKurs(t.titel, einstellungen));
  return { anlegen, verschieben, ohneGegenstueck };
}

// Neuer eigener Kurs aus dem Anbieter-Termin (ohne Kürzel, wie "Ich gebe").
export function alsEigenerKurs(p) {
  return { titel: norm(p.titel), start: p.start, ende: p.ende, ort: p.ort, kategorien: p.kategorien, text: p.text };
}
