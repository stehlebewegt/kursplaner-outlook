// Kern von "Tag neu planen": Aus den Kursen eines Tages werden Anfahrt,
// Vorbereitung, Puffer, Nachbereitung und Rückfahrt komplett neu berechnet.
// Reine Logik ohne Outlook, damit sie sich testen lässt.
//
// Termin-Format (Ein- und Ausgabe):
//   { titel, start: "2026-10-14T08:45", ende: "2026-10-14T09:30",
//     kategorien: [...], ort, generiert: true|false, id? }
// "generiert" markiert Blöcke, die das Add-in selbst angelegt hat. Nur diese
// werden beim Neuplanen verändert oder gelöscht.

const KUERZEL_MUSTER = /^\s*\(([A-Z])\)\s*/;

export function leseKuerzel(titel) {
  const m = titel.match(KUERZEL_MUSTER);
  return m ? m[1] : "";
}

export function ohneKuerzel(titel) {
  return titel.replace(KUERZEL_MUSTER, "").trim();
}

function zuMinuten(zeit) {
  const [h, m] = zeit.slice(11, 16).split(":").map(Number);
  return h * 60 + m;
}

function zuZeit(datum, minuten) {
  const h = String(Math.floor(minuten / 60)).padStart(2, "0");
  const m = String(minuten % 60).padStart(2, "0");
  return `${datum}T${h}:${m}`;
}

export function findeAnbieter(termin, einstellungen) {
  const kats = termin.kategorien || [];
  for (const a of einstellungen.anbieter) {
    if (kats.some((k) => a.kategorien.includes(k) || k === a.kategorieVertretung)) return a;
  }
  const ort = termin.ort || "";
  return einstellungen.anbieter.find((a) => a.ortErkennung.some((t) => ort.includes(t))) || null;
}

// Anbieter nur über die Kategorie. Ein Termin, der bloß am Ort eines Anbieters
// liegt (z. B. "Besprechung PTM"), ist kein Kurs.
export function kursAnbieter(termin, einstellungen) {
  const kats = termin.kategorien || [];
  return (
    einstellungen.anbieter.find((a) => kats.some((k) => a.kategorien.includes(k) || k === a.kategorieVertretung)) || null
  );
}

// Kurse, die tatsächlich stattfinden: Anbieter-Kategorie, Kürzel aktiv, nicht generiert.
export function aktiveKurse(termine, einstellungen) {
  return termine
    .filter((t) => !t.generiert)
    .map((t) => ({ termin: t, anbieter: kursAnbieter(t, einstellungen) }))
    .filter(({ termin, anbieter }) => {
      if (!anbieter) return false;
      const k = einstellungen.kuerzel[leseKuerzel(termin.titel)];
      return k ? k.aktiv : false;
    })
    .map(({ termin, anbieter }) => ({
      titel: ohneKuerzel(termin.titel),
      vollerTitel: termin.titel.trim(),
      start: zuMinuten(termin.start),
      ende: zuMinuten(termin.ende),
      anbieter,
    }))
    .sort((a, b) => a.start - b.start);
}

// Kurse am Stück (gleicher Anbieter, kleine Lücke) zu Blöcken zusammenfassen.
export function bildeBloecke(kurse, einstellungen) {
  const bloecke = [];
  for (const k of kurse) {
    const letzter = bloecke[bloecke.length - 1];
    const vorher = letzter && letzter.kurse[letzter.kurse.length - 1];
    if (
      vorher &&
      vorher.anbieter === k.anbieter &&
      k.start - vorher.ende <= einstellungen.maxLueckeAmStueckMin
    ) {
      letzter.kurse.push(k);
    } else {
      bloecke.push({ anbieter: k.anbieter, kurse: [k] });
    }
  }
  return bloecke;
}

function fahrzeit(von, nach, einstellungen) {
  const schluessel = `${von.kuerzel}-${nach.kuerzel}`;
  return einstellungen.fahrzeitZwischen[schluessel] ?? nach.fahrzeitMin;
}

// Liefert die Soll-Liste aller generierten Blöcke für einen Tag.
export function planeTag(datum, termine, einstellungen) {
  const e = einstellungen;
  const bloecke = bildeBloecke(aktiveKurse(termine, e), e);
  // Lohnt sich zwischen zwei Blöcken die Fahrt nach Hause?
  const heimfahrt = (von, nach, pause) =>
    pause >= Math.max(von.anbieter.fahrzeitMin + nach.anbieter.fahrzeitMin, e.pauseHeimfahrtMin ?? 0);
  const soll = [];
  const neu = (titel, start, ende, kategorie, anbieter) =>
    soll.push({
      titel,
      start: zuZeit(datum, start),
      ende: zuZeit(datum, ende),
      kategorien: [kategorie],
      ort: anbieter.ort,
      generiert: true,
    });

  // Private Termine mit Fahrzeit (im Termin gespeichert): Hin- und Rückfahrt.
  // Termine kurz hintereinander (z. B. ein Lehrgangstag mit vielen Einheiten)
  // sind eine Fahrt: Anfahrt vor dem ersten, Rückfahrt nach dem letzten. Heim
  // geht es dazwischen nur, wenn sich die Pause lohnt (wie bei Kursen).
  const privat = privateMitFahrzeit(termine, e).sort((a, b) => (a.start < b.start ? -1 : 1));
  const gruppen = [];
  for (const t of privat) {
    const g = gruppen[gruppen.length - 1];
    const lohnt = g && zuMinuten(t.start) - g.bis >= Math.max(g.letzter.fahrzeit + t.fahrzeit, e.pauseHeimfahrtMin ?? 0);
    if (g && !lohnt) {
      if (zuMinuten(t.ende) >= g.bis) g.letzter = t;
      g.bis = Math.max(g.bis, zuMinuten(t.ende));
    } else gruppen.push({ erster: t, letzter: t, bis: zuMinuten(t.ende) });
  }
  for (const { erster, letzter, bis } of gruppen) {
    const von = zuMinuten(erster.start);
    neu(`Anfahrt - ${erster.titel.trim()}`, von - erster.fahrzeit, von, e.kategorieFahrt, { ort: erster.ort });
    neu(`Rückfahrt - ${letzter.titel.trim()}`, bis, bis + letzter.fahrzeit, e.kategorieFahrt, { ort: letzter.ort });
  }

  bloecke.forEach((b, i) => {
    const erster = b.kurse[0];
    const letzter = b.kurse[b.kurse.length - 1];
    const vor = bloecke[i - 1];
    const nach = bloecke[i + 1];
    const vorStart = erster.start - e.vorbereitungMin;
    const nachEnde = letzter.ende + e.nachbereitungMin;

    // Anfahrt: von zu Hause, oder direkt vom vorherigen Anbieter, wenn
    // die Zeit dazwischen für Heimfahrt und erneute Anfahrt nicht reicht.
    const direktVomVorherigen = vor && !heimfahrt(vor, b, vorStart - vor.ende);
    if (direktVomVorherigen && vor.anbieter === b.anbieter) {
      // Beim selben Anbieter geblieben: keine Fahrt.
    } else if (!direktVomVorherigen) {
      neu(`Anfahrt - ${erster.vollerTitel}`, vorStart - b.anbieter.fahrzeitMin, vorStart, e.kategorieFahrt, b.anbieter);
    } else {
      const dauer = fahrzeit(vor.anbieter, b.anbieter, e);
      neu(`Fahrt ${vor.anbieter.kuerzel} → ${b.anbieter.kuerzel} - ${erster.vollerTitel}`, vorStart - dauer, vorStart, e.kategorieFahrt, b.anbieter);
    }
    neu(`Vorbereitung - ${erster.vollerTitel}`, vorStart, erster.start, e.kategoriePuffer, b.anbieter);

    for (let j = 1; j < b.kurse.length; j++) {
      const a = b.kurse[j - 1];
      const k = b.kurse[j];
      if (k.start > a.ende) {
        neu(`Nach-/Vorbereitung - ${k.vollerTitel}`, a.ende, k.start, e.kategoriePuffer, b.anbieter);
      }
    }

    neu(`Nachbereitung - ${letzter.vollerTitel}`, letzter.ende, nachEnde, e.kategoriePuffer, b.anbieter);
    b.ende = nachEnde;

    const direktZumNaechsten = nach && !heimfahrt(b, nach, nach.kurse[0].start - e.vorbereitungMin - nachEnde);
    if (!direktZumNaechsten) {
      neu(`Rückfahrt - ${letzter.vollerTitel}`, nachEnde, nachEnde + b.anbieter.fahrzeitMin, e.kategorieFahrt, b.anbieter);
    }
  });
  return soll;
}

// Blöcke, die Michael vor dem Add-in von Hand angelegt hat: gleiche Titel-Muster
// und Kategorie, Ort bei einem Anbieter. Die werden beim ersten Neuplanen
// übernommen, sonst gäbe es alles doppelt. Private Fahrten (Arzt usw.) passen
// nicht, weil ihr Ort keinem Anbieter gehört.
const ALTE_BLOCK_TITEL = /^(Anfahrt|Vorbereitung|Nach-\/Vorbereitung|Nachbereitung|Rückfahrt|Fahrt [A-Z]+ → [A-Z]+) - /;

export const istBlockTitel = (titel) => ALTE_BLOCK_TITEL.test(titel.trim());

export function istAlterBlock(termin, einstellungen, kursTitel = []) {
  if (termin.generiert) return false;
  const kats = termin.kategorien || [];
  const passendeKategorie =
    kats.includes(einstellungen.kategorieFahrt) || kats.includes(einstellungen.kategoriePuffer);
  const ort = termin.ort || "";
  const beiAnbieter = einstellungen.anbieter.some((a) => a.ortErkennung.some((t) => ort.includes(t)));
  // Manche alten Rückfahrten haben die Heimadresse als Ort; die erkennen wir
  // daran, dass ihr Titel einen Kurs des Tages nennt.
  const titel = termin.titel.trim();
  const nenntKurs = kursTitel.includes(titel.replace(ALTE_BLOCK_TITEL, "").trim());
  return ALTE_BLOCK_TITEL.test(titel) && passendeKategorie && (beiAnbieter || nenntKurs);
}

// Private Termine: kein Kurs, nicht generiert, mit Fahrzeit in Minuten.
export function privateMitFahrzeit(termine, einstellungen) {
  return termine.filter((t) => !t.generiert && !kursAnbieter(t, einstellungen) && t.fahrzeit > 0);
}

const dauer = (t) => zuMinuten(t.ende) - zuMinuten(t.start);

export function markiereAlteBloecke(termine, einstellungen) {
  const e = einstellungen;
  // Privater Termin ohne gespeicherte Fahrzeit, aber mit einer von Hand
  // angelegten "Anfahrt - <Titel>": Die Dauer dieser Anfahrt wird seine
  // Fahrzeit (und im Termin gemerkt, siehe fahrzeitMerken).
  termine = termine.map((t) => {
    if (t.generiert || t.fahrzeit != null || kursAnbieter(t, e)) return t;
    const anfahrt = termine.find(
      (b) =>
        !b.generiert &&
        b.titel.trim() === `Anfahrt - ${t.titel.trim()}` &&
        (b.kategorien || []).includes(e.kategorieFahrt)
    );
    return anfahrt && dauer(anfahrt) > 0 ? { ...t, fahrzeit: dauer(anfahrt), fahrzeitMerken: true } : t;
  });
  const titel = [
    ...termine.filter((t) => kursAnbieter(t, e)),
    ...privateMitFahrzeit(termine, e),
  ].map((t) => t.titel.trim());
  return termine.map((t) => (istAlterBlock(t, e, titel) ? { ...t, generiert: true, uebernommen: true } : t));
}

// Vergleicht Soll mit den vorhandenen generierten Blöcken.
// Ergebnis: was angelegt und was gelöscht werden muss. Unveränderte bleiben stehen.
export function vergleiche(soll, termine) {
  const vorhanden = termine.filter((t) => t.generiert);
  const gleich = (a, b) => a.titel.trim() === b.titel.trim() && a.start === b.start && a.ende === b.ende;
  return {
    anlegen: soll.filter((s) => !vorhanden.some((v) => gleich(s, v))),
    loeschen: vorhanden.filter((v) => !soll.some((s) => gleich(s, v))),
  };
}

// Private Termine, für die noch niemand eine Fahrzeit festgelegt hat und für
// die es auch keine von Hand angelegte Anfahrt gibt: Kandidaten fürs Berechnen.
export function ohneFahrzeit(termine, einstellungen) {
  // Anfahrt muss am selben Tag liegen (gleicher Titel an anderen Tagen zählt nicht).
  const titel = new Set(termine.map((t) => `${t.start.slice(0, 10)} ${t.titel.trim()}`));
  return termine.filter(
    (t) =>
      !t.generiert &&
      t.fahrzeit == null &&
      !kursAnbieter(t, einstellungen) &&
      !istBlockTitel(t.titel) &&
      !titel.has(`${t.start.slice(0, 10)} Anfahrt - ${t.titel.trim()}`)
  );
}
