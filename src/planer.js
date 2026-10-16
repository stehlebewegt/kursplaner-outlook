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

// Kurse, die tatsächlich stattfinden: Anbieter erkannt, Kürzel aktiv, nicht generiert.
export function aktiveKurse(termine, einstellungen) {
  return termine
    .filter((t) => !t.generiert)
    .map((t) => ({ termin: t, anbieter: findeAnbieter(t, einstellungen) }))
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

  bloecke.forEach((b, i) => {
    const erster = b.kurse[0];
    const letzter = b.kurse[b.kurse.length - 1];
    const vor = bloecke[i - 1];
    const nach = bloecke[i + 1];
    const vorStart = erster.start - e.vorbereitungMin;
    const nachEnde = letzter.ende + e.nachbereitungMin;

    // Anfahrt: von zu Hause, oder direkt vom vorherigen Anbieter, wenn
    // die Zeit dazwischen für Heimfahrt und erneute Anfahrt nicht reicht.
    const direktVomVorherigen =
      vor && vorStart - (vor.ende ?? 0) < vor.anbieter.fahrzeitMin + b.anbieter.fahrzeitMin;
    if (!direktVomVorherigen) {
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

    const direktZumNaechsten =
      nach &&
      nach.kurse[0].start - e.vorbereitungMin - nachEnde < b.anbieter.fahrzeitMin + nach.anbieter.fahrzeitMin;
    if (!direktZumNaechsten) {
      neu(`Rückfahrt - ${letzter.vollerTitel}`, nachEnde, nachEnde + b.anbieter.fahrzeitMin, e.kategorieFahrt, b.anbieter);
    }
  });
  return soll;
}

// Vergleicht Soll mit den vorhandenen generierten Blöcken.
// Ergebnis: was angelegt und was gelöscht werden muss. Unveränderte bleiben stehen.
export function vergleiche(soll, termine) {
  const vorhanden = termine.filter((t) => t.generiert);
  const gleich = (a, b) => a.titel === b.titel && a.start === b.start && a.ende === b.ende;
  return {
    anlegen: soll.filter((s) => !vorhanden.some((v) => gleich(s, v))),
    loeschen: vorhanden.filter((v) => !soll.some((s) => gleich(s, v))),
  };
}
