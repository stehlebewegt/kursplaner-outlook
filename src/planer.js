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
  const kurs = kursAnbieter(termin, einstellungen);
  if (kurs) return kurs;
  const ort = termin.ort || "";
  return einstellungen.anbieter.find((a) => a.ortErkennung.some((t) => ort.includes(t))) || null;
}

// Vertretung in einem anderen Verein (Kategorie aus einstellungen.vertretungExtern).
export const istExtern = (termin, e) =>
  !!e.vertretungExtern && (termin.kategorien || []).some((k) => e.vertretungExtern.kategorien.includes(k));

// Für Vertretungen in anderen Vereinen gibt es keinen festen Anbieter: Ort und
// Fahrzeit kommen aus dem Termin. schluessel unterscheidet die Orte (für Strecken).
function externAnbieter(termin, e) {
  const x = e.vertretungExtern;
  const adresse = (termin.adresse || termin.ort || "").trim();
  return {
    ...x,
    ort: termin.ort || "",
    adresse,
    fahrzeitMin: termin.fahrzeit > 0 ? termin.fahrzeit : x.fahrzeitMin,
    schluessel: `${x.kuerzel}@${adresse}`,
    extern: true,
  };
}
export const streckenSchluessel = (a) => a.schluessel || a.kuerzel;
export const gleicherAnbieter = (a, b) => a === b || (!!a && !!b && a.kuerzel === b.kuerzel && a.ort === b.ort);

// Anbieter nur über die Kategorie. Ein Termin, der bloß am Ort eines Anbieters
// liegt (z. B. "Besprechung PTM"), ist kein Kurs.
export function kursAnbieter(termin, einstellungen) {
  const kats = termin.kategorien || [];
  return (
    einstellungen.anbieter.find((a) => kats.some((k) => a.kategorien.includes(k) || k === a.kategorieVertretung)) ||
    (istExtern(termin, einstellungen) ? externAnbieter(termin, einstellungen) : null)
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
      gleicherAnbieter(vorher.anbieter, k.anbieter) &&
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

const tagDavor = (datum, n) => new Date(Date.parse(`${datum}T12:00:00Z`) + n * 86400000).toISOString().slice(0, 10);
const laeuftAm = (t, datum) => t.start.slice(0, 10) <= datum && t.ende.slice(0, 10) >= datum;
const mitUebernachtung = (t, e) => !t.generiert && (t.kategorien || []).includes(e.kategorieUebernachtung);

// Termine, die für das Planen eines Tages nötig sind: alles, was an dem Tag
// läuft, plus Übernachtungs-Termine vom Vortag und Folgetag.
export function termineFuerTag(alle, datum, einstellungen) {
  const nachbar = (t) => laeuftAm(t, tagDavor(datum, -1)) || laeuftAm(t, tagDavor(datum, 1));
  return alle.filter((t) => laeuftAm(t, datum) || (nachbar(t) && mitUebernachtung(t, einstellungen)));
}

// Inhalt für Fahrt-Blöcke: ein kurzer Link, der im Auto direkt die Navigation startet
// (auf dem iPhone in der Google-Maps-App, wenn sie installiert ist). Start ist
// der aktuelle Standort. Die Heimadresse kennt nur die Automatik (Secret),
// ohne sie bekommt die Rückfahrt keinen Link.
export const ROUTEN_LINK = "https://www.google.com/maps/dir/?api=1";
// Outlook-Felder für die Erinnerung eines Blocks: Minuten vor Beginn (0 = zum
// Beginn), leer oder -1 = keine.
export const KEINE_ERINNERUNG = -1;
export const erinnerungsFelder = (min) =>
  min == null || min < 0 ? { isReminderOn: false } : { isReminderOn: true, reminderMinutesBeforeStart: min };

// Woran ein Block mit Link erkannt wird (Outlook liefert als Vorschau nur den Text).
export const ROUTEN_BESCHRIFTUNG = "Route in Google Maps starten";
export const routenUrl = (ziel) => `${ROUTEN_LINK}&destination=${encodeURIComponent(ziel)}&travelmode=driving`;
export const routenHtml = (ziel) => `<p><a href="${routenUrl(ziel).replace(/&/g, "&amp;")}">🧭 ${ROUTEN_BESCHRIFTUNG}</a></p>`;

// Tatsächliche Strecke in der Notiz der Kursfahrten. Die Werte kennt nur die
// Automatik (sie hängen an der Heimadresse); einstellungen.strecken hat
// Schlüssel wie "heim-SVE", "SVE-heim", "SVE-PTM" mit { km, minuten }.
export const STRECKEN_BESCHRIFTUNG = "Strecke laut OpenStreetMap";
const zahl = (x) => String(x).replace(".", ",");
export const streckenHtml = (s, eingeplant) =>
  s ? `<p>📏 ${STRECKEN_BESCHRIFTUNG}: ${zahl(s.km)} km, ca. ${s.minuten} Min. (eingeplant ${eingeplant} Min.)</p>` : "";

// Liefert die Soll-Liste aller generierten Blöcke für einen Tag.
export function planeTag(datum, alleTermine, einstellungen) {
  const e = einstellungen;
  const termine = alleTermine.filter((t) => laeuftAm(t, datum));
  const nachbarn = alleTermine.filter((t) => !laeuftAm(t, datum) && mitUebernachtung(t, e));
  const uebernachtetDavor = nachbarn.some((t) => t.ende.slice(0, 10) === tagDavor(datum, -1));
  const uebernachtetDanach = nachbarn.some((t) => t.start.slice(0, 10) === tagDavor(datum, 1));
  const bloecke = bildeBloecke(aktiveKurse(termine, e), e);
  // Lohnt sich zwischen zwei Blöcken die Fahrt nach Hause?
  const heimfahrt = (von, nach, pause) =>
    pause >= Math.max(von.anbieter.fahrzeitMin + nach.anbieter.fahrzeitMin, e.pauseHeimfahrtMin ?? 0);
  const soll = [];
  const neu = (titel, start, ende, kategorie, anbieter, ziel, strecke) =>
    soll.push({
      titel,
      start: zuZeit(datum, start),
      ende: zuZeit(datum, ende),
      kategorien: [kategorie],
      ort: anbieter.ort,
      generiert: true,
      ...(ziel ? { html: routenHtml(ziel) + streckenHtml(e.strecken?.[strecke], ende - start) } : {}),
    });
  // Hinfahrten bekommen eine Erinnerung, damit das Handy rechtzeitig meldet
  // (in der Erinnerung steht der Routen-Link gleich mit drin).
  // Vergangene Tage nie mit Erinnerung, sonst meldet Outlook sie als überfällig.
  const vergangen = e.heute && datum < e.heute;
  const mitErinnerung = () => {
    if (e.erinnerungVorFahrtMin && !vergangen) soll[soll.length - 1].erinnerung = e.erinnerungVorFahrtMin;
  };
  // Nach dem Kurs: Erinnerung zum Beginn der Nachbereitung (= Kursende), damit
  // das Protokoll gleich ausgefüllt wird.
  const protokollErinnerung = () => {
    if (e.erinnerungProtokoll && !vergangen) soll[soll.length - 1].erinnerung = 0;
  };

  // Private Termine mit Fahrzeit (im Termin gespeichert): Hin- und Rückfahrt.
  // Termine kurz hintereinander (z. B. ein Lehrgangstag mit vielen Einheiten)
  // sind eine Fahrt: Anfahrt vor dem ersten, Rückfahrt nach dem letzten. Heim
  // geht es dazwischen nur, wenn sich die Pause lohnt (wie bei Kursen).
  // Mehrtägige Termine (z. B. Event von Donnerstag bis Samstag): Anfahrt nur am
  // ersten Tag, Rückfahrt nur am letzten, dazwischen keine Fahrt. Genauso bei
  // Terminen mit Kategorie "Übernachtung" an aufeinanderfolgenden Tagen.
  const beginn = (t) => (t.start.slice(0, 10) < datum ? 0 : zuMinuten(t.start));
  const schluss = (t) => (t.ende.slice(0, 10) > datum ? 24 * 60 : zuMinuten(t.ende));
  const privat = privateMitFahrzeit(termine, e).sort((a, b) => (a.start < b.start ? -1 : 1));
  const gruppen = [];
  for (const t of privat) {
    const g = gruppen[gruppen.length - 1];
    const lohnt = g && beginn(t) - g.bis >= Math.max(g.letzter.fahrzeit + t.fahrzeit, e.pauseHeimfahrtMin ?? 0);
    if (g && !lohnt) {
      if (schluss(t) >= g.bis) g.letzter = t;
      g.bis = Math.max(g.bis, schluss(t));
    } else gruppen.push({ erster: t, letzter: t, bis: schluss(t) });
  }
  for (const { erster, letzter, bis } of gruppen) {
    const von = beginn(erster);
    if (erster.start.slice(0, 10) === datum && !(uebernachtetDavor && mitUebernachtung(erster, e)))
      neu(`Anfahrt - ${erster.titel.trim()}`, von - erster.fahrzeit, von, e.kategorieFahrt, { ort: erster.ort }, erster.adresse || erster.ort) && mitErinnerung();
    if (letzter.ende.slice(0, 10) === datum && !(uebernachtetDanach && mitUebernachtung(letzter, e)))
      neu(`Rückfahrt - ${letzter.titel.trim()}`, bis, bis + letzter.fahrzeit, e.kategorieFahrt, { ort: letzter.ort }, e.heimadresse);
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
    if (direktVomVorherigen && gleicherAnbieter(vor.anbieter, b.anbieter)) {
      // Beim selben Anbieter geblieben: keine Fahrt.
    } else if (!direktVomVorherigen) {
      neu(`Anfahrt - ${erster.vollerTitel}`, vorStart - b.anbieter.fahrzeitMin, vorStart, e.kategorieFahrt, b.anbieter, b.anbieter.adresse || b.anbieter.ort, `heim-${streckenSchluessel(b.anbieter)}`) && mitErinnerung();
    } else {
      const dauer = fahrzeit(vor.anbieter, b.anbieter, e);
      neu(`Fahrt ${vor.anbieter.kuerzel} → ${b.anbieter.kuerzel} - ${erster.vollerTitel}`, vorStart - dauer, vorStart, e.kategorieFahrt, b.anbieter, b.anbieter.adresse || b.anbieter.ort, `${streckenSchluessel(vor.anbieter)}-${streckenSchluessel(b.anbieter)}`) && mitErinnerung();
    }
    neu(`Vorbereitung - ${erster.vollerTitel}`, vorStart, erster.start, e.kategoriePuffer, b.anbieter);

    for (let j = 1; j < b.kurse.length; j++) {
      const a = b.kurse[j - 1];
      const k = b.kurse[j];
      if (k.start > a.ende) {
        neu(`Nach-/Vorbereitung - ${k.vollerTitel}`, a.ende, k.start, e.kategoriePuffer, b.anbieter);
        protokollErinnerung();
      }
    }

    neu(`Nachbereitung - ${letzter.vollerTitel}`, letzter.ende, nachEnde, e.kategoriePuffer, b.anbieter);
    protokollErinnerung();
    b.ende = nachEnde;

    const direktZumNaechsten = nach && !heimfahrt(b, nach, nach.kurse[0].start - e.vorbereitungMin - nachEnde);
    if (!direktZumNaechsten) {
      neu(`Rückfahrt - ${letzter.vollerTitel}`, nachEnde, nachEnde + b.anbieter.fahrzeitMin, e.kategorieFahrt, b.anbieter, e.heimadresse, `${streckenSchluessel(b.anbieter)}-heim`);
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
      (!kursAnbieter(t, einstellungen) || istExtern(t, einstellungen)) &&
      !istBlockTitel(t.titel) &&
      !titel.has(`${t.start.slice(0, 10)} Anfahrt - ${t.titel.trim()}`)
  );
}
