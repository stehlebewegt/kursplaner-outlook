// Die Aktionen hinter den Buttons. "api" ist der Kalender-Zugriff
// (in Outlook: web/graph.js, in den Tests: ein Kalender im Speicher).
import { planeTag, ROUTEN_BESCHRIFTUNG, erinnerungsFelder, vergleiche, markiereAlteBloecke, leseKuerzel, ohneKuerzel, findeAnbieter } from "./planer.js";

const datumVon = (t) => t.start.slice(0, 10);

// Berechnet die Änderungen für einen Tag, ohne etwas zu tun (Vorschau).
export async function vorschauTag(api, datum, einstellungen) {
  const termine = markiereAlteBloecke(await api.holeTag(datum), einstellungen);
  const soll = planeTag(datum, termine, einstellungen);
  const { anlegen, loeschen } = vergleiche(soll, termine);
  const uebernehmen = termine.filter((t) => t.uebernommen && !loeschen.includes(t));
  const fahrzeitMerken = termine.filter((t) => t.fahrzeitMerken);
  // Vorhandene Blöcke, denen der Routen-Link fehlt oder deren Erinnerung nicht
  // stimmt, werden nachgebessert. Nur wenn der Abruf diese Angaben mitliefert
  // (vorschau / erinnerung), sonst wissen wir es nicht.
  const nachbessern = termine
    .filter((t) => t.generiert && !loeschen.includes(t))
    .map((t) => ({ t, s: soll.find((s) => s.titel.trim() === t.titel.trim() && s.start === t.start && s.ende === t.ende) }))
    .filter(({ s }) => s)
    .map(({ t, s }) => {
      const felder = {};
      if (s.html && typeof t.vorschau === "string" && !t.vorschau.includes(ROUTEN_BESCHRIFTUNG))
        felder.body = { contentType: "html", content: s.html };
      if (t.erinnerung !== undefined && (s.erinnerung || 0) !== t.erinnerung) Object.assign(felder, erinnerungsFelder(s.erinnerung));
      return { id: t.id, felder };
    })
    .filter((n) => Object.keys(n.felder).length);
  return { anlegen, loeschen, uebernehmen, fahrzeitMerken, nachbessern };
}

export async function planeTagNeu(api, datum, einstellungen) {
  const v = await vorschauTag(api, datum, einstellungen);
  for (const t of v.loeschen) await api.loesche(t.id);
  for (const t of v.anlegen) await api.legeAn(t, { generiert: true, html: t.html });
  if (einstellungen.routenLinkNachtragen) for (const n of v.nachbessern) await api.aendere(n.id, n.felder);
  // Von Hand angelegte Blöcke, die passen, bekommen die Markierung.
  for (const t of v.uebernehmen) await api.markiere(t.id);
  for (const t of v.fahrzeitMerken) await api.setzeFahrzeit(t.id, t.fahrzeit);
  return v;
}

// "Ich gebe" / "Ich vertrete": Kurs aus dem Anbieter-Kalender übernehmen.
export async function uebernehmeKurs(api, quelle, { vertretung }, einstellungen) {
  const anbieter = findeAnbieter(quelle, einstellungen);
  if (!anbieter) throw new Error("Zu diesem Kurs ist kein Anbieter hinterlegt.");
  const titel = (vertretung ? "(V) " : "") + ohneKuerzel(quelle.titel);
  const vorhanden = await api.holeTag(datumVon(quelle));
  const doppelt = vorhanden.find(
    (t) => ohneKuerzel(t.titel) === ohneKuerzel(quelle.titel) && t.start === quelle.start && !t.generiert
  );
  if (doppelt) return { doppelt, aenderungen: null };
  const kategorien = vertretung
    ? [anbieter.kategorieVertretung, ...quelle.kategorien.filter((k) => k !== anbieter.kategorieVertretung)]
    : quelle.kategorien;
  const kurs = await api.legeAn(
    { titel, start: quelle.start, ende: quelle.ende, ort: quelle.ort || anbieter.ort, kategorien },
    { text: quelle.text }
  );
  const aenderungen = await planeTagNeu(api, datumVon(quelle), einstellungen);
  return { kurs, aenderungen };
}

// Kürzel (U), (A), (F) setzen oder entfernen ("" = regulär) und den Tag neu planen.
export async function setzeKuerzel(api, kurs, kuerzel, einstellungen) {
  const alt = leseKuerzel(kurs.titel);
  if (alt === "V" && kuerzel === "") kuerzel = "V";
  const titel = (kuerzel ? `(${kuerzel}) ` : "") + ohneKuerzel(kurs.titel);
  const aktiv = einstellungen.kuerzel[kuerzel]?.aktiv ?? true;
  await api.aendere(kurs.id, { subject: titel, showAs: aktiv ? "busy" : "free" });
  return { ...(await planeTagNeu(api, datumVon(kurs), einstellungen)), titel };
}

// Fahrzeit für einen privaten Termin setzen (0 = keine) und den Tag neu planen.
export async function setzeFahrzeit(api, termin, minuten, einstellungen) {
  await api.setzeFahrzeit(termin.id, minuten);
  return planeTagNeu(api, datumVon(termin), einstellungen);
}
