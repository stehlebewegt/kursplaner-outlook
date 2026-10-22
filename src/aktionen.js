// Die Aktionen hinter den Buttons. "api" ist der Kalender-Zugriff
// (in Outlook: web/graph.js, in den Tests: ein Kalender im Speicher).
import { planeTag, ROUTEN_BESCHRIFTUNG, STRECKEN_BESCHRIFTUNG, erinnerungsFelder, KEINE_ERINNERUNG, vergleiche, markiereAlteBloecke, leseKuerzel, ohneKuerzel, findeAnbieter, kursAnbieter } from "./planer.js";

const datumVon = (t) => t.start.slice(0, 10);

// Berechnet die Änderungen für einen Tag, ohne etwas zu tun (Vorschau).
export async function vorschauTag(api, datum, einstellungen) {
  const termine = markiereAlteBloecke(await api.holeTag(datum), einstellungen);
  const soll = planeTag(datum, termine, einstellungen);
  const { anlegen, loeschen } = vergleiche(soll, termine);
  const uebernehmen = termine.filter((t) => t.uebernommen && !loeschen.includes(t));
  const fahrzeitMerken = termine.filter((t) => t.fahrzeitMerken);
  const heute = einstellungen.heute || new Date().toLocaleDateString("sv-SE");
  // Vorhandene Blöcke, denen der Routen-Link fehlt oder deren Erinnerung nicht
  // stimmt, werden nachgebessert. Nur wenn der Abruf diese Angaben mitliefert
  // (vorschau / erinnerung), sonst wissen wir es nicht.
  const nachbessern = termine
    .filter((t) => t.generiert && !loeschen.includes(t))
    .map((t) => ({ t, s: soll.find((s) => s.titel.trim() === t.titel.trim() && s.start === t.start && s.ende === t.ende) }))
    .filter(({ s }) => s)
    .map(({ t, s }) => {
      const felder = {};
      // Strecke fehlt in der Notiz: nur nachtragen, wenn sie diesmal bekannt ist
      // (in der Automatik), das Add-in lässt vorhandene Notizen in Ruhe.
      const fehlt = (text) => s.html.includes(text) && !t.vorschau.includes(text);
      if (s.html && typeof t.vorschau === "string" && (fehlt(ROUTEN_BESCHRIFTUNG) || fehlt(STRECKEN_BESCHRIFTUNG)))
        felder.body = { contentType: "html", content: s.html };
      // Erinnerung nie wieder einschalten, wenn der Block heute oder früher ist:
      // Wegklicken schaltet sie in Outlook aus, sonst käme sie alle 5 Minuten neu.
      const sollAn = (s.erinnerung ?? KEINE_ERINNERUNG) !== KEINE_ERINNERUNG;
      const weggeklickt = sollAn && t.erinnerung === KEINE_ERINNERUNG && datum <= heute;
      if (t.erinnerung !== undefined && (s.erinnerung ?? KEINE_ERINNERUNG) !== t.erinnerung && !weggeklickt)
        Object.assign(felder, erinnerungsFelder(s.erinnerung));
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

// Urlaub (oder Absage) für einen Zeitraum: alle regulären eigenen Kurse von
// "von" bis "bis" bekommen das Kürzel, danach werden die Tage neu geplant.
// kuerzel "" hebt (U)/(A) im Zeitraum wieder auf. Vertretungen (V) und
// Hospitationen (H) bleiben, wie sie sind; sie werden nur zurückgegeben.
export function kurseImZeitraum(termine, kuerzel, einstellungen) {
  const kurse = termine.filter((t) => !t.generiert && kursAnbieter(t, einstellungen));
  const alt = (t) => leseKuerzel(t.titel);
  // Setzen nur bei regulären Kursen; Aufheben nimmt (U) und (A) wieder weg.
  const aendern = kurse.filter((t) => (kuerzel ? alt(t) === "" : ["U", "A"].includes(alt(t))));
  const bleiben = kurse.filter((t) => ["V", "H"].includes(alt(t)));
  return { aendern, bleiben };
}

export async function setzeZeitraum(api, von, bis, kuerzel, einstellungen, { schreiben = true } = {}) {
  const { aendern, bleiben } = kurseImZeitraum(await api.holeVerlauf(von, bis), kuerzel, einstellungen);
  if (!schreiben) return { aendern, bleiben, tage: [] };
  const aktiv = einstellungen.kuerzel[kuerzel]?.aktiv ?? true;
  for (const t of aendern)
    await api.aendere(t.id, { subject: (kuerzel ? `(${kuerzel}) ` : "") + ohneKuerzel(t.titel), showAs: aktiv ? "busy" : "free" });
  const tage = [...new Set(aendern.map(datumVon))].sort();
  for (const d of tage) await planeTagNeu(api, d, einstellungen);
  return { aendern, bleiben, tage };
}
