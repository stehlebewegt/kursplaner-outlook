// Monatsauswertung pro Anbieter: welche Kurse geleistet, welche noch geplant,
// regulär oder Vertretung, mit Teilnehmern aus dem Kursprotokoll und den
// Fahrten zu den Kursen, die stattgefunden haben (bzw. noch geplant sind).
// (U) und (A) zählen nicht, (F) wird als "entfällt (Ferien)" nur aufgeführt.
import { leseKuerzel, ohneKuerzel, kursAnbieter } from "./planer.js";
import { leseProtokoll } from "./protokoll.js";

const minuten = (t) => (Date.parse(t.ende) - Date.parse(t.start)) / 60000;

export function erstelleAuswertung(termine, anbieter, jetzt, einstellungen) {
  const zeilen = termine
    .filter((t) => !t.generiert && kursAnbieter(t, einstellungen) === anbieter)
    .filter((t) => !["U", "A"].includes(leseKuerzel(t.titel)))
    .map((t) => {
      const k = leseKuerzel(t.titel);
      const p = leseProtokoll(t.text);
      return {
        datum: t.start.slice(0, 10),
        von: t.start.slice(11, 16),
        bis: t.ende.slice(11, 16),
        kurs: ohneKuerzel(t.titel).trim(),
        vollerTitel: t.titel.trim(),
        art: k === "V" ? "Vertretung" : "regulär",
        status: k === "F" ? "entfällt (Ferien)" : t.start <= jetzt ? "geleistet" : "geplant",
        minuten: minuten(t),
        anwesend: p.anwesend === "" ? null : parseInt(p.anwesend, 10) || 0,
        thema: p.thema,
      };
    })
    .sort((a, b) => (a.datum + a.von < b.datum + b.von ? -1 : 1));

  const zaehle = (f) => zeilen.filter(f);
  const geleistet = zaehle((z) => z.status === "geleistet");
  const geplant = zaehle((z) => z.status === "geplant");
  const mitTeilnehmern = geleistet.filter((z) => z.anwesend !== null);
  const summe = {
    geleistet: geleistet.length,
    geleistetStunden: geleistet.reduce((s, z) => s + z.minuten, 0) / 60,
    davonVertretung: geleistet.filter((z) => z.art === "Vertretung").length,
    geplant: geplant.length,
    geplantStunden: geplant.reduce((s, z) => s + z.minuten, 0) / 60,
    ferien: zaehle((z) => z.status === "entfällt (Ferien)").length,
    teilnehmer: mitTeilnehmern.reduce((s, z) => s + z.anwesend, 0),
    teilnehmerSchnitt: mitTeilnehmern.length
      ? Math.round((mitTeilnehmern.reduce((s, z) => s + z.anwesend, 0) / mitTeilnehmern.length) * 10) / 10
      : null,
    ohneProtokoll: geleistet.length - mitTeilnehmern.length,
  };
  const fahrten = sammleFahrten(termine, zeilen, einstellungen, anbieter);
  const gefahren = fahrten.filter((f) => f.status === "geleistet");
  summe.fahrten = gefahren.length;
  summe.fahrtStunden = gefahren.reduce((s, f) => s + f.minuten, 0) / 60;
  return { zeilen, summe, fahrten };
}

const FAHRT_TITEL = /^(Anfahrt|Rückfahrt|Fahrt [A-Z]+ → [A-Z]+) - (.+)$/;

// Fahrt-Blöcke (Kategorie Fahrzeit) an Tagen, an denen bei diesem Anbieter
// ein Kurs stattfand: Der Titel nennt einen Kurs des Tages, oder der Ort ist
// der Anbieter. Abgesagte Kurse haben keine Fahrt, also zählt nur, was
// wirklich gefahren wurde. "Fahrtkosten nicht abrechenbar" fällt raus.
function sammleFahrten(termine, zeilen, e, anbieter) {
  const zaehlt = zeilen.filter((z) => z.status !== "entfällt (Ferien)");
  const kurse = new Map(zaehlt.map((z) => [`${z.datum} ${z.vollerTitel}`, z]));
  const tage = new Map();
  for (const z of zaehlt) if (!tage.has(z.datum) || z.status === "geleistet") tage.set(z.datum, z.status);
  const beiAnbieter = (t) => anbieter.ortErkennung.some((o) => (t.ort || "").includes(o));
  return termine
    .filter((t) => (t.kategorien || []).includes(e.kategorieFahrt))
    .filter((t) => !(t.kategorien || []).includes(e.kategorieNichtAbrechenbar))
    .map((t) => {
      const m = t.titel.trim().match(FAHRT_TITEL);
      if (!m) return null;
      const datum = t.start.slice(0, 10);
      const kurs = kurse.get(`${datum} ${m[2].trim()}`);
      const status = kurs ? kurs.status : beiAnbieter(t) && tage.get(datum);
      if (!status) return null;
      return {
        datum,
        von: t.start.slice(11, 16),
        bis: t.ende.slice(11, 16),
        fahrt: m[1] === "Anfahrt" ? "Hinfahrt" : m[1],
        kurs: ohneKuerzel(m[2]).trim(),
        minuten: minuten(t),
        status,
      };
    })
    .filter(Boolean)
    .sort((a, b) => (a.datum + a.von < b.datum + b.von ? -1 : 1));
}

export const MONATE = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"];

// "2026-10" → { von: "2026-10-01", bis: "2026-10-31", name: "Oktober 2026" }
export function monatsZeitraum(monat) {
  const [j, m] = monat.split("-").map(Number);
  const letzter = new Date(Date.UTC(j, m, 0)).getUTCDate();
  return { von: `${monat}-01`, bis: `${monat}-${String(letzter).padStart(2, "0")}`, name: `${MONATE[m - 1]} ${j}` };
}
