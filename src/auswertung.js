// Monatsauswertung pro Anbieter: welche Kurse geleistet, welche noch geplant,
// regulär oder Vertretung, mit Teilnehmern aus dem Kursprotokoll und den
// Fahrten zu den Kursen, die stattgefunden haben (bzw. noch geplant sind).
// (U) und (A) zählen nicht, (F) wird als "entfällt (Ferien)" nur aufgeführt.
import { leseKuerzel, ohneKuerzel, kursAnbieter } from "./planer.js";
import { leseProtokoll } from "./protokoll.js";

const minuten = (t) => (Date.parse(t.ende) - Date.parse(t.start)) / 60000;

// strecke(von, nach): berechnete Fahrt { km, minuten } zwischen "heim" und
// Anbieter-Kürzeln (siehe automatik/route.js); ohne sie bleiben km leer.
export function erstelleAuswertung(termine, anbieter, jetzt, einstellungen, strecke = () => null) {
  const zeilen = termine
    .filter((t) => !t.generiert && kursAnbieter(t, einstellungen)?.kuerzel === anbieter.kuerzel)
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
        art: k === "V" ? "Vertretung" : k === "H" ? "Hospitation" : "regulär",
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
  const fahrten = sammleFahrten(termine, zeilen, einstellungen, anbieter).map((f) => {
    const r = strecke(f.von, f.nach);
    return { ...f, km: r?.km ?? null, minutenBerechnet: r?.minuten ?? null };
  });
  const gefahren = fahrten.filter((f) => f.status === "geleistet");
  summe.fahrten = gefahren.length;
  summe.fahrtKm = gefahren.every((f) => f.km !== null) ? Math.round(gefahren.reduce((s, f) => s + f.km, 0) * 10) / 10 : null;
  return { zeilen, summe, fahrten };
}

const FAHRT_TITEL = /^(Anfahrt|🚗 Jetzt losfahren|Rückfahrt|Fahrt ([A-Z]+) → ([A-Z]+)) - (.+)$/;

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
      const kurs = kurse.get(`${datum} ${m[4].trim()}`);
      const status = kurs ? kurs.status : beiAnbieter(t) && tage.get(datum);
      if (!status) return null;
      // Vertretung in anderen Vereinen: Ziel ist der Ort aus dem Fahrt-Block.
      const ziel = anbieter.kuerzel === e.vertretungExtern?.kuerzel ? t.ort || anbieter.kuerzel : anbieter.kuerzel;
      const hin = m[1] === "Anfahrt" || m[1] === "🚗 Jetzt losfahren";
      const [von, nach] = hin ? ["heim", ziel] : m[1] === "Rückfahrt" ? [ziel, "heim"] : [m[2], m[3]];
      return {
        datum,
        uhrzeit: t.start.slice(11, 16),
        fahrt: hin ? "Hinfahrt" : m[1] === "Rückfahrt" ? "Rückfahrt" : "Weiterfahrt",
        von,
        nach,
        kurs: ohneKuerzel(m[4]).trim(),
        minuten: minuten(t),
        status,
      };
    })
    .filter(Boolean)
    .sort((a, b) => (a.datum + a.uhrzeit < b.datum + b.uhrzeit ? -1 : 1));
}

// Fahrten-Tabelle (für PDF und E-Mail gleich): eine Zeile pro Tag, nur
// tatsächlich gefahrene. Mit berechneter Strecke km und Fahrminuten laut
// Routenplaner, sonst die Minuten aus dem Kalender.
const ortName = (k) => (k === "heim" ? "Zuhause" : k);
const zahlDe = (x) => String(Math.round(x * 10) / 10).replace(".", ",");
export function fahrtenTabelle(fahrten) {
  const gefahren = fahrten.filter((f) => f.status === "geleistet");
  const mitKm = gefahren.some((f) => f.km !== null && f.km !== undefined);
  const kopf = mitKm ? ["Datum", "Strecke", "Fahrten", "km", "Min."] : ["Datum", "Strecke", "Fahrten", "Min. (Kalender)"];
  const tage = new Map();
  for (const f of gefahren) tage.set(f.datum, [...(tage.get(f.datum) || []), f]);
  const reihen = [...tage].map(([datum, fs]) => {
    const orte = [fs[0].von];
    for (const f of fs) {
      if (orte[orte.length - 1] !== f.von) orte.push(f.von);
      orte.push(f.nach);
    }
    const r = [datum, orte.map(ortName).join(" – "), fs.length];
    if (!mitKm) return [...r, fs.reduce((s, f) => s + f.minuten, 0)];
    const fehlt = fs.some((f) => f.km === null);
    return [...r, fehlt ? "–" : zahlDe(fs.reduce((s, f) => s + f.km, 0)), fehlt ? "–" : fs.reduce((s, f) => s + f.minutenBerechnet, 0)];
  });
  return { kopf, reihen };
}

export function fahrtenSumme(summe) {
  return `Fahrten: ${summe.fahrten ?? 0}${summe.fahrtKm != null ? `, zusammen ${zahlDe(summe.fahrtKm)} km` : ""}`;
}

export const MONATE = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"];

// "2026-10" → { von: "2026-10-01", bis: "2026-10-31", name: "Oktober 2026" }
export function monatsZeitraum(monat) {
  const [j, m] = monat.split("-").map(Number);
  const letzter = new Date(Date.UTC(j, m, 0)).getUTCDate();
  return { von: `${monat}-01`, bis: `${monat}-${String(letzter).padStart(2, "0")}`, name: `${MONATE[m - 1]} ${j}` };
}

// Stand eines Berichts zum Merken (kompakt, landet am E-Mail-Entwurf).
const schluessel = (z) => `${z.datum} ${z.von} ${z.kurs}`;
export const stand = (a, jetzt) => ({ stand: jetzt, z: a.zeilen.map((z) => [z.datum, z.von, z.bis, z.kurs, z.art, z.status]) });

// Nachtrag für den Vormonat: was beim letzten Bericht noch offen war (ab
// dessen Stand) mit dem tatsächlichen Ergebnis, dazu neue, entfallene und
// geänderte Kurse. alt = gemerkter Stand oder null (dann ab abDatum, ohne Vergleich).
export function nachtrag(a, alt, abDatum) {
  const ab = alt ? alt.stand.slice(0, 10) : abDatum;
  const vorher = new Map((alt?.z || []).map(([datum, von, bis, kurs, art, status]) => [schluessel({ datum, von, kurs }), { datum, von, bis, kurs, art, status }]));
  const zeilen = [];
  for (const z of a.zeilen) {
    const v = vorher.get(schluessel(z));
    vorher.delete(schluessel(z));
    const aenderung = !alt ? "" : !v ? "neu" : v.art !== z.art || v.bis !== z.bis || (v.status !== z.status && v.status !== "geplant") ? "geändert" : "";
    if (z.datum >= ab || aenderung) zeilen.push({ ...z, aenderung });
  }
  for (const v of vorher.values()) zeilen.push({ ...v, status: "entfallen", aenderung: "entfallen" });
  zeilen.sort((x, y) => (x.datum + x.von < y.datum + y.von ? -1 : 1));
  const geleistet = zeilen.filter((z) => z.status === "geleistet");
  const fahrten = a.fahrten.filter((f) => f.datum >= ab);
  const gefahren = fahrten.filter((f) => f.status === "geleistet");
  return {
    ab,
    zeilen,
    fahrten,
    summe: {
      geleistet: geleistet.length,
      geleistetStunden: geleistet.reduce((s, z) => s + z.minuten, 0) / 60,
      neu: zeilen.filter((z) => z.aenderung === "neu").length,
      entfallen: zeilen.filter((z) => z.aenderung === "entfallen").length,
      fahrten: gefahren.length,
      fahrtKm: gefahren.every((f) => f.km !== null) ? Math.round(gefahren.reduce((s, f) => s + f.km, 0) * 10) / 10 : null,
    },
  };
}
