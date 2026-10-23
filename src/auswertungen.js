// Auswertungen wie in Michaels Excel-Kursbuch (Kursbericht, Wie oft wo, Geld,
// Prüfliste). Rechnet nur aus Kalender-Terminen, lädt nichts.
import { leseKuerzel, ohneKuerzel, kursAnbieter } from "./planer.js";
import { leseProtokoll, gesamt, teileMaterial } from "./protokoll.js";
import { ampel } from "./kursdetails.js";
import { MONATE } from "./auswertung.js";

export const ARTEN = { "": "Regulär", V: "Vertretung", H: "Hospitation", U: "Urlaub", A: "Abgesagt", F: "Ferien" };
const TAG = ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"];
const zahl = (v) => (String(v ?? "").trim() === "" ? null : parseInt(v, 10) || 0);
const istBlock = (t) => t.generiert || /^(Anfahrt|🚗 Jetzt losfahren|Rückfahrt|Fahrt |Vorbereitung|Nachbereitung|Nach-\/Vorbereitung)/.test(t.titel);

// Alle Kurse mit Anbieter, Art und Protokoll.
export function kursZeilen(termine, e) {
  return termine
    .filter((t) => !istBlock(t))
    .map((t) => ({ t, a: kursAnbieter(t, e) }))
    .filter(({ a }) => a)
    .map(({ t, a }) => {
      const k = leseKuerzel(t.titel);
      const p = leseProtokoll(t.text || "");
      return {
        t,
        datum: t.start.slice(0, 10),
        beginn: t.start.slice(11, 16),
        tag: TAG[new Date(`${t.start.slice(0, 10)}T12:00:00Z`).getUTCDay()],
        anbieter: a.kuerzel,
        name: ohneKuerzel(t.titel).replace(/\s*\([^)]+\)\s*$/, "").trim(),
        art: ARTEN[k] ?? "Regulär",
        findetStatt: e.kuerzel[k]?.aktiv !== false,
        p,
        anwesend: zahl(p.anwesend),
        unentschuldigt: zahl(p.unentschuldigt),
        gesamt: zahl(gesamt(p)),
      };
    })
    .sort((x, y) => (x.t.start < y.t.start ? -1 : 1));
}

const zaehle = (liste) => {
  const n = new Map();
  for (const x of liste) n.set(x, (n.get(x) || 0) + 1);
  return [...n].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], "de"));
};

// 1. Kursbericht: Anwesenheit und Ampel je Kurs, Themen, Material.
// anbieter "" = alle; nurEigene ohne Vertretungen und Hospitationen.
export function kursbericht(termine, { von, bis, anbieter = "", nurEigene = false, jetzt }, e) {
  const gehalten = kursZeilen(termine, e).filter(
    (z) => z.datum >= von && z.datum <= bis && z.t.ende <= jetzt && z.findetStatt && (!anbieter || z.anbieter === anbieter) && (!nurEigene || z.art === "Regulär")
  );
  const mit = gehalten.filter((z) => z.anwesend != null);
  const s = (liste, k) => liste.reduce((x, z) => x + (z[k] || 0), 0);
  const quote = (liste) => (s(liste, "gesamt") ? s(liste, "anwesend") / s(liste, "gesamt") : null);
  const gruppen = new Map();
  for (const z of gehalten) {
    const key = `${z.anbieter}|${z.tag}|${z.beginn}`;
    if (!gruppen.has(key)) gruppen.set(key, []);
    gruppen.get(key).push(z);
  }
  const reihenfolge = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"];
  const kurse = [...gruppen.values()]
    .map((g) => {
      const m = g.filter((z) => z.anwesend != null);
      const q = quote(m);
      return {
        anbieter: g[0].anbieter,
        name: zaehle(g.map((z) => z.name))[0][0],
        tag: g[0].tag,
        beginn: g[0].beginn,
        termine: g.length,
        mitProtokoll: m.length,
        schnittGesamt: m.length ? s(m, "gesamt") / m.length : null,
        schnittAnwesend: m.length ? s(m, "anwesend") / m.length : null,
        quote: q,
        unentschuldigt: s(m, "gesamt") ? s(m, "unentschuldigt") / s(m, "gesamt") : null,
        ampel: ampel(q),
      };
    })
    .sort((a, b) => a.anbieter.localeCompare(b.anbieter) || reihenfolge.indexOf(a.tag) - reihenfolge.indexOf(b.tag) || a.beginn.localeCompare(b.beginn));
  return {
    gehalten: gehalten.length,
    mitProtokoll: mit.length,
    ohneProtokoll: gehalten.length - mit.length,
    schnittAnwesend: mit.length ? s(mit, "anwesend") / mit.length : null,
    quote: quote(mit),
    unentschuldigt: s(mit, "gesamt") ? s(mit, "unentschuldigt") / s(mit, "gesamt") : null,
    kurse,
    themen: zaehle(gehalten.map((z) => z.p.thema).filter(Boolean)),
    material: zaehle(gehalten.flatMap((z) => teileMaterial(z.p.material))),
    besonderheiten: gehalten
      .filter((z) => z.p.besonderheiten || /^ja/i.test(z.p.vorkommnis))
      .map((z) => ({ datum: z.datum, kurs: z.name, anbieter: z.anbieter, text: z.p.besonderheiten, vorkommnis: /^ja/i.test(z.p.vorkommnis) })),
  };
}

// 2. Wie oft wo: Kurse je Monat und Anbieter nach Art (nur vergangene).
export const WIE_OFT_ARTEN = ["Regulär", "Vertretung", "Hospitation", "Urlaub", "Abgesagt"];
export function wieOftWo(termine, jahr, jetzt, e) {
  const zeilen = kursZeilen(termine, e).filter((z) => z.datum.startsWith(`${jahr}-`) && z.t.ende <= jetzt);
  const anbieter = [...new Set([...e.anbieter.map((a) => a.kuerzel), ...zeilen.map((z) => z.anbieter)])].filter((k) => zeilen.some((z) => z.anbieter === k));
  const zelle = (liste) => {
    const r = Object.fromEntries(WIE_OFT_ARTEN.map((a) => [a, liste.filter((z) => z.art === a).length]));
    r.gehalten = liste.filter((z) => z.findetStatt).length;
    return r;
  };
  const monate = MONATE.map((name, i) => {
    const m = zeilen.filter((z) => Number(z.datum.slice(5, 7)) === i + 1);
    return { name, je: Object.fromEntries(anbieter.map((k) => [k, zelle(m.filter((z) => z.anbieter === k))])), gehalten: m.filter((z) => z.findetStatt).length };
  });
  return { anbieter, monate, summe: { je: Object.fromEntries(anbieter.map((k) => [k, zelle(zeilen.filter((z) => z.anbieter === k))])), gehalten: zeilen.filter((z) => z.findetStatt).length } };
}

// 3. Geld (nur für Michael): bezahlte Kurse = regulär und Vertretung, die
// stattfinden; erbracht bis jetzt, der Rest geplant. km je Monat und Anbieter
// kommen aus dem Fahrtenbuch (kmJe(kuerzel, monat 1-12) oder null).
export function geld(termine, jahr, jetzt, e, kmJe = () => null) {
  const zeilen = kursZeilen(termine, e).filter((z) => z.datum.startsWith(`${jahr}-`) && z.findetStatt && ["Regulär", "Vertretung"].includes(z.art));
  return e.anbieter
    .filter((a) => a.satzJeKurs > 0)
    .map((a) => {
      const monate = MONATE.map((name, i) => {
        const m = zeilen.filter((z) => z.anbieter === a.kuerzel && Number(z.datum.slice(5, 7)) === i + 1);
        const erbracht = m.filter((z) => z.t.ende <= jetzt).length;
        const km = kmJe(a.kuerzel, i + 1);
        const schluessel = `${jahr}-${String(i + 1).padStart(2, "0")}-${a.kuerzel}`;
        const betrag = m.length * a.satzJeKurs;
        return {
          name,
          schluessel,
          kurse: m.length,
          erbracht,
          geplant: m.length - erbracht,
          betrag,
          km,
          fahrtkosten: km == null ? null : km * (e.kmSatz ?? 0),
          ausbezahlt: !!e.ausbezahlt?.[schluessel],
          ueberMonatsgrenze: a.grenzeMonat > 0 && betrag > a.grenzeMonat,
        };
      });
      const betrag = monate.reduce((x, m) => x + m.betrag, 0);
      return {
        kuerzel: a.kuerzel,
        name: a.name,
        satz: a.satzJeKurs,
        grenzeJahr: a.grenzeJahr,
        grenzeMonat: a.grenzeMonat,
        monate,
        summe: {
          kurse: monate.reduce((x, m) => x + m.kurse, 0),
          betrag,
          ausbezahlt: monate.filter((m) => m.ausbezahlt).reduce((x, m) => x + m.betrag, 0),
          km: monate.every((m) => m.km == null) ? null : monate.reduce((x, m) => x + (m.km || 0), 0),
          fahrtkosten: monate.every((m) => m.fahrtkosten == null) ? null : monate.reduce((x, m) => x + (m.fahrtkosten || 0), 0),
        },
        anteilJahr: a.grenzeJahr > 0 ? betrag / a.grenzeJahr : null,
      };
    });
}

// 4. Prüfliste: was fehlt oder nicht passt (dieses Jahr bis jetzt).
// Protokolle zählen erst ab dem ersten Protokoll im Jahr (vorher gab es keine).
export function pruefliste(termine, jahr, jetzt, e) {
  const zeilen = kursZeilen(termine, e).filter((z) => z.datum.startsWith(`${jahr}-`));
  const vorbei = zeilen.filter((z) => z.t.ende <= jetzt && z.findetStatt);
  const ab = vorbei.find((z) => z.anwesend != null || z.p.thema)?.datum ?? "9999";
  const doppelt = zeilen.filter((z, i) =>
    zeilen.some((y, j) => i !== j && y.datum === z.datum && y.beginn === z.beginn && y.anbieter === z.anbieter)
  );
  return {
    protokollAb: ab === "9999" ? null : ab,
    punkte: [
      { id: "teilnehmer", titel: "Gehaltene Kurse ohne Teilnehmerzahl", tun: "Protokoll öffnen und Anwesend eintragen.", termine: vorbei.filter((z) => z.datum >= ab && z.art !== "Hospitation" && z.anwesend == null) },
      { id: "thema", titel: "Mit Teilnehmern, aber ohne Thema", tun: "Thema im Protokoll ergänzen.", termine: vorbei.filter((z) => z.anwesend != null && !z.p.thema) },
      { id: "doppelt", titel: "Doppelte Termine (gleicher Tag, Beginn, Anbieter)", tun: "In Outlook prüfen und einen löschen.", termine: doppelt },
      { id: "ort", titel: "Vertretung in anderem Verein ohne Adresse", tun: "Adresse im Termin eintragen, sonst fehlt die Fahrt.", termine: zeilen.filter((z) => z.anbieter === e.vertretungExtern?.kuerzel && !(z.t.ort || "").trim()) },
    ].map((x) => ({ ...x, termine: x.termine.map((z) => z.t) })),
  };
}
