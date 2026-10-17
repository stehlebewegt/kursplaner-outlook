// PDF der Monatsauswertung (mit jsPDF, liegt in web/vendor).
// Blatt 1: laufender Monat. Blatt 2: Nachtrag zum Vormonat (falls vorhanden).
// Jedes Blatt beginnt auf einer neuen Seite; die Schrift wird kleiner, bis
// der Abschnitt auf eine Seite passt (höchstens auf 60 %), sonst geht er
// auf weiteren Seiten weiter.
import { fahrtenTabelle, fahrtenSumme } from "../src/auswertung.js";
const datumDe = (d) => `${d.slice(8, 10)}.${d.slice(5, 7)}.${d.slice(0, 4)}`;
const zahl = (x) => String(Math.round(x * 100) / 100).replace(".", ",");

const OBEN = 20;
const UNTEN = 280;
const MIN_FAKTOR = 0.6;

// Zeichnet die Teile eines Blatts ab neuer Seite mit Schriftfaktor f.
// trocken = nur messen (liefert die benötigte Höhe, ohne zu zeichnen).
function blatt(doc, teile, f, trocken) {
  let y = OBEN;
  const neueSeite = () => {
    if (!trocken) doc.addPage();
    y = OBEN;
  };
  const schrift = (stil, groesse) => !trocken && doc.setFont("helvetica", stil).setFontSize(groesse * f);
  const text = (t, x) => !trocken && doc.text(t, x, y);
  for (const t of teile) {
    if (t.art === "abstand") y += t.mm * f;
    else if (t.art === "zeile") {
      schrift(t.fett ? "bold" : "normal", t.groesse);
      if (!trocken && t.grau) doc.setTextColor(110);
      text(t.text, 15);
      if (!trocken) doc.setTextColor(0);
      y += t.hoehe * f;
    } else if (t.art === "block") {
      for (const [k, v] of t.paare) {
        schrift("bold", 11);
        text(k, 15);
        schrift("normal", 11);
        text(v, 55);
        y += 6 * f;
      }
    } else if (t.art === "tabelle") {
      if (!trocken && y > UNTEN - 20 * f) neueSeite();
      schrift("bold", 11);
      text(t.titel, 15);
      y += 6 * f;
      const kopf = () => {
        schrift("bold", 9);
        for (const [k, x] of t.spalten) text(k, x);
        y += 2 * f;
        if (!trocken) doc.line(15, y, 195, y);
        y += 5 * f;
      };
      kopf();
      for (const werte of t.reihen) {
        if (!trocken && y > UNTEN) {
          neueSeite();
          kopf();
        }
        schrift("normal", 9);
        if (!trocken)
          werte.forEach((w, i) => {
            const breite = (t.spalten[i + 1]?.[1] ?? 197) - t.spalten[i][1] - 2;
            doc.text(doc.splitTextToSize(String(w), breite)[0], t.spalten[i][1], y);
          });
        y += 6 * f;
      }
      y += 6 * f;
    }
  }
  return y - OBEN;
}

function kopfteile(titel, unterzeile, stand) {
  return [
    { art: "zeile", text: titel, groesse: 15, fett: true, hoehe: 7 },
    { art: "zeile", text: unterzeile, groesse: 11, hoehe: 5 },
    { art: "zeile", text: `Stand: ${datumDe(stand.slice(0, 10))}, ${stand.slice(11, 16)} Uhr`, groesse: 9, grau: true, hoehe: 10 },
  ];
}

const fahrtenTeil = (titel, fahrten) => {
  const { kopf, reihen } = fahrtenTabelle(fahrten);
  const x = [15, 40, 140, 160, 180];
  return { art: "tabelle", titel, spalten: kopf.map((k, i) => [k, x[i]]), reihen: reihen.map((r) => [datumDe(r[0]), ...r.slice(1)]) };
};

export function erstellePdf({ zeilen, summe, fahrten = [] }, anbieter, zeitraum, stand, { teilnehmer = false, nachtrag = null, jsPDF = window.jspdf?.jsPDF } = {}) {
  const doc = new jsPDF({ unit: "mm", format: "a4" });
  const unterzeile = `${anbieter.name} · Übungsleiter: Michael Stehle`;

  const paare = [
    ["Geleistet", `${summe.geleistet} Einheiten (${zahl(summe.geleistetStunden)} Std.), davon ${summe.davonVertretung} Vertretung`],
    summe.geplant && ["Noch geplant", `${summe.geplant} Einheiten (${zahl(summe.geplantStunden)} Std.)`],
    summe.ferien && ["Entfällt (Ferien)", `${summe.ferien} Einheiten`],
    ["Fahrten", fahrtenSumme(summe).replace("Fahrten: ", "")],
  ].filter(Boolean);
  if (teilnehmer)
    paare.push(["Teilnehmer", summe.teilnehmerSchnitt === null ? "keine Angaben" : `${summe.teilnehmer} gesamt, im Schnitt ${zahl(summe.teilnehmerSchnitt)}`]);
  const kursSpalten = [["Datum", 15], ["Kurszeit", 38], ["Kurs", 62], ["Art", 132], ["Status", 158]];
  if (teilnehmer) kursSpalten.push(["TN", 190]);
  const blatt1 = [
    ...kopfteile(`Monatsauswertung ${zeitraum.name}`, unterzeile, stand),
    { art: "block", paare },
    { art: "abstand", mm: 6 },
    {
      art: "tabelle",
      titel: "Kurse",
      spalten: kursSpalten,
      reihen: zeilen.map((z) => {
        const w = [datumDe(z.datum), `${z.von}–${z.bis}`, z.kurs, z.art, z.status];
        if (teilnehmer) w.push(z.anwesend ?? "–");
        return w;
      }),
    },
  ];
  if (fahrten.some((f) => f.status === "geleistet")) blatt1.push(fahrtenTeil("Fahrten", fahrten));

  const blaetter = [blatt1];
  // Nachtrag Vormonat: was beim letzten Bericht noch geplant war, jetzt mit Ergebnis.
  if (nachtrag?.zeilen.length) {
    const n = nachtrag.summe;
    const blatt2 = [
      ...kopfteile(`Nachtrag ${nachtrag.name}`, unterzeile, stand),
      { art: "zeile", text: `Kurse ab ${datumDe(nachtrag.ab)}, die bei der letzten Auswertung noch geplant waren, und Änderungen seitdem.`, groesse: 10, hoehe: 8 },
      {
        art: "block",
        paare: [
          ["Geleistet", `${n.geleistet} Einheiten (${zahl(n.geleistetStunden)} Std.)${n.neu ? `, davon ${n.neu} neu dazugekommen` : ""}`],
          n.entfallen && ["Entfallen", `${n.entfallen} Einheiten`],
          ["Fahrten", fahrtenSumme(n).replace("Fahrten: ", "")],
        ].filter(Boolean),
      },
      { art: "abstand", mm: 6 },
      {
        art: "tabelle",
        titel: "Kurse",
        spalten: [["Datum", 15], ["Kurszeit", 38], ["Kurs", 62], ["Art", 120], ["Status", 145], ["Änderung", 172]],
        reihen: nachtrag.zeilen.map((z) => [datumDe(z.datum), `${z.von}–${z.bis}`, z.kurs, z.art, z.status, z.aenderung || ""]),
      },
    ];
    if (nachtrag.fahrten.some((f) => f.status === "geleistet")) blatt2.push(fahrtenTeil("Fahrten", nachtrag.fahrten));
    blaetter.push(blatt2);
  }

  blaetter.forEach((teile, i) => {
    if (i) doc.addPage();
    const f = Math.min(1, Math.max(MIN_FAKTOR, (UNTEN - OBEN) / blatt(doc, teile, 1, true)));
    blatt(doc, teile, f, false);
  });

  // Fußzeile mit Seitenzahl, damit man einzelne Blätter gezielt drucken kann.
  const seiten = doc.getNumberOfPages();
  for (let s = 1; s <= seiten; s++) {
    doc.setPage(s);
    doc.setFont("helvetica", "normal").setFontSize(8).setTextColor(110);
    doc.text(`Seite ${s} von ${seiten}`, 195, 290, { align: "right" });
    doc.setTextColor(0);
  }
  return doc;
}
