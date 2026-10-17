// PDF der Monatsauswertung (mit jsPDF, liegt in web/vendor).
import { fahrtenTabelle, fahrtenSumme } from "../src/auswertung.js";
const datumDe = (d) => `${d.slice(8, 10)}.${d.slice(5, 7)}.${d.slice(0, 4)}`;
const zahl = (x) => String(Math.round(x * 100) / 100).replace(".", ",");

export function erstellePdf({ zeilen, summe, fahrten = [] }, anbieter, zeitraum, stand, { teilnehmer = false, nachtrag = null, jsPDF = window.jspdf?.jsPDF } = {}) {
  const doc = new jsPDF({ unit: "mm", format: "a4" });
  let y = 20;
  doc.setFont("helvetica", "bold").setFontSize(15);
  doc.text(`Monatsauswertung ${zeitraum.name}`, 15, y);
  y += 7;
  doc.setFont("helvetica", "normal").setFontSize(11);
  doc.text(`${anbieter.name} · Übungsleiter: Michael Stehle`, 15, y);
  y += 5;
  doc.setFontSize(9).setTextColor(110);
  doc.text(`Stand: ${datumDe(stand.slice(0, 10))}, ${stand.slice(11, 16)} Uhr`, 15, y);
  doc.setTextColor(0);
  y += 10;

  doc.setFontSize(11);
  const block = [
    ["Geleistet", `${summe.geleistet} Einheiten (${zahl(summe.geleistetStunden)} Std.), davon ${summe.davonVertretung} Vertretung`],
    summe.geplant && ["Noch geplant", `${summe.geplant} Einheiten (${zahl(summe.geplantStunden)} Std.)`],
    summe.ferien && ["Entfällt (Ferien)", `${summe.ferien} Einheiten`],
    ["Fahrten", fahrtenSumme(summe).replace("Fahrten: ", "")],
  ].filter(Boolean);
  if (teilnehmer)
    block.push(["Teilnehmer", summe.teilnehmerSchnitt === null ? "keine Angaben" : `${summe.teilnehmer} gesamt, im Schnitt ${zahl(summe.teilnehmerSchnitt)}`]);
  for (const [k, v] of block) {
    doc.setFont("helvetica", "bold").text(k, 15, y);
    doc.setFont("helvetica", "normal").text(v, 55, y);
    y += 6;
  }
  y += 6;

  // Tabelle mit Kopfzeile, Seitenumbruch wiederholt den Kopf.
  const tabelle = (titel, spalten, reihen) => {
    if (y > 260) {
      doc.addPage();
      y = 20;
    }
    doc.setFont("helvetica", "bold").setFontSize(11).text(titel, 15, y);
    y += 6;
    const kopf = () => {
      doc.setFont("helvetica", "bold").setFontSize(9);
      for (const [t, x] of spalten) doc.text(t, x, y);
      y += 2;
      doc.line(15, y, 195, y);
      y += 5;
      doc.setFont("helvetica", "normal");
    };
    kopf();
    for (const werte of reihen) {
      if (y > 280) {
        doc.addPage();
        y = 20;
        kopf();
      }
      werte.forEach((w, i) => {
        const breite = (spalten[i + 1]?.[1] ?? 197) - spalten[i][1] - 2;
        doc.text(doc.splitTextToSize(String(w), breite)[0], spalten[i][1], y);
      });
      y += 6;
    }
    y += 6;
  };

  const kursSpalten = [["Datum", 15], ["Kurszeit", 38], ["Kurs", 62], ["Art", 132], ["Status", 158]];
  if (teilnehmer) kursSpalten.push(["TN", 190]);
  tabelle(
    "Kurse",
    kursSpalten,
    zeilen.map((z) => {
      const w = [datumDe(z.datum), `${z.von}–${z.bis}`, z.kurs, z.art, z.status];
      if (teilnehmer) w.push(z.anwesend ?? "–");
      return w;
    })
  );
  if (fahrten.length) {
    const { kopf, reihen } = fahrtenTabelle(fahrten);
    const x = [15, 40, 140, 160, 180];
    tabelle("Fahrten", kopf.map((k, i) => [k, x[i]]), reihen.map((r) => [datumDe(r[0]), ...r.slice(1)]));
  }

  // Nachtrag Vormonat: was beim letzten Bericht noch geplant war, jetzt mit Ergebnis.
  if (nachtrag?.zeilen.length) {
    if (y > 230) {
      doc.addPage();
      y = 20;
    }
    const n = nachtrag.summe;
    doc.setFont("helvetica", "bold").setFontSize(13).text(`Nachtrag ${nachtrag.name}`, 15, y);
    y += 6;
    doc.setFont("helvetica", "normal").setFontSize(10);
    const zeilenText = [
      `Kurse ab ${datumDe(nachtrag.ab)}, die bei der letzten Auswertung noch geplant waren, und Änderungen seitdem.`,
      `Geleistet: ${n.geleistet} Einheiten (${zahl(n.geleistetStunden)} Std.)` +
        (n.neu ? `, davon ${n.neu} neu dazugekommen` : "") +
        (n.entfallen ? ` · entfallen: ${n.entfallen}` : "") +
        ` · ${fahrtenSumme(n)}`,
    ];
    for (const t of zeilenText) {
      doc.text(t, 15, y);
      y += 5;
    }
    y += 4;
    tabelle(
      `Kurse ${nachtrag.name}`,
      [["Datum", 15], ["Kurszeit", 38], ["Kurs", 62], ["Art", 120], ["Status", 145], ["Änderung", 172]],
      nachtrag.zeilen.map((z) => [datumDe(z.datum), `${z.von}–${z.bis}`, z.kurs, z.art, z.status, z.aenderung || ""])
    );
    if (nachtrag.fahrten.some((f) => f.status === "geleistet")) {
      const { kopf, reihen } = fahrtenTabelle(nachtrag.fahrten);
      const x = [15, 40, 140, 160, 180];
      tabelle(`Fahrten ${nachtrag.name}`, kopf.map((k, i) => [k, x[i]]), reihen.map((r) => [datumDe(r[0]), ...r.slice(1)]));
    }
  }
  return doc;
}
