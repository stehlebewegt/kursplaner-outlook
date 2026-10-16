// PDF der Monatsauswertung (mit jsPDF, liegt in web/vendor).
const datumDe = (d) => `${d.slice(8, 10)}.${d.slice(5, 7)}.${d.slice(0, 4)}`;
const zahl = (x) => String(Math.round(x * 100) / 100).replace(".", ",");

export function erstellePdf({ zeilen, summe }, anbieter, zeitraum, stand) {
  const { jsPDF } = window.jspdf;
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
    ["Noch geplant", `${summe.geplant} Einheiten (${zahl(summe.geplantStunden)} Std.)`],
    ["Entfällt (Ferien)", `${summe.ferien} Einheiten`],
    ["Teilnehmer", summe.teilnehmerSchnitt === null ? "keine Angaben" : `${summe.teilnehmer} gesamt, im Schnitt ${zahl(summe.teilnehmerSchnitt)}`],
  ];
  for (const [k, v] of block) {
    doc.setFont("helvetica", "bold").text(k, 15, y);
    doc.setFont("helvetica", "normal").text(v, 55, y);
    y += 6;
  }
  y += 6;

  const spalten = [
    ["Datum", 15], ["Zeit", 38], ["Kurs", 62], ["Art", 132], ["Status", 158], ["TN", 190],
  ];
  const kopf = () => {
    doc.setFont("helvetica", "bold").setFontSize(9);
    for (const [t, x] of spalten) doc.text(t, x, y);
    y += 2;
    doc.line(15, y, 195, y);
    y += 5;
    doc.setFont("helvetica", "normal");
  };
  kopf();
  for (const z of zeilen) {
    if (y > 280) {
      doc.addPage();
      y = 20;
      kopf();
    }
    const werte = [datumDe(z.datum), `${z.von}–${z.bis}`, z.kurs, z.art, z.status, z.anwesend ?? "–"];
    werte.forEach((w, i) => doc.text(doc.splitTextToSize(String(w), i === 2 ? 68 : 40)[0], spalten[i][1], y));
    y += 6;
  }
  return doc;
}
