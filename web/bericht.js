// PDF der Monatsauswertung (mit jsPDF, liegt in web/vendor).
const datumDe = (d) => `${d.slice(8, 10)}.${d.slice(5, 7)}.${d.slice(0, 4)}`;
const zahl = (x) => String(Math.round(x * 100) / 100).replace(".", ",");

export function erstellePdf({ zeilen, summe, fahrten = [] }, anbieter, zeitraum, stand, { teilnehmer = false, jsPDF = window.jspdf?.jsPDF } = {}) {
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
    ["Fahrten", `${summe.fahrten ?? 0} Fahrten (${zahl(summe.fahrtStunden ?? 0)} Std.)`],
  ];
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

  const kursSpalten = [["Datum", 15], ["Zeit", 38], ["Kurs", 62], ["Art", 132], ["Status", 158]];
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
  if (fahrten.length)
    tabelle(
      "Fahrten",
      [["Datum", 15], ["Zeit", 38], ["Fahrt", 62], ["zum Kurs", 100], ["Min.", 165], ["Status", 178]],
      fahrten.map((f) => [datumDe(f.datum), `${f.von}–${f.bis}`, f.fahrt, f.kurs, f.minuten, f.status])
    );
  return doc;
}
