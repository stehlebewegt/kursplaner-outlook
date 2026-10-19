// Wochenvorschau als PDF (gleicher Inhalt wie die HTML-Mail), mit jsPDF.
import { wochenDaten, plusTage } from "./hinweise.js";
import { ohneKuerzel } from "./planer.js";

const WOCHENTAGE = ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"];
const tagLang = (d) => `${WOCHENTAGE[new Date(`${d}T12:00:00Z`).getUTCDay()]}, ${d.slice(8, 10)}.${d.slice(5, 7)}.${d.slice(0, 4)}`;
const kurz = (d) => `${d.slice(8, 10)}.${d.slice(5, 7)}.`;
const zeit = (t) => `${t.start.slice(11, 16)}–${t.ende.slice(11, 16)}`;
const FARBE = { SVE: [47, 107, 86], PTM: [47, 93, 124] };

export function wochenPdf(termine, montag, einstellungen, jsPDF) {
  const w = wochenDaten(termine, montag, einstellungen);
  const doc = new jsPDF({ unit: "mm", format: "a4" });
  // Kopf: dunkler Balken wie in der Mail.
  doc.setFillColor(47, 61, 68).rect(0, 0, 210, 30, "F");
  doc.setTextColor(255).setFont("helvetica", "bold").setFontSize(18).text(`Deine Woche ab ${kurz(montag)}`, 15, 15);
  doc.setFont("helvetica", "normal").setFontSize(10)
    .text(`${w.kurse} Kurse · ${String(w.fahrStd).replace(".", ",")} Std. Fahrzeit · ${kurz(montag)} bis ${kurz(plusTage(montag, 6))}`, 15, 23);
  let y = 42;
  const platz = (mm) => {
    if (y + mm > 285) {
      doc.addPage();
      y = 20;
    }
  };
  for (const tag of w.tage) {
    const frei = !tag.kurse.length && !tag.privat.length;
    platz(16);
    doc.setTextColor(frei ? 150 : 53, frei ? 160 : 80, frei ? 165 : 90).setFont("helvetica", "bold").setFontSize(12).text(tagLang(tag.datum), 15, y);
    doc.setDrawColor(frei ? 200 : 53, frei ? 205 : 80, frei ? 208 : 90).setLineWidth(0.5).line(15, y + 2, 195, y + 2);
    y += 8;
    if (frei) {
      doc.setTextColor(140).setFont("helvetica", "normal").setFontSize(10).text("frei", 15, y);
      y += 8;
      continue;
    }
    const reihe = (z, titel, marke, farbe, grau) => {
      platz(7);
      doc.setFont("helvetica", "normal").setFontSize(10).setTextColor(grau ? 150 : 90).text(z, 15, y);
      doc.setTextColor(grau ? 150 : 38).text(titel, 45, y);
      if (grau) {
        const b = doc.getTextWidth(titel);
        doc.setDrawColor(150).setLineWidth(0.2).line(45, y - 1.2, 45 + b, y - 1.2);
      }
      if (marke) doc.setFont("helvetica", "bold").setFontSize(8).setTextColor(...farbe).text(marke, 195, y, { align: "right" });
      y += 6.5;
    };
    for (const { t, anbieter, aktiv, zusatz } of tag.kurse)
      reihe(zeit(t), ohneKuerzel(t.titel), anbieter + (zusatz ? ` · ${zusatz}` : ""), FARBE[anbieter] || [55, 71, 79], !aktiv);
    for (const t of tag.privat) reihe(zeit(t), t.titel.trim(), "privat", [91, 103, 112], false);
    if (tag.fahrten) {
      platz(7);
      doc.setFont("helvetica", "italic").setFontSize(9).setTextColor(91, 103, 112)
        .text(`${Math.round(tag.fahrMin)} Min. Fahrzeit (${tag.fahrten} ${tag.fahrten === 1 ? "Fahrt" : "Fahrten"})`, 45, y);
      y += 6.5;
    }
    y += 3;
  }
  return doc;
}
