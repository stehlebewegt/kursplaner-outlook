// Fahrten eines Jahres für die Steuer: alle tatsächlich gefahrenen Fahrten zu
// Kursen, eine Zeile pro Tag, mit km laut Routenplaner. Ohne Geldbeträge.
import { erstelleAuswertung, fahrtenTabelle, MONATE } from "./auswertung.js";

// strecke(von, nach) wie bei der Monatsauswertung.
export function fahrtenJahr(termine, jahr, jetzt, einstellungen, strecke = () => null) {
  const imJahr = termine.filter((t) => t.start.startsWith(`${jahr}-`));
  const reihen = [];
  for (const anbieter of einstellungen.anbieter) {
    const a = erstelleAuswertung(imJahr, anbieter, jetzt, einstellungen, strecke);
    const tab = fahrtenTabelle(a.fahrten);
    const mitKm = tab.kopf.includes("km");
    for (const r of tab.reihen)
      reihen.push({ datum: r[0], anbieter: anbieter.kuerzel, strecke: r[1], fahrten: r[2], km: mitKm && r[3] !== "–" ? Number(String(r[3]).replace(",", ".")) : null });
  }
  reihen.sort((a, b) => (a.datum + a.anbieter < b.datum + b.anbieter ? -1 : 1));
  const monate = MONATE.map((name, i) => {
    const m = reihen.filter((r) => Number(r.datum.slice(5, 7)) === i + 1);
    return { name, fahrten: m.reduce((s, r) => s + r.fahrten, 0), km: m.reduce((s, r) => s + (r.km ?? 0), 0), ohneKm: m.filter((r) => r.km === null).length };
  });
  const runde = (x) => Math.round(x * 10) / 10;
  return {
    reihen,
    monate: monate.map((m) => ({ ...m, km: runde(m.km) })),
    summe: { tage: reihen.length, fahrten: reihen.reduce((s, r) => s + r.fahrten, 0), km: runde(reihen.reduce((s, r) => s + (r.km ?? 0), 0)), ohneKm: reihen.filter((r) => r.km === null).length },
  };
}

const zahl = (x) => String(x).replace(".", ",");
const datumDe = (d) => `${d.slice(8, 10)}.${d.slice(5, 7)}.${d.slice(0, 4)}`;

export function fahrtenPdf(f, jahr, einstellungen, jsPDF, stand) {
  const doc = new jsPDF({ unit: "mm", format: "a4" });
  doc.setFillColor(47, 61, 68).rect(0, 0, 210, 30, "F");
  doc.setTextColor(255).setFont("helvetica", "bold").setFontSize(18).text(`Fahrten ${jahr}`, 15, 15);
  doc.setFont("helvetica", "normal").setFontSize(10)
    .text(`Michael Stehle · Fahrten zu Kursen · ${f.summe.fahrten} Fahrten an ${f.summe.tage} Tagen · ${zahl(f.summe.km)} km`, 15, 23);
  let y = 42;
  doc.setTextColor(38).setFont("helvetica", "bold").setFontSize(11).text("Ziele", 15, y);
  y += 6;
  doc.setFont("helvetica", "normal").setFontSize(9);
  for (const a of einstellungen.anbieter) {
    doc.text(`${a.kuerzel} = ${a.name}, ${a.ort.replace(/\n/g, ", ")}`, 15, y);
    y += 5;
  }
  doc.text("Start und Ziel „Zuhause“ = Wohnadresse. km laut Routenplaner (OpenStreetMap), einfache Strecke je Fahrt.", 15, y);
  y += 10;

  doc.setFont("helvetica", "bold").setFontSize(11).text("Übersicht nach Monaten", 15, y);
  y += 6;
  doc.setFontSize(9).text("Monat", 15, y).text("Fahrten", 80, y, { align: "right" }).text("km", 110, y, { align: "right" });
  doc.setDrawColor(180).line(15, y + 1.5, 110, y + 1.5);
  y += 6;
  doc.setFont("helvetica", "normal");
  for (const m of f.monate.filter((m) => m.fahrten)) {
    doc.text(m.name, 15, y).text(String(m.fahrten), 80, y, { align: "right" }).text(zahl(m.km) + (m.ohneKm ? " *" : ""), 110, y, { align: "right" });
    y += 5;
  }
  doc.setFont("helvetica", "bold").text("Gesamt", 15, y).text(String(f.summe.fahrten), 80, y, { align: "right" }).text(zahl(f.summe.km), 110, y, { align: "right" });
  y += 12;

  const kopf = () => {
    doc.setFont("helvetica", "bold").setFontSize(9).setTextColor(38);
    doc.text("Datum", 15, y).text("Strecke", 42, y).text("Fahrten", 170, y, { align: "right" }).text("km", 195, y, { align: "right" });
    doc.setDrawColor(180).line(15, y + 1.5, 195, y + 1.5);
    y += 6;
    doc.setFont("helvetica", "normal");
  };
  doc.setFont("helvetica", "bold").setFontSize(11).text("Alle Fahrten", 15, y);
  y += 6;
  kopf();
  for (const r of f.reihen) {
    if (y > 282) {
      doc.addPage();
      y = 20;
      kopf();
    }
    doc.text(datumDe(r.datum), 15, y).text(r.strecke, 42, y).text(String(r.fahrten), 170, y, { align: "right" }).text(r.km === null ? "*" : zahl(r.km), 195, y, { align: "right" });
    y += 5;
  }
  y += 4;
  if (y > 282) {
    doc.addPage();
    y = 20;
  }
  doc.setFontSize(8).setTextColor(110);
  if (f.summe.ohneKm) doc.text("* Strecke konnte nicht berechnet werden, km fehlen in der Summe.", 15, y), (y += 4);
  doc.text(`Erstellt am ${datumDe(stand.slice(0, 10))} aus dem Kalender (Kursplaner).`, 15, y);
  return doc;
}
