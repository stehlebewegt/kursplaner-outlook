// Monatsauswertung als HTML-Tabelle für den Text eines E-Mail-Entwurfs.
const datumDe = (d) => `${d.slice(8, 10)}.${d.slice(5, 7)}.${d.slice(0, 4)}`;
const zahl = (x) => String(Math.round(x * 100) / 100).replace(".", ",");
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

function tabelle(kopf, reihen) {
  const zelle = (tag, w) => `<${tag} style="border:1px solid #ccc;padding:2px 6px">${esc(w)}</${tag}>`;
  return (
    `<table style="border-collapse:collapse;font-size:13px">` +
    `<tr>${kopf.map((k) => zelle("th", k)).join("")}</tr>` +
    reihen.map((r) => `<tr>${r.map((w) => zelle("td", w)).join("")}</tr>`).join("") +
    `</table>`
  );
}

export function berichtHtml({ zeilen, summe, fahrten = [] }, anbieter, zeitraum, { teilnehmer = false } = {}) {
  const kursKopf = ["Datum", "Zeit", "Kurs", "Art", "Status"];
  if (teilnehmer) kursKopf.push("TN");
  const kursReihen = zeilen.map((z) => {
    const r = [datumDe(z.datum), `${z.von}–${z.bis}`, z.kurs, z.art, z.status];
    if (teilnehmer) r.push(z.anwesend ?? "–");
    return r;
  });
  const summen = [
    `Geleistet: ${summe.geleistet} Einheiten (${zahl(summe.geleistetStunden)} Std.), davon ${summe.davonVertretung} Vertretung`,
    summe.geplant ? `Noch geplant: ${summe.geplant} Einheiten (${zahl(summe.geplantStunden)} Std.)` : "",
    summe.ferien ? `Entfällt (Ferien): ${summe.ferien} Einheiten` : "",
    `Fahrten: ${summe.fahrten ?? 0} (${zahl(summe.fahrtStunden ?? 0)} Std.)`,
  ].filter(Boolean);
  if (teilnehmer && summe.teilnehmerSchnitt !== null) summen.push(`Teilnehmer: ${summe.teilnehmer} gesamt, im Schnitt ${zahl(summe.teilnehmerSchnitt)}`);
  return (
    `<p>Hallo,</p><p>hier meine Monatsauswertung für ${esc(zeitraum.name)} (${esc(anbieter.name)}).</p>` +
    `<p>${summen.map(esc).join("<br>")}</p>` +
    `<h3>Kurse</h3>` +
    tabelle(kursKopf, kursReihen) +
    (fahrten.length
      ? `<h3>Fahrten</h3>` +
        tabelle(["Datum", "Zeit", "Fahrt", "zum Kurs", "Minuten", "Status"], fahrten.map((f) => [datumDe(f.datum), `${f.von}–${f.bis}`, f.fahrt, f.kurs, f.minuten, f.status]))
      : "") +
    `<p>Viele Grüße<br>Michael Stehle</p>`
  );
}
