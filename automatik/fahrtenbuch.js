// Anfang Januar legt die Automatik einen Entwurf mit dem PDF aller Fahrten des
// Vorjahres an (einmal, Kennung am Entwurf). probe = "JJJJ": sofort, mit PROBE.
// Öffentliche Ausgabe nur mit Zahlen, ohne km.
import { BERICHT, ladeJsPdf, berechneStrecken } from "./monatsbericht.js";
import { erstelleAuswertung } from "../src/auswertung.js";
import { fahrtenJahr, fahrtenPdf } from "../src/fahrtenbuch.js";

export async function fahrtenbuch({ graph, holeBereich, rechne, einstellungen, heute, jetzt, probe, ausgabe }) {
  if (!probe && heute.slice(5, 7) !== "01") return;
  const jahr = probe || String(Number(heute.slice(0, 4)) - 1);
  const kennung = `fahrten ${jahr}${probe ? " probe" : ""}`;
  if (!probe) {
    const da = await graph("GET", `/messages?$top=1&$select=id&$filter=singleValueExtendedProperties/any(ep: ep/id eq '${BERICHT}' and ep/value eq '${kennung}')`);
    if (da.value.length) return;
  }
  const termine = await holeBereich(`${jahr}-01-01`, `${jahr}-12-31`);
  const vorab = einstellungen.anbieter.flatMap((a) => erstelleAuswertung(termine, a, jetzt, einstellungen).fahrten);
  const strecken = await berechneStrecken(vorab, rechne, einstellungen);
  const f = fahrtenJahr(termine, jahr, jetzt, einstellungen, (v, n) => strecken.get(`${v} ${n}`));
  const pdf = fahrtenPdf(f, jahr, einstellungen, ladeJsPdf(), jetzt);
  await graph("POST", "/messages", {
    subject: `${probe ? "PROBE: " : ""}Fahrten ${jahr} fürs Finanzamt`,
    body: {
      contentType: "text",
      content: `Im Anhang alle Fahrten zu deinen Kursen ${jahr}: ${f.summe.fahrten} Fahrten an ${f.summe.tage} Tagen, zusammen ${String(f.summe.km).replace(".", ",")} km.\n\nNur ein Entwurf für dich, wird nicht gesendet.`,
    },
    singleValueExtendedProperties: [{ id: BERICHT, value: kennung }],
    attachments: [
      {
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: `Fahrten ${jahr}.pdf`,
        contentType: "application/pdf",
        contentBytes: Buffer.from(pdf.output("arraybuffer")).toString("base64"),
      },
    ],
  });
  ausgabe(`Fahrten ${jahr}: Entwurf angelegt (${f.summe.tage} Tage, ${f.summe.fahrten} Fahrten${f.summe.ohneKm ? `, ${f.summe.ohneKm} Tage ohne km` : ""})`);
}
