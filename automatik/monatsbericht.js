// Monatsbericht: am eingestellten Tag (einstellungen.monatsbericht.tag) legt
// die Automatik pro Anbieter einen E-Mail-Entwurf mit dem PDF des Vormonats
// an. Verschickt wird nie, das macht Michael selbst.
// Braucht die Anwendungsberechtigung Mail.ReadWrite in Entra.
import { readFileSync } from "node:fs";
import { erstelleAuswertung, monatsZeitraum } from "../src/auswertung.js";
import { erstellePdf } from "../web/bericht.js";

// jsPDF ist eine Browser-Datei (UMD); im Repo gilt .js als ES-Modul, darum von Hand laden.
export function ladeJsPdf() {
  const modul = { exports: {} };
  new Function("module", "exports", readFileSync(new URL("../web/vendor/jspdf.umd.min.js", import.meta.url), "utf8"))(modul, modul.exports);
  return modul.exports.jsPDF;
}
const jsPDF = ladeJsPdf();
// Merkt am Entwurf, für welchen Anbieter und Monat er ist, damit er nur einmal entsteht.
export const BERICHT = "String {6f1c2a4e-3b7d-4c1a-9e55-8a2d9c0b7f31} Name KursplanerBericht";

export function vormonat(heute) {
  const [j, m] = heute.split("-").map(Number);
  return m === 1 ? `${j - 1}-12` : `${j}-${String(m - 1).padStart(2, "0")}`;
}

const fuelle = (vorlage, werte) => vorlage.replace(/\{(\w+)\}/g, (_, k) => werte[k] ?? "");

// probe = "JJJJ-MM": sofort für diesen Monat, Betreff mit "PROBE", zählt nicht als erledigt.
export async function monatsberichte({ graph, holeBereich, rechne, einstellungen, heute, jetzt, probe, ausgabe }) {
  const mb = einstellungen.monatsbericht;
  if (!probe && Number(heute.slice(8, 10)) !== mb.tag) return;
  const monat = probe || vormonat(heute);
  const zeitraum = monatsZeitraum(monat);
  const termine = await holeBereich(zeitraum.von, zeitraum.bis);

  for (const anbieter of einstellungen.anbieter) {
    const kennung = `${anbieter.kuerzel} ${monat}${probe ? " probe" : ""}`;
    if (!probe) {
      const da = await graph(
        "GET",
        `/messages?$top=1&$select=id&$filter=singleValueExtendedProperties/any(ep: ep/id eq '${BERICHT}' and ep/value eq '${kennung}')`
      );
      if (da.value.length) continue;
    }
    // Strecken einmal pro Paar berechnen (Zuhause und Anbieter-Adressen).
    const ortVon = (k) => (k === "heim" ? null : einstellungen.anbieter.find((a) => a.kuerzel === k)?.ort);
    const vorab = erstelleAuswertung(termine, anbieter, jetzt, einstellungen);
    const strecken = new Map();
    if (rechne) {
      for (const f of vorab.fahrten) {
        const paar = `${f.von} ${f.nach}`;
        if (strecken.has(paar) || ortVon(f.von) === undefined || ortVon(f.nach) === undefined) continue;
        try {
          strecken.set(paar, await rechne.strecke(ortVon(f.von), ortVon(f.nach)));
        } catch {
          strecken.set(paar, null);
        }
      }
    }
    const a = erstelleAuswertung(termine, anbieter, jetzt, einstellungen, (von, nach) => strecken.get(`${von} ${nach}`));
    if (!a.zeilen.length) {
      ausgabe(`Monatsbericht ${kennung}: keine Kurse, kein Entwurf`);
      continue;
    }
    const pdf = erstellePdf(a, anbieter, zeitraum, jetzt, { teilnehmer: mb.teilnehmer, jsPDF });
    const werte = { monat: zeitraum.name, anbieter: anbieter.name };
    await graph("POST", "/messages", {
      subject: (probe ? "PROBE: " : "") + fuelle(mb.betreff, werte),
      body: { contentType: "text", content: fuelle(mb.text, werte) },
      attachments: [
        {
          "@odata.type": "#microsoft.graph.fileAttachment",
          name: `Monatsauswertung ${anbieter.kuerzel} ${zeitraum.name}.pdf`,
          contentType: "application/pdf",
          contentBytes: Buffer.from(pdf.output("arraybuffer")).toString("base64"),
        },
      ],
      singleValueExtendedProperties: [{ id: BERICHT, value: kennung }],
    });
    const ohneKm = [...strecken.values()].filter((s) => !s).length;
    ausgabe(`Monatsbericht ${kennung}: Entwurf angelegt (${a.zeilen.length} Kurse, ${a.summe.fahrten} Fahrten${ohneKm ? `, ${ohneKm} Strecken ohne km` : ""})`);
  }
}
