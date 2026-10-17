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

// Ansprechpartner aus Outlook-Kontakten mit Kategorie "Monatsbericht <Kürzel>".
// Ohne Berechtigung (Contacts.Read) oder Kontakt: kein Empfänger, "zusammen".
export function anschreiben(kontakte, mb) {
  const leute = kontakte
    .map((k) => ({ name: (k.givenName || k.displayName || "").trim(), mail: k.emailAddresses?.[0]?.address }))
    .filter((k) => k.mail);
  const namen = leute.map((k) => k.name).filter(Boolean);
  const anrede = namen.length ? namen.join(namen.length === 2 ? " und " : ", ").replace(/, ([^,]*)$/, " und $1") : "zusammen";
  return { an: leute.map((k) => k.mail), anrede, vorlage: leute.length === 1 && namen.length === 1 ? mb.text : mb.textGruppe };
}

// Sucht im Standard-Kontaktordner und in allen Unterordnern.
async function holeKontakte(graph, kuerzel) {
  const filter = `$top=10&$select=givenName,displayName,emailAddresses&$filter=categories/any(c:c eq 'Monatsbericht ${kuerzel}')`;
  try {
    const gefunden = (await graph("GET", `/contacts?${filter}`)).value;
    const ordner = (await graph("GET", "/contactFolders?$top=50&$select=id")).value;
    for (const o of ordner) gefunden.push(...(await graph("GET", `/contactFolders/${o.id}/contacts?${filter}`)).value);
    const gesehen = new Set();
    return gefunden.filter((k) => {
      const m = k.emailAddresses?.[0]?.address?.toLowerCase();
      return m && !gesehen.has(m) && gesehen.add(m);
    });
  } catch (e) {
    if (/ 403 /.test(e.message)) return null;
    throw e;
  }
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
    const kontakte = await holeKontakte(graph, anbieter.kuerzel);
    const { an, anrede, vorlage } = anschreiben(kontakte || [], mb);
    const werte = { monat: zeitraum.name, anbieter: anbieter.name, anrede };
    await graph("POST", "/messages", {
      subject: (probe ? "PROBE: " : "") + fuelle(mb.betreff, werte),
      body: { contentType: "text", content: fuelle(vorlage, werte) },
      toRecipients: an.map((address) => ({ emailAddress: { address } })),
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
    const empf = kontakte === null ? "Kontakte-Berechtigung fehlt" : `${an.length} Empfänger`;
    ausgabe(`Monatsbericht ${kennung}: Entwurf angelegt (${a.zeilen.length} Kurse, ${a.summe.fahrten} Fahrten, ${empf}${ohneKm ? `, ${ohneKm} Strecken ohne km` : ""})`);
  }
}
