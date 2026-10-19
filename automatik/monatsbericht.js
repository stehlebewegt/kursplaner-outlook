// Monatsbericht: am eingestellten Tag (einstellungen.monatsbericht.tag) legt
// die Automatik pro Anbieter einen E-Mail-Entwurf mit dem PDF des laufenden
// Monats an (bis heute geleistet, danach geplant), dazu ein Nachtrag zum
// Vormonat. Der Stand jedes Berichts wird am Entwurf gemerkt, damit der
// Nachtrag zeigen kann, was sich seitdem geändert hat.
// Verschickt wird nie, das macht Michael selbst.
// Braucht die Anwendungsberechtigung Mail.ReadWrite in Entra.
import { readFileSync } from "node:fs";
import { erstelleAuswertung, monatsZeitraum, stand, nachtrag } from "../src/auswertung.js";
import { erstellePdf } from "../web/bericht.js";
import { SIGNATUR_BETREFF, leseSignatur, htmlZuText, mitSignatur } from "../src/signatur.js";

// jsPDF ist eine Browser-Datei (UMD); im Repo gilt .js als ES-Modul, darum von Hand laden.
export function ladeJsPdf() {
  const modul = { exports: {} };
  new Function("module", "exports", readFileSync(new URL("../web/vendor/jspdf.umd.min.js", import.meta.url), "utf8"))(modul, modul.exports);
  return modul.exports.jsPDF;
}
const jsPDF = ladeJsPdf();
// Merkt am Entwurf, für welchen Anbieter und Monat er ist, damit er nur einmal entsteht.
export const BERICHT = "String {6f1c2a4e-3b7d-4c1a-9e55-8a2d9c0b7f31} Name KursplanerBericht";
// Gemerkter Stand (JSON, siehe stand() in src/auswertung.js).
export const STAND = "String {6f1c2a4e-3b7d-4c1a-9e55-8a2d9c0b7f31} Name KursplanerBerichtStand";

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
// weitere = zusätzliche Postfächer (z. B. gemeinsames Adressbuch), deren Kontakte mit durchsucht werden.
export async function holeKontakte(graph, kuerzel, weitere = []) {
  const filter = `$top=10&$select=givenName,displayName,emailAddresses&$filter=categories/any(c:c eq 'Monatsbericht ${kuerzel}')`;
  const basen = ["", ...weitere.map((p) => `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(p)}`)];
  try {
    const gefunden = [];
    for (const b of basen) {
      try {
        gefunden.push(...(await graph("GET", `${b}/contacts?${filter}`)).value);
        const ordner = (await graph("GET", `${b}/contactFolders?$top=50&$select=id`)).value;
        for (const o of ordner) gefunden.push(...(await graph("GET", `${b}/contactFolders/${o.id}/contacts?${filter}`)).value);
      } catch (e) {
        if (!b || !/ 404 /.test(e.message)) throw e; // unbekanntes Zusatzpostfach überspringen
      }
    }
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

// Gemerkter Stand des letzten Berichts (auch gesendet oder im Papierkorb), sonst null.
async function holeStand(graph, kennungen) {
  for (const k of kennungen) {
    const r = await graph(
      "GET",
      `/messages?$top=1&$select=id&$filter=singleValueExtendedProperties/any(ep: ep/id eq '${BERICHT}' and ep/value eq '${k}')` +
        `&$expand=singleValueExtendedProperties($filter=id eq '${STAND}')`
    );
    const wert = r.value[0]?.singleValueExtendedProperties?.[0]?.value;
    if (wert) {
      try {
        return JSON.parse(wert);
      } catch {
        return null;
      }
    }
  }
  return null;
}

// Strecken einmal pro Paar berechnen (Zuhause und Anbieter-Adressen).
export async function berechneStrecken(fahrten, rechne, einstellungen) {
  const ortVon = (k) => (k === "heim" ? null : einstellungen.anbieter.find((a) => a.kuerzel === k)?.ort);
  const strecken = new Map();
  if (!rechne) return strecken;
  for (const f of fahrten) {
    const paar = `${f.von} ${f.nach}`;
    if (strecken.has(paar) || ortVon(f.von) === undefined || ortVon(f.nach) === undefined) continue;
    try {
      strecken.set(paar, await rechne.strecke(ortVon(f.von), ortVon(f.nach)));
    } catch {
      strecken.set(paar, null);
    }
  }
  return strecken;
}

// probe = "JJJJ-MM": sofort für diesen Monat, Betreff mit "PROBE", zählt nicht als erledigt.
// Signatur aus dem Entwurf "Kursplaner-Signatur" (siehe src/signatur.js); ohne: Text wie bisher.
export async function holeSignatur(graph) {
  try {
    const { value } = await graph(
      "GET",
      `/mailFolders/drafts/messages?$filter=${encodeURIComponent(`subject eq '${SIGNATUR_BETREFF}'`)}&$select=body&$top=1`
    );
    const b = value[0]?.body;
    return (b && leseSignatur(b.contentType === "html" ? htmlZuText(b.content) : b.content)) || null;
  } catch {
    return null;
  }
}

export async function monatsberichte({ graph, holeBereich, rechne, einstellungen, heute, jetzt, probe, ausgabe, kontaktPostfaecher = [] }) {
  const mb = einstellungen.monatsbericht;
  let sig;
  if (!probe && Number(heute.slice(8, 10)) !== mb.tag) return;
  const monat = probe || heute.slice(0, 7);
  const vm = vormonat(`${monat}-01`);
  const zeitraum = monatsZeitraum(monat);
  const vZeitraum = monatsZeitraum(vm);
  const alle = await holeBereich(vZeitraum.von, zeitraum.bis);
  const im = (z) => alle.filter((t) => t.start.slice(0, 10) >= z.von && t.start.slice(0, 10) <= z.bis);
  const termine = im(zeitraum);
  const vTermine = im(vZeitraum);

  for (const anbieter of einstellungen.anbieter) {
    const kennung = `${anbieter.kuerzel} ${monat}${probe ? " probe" : ""}`;
    if (!probe) {
      const da = await graph(
        "GET",
        `/messages?$top=1&$select=id&$filter=singleValueExtendedProperties/any(ep: ep/id eq '${BERICHT}' and ep/value eq '${kennung}')`
      );
      if (da.value.length) continue;
    }
    const vorab = [erstelleAuswertung(termine, anbieter, jetzt, einstellungen), erstelleAuswertung(vTermine, anbieter, jetzt, einstellungen)];
    const strecken = await berechneStrecken(vorab.flatMap((x) => x.fahrten), rechne, einstellungen);
    const strecke = (von, nach) => strecken.get(`${von} ${nach}`);
    const a = erstelleAuswertung(termine, anbieter, jetzt, einstellungen, strecke);
    const alt = await holeStand(graph, [`${anbieter.kuerzel} ${vm}`, ...(probe ? [`${anbieter.kuerzel} ${vm} probe`] : [])]);
    const n = nachtrag(erstelleAuswertung(vTermine, anbieter, jetzt, einstellungen, strecke), alt, `${vm}-${String(mb.tag).padStart(2, "0")}`);
    n.name = vZeitraum.name;
    if (!a.zeilen.length && !n.zeilen.length) {
      ausgabe(`Monatsbericht ${kennung}: keine Kurse, kein Entwurf`);
      continue;
    }
    const pdf = erstellePdf(a, anbieter, zeitraum, jetzt, { teilnehmer: mb.teilnehmer, nachtrag: n, jsPDF });
    const kontakte = await holeKontakte(graph, anbieter.kuerzel, kontaktPostfaecher);
    const { an, anrede, vorlage } = anschreiben(kontakte || [], mb);
    const werte = { monat: zeitraum.name, vormonat: vZeitraum.name, anbieter: anbieter.name, anrede };
    werte.nachtrag = n.zeilen.length ? fuelle(mb.nachtragText || "", werte) : "";
    await graph("POST", "/messages", {
      subject: (probe ? "PROBE: " : "") + fuelle(mb.betreff, werte),
      body: mitSignatur(fuelle(vorlage, werte), (sig ??= await holeSignatur(graph))),
      toRecipients: an.map((address) => ({ emailAddress: { address } })),
      attachments: [
        {
          "@odata.type": "#microsoft.graph.fileAttachment",
          name: `Monatsauswertung ${anbieter.kuerzel} ${zeitraum.name}.pdf`,
          contentType: "application/pdf",
          contentBytes: Buffer.from(pdf.output("arraybuffer")).toString("base64"),
        },
      ],
      singleValueExtendedProperties: [
        { id: BERICHT, value: kennung },
        { id: STAND, value: JSON.stringify(stand(a, jetzt)) },
      ],
    });
    const ohneKm = [...strecken.values()].filter((s) => !s).length;
    const empf = kontakte === null ? "Kontakte-Berechtigung fehlt" : `${an.length} Empfänger`;
    const nt = n.zeilen.length ? `, Nachtrag ${n.zeilen.length} Kurse${alt ? ` (${n.summe.neu} neu, ${n.summe.entfallen} entfallen)` : " ohne alten Stand"}` : "";
    ausgabe(
      `Monatsbericht ${kennung}: Entwurf angelegt (${a.zeilen.length} Kurse, ${a.summe.geleistet} geleistet, ${a.summe.geplant} geplant, ${a.summe.fahrten} Fahrten${nt}, ${empf}${ohneKm ? `, ${ohneKm} Strecken ohne km` : ""})`
    );
  }
}

// Einmalig: Monatsabschlüsse (ohne Nachtrag) für mehrere Monate, pro Anbieter
// ein Entwurf ohne Empfänger mit einem PDF je Monat. bereich = "JJJJ-MM:JJJJ-MM".
export async function abschluesse({ graph, holeBereich, rechne, einstellungen, jetzt, bereich, ausgabe }) {
  const [von, bis] = bereich.split(":");
  const monate = [];
  for (let m = von; m <= bis; m = (([j, mm]) => (mm === 12 ? `${j + 1}-01` : `${j}-${String(mm + 1).padStart(2, "0")}`))(m.split("-").map(Number)))
    monate.push(m);
  const alle = await holeBereich(monatsZeitraum(von).von, monatsZeitraum(bis).bis);
  for (const anbieter of einstellungen.anbieter) {
    const anhaenge = [];
    let kurse = 0;
    for (const monat of monate) {
      const z = monatsZeitraum(monat);
      const termine = alle.filter((t) => t.start.slice(0, 10) >= z.von && t.start.slice(0, 10) <= z.bis);
      const vorab = erstelleAuswertung(termine, anbieter, jetzt, einstellungen);
      if (!vorab.zeilen.length) continue;
      const strecken = await berechneStrecken(vorab.fahrten, rechne, einstellungen);
      const a = erstelleAuswertung(termine, anbieter, jetzt, einstellungen, (v, n) => strecken.get(`${v} ${n}`));
      const pdf = erstellePdf(a, anbieter, z, jetzt, { teilnehmer: einstellungen.monatsbericht.teilnehmer, jsPDF });
      anhaenge.push({
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: `Monatsauswertung ${anbieter.kuerzel} ${monat} ${z.name}.pdf`,
        contentType: "application/pdf",
        contentBytes: Buffer.from(pdf.output("arraybuffer")).toString("base64"),
      });
      kurse += a.zeilen.length;
    }
    if (!anhaenge.length) {
      ausgabe(`Abschlüsse ${anbieter.kuerzel}: keine Kurse`);
      continue;
    }
    await graph("POST", "/messages", {
      subject: `ABSCHLÜSSE ${anbieter.kuerzel} ${von} bis ${bis} (nur zum Speichern)`,
      body: { contentType: "text", content: `Monatsabschlüsse ${anbieter.name}, ${anhaenge.length} PDFs. Zum Abspeichern, nicht zum Versenden.` },
      attachments: anhaenge,
    });
    ausgabe(`Abschlüsse ${anbieter.kuerzel}: Entwurf mit ${anhaenge.length} PDFs (${kurse} Kurse)`);
  }
}
