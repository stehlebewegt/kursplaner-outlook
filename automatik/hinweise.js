// Für Michael: morgens Hinweise als Entwurf (fehlende Protokolle, Überschneidungen),
// sonntagabends die Wochenvorschau als ungelesene Mail im Posteingang. Jede
// entsteht nur einmal (Kennung an der Nachricht). Gesendet wird nichts. Die
// öffentliche Ausgabe enthält nur Zahlen.
import { BERICHT, ladeJsPdf } from "./monatsbericht.js";
import { wochenPdf } from "../src/wochenPdf.js";
import { fehlendeProtokolle, ueberschneidungen, hinweisHtml, wochenHtml, plusTage } from "../src/hinweise.js";

async function schonDa(graph, kennung) {
  const da = await graph(
    "GET",
    `/messages?$top=1&$select=id&$filter=singleValueExtendedProperties/any(ep: ep/id eq '${BERICHT}' and ep/value eq '${kennung}')`
  );
  return da.value.length > 0;
}

const entwurf = (graph, kennung, betreff, html, anhaenge = []) =>
  graph("POST", "/messages", {
    subject: betreff,
    body: { contentType: "html", content: html },
    singleValueExtendedProperties: [{ id: BERICHT, value: kennung }],
    attachments: anhaenge,
  });
// MAPI-Nachrichtenflags: 1 = gelesen, ohne „nicht gesendet“. So liegt die
// Nachricht als normale Mail im Posteingang statt als Entwurf; danach auf
// ungelesen gestellt, damit sie auffällt. Braucht nur Mail.ReadWrite.
export const NACHRICHTEN_FLAGS = "Integer 0x0E07";
export async function inPosteingang(graph, kennung, betreff, html, anhaenge = [], postfach) {
  const an = postfach ? { emailAddress: { name: "Kursplaner", address: postfach } } : undefined;
  const mail = await graph("POST", "/mailFolders/inbox/messages", {
    subject: betreff,
    body: { contentType: "html", content: html },
    ...(an && { from: an, toRecipients: [an] }),
    singleValueExtendedProperties: [
      { id: BERICHT, value: kennung },
      { id: NACHRICHTEN_FLAGS, value: "1" },
    ],
    attachments: anhaenge,
  });
  await graph("PATCH", `/messages/${mail.id}`, { isRead: false });
}
const pdfAnhang = (name, doc) => ({
  "@odata.type": "#microsoft.graph.fileAttachment",
  name,
  contentType: "application/pdf",
  contentBytes: Buffer.from(doc.output("arraybuffer")).toString("base64"),
});

// jetzt = "JJJJ-MM-TTTHH:MM" deutsche Zeit.
export async function hinweise({ graph, holeBereich, einstellungen, jetzt, ausgabe, erzwingen = false, holeMails = async () => [], postfach }) {
  const e = einstellungen.hinweise;
  if (!e) return;
  const heute = jetzt.slice(0, 10);
  const stunde = Number(jetzt.slice(11, 13));

  // Nur in einem Fenster am Morgen suchen; ohne Hinweise entsteht kein Entwurf,
  // dann wird in diesem Fenster bei jedem Lauf erneut (kurz) geprüft.
  if (erzwingen || (stunde >= e.morgensAb && stunde < e.morgensAb + 2)) {
    const kennung = `hinweise ${heute}${erzwingen ? " probe" : ""}`;
    if (erzwingen || !(await schonDa(graph, kennung))) {
      // Protokolle: Kurse der letzten Tage, die schon vorbei sind.
      const zurueck = await holeBereich(plusTage(heute, -e.protokollTage), plusTage(heute, -1));
      const ohne = fehlendeProtokolle(zurueck, jetzt, einstellungen);
      for (const t of ohne) {
        const voll = await graph("GET", `/events/${t.id}?$select=body`);
        t.text = voll.body?.content || "";
      }
      // Body kommt als HTML, wenn er so gespeichert ist: Tags entfernen.
      for (const t of ohne) t.text = t.text.replace(/<br\s*\/?>/gi, "\n").replace(/<\/(p|div)>/gi, "\n").replace(/<[^>]+>/g, "").replace(/&nbsp;/g, " ");
      const protokolle = fehlendeProtokolle(ohne, jetzt, einstellungen);
      const paare = ueberschneidungen(await holeBereich(heute, plusTage(heute, e.ueberschneidungTage)), einstellungen);
      let mails = [];
      try {
        mails = await holeMails();
      } catch (err) {
        ausgabe(`Hinweise: Mails nicht geprüft (${err.message.slice(0, 60)})`);
      }
      const anzahl = protokolle.length + paare.length + mails.length;
      if (anzahl) {
        await entwurf(graph, kennung, `${erzwingen ? "PROBE: " : ""}Kursplaner: ${anzahl} Hinweise`, hinweisHtml({ protokolle, paare, mails }));
      }
      ausgabe(`Hinweise ${heute}: ${protokolle.length} Protokolle fehlen, ${paare.length} Überschneidungen, ${mails.length} Mails offen`);
    }
  }

  // Sonntag ab der eingestellten Stunde: Vorschau auf die kommende Woche.
  const sonntag = new Date(`${heute}T12:00:00Z`).getUTCDay() === 0;
  if (erzwingen || (sonntag && stunde >= e.wocheAb)) {
    const montag = sonntag ? plusTage(heute, 1) : plusTage(heute, (8 - new Date(`${heute}T12:00:00Z`).getUTCDay()) % 7 || 7);
    const kennung = `woche ${montag}${erzwingen ? " probe" : ""}`;
    if (erzwingen || !(await schonDa(graph, kennung))) {
      const termine = await holeBereich(montag, plusTage(montag, 6));
      await inPosteingang(
        graph,
        kennung,
        `${erzwingen ? "PROBE: " : ""}Kursplaner: Woche ab ${montag.slice(8, 10)}.${montag.slice(5, 7)}.`,
        wochenHtml(termine, montag, einstellungen),
        [pdfAnhang(`Woche ab ${montag.slice(8, 10)}.${montag.slice(5, 7)}.pdf`, wochenPdf(termine, montag, einstellungen, ladeJsPdf()))],
        postfach
      );
      ausgabe(`Wochenvorschau ab ${montag}: im Posteingang`);
    }
  }
}
