// Entwürfe für Michael: morgens Hinweise (fehlende Protokolle, Überschneidungen),
// sonntagabends die Wochenvorschau. Jeder Entwurf entsteht nur einmal (Kennung
// am Entwurf). Gesendet wird nichts. Die öffentliche Ausgabe enthält nur Zahlen.
import { BERICHT } from "./monatsbericht.js";
import { fehlendeProtokolle, ueberschneidungen, hinweisHtml, wochenHtml, plusTage } from "../src/hinweise.js";

async function schonDa(graph, kennung) {
  const da = await graph(
    "GET",
    `/messages?$top=1&$select=id&$filter=singleValueExtendedProperties/any(ep: ep/id eq '${BERICHT}' and ep/value eq '${kennung}')`
  );
  return da.value.length > 0;
}

const entwurf = (graph, kennung, betreff, html) =>
  graph("POST", "/messages", {
    subject: betreff,
    body: { contentType: "html", content: html },
    singleValueExtendedProperties: [{ id: BERICHT, value: kennung }],
  });

// jetzt = "JJJJ-MM-TTTHH:MM" deutsche Zeit.
export async function hinweise({ graph, holeBereich, einstellungen, jetzt, ausgabe, erzwingen = false }) {
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
      if (protokolle.length || paare.length) {
        await entwurf(
          graph,
          kennung,
          `${erzwingen ? "PROBE: " : ""}Kursplaner: ${protokolle.length + paare.length} Hinweise`,
          hinweisHtml({ protokolle, paare })
        );
      }
      ausgabe(`Hinweise ${heute}: ${protokolle.length} Protokolle fehlen, ${paare.length} Überschneidungen`);
    }
  }

  // Sonntag ab der eingestellten Stunde: Vorschau auf die kommende Woche.
  const sonntag = new Date(`${heute}T12:00:00Z`).getUTCDay() === 0;
  if (erzwingen || (sonntag && stunde >= e.wocheAb)) {
    const montag = sonntag ? plusTage(heute, 1) : plusTage(heute, (8 - new Date(`${heute}T12:00:00Z`).getUTCDay()) % 7 || 7);
    const kennung = `woche ${montag}${erzwingen ? " probe" : ""}`;
    if (erzwingen || !(await schonDa(graph, kennung))) {
      const termine = await holeBereich(montag, plusTage(montag, 6));
      await entwurf(
        graph,
        kennung,
        `${erzwingen ? "PROBE: " : ""}Kursplaner: Woche ab ${montag.slice(8, 10)}.${montag.slice(5, 7)}.`,
        wochenHtml(termine, montag, einstellungen)
      );
      ausgabe(`Wochenvorschau ab ${montag}: Entwurf angelegt`);
    }
  }
}
