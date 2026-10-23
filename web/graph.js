// Zugriff auf Michaels Kalender über Microsoft Graph.
// Anmeldung über "Nested App Authentication" (NAA): Outlook reicht die
// Anmeldung an das Add-in weiter, kein extra Login-Fenster nötig.
import { config } from "./config.js";
import { einstellungen } from "../src/einstellungen.js";
import { termineFuerTag, erinnerungsFelder } from "../src/planer.js";
import { SIGNATUR_BETREFF, leseSignatur, htmlZuText, mitSignatur } from "../src/signatur.js";
import { holeEigene, speichereEigene } from "../src/eigeneEinstellungen.js";

const GRAPH = "https://graph.microsoft.com/v1.0";
// Unsichtbare Markierung an jedem Block, den das Add-in anlegt.
export const MARKIERUNG = "String {6f1c2a4e-3b7d-4c1a-9e55-8a2d9c0b7f31} Name KursplanerGeneriert";
// Fahrzeit in Minuten an privaten Terminen (Hin- und Rückfahrt).
export const FAHRZEIT = "String {6f1c2a4e-3b7d-4c1a-9e55-8a2d9c0b7f31} Name KursplanerFahrzeit";

let msalApp;
// Das Kursplaner-Fenster (app/) meldet sich selbst an und setzt hier seine
// Token-Quelle; im Add-in bleibt es bei NAA.
let tokenQuelle = null;
export const nutzeAnmeldung = (quelle) => (tokenQuelle = quelle);

async function token() {
  if (tokenQuelle) return tokenQuelle();
  if (!msalApp) {
    msalApp = await msal.createNestablePublicClientApplication({
      auth: { clientId: config.clientId, authority: `https://login.microsoftonline.com/${config.tenantId}` },
    });
  }
  const anfrage = { scopes: config.scopes };
  try {
    return (await msalApp.acquireTokenSilent(anfrage)).accessToken;
  } catch {
    return (await msalApp.acquireTokenPopup(anfrage)).accessToken;
  }
}

const warte = (ms) => new Promise((r) => setTimeout(r, ms));

async function graph(methode, pfad, body, versuch = 0) {
  const antwort = await fetch(pfad.startsWith("http") ? pfad : GRAPH + pfad, {
    method: methode,
    headers: {
      Authorization: `Bearer ${await token()}`,
      "Content-Type": "application/json",
      Prefer: `outlook.timezone="${config.zeitzone}", outlook.body-content-type="text"`,
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  // Gebremst: kurz warten (2, 4, 8, 16 Sekunden) und neu versuchen.
  if ((antwort.status === 429 || antwort.status === 503) && versuch < 4) {
    const sekunden = Number(antwort.headers.get("Retry-After")) || 2 ** (versuch + 1);
    await warte(sekunden * 1000);
    return graph(methode, pfad, body, versuch + 1);
  }
  if (!antwort.ok) {
    const text = await antwort.text();
    throw new Error(`Graph ${methode} ${pfad}: ${antwort.status} ${text.slice(0, 200)}`);
  }
  return antwort.status === 204 ? null : antwort.json();
}

// Graph-Termin in das Format von planer.js umwandeln.
function zuTermin(e) {
  const props = e.singleValueExtendedProperties || [];
  const prop = props.find((p) => p.id === MARKIERUNG);
  // null = nie eingestellt, 0 = bewusst keine Fahrzeit.
  const wert = props.find((p) => p.id === FAHRZEIT)?.value;
  // "" = bitte (neu) berechnen, -1 = Adresse nicht gefunden.
  const fahrzeit = wert === undefined || wert === "" ? null : Number(wert) || 0;
  return {
    id: e.id,
    titel: e.subject || "",
    start: e.start.dateTime.slice(0, 16),
    ende: e.end.dateTime.slice(0, 16),
    kategorien: e.categories || [],
    ort: e.location?.displayName || "",
    generiert: prop?.value === "1",
    fahrzeit,
  };
}

const AUSWAHL =
  "$select=subject,start,end,categories,location,showAs" +
  `&$expand=singleValueExtendedProperties($filter=id eq '${MARKIERUNG}' or id eq '${FAHRZEIT}')`;

// Alle Termine eines Tages im eigenen Kalender (plus Übernachtungs-Termine
// vom Vortag und Folgetag, siehe termineFuerTag).
export async function holeTag(datum) {
  const tag = (n) => new Date(Date.parse(`${datum}T12:00:00Z`) + n * 86400000).toISOString().slice(0, 10);
  let url = `/me/calendarView?startDateTime=${tag(-1)}T00:00:00&endDateTime=${tag(1)}T23:59:59&$top=100&${AUSWAHL}`;
  const termine = [];
  while (url) {
    const seite = await graph("GET", url);
    termine.push(...seite.value.map(zuTermin));
    url = seite["@odata.nextLink"];
  }
  return termineFuerTag(termine, datum, einstellungen);
}

// Kurse mit Termintext über einen Zeitraum, für "letzte Themen" im Protokoll.
export async function holeVerlauf(von, bis) {
  let url =
    `/me/calendarView?startDateTime=${von}T00:00:00&endDateTime=${bis}T23:59:59&$top=200` +
    `&$select=subject,start,end,categories,location,body`;
  const termine = [];
  while (url) {
    const seite = await graph("GET", url);
    termine.push(...seite.value.map((e) => ({ ...zuTermin(e), text: e.body?.content || "" })));
    url = seite["@odata.nextLink"];
  }
  return termine;
}

export async function holeTermin(id) {
  const e = await graph("GET", `/me/events/${id}?${AUSWAHL.replace("showAs", "showAs,body")}`);
  const t = zuTermin(e);
  const kalender = await findeKalender(t);
  return { ...t, text: e.body?.content || "", kalender: kalender.name, eigenerKalender: kalender.isDefaultCalendar };
}

// In welchem Kalender liegt der Termin? Bei Serienterminen liefert Graph über
// /events/{id}/calendar nicht zuverlässig den richtigen, deshalb schauen wir
// in jedem Kalender an dem Tag nach.
let kalenderListe;

async function findeKalender(t) {
  kalenderListe ??= (await graph("GET", "/me/calendars?$select=id,name,isDefaultCalendar&$top=50")).value;
  const datum = t.start.slice(0, 10);
  // Nur eigener Kalender (zuerst) und Anbieter-Kalender, nacheinander, und
  // aufhören, sobald die ID gefunden ist: zu viele Abfragen bremst Microsoft
  // aus (Fehler 429).
  const relevant = kalenderListe
    .filter((k) => k.isDefaultCalendar || einstellungen.anbieter.some((a) => a.kalender === k.name))
    .sort((a, b) => b.isDefaultCalendar - a.isDefaultCalendar);
  const tage = [];
  for (const k of relevant) {
    const { value } = await graph(
      "GET",
      `/me/calendars/${k.id}/calendarView?startDateTime=${datum}T00:00:00&endDateTime=${datum}T23:59:59&$select=subject,start&$top=100`
    );
    if (value.some((e) => e.id === t.id)) return k;
    tage.push({ k, termine: value });
  }
  // Gleiche Titel gibt es oft in zwei Kalendern (Anbieter-Kalender und eigene
  // Kopie), deshalb der Titel nur als Notlösung.
  const gleich = (e) => (e.subject || "").trim() === t.titel.trim() && e.start.dateTime.slice(0, 16) === t.start;
  const perTitel = tage.filter(({ termine }) => termine.some(gleich));
  if (perTitel.length === 1) return perTitel[0].k;
  return kalenderListe.find((k) => k.isDefaultCalendar) || { name: "?", isDefaultCalendar: false };
}

// E-Mail-Entwurf mit PDF im Anhang, landet im Ordner "Entwürfe".
export async function legeEntwurfAn({ an, betreff, text, dateiname, pdfBase64 }) {
  const m = await graph("POST", "/me/messages", {
    subject: betreff,
    body: mitSignatur(text, await holeSignatur()),
    toRecipients: an ? [{ emailAddress: { address: an } }] : [],
    attachments: [
      { "@odata.type": "#microsoft.graph.fileAttachment", name: dateiname, contentType: "application/pdf", contentBytes: pdfBase64 },
    ],
  });
  return m.id;
}

export async function legeAn(t, { generiert, text, html, frei } = {}) {
  const body = {
    subject: t.titel,
    start: { dateTime: `${t.start}:00`, timeZone: config.zeitzone },
    end: { dateTime: `${t.ende}:00`, timeZone: config.zeitzone },
    location: { displayName: t.ort || "" },
    categories: t.kategorien || [],
    showAs: frei ? "free" : "busy",
    ...erinnerungsFelder(t.erinnerung),
  };
  if (html) body.body = { contentType: "html", content: html };
  else if (text) body.body = { contentType: "text", content: text };
  if (generiert) body.singleValueExtendedProperties = [{ id: MARKIERUNG, value: "1" }];
  return zuTermin(await graph("POST", "/me/events", body));
}

// Übernimmt einen von Hand angelegten Block: markiert ihn als generiert.
export async function markiere(id) {
  await graph("PATCH", `/me/events/${id}`, { singleValueExtendedProperties: [{ id: MARKIERUNG, value: "1" }] });
}

export async function setzeFahrzeit(id, minuten) {
  await graph("PATCH", `/me/events/${id}`, { singleValueExtendedProperties: [{ id: FAHRZEIT, value: String(minuten) }] });
}

export async function aendere(id, felder) {
  await graph("PATCH", `/me/events/${id}`, felder);
}

// Termin auf neue Zeit legen (start/ende "JJJJ-MM-TTTHH:MM", Ortszeit).
export async function verschiebe(id, start, ende) {
  await graph("PATCH", `/me/events/${id}`, {
    start: { dateTime: `${start}:00`, timeZone: config.zeitzone },
    end: { dateTime: `${ende}:00`, timeZone: config.zeitzone },
  });
}

export async function loesche(id) {
  await graph("DELETE", `/me/events/${id}`);
}

// Kurse aller Anbieter-Kalender im Zeitraum (mit Text und Kalendername),
// für „Vertretung übernehmen“ im Kursplaner-Fenster.
export async function holeAnbieterKurse(von, bis) {
  kalenderListe ??= (await graph("GET", "/me/calendars?$select=id,name,isDefaultCalendar&$top=50")).value;
  const kurse = [];
  for (const a of einstellungen.anbieter) {
    const k = kalenderListe.find((k) => k.name === a.kalender);
    if (!k) continue;
    let url =
      `/me/calendars/${k.id}/calendarView?startDateTime=${von}T00:00:00&endDateTime=${bis}T23:59:59&$top=200` +
      `&$select=subject,start,end,categories,location,body,isAllDay`;
    while (url) {
      const seite = await graph("GET", url);
      for (const e of seite.value.filter((e) => !e.isAllDay))
        kurse.push({ ...zuTermin(e), text: e.body?.content || "", kalender: a.kalender });
      url = seite["@odata.nextLink"];
    }
  }
  return kurse.sort((a, b) => (a.start < b.start ? -1 : 1));
}

// Adresse aus dem Signatur-Entwurf (für km im Kursplaner-Fenster; bleibt im Browser).
export async function heimAdresse() {
  return (await holeSignatur())?.Adresse || null;
}

// --- Neuer Kurs im Anbieter-Kalender ---
// name null = eigener Kalender.
async function anbieterKalender(name) {
  kalenderListe ??= (await graph("GET", "/me/calendars?$select=id,name,isDefaultCalendar&$top=50")).value;
  const k = kalenderListe.find((k) => (name ? k.name === name : k.isDefaultCalendar));
  if (!k) throw new Error(`Den Kalender „${name}“ gibt es nicht.`);
  return k;
}

// Titel der nächsten Wochen im Anbieter-Kalender, als Vorschläge fürs Formular.
export async function holeKursTitel(kalenderName, von, bis) {
  const k = await anbieterKalender(kalenderName);
  const { value } = await graph(
    "GET",
    `/me/calendars/${k.id}/calendarView?startDateTime=${von}T00:00:00&endDateTime=${bis}T23:59:59&$select=subject,start,end&$top=300`
  );
  return value.map((e) => ({ titel: e.subject || "", start: e.start.dateTime.slice(0, 16), ende: e.end.dateTime.slice(0, 16) }));
}

export async function legeKursAn(kalenderName, ereignis) {
  const k = await anbieterKalender(kalenderName);
  return graph("POST", `/me/calendars/${k.id}/events`, ereignis);
}

// Einfacher E-Mail-Entwurf ohne Anhang (Vorlagen zum Kurs).
export async function legeMailEntwurfAn({ an, betreff, text }) {
  const m = await graph("POST", "/me/messages", {
    subject: betreff,
    body: mitSignatur(text, await holeSignatur()),
    toRecipients: an ? [{ emailAddress: { address: an } }] : [],
  });
  return m.id;
}

// Signatur aus dem Entwurf "Kursplaner-Signatur" (siehe src/signatur.js).
// Fehlt er oder klappt das Lesen nicht, geht der Entwurf ohne Signatur raus.
let signatur;
async function holeSignatur() {
  if (signatur !== undefined) return signatur;
  try {
    const { value } = await graph(
      "GET",
      `/me/mailFolders/drafts/messages?$filter=${encodeURIComponent(`subject eq '${SIGNATUR_BETREFF}'`)}&$select=body&$top=1`
    );
    const b = value[0]?.body;
    signatur = b ? leseSignatur(b.contentType === "html" ? htmlZuText(b.content) : b.content) : null;
  } catch {
    signatur = null;
  }
  return signatur;
}

// Eigene Einstellungen im Postfach (siehe src/eigeneEinstellungen.js).
export const leseEinstellungen = () => holeEigene(graph, "/me");
export const speichereEinstellungen = (daten) => speichereEigene(graph, daten, "/me");
