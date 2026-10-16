// Zugriff auf Michaels Kalender über Microsoft Graph.
// Anmeldung über "Nested App Authentication" (NAA): Outlook reicht die
// Anmeldung an das Add-in weiter, kein extra Login-Fenster nötig.
import { config } from "./config.js";

const GRAPH = "https://graph.microsoft.com/v1.0";
// Unsichtbare Markierung an jedem Block, den das Add-in anlegt.
export const MARKIERUNG = "String {6f1c2a4e-3b7d-4c1a-9e55-8a2d9c0b7f31} Name KursplanerGeneriert";

let msalApp;

async function token() {
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

async function graph(methode, pfad, body) {
  const antwort = await fetch(pfad.startsWith("http") ? pfad : GRAPH + pfad, {
    method: methode,
    headers: {
      Authorization: `Bearer ${await token()}`,
      "Content-Type": "application/json",
      Prefer: `outlook.timezone="${config.zeitzone}", outlook.body-content-type="text"`,
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (!antwort.ok) {
    const text = await antwort.text();
    throw new Error(`Graph ${methode} ${pfad}: ${antwort.status} ${text.slice(0, 200)}`);
  }
  return antwort.status === 204 ? null : antwort.json();
}

// Graph-Termin in das Format von planer.js umwandeln.
function zuTermin(e) {
  const prop = (e.singleValueExtendedProperties || []).find((p) => p.id === MARKIERUNG);
  return {
    id: e.id,
    titel: e.subject || "",
    start: e.start.dateTime.slice(0, 16),
    ende: e.end.dateTime.slice(0, 16),
    kategorien: e.categories || [],
    ort: e.location?.displayName || "",
    generiert: prop?.value === "1",
  };
}

const AUSWAHL =
  "$select=subject,start,end,categories,location,showAs" +
  `&$expand=singleValueExtendedProperties($filter=id eq '${MARKIERUNG}')`;

// Alle Termine eines Tages im eigenen Kalender.
export async function holeTag(datum) {
  let url = `/me/calendarView?startDateTime=${datum}T00:00:00&endDateTime=${datum}T23:59:59&$top=100&${AUSWAHL}`;
  const termine = [];
  while (url) {
    const seite = await graph("GET", url);
    termine.push(...seite.value.map(zuTermin));
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
async function findeKalender(t) {
  const { value: kalender } = await graph("GET", "/me/calendars?$select=id,name,isDefaultCalendar&$top=50");
  const datum = t.start.slice(0, 10);
  const tage = await Promise.all(
    kalender.map(async (k) => ({
      k,
      termine: (
        await graph(
          "GET",
          `/me/calendars/${k.id}/calendarView?startDateTime=${datum}T00:00:00&endDateTime=${datum}T23:59:59&$select=subject,start&$top=100`
        )
      ).value,
    }))
  );
  // Erst über die ID suchen. Gleiche Titel gibt es oft in zwei Kalendern
  // (Anbieter-Kalender und eigene Kopie), deshalb der Titel nur als Notlösung.
  const perId = tage.find(({ termine }) => termine.some((e) => e.id === t.id));
  if (perId) return perId.k;
  const gleich = (e) => (e.subject || "").trim() === t.titel.trim() && e.start.dateTime.slice(0, 16) === t.start;
  const perTitel = tage.filter(({ termine }) => termine.some(gleich));
  if (perTitel.length === 1) return perTitel[0].k;
  return kalender.find((k) => k.isDefaultCalendar) || { name: "?", isDefaultCalendar: false };
}

export async function legeAn(t, { generiert, text, frei } = {}) {
  const body = {
    subject: t.titel,
    start: { dateTime: `${t.start}:00`, timeZone: config.zeitzone },
    end: { dateTime: `${t.ende}:00`, timeZone: config.zeitzone },
    location: { displayName: t.ort || "" },
    categories: t.kategorien || [],
    showAs: frei ? "free" : "busy",
    isReminderOn: false,
  };
  if (text) body.body = { contentType: "text", content: text };
  if (generiert) body.singleValueExtendedProperties = [{ id: MARKIERUNG, value: "1" }];
  return zuTermin(await graph("POST", "/me/events", body));
}

// Übernimmt einen von Hand angelegten Block: markiert ihn als generiert.
export async function markiere(id) {
  await graph("PATCH", `/me/events/${id}`, { singleValueExtendedProperties: [{ id: MARKIERUNG, value: "1" }] });
}

export async function aendere(id, felder) {
  await graph("PATCH", `/me/events/${id}`, felder);
}

export async function loesche(id) {
  await graph("DELETE", `/me/events/${id}`);
}
