// Startet die Automatik (läuft in GitHub Actions, siehe .github/workflows).
// Anmeldung als App mit Geheimnis, nicht als Michael: braucht die
// Anwendungsberechtigung Calendars.ReadWrite in Entra.
//
// Umgebungsvariablen (als GitHub-Secrets hinterlegt, nie ins Repo):
//   KURSPLANER_SECRET   Geheimer Clientschlüssel der Entra-App
//   KURSPLANER_POSTFACH E-Mail-Adresse des Kalenders
//   SCHREIBEN=1         Änderungen wirklich ausführen (sonst nur anzeigen)
//   TAGE                wie viele Tage ab heute (Standard 56 = 8 Wochen)
//
// Die Ausgabe ist öffentlich lesbar, deshalb nur Datum, Uhrzeit und Anzahl,
// keine Titel.
import { config } from "../web/config.js";
import { MARKIERUNG } from "../web/graph.js";
import { einstellungen } from "../src/einstellungen.js";
import { planeZeitraum, naechsteTage } from "./planen.js";

const { KURSPLANER_SECRET: secret, KURSPLANER_POSTFACH: postfach } = process.env;
if (!secret || !postfach) throw new Error("KURSPLANER_SECRET oder KURSPLANER_POSTFACH fehlt.");
const schreiben = process.env.SCHREIBEN === "1";
const anzahlTage = Number(process.env.TAGE) || 56;

async function holeToken() {
  const antwort = await fetch(`https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`, {
    method: "POST",
    body: new URLSearchParams({
      client_id: config.clientId,
      client_secret: secret,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials",
    }),
  });
  if (!antwort.ok) throw new Error(`Anmeldung fehlgeschlagen: ${antwort.status} ${(await antwort.text()).slice(0, 300)}`);
  return (await antwort.json()).access_token;
}

const token = await holeToken();
const basis = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(postfach)}`;
const warte = (ms) => new Promise((r) => setTimeout(r, ms));

async function graph(methode, pfad, body, versuch = 0) {
  const antwort = await fetch(pfad.startsWith("http") ? pfad : basis + pfad, {
    method: methode,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      Prefer: `outlook.timezone="${config.zeitzone}", outlook.body-content-type="text"`,
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  if ((antwort.status === 429 || antwort.status === 503) && versuch < 4) {
    await warte((Number(antwort.headers.get("Retry-After")) || 2 ** (versuch + 1)) * 1000);
    return graph(methode, pfad, body, versuch + 1);
  }
  if (!antwort.ok) throw new Error(`Graph ${methode}: ${antwort.status} ${(await antwort.text()).slice(0, 300)}`);
  return antwort.status === 204 ? null : antwort.json();
}

function zuTermin(e) {
  const prop = (e.singleValueExtendedProperties || []).find((p) => p.id === MARKIERUNG);
  return {
    id: e.id,
    titel: e.subject || "",
    start: e.start.dateTime.slice(0, 16),
    ende: e.end.dateTime.slice(0, 16),
    kategorien: e.categories || [],
    ort: e.location?.displayName || "",
    showAs: e.showAs,
    generiert: prop?.value === "1",
  };
}

const api = {
  async holeBereich(von, bis) {
    let url =
      `/calendarView?startDateTime=${von}T00:00:00&endDateTime=${bis}T23:59:59&$top=200` +
      `&$select=subject,start,end,categories,location,showAs,isAllDay` +
      `&$expand=singleValueExtendedProperties($filter=id eq '${MARKIERUNG}')`;
    const termine = [];
    while (url) {
      const seite = await graph("GET", url);
      termine.push(...seite.value.filter((e) => !e.isAllDay).map(zuTermin));
      url = seite["@odata.nextLink"];
    }
    return termine;
  },
  async legeAn(t) {
    await graph("POST", "/events", {
      subject: t.titel,
      start: { dateTime: `${t.start}:00`, timeZone: config.zeitzone },
      end: { dateTime: `${t.ende}:00`, timeZone: config.zeitzone },
      location: { displayName: t.ort || "" },
      categories: t.kategorien || [],
      showAs: "busy",
      isReminderOn: false,
      singleValueExtendedProperties: [{ id: MARKIERUNG, value: "1" }],
    });
  },
  markiere: (id) => graph("PATCH", `/events/${id}`, { singleValueExtendedProperties: [{ id: MARKIERUNG, value: "1" }] }),
  aendere: (id, felder) => graph("PATCH", `/events/${id}`, felder),
  loesche: (id) => graph("DELETE", `/events/${id}`),
};

const heute = new Date().toLocaleDateString("sv-SE", { timeZone: "Europe/Berlin" });
const ergebnis = await planeZeitraum(api, naechsteTage(heute, anzahlTage), einstellungen, { schreiben });
const zeit = (t) => `${t.start.slice(11)}–${t.ende.slice(11)}`;
console.log(`${schreiben ? "Ausgeführt" : "Nur Vorschau"}: ${heute} + ${anzahlTage} Tage, ${ergebnis.length} Tage mit Änderungen`);
for (const r of ergebnis) {
  console.log(
    `${r.datum}: +${r.anlegen.length} [${r.anlegen.map(zeit).join(", ")}]` +
      ` −${r.loeschen.length} [${r.loeschen.map(zeit).join(", ")}]` +
      ` übernommen ${r.uebernehmen.length}, freigegeben ${r.freigeben.length}`
  );
}
