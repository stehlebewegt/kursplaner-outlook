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
import { MARKIERUNG, FAHRZEIT } from "../web/graph.js";
import { einstellungen } from "../src/einstellungen.js";
import { planeZeitraum, naechsteTage } from "./planen.js";
import { gleicheAb, alsEigenerKurs } from "../src/dauerkurse.js";
import { ohneFahrzeit } from "../src/planer.js";
import { routenRechner, hatAdresse } from "./route.js";

// Fehler und Ergebnis zusätzlich als GitHub-Hinweis ausgeben, damit sie
// auf der Übersichtsseite des Laufs stehen.
process.on("uncaughtException", (e) => {
  console.log(`::error::${String(e.message).replace(/\n/g, " ")}`);
  process.exit(1);
});

const { KURSPLANER_SECRET: secret, KURSPLANER_POSTFACH: postfach } = process.env;
if (!secret || !postfach) throw new Error(`Secret fehlt: ${[!secret && "KURSPLANER_SECRET", !postfach && "KURSPLANER_POSTFACH"].filter(Boolean).join(", ")}`);
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
  const props = e.singleValueExtendedProperties || [];
  const prop = props.find((p) => p.id === MARKIERUNG);
  // null = nie eingestellt, 0 = bewusst keine Fahrzeit.
  const wert = props.find((p) => p.id === FAHRZEIT)?.value;
  const fahrzeit = wert === undefined ? null : Number(wert) || 0;
  return {
    id: e.id,
    titel: e.subject || "",
    start: e.start.dateTime.slice(0, 16),
    ende: e.end.dateTime.slice(0, 16),
    kategorien: e.categories || [],
    ort: e.location?.displayName || "",
    showAs: e.showAs,
    generiert: prop?.value === "1",
    fahrzeit,
  };
}

const api = {
  async holeBereich(von, bis) {
    let url =
      `/calendarView?startDateTime=${von}T00:00:00&endDateTime=${bis}T23:59:59&$top=200` +
      `&$select=subject,start,end,categories,location,showAs,isAllDay` +
      `&$expand=singleValueExtendedProperties($filter=id eq '${MARKIERUNG}' or id eq '${FAHRZEIT}')`;
    const termine = [];
    while (url) {
      const seite = await graph("GET", url);
      termine.push(...seite.value.filter((e) => !e.isAllDay).map(zuTermin));
      url = seite["@odata.nextLink"];
    }
    return termine;
  },
  // Kurse aus den Anbieter-Kalendern, mit originalStart (Ortszeit) bei
  // Serienterminen, damit verschobene Termine erkennbar sind.
  async holeAnbieter(von, bis) {
    const kalender = (await graph("GET", "/calendars?$select=id,name&$top=50")).value;
    const termine = [];
    for (const a of einstellungen.anbieter) {
      const k = kalender.find((k) => k.name === a.kalender);
      if (!k) continue;
      let url =
        `/calendars/${k.id}/calendarView?startDateTime=${von}T00:00:00&endDateTime=${bis}T23:59:59&$top=200` +
        `&$select=subject,start,end,categories,location,isAllDay,originalStart,body`;
      while (url) {
        const seite = await graph("GET", url);
        for (const e of seite.value.filter((e) => !e.isAllDay)) {
          termine.push({ ...zuTermin(e), text: e.body?.content || "", originalStart: e.originalStart ? ortszeit(e.originalStart) : null });
        }
        url = seite["@odata.nextLink"];
      }
    }
    return termine;
  },
  async legeAn(t, { generiert = true, text } = { generiert: true }) {
    await graph("POST", "/events", {
      subject: t.titel,
      start: { dateTime: `${t.start}:00`, timeZone: config.zeitzone },
      end: { dateTime: `${t.ende}:00`, timeZone: config.zeitzone },
      location: { displayName: t.ort || "" },
      categories: t.kategorien || [],
      showAs: "busy",
      isReminderOn: false,
      ...(text ? { body: { contentType: "text", content: text } } : {}),
      ...(generiert ? { singleValueExtendedProperties: [{ id: MARKIERUNG, value: "1" }] } : {}),
    });
  },
  setzeFahrzeit: (id, minuten) =>
    graph("PATCH", `/events/${id}`, { singleValueExtendedProperties: [{ id: FAHRZEIT, value: String(minuten) }] }),
  markiere: (id) => graph("PATCH", `/events/${id}`, { singleValueExtendedProperties: [{ id: MARKIERUNG, value: "1" }] }),
  aendere: (id, felder) => graph("PATCH", `/events/${id}`, felder),
  loesche: (id) => graph("DELETE", `/events/${id}`),
};

// UTC-Zeitpunkt (z. B. "2026-11-04T14:00:00Z") als deutsche Ortszeit "JJJJ-MM-TTTHH:MM".
function ortszeit(utc) {
  const iso = /[zZ]|[+-]\d\d:\d\d$/.test(utc) ? utc : `${utc}Z`;
  return new Date(iso).toLocaleString("sv-SE", { timeZone: "Europe/Berlin" }).replace(" ", "T").slice(0, 16);
}

const heute = new Date().toLocaleDateString("sv-SE", { timeZone: "Europe/Berlin" });
const tage = naechsteTage(heute, anzahlTage);
const zeit = (t) => `${t.start.slice(11)}–${t.ende.slice(11)}`;
const tagZeit = (t) => `${t.start.slice(8, 10)}.${t.start.slice(5, 7)}. ${t.start.slice(11)}`;
const zeilen = [];
const ausgabe = (z) => (console.log(z), zeilen.push(z));

// Erst Dauerkurse abgleichen, dann die Tage planen.
const dauer = gleicheAb(
  await api.holeAnbieter(tage[0], tage[tage.length - 1]),
  await api.holeBereich(tage[0], tage[tage.length - 1]),
  einstellungen
);
const dauerSchreiben = schreiben && einstellungen.dauerkurse.schreiben;
if (dauerSchreiben) {
  for (const p of dauer.anlegen) {
    const kurs = alsEigenerKurs(p);
    await api.legeAn(kurs, { generiert: false, text: kurs.text });
  }
  for (const v of dauer.verschieben) {
    await api.aendere(v.termin.id, {
      start: { dateTime: `${v.start}:00`, timeZone: config.zeitzone },
      end: { dateTime: `${v.ende}:00`, timeZone: config.zeitzone },
    });
  }
}
ausgabe(
  `Dauerkurse (${dauerSchreiben ? "ausgeführt" : "nur Vorschau"}): neu ${dauer.anlegen.length} [${dauer.anlegen.map(tagZeit).join(", ")}]` +
    `, verschoben ${dauer.verschieben.length} [${dauer.verschieben.map((v) => `${tagZeit(v.termin)}→${v.start.slice(11)}`).join(", ")}]` +
    `, eigene ohne Gegenstück ${dauer.ohneGegenstueck.length} [${dauer.ohneGegenstueck.map(tagZeit).join(", ")}]`
);

// Private Termine mit Adresse: Fahrzeit ab zu Hause berechnen und im Termin
// merken (einmal pro Termin). Heimadresse nur als Secret, nie im Code.
const heimadresse = process.env.KURSPLANER_HEIMADRESSE;
if (heimadresse) {
  const rechne = routenRechner({ heimadresse });
  const offen = ohneFahrzeit(await api.holeBereich(tage[0], tage[tage.length - 1]), einstellungen)
    .filter((t) => hatAdresse(t.ort))
    .slice(0, 20);
  const berechnet = [];
  for (const t of offen) {
    let minuten;
    try {
      minuten = await rechne(t.ort);
    } catch {
      continue; // Netz-Problem: beim nächsten Lauf nochmal
    }
    // Nicht gefunden: 0 merken, dann fragt die Automatik nicht jedes Mal neu.
    if (schreiben) await api.setzeFahrzeit(t.id, minuten ?? 0);
    berechnet.push(`${tagZeit(t)}=${minuten ?? "?"}`);
  }
  ausgabe(`Fahrzeiten berechnet: ${berechnet.length} [${berechnet.join(", ")}]`);
}

const ergebnis = await planeZeitraum(api, tage, einstellungen, { schreiben });
ausgabe(`${schreiben ? "Ausgeführt" : "Nur Vorschau"}: ${heute} + ${anzahlTage} Tage, ${ergebnis.length} Tage mit Änderungen`);
for (const r of ergebnis) {
  ausgabe(
    `${r.datum}: +${r.anlegen.length} [${r.anlegen.map(zeit).join(", ")}]` +
      ` −${r.loeschen.length} [${r.loeschen.map(zeit).join(", ")}]` +
      ` übernommen ${r.uebernehmen.length}, freigegeben ${r.freigeben.length}, Fahrzeit erkannt ${r.fahrzeit.length}`
  );
}
console.log(`::notice::${zeilen.join("%0A")}`);
