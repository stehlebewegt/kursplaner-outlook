// Startet die Automatik (läuft in GitHub Actions, siehe .github/workflows).
// Anmeldung als App mit Geheimnis, nicht als Michael: braucht die
// Anwendungsberechtigung Calendars.ReadWrite in Entra.
//
// Umgebungsvariablen (als GitHub-Secrets hinterlegt, nie ins Repo):
//   KURSPLANER_SECRET   Geheimer Clientschlüssel der Entra-App
//   KURSPLANER_POSTFACH E-Mail-Adresse des Kalenders
//   KURSPLANER_KONTAKTE weitere Postfächer mit Kontakten, kommagetrennt (optional)
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
import { ohneFahrzeit, kursAnbieter, ROUTEN_BESCHRIFTUNG, routenHtml, erinnerungsFelder, KEINE_ERINNERUNG } from "../src/planer.js";
import { routenRechner, hatAdresse } from "./route.js";
import { monatsberichte, abschluesse } from "./monatsbericht.js";
import { hinweise } from "./hinweise.js";
import { holeFerien, holeFeiertage } from "./ferien.js";
import { ferienKurse } from "../src/ferien.js";

// Ziel für den Routen-Link der Rückfahrten. Nur zur Laufzeit aus dem Secret,
// nie im Repo und nie in der Ausgabe.
einstellungen.heimadresse = process.env.KURSPLANER_HEIMADRESSE || "";

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

function adresseVon(l) {
  const a = l?.address || {};
  const teile = [a.street, [a.postalCode, a.city].filter(Boolean).join(" "), a.countryOrRegion].map((x) => (x || "").trim()).filter(Boolean);
  return a.street || a.city ? teile.join(", ") : l?.displayName || "";
}

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
    // Für die Routenberechnung lieber die hinterlegte Adresse als den Anzeigenamen
    // (z. B. "Parkplatz kostenloser Teil" mit Adresse Bleichestraße 15, Bad Waldsee).
    adresse: adresseVon(e.location),
    showAs: e.showAs,
    ...(e.bodyPreview !== undefined ? { vorschau: e.bodyPreview } : {}),
    ...(e.isReminderOn !== undefined ? { erinnerung: e.isReminderOn ? e.reminderMinutesBeforeStart : KEINE_ERINNERUNG } : {}),
    generiert: prop?.value === "1",
    fahrzeit,
  };
}

const api = {
  async holeBereich(von, bis) {
    let url =
      `/calendarView?startDateTime=${von}T00:00:00&endDateTime=${bis}T23:59:59&$top=200` +
      `&$select=subject,start,end,categories,location,showAs,isAllDay,bodyPreview,isReminderOn,reminderMinutesBeforeStart` +
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
  async legeAn(t, { generiert = true, text, html } = { generiert: true }) {
    await graph("POST", "/events", {
      subject: t.titel,
      start: { dateTime: `${t.start}:00`, timeZone: config.zeitzone },
      end: { dateTime: `${t.ende}:00`, timeZone: config.zeitzone },
      location: { displayName: t.ort || "" },
      categories: t.kategorien || [],
      showAs: "busy",
      ...erinnerungsFelder(t.erinnerung),
      ...(html ? { body: { contentType: "html", content: html } } : text ? { body: { contentType: "text", content: text } } : {}),
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

// Einmalig Routen-Links in vorhandene Fahrt-Blöcke schreiben, ohne neu zu
// planen: ROUTENLINKS="JJJJ-MM-TT:JJJJ-MM-TT". Ziel ist der Ort des Blocks,
// bei Rückfahrten die Heimadresse. Schreiben nur mit NACHBERECHNEN_SCHREIBEN=ja.
if (process.env.ROUTENLINKS) {
  const [von, bis] = process.env.ROUTENLINKS.split(":");
  const wirklich = schreiben && process.env.NACHBERECHNEN_SCHREIBEN === "ja";
  const bloecke = (await api.holeBereich(von, bis)).filter(
    (t) => t.generiert && t.kategorien.includes(einstellungen.kategorieFahrt) && !(t.vorschau || "").includes(ROUTEN_BESCHRIFTUNG)
  );
  let gesetzt = 0;
  for (const t of bloecke) {
    const ziel = t.titel.startsWith("Rückfahrt") ? einstellungen.heimadresse : t.adresse || t.ort;
    if (!ziel) continue;
    if (wirklich) await api.aendere(t.id, { body: { contentType: "html", content: routenHtml(ziel) } });
    gesetzt++;
  }
  ausgabe(`Routen-Links ${wirklich ? "geschrieben" : "Probelauf"}: ${bloecke.length} Fahrt-Blöcke ohne Link, ${gesetzt} mit Ziel [${bloecke.map((t) => t.start.slice(5, 16).replace("T", " ")).join(", ")}]`);
  console.log(`::notice::${zeilen.join("%0A")}`);
  process.exit(0);
}

// Einmalige Nachberechnung (per Hand gestartet): NACHBERECHNEN="JJJJ-MM-TT:JJJJ-MM-TT"
// plant nur die Puffer-Blöcke dieser Tage neu, sonst nichts. Geschrieben wird
// nur mit NACHBERECHNEN_SCHREIBEN=ja, sonst ist es ein Probelauf.
if (process.env.NACHBERECHNEN) {
  const [von, bis] = process.env.NACHBERECHNEN.split(":");
  const wirklich = schreiben && process.env.NACHBERECHNEN_SCHREIBEN === "ja";
  const bereich = naechsteTage(von, Math.round((Date.parse(bis) - Date.parse(von)) / 86400000) + 1);
  // Erst private Termine mit Adresse ohne Fahrzeit berechnen (ohne Obergrenze),
  // damit die Planung danach ihre Fahrt-Blöcke anlegen kann.
  if (process.env.KURSPLANER_HEIMADRESSE) {
    const rechnePrivat = routenRechner({ heimadresse: process.env.KURSPLANER_HEIMADRESSE });
    const offen = ohneFahrzeit(await api.holeBereich(von, bis), einstellungen).filter((t) => hatAdresse(t.adresse));
    let gefunden = 0;
    for (const t of offen) {
      let minuten;
      try {
        minuten = await rechnePrivat(t.adresse);
      } catch {
        continue;
      }
      if (minuten != null) gefunden++;
      if (wirklich) await api.setzeFahrzeit(t.id, minuten ?? -1);
    }
    ausgabe(`Private Termine mit Adresse ohne Fahrzeit: ${offen.length}, davon Route gefunden ${gefunden}`);
  }
  const monate = new Map();
  for (const monat of [...new Set(bereich.map((d) => d.slice(0, 7)))]) {
    const r = await planeZeitraum(api, bereich.filter((d) => d.startsWith(monat)), einstellungen, { schreiben: wirklich });
    monate.set(monat, r);
  }
  ausgabe(`Nachberechnung ${von} bis ${bis} (${wirklich ? "AUSGEFÜHRT" : "nur Probelauf, nichts geändert"}):`);
  for (const [monat, r] of monate) {
    const summe = (k) => r.reduce((s, x) => s + x[k].length, 0);
    ausgabe(`${monat}: ${r.length} Tage mit Änderungen, Blöcke +${summe("anlegen")} −${summe("loeschen")}, übernommen ${summe("uebernehmen")}, freigegeben ${summe("freigeben")}`);
  }
  console.log(`::notice::${zeilen.join("%0A")}`);
  process.exit(0);
}

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

// Schulferien: reguläre Kurse bei Anbietern ohne Ferienkurse bekommen (F) und
// werden frei; die Planung danach nimmt dann ihre Puffer weg.
try {
  const ferien = await holeFerien(tage[0], tage[tage.length - 1]);
  // Feiertage getrennt holen: fällt der Abruf aus, gelten trotzdem die Ferien.
  const feiertage = await holeFeiertage(tage[0], tage[tage.length - 1]).catch(() => []);
  const betroffen = ferienKurse(await api.holeBereich(tage[0], tage[tage.length - 1]), ferien, einstellungen, feiertage);
  const setzen = schreiben && einstellungen.ferienAutomatisch;
  if (setzen) for (const t of betroffen) await api.aendere(t.id, { subject: `(F) ${t.titel.trim()}`, showAs: "free" });
  ausgabe(
    `Ferien BW: ${ferien.map((f) => `${f.von}–${f.bis}`).join(", ") || "keine"}; Feiertage: ${feiertage.map((f) => f.von).join(", ") || "keine"}; Kurse auf (F) ${setzen ? "gesetzt" : "zu setzen (Probe)"}: ${betroffen.length} [${betroffen.map(tagZeit).join(", ")}]`
  );
} catch (e) {
  ausgabe(`Ferien: ${e.message.slice(0, 100)}`);
}

// Private Termine mit Adresse: Fahrzeit ab zu Hause berechnen und im Termin
// merken (einmal pro Termin). Heimadresse nur als Secret, nie im Code.
const heimadresse = process.env.KURSPLANER_HEIMADRESSE;
const rechne = heimadresse ? routenRechner({ heimadresse }) : null;
if (heimadresse) {
  const offen = ohneFahrzeit(await api.holeBereich(tage[0], tage[tage.length - 1]), einstellungen)
    .filter((t) => hatAdresse(t.adresse))
    .slice(0, 20);
  const berechnet = [];
  for (const t of offen) {
    let minuten;
    try {
      minuten = await rechne(t.adresse);
    } catch {
      continue; // Netz-Problem: beim nächsten Lauf nochmal
    }
    // Nicht gefunden: -1 merken, dann fragt die Automatik nicht jedes Mal neu,
    // und der Kursplaner kann es anzeigen.
    if (schreiben) await api.setzeFahrzeit(t.id, minuten ?? -1);
    berechnet.push(`${tagZeit(t)}=${minuten ?? "?"}`);
  }
  ausgabe(`Fahrzeiten berechnet: ${berechnet.length} [${berechnet.join(", ")}]`);
  // Überblick: private Termine mit Adresse und ihr Fahrzeit-Stand (ohne Titel).
  const privat = (await api.holeBereich(tage[0], tage[6])).filter(
    (t) => !t.generiert && hatAdresse(t.adresse) && !kursAnbieter(t, einstellungen)
  );
  ausgabe(`Private Termine (7 Tage): ${privat.map((t) => `${tagZeit(t)}=${t.fahrzeit}`).join(", ")}`);
}

const ergebnis = await planeZeitraum(api, tage, einstellungen, { schreiben });
ausgabe(`${schreiben ? "Ausgeführt" : "Nur Vorschau"}: ${heute} + ${anzahlTage} Tage, ${ergebnis.length} Tage mit Änderungen`);
for (const r of ergebnis) {
  ausgabe(
    `${r.datum}: +${r.anlegen.length} [${r.anlegen.map(zeit).join(", ")}]` +
      ` −${r.loeschen.length} [${r.loeschen.map(zeit).join(", ")}]` +
      ` übernommen ${r.uebernehmen.length}, freigegeben ${r.freigeben.length}, Fahrzeit erkannt ${r.fahrzeit.length}` +
      (r.nachbessern.length ? `, nachzubessern ${r.nachbessern.length}${einstellungen.routenLinkNachtragen ? " (erledigt)" : ""}` : "")
  );
}
// Monatsbericht als Entwurf (am eingestellten Tag oder als Probe per Hand).
if (process.env.ABSCHLUESSE) {
  await abschluesse({
    graph,
    holeBereich: api.holeBereich,
    rechne,
    einstellungen,
    jetzt: new Date().toLocaleString("sv-SE", { timeZone: "Europe/Berlin" }).replace(" ", "T").slice(0, 16),
    bereich: process.env.ABSCHLUESSE,
    ausgabe,
  });
}
try {
  await hinweise({
    graph,
    holeBereich: api.holeBereich,
    einstellungen,
    jetzt: new Date().toLocaleString("sv-SE", { timeZone: "Europe/Berlin" }).replace(" ", "T").slice(0, 16),
    ausgabe,
    erzwingen: process.env.HINWEISE_PROBE === "ja",
  });
} catch (e) {
  ausgabe(`Hinweise fehlgeschlagen: ${e.message.slice(0, 120)}`);
}
try {
  await monatsberichte({
    graph,
    holeBereich: api.holeBereich,
    rechne,
    einstellungen,
    heute,
    jetzt: new Date().toLocaleString("sv-SE", { timeZone: "Europe/Berlin" }).replace(" ", "T").slice(0, 16),
    probe: process.env.MONATSBERICHT_PROBE || "",
    ausgabe,
    kontaktPostfaecher: (process.env.KURSPLANER_KONTAKTE || "").split(",").map((p) => p.trim()).filter(Boolean),
  });
} catch (e) {
  ausgabe(/ 403 /.test(e.message) ? "Monatsbericht: Mail-Berechtigung (Mail.ReadWrite) fehlt noch" : `Monatsbericht fehlgeschlagen: ${e.message.slice(0, 120)}`);
}
console.log(`::notice::${zeilen.join("%0A")}`);
