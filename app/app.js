// Kursplaner als eigenes Fenster (installierbar über Edge/Chrome): alles,
// was das Add-in kann, ohne dass in Outlook etwas markiert sein muss.
// Anmeldung mit MSAL (Weiterleitung), Daten direkt aus Outlook über Graph.
import * as api from "../web/graph.js";
import { config } from "../web/config.js";
import { einstellungen } from "../src/einstellungen.js";
import { kursAnbieter, leseKuerzel, ohneKuerzel, routenUrl } from "../src/planer.js";
import { leseProtokoll, AUSFALLGRUENDE, leseAusfallgrund, mitAusfallgrund } from "../src/protokoll.js";
import { kursDetails } from "../src/kursdetails.js";
import { kursbericht, wieOftWo, WIE_OFT_ARTEN, geld, pruefliste } from "../src/auswertungen.js";
import { auszug } from "../src/eigeneEinstellungen.js";
import { wendeAn } from "../src/eigeneEinstellungen.js";
import { setzeKuerzel, verschiebeKurs } from "../src/aktionen.js";
import { fahrtenJahr, fahrtenPdf, alleZiele } from "../src/fahrtenbuch.js";
import { erstelleAuswertung } from "../src/auswertung.js";
import { adressVarianten } from "../automatik/route.js";
import {
  $, zustand, nutzeHost, mitWarten, zeige, zeigeAenderungen, zeigeProtokoll, richteAuswertungEin, richteNeuerKursEin,
  richteMailEin, richteUrlaubEin, uebernehme,
} from "../web/bereiche.js";

const html = (t) => String(t).replace(/[&<>"]/g, (z) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[z]);
const jetztLokal = () => new Date().toLocaleString("sv-SE", { timeZone: "Europe/Berlin" }).replace(" ", "T").slice(0, 16);
const plusTage = (datum, n) => new Date(Date.parse(`${datum}T12:00:00Z`) + n * 86400000).toISOString().slice(0, 10);
const TAGE = ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"];
const tagName = (d) => TAGE[new Date(`${d}T12:00:00Z`).getUTCDay()];
const tagDe = (d) => `${d.slice(8, 10)}.${d.slice(5, 7)}.`;

// --- Anmeldung ---
const msalApp = new msal.PublicClientApplication({
  auth: {
    clientId: config.clientId,
    authority: `https://login.microsoftonline.com/${config.tenantId}`,
    redirectUri: location.origin + location.pathname,
  },
  cache: { cacheLocation: "localStorage" },
});
let konto;
async function token() {
  try {
    return (await msalApp.acquireTokenSilent({ scopes: config.scopes, account: konto })).accessToken;
  } catch {
    await msalApp.acquireTokenRedirect({ scopes: config.scopes, account: konto });
    return new Promise(() => {}); // Seite wird gleich neu geladen
  }
}

// --- Was im Add-in Outlook erledigt, macht hier der Browser ---
const speicher = {
  lies: (k) => {
    try {
      return localStorage.getItem(`kursplaner-${k}`);
    } catch {
      return null;
    }
  },
  merke: (k, v) => {
    try {
      localStorage.setItem(`kursplaner-${k}`, String(v));
    } catch {
      /* ohne Speicher eben nicht gemerkt */
    }
  },
};
nutzeHost({
  ...speicher,
  oeffneEntwurf: () => false, // liegt in den Entwürfen, Outlook zeigt ihn dort
  terminText: (t) => Promise.resolve(t.text || ""),
  istSerie: () => Promise.resolve(false), // hier immer ein einzelner Termin
  schreibeInsFenster: () => Promise.resolve(false),
  nachProtokoll: () => schliesseLeiste(),
});

// Ergebnis-Meldungen unten rechts: Klick oder nach einer Weile weg.
let weg;
new MutationObserver(() => {
  clearTimeout(weg);
  if ($("ergebnis").innerHTML) weg = setTimeout(() => ($("ergebnis").innerHTML = ""), 10000);
}).observe($("ergebnis"), { childList: true });
$("ergebnis").onclick = () => ($("ergebnis").innerHTML = "");

// Die Formulare (Protokoll, Neuer Kurs, Urlaub, Monatsauswertung, E-Mail) gibt
// es nur einmal, in taskpane.html. Hier werden sie an ihre Plätze gesetzt.
async function holeBereiche() {
  const seite = new DOMParser().parseFromString(await (await fetch("../taskpane.html")).text(), "text/html");
  document.querySelectorAll("[data-bereich]").forEach((platz) => {
    const bereich = seite.getElementById(platz.dataset.bereich);
    bereich.hidden = false;
    bereich.open = true;
    platz.replaceWith(document.importNode(bereich, true));
  });
}

// --- Ansichten ---
const ANSICHTEN = {
  heute: () => zeigeHeute(),
  protokolle: () => zeigeOffeneProtokolle(),
  woche: () => zeigeWoche(),
  neu: () => {},
  vertretung: () => zeigeVertretungen(),
  urlaub: () => {},
  auswertung: () => {},
  fahrten: () => {},
  einstellungen: () => zeigeEinstellungen(),
  kursbericht: () => richteKursberichtEin(),
  wieoft: () => zeigeWieOft(),
  geld: () => zeigeGeld(),
  pruefliste: () => zeigePruefliste(),
};
let aktuell = "heute";
function oeffne(name) {
  document.querySelectorAll("nav button, #unten button").forEach((b) => b.classList.toggle("an", b.dataset.ansicht === name));
  // „Mehr“ ist aktiv, wenn die Ansicht nicht unten in der Leiste steht.
  $("mehr").classList.toggle("an", !document.querySelector(`#unten [data-ansicht="${name}"]`));
  document.querySelector("nav").classList.remove("offen");
  document.querySelectorAll("section.ansicht").forEach((s) => (s.hidden = s.id !== `ansicht-${name}`));
  $("leiste").hidden = true;
  aktuell = name;
  ANSICHTEN[name]();
}

// Wie ein Kurs in der Liste heißt und wo er stattfindet.
const FARBEN = { SVE: "#5B787D", PTM: "#C4A870", VTR: "#8C6BB1" };
const KUERZEL = { U: "Urlaub", A: "abgesagt", F: "Ferien", V: "Vertretung", H: "Hospitation" };
function beschreibe(t) {
  const a = kursAnbieter(t, einstellungen);
  const k = leseKuerzel(t.titel);
  const ohne = ohneKuerzel(t.titel);
  const klammer = ohne.match(/^(.*?)\s*\(([^)]+)\)\s*$/);
  const name = klammer ? klammer[1] : ohne;
  const wer = klammer ? klammer[2] : "";
  const wo = a.extern ? `${wer} · anderer Verein` : `${a.kuerzel} · ${a.name}${wer && wer !== einstellungen.dauerkurse.kursleitung ? ` · für ${wer}` : ""}`;
  return { a, k, name, wo, faelltAus: einstellungen.kuerzel[k]?.aktiv === false };
}

// Kurse als Karten mit Knöpfen (Woche, Heute, Offene Protokolle).
const istKurs = (t) =>
  !t.generiert && kursAnbieter(t, einstellungen) && !/^(Anfahrt|🚗 Jetzt losfahren|Rückfahrt|Fahrt |Vorbereitung|Nachbereitung|Nach-\/Vorbereitung)/.test(t.titel);
function zeigeKurse(kurse, ziel, neuLaden, { mitTag = true, knoepfe = ["protokoll", "mail", "verschieben", "kuerzel"] } = {}) {
  let letzterTag = "";
  ziel.innerHTML = kurse
    .map((t, i) => {
      const d = t.start.slice(0, 10);
      const b = beschreibe(t);
      const kopf = mitTag && d !== letzterTag ? `<div class="tag">${tagName(d)}, ${tagDe(d)}</div>` : "";
      letzterTag = d;
      return (
        kopf +
        `<div class="kurs${b.faelltAus ? " faellt-aus" : ""}" data-details="${i}" title="Details anzeigen">` +
        `<div class="streifen" style="background:${FARBEN[b.a.kuerzel] || "#5B787D"}"></div>` +
        `<div class="zeit">${t.start.slice(11)}–${t.ende.slice(11)}</div>` +
        `<div class="was">${html(b.name)}${KUERZEL[b.k] ? `<span class="marke-klein">${KUERZEL[b.k]}${b.k === "A" && leseAusfallgrund(t.text) ? ` · ${html(leseAusfallgrund(t.text))}` : ""}</span>` : ""}<small>${html(b.wo)}</small></div>` +
        `<div class="knoepfe">` +
        (knoepfe.includes("protokoll") && brauchtProtokoll(t) ? `<button class="knopf" data-protokoll="${i}">Protokoll</button>` : "") +
        (knoepfe.includes("mail") ? `<button class="knopf" data-mail="${i}">E-Mail</button>` : "") +
        (knoepfe.includes("verschieben") && !b.faelltAus ? `<button class="knopf" data-verschieben="${i}">Verschieben</button>` : "") +
        (!knoepfe.includes("kuerzel")
          ? ""
          : b.faelltAus
            ? `<button class="knopf" data-kuerzel="" data-kurs="${i}">Findet statt</button>`
            : `<button class="knopf" data-kuerzel="A" data-kurs="${i}">Absagen</button>`) +
        `</div></div>`
      );
    })
    .join("");
  // Klick auf den Kurs (nicht auf einen Knopf): Kursdetails.
  ziel.querySelectorAll("[data-details]").forEach((k) => {
    k.onclick = (e) => {
      if (!e.target.closest("button")) oeffneLeiste(kurse[k.dataset.details], "details", neuLaden);
    };
  });
  ziel.querySelectorAll("[data-protokoll]").forEach((b) => (b.onclick = () => oeffneLeiste(kurse[b.dataset.protokoll], "protokoll")));
  ziel.querySelectorAll("[data-mail]").forEach((b) => (b.onclick = () => oeffneLeiste(kurse[b.dataset.mail], "mail")));
  ziel.querySelectorAll("[data-verschieben]").forEach((b) => (b.onclick = () => oeffneLeiste(kurse[b.dataset.verschieben], "verschieben")));
  ziel.querySelectorAll("[data-kuerzel]").forEach((b) => (b.onclick = () => absagenOderStatt(kurse[b.dataset.kurs], b, neuLaden)));
}

// Kleine Auswahl mit Knöpfen (statt confirm), gibt die Wahl zurück oder null.
function waehle(titel, text, optionen) {
  const d = $("auswahl");
  $("auswahl-titel").textContent = titel;
  $("auswahl-text").textContent = text;
  return new Promise((fertig) => {
    $("auswahl-knoepfe").innerHTML = optionen.map((o, i) => `<button class="zweit" data-i="${i}">${html(o)}</button>`).join("");
    $("auswahl-knoepfe").querySelectorAll("button").forEach((b) => (b.onclick = () => (d.close(), fertig(optionen[b.dataset.i]))));
    $("auswahl-zu").onclick = () => d.close();
    d.onclose = () => fertig(null); // zählt nur, wenn vorher nichts gewählt wurde
    d.showModal();
  });
}

// (A) setzen (mit Ausfallgrund) oder wieder wegnehmen, nach Rückfrage.
const OHNE_GRUND = "Ohne Grund";
async function absagenOderStatt(t, b, neuLaden) {
  const absagen = b.dataset.kuerzel === "A";
  const kurs = `„${ohneKuerzel(t.titel)}“ am ${tagDe(t.start)}`;
  let grund = "";
  if (absagen) {
    const wahl = await waehle("Kurs absagen", `${kurs}: Warum fällt er aus? Fahrten und Vor-/Nachbereitung fallen dann weg.`, [...AUSFALLGRUENDE, OHNE_GRUND]);
    if (!wahl) return;
    grund = wahl === OHNE_GRUND ? "" : wahl;
  } else if (!confirm(`${kurs} findet doch statt? Fahrten und Vor-/Nachbereitung kommen wieder dazu.`)) return;
  mitWarten(b, async () => {
    const v = await setzeKuerzel(api, t, b.dataset.kuerzel, einstellungen);
    const text = mitAusfallgrund(t.text || "", grund);
    if (text !== (t.text || "").replace(/\r/g, "").trimEnd()) await api.aendere(t.id, { body: { contentType: "text", content: text } });
    zeigeAenderungen(v, absagen ? `Kurs abgesagt${grund ? ` (${grund})` : ""}` : "Kurs findet statt");
    $("leiste").hidden = true;
    await neuLaden();
  });
}

// --- Heute: nächste Abfahrt mit Navigation, Kurse von heute (oder vom nächsten Kurstag) ---
const ROUTE = /https:\/\/www\.google\.com\/maps\/dir\/[^\s<>"]+/;
async function zeigeHeute() {
  const heute = jetztLokal().slice(0, 10);
  $("h-unter").textContent = `${tagName(heute)}, ${tagDe(heute)}${heute.slice(0, 4)}`;
  $("h-fahrt").innerHTML = "";
  $("h-frei").innerHTML = "";
  zeigeWartet();
  zeigeAufgaben();
  zeigeNachrichten();
  $("h-liste").innerHTML = `<div class="leise">Wird geladen …</div>`;
  $("h-kalender").innerHTML = `<div class="leise">Wird geladen …</div>`;
  let termine;
  try {
    termine = await api.holeVerlauf(heute, plusTage(heute, 14));
  } catch (e) {
    $("h-liste").innerHTML = `<div class="fehler">Konnte den Kalender nicht laden: ${html(e.message)}</div>`;
    $("h-kalender").innerHTML = "";
    return;
  }
  zeigeKalender(termine, heute);
  const jetzt = jetztLokal();
  // Nächste Fahrt heute, die noch nicht vorbei ist.
  const fahrt = termine
    .filter((t) => t.start.slice(0, 10) === heute && t.ende > jetzt && (t.kategorien || []).includes(einstellungen.kategorieFahrt))
    .sort((a, b) => (a.start < b.start ? -1 : 1))[0];
  if (fahrt) {
    const link = (fahrt.text || "").match(ROUTE)?.[0] || (fahrt.titel.startsWith("Rückfahrt") || !fahrt.ort ? "" : routenUrl(fahrt.ort));
    $("h-fahrt").innerHTML =
      `<div class="karte fahrt"><div class="leise">${fahrt.start > jetzt ? "Nächste Abfahrt" : "Du bist unterwegs"}</div>` +
      `<div class="gross">${fahrt.start.slice(11)} Uhr</div><div>${html(fahrt.titel)}</div>` +
      (link ? `<a class="knopf-link" href="${html(link)}" target="_blank" rel="noopener">🧭 Route starten</a>` : "") +
      `</div>`;
  }
  const kurse = termine.filter(istKurs).sort((a, b) => (a.start < b.start ? -1 : 1));
  const heuteKurse = kurse.filter((t) => t.start.slice(0, 10) === heute);
  if (heuteKurse.length) {
    $("h-titel").textContent = "Deine Kurse heute";
    zeigeKurse(heuteKurse, $("h-liste"), zeigeHeute, { mitTag: false });
    return;
  }
  // Kein Kurs heute: deutlich als eigene Karte, darunter die Vorschau.
  const naechster = kurse.find((t) => t.start.slice(0, 10) > heute && !beschreibe(t).faelltAus);
  const frei = (zeile) =>
    `<div class="frei"><div class="symbol">🌿</div><div><div class="gross">Heute keine Kurse</div><div class="klein">${zeile}</div></div></div>`;
  if (!naechster) {
    $("h-frei").innerHTML = frei("Auch in den nächsten zwei Wochen steht kein Kurs in deinem Kalender.");
    $("h-titel").textContent = "";
    $("h-liste").innerHTML = "";
    return;
  }
  const tag = naechster.start.slice(0, 10);
  const tage = Math.round((new Date(`${tag}T12:00:00Z`) - new Date(`${heute}T12:00:00Z`)) / 864e5);
  $("h-frei").innerHTML = frei(`Nächster Kurstag: ${tagName(tag)}, ${tagDe(tag)} (${tage === 1 ? "morgen" : `in ${tage} Tagen`})`);
  $("h-titel").textContent = `Vorschau ${tagName(tag)}, ${tagDe(tag)}`;
  zeigeKurse(kurse.filter((t) => t.start.slice(0, 10) === tag), $("h-liste"), zeigeHeute, { mitTag: false });
}

// --- Überblick wie „Outlook Heute“ ---
// Kalender: die eigentlichen Termine von heute und den nächsten zwei Tagen,
// ohne Fahrt- und Vor-/Nachbereitungsblöcke.
const PUFFER_TITEL = /^(🚗\s*)?(Jetzt losfahren|Anfahrt|Rückfahrt|Vorbereitung|Nachbereitung|Nach-\/Vorbereitung) - /;
const istPuffer = (t) =>
  (t.kategorien || []).some((k) => k === einstellungen.kategorieFahrt || k === einstellungen.kategoriePuffer) || PUFFER_TITEL.test(t.titel);
function zeigeKalender(termine, heute) {
  const jetzt = jetztLokal();
  const tage = [0, 1, 2].map((n) => plusTage(heute, n));
  const name = (d, n) => (n === 0 ? "Heute" : n === 1 ? "Morgen" : tagName(d));
  $("h-kalender").innerHTML = tage
    .map((d, n) => {
      const ende = plusTage(d, 1);
      const amTag = termine
        .filter((t) => !istPuffer(t) && t.start < `${ende}T00:00` && t.ende > `${d}T00:00`)
        .sort((a, b) => (a.start < b.start ? -1 : 1));
      const zeilen = amTag.map((t) => {
        const ganz = t.start.slice(11) === "00:00" && t.ende.slice(11) === "00:00";
        const zeit = ganz ? "ganztägig" : `${t.start.slice(0, 10) < d ? "…" : t.start.slice(11)} – ${t.ende.slice(0, 10) > d ? "…" : t.ende.slice(11)}`;
        const art = t.ende <= jetzt ? "vorbei" : t.start <= jetzt ? "jetzt" : "";
        return (
          `<div class="k-zeile ${art}"><span class="zeit">${zeit}</span><span>${html(t.titel)}</span>` +
          (t.ort ? `<span class="ort">${html(t.ort)}</span>` : "") +
          `</div>`
        );
      });
      return `<div class="k-tag">${name(d, n)}, ${tagDe(d)}</div>` + (zeilen.join("") || `<div class="leise">Keine Termine</div>`);
    })
    .join("");
}

async function zeigeNachrichten() {
  const ziel = $("h-nachrichten");
  try {
    const z = await api.holeOrdnerZahlen();
    ziel.innerHTML = [
      ["Posteingang (ungelesen)", z.posteingang],
      ["Entwürfe", z.entwuerfe],
      ["Postausgang", z.postausgang],
    ]
      .map(([n, zahl]) => `<div class="n-zeile"><span>${n}</span><b>${zahl}</b></div>`)
      .join("");
  } catch (e) {
    ziel.innerHTML = `<div class="fehler">Konnte die Ordner nicht lesen: ${html(e.message)}</div>`;
  }
}

// Aufgaben kommen aus Microsoft To Do (dort landen auch die Outlook-Aufgaben).
// Dafür braucht es einmalig deine Zustimmung, deshalb ein eigener Schlüssel.
const AUFGABEN = ["Tasks.ReadWrite"];
async function aufgabenGraph(methode, pfad, body) {
  const t = (await msalApp.acquireTokenSilent({ scopes: AUFGABEN, account: konto })).accessToken;
  const antwort = await fetch(`https://graph.microsoft.com/v1.0${pfad}`, {
    method: methode,
    headers: { Authorization: `Bearer ${t}`, "Content-Type": "application/json", Prefer: `outlook.timezone="${config.zeitzone}"` },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (!antwort.ok) throw new Error(`${antwort.status} ${(await antwort.text()).slice(0, 160)}`);
  return antwort.json();
}

async function zeigeAufgaben() {
  const ziel = $("h-aufgaben");
  ziel.innerHTML = `<div class="leise">Wird geladen …</div>`;
  let aufgaben = [];
  try {
    const listen = (await aufgabenGraph("GET", "/me/todo/lists?$top=50")).value;
    for (const l of listen) {
      const filter = encodeURIComponent("status ne 'completed'");
      const antwort = await aufgabenGraph("GET", `/me/todo/lists/${l.id}/tasks?$filter=${filter}&$top=100`);
      aufgaben.push(...antwort.value.map((a) => ({ ...a, liste: l.id })));
    }
  } catch (e) {
    if (e instanceof msal.InteractionRequiredAuthError || /consent|interaction/i.test(e.errorCode || e.message)) {
      ziel.innerHTML =
        `<div class="leise">Damit der Kursplaner deine Aufgaben zeigen darf, musst du einmal zustimmen.</div>` +
        `<button id="h-aufgaben-ok">Aufgaben anzeigen</button>`;
      $("h-aufgaben-ok").onclick = () => msalApp.acquireTokenRedirect({ scopes: AUFGABEN, account: konto });
      return;
    }
    ziel.innerHTML = `<div class="fehler">Konnte die Aufgaben nicht laden: ${html(e.message)}</div>`;
    return;
  }
  const heute = jetztLokal().slice(0, 10);
  const faellig = (a) => a.dueDateTime?.dateTime?.slice(0, 10) || "";
  // Wie in Outlook: mit Datum zuerst (älteste oben), ohne Datum am Ende.
  aufgaben.sort((a, b) => (faellig(a) || "9") < (faellig(b) || "9") ? -1 : (faellig(a) || "9") > (faellig(b) || "9") ? 1 : a.title.localeCompare(b.title));
  if (!aufgaben.length) {
    ziel.innerHTML = `<div class="w-leer">✓ Keine offenen Aufgaben.</div>`;
    return;
  }
  ziel.innerHTML = aufgaben
    .map((a, i) => {
      const f = faellig(a);
      const zu = f && f < heute;
      const wann = f ? (f === heute ? "heute fällig" : `${zu ? "überfällig seit" : "fällig"} ${tagDe(f)}${f.slice(0, 4)}`) : "";
      return (
        `<label class="a-zeile${zu ? " ueberfaellig" : ""}"><input type="checkbox" data-aufgabe="${i}">` +
        `<span>${a.importance === "high" ? `<span class="wichtig">! </span>` : ""}${html(a.title)}` +
        (wann ? `<span class="faellig">${wann}</span>` : "") +
        `</span></label>`
      );
    })
    .join("");
  ziel.querySelectorAll("[data-aufgabe]").forEach(
    (box) =>
      (box.onchange = async () => {
        const a = aufgaben[box.dataset.aufgabe];
        box.disabled = true;
        try {
          await aufgabenGraph("PATCH", `/me/todo/lists/${a.liste}/tasks/${a.id}`, { status: "completed" });
          box.closest(".a-zeile").remove();
          if (!ziel.querySelector(".a-zeile")) ziel.innerHTML = `<div class="w-leer">✓ Keine offenen Aufgaben.</div>`;
        } catch (e) {
          box.checked = false;
          box.disabled = false;
          zeige(`<div class="karte fehler">Konnte die Aufgabe nicht abhaken: ${html(e.message)}</div>`);
        }
      })
  );
}

// „Wartet auf dich“ unter den Kursen: offene Protokolle, Prüfliste, Grenzen.
async function zeigeWartet() {
  const ziel = $("h-wartet");
  const j = Number(jetztLokal().slice(0, 4));
  let offen, p, knapp;
  try {
    offen = await ladeOffeneProtokolle();
    const termine = await ladeJahr(j);
    p = pruefliste(termine, j, jetztLokal(), einstellungen);
    knapp = geld(termine, j, jetztLokal(), einstellungen).filter((a) => a.anteilJahr != null && a.anteilJahr > 0.85);
  } catch {
    ziel.innerHTML = "";
    return;
  }
  zeigeAnzahl(offen.length);
  const zeilen = [];
  if (offen.length)
    zeilen.push({
      text: offen.length === 1 ? "1 Protokoll offen" : `${offen.length} Protokolle offen`,
      klein: offen.slice(0, 3).map((t) => `${tagDe(t.start)} ${beschreibe(t).name}`).join(" · ") + (offen.length > 3 ? " …" : ""),
      ansicht: "protokolle",
    });
  // „Ohne Teilnehmerzahl“ sind dieselben wie die offenen Protokolle.
  for (const x of p.punkte.filter((x) => x.id !== "teilnehmer" && x.termine.length))
    zeilen.push({ text: `${x.titel}: ${x.termine.length}`, klein: x.tun, ansicht: "pruefliste" });
  for (const a of knapp) zeilen.push({ text: `${a.name}: ${prozent(a.anteilJahr)} der Jahresgrenze`, klein: "Rest des Jahres im Blick behalten.", ansicht: "geld" });
  ziel.innerHTML =
    `<div class="tag">Wartet auf dich</div>` +
    (zeilen.length
      ? `<div class="karte wartet">${zeilen
          .map((z) => `<button class="w-zeile" data-ziel="${z.ansicht}"><b>${html(z.text)}</b><span>${html(z.klein)}</span><i>›</i></button>`)
          .join("")}</div>`
      : `<div class="w-leer">✓ Nichts offen. Alle Protokolle sind eingetragen.</div>`);
  ziel.querySelectorAll("[data-ziel]").forEach((b) => (b.onclick = () => oeffne(b.dataset.ziel)));
}

// --- Offene Protokolle: vergangene Kurse (60 Tage), die stattfanden, ohne Protokoll ---
// Nur selbst gehaltene Kurse (regulär, Vertretung); Urlaub, Ferien, Absage, Hospitation nicht.
const brauchtProtokoll = (t) => {
  const b = beschreibe(t);
  return !b.faelltAus && b.k !== "H";
};
const ohneProtokoll = (t) => {
  const p = leseProtokoll(t.text || "");
  return p.anwesend === "" && !p.thema;
};
async function ladeOffeneProtokolle() {
  const heute = jetztLokal().slice(0, 10);
  const jetzt = jetztLokal();
  const termine = await api.holeVerlauf(plusTage(heute, -60), heute);
  return termine
    .filter((t) => istKurs(t) && t.ende <= jetzt && brauchtProtokoll(t) && ohneProtokoll(t))
    .sort((a, b) => (a.start < b.start ? 1 : -1));
}
function zeigeAnzahl(n) {
  for (const id of ["p-anzahl", "p-anzahl2"]) {
    $(id).textContent = n ? String(n) : "";
    $(id).hidden = !n;
  }
}
async function zeigeOffeneProtokolle() {
  $("o-unter").textContent = "Wird geladen …";
  $("o-liste").innerHTML = "";
  try {
    const offen = await ladeOffeneProtokolle();
    zeigeAnzahl(offen.length);
    $("o-unter").textContent = offen.length
      ? `${offen.length} ${offen.length === 1 ? "Kurs" : "Kurse"} der letzten 60 Tage ohne Protokoll, die neuesten zuerst.`
      : "Alles erledigt.";
    if (!offen.length) {
      $("o-liste").innerHTML = `<div class="leer">Alle Kurse der letzten 60 Tage haben ein Protokoll. 👍</div>`;
      return;
    }
    zeigeKurse(offen, $("o-liste"), zeigeOffeneProtokolle, { knoepfe: ["protokoll"] });
  } catch (e) {
    $("o-unter").innerHTML = `<span class="fehler">Konnte den Kalender nicht laden: ${html(e.message)}</span>`;
  }
}

// --- Meine Woche ---
let wochenStart;
const montag = (d) => plusTage(d, -((new Date(`${d}T12:00:00Z`).getUTCDay() + 6) % 7));

async function zeigeWoche() {
  wochenStart ??= montag(jetztLokal().slice(0, 10));
  const bis = plusTage(wochenStart, 6);
  $("w-unter").textContent = "Wird geladen …";
  $("w-liste").innerHTML = "";
  let kurseDerWoche;
  try {
    kurseDerWoche = (await api.holeVerlauf(wochenStart, bis)).filter(istKurs).sort((a, b) => (a.start < b.start ? -1 : 1));
  } catch (e) {
    $("w-unter").innerHTML = `<span class="fehler">Konnte den Kalender nicht laden: ${html(e.message)}</span>`;
    return;
  }
  const finden = kurseDerWoche.filter((t) => !beschreibe(t).faelltAus);
  const tage = new Set(finden.map((t) => t.start.slice(0, 10)));
  const von = `${Number(wochenStart.slice(8, 10))}.${Number(wochenStart.slice(5, 7))}.`;
  $("w-unter").textContent =
    `${von} – ${tagDe(bis)}${bis.slice(0, 4)} · ${finden.length} ${finden.length === 1 ? "Kurs" : "Kurse"}` +
    (tage.size ? ` an ${tage.size} ${tage.size === 1 ? "Tag" : "Tagen"}` : "") +
    (kurseDerWoche.length > finden.length ? ` · ${kurseDerWoche.length - finden.length} ${kurseDerWoche.length - finden.length === 1 ? "fällt" : "fallen"} aus` : "");
  if (!kurseDerWoche.length) {
    $("w-liste").innerHTML = `<div class="leer">In dieser Woche stehen keine Kurse in deinem Kalender.</div>`;
    return;
  }
  zeigeKurse(kurseDerWoche, $("w-liste"), zeigeWoche);
}
$("w-zurueck").onclick = () => ((wochenStart = plusTage(wochenStart, -7)), zeigeWoche());
$("w-vor").onclick = () => ((wochenStart = plusTage(wochenStart, 7)), zeigeWoche());
$("w-heute").onclick = () => ((wochenStart = montag(jetztLokal().slice(0, 10))), zeigeWoche());

// Seitenleiste: Protokoll, E-Mail-Vorlagen oder Verschieben zu einem Kurs.
async function oeffneLeiste(t, was, neuLaden = () => ANSICHTEN[aktuell]()) {
  zustand.termin = t;
  $("bereich-details").hidden = was !== "details";
  $("leiste-titel").textContent = ohneKuerzel(t.titel);
  $("leiste-info").textContent = `${tagName(t.start.slice(0, 10))}, ${tagDe(t.start)} · ${t.start.slice(11)}–${t.ende.slice(11)}`;
  $("bereich-protokoll").hidden = was !== "protokoll";
  $("bereich-mail").hidden = was !== "mail";
  $("bereich-verschieben").hidden = was !== "verschieben";
  $("leiste").hidden = false;
  $("leiste").scrollTop = 0;
  // In den Details steht die Leitung als Schild, im Titel reicht der Kursname.
  if (was === "details") $("leiste-titel").textContent = ohneKuerzel(t.titel).replace(/\s*\([^)]+\)\s*$/, "");
  if (was === "details") await zeigeDetails(t, neuLaden);
  else if (was === "protokoll") await zeigeProtokoll();
  else if (was === "mail") richteMailEin(kursAnbieter(t, einstellungen));
  else await richteVerschiebenEin(t);
}

// --- Kursdetails: wann, wo, Protokoll, wie läuft der Kurs, was habe ich gemacht ---
const AMPEL_TEXT = { gut: "gut", beobachten: "beobachten", kritisch: "kritisch" };
const TREND = { steigt: "↗ Es kommen mehr", sinkt: "↘ Es kommen weniger", gleich: "→ Gleichbleibend" };
// „Halle zu“ bei Absagen mit Grund, sonst Urlaub, Ferien, abgesagt.
const ausfallText = (t) => (leseKuerzel(t.titel) === "A" && leseAusfallgrund(t.text)) || KUERZEL[leseKuerzel(t.titel)] || "fällt aus";
const komma = (z, n = 1) => z.toFixed(n).replace(".", ",");
async function zeigeDetails(t, neuLaden) {
  const ziel = $("bereich-details");
  const b = beschreibe(t);
  const datum = t.start.slice(0, 10);
  const heute = jetztLokal().slice(0, 10);
  const wer = ohneKuerzel(t.titel).match(/\(([^)]+)\)\s*$/)?.[1] || "";
  const ort = t.ort || b.a.adresse || b.a.ort || "";
  ziel.innerHTML =
    `<div class="chips"><span class="chip">${html(b.a.extern ? "Anderer Verein" : `${b.a.kuerzel} · ${b.a.name}`)}</span>` +
    `<span class="chip">${KUERZEL[b.k] || "Regulär"}</span>${wer ? `<span class="chip">Leitung ${html(wer)}</span>` : ""}</div>` +
    `<div class="d-ort"><span id="d-ort">📍 ${html(ort || "Ort fehlt im Termin")}</span>` +
    (ort ? `<a class="knopf-link" href="${html(routenUrl(ort))}" target="_blank" rel="noopener">🧭 Route</a>` : "") +
    `</div>` +
    `<div class="d-knoepfe"><button id="d-protokoll">Protokoll</button><button class="zweit" id="d-mail">E-Mail</button>` +
    (b.faelltAus ? "" : `<button class="zweit" id="d-verschieben">Verschieben</button>`) +
    `<button class="zweit" id="d-kuerzel" data-kuerzel="${b.faelltAus ? "" : "A"}">${b.faelltAus ? "Findet statt" : "Absagen"}</button></div>` +
    `<div id="d-inhalt" class="leise" style="padding-top:14px">Verlauf wird geladen …</div>`;
  $("d-protokoll").onclick = () => oeffneLeiste(t, "protokoll");
  $("d-mail").onclick = () => oeffneLeiste(t, "mail");
  if ($("d-verschieben")) $("d-verschieben").onclick = () => oeffneLeiste(t, "verschieben");
  $("d-kuerzel").onclick = () => absagenOderStatt(t, $("d-kuerzel"), neuLaden);

  let termine;
  try {
    // Ein Abruf: gut vier Monate zurück (Verlauf) und fünf Wochen voraus (nächste Termine).
    termine = await api.holeVerlauf(plusTage(datum < heute ? datum : heute, -120), plusTage(datum, 35));
  } catch (e) {
    $("d-inhalt").innerHTML = `<span class="fehler">Verlauf nicht geladen: ${html(e.message)}</span>`;
    return;
  }
  if (zustand.termin !== t) return; // inzwischen ein anderer Kurs geöffnet
  const d = kursDetails(t, termine.filter(istKurs), { jetzt: jetztLokal(), ohneKuerzel, faelltAus: (x) => beschreibe(x).faelltAus });

  // Abfahrt: der letzte Fahrt-Block am Kurstag, der vor dem Kurs endet.
  const fahrt = termine
    .filter((x) => x.start.slice(0, 10) === datum && /^(Anfahrt|🚗 Jetzt losfahren|Fahrt )/.test(x.titel) && x.ende <= t.start)
    .sort((x, y) => (x.ende < y.ende ? 1 : -1))[0];
  if (fahrt) $("d-ort").insertAdjacentHTML("beforeend", ` · Abfahrt ${fahrt.start.slice(11)}`);

  const p = d.dieser.p;
  const vorbei = t.ende <= jetztLokal();
  const dieser = b.faelltAus
    ? `<div class="protokoll-kurz">Dieser Termin ${KUERZEL[b.k] === "Urlaub" ? "fällt wegen Urlaub aus" : `ist ${KUERZEL[b.k] || "abgesagt"}`}` +
      `${leseAusfallgrund(t.text) ? ` (${html(leseAusfallgrund(t.text))})` : ""}.</div>`
    : p.thema || d.dieser.anwesend != null
      ? `<div class="protokoll-kurz"><b>${html(p.thema || "ohne Thema")}</b>` +
        (p.material ? `<div class="leise">${html(p.material)}</div>` : "") +
        (d.dieser.anwesend != null ? `<div>${d.dieser.anwesend} von ${d.dieser.gesamt} da</div>` : "") +
        (p.besonderheiten || /^ja/i.test(p.vorkommnis) ? `<span class="marke-klein" style="margin:6px 0 0;display:inline-block">Besonderheit notiert</span>` : "") +
        `</div>`
      : `<div class="protokoll-kurz${vorbei ? " offen" : ""}"><b>${vorbei ? "Protokoll noch offen" : "Protokoll nach dem Kurs"}</b>` +
        (d.verlauf[0] ? `<div class="leise">Zuletzt: ${html(d.verlauf[0].p.thema || "ohne Thema")} am ${tagDe(d.verlauf[0].t.start)}</div>` : "") +
        `</div>`;

  const hoechst = Math.max(1, ...d.balken.map((v) => v.gesamt || v.anwesend || 0));
  const balken = d.balken
    .map((v) =>
      v.ausfall || v.anwesend == null
        ? `<div><i class="aus" style="height:14px"></i>${tagDe(v.t.start)}<br>${v.ausfall ? html(ausfallText(v.t)) : "ohne"}</div>`
        : `<div title="${v.anwesend} von ${v.gesamt}"><i style="height:${Math.max(6, Math.round((v.anwesend / hoechst) * 60))}px"></i>${tagDe(v.t.start)}<br>${v.anwesend}</div>`
    )
    .join("");
  const kachel = (wert, text) => `<div class="kachel"><b>${wert}</b><span>${text}</span></div>`;
  const wie = d.gehalten
    ? `<div class="kacheln">${kachel(d.gehalten, "Termine gehalten")}` +
      kachel(d.schnittAnwesend == null ? "–" : komma(d.schnittAnwesend), "Ø anwesend") +
      kachel(
        d.quoteAmpel == null ? (d.quote == null ? "–" : `${Math.round(d.quote * 100)} %`) : `${Math.round(d.quoteAmpel * 100)} %`,
        `${d.ampel ? `<i class="ampel ${d.ampel}"></i>` : ""}Anwesenheit${d.ampel ? ` (${AMPEL_TEXT[d.ampel]})` : ""}`
      ) +
      kachel(d.schnittGesamt == null ? "–" : komma(d.schnittGesamt, 0), "Teilnehmer") +
      `</div>` +
      (d.balken.length ? `<div class="balken">${balken}</div>` : "") +
      (d.trend ? `<div class="trend ${d.trend.richtung}">${TREND[d.trend.richtung]} · Ø ${komma(d.trend.jetzt)} anwesend in den letzten 4 Terminen, davor ${komma(d.trend.vorher)}</div>` : "") +
      (d.ausfaelle.length
        ? `<div class="d-ausfall"><b>Ausgefallen:</b> ${d.ausfaelle.map((x) => `${tagDe(x.start)} ${html(ausfallText(x))}`).join(" · ")}</div>`
        : "") +
      `<div class="leise" style="margin-top:8px">${d.mitProtokoll} von ${d.gehalten} mit Teilnehmerzahl. Ampel aus den letzten 8 Protokollen, ab 3: ab 70 % gut, unter 55 % kritisch.</div>`
    : `<div class="leise">Noch keine gehaltenen Termine in den letzten vier Monaten.</div>`;
  const verlauf = d.verlauf.length
    ? `<table class="verlauf">${d.verlauf
        .map(
          (v) =>
            `<tr><td>${tagDe(v.t.start)}</td><td>${html(v.p.thema || "ohne Thema")}` +
            (v.p.besonderheiten || /^ja/i.test(v.p.vorkommnis) ? ` <span class="marke-klein">Besonderheit</span>` : "") +
            (v.p.material ? `<br><span class="leise">${html(v.p.material)}</span>` : "") +
            `</td><td class="rechts">${v.anwesend != null ? `${v.anwesend} / ${v.gesamt}` : ""}</td></tr>`
        )
        .join("")}</table>` +
      (d.themen.some(([, n]) => n > 1) ? `<div class="chips">${d.themen.map(([th, n]) => `<span class="chip rand">${html(th)} ${n}</span>`).join("")}</div>` : "")
    : `<div class="leise">Noch keine Protokolle zu diesem Kurs.</div>`;
  const kommend = d.kommend.length
    ? d.kommend
        .map((v) => {
          const k = KUERZEL[leseKuerzel(v.t.titel)];
          return `<div class="d-zeile"><span>${tagName(v.t.start.slice(0, 10)).slice(0, 2)} ${tagDe(v.t.start)} · ${v.t.start.slice(11)}</span>${k ? `<span class="marke-klein">${k}</span>` : `<span class="leise">regulär</span>`}</div>`;
        })
        .join("")
    : `<div class="leise">In den nächsten fünf Wochen kein weiterer Termin.</div>`;

  $("d-inhalt").className = "";
  $("d-inhalt").style.paddingTop = "0";
  $("d-inhalt").innerHTML =
    `<div class="d-teil"><h3>Dieser Termin</h3>${dieser}</div>` +
    `<div class="d-teil"><h3>Wie läuft der Kurs?</h3>${wie}</div>` +
    `<div class="d-teil"><h3>Was habe ich gemacht?</h3>${verlauf}</div>` +
    `<div class="d-teil"><h3>Nächste Termine</h3>${kommend}</div>`;
}

// --- Kurs verschieben ---
const minuten = (hhmm) => Number(hhmm.slice(0, 2)) * 60 + Number(hhmm.slice(3, 5));
const uhr = (m) => `${String(Math.floor(m / 60) % 24).padStart(2, "0")}:${String(m % 60).padStart(2, "0")}`;
let anbieterKopie = null;
async function richteVerschiebenEin(t) {
  const dauer = minuten(t.ende.slice(11)) - minuten(t.start.slice(11));
  $("vs-datum").value = t.start.slice(0, 10);
  $("vs-beginn").value = t.start.slice(11);
  $("vs-ende").value = t.ende.slice(11);
  $("vs-beginn").oninput = () => $("vs-beginn").value && ($("vs-ende").value = uhr(minuten($("vs-beginn").value) + dauer));
  $("vs-hinweis").textContent = "";
  $("vs-anbieter-zeile").hidden = true;
  anbieterKopie = null;
  // Steht der Kurs auch im Anbieter-Kalender? Dann dort mitverschieben (Standard).
  const a = kursAnbieter(t, einstellungen);
  if (a.kalender) {
    try {
      const tag = t.start.slice(0, 10);
      anbieterKopie =
        (await api.holeAnbieterKurse(tag, tag)).find(
          (k) => k.kalender === a.kalender && ohneKuerzel(k.titel) === ohneKuerzel(t.titel).replace(/^\(V\)\s*/, "") && k.start === t.start
        ) || null;
    } catch {
      anbieterKopie = null;
    }
  }
  if (anbieterKopie) {
    $("vs-anbieter-text").textContent = `Auch im Kalender „${a.kalender}“ verschieben`;
    $("vs-anbieter").checked = true;
    $("vs-anbieter-zeile").hidden = false;
  }
  $("vs-los").onclick = () =>
    mitWarten($("vs-los"), async () => {
      const neu = { datum: $("vs-datum").value, beginn: $("vs-beginn").value, ende: $("vs-ende").value };
      if (!neu.datum || !neu.beginn || !neu.ende || neu.ende <= neu.beginn) {
        $("vs-hinweis").innerHTML = `<span class="fehler">Bitte Tag, Beginn und Ende prüfen.</span>`;
        return;
      }
      const mit = anbieterKopie && $("vs-anbieter").checked ? anbieterKopie : null;
      const v = await verschiebeKurs(api, t, neu, einstellungen, { anbieterKopie: mit });
      zeige(
        `<div class="karte"><strong>Kurs verschoben</strong><div>${html(ohneKuerzel(t.titel))}: ${tagName(neu.datum)}, ${tagDe(neu.datum)} ${neu.beginn}–${neu.ende}.</div>` +
          `<div class="leise">Fahrten und Vor-/Nachbereitung sind angepasst${mit ? ", der Anbieter-Kalender auch" : ""}.</div></div>`
      );
      $("leiste").hidden = true;
      ANSICHTEN[aktuell]();
    });
}
function schliesseLeiste() {
  $("leiste").hidden = true;
  ANSICHTEN[aktuell]();
  if (aktuell !== "protokolle") ladeOffeneProtokolle().then((o) => zeigeAnzahl(o.length)).catch(() => {});
}
$("leiste-zu").onclick = schliesseLeiste;

// --- Vertretung übernehmen ---
async function zeigeVertretungen() {
  const heute = jetztLokal().slice(0, 10);
  const bis = plusTage(heute, 28);
  $("v-liste").innerHTML = `<div class="leise">Wird geladen …</div>`;
  try {
    const [anbieter, eigene] = await Promise.all([api.holeAnbieterKurse(heute, bis), api.holeVerlauf(heute, bis)]);
    const schon = new Set(eigene.map((t) => `${t.start} ${ohneKuerzel(t.titel)}`));
    const offen = anbieter.filter(
      (t) => t.start >= jetztLokal() && !schon.has(`${t.start} ${ohneKuerzel(t.titel)}`) && !["U", "A", "F"].includes(leseKuerzel(t.titel))
    );
    if (!offen.length) {
      $("v-liste").innerHTML = `<div class="leer">Alle Kurse der nächsten vier Wochen stehen schon in deinem Kalender.</div>`;
      return;
    }
    let letzterTag = "";
    $("v-liste").innerHTML = offen
      .map((t, i) => {
        const d = t.start.slice(0, 10);
        const b = beschreibe(t);
        const kopf = d !== letzterTag ? `<div class="tag">${tagName(d)}, ${tagDe(d)}</div>` : "";
        letzterTag = d;
        return (
          kopf +
          `<div class="kurs"><div class="streifen" style="background:${FARBEN[b.a.kuerzel] || "#5B787D"}"></div>` +
          `<div class="zeit">${t.start.slice(11)}–${t.ende.slice(11)}</div>` +
          `<div class="was">${html(b.name)}<small>${html(b.wo)}</small></div>` +
          `<div class="knoepfe"><button class="knopf" data-vertrete="${i}">Ich vertrete</button><button class="knopf" data-gebe="${i}">Ich gebe ihn</button></div></div>`
        );
      })
      .join("");
    const los = (t, vertretung) => {
      zustand.termin = t;
      uebernehme(vertretung);
      setTimeout(zeigeVertretungen, 4000);
    };
    document.querySelectorAll("[data-vertrete]").forEach((b) => (b.onclick = () => los(offen[b.dataset.vertrete], true)));
    document.querySelectorAll("[data-gebe]").forEach((b) => (b.onclick = () => los(offen[b.dataset.gebe], false)));
  } catch (e) {
    $("v-liste").innerHTML = `<div class="fehler">Konnte die Anbieter-Kalender nicht laden: ${html(e.message)}</div>`;
  }
}
$("v-extern").onclick = () => {
  oeffne("neu");
  $("n-anbieter").value = einstellungen.vertretungExtern.kuerzel;
  $("n-anbieter").dispatchEvent(new Event("change"));
};

// --- Fahrten fürs Finanzamt (km über OpenStreetMap, direkt aus dem Browser) ---
// Die Heimadresse steht in deinem Signatur-Entwurf und verlässt dein Gerät nur
// Richtung OpenStreetMap, wie bei der Automatik.
function routen() {
  const koordinaten = new Map();
  let letzte = 0;
  const finde = async (ort) => {
    for (const v of adressVarianten(ort)) {
      if (!koordinaten.has(v)) {
        await new Promise((r) => setTimeout(r, Math.max(0, 1100 - (Date.now() - letzte))));
        letzte = Date.now();
        const a = await fetch(`https://nominatim.openstreetmap.org/search?format=json&limit=1&countrycodes=de&q=${encodeURIComponent(v)}`);
        const t = a.ok ? (await a.json())[0] : null;
        koordinaten.set(v, t ? { lat: t.lat, lon: t.lon } : null);
      }
      if (koordinaten.get(v)) return koordinaten.get(v);
    }
    return null;
  };
  return async (von, nach) => {
    const [a, b] = [await finde(von), await finde(nach)];
    if (!a || !b) return null;
    const r = await fetch(`https://router.project-osrm.org/route/v1/driving/${a.lon},${a.lat};${b.lon},${b.lat}?overview=false`);
    const route = r.ok ? (await r.json()).routes?.[0] : null;
    return route ? { km: Math.round(route.distance / 100) / 10, minuten: Math.ceil(route.duration / 60) } : null;
  };
}

// Strecken (km) für alle Fahrten in "termine", einmal je Paar berechnet.
// Liefert strecke(von, nach) für die Auswertungen; ohne Heimadresse ohne km.
const streckenSpeicher = new Map();
async function streckenFuer(termine, stand, hinweis = () => {}) {
  // km ab Zuhause aus den Einstellungen (wie im Kursbuch) gehen vor, dann OpenStreetMap.
  const fest = (v, n) => {
    const k = v === "heim" ? n : n === "heim" ? v : null;
    const a = k && einstellungen.anbieter.find((x) => x.kuerzel === k && x.kmAbZuhause > 0);
    return a ? { km: a.kmAbZuhause, minuten: a.fahrzeitMin } : null;
  };
  let heim = null;
  try {
    heim = await api.heimAdresse();
  } catch {
    /* ohne Adresse eben nur die festen km */
  }
  if (heim) {
    hinweis("Strecken werden berechnet …");
    const rechne = routen();
    const ortVon = (k) => (k === "heim" ? heim : einstellungen.anbieter.find((a) => a.kuerzel === k)?.ort ?? (/^[A-Z]{2,4}$/.test(k) ? null : k));
    for (const a of alleZiele(einstellungen))
      for (const f of erstelleAuswertung(termine, a, stand, einstellungen).fahrten) {
        const paar = `${f.von} ${f.nach}`;
        if (fest(f.von, f.nach) || streckenSpeicher.has(paar) || !ortVon(f.von) || !ortVon(f.nach)) continue;
        streckenSpeicher.set(paar, await rechne(ortVon(f.von), ortVon(f.nach)).catch(() => null));
      }
  }
  const strecke = (v, n) => fest(v, n) ?? streckenSpeicher.get(`${v} ${n}`) ?? null;
  const fehlt = [...streckenSpeicher].filter(([, r]) => !r).length;
  return { heim, strecke, fehlt };
}

function richteFahrtenEin() {
  const jahr = Number(jetztLokal().slice(0, 4));
  $("fa-jahr").innerHTML = [jahr, jahr - 1, jahr - 2].map((j) => `<option>${j}</option>`).join("");
  // Im Januar ist meist das Vorjahr gemeint.
  if (jetztLokal().slice(5, 7) === "01") $("fa-jahr").value = String(jahr - 1);
  $("fa-pdf").onclick = () =>
    mitWarten($("fa-pdf"), async () => {
      const j = $("fa-jahr").value;
      $("fa-hinweis").textContent = "Kalender wird gelesen …";
      const termine = await api.holeVerlauf(`${j}-01-01`, `${j}-12-31`);
      const stand = jetztLokal();
      const { heim, strecke } = await streckenFuer(termine, stand, (t) => ($("fa-hinweis").textContent = t));
      const f = fahrtenJahr(termine, j, stand, einstellungen, strecke);
      fahrtenPdf(f, j, einstellungen, window.jspdf.jsPDF, stand).save(`Fahrten ${j}.pdf`);
      $("fa-hinweis").textContent =
        `${f.summe.fahrten} Fahrten an ${f.summe.tage} Tagen, ${String(f.summe.km).replace(".", ",")} km.` +
        (f.summe.ohneKm ? ` Bei ${f.summe.ohneKm} Tagen fehlen die km${heim ? "" : " (keine Adresse im Signatur-Entwurf)"}: unter ⚙️ Einstellungen beim Anbieter „km ab Zuhause“ eintragen.` : "");
    });
}

// --- Start ---
async function start() {
  $("jahr").textContent = jetztLokal().slice(0, 4);
  await msalApp.initialize();
  try {
    konto = (await msalApp.handleRedirectPromise())?.account;
  } catch (e) {
    $("anmelden-fehler").textContent = `Anmeldung hat nicht geklappt: ${e.message}`;
  }
  konto ??= msalApp.getAllAccounts()[0];
  if (!konto) {
    $("anmelden-seite").hidden = false;
    $("anmelden").onclick = () => msalApp.loginRedirect({ scopes: config.scopes, prompt: "select_account" });
    return;
  }
  api.nutzeAnmeldung(token);
  // Eigene Einstellungen aus dem Postfach, bevor irgendetwas geplant wird.
  try {
    wendeAn(einstellungen, await api.leseEinstellungen());
    einstellungen.jetzt = jetztLokal; // keine Erinnerung für schon vorbeigegangene Abfahrten
  } catch {
    /* dann eben die Standardwerte */
  }
  $("wer").textContent = konto.name || konto.username;
  $("haupt").hidden = false;
  await holeBereiche();
  richteNeuerKursEin();
  richteUrlaubEin();
  richteAuswertungEin();
  richteFahrtenEin();
  document.querySelectorAll("nav [data-ansicht], #unten [data-ansicht]").forEach((b) => (b.onclick = () => oeffne(b.dataset.ansicht)));
  $("unten").hidden = false;
  $("mehr").onclick = () => document.querySelector("nav").classList.toggle("offen");
  document.querySelectorAll(".abmelden").forEach(
    (b) =>
      (b.onclick = () => {
        if (confirm("Vom Kursplaner abmelden? Beim nächsten Öffnen meldest du dich neu an.")) {
          msalApp.logoutRedirect({ account: konto, postLogoutRedirectUri: location.origin + location.pathname });
        }
      })
  );
  // Start immer mit „Heute“; die Zahl offener Protokolle steht im Menü.
  oeffne("heute");
  ladeOffeneProtokolle().then((o) => zeigeAnzahl(o.length)).catch(() => {});
}
start().catch((e) => zeige(`<div class="karte fehler">Das hat nicht geklappt: ${html(e.message)}</div>`));

if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(() => {});

// --- Einstellungen: Zeiten, Anbieter, Material, E-Mail-Vorlagen, Monatsbericht ---
// Gespeichert in Michaels Postfach (src/eigeneEinstellungen.js), nicht im Code.
const ZEITEN = [
  ["vorbereitungMin", "Vorbereitung vor dem Kurs (Min.)"],
  ["nachbereitungMin", "Nachbereitung nach dem Kurs (Min.)"],
  ["maxLueckeAmStueckMin", "Kurse gelten als „am Stück“ bis zu einer Lücke von (Min.)"],
  ["pauseHeimfahrtMin", "Heimfahren zwischen zwei Kursen erst ab einer Pause von (Min.)"],
  ["fahrzeitRundungMin", "Berechnete Fahrzeit aufrunden auf (Min.)"],
  ["erinnerungVorFahrtMin", "Erinnerung aufs Handy: Minuten vor der Abfahrt (0 = genau zur Abfahrt, -1 = keine)"],
];
const feld = (id, label, wert, typ = "text") =>
  `<label for="${id}">${label}</label>` +
  (typ === "textarea" ? `<textarea id="${id}" rows="6">${html(wert ?? "")}</textarea>` : `<input id="${id}" type="${typ}" value="${html(wert ?? "")}">`);
const liste = (text) => text.split(",").map((x) => x.trim()).filter(Boolean);

function anbieterKarte(a, i) {
  return (
    `<div class="e-karte" data-anbieter="${i}">` +
    `<div class="reihe"><div>${feld(`ea-kuerzel-${i}`, "Kürzel (z. B. SVE)", a.kuerzel)}</div><div>${feld(`ea-name-${i}`, "Name", a.name)}</div></div>` +
    feld(`ea-ort-${i}`, "Adresse", a.ort) +
    `<div class="reihe"><div>${feld(`ea-fahrzeit-${i}`, "Fahrzeit (Min.)", a.fahrzeitMin, "number")}</div><div>${feld(`ea-kalender-${i}`, "Anbieter-Kalender in Outlook", a.kalender)}</div></div>` +
    feld(`ea-kategorien-${i}`, "Kategorien der Kurse (mit Komma getrennt)", (a.kategorien || []).join(", ")) +
    feld(`ea-vertretung-${i}`, "Kategorie für Vertretungen", a.kategorieVertretung) +
    feld(`ea-erkennung-${i}`, "Woran der Ort erkannt wird (Wörter aus der Adresse, mit Komma)", (a.ortErkennung || []).join(", ")) +
    `<div class="drei"><div>${feld(`ea-satz-${i}`, "Satz je Kurs (€)", a.satzJeKurs || "", "number")}</div><div>${feld(`ea-gjahr-${i}`, "Grenze im Jahr (€)", a.grenzeJahr || "", "number")}</div><div>${feld(`ea-gmonat-${i}`, "Grenze im Monat (€)", a.grenzeMonat || "", "number")}</div></div>` +
    feld(`ea-km-${i}`, "km ab Zuhause (einfache Strecke)", a.kmAbZuhause || "", "number") +
    `<label class="haken-zeile"><input id="ea-ferien-${i}" type="checkbox"${a.ferienFrei ? " checked" : ""}> In den Schulferien keine Kurse</label>` +
    `<label class="haken-zeile"><input id="ea-feiertag-${i}" type="checkbox"${a.feiertagFrei ? " checked" : ""}> An Feiertagen keine Kurse</label>` +
    `<button class="weg-knopf" data-weg-anbieter="${i}">Anbieter entfernen</button></div>`
  );
}
function vorlageKarte(v, i) {
  return (
    `<div class="e-karte" data-vorlage-nr="${i}">` +
    `<div class="reihe"><div>${feld(`ev-knopf-${i}`, "Knopf", v.knopf)}</div><div><label for="ev-an-${i}">Empfänger</label>` +
    `<select id="ev-an-${i}"><option value="">leer lassen</option><option value="anbieter"${v.an === "anbieter" ? " selected" : ""}>gemerkter Empfänger des Anbieters</option></select></div></div>` +
    feld(`ev-betreff-${i}`, "Betreff", v.betreff) +
    feld(`ev-text-${i}`, "Text", v.text, "textarea") +
    `<button class="weg-knopf" data-weg-vorlage="${i}">Vorlage entfernen</button></div>`
  );
}

let entwurf; // Arbeitskopie der Einstellungen in der Ansicht
function zeigeEinstellungen() {
  entwurf ??= {
    ...Object.fromEntries(ZEITEN.map(([k]) => [k, einstellungen[k]])),
    anbieter: einstellungen.anbieter.map((a) => ({ ...a })),
    material: [...einstellungen.material],
    vorlagen: einstellungen.vorlagen.map((v) => ({ ...v })),
    monatsbericht: { ...einstellungen.monatsbericht },
  };
  const e = entwurf;
  $("e-inhalt").className = "";
  $("e-inhalt").innerHTML =
    `<div class="e-block"><h2>Zeiten</h2>${ZEITEN.map(([k, l]) => feld(`ez-${k}`, l, e[k], "number")).join("")}</div>` +
    `<div class="e-block"><h2>Anbieter</h2><div class="leise">Feste Anbieter mit eigenem Kalender. Vertretungen in anderen Vereinen brauchen keinen Eintrag.</div>` +
    e.anbieter.map(anbieterKarte).join("") +
    `<button class="zweit" id="e-anbieter-neu">Anbieter hinzufügen</button></div>` +
    `<div class="e-block"><h2>Material</h2><div class="leise">Zum Anhaken im Kursprotokoll, eins pro Zeile.</div>` +
    `<textarea id="e-material" rows="8">${html(e.material.join("\n"))}</textarea></div>` +
    `<div class="e-block"><h2>E-Mail-Vorlagen</h2><div class="leise">Platzhalter: {kurs}, {wochentag}, {datum}, {kurzdatum}, {zeit}, {anbieter}</div>` +
    e.vorlagen.map(vorlageKarte).join("") +
    `<button class="zweit" id="e-vorlage-neu">Vorlage hinzufügen</button></div>` +
    `<div class="e-block"><h2>Monatsauswertung</h2>` +
    feld("em-tag", "Entwurf kommt am Tag des Monats", e.monatsbericht.tag, "number") +
    feld("em-betreff", "Betreff ({monat}, {anbieter})", e.monatsbericht.betreff) +
    feld("em-text", "Text an eine Person ({anrede}, {monat}, {nachtrag})", e.monatsbericht.text, "textarea") +
    feld("em-gruppe", "Text an mehrere Personen", e.monatsbericht.textGruppe, "textarea") +
    `</div>` +
    `<button id="e-speichern">Einstellungen speichern</button>`;

  $("e-anbieter-neu").onclick = () => {
    uebernimmFormular();
    e.anbieter.push({ kuerzel: "", name: "", kalender: "", ort: "", ortErkennung: [], fahrzeitMin: 20, kategorien: [], kategorieVertretung: "", feiertagFrei: true, empfaenger: "" });
    zeigeEinstellungen();
  };
  $("e-vorlage-neu").onclick = () => {
    uebernimmFormular();
    e.vorlagen.push({ id: `eigene-${Date.now()}`, knopf: "Neue Vorlage", an: "", betreff: "{kurs} am {wochentag}, {kurzdatum}", text: "Hallo,\n\n\n\nViele Grüße\nMichael" });
    zeigeEinstellungen();
  };
  document.querySelectorAll("[data-weg-anbieter]").forEach((b) => (b.onclick = () => {
    if (!confirm("Diesen Anbieter entfernen? Seine Kurse bekommen dann keine Fahrten mehr.")) return;
    uebernimmFormular();
    e.anbieter.splice(Number(b.dataset.wegAnbieter), 1);
    zeigeEinstellungen();
  }));
  document.querySelectorAll("[data-weg-vorlage]").forEach((b) => (b.onclick = () => {
    uebernimmFormular();
    e.vorlagen.splice(Number(b.dataset.wegVorlage), 1);
    zeigeEinstellungen();
  }));
  $("e-speichern").onclick = () =>
    mitWarten($("e-speichern"), async () => {
      uebernimmFormular();
      const fehler = e.anbieter.some((a) => !a.kuerzel || !a.name || !a.kalender || !a.ort);
      if (fehler) {
        zeige(`<div class="karte fehler">Bei jedem Anbieter brauchen Kürzel, Name, Kalender und Adresse einen Eintrag.</div>`);
        return;
      }
      // Immer alles speichern (auch „ausbezahlt“ aus der Geld-Ansicht).
      await api.speichereEinstellungen({ ...auszug(einstellungen), ...e });
      wendeAn(einstellungen, e);
      zeige(`<div class="karte"><strong>Gespeichert</strong><div class="leise">Gilt ab sofort hier, im Add-in beim nächsten Öffnen und in der Automatik beim nächsten Lauf.</div></div>`);
    });
}

// Eingaben aus dem Formular in die Arbeitskopie übernehmen.
function uebernimmFormular() {
  const e = entwurf;
  const wert = (id) => $(id)?.value ?? "";
  for (const [k] of ZEITEN) if (wert(`ez-${k}`) !== "") e[k] = Number(wert(`ez-${k}`));
  e.anbieter = e.anbieter.map((a, i) => ({
    ...a,
    kuerzel: wert(`ea-kuerzel-${i}`).trim().toUpperCase(),
    name: wert(`ea-name-${i}`).trim(),
    ort: wert(`ea-ort-${i}`).trim(),
    fahrzeitMin: Number(wert(`ea-fahrzeit-${i}`)) || 0,
    kalender: wert(`ea-kalender-${i}`).trim(),
    kategorien: liste(wert(`ea-kategorien-${i}`)),
    kategorieVertretung: wert(`ea-vertretung-${i}`).trim(),
    ortErkennung: liste(wert(`ea-erkennung-${i}`)),
    ferienFrei: !!$(`ea-ferien-${i}`)?.checked,
    satzJeKurs: Number(wert(`ea-satz-${i}`)) || 0,
    grenzeJahr: Number(wert(`ea-gjahr-${i}`)) || 0,
    grenzeMonat: Number(wert(`ea-gmonat-${i}`)) || 0,
    kmAbZuhause: Number(wert(`ea-km-${i}`)) || 0,
    feiertagFrei: !!$(`ea-feiertag-${i}`)?.checked,
  }));
  e.material = wert("e-material").split("\n").map((x) => x.trim()).filter(Boolean);
  e.vorlagen = e.vorlagen.map((v, i) => ({ ...v, knopf: wert(`ev-knopf-${i}`).trim(), an: wert(`ev-an-${i}`), betreff: wert(`ev-betreff-${i}`), text: wert(`ev-text-${i}`) }));
  e.monatsbericht = { ...e.monatsbericht, tag: Number(wert("em-tag")) || e.monatsbericht.tag, betreff: wert("em-betreff"), text: wert("em-text"), textGruppe: wert("em-gruppe") };
}

// --- Auswertungen wie im Excel-Kursbuch: Kursbericht, Wie oft wo, Geld, Prüfliste ---
const jahrCache = new Map(); // ein Jahr Kalender, 5 Minuten gemerkt
async function ladeJahr(j) {
  const c = jahrCache.get(j);
  if (c && Date.now() - c.zeit < 300000) return c.termine;
  const termine = await api.holeVerlauf(`${j}-01-01`, `${j}-12-31`);
  jahrCache.set(j, { zeit: Date.now(), termine });
  return termine;
}
const euro = (x) => x.toLocaleString("de-DE", { style: "currency", currency: "EUR" });
const prozent = (x) => (x == null ? "–" : `${Math.round(x * 100)} %`);
const zahlDe = (x, n = 1) => (x == null ? "–" : x.toLocaleString("de-DE", { maximumFractionDigits: n }));
const datumDe = (d) => `${d.slice(8, 10)}.${d.slice(5, 7)}.${d.slice(0, 4)}`;
const ampelPunkt = (a) => (a ? `<i class="ampel ${a}"></i>${a}` : "");
function jahrPillen(ziel, gewaehlt, weiter) {
  const j = Number(jetztLokal().slice(0, 4));
  ziel.innerHTML = [j, j - 1].map((x) => `<button class="pille${x === gewaehlt ? " an" : ""}" data-jahr="${x}">${x}</button>`).join("");
  ziel.querySelectorAll("[data-jahr]").forEach((b) => (b.onclick = () => weiter(Number(b.dataset.jahr))));
}
const ladeFehler = (ziel, e) => (ziel.innerHTML = `<div class="karte fehler">Konnte den Kalender nicht laden: ${html(e.message)}</div>`);

// 1. Kursbericht
let kbEingerichtet = false;
function richteKursberichtEin() {
  if (kbEingerichtet) return;
  kbEingerichtet = true;
  const heute = jetztLokal().slice(0, 10);
  const setze = (art) => {
    const j = heute.slice(0, 4);
    const m = heute.slice(0, 7);
    const vor = plusTage(`${m}-01`, -1);
    const [v, b] = art === "jahr" ? [`${j}-01-01`, heute] : art === "vormonat" ? [`${vor.slice(0, 7)}-01`, vor] : [`${m}-01`, heute];
    $("kb-von").value = v;
    $("kb-bis").value = b;
  };
  setze("monat");
  $("kb-anbieter").innerHTML =
    `<option value="">Alle</option>` + alleZiele(einstellungen).map((a) => `<option value="${a.kuerzel}">${html(a.kuerzel)} · ${html(a.name)}</option>`).join("");
  document.querySelectorAll("#kb-schnell [data-zeitraum]").forEach((b) => (b.onclick = () => (setze(b.dataset.zeitraum), zeigeKursbericht())));
  $("kb-zeigen").onclick = () => zeigeKursbericht();
  $("kb-pdf").onclick = () => mitWarten($("kb-pdf"), async () => kursberichtPdf(await rechneKursbericht()));
  zeigeKursbericht();
}
async function rechneKursbericht() {
  const von = $("kb-von").value;
  const bis = $("kb-bis").value;
  const termine = await api.holeVerlauf(von, bis);
  const opt = { von, bis, anbieter: $("kb-anbieter").value, nurEigene: $("kb-eigene").checked, jetzt: jetztLokal() };
  return { ...kursbericht(termine, opt, einstellungen), ...opt, besonders: $("kb-besonders").checked };
}
async function zeigeKursbericht() {
  const ziel = $("kb-ergebnis");
  ziel.innerHTML = `<div class="leise" style="margin-top:12px">Wird geladen …</div>`;
  let k;
  try {
    k = await rechneKursbericht();
  } catch (e) {
    return ladeFehler(ziel, e);
  }
  const kachel = (w, t) => `<div class="kachel"><b>${w}</b><span>${t}</span></div>`;
  ziel.innerHTML =
    `<div class="kacheln fuenf">${kachel(k.mitProtokoll, "mit Protokoll")}${kachel(zahlDe(k.schnittAnwesend), "Ø anwesend")}` +
    `${kachel(prozent(k.quote), "Anwesenheit")}${kachel(prozent(k.unentschuldigt), "unentschuldigt")}${kachel(k.ohneProtokoll, "ohne Protokoll")}</div>` +
    (k.kurse.length
      ? `<div class="tabelle-rahmen"><table><tr><th>Kurs</th><th>Tag</th><th class="z">Termine</th><th class="z">Ø Teiln.</th><th class="z">Ø anwesend</th><th class="z">Anwesenheit</th><th class="z">unentsch.</th><th>Ampel</th></tr>` +
        k.kurse
          .map(
            (x) =>
              `<tr><td>${html(x.name)}<br><span class="leise">${x.anbieter}</span></td><td>${x.tag} ${x.beginn}</td><td class="z">${x.termine}</td>` +
              `<td class="z">${zahlDe(x.schnittGesamt, 0)}</td><td class="z">${zahlDe(x.schnittAnwesend)}</td><td class="z">${prozent(x.quote)}</td><td class="z">${prozent(x.unentschuldigt)}</td><td>${ampelPunkt(x.ampel)}</td></tr>`
          )
          .join("") +
        `</table></div>`
      : `<div class="leer" style="margin-top:12px">In diesem Zeitraum keine gehaltenen Kurse.</div>`) +
    (k.themen.length ? `<div class="karte"><h2 style="font-size:20px;margin:0 0 8px">Themen</h2><div class="chips">${k.themen.map(([t, n]) => `<span class="chip rand">${html(t)} ${n}</span>`).join("")}</div></div>` : "") +
    (k.material.length ? `<div class="karte"><h2 style="font-size:20px;margin:0 0 8px">Material</h2><div class="chips">${k.material.map(([t, n]) => `<span class="chip rand">${html(t)} ${n}</span>`).join("")}</div></div>` : "") +
    (k.besonders && k.besonderheiten.length
      ? `<div class="karte"><h2 style="font-size:20px;margin:0 0 8px">Besonderheiten</h2>${k.besonderheiten
          .map((b) => `<div style="margin:8px 0"><b>${tagDe(b.datum)} ${html(b.kurs)}</b>${b.vorkommnis ? ` <span class="marke-klein">Vorkommnis</span>` : ""}<div>${html(b.text || "")}</div></div>`)
          .join("")}</div>`
      : k.besonderheiten.length
        ? `<div class="leise">${k.besonderheiten.length} Besonderheiten notiert (zum Anzeigen oben ankreuzen).</div>`
        : "");
}
function kursberichtPdf(k) {
  const doc = new window.jspdf.jsPDF({ unit: "mm", format: "a4" });
  let y = 18;
  const zeile = (text, groesse = 10, fett = false) => {
    if (y > 280) {
      doc.addPage();
      y = 18;
    }
    doc.setFont("helvetica", fett ? "bold" : "normal");
    doc.setFontSize(groesse);
    const teile = doc.splitTextToSize(text, 180);
    doc.text(teile, 15, y);
    y += teile.length * groesse * 0.45 + 1.5;
  };
  const anb = k.anbieter ? alleZiele(einstellungen).find((a) => a.kuerzel === k.anbieter)?.name : "alle Anbieter";
  zeile("Kursbericht", 18, true);
  zeile(`${datumDe(k.von)} bis ${datumDe(k.bis)} · ${anb}${k.nurEigene ? " · nur eigene Kurse" : ""} · Michael Stehle`, 10);
  y += 3;
  zeile(`Termine mit Protokoll: ${k.mitProtokoll}   Ø anwesend: ${zahlDe(k.schnittAnwesend)}   Anwesenheit: ${prozent(k.quote)}   unentschuldigt: ${prozent(k.unentschuldigt)}`, 10, true);
  y += 3;
  zeile("Anwesenheit je Kurs", 13, true);
  for (const x of k.kurse)
    zeile(`${x.anbieter} · ${x.name} · ${x.tag} ${x.beginn}:  ${x.termine} Termine, Ø ${zahlDe(x.schnittAnwesend)} von ${zahlDe(x.schnittGesamt, 0)} anwesend (${prozent(x.quote)}), unentschuldigt ${prozent(x.unentschuldigt)}${x.ampel ? `, ${x.ampel}` : ""}`);
  y += 3;
  if (k.themen.length) {
    zeile("Themen", 13, true);
    zeile(k.themen.map(([t, n]) => `${t} (${n})`).join(", "));
  }
  if (k.material.length) {
    zeile("Material", 13, true);
    zeile(k.material.map(([t, n]) => `${t} (${n})`).join(", "));
  }
  if (k.besonders && k.besonderheiten.length) {
    y += 3;
    zeile("Besonderheiten (vertraulich)", 13, true);
    for (const b of k.besonderheiten) zeile(`${datumDe(b.datum)} ${b.kurs}${b.vorkommnis ? " (Vorkommnis)" : ""}: ${b.text || ""}`);
  }
  doc.save(`Kursbericht ${k.von} bis ${k.bis}${k.anbieter ? ` ${k.anbieter}` : ""}.pdf`);
}

// 2. Wie oft wo
let woJahr;
async function zeigeWieOft(j = (woJahr ??= Number(jetztLokal().slice(0, 4)))) {
  woJahr = j;
  jahrPillen($("wo-jahre"), j, zeigeWieOft);
  const ziel = $("wo-ergebnis");
  ziel.innerHTML = `<div class="leise" style="padding:14px">Wird geladen …</div>`;
  let w;
  try {
    w = wieOftWo(await ladeJahr(j), j, jetztLokal(), einstellungen);
  } catch (e) {
    return ladeFehler(ziel, e);
  }
  const kurz = { Regulär: "regulär", Vertretung: "Vertr.", Hospitation: "Hosp.", Urlaub: "Urlaub", Abgesagt: "abges.", gehalten: "gehalten" };
  const spalten = [...WIE_OFT_ARTEN, "gehalten"];
  const zellen = (je) => w.anbieter.map((k) => spalten.map((a) => `<td class="z${je[k][a] ? "" : " null"}">${a === "gehalten" ? `<b>${je[k][a]}</b>` : je[k][a]}</td>`).join("")).join("");
  ziel.innerHTML =
    `<table><tr><th></th>${w.anbieter.map((k) => `<th colspan="${spalten.length}" style="text-align:center">${k}</th>`).join("")}<th></th></tr>` +
    `<tr><th>Monat</th>${w.anbieter.map(() => spalten.map((a) => `<th class="z">${kurz[a]}</th>`).join("")).join("")}<th class="z">gesamt</th></tr>` +
    w.monate.map((m) => `<tr><td>${m.name}</td>${zellen(m.je)}<td class="z"><b>${m.gehalten}</b></td></tr>`).join("") +
    `<tr class="summe"><td>Summe</td>${zellen(w.summe.je)}<td class="z">${w.summe.gehalten}</td></tr></table>`;
}

// 3. Geld
let gJahr;
const kmCache = new Map();
async function zeigeGeld(j = (gJahr ??= Number(jetztLokal().slice(0, 4)))) {
  gJahr = j;
  jahrPillen($("g-jahre"), j, zeigeGeld);
  const ziel = $("g-ergebnis");
  if (!einstellungen.anbieter.some((a) => a.satzJeKurs > 0)) return zeigeSaetzeEingeben(ziel);
  ziel.innerHTML = `<div class="leise">Wird geladen …</div>`;
  let termine;
  try {
    termine = await ladeJahr(j);
  } catch (e) {
    return ladeFehler(ziel, e);
  }
  const zeichne = (kmJe, kmText) => {
    if (gJahr !== j || aktuell !== "geld") return;
    const liste = geld(termine, j, jetztLokal(), einstellungen, kmJe);
    ziel.innerHTML =
      liste
        .map((a) => {
          const anteil = a.anteilJahr;
          const balken = anteil == null ? "" : `<div class="grenze"><i class="${anteil > 1 ? "ueber" : anteil > 0.85 ? "warn" : ""}" style="width:${Math.min(100, anteil * 100)}%"></i></div>` +
            `<div class="leise">${euro(a.summe.betrag)} von ${euro(a.grenzeJahr)} im Jahr (inkl. geplanter Kurse) · noch ${euro(Math.max(0, a.grenzeJahr - a.summe.betrag))}</div>`;
          return (
            `<div class="karte"><h2 style="font-size:22px;margin:0">${html(a.kuerzel)} · ${html(a.name)}</h2>` +
            `<div class="leise">${euro(a.satz)} je Kurs${a.grenzeMonat ? ` · Monatsgrenze ${euro(a.grenzeMonat)}` : ""}</div>${balken}` +
            `<div class="tabelle-rahmen"><table><tr><th>Monat</th><th class="z">Kurse</th><th class="z">Betrag</th><th class="z">km</th><th class="z">Fahrtkosten</th><th>ausbezahlt</th></tr>` +
            a.monate
              .filter((m) => m.kurse || m.km)
              .map(
                (m) =>
                  `<tr${m.erbracht ? "" : ' class="geplant"'}><td>${m.name}</td><td class="z">${m.kurse}${m.geplant ? `<br><span class="leise" style="font-size:12px">${m.geplant} geplant</span>` : ""}</td>` +
                  `<td class="z">${euro(m.betrag)}${m.ueberMonatsgrenze ? ` <span class="marke-klein">über Grenze</span>` : ""}</td><td class="z">${zahlDe(m.km, 0)}</td><td class="z">${m.fahrtkosten == null ? "–" : euro(m.fahrtkosten)}</td>` +
                  `<td><label class="haken-zeile" style="margin:0"><input type="checkbox" data-ausbezahlt="${m.schluessel}"${m.ausbezahlt ? " checked" : ""}></label></td></tr>`
              )
              .join("") +
            `<tr class="summe"><td>Summe</td><td class="z">${a.summe.kurse}</td><td class="z">${euro(a.summe.betrag)}</td><td class="z">${zahlDe(a.summe.km, 0)}</td><td class="z">${a.summe.fahrtkosten == null ? "–" : euro(a.summe.fahrtkosten)}</td><td>${euro(a.summe.ausbezahlt)}</td></tr>` +
            `</table></div></div>`
          );
        })
        .join("") +
      `<div class="leise">${kmText} Sätze und Grenzen änderst du unter ⚙️ Einstellungen beim Anbieter.</div>`;
    ziel.querySelectorAll("[data-ausbezahlt]").forEach(
      (c) =>
        (c.onchange = async () => {
          const neu = { ...(einstellungen.ausbezahlt || {}) };
          if (c.checked) neu[c.dataset.ausbezahlt] = true;
          else delete neu[c.dataset.ausbezahlt];
          c.disabled = true;
          try {
            await api.speichereEinstellungen({ ...auszug(einstellungen), ausbezahlt: neu });
            einstellungen.ausbezahlt = neu;
            zeichne(kmJe, kmText);
          } catch (e) {
            c.checked = !c.checked;
            c.disabled = false;
            zeige(`<div class="karte fehler">Nicht gespeichert: ${html(e.message)}</div>`);
          }
        })
    );
  };
  const kmFertig = kmCache.get(j);
  if (kmFertig) return zeichne(kmFertig.kmJe, kmFertig.text);
  zeichne(() => null, "km werden berechnet …");
  // km wie bei den Fahrten fürs Finanzamt, dann neu zeichnen.
  try {
    const stand = jetztLokal();
    const { heim, strecke, fehlt } = await streckenFuer(termine, stand);
    const f = fahrtenJahr(termine, j, stand, einstellungen, strecke);
    const kmJe = (k, monat) => {
      const r = f.reihen.filter((x) => x.anbieter === k && Number(x.datum.slice(5, 7)) === monat);
      return r.some((x) => x.km != null) ? r.reduce((s, x) => s + (x.km ?? 0), 0) : null;
    };
    const festeKm = einstellungen.anbieter.some((a) => a.kmAbZuhause > 0);
    const text =
      f.summe.ohneKm || !f.reihen.length
        ? `Bei ${f.summe.ohneKm} Fahrtagen fehlen die km${heim ? (fehlt ? " (OpenStreetMap hat nicht geantwortet)" : "") : " (keine Adresse im Signatur-Entwurf)"}. ` +
          `Am einfachsten: unter ⚙️ Einstellungen beim Anbieter „km ab Zuhause“ eintragen.`
        : `km ${festeKm ? "aus deinen Einstellungen" : "laut OpenStreetMap"}, Fahrtkosten mit ${euro(einstellungen.kmSatz)} je km.`;
    kmCache.set(j, { kmJe, text });
    zeichne(kmJe, text);
  } catch {
    zeichne(() => null, "km konnten nicht berechnet werden.");
  }
}
function zeigeSaetzeEingeben(ziel) {
  ziel.innerHTML =
    `<div class="karte"><h2 style="font-size:22px;margin:0 0 4px">Sätze eintragen</h2>` +
    `<div class="leise">Einmal eintragen, dann rechnet der Kursplaner selbst. Die Werte liegen nur in deinem Postfach.</div>` +
    einstellungen.anbieter
      .map(
        (a, i) =>
          `<h3 style="margin:16px 0 0">${html(a.kuerzel)} · ${html(a.name)}</h3><div class="drei">` +
          `<div><label for="gs-satz-${i}">Satz je Kurs (€)</label><input id="gs-satz-${i}" type="number" step="0.5" min="0"></div>` +
          `<div><label for="gs-jahr-${i}">Grenze im Jahr (€)</label><input id="gs-jahr-${i}" type="number" step="50" min="0" placeholder="z. B. Übungsleiterpauschale"></div>` +
          `<div><label for="gs-monat-${i}">Grenze im Monat (€)</label><input id="gs-monat-${i}" type="number" step="10" min="0" placeholder="z. B. Minijob"></div></div>` +
          `<label for="gs-km-${i}">km ab Zuhause (einfache Strecke)</label><input id="gs-km-${i}" type="number" step="1" min="0">`
      )
      .join("") +
    `<label for="gs-km">Fahrtkosten je km (€)</label><input id="gs-km" type="number" step="0.01" min="0" value="${einstellungen.kmSatz}">` +
    `<button id="gs-speichern">Speichern</button></div>`;
  $("gs-speichern").onclick = () =>
    mitWarten($("gs-speichern"), async () => {
      const anbieter = einstellungen.anbieter.map((a, i) => ({
        ...a,
        satzJeKurs: Number($(`gs-satz-${i}`).value) || 0,
        grenzeJahr: Number($(`gs-jahr-${i}`).value) || 0,
        grenzeMonat: Number($(`gs-monat-${i}`).value) || 0,
        kmAbZuhause: Number($(`gs-km-${i}`).value) || 0,
      }));
      const neu = { ...auszug(einstellungen), anbieter, kmSatz: Number($("gs-km").value) || 0 };
      await api.speichereEinstellungen(neu);
      wendeAn(einstellungen, neu);
      entwurf = undefined;
      zeigeGeld();
    });
}

// 4. Prüfliste
async function zeigePruefliste() {
  const ziel = $("pl-ergebnis");
  ziel.innerHTML = `<div class="leise">Wird geladen …</div>`;
  const j = Number(jetztLokal().slice(0, 4));
  let p;
  try {
    p = pruefliste(await ladeJahr(j), j, jetztLokal(), einstellungen);
  } catch (e) {
    return ladeFehler(ziel, e);
  }
  // Grenzen aus „Geld“, wenn Sätze eingetragen sind.
  const knapp = geld(await ladeJahr(j), j, jetztLokal(), einstellungen).filter((a) => a.anteilJahr != null && a.anteilJahr > 0.85);
  const offen = p.punkte.reduce((s, x) => s + x.termine.length, 0) + knapp.length;
  $("pl-anzahl").textContent = offen ? String(offen) : "";
  $("pl-anzahl").hidden = !offen;
  $("pl-unter").textContent = `Was fehlt oder nicht passt, ${j} bis heute.` + (p.protokollAb ? ` Protokolle zählen ab ${datumDe(p.protokollAb)} (dein erstes Protokoll).` : "");
  ziel.innerHTML =
    p.punkte
      .map(
        (x, i) =>
          `<div class="karte"><div class="pl-punkt"><div><b>${x.titel}</b><div class="leise">${x.termine.length ? x.tun : "Alles in Ordnung."}</div></div>` +
          `<div class="gross${x.termine.length ? "" : " ok"}">${x.termine.length || "✓"}</div></div><div id="pl-liste-${i}"></div></div>`
      )
      .join("") +
    knapp.map((a) => `<div class="karte"><div class="pl-punkt"><div><b>${html(a.kuerzel)}: Grenze fast erreicht</b><div class="leise">${euro(a.summe.betrag)} von ${euro(a.grenzeJahr)} (${prozent(a.anteilJahr)}) inkl. geplanter Kurse.</div></div><div class="gross">!</div></div></div>`).join("") +
    (offen ? "" : `<div class="leer">Alles grün. 👍</div>`);
  p.punkte.forEach((x, i) => {
    if (x.termine.length) zeigeKurse([...x.termine].sort((a, b) => (a.start < b.start ? 1 : -1)), $(`pl-liste-${i}`), () => ((jahrCache.clear(), zeigePruefliste())), { knoepfe: ["protokoll"] });
  });
}
