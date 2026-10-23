// Kursplaner als eigenes Fenster (installierbar über Edge/Chrome): alles,
// was das Add-in kann, ohne dass in Outlook etwas markiert sein muss.
// Anmeldung mit MSAL (Weiterleitung), Daten direkt aus Outlook über Graph.
import * as api from "../web/graph.js";
import { config } from "../web/config.js";
import { einstellungen } from "../src/einstellungen.js";
import { kursAnbieter, leseKuerzel, ohneKuerzel, routenUrl } from "../src/planer.js";
import { leseProtokoll } from "../src/protokoll.js";
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
  !t.generiert && kursAnbieter(t, einstellungen) && !/^(Anfahrt|Rückfahrt|Fahrt |Vorbereitung|Nachbereitung|Nach-\/Vorbereitung)/.test(t.titel);
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
        `<div class="kurs${b.faelltAus ? " faellt-aus" : ""}">` +
        `<div class="streifen" style="background:${FARBEN[b.a.kuerzel] || "#5B787D"}"></div>` +
        `<div class="zeit">${t.start.slice(11)}–${t.ende.slice(11)}</div>` +
        `<div class="was">${html(b.name)}${KUERZEL[b.k] ? `<span class="marke-klein">${KUERZEL[b.k]}</span>` : ""}<small>${html(b.wo)}</small></div>` +
        `<div class="knoepfe">` +
        (knoepfe.includes("protokoll") ? `<button class="knopf" data-protokoll="${i}">Protokoll</button>` : "") +
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
  ziel.querySelectorAll("[data-protokoll]").forEach((b) => (b.onclick = () => oeffneLeiste(kurse[b.dataset.protokoll], "protokoll")));
  ziel.querySelectorAll("[data-mail]").forEach((b) => (b.onclick = () => oeffneLeiste(kurse[b.dataset.mail], "mail")));
  ziel.querySelectorAll("[data-verschieben]").forEach((b) => (b.onclick = () => oeffneLeiste(kurse[b.dataset.verschieben], "verschieben")));
  ziel.querySelectorAll("[data-kuerzel]").forEach((b) => {
    b.onclick = () => {
      const t = kurse[b.dataset.kurs];
      const absagen = b.dataset.kuerzel === "A";
      const frage = absagen
        ? `„${ohneKuerzel(t.titel)}“ am ${tagDe(t.start)} absagen? Fahrten und Vor-/Nachbereitung fallen dann weg.`
        : `„${ohneKuerzel(t.titel)}“ am ${tagDe(t.start)} findet doch statt? Fahrten und Vor-/Nachbereitung kommen wieder dazu.`;
      if (!confirm(frage)) return;
      mitWarten(b, async () => {
        const v = await setzeKuerzel(api, t, b.dataset.kuerzel, einstellungen);
        zeigeAenderungen(v, absagen ? "Kurs abgesagt" : "Kurs findet statt");
        await neuLaden();
      });
    };
  });
}

// --- Heute: nächste Abfahrt mit Navigation, Kurse von heute (oder vom nächsten Kurstag) ---
const ROUTE = /https:\/\/www\.google\.com\/maps\/dir\/[^\s<>"]+/;
async function zeigeHeute() {
  const heute = jetztLokal().slice(0, 10);
  $("h-unter").textContent = `${tagName(heute)}, ${tagDe(heute)}${heute.slice(0, 4)}`;
  $("h-fahrt").innerHTML = "";
  $("h-liste").innerHTML = `<div class="leise">Wird geladen …</div>`;
  let termine;
  try {
    termine = await api.holeVerlauf(heute, plusTage(heute, 14));
  } catch (e) {
    $("h-liste").innerHTML = `<div class="fehler">Konnte den Kalender nicht laden: ${html(e.message)}</div>`;
    return;
  }
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
  const naechster = kurse.find((t) => t.start.slice(0, 10) > heute && !beschreibe(t).faelltAus);
  if (!naechster) {
    $("h-titel").textContent = "Heute keine Kurse";
    $("h-liste").innerHTML = `<div class="leer">In den nächsten zwei Wochen stehen keine Kurse in deinem Kalender.</div>`;
    return;
  }
  const tag = naechster.start.slice(0, 10);
  $("h-titel").textContent = `Heute keine Kurse. Nächster Kurstag: ${tagName(tag)}, ${tagDe(tag)}`;
  zeigeKurse(kurse.filter((t) => t.start.slice(0, 10) === tag), $("h-liste"), zeigeHeute, { mitTag: false });
}

// --- Offene Protokolle: vergangene Kurse (60 Tage), die stattfanden, ohne Protokoll ---
const ohneProtokoll = (t) => {
  const p = leseProtokoll(t.text || "");
  return p.anwesend === "" && !p.thema;
};
async function ladeOffeneProtokolle() {
  const heute = jetztLokal().slice(0, 10);
  const jetzt = jetztLokal();
  const termine = await api.holeVerlauf(plusTage(heute, -60), heute);
  return termine
    .filter((t) => istKurs(t) && t.ende <= jetzt && !beschreibe(t).faelltAus && ohneProtokoll(t))
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
async function oeffneLeiste(t, was) {
  zustand.termin = t;
  $("leiste-titel").textContent = ohneKuerzel(t.titel);
  $("leiste-info").textContent = `${tagName(t.start.slice(0, 10))}, ${tagDe(t.start)} · ${t.start.slice(11)}–${t.ende.slice(11)}`;
  $("bereich-protokoll").hidden = was !== "protokoll";
  $("bereich-mail").hidden = was !== "mail";
  $("bereich-verschieben").hidden = was !== "verschieben";
  $("leiste").hidden = false;
  if (was === "protokoll") await zeigeProtokoll();
  else if (was === "mail") richteMailEin(kursAnbieter(t, einstellungen));
  else await richteVerschiebenEin(t);
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
$("leiste-zu").onclick = () => {
  $("leiste").hidden = true;
  ANSICHTEN[aktuell]();
  if (aktuell !== "protokolle") ladeOffeneProtokolle().then((o) => zeigeAnzahl(o.length)).catch(() => {});
};

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
      const heim = await api.heimAdresse();
      const strecken = new Map();
      if (heim) {
        $("fa-hinweis").textContent = "Strecken werden berechnet …";
        const strecke = routen();
        const ortVon = (k) => (k === "heim" ? heim : einstellungen.anbieter.find((a) => a.kuerzel === k)?.ort ?? (/^[A-Z]{2,4}$/.test(k) ? null : k));
        for (const a of alleZiele(einstellungen))
          for (const f of erstelleAuswertung(termine, a, stand, einstellungen).fahrten) {
            const paar = `${f.von} ${f.nach}`;
            if (strecken.has(paar) || !ortVon(f.von) || !ortVon(f.nach)) continue;
            strecken.set(paar, await strecke(ortVon(f.von), ortVon(f.nach)).catch(() => null));
          }
      }
      const f = fahrtenJahr(termine, j, stand, einstellungen, (v, n) => strecken.get(`${v} ${n}`) ?? null);
      fahrtenPdf(f, j, einstellungen, window.jspdf.jsPDF, stand).save(`Fahrten ${j}.pdf`);
      $("fa-hinweis").textContent =
        `${f.summe.fahrten} Fahrten an ${f.summe.tage} Tagen, ${String(f.summe.km).replace(".", ",")} km.` +
        (heim ? (f.summe.ohneKm ? ` Bei ${f.summe.ohneKm} Tagen fehlen die km.` : "") : " Ohne km: im Signatur-Entwurf steht keine Adresse.");
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
  ["erinnerungVorFahrtMin", "Erinnerung vor der Abfahrt (Min., 0 = aus)"],
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
      await api.speichereEinstellungen(e);
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
    feiertagFrei: !!$(`ea-feiertag-${i}`)?.checked,
  }));
  e.material = wert("e-material").split("\n").map((x) => x.trim()).filter(Boolean);
  e.vorlagen = e.vorlagen.map((v, i) => ({ ...v, knopf: wert(`ev-knopf-${i}`).trim(), an: wert(`ev-an-${i}`), betreff: wert(`ev-betreff-${i}`), text: wert(`ev-text-${i}`) }));
  e.monatsbericht = { ...e.monatsbericht, tag: Number(wert("em-tag")) || e.monatsbericht.tag, betreff: wert("em-betreff"), text: wert("em-text"), textGruppe: wert("em-gruppe") };
}
