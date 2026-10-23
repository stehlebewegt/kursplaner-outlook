// Outlook-Add-in: Bereiche je nach geöffnetem Termin. Die Bereiche selbst
// stehen in bereiche.js (gemeinsam mit dem Kursplaner-Fenster in app/).
import * as api from "./graph.js";
import { config } from "./config.js";
import { einstellungen } from "../src/einstellungen.js";
import { findeAnbieter } from "../src/planer.js";
import { vorschauTag, planeTagNeu, setzeKuerzel } from "../src/aktionen.js";
import { wendeAn } from "../src/eigeneEinstellungen.js";
import {
  $, zustand, nutzeHost, mitWarten, zeigeAenderungen, zeigeProtokoll, richteAuswertungEin, richteNeuerKursEin,
  richteMailEin, richteUrlaubEin, richteFahrzeitEin, uebernehme,
} from "./bereiche.js";

// Was nur in Outlook geht: Einstellungen im Postfach, der offene Termin.
const officeOk = (r) => r.status === Office.AsyncResultStatus.Succeeded;
nutzeHost({
  lies: (k) => Office.context.roamingSettings.get(k),
  merke: (k, v) => {
    Office.context.roamingSettings.set(k, v);
    Office.context.roamingSettings.saveAsync();
  },
  oeffneEntwurf: (id) => {
    try {
      Office.context.mailbox.displayMessageForm(Office.context.mailbox.convertToEwsId(id, Office.MailboxEnums.RestVersion.v2_0));
      return true;
    } catch {
      return false; // dann eben aus den Entwürfen öffnen
    }
  },
  // Text des offenen Termins: im Bearbeiten-Fenster der aktuelle Stand
  // (auch noch nicht gespeichert), sonst der aus Graph.
  terminText: (t) => {
    const item = Office.context.mailbox.item;
    if (!item.body?.setAsync) return Promise.resolve(t.text);
    return new Promise((ok) => item.body.getAsync(Office.CoercionType.Text, (r) => ok(officeOk(r) ? r.value : t.text)));
  },
  // Ist die ganze Serie offen statt nur ein Termin? Dann darf kein Protokoll
  // hinein, sonst stünde es in jedem Termin der Serie.
  istSerie: () => {
    const item = Office.context.mailbox.item;
    // Ein einzelner Termin aus einer Serie hat eine seriesId, die Serie selbst nicht.
    if (!item.recurrence?.getAsync || item.seriesId) return Promise.resolve(false);
    return new Promise((ok) => item.recurrence.getAsync((r) => ok(officeOk(r) && !!r.value)));
  },
  schreibeInsFenster: (text) => {
    const item = Office.context.mailbox.item;
    if (!item.body?.setAsync) return Promise.resolve(false);
    return new Promise((ok, fehler) =>
      item.body.setAsync(text, { coercionType: Office.CoercionType.Text }, (r) =>
        officeOk(r) ? ok(true) : fehler(new Error(r.error.message))
      )
    );
  },
});

// Eigene Termine öffnen sich zum Bearbeiten; dort gibt es itemId nicht direkt.
function eigeneId(item) {
  if (item.itemId) return Promise.resolve(item.itemId);
  return new Promise((ok, fehler) =>
    item.getItemIdAsync((r) =>
      r.status === Office.AsyncResultStatus.Succeeded
        ? ok(r.value)
        : fehler(new Error("Bitte den Termin erst speichern, dann den Kursplaner öffnen."))
    )
  );
}

// Auf-/Zugeklappt merken, damit jeder Bereich so bleibt, wie du ihn zuletzt hattest.
function richteKlappenEin() {
  document.querySelectorAll("details.karte").forEach((d) => {
    const schluessel = `klappe-${d.id}`;
    let gemerkt = null;
    try {
      gemerkt = localStorage.getItem(schluessel);
    } catch {
      /* ohne Speicher eben Standard */
    }
    d.open = gemerkt === null ? d.dataset.auf === "1" : gemerkt === "1";
    // Nur echte Klicks merken, nicht das Aufklappen durch das Add-in selbst.
    d.querySelector("summary").addEventListener("click", () => {
      try {
        localStorage.setItem(schluessel, d.open ? "0" : "1");
      } catch {
        /* egal */
      }
    });
  });
}

const NUR_BEREICH = {
  auswertung: { titel: "Monatsauswertung", id: "bereich-auswertung", einrichten: () => richteAuswertungEin() },
  neu: { titel: "Neuer Kurs", id: "bereich-neu", einrichten: () => richteNeuerKursEin() },
  urlaub: { titel: "Urlaub", id: "bereich-urlaub", einrichten: () => richteUrlaubEin() },
};

async function start() {
  if (config.clientId.startsWith("HIER")) {
    $("info").textContent = "Noch nicht eingerichtet: Die App-Registrierung fehlt in web/config.js.";
    return;
  }
  richteKlappenEin();
  // Eigene Einstellungen aus dem Postfach (im Kursplaner-Fenster gepflegt).
  try {
    wendeAn(einstellungen, await api.leseEinstellungen());
  } catch {
    /* dann eben die Standardwerte */
  }
  // Eigener Knopf im Reiter „Kursplaner“: nur dieser eine Bereich, egal was
  // gerade ausgewählt ist (auch aus einer neuen E-Mail heraus).
  const nur = NUR_BEREICH[new URLSearchParams(location.search).get("bereich")];
  if (nur) {
    $("titel").textContent = nur.titel;
    $("info").textContent = "";
    nur.einrichten();
    const d = $(nur.id);
    d.hidden = false;
    d.open = true;
    return;
  }
  const item = Office.context.mailbox.item;
  // Ohne Termin (z. B. aus einer E-Mail geöffnet): nur die Monatsauswertung.
  if (item?.itemType !== Office.MailboxEnums.ItemType.Appointment) {
    $("info").textContent = "Für die Monatsauswertung und neue Kurse brauchst du keinen Termin.";
    $("bereich-auswertung").hidden = false;
    $("bereich-auswertung").open = true;
    richteAuswertungEin();
    richteNeuerKursEin();
    richteUrlaubEin();
    return;
  }
  const ewsId = await eigeneId(item);
  const id = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
  const termin = (zustand.termin = await api.holeTermin(id));
  const datum = termin.start.slice(0, 10);
  const anbieter = findeAnbieter(termin, einstellungen);
  $("titel").textContent = termin.titel;
  $("info").textContent = `${datum.split("-").reverse().join(".")} · ${termin.start.slice(11)}–${termin.ende.slice(11)} · ${termin.kalender}`;

  const istAnbieterKalender = einstellungen.anbieter.some((a) => a.kalender === termin.kalender);
  $("bereich-anbieter").hidden = !(istAnbieterKalender && anbieter);
  $("bereich-eigen").hidden = !(termin.eigenerKalender && anbieter && !termin.generiert);

  if (termin.eigenerKalender && !anbieter && !termin.generiert) richteFahrzeitEin();
  $("bereich-planen").hidden = false;
  $("bereich-auswertung").hidden = false;
  richteAuswertungEin();
  richteNeuerKursEin();
  richteUrlaubEin();
  if (anbieter && !termin.generiert) richteMailEin(anbieter);
  if (!$("bereich-eigen").hidden) await zeigeProtokoll();

  $("ich-gebe").onclick = () => uebernehme(false);
  $("ich-vertrete").onclick = () => uebernehme(true);
  document.querySelectorAll("[data-kuerzel]").forEach((b) => {
    b.onclick = () =>
      mitWarten(b, async () => {
        const v = await setzeKuerzel(api, zustand.termin, b.dataset.kuerzel, einstellungen);
        // Titel lokal nachziehen statt den Termin neu zu laden (spart Abfragen).
        zustand.termin = { ...zustand.termin, titel: v.titel ?? zustand.termin.titel };
        $("titel").textContent = zustand.termin.titel;
        // Ist der Termin zum Bearbeiten offen, auch dort den Titel anpassen,
        // sonst überschreibt "Speichern" die Änderung wieder.
        const item = Office.context.mailbox.item;
        if (item.subject?.setAsync) item.subject.setAsync(zustand.termin.titel);
        zeigeAenderungen(v, "Erledigt");
      });
  });
  $("vorschau").onclick = () =>
    mitWarten($("vorschau"), async () => {
      const v = await vorschauTag(api, datum, einstellungen);
      $("ausfuehren").hidden = zeigeAenderungen(v, "Das würde sich ändern") === 0;
    });
  $("ausfuehren").onclick = () =>
    mitWarten($("ausfuehren"), async () => {
      const v = await planeTagNeu(api, datum, einstellungen);
      $("ausfuehren").hidden = true;
      zeigeAenderungen(v, "Erledigt");
    });
}

// Angeheftet: Wählt Michael einen anderen Termin oder eine andere E-Mail,
// lädt der Kursplaner einfach neu.
Office.onReady(() => {
  try {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => location.reload());
  } catch {
    /* ältere Outlook-Versionen: dann eben ohne */
  }
});
Office.onReady(() => start().catch((e) => ($("info").innerHTML = `<span class="fehler">${e.message}</span>`)));

