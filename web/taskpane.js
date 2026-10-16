import * as api from "./graph.js";
import { config } from "./config.js";
import { einstellungen } from "../src/einstellungen.js";
import { findeAnbieter } from "../src/planer.js";
import { FELDER, leseProtokoll, schreibeProtokoll, gesamt } from "../src/protokoll.js";
import { vorschauTag, planeTagNeu, uebernehmeKurs, setzeKuerzel } from "../src/aktionen.js";

const $ = (id) => document.getElementById(id);
let termin;

const zeile = (t, klasse) => `<li class="${klasse}">${t.start.slice(11)}–${t.ende.slice(11)} ${t.titel}</li>`;

function zeige(html) {
  $("ergebnis").innerHTML = html;
}

function zeigeAenderungen(v, ueberschrift) {
  const zeilen = [...v.loeschen.map((t) => zeile(t, "weg")), ...v.anlegen.map((t) => zeile(t, "neu"))];
  zeige(
    `<div class="karte"><strong>${ueberschrift}</strong>` +
      (zeilen.length ? `<ul>${zeilen.join("")}</ul>` : `<div class="leise">Alles passt, nichts zu ändern.</div>`) +
      `</div>`
  );
  return zeilen.length;
}

async function mitWarten(knopf, arbeit) {
  const alle = document.querySelectorAll("button");
  alle.forEach((b) => (b.disabled = true));
  try {
    await arbeit();
  } catch (e) {
    zeige(`<div class="karte fehler">Das hat nicht geklappt: ${e.message}</div>`);
  } finally {
    alle.forEach((b) => (b.disabled = false));
  }
}

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

async function start() {
  if (config.clientId.startsWith("HIER")) {
    $("info").textContent = "Noch nicht eingerichtet: Die App-Registrierung fehlt in web/config.js.";
    return;
  }
  const ewsId = await eigeneId(Office.context.mailbox.item);
  const id = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
  termin = await api.holeTermin(id);
  const datum = termin.start.slice(0, 10);
  const anbieter = findeAnbieter(termin, einstellungen);
  $("titel").textContent = termin.titel;
  $("info").textContent = `${datum.split("-").reverse().join(".")} · ${termin.start.slice(11)}–${termin.ende.slice(11)} · ${termin.kalender}`;

  const istAnbieterKalender = einstellungen.anbieter.some((a) => a.kalender === termin.kalender);
  $("bereich-anbieter").hidden = !(istAnbieterKalender && anbieter);
  $("bereich-eigen").hidden = !(termin.eigenerKalender && anbieter && !termin.generiert);

  if (!$("bereich-eigen").hidden) await zeigeProtokoll();

  $("ich-gebe").onclick = () => uebernehme(false);
  $("ich-vertrete").onclick = () => uebernehme(true);
  document.querySelectorAll("[data-kuerzel]").forEach((b) => {
    b.onclick = () =>
      mitWarten(b, async () => {
        const v = await setzeKuerzel(api, termin, b.dataset.kuerzel, einstellungen);
        // Titel lokal nachziehen statt den Termin neu zu laden (spart Abfragen).
        termin = { ...termin, titel: v.titel ?? termin.titel };
        $("titel").textContent = termin.titel;
        // Ist der Termin zum Bearbeiten offen, auch dort den Titel anpassen,
        // sonst überschreibt "Speichern" die Änderung wieder.
        const item = Office.context.mailbox.item;
        if (item.subject?.setAsync) item.subject.setAsync(termin.titel);
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

// Text des offenen Termins: im Bearbeiten-Fenster der aktuelle Stand
// (auch noch nicht gespeichert), sonst der aus Graph.
function terminText() {
  const item = Office.context.mailbox.item;
  if (!item.body?.setAsync) return Promise.resolve(termin.text);
  return new Promise((ok) =>
    item.body.getAsync(Office.CoercionType.Text, (r) => ok(r.status === Office.AsyncResultStatus.Succeeded ? r.value : termin.text))
  );
}

function protokollAusFormular() {
  return Object.fromEntries(FELDER.map(([k]) => [k, $(`p-${k}`)?.value ?? ""]));
}

// Ist die ganze Serie offen statt nur ein Termin? Dann darf kein Protokoll
// hinein, sonst stünde es in jedem Termin der Serie.
function istSerie() {
  const item = Office.context.mailbox.item;
  if (!item.recurrence?.getAsync) return Promise.resolve(false);
  return new Promise((ok) =>
    item.recurrence.getAsync((r) => ok(r.status === Office.AsyncResultStatus.Succeeded && !!r.value))
  );
}

async function zeigeProtokoll() {
  $("bereich-protokoll").hidden = false;
  if (await istSerie()) {
    $("p-formular").hidden = true;
    $("p-serie").hidden = false;
    return;
  }
  const p = leseProtokoll(await terminText());
  for (const [k] of FELDER) if ($(`p-${k}`)) $(`p-${k}`).value = p[k];
  const zeigeGesamt = () => ($("p-gesamt").textContent = `Teilnehmer gesamt: ${gesamt(protokollAusFormular()) || "–"}`);
  ["p-anwesend", "p-entschuldigt", "p-unentschuldigt"].forEach((id) => ($(id).oninput = zeigeGesamt));
  zeigeGesamt();
  $("p-speichern").onclick = () =>
    mitWarten($("p-speichern"), async () => {
      const text = schreibeProtokoll(protokollAusFormular());
      const item = Office.context.mailbox.item;
      if (item.body?.setAsync) {
        // Termin ist zum Bearbeiten offen: in das Fenster schreiben, Michael speichert.
        await new Promise((ok, fehler) =>
          item.body.setAsync(text, { coercionType: Office.CoercionType.Text }, (r) =>
            r.status === Office.AsyncResultStatus.Succeeded ? ok() : fehler(new Error(r.error.message))
          )
        );
        zeige(`<div class="karte"><strong>Protokoll eingetragen</strong><div class="leise">Jetzt noch oben auf „Speichern &amp; schließen“ klicken.</div></div>`);
      } else {
        await api.aendere(termin.id, { body: { contentType: "text", content: text } });
        zeige(`<div class="karte"><strong>Protokoll gespeichert</strong></div>`);
      }
    });
}

function uebernehme(vertretung) {
  mitWarten(null, async () => {
    const { doppelt, aenderungen } = await uebernehmeKurs(api, termin, { vertretung }, einstellungen);
    if (doppelt) {
      zeige(`<div class="karte">Der Kurs steht schon in deinem Kalender: <strong>${doppelt.titel}</strong></div>`);
      return;
    }
    zeigeAenderungen(aenderungen, vertretung ? "Vertretung eingetragen" : "Kurs eingetragen");
  });
}

Office.onReady(() => start().catch((e) => ($("info").innerHTML = `<span class="fehler">${e.message}</span>`)));
