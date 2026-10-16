import * as api from "./graph.js";
import { config } from "./config.js";
import { einstellungen } from "../src/einstellungen.js";
import { findeAnbieter } from "../src/planer.js";
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

async function start() {
  if (config.clientId.startsWith("HIER")) {
    $("info").textContent = "Noch nicht eingerichtet: Die App-Registrierung fehlt in web/config.js.";
    return;
  }
  const item = Office.context.mailbox.item;
  const id = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0);
  termin = await api.holeTermin(id);
  const datum = termin.start.slice(0, 10);
  const anbieter = findeAnbieter(termin, einstellungen);
  $("titel").textContent = termin.titel;
  $("info").textContent = `${datum.split("-").reverse().join(".")} · ${termin.start.slice(11)}–${termin.ende.slice(11)} · ${termin.kalender}`;

  const istAnbieterKalender = einstellungen.anbieter.some((a) => a.kalender === termin.kalender);
  $("bereich-anbieter").hidden = !(istAnbieterKalender && anbieter);
  $("bereich-eigen").hidden = !(termin.eigenerKalender && anbieter && !termin.generiert);

  $("ich-gebe").onclick = () => uebernehme(false);
  $("ich-vertrete").onclick = () => uebernehme(true);
  document.querySelectorAll("[data-kuerzel]").forEach((b) => {
    b.onclick = () =>
      mitWarten(b, async () => {
        const v = await setzeKuerzel(api, termin, b.dataset.kuerzel, einstellungen);
        termin = await api.holeTermin(termin.id);
        $("titel").textContent = termin.titel;
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
