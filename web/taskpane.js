import * as api from "./graph.js";
import { config } from "./config.js";
import { einstellungen } from "../src/einstellungen.js";
import { findeAnbieter, ohneKuerzel } from "../src/planer.js";
import { FELDER, leseProtokoll, schreibeProtokoll, gesamt, auswertungVerlauf, teileMaterial } from "../src/protokoll.js";
import { erstelleAuswertung, monatsZeitraum } from "../src/auswertung.js";
import { erstellePdf } from "./bericht.js";
import { vorschauTag, planeTagNeu, uebernehmeKurs, setzeKuerzel, setzeFahrzeit } from "../src/aktionen.js";

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

async function start() {
  if (config.clientId.startsWith("HIER")) {
    $("info").textContent = "Noch nicht eingerichtet: Die App-Registrierung fehlt in web/config.js.";
    return;
  }
  richteKlappenEin();
  const item = Office.context.mailbox.item;
  // Ohne Termin (z. B. aus einer E-Mail geöffnet): nur die Monatsauswertung.
  if (item?.itemType !== Office.MailboxEnums.ItemType.Appointment) {
    $("info").textContent = "Für die Monatsauswertung brauchst du keinen Termin.";
    $("bereich-auswertung").hidden = false;
    $("bereich-auswertung").open = true;
    richteAuswertungEin();
    return;
  }
  const ewsId = await eigeneId(item);
  const id = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
  termin = await api.holeTermin(id);
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
  const p = Object.fromEntries(FELDER.map(([k]) => [k, $(`p-${k}`)?.value ?? ""]));
  p.material = [...document.querySelectorAll("#p-material-liste input:checked")].map((c) => c.value).join(", ");
  p.vorkommnis = $("p-vorkommnis-ja").checked ? "ja" : "nein";
  return p;
}

const html = (t) => String(t).replace(/[&<>"]/g, (z) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[z]);

function zeigeMaterial(alle, gewaehlt) {
  const liste = [...new Set([...alle, ...gewaehlt])].sort((a, b) => a.localeCompare(b, "de"));
  $("p-material-liste").innerHTML = liste
    .map((m) => `<label><input type="checkbox" value="${html(m)}"${gewaehlt.includes(m) ? " checked" : ""}> ${html(m)}</label>`)
    .join("");
}

function setzeMaterial(gewaehlt) {
  const alle = [...document.querySelectorAll("#p-material-liste input")].map((c) => c.value);
  zeigeMaterial(alle, gewaehlt);
}

// Ist die ganze Serie offen statt nur ein Termin? Dann darf kein Protokoll
// hinein, sonst stünde es in jedem Termin der Serie.
function istSerie() {
  const item = Office.context.mailbox.item;
  // Ein einzelner Termin aus einer Serie hat eine seriesId, die Serie selbst nicht.
  if (!item.recurrence?.getAsync || item.seriesId) return Promise.resolve(false);
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
  $("p-vorkommnis-ja").checked = p.vorkommnis.toLowerCase() === "ja";
  zeigeMaterial(einstellungen.material, teileMaterial(p.material));

  // Frühere Kurse der letzten 6 Monate: letzte Themen, Themenliste, Material.
  try {
    const datum = termin.start.slice(0, 10);
    const von = new Date(Date.parse(datum) - 183 * 86400000).toISOString().slice(0, 10);
    const kurse = (await api.holeVerlauf(von, datum)).filter((t) => findeAnbieter(t, einstellungen) && !t.generiert);
    const v = auswertungVerlauf(kurse, ohneKuerzel(termin.titel), datum, einstellungen.material, ohneKuerzel);
    $("p-themen").innerHTML = v.themen.map((t) => `<option value="${html(t)}">`).join("");
    zeigeMaterial(v.material, teileMaterial(p.material));
    if (v.letzte.length) {
      const tag = (d) => d.slice(8, 10) + "." + d.slice(5, 7) + ".";
      $("p-letzte").innerHTML =
        "<strong>Letzte Themen</strong>" + v.letzte.map((x) => `<div>${tag(x.datum)} ${html(x.p.thema)}</div>`).join("");
      $("p-letzte").hidden = false;
      $("p-wie-letztes").hidden = false;
      $("p-wie-letztes").onclick = () => {
        const l = v.letzte[0].p;
        $("p-thema").value = l.thema;
        setzeMaterial(teileMaterial(l.material));
        if (!$("p-erstvertretung").value) $("p-erstvertretung").value = l.erstvertretung;
      };
      if (!p.erstvertretung) $("p-erstvertretung").value = v.letzte[0].p.erstvertretung;
    }
  } catch (e) {
    $("p-letzte").textContent = `Frühere Themen konnten nicht geladen werden: ${e.message}`;
    $("p-letzte").hidden = false;
  }

  $("p-material-dazu").onclick = () => {
    const neu = $("p-material-neu").value.trim();
    if (!neu) return;
    const alle = [...document.querySelectorAll("#p-material-liste input")].map((c) => c.value);
    zeigeMaterial([...alle, neu], [...teileMaterial(protokollAusFormular().material), neu]);
    $("p-material-neu").value = "";
  };
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

// --- Monatsauswertung ---
const jetztLokal = () => new Date().toLocaleString("sv-SE", { timeZone: "Europe/Berlin" }).replace(" ", "T").slice(0, 16);
const merke = (k, v) => {
  Office.context.roamingSettings.set(k, v);
  Office.context.roamingSettings.saveAsync();
};

function richteAuswertungEin() {
  const heute = jetztLokal().slice(0, 10);
  // Bis zum 10. ist meist der Vormonat gemeint.
  const d = new Date(`${heute.slice(0, 7)}-01T12:00:00Z`);
  if (Number(heute.slice(8, 10)) <= 10) d.setUTCMonth(d.getUTCMonth() - 1);
  $("a-monat").value = d.toISOString().slice(0, 7);
  $("a-anbieter").innerHTML = einstellungen.anbieter.map((a) => `<option value="${a.kuerzel}">${html(a.name)}</option>`).join("");
  const anbieter = termin && findeAnbieter(termin, einstellungen);
  if (anbieter) $("a-anbieter").value = anbieter.kuerzel;
  const zeigeEmpfaenger = () => ($("a-empfaenger").value = Office.context.roamingSettings.get(`empfaenger-${$("a-anbieter").value}`) || "");
  $("a-anbieter").onchange = () => {
    zeigeEmpfaenger();
    $("a-aktionen").hidden = true;
    $("a-ergebnis").innerHTML = "";
  };
  zeigeEmpfaenger();

  let bericht;
  $("a-anzeigen").onclick = () =>
    mitWarten($("a-anzeigen"), async () => {
      const anbieter = einstellungen.anbieter.find((a) => a.kuerzel === $("a-anbieter").value);
      const zeitraum = monatsZeitraum($("a-monat").value);
      const stand = jetztLokal();
      const a = erstelleAuswertung(await api.holeVerlauf(zeitraum.von, zeitraum.bis), anbieter, stand, einstellungen);
      bericht = { a, anbieter, zeitraum, stand };
      const s = a.summe;
      const tag = (d) => `${d.slice(8, 10)}.${d.slice(5, 7)}.`;
      $("a-ergebnis").innerHTML =
        `<div style="margin-top:8px"><strong>${html(zeitraum.name)} · ${html(anbieter.kuerzel)}</strong></div>` +
        `<div>Geleistet: ${s.geleistet} (davon ${s.davonVertretung} Vertretung) · Geplant: ${s.geplant} · Ferien: ${s.ferien}</div>` +
        `<div>Teilnehmer: ${s.teilnehmerSchnitt === null ? "keine Angaben" : `${s.teilnehmer}, Schnitt ${s.teilnehmerSchnitt}`}</div>` +
        (s.ohneProtokoll ? `<div class="fehler">${s.ohneProtokoll} geleistete Kurse ohne Teilnehmerzahl</div>` : "") +
        `<table><tr><th>Tag</th><th>Kurs</th><th>Status</th><th>TN</th></tr>` +
        a.zeilen
          .map((z) => `<tr><td>${tag(z.datum)}</td><td>${html(z.kurs)}${z.art === "Vertretung" ? " (V)" : ""}</td><td>${z.status}</td><td>${z.anwesend ?? "–"}</td></tr>`)
          .join("") +
        `</table>`;
      $("a-aktionen").hidden = false;
    });

  const dateiname = () => `Monatsauswertung ${bericht.anbieter.kuerzel} ${bericht.zeitraum.name}.pdf`;
  $("a-pdf").onclick = () => erstellePdf(bericht.a, bericht.anbieter, bericht.zeitraum, bericht.stand).save(dateiname());
  $("a-mail").onclick = () =>
    mitWarten($("a-mail"), async () => {
      const an = $("a-empfaenger").value.trim();
      if (an) merke(`empfaenger-${bericht.anbieter.kuerzel}`, an);
      const pdf = erstellePdf(bericht.a, bericht.anbieter, bericht.zeitraum, bericht.stand);
      const id = await api.legeEntwurfAn({
        an,
        betreff: `Monatsauswertung ${bericht.zeitraum.name} – Michael Stehle`,
        text: `Hallo,\n\nanbei meine Monatsauswertung für ${bericht.zeitraum.name}.\n\nViele Grüße\nMichael Stehle`,
        dateiname: dateiname(),
        pdfBase64: pdf.output("datauristring").split(",")[1],
      });
      zeige(`<div class="karte"><strong>E-Mail-Entwurf erstellt</strong><div class="leise">Er liegt in deinen Entwürfen und öffnet sich gleich.</div></div>`);
      try {
        Office.context.mailbox.displayMessageForm(Office.context.mailbox.convertToEwsId(id, Office.MailboxEnums.RestVersion.v2_0));
      } catch {
        /* dann eben aus den Entwürfen öffnen */
      }
    });
}

// Privater Termin: Fahrzeit wählen, Hin- und Rückfahrt werden angelegt und
// wandern mit, wenn der Termin verschoben wird.
function richteFahrzeitEin() {
  $("bereich-fahrzeit").hidden = false;
  const schluessel = `fahrzeit-ort-${termin.ort.trim().toLowerCase()}`;
  const vorschlag = termin.ort.trim() ? Number(Office.context.roamingSettings.get(schluessel)) || 0 : 0;
  const markiere = () =>
    document.querySelectorAll("[data-fahrzeit]").forEach((b) => {
      const m = Number(b.dataset.fahrzeit);
      b.classList.toggle("zweit", m !== termin.fahrzeit);
    });
  $("f-info").textContent = termin.fahrzeit
    ? `Eingestellt: ${termin.fahrzeit} Minuten pro Strecke.`
    : vorschlag
      ? `Letztes Mal an diesem Ort: ${vorschlag} Minuten.`
      : "Wie lange fährst du pro Strecke?";
  markiere();
  document.querySelectorAll("[data-fahrzeit]").forEach((b) => {
    b.onclick = () =>
      mitWarten(b, async () => {
        const minuten = Number(b.dataset.fahrzeit);
        const v = await setzeFahrzeit(api, termin, minuten, einstellungen);
        termin = { ...termin, fahrzeit: minuten };
        if (minuten && termin.ort.trim()) merke(schluessel, minuten);
        $("f-info").textContent = minuten ? `Eingestellt: ${minuten} Minuten pro Strecke.` : "Keine Fahrzeit.";
        markiere();
        zeigeAenderungen(v, "Erledigt");
      });
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
