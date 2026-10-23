// Die Bereiche des Kursplaners (Protokoll, Monatsauswertung, Neuer Kurs,
// Urlaub, E-Mail, Fahrzeit). Gemeinsam für das Outlook-Add-in (taskpane.js)
// und das eigene Kursplaner-Fenster (app/app.js). Was nur in Outlook geht
// (Einstellungen merken, geöffneten Termin lesen/schreiben, Entwurf öffnen),
// kommt über "host" (siehe nutzeHost).
import * as api from "./graph.js";
import { config } from "./config.js";
import { einstellungen } from "../src/einstellungen.js";
import { findeAnbieter, ohneKuerzel } from "../src/planer.js";
import { FELDER, leseProtokoll, schreibeProtokoll, gesamt, auswertungVerlauf, teileMaterial } from "../src/protokoll.js";
import { erstelleAuswertung, monatsZeitraum } from "../src/auswertung.js";
import { erstellePdf } from "./bericht.js";
import { fuelleVorlage } from "../src/vorlagen.js";
import { pruefeKurs, kursEreignis, kursTitel, anzahlTermine } from "../src/neuerKurs.js";
import { uebernehmeKurs, setzeFahrzeit, setzeZeitraum } from "../src/aktionen.js";

// Der Termin, um den es gerade geht (null = keiner).
export const zustand = { termin: null };

// host: { lies(k), merke(k, v), oeffneEntwurf(id) → true wenn geöffnet,
//   terminText(t), istSerie(), schreibeInsFenster(text) → true wenn geschrieben }
let host;
export const nutzeHost = (h) => (host = h);

export const $ = (id) => document.getElementById(id);

const zeile = (t, klasse) => `<li class="${klasse}">${t.start.slice(11)}–${t.ende.slice(11)} ${t.titel}</li>`;

export function zeige(html) {
  $("ergebnis").innerHTML = html;
}

export function zeigeAenderungen(v, ueberschrift) {
  const zeilen = [...v.loeschen.map((t) => zeile(t, "weg")), ...v.anlegen.map((t) => zeile(t, "neu"))];
  zeige(
    `<div class="karte"><strong>${ueberschrift}</strong>` +
      (zeilen.length ? `<ul>${zeilen.join("")}</ul>` : `<div class="leise">Alles passt, nichts zu ändern.</div>`) +
      `</div>`
  );
  return zeilen.length;
}

export async function mitWarten(knopf, arbeit) {
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

// Text des offenen Termins: im Bearbeiten-Fenster der aktuelle Stand
// (auch noch nicht gespeichert), sonst der aus Graph.
function terminText() {
  return host.terminText(zustand.termin);
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
  return host.istSerie();
}

export async function zeigeProtokoll() {
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
    const datum = zustand.termin.start.slice(0, 10);
    const von = new Date(Date.parse(datum) - 183 * 86400000).toISOString().slice(0, 10);
    const kurse = (await api.holeVerlauf(von, datum)).filter((t) => findeAnbieter(t, einstellungen) && !t.generiert);
    const v = auswertungVerlauf(kurse, ohneKuerzel(zustand.termin.titel), datum, einstellungen.material, ohneKuerzel);
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
      // Termin ist zum Bearbeiten offen: in das Fenster schreiben, Michael speichert.
      if (await host.schreibeInsFenster(text)) {
        zeige(`<div class="karte"><strong>Protokoll eingetragen</strong><div class="leise">Jetzt noch oben auf „Speichern &amp; schließen“ klicken.</div></div>`);
      } else {
        await api.aendere(zustand.termin.id, { body: { contentType: "text", content: text } });
        zeige(`<div class="karte"><strong>Protokoll gespeichert</strong></div>`);
        host.nachProtokoll?.(); // App: Seitenleiste zu, Liste neu
      }
    });
}

// --- Monatsauswertung ---
const jetztLokal = () => new Date().toLocaleString("sv-SE", { timeZone: "Europe/Berlin" }).replace(" ", "T").slice(0, 16);
const merke = (k, v) => host.merke(k, v);

export function richteAuswertungEin() {
  const heute = jetztLokal().slice(0, 10);
  // Bis zum 10. ist meist der Vormonat gemeint.
  const d = new Date(`${heute.slice(0, 7)}-01T12:00:00Z`);
  if (Number(heute.slice(8, 10)) <= 10) d.setUTCMonth(d.getUTCMonth() - 1);
  $("a-monat").value = d.toISOString().slice(0, 7);
  $("a-anbieter").innerHTML = einstellungen.anbieter.map((a) => `<option value="${a.kuerzel}">${html(a.name)}</option>`).join("");
  const anbieter = zustand.termin && findeAnbieter(zustand.termin, einstellungen);
  if (anbieter) $("a-anbieter").value = anbieter.kuerzel;
  const zeigeEmpfaenger = () => ($("a-empfaenger").value = host.lies(`empfaenger-${$("a-anbieter").value}`) || "");
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
        `<div>Fahrten: ${s.fahrten}${s.fahrtKm != null ? `, ${String(s.fahrtKm).replace(".", ",")} km` : ""}</div>` +
        `<table><tr><th>Tag</th><th>Kurs</th><th>Status</th></tr>` +
        a.zeilen
          .map((z) => `<tr><td>${tag(z.datum)}</td><td>${html(z.kurs)}${z.art === "Vertretung" ? " (V)" : ""}</td><td>${z.status}</td></tr>`)
          .join("") +
        `</table>`;
      $("a-aktionen").hidden = false;
    });

  const dateiname = () => `Monatsauswertung ${bericht.anbieter.kuerzel} ${bericht.zeitraum.name}.pdf`;
  $("a-pdf").onclick = () => erstellePdf(bericht.a, bericht.anbieter, bericht.zeitraum, bericht.stand, einstellungen.monatsbericht).save(dateiname());
  $("a-mail").onclick = () =>
    mitWarten($("a-mail"), async () => {
      const an = $("a-empfaenger").value.trim();
      if (an) merke(`empfaenger-${bericht.anbieter.kuerzel}`, an);
      const pdf = erstellePdf(bericht.a, bericht.anbieter, bericht.zeitraum, bericht.stand, einstellungen.monatsbericht);
      const id = await api.legeEntwurfAn({
        an,
        betreff: `Monatsauswertung ${bericht.zeitraum.name} – Michael Stehle`,
        text: `Hallo,\n\nanbei meine Monatsauswertung für ${bericht.zeitraum.name}.\n\nViele Grüße\nMichael Stehle`,
        dateiname: dateiname(),
        pdfBase64: pdf.output("datauristring").split(",")[1],
      });
      const offen = host.oeffneEntwurf(id);
      zeige(`<div class="karte"><strong>E-Mail-Entwurf erstellt</strong><div class="leise">Er liegt in deinen Entwürfen${offen ? " und öffnet sich gleich" : ""}.</div></div>`);
    });
}

// --- Neuer Kurs: im Anbieter-Kalender anlegen, einmalig oder jede Woche ---
export function richteNeuerKursEin() {
  $("bereich-neu").hidden = false;
  const heute = jetztLokal().slice(0, 10);
  // Neben den festen Anbietern: Vertretung in einem anderen Verein (wechselnd).
  const extern = einstellungen.vertretungExtern;
  const ziele = [...einstellungen.anbieter, ...(extern ? [{ ...extern, name: "Anderer Verein (Vertretung)" }] : [])];
  const anbieterVon = () => ziele.find((a) => a.kuerzel === $("n-anbieter").value);
  const istExternGewaehlt = () => !!extern && $("n-anbieter").value === extern.kuerzel;
  $("n-anbieter").innerHTML = ziele.map((a) => `<option value="${a.kuerzel}">${html(a.name)}</option>`).join("");
  const offen = zustand.termin && findeAnbieter(zustand.termin, einstellungen);
  if (offen) $("n-anbieter").value = offen.kuerzel;
  // Vereine mit ihrer letzten Adresse merken: { "TSV Beispiel": "Halle, Straße, Ort" }.
  const vereine = () => {
    try {
      return JSON.parse(host.lies("vereine") || "{}");
    } catch {
      return {};
    }
  };
  $("n-datum").value = heute;

  // Vorschläge aus den nächsten Wochen im Anbieter-Kalender: Kursnamen,
  // Kursleitungen und die übliche Dauer eines Kurses.
  let dauer = {};
  const ladeVorschlaege = async () => {
    const a = anbieterVon();
    $("n-kategorie").innerHTML = a.kategorien.map((k) => `<option>${html(k)}</option>`).join("");
    $("n-ort").value = a.ort;
    // Anderer Verein: nur in meinen Kalender, statt Kursleitung der Verein, Ort von Hand.
    const ext = istExternGewaehlt();
    $("n-wohin").value = ext ? "eigen" : $("n-wohin").value;
    $("n-wohin").disabled = ext;
    $("n-leitung-label").textContent = ext ? "Verein" : "Kursleitung";
    $("n-leitung").value = ext ? "" : einstellungen.dauerkurse.kursleitung;
    $("n-ort").placeholder = ext ? "Adresse der Halle (für Fahrzeit und Navigation)" : "";
    if (ext) {
      const v = vereine();
      $("n-leitungen").innerHTML = Object.keys(v).sort((x, y) => x.localeCompare(y, "de")).map((n) => `<option value="${html(n)}">`).join("");
      $("n-namen").innerHTML = "";
      dauer = {};
      return;
    }
    try {
      const bis = new Date(Date.parse(`${heute}T12:00:00Z`) + 28 * 86400000).toISOString().slice(0, 10);
      const kurse = await api.holeKursTitel(a.kalender, heute, bis);
      const teile = kurse.map((k) => k.titel.match(/^(?:\([A-Z]\)\s*)?(.*?)\s*\(([^)]+)\)\s*$/)).filter(Boolean);
      const sortiert = (liste) => [...new Set(liste)].sort((x, y) => x.localeCompare(y, "de"));
      $("n-namen").innerHTML = sortiert(teile.map((m) => m[1])).map((n) => `<option value="${html(n)}">`).join("");
      $("n-leitungen").innerHTML = sortiert(teile.map((m) => m[2])).map((n) => `<option value="${html(n)}">`).join("");
      dauer = {};
      kurse.forEach((k) => {
        const m = k.titel.match(/^(?:\([A-Z]\)\s*)?(.*?)\s*\(/);
        if (m) dauer[m[1]] ??= (Date.parse(k.ende) - Date.parse(k.start)) / 60000;
      });
    } catch {
      /* ohne Vorschläge geht es auch */
    }
  };
  ladeVorschlaege();
  $("n-anbieter").onchange = () => {
    ladeVorschlaege();
    $("n-anlegen").hidden = true;
  };
  // Beginn gesetzt und Ende leer: übliche Dauer dieses Kurses (sonst 45 Min.) vorschlagen.
  const endeVorschlagen = () => {
    if (!$("n-beginn").value || $("n-ende").dataset.vonHand) return;
    const [h, m] = $("n-beginn").value.split(":").map(Number);
    const ende = h * 60 + m + (dauer[$("n-name").value.trim()] || 45);
    $("n-ende").value = `${String(Math.floor(ende / 60) % 24).padStart(2, "0")}:${String(ende % 60).padStart(2, "0")}`;
  };
  $("n-beginn").oninput = endeVorschlagen;
  $("n-name").onchange = endeVorschlagen;
  // Bekannter Verein: letzte Adresse vorschlagen.
  $("n-leitung").onchange = () => {
    const adr = istExternGewaehlt() && vereine()[$("n-leitung").value.trim()];
    if (adr && !$("n-ort").value.trim()) $("n-ort").value = adr;
  };
  $("n-ende").oninput = () => ($("n-ende").dataset.vonHand = "1");
  $("n-serie").onchange = () => ($("n-bis-feld").hidden = !$("n-serie").checked);
  document.querySelectorAll("#bereich-neu input, #bereich-neu select").forEach((f) =>
    f.addEventListener("input", () => ($("n-anlegen").hidden = true))
  );

  const formular = () => ({
    anbieter: anbieterVon(),
    kategorie: $("n-kategorie").value,
    name: $("n-name").value,
    leitung: $("n-leitung").value,
    datum: $("n-datum").value,
    beginn: $("n-beginn").value,
    ende: $("n-ende").value,
    ort: $("n-ort").value,
    serie: $("n-serie").checked,
    bis: $("n-serie").checked ? $("n-bis").value : "",
    zeitzone: config.zeitzone,
    nurEigen: istExternGewaehlt() || $("n-wohin").value === "eigen",
    extern: istExternGewaehlt(),
  });
  const zielKalender = (f) => (f.nurEigen ? null : f.anbieter.kalender);
  const zielName = (f) => (f.nurEigen ? "deinem Kalender" : `„${f.anbieter.kalender}“`);
  const tagDe = (d) => d.split("-").reverse().join(".");
  const meiner = (f) => f.leitung.trim().toLowerCase() === einstellungen.dauerkurse.kursleitung.toLowerCase();

  $("n-pruefen").onclick = () =>
    mitWarten($("n-pruefen"), async () => {
      const f = formular();
      const fehler = pruefeKurs(f);
      if (fehler.length) {
        $("n-hinweis").innerHTML = `<span class="fehler">${fehler.map(html).join("<br>")}</span>`;
        $("n-anlegen").hidden = true;
        return;
      }
      // Gibt es den Kurs am ersten Tag schon?
      const amTag = await api.holeKursTitel(zielKalender(f), f.datum, f.datum);
      const doppelt = amTag.find((k) => k.titel.trim() === kursTitel(f) && k.start === `${f.datum}T${f.beginn}`);
      const n = anzahlTermine(f);
      const wie = !f.serie
        ? `einmalig am ${tagDe(f.datum)}`
        : `jeden ${new Date(`${f.datum}T12:00:00Z`).toLocaleDateString("de-DE", { weekday: "long", timeZone: "UTC" })} ab ${tagDe(f.datum)}` +
          (f.bis ? ` bis ${tagDe(f.bis)} (${n} Termine)` : ", ohne Ende");
      $("n-hinweis").innerHTML =
        `<strong>${html(kursTitel(f))}</strong><br>${f.beginn}–${f.ende} Uhr, ${html(wie)}<br>` +
        `Kommt in ${html(zielName(f))} · ${html(f.kategorie)}<br>` +
        (f.extern
          ? "Vertretung bei einem anderen Verein, nur in deinem Kalender. Die Fahrzeit rechnet die Automatik aus der Adresse, dann kommen Fahrten und Vor-/Nachbereitung dazu."
          : f.nurEigen
          ? "Nur in deinem Kalender, der Anbieter-Kalender bleibt unverändert. Fahrten und Vor-/Nachbereitung kommen automatisch dazu."
          : meiner(f)
          ? "Du bist Kursleitung: Die Termine kommen danach automatisch in deinen Kalender, mit Fahrten und Vor-/Nachbereitung."
          : "Nur im Anbieter-Kalender. Mit „Ich gebe diesen Kurs“ holst du einzelne Termine später in deinen Kalender.") +
        (doppelt ? `<br><span class="fehler">Achtung: Diesen Kurs gibt es am ${tagDe(f.datum)} um ${f.beginn} schon.</span>` : "");
      $("n-anlegen").hidden = false;
    });

  $("n-anlegen").onclick = () =>
    mitWarten($("n-anlegen"), async () => {
      const f = formular();
      if (pruefeKurs(f).length) return;
      await api.legeKursAn(zielKalender(f), kursEreignis(f));
      if (f.extern) merke("vereine", JSON.stringify({ ...vereine(), [f.leitung.trim()]: f.ort.trim() }));
      $("n-anlegen").hidden = true;
      $("n-hinweis").innerHTML = "";
      zeige(
        `<div class="karte"><strong>Kurs angelegt</strong><div>${html(kursTitel(f))} in ${html(zielName(f))}.</div>` +
          (f.nurEigen
            ? `<div class="leise">Fahrten und Vor-/Nachbereitung trägt die Automatik in ein paar Minuten ein.</div>`
            : meiner(f) ? `<div class="leise">In ein paar Minuten steht er auch in deinem Kalender (die nächsten 8 Wochen, danach rollend).</div>` : "") +
          `</div>`
      );
    });
}

// --- E-Mail zum Kurs: Vorlagen mit Kurs, Datum und Uhrzeit ---
export function richteMailEin(anbieter) {
  $("bereich-mail").hidden = false;
  $("m-knoepfe").innerHTML = einstellungen.vorlagen
    .map((v) => `<button class="zweit" data-vorlage="${v.id}">${html(v.knopf)}</button>`)
    .join("");
  document.querySelectorAll("[data-vorlage]").forEach((b) => {
    b.onclick = () =>
      mitWarten(b, async () => {
        const v = einstellungen.vorlagen.find((x) => x.id === b.dataset.vorlage);
        const { betreff, text } = fuelleVorlage(v, zustand.termin, anbieter);
        // Empfänger des Anbieters: derselbe, der bei der Monatsauswertung gemerkt ist.
        const an = v.an === "anbieter" ? host.lies(`empfaenger-${anbieter.kuerzel}`) || "" : "";
        const id = await api.legeMailEntwurfAn({ an, betreff, text });
        const offen = host.oeffneEntwurf(id);
        zeige(
          `<div class="karte"><strong>Entwurf erstellt</strong><div class="leise">${html(betreff)}${an ? "" : "<br>Empfänger bitte noch eintragen."}${offen ? "" : "<br>Er liegt in deinen Entwürfen."}</div></div>`
        );
      });
  });
}

// --- Urlaub für mehrere Tage: Kürzel für alle regulären Kurse im Zeitraum ---
export function richteUrlaubEin() {
  $("bereich-urlaub").hidden = false;
  const tagDe = (d) => `${d.slice(8, 10)}.${d.slice(5, 7)}.`;
  const liste = (kurse) => `<ul>${kurse.map((t) => `<li>${tagDe(t.start)} ${t.start.slice(11)} ${html(ohneKuerzel(t.titel))}</li>`).join("")}</ul>`;
  const eingabe = () => ({ von: $("u-von").value, bis: $("u-bis").value || $("u-von").value, kuerzel: $("u-kuerzel").value });
  ["u-von", "u-bis", "u-kuerzel"].forEach((id) => $(id).addEventListener("input", () => ($("u-ausfuehren").hidden = true)));
  $("u-von").addEventListener("change", () => {
    if (!$("u-bis").value || $("u-bis").value < $("u-von").value) $("u-bis").value = $("u-von").value;
  });

  $("u-vorschau").onclick = () =>
    mitWarten($("u-vorschau"), async () => {
      const { von, bis, kuerzel } = eingabe();
      if (!von || bis < von) {
        $("u-hinweis").innerHTML = `<span class="fehler">Bitte Von und Bis wählen.</span>`;
        return;
      }
      const v = await setzeZeitraum(api, von, bis, kuerzel, einstellungen, { schreiben: false });
      const was = kuerzel ? `bekommen (${kuerzel})` : "finden wieder statt";
      $("u-hinweis").innerHTML =
        (v.aendern.length ? `<strong>${v.aendern.length} Kurse ${was}:</strong>${liste(v.aendern)}` : "In dem Zeitraum gibt es dafür keine Kurse.") +
        (v.bleiben.length ? `<strong>Bleiben so (Vertretung/Hospitation), bitte selbst klären:</strong>${liste(v.bleiben)}` : "");
      $("u-ausfuehren").hidden = !v.aendern.length;
    });

  $("u-ausfuehren").onclick = () =>
    mitWarten($("u-ausfuehren"), async () => {
      const { von, bis, kuerzel } = eingabe();
      const v = await setzeZeitraum(api, von, bis, kuerzel, einstellungen);
      $("u-ausfuehren").hidden = true;
      $("u-hinweis").innerHTML = "";
      zeige(`<div class="karte"><strong>Erledigt</strong><div>${v.aendern.length} Kurse geändert, ${v.tage.length} Tage neu geplant (Fahrten und Vor-/Nachbereitung).</div></div>`);
    });
}

// Privater Termin: Fahrzeit wählen, Hin- und Rückfahrt werden angelegt und
// wandern mit, wenn der Termin verschoben wird.
export function richteFahrzeitEin() {
  $("bereich-fahrzeit").hidden = false;
  const schluessel = `fahrzeit-ort-${zustand.termin.ort.trim().toLowerCase()}`;
  const vorschlag = zustand.termin.ort.trim() ? Number(host.lies(schluessel)) || 0 : 0;
  const markiere = () =>
    document.querySelectorAll("[data-fahrzeit]").forEach((b) => {
      const m = Number(b.dataset.fahrzeit);
      b.classList.toggle("zweit", m !== zustand.termin.fahrzeit);
    });
  const f = zustand.termin.fahrzeit;
  $("f-info").textContent =
    f > 0
      ? `Eingestellt: ${f} Minuten pro Strecke.`
      : f === -1
        ? "Die Adresse wurde nicht gefunden. Bitte wähle die Fahrzeit selbst."
        : f === 0
          ? "Keine Fahrzeit."
          : zustand.termin.ort.trim()
            ? "Die Fahrzeit wird automatisch ab deiner Adresse berechnet (innerhalb von 15 Minuten). Oder wähle selbst:"
            : vorschlag
              ? `Letztes Mal an diesem Ort: ${vorschlag} Minuten.`
              : "Wie lange fährst du pro Strecke?";
  markiere();
  document.querySelectorAll("[data-fahrzeit]").forEach((b) => {
    b.onclick = () =>
      mitWarten(b, async () => {
        const minuten = Number(b.dataset.fahrzeit);
        const v = await setzeFahrzeit(api, zustand.termin, minuten, einstellungen);
        zustand.termin = { ...zustand.termin, fahrzeit: minuten };
        if (minuten && zustand.termin.ort.trim()) merke(schluessel, minuten);
        $("f-info").textContent = minuten ? `Eingestellt: ${minuten} Minuten pro Strecke.` : "Keine Fahrzeit.";
        markiere();
        zeigeAenderungen(v, "Erledigt");
      });
  });
}

export function uebernehme(vertretung) {
  mitWarten(null, async () => {
    const { doppelt, aenderungen } = await uebernehmeKurs(api, zustand.termin, { vertretung }, einstellungen);
    if (doppelt) {
      zeige(`<div class="karte">Der Kurs steht schon in deinem Kalender: <strong>${doppelt.titel}</strong></div>`);
      return;
    }
    zeigeAenderungen(aenderungen, vertretung ? "Vertretung eingetragen" : "Kurs eingetragen");
  });
}
