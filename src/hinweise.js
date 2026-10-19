// Tägliche Hinweise und Wochenvorschau, als Entwurf in Michaels Postfach.
// Reine Logik ohne Outlook, damit sie sich testen lässt.
import { kursAnbieter, leseKuerzel, ohneKuerzel } from "./planer.js";
import { leseProtokoll } from "./protokoll.js";

const datumDe = (d) => `${d.slice(8, 10)}.${d.slice(5, 7)}.`;
const zeit = (t) => `${t.start.slice(11, 16)}–${t.ende.slice(11, 16)}`;
const WOCHENTAGE = ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"];
const wochentag = (d) => WOCHENTAGE[new Date(`${d}T12:00:00Z`).getUTCDay()];
// Kurs, der wirklich stattfindet (Anbieter-Kategorie, Kürzel aktiv).
const istKurs = (t, e) => !t.generiert && kursAnbieter(t, e) && e.kuerzel[leseKuerzel(t.titel)]?.aktiv;
export const plusTage = (d, n) => new Date(Date.parse(`${d}T12:00:00Z`) + n * 86400000).toISOString().slice(0, 10);

// Kurse, die schon vorbei sind und noch keine Anwesenden im Protokoll haben.
// termine brauchen dafür ihren Text (t.text).
export function fehlendeProtokolle(termine, jetzt, einstellungen) {
  return termine
    .filter((t) => istKurs(t, einstellungen) && t.ende <= jetzt)
    .filter((t) => leseProtokoll(t.text || "").anwesend === "");
}

// Private Termine (beschäftigt), die sich mit einem Kurs oder einem Fahrt-/
// Puffer-Block überschneiden. Liefert Paare [privat, anderer].
export function ueberschneidungen(termine, einstellungen) {
  const privat = termine.filter((t) => !t.generiert && !kursAnbieter(t, einstellungen) && t.showAs !== "free");
  const andere = termine.filter((t) => t.generiert || istKurs(t, einstellungen));
  const paare = [];
  for (const p of privat)
    for (const a of andere) if (p.start < a.ende && a.start < p.ende) paare.push([p, a]);
  return paare;
}

// --- Darstellung als HTML-Mail (Outlook braucht Tabellen und Inline-Styles) ---
const esc = (x) => String(x).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const SCHRIFT = "font-family:Segoe UI,Arial,sans-serif;color:#263238;";
// Outlook am PC (Word-Darstellung) kennt bei <div> weder Innenabstand noch
// Rundung, deshalb Kopf und Rahmen als Tabelle mit bgcolor und Zell-Abstand.
const rahmen = (titel, unter, inhalt) =>
  `<table role="presentation" width="640" cellpadding="0" cellspacing="0" border="0" style="${SCHRIFT}width:640px;max-width:100%;border-collapse:collapse">` +
  `<tr><td bgcolor="#2f3d44" style="background:#2f3d44;color:#ffffff;padding:16px 20px">` +
  `<div style="font-size:20px;font-weight:700;color:#ffffff;line-height:26px">${esc(titel)}</div>` +
  (unter ? `<div style="font-size:13px;color:#d5dde1;line-height:20px;padding-top:4px">${esc(unter)}</div>` : "") +
  `</td></tr>` +
  `<tr><td style="border:1px solid #dfe3e6;border-top:0;padding:6px 20px 16px">${inhalt}</td></tr>` +
  `<tr><td style="font-size:11px;color:#8a949a;padding-top:8px">Nur ein Entwurf für dich, wird nicht gesendet. Kann gelöscht werden.</td></tr></table>`;
const abschnitt = (titel, farbe) =>
  `<div style="font-size:15px;font-weight:700;color:${farbe};border-bottom:2px solid ${farbe};padding:12px 0 4px;margin-bottom:6px">${esc(titel)}</div>`;
const zeile = (links, mitte, rechts = "", grau = false) =>
  `<tr style="${grau ? "color:#8a949a;" : ""}"><td style="padding:4px 12px 4px 0;width:120px;white-space:nowrap;font-size:13px;color:#5b6770;vertical-align:top">${links}</td>` +
  `<td style="padding:4px 8px 4px 0;font-size:14px">${mitte}</td><td style="padding:4px 0;text-align:right;white-space:nowrap">${rechts}</td></tr>`;
const tabelle = (zeilen) => `<table style="border-collapse:collapse;width:100%">${zeilen.join("")}</table>`;
const marke = (text, hinter, vorne) =>
  `<span style="display:inline-block;background:${hinter};color:${vorne};font-size:11px;font-weight:700;border-radius:9px;padding:1px 8px">${esc(text)}</span>`;
const ANBIETER_FARBE = { SVE: ["#e6f0ec", "#2f6b56"], PTM: ["#e6eef6", "#2f5d7c"] };
const tagDe = (d) => `${wochentag(d)} ${datumDe(d)}`;

export function hinweisHtml({ protokolle, paare }) {
  let inhalt = "";
  if (protokolle.length)
    inhalt +=
      abschnitt(`Kursprotokoll fehlt noch (${protokolle.length})`, "#a0612b") +
      tabelle(protokolle.map((t) => zeile(`${tagDe(t.start.slice(0, 10))} ${t.start.slice(11, 16)}`, esc(t.titel.trim()), marke("Anwesende fehlen", "#faf0e1", "#8a5a1c"))));
  if (paare.length)
    inhalt +=
      abschnitt(`Überschneidungen (${paare.length})`, "#b23b3b") +
      tabelle(paare.map(([p, a]) => zeile(tagDe(p.start.slice(0, 10)), `<b>${esc(p.titel.trim())}</b> ${zeit(p)}<br><span style="color:#5b6770">mit ${esc(a.titel.trim())} ${zeit(a)}</span>`)));
  const anzahl = protokolle.length + paare.length;
  return rahmen(`${anzahl} ${anzahl === 1 ? "Hinweis" : "Hinweise"} aus dem Kursplaner`, "", inhalt);
}

// Daten einer Woche (montag = "JJJJ-MM-TT"): pro Tag Kurse, Termine, Fahrzeit.
export function wochenDaten(termine, montag, einstellungen) {
  const fahrt = einstellungen.kategorieFahrt;
  const tage = [];
  for (let i = 0; i < 7; i++) {
    const d = plusTage(montag, i);
    const amTag = termine.filter((t) => t.start.slice(0, 10) === d).sort((a, b) => (a.start < b.start ? -1 : 1));
    const fahrten = amTag.filter((t) => t.generiert && (t.kategorien || []).includes(fahrt));
    tage.push({
      datum: d,
      kurse: amTag
        .filter((t) => !t.generiert && kursAnbieter(t, einstellungen))
        .map((t) => {
          const kz = leseKuerzel(t.titel);
          const bed = einstellungen.kuerzel[kz];
          return { t, anbieter: kursAnbieter(t, einstellungen).kuerzel, aktiv: !!bed?.aktiv, zusatz: kz ? bed?.bedeutung || kz : "" };
        }),
      privat: amTag.filter((t) => !t.generiert && !kursAnbieter(t, einstellungen) && t.showAs !== "free"),
      fahrten: fahrten.length,
      fahrMin: fahrten.reduce((s, t) => s + (Date.parse(t.ende) - Date.parse(t.start)) / 60000, 0),
    });
  }
  return {
    tage,
    kurse: tage.reduce((s, x) => s + x.kurse.filter((k) => k.aktiv).length, 0),
    fahrStd: Math.round((tage.reduce((s, x) => s + x.fahrMin, 0) / 60) * 10) / 10,
  };
}

export function wochenHtml(termine, montag, einstellungen) {
  const w = wochenDaten(termine, montag, einstellungen);
  let inhalt = "";
  for (const tag of w.tage) {
    const frei = !tag.kurse.length && !tag.privat.length;
    inhalt += abschnitt(tagDe(tag.datum), frei ? "#9aa4a9" : "#35505a");
    if (frei) {
      inhalt += `<div style="font-size:13px;color:#8a949a;padding:2px 0 4px">frei</div>`;
      continue;
    }
    const z = [
      ...tag.kurse.map(({ t, anbieter, aktiv, zusatz }) => {
        const [h, v] = ANBIETER_FARBE[anbieter] || ["#eceff1", "#37474f"];
        const titel = aktiv ? esc(ohneKuerzel(t.titel)) : `<s>${esc(ohneKuerzel(t.titel))}</s>`;
        return zeile(zeit(t), titel, marke(anbieter, h, v) + (zusatz ? " " + marke(zusatz, "#f3e7d3", "#8a5a1c") : ""), !aktiv);
      }),
      ...tag.privat.map((t) => zeile(zeit(t), esc(t.titel.trim()), marke("privat", "#eceff1", "#5b6770"))),
    ];
    if (tag.fahrten) z.push(zeile("", `<span style="color:#5b6770;font-size:13px">🚗 ${Math.round(tag.fahrMin)} Min. Fahrzeit (${tag.fahrten} ${tag.fahrten === 1 ? "Fahrt" : "Fahrten"})</span>`));
    inhalt += tabelle(z);
  }
  const bis = plusTage(montag, 6);
  return rahmen(`Deine Woche ab ${tagDe(montag)}`, `${w.kurse} Kurse · ${String(w.fahrStd).replace(".", ",")} Std. Fahrzeit · bis ${tagDe(bis)}`, inhalt);
}
