// Einheitliche Signatur für alle Kursplaner-Entwürfe an Anbieter.
// Die Daten stehen nicht im Repo (öffentlich), sondern im Entwurf
// "Kursplaner-Signatur" in Michaels Postfach, als Zeilen "Feld: Wert".
export const SIGNATUR_BETREFF = "Kursplaner-Signatur";
const FELDER = ["Name", "Motto", "Slogan", "Adresse", "Telefon", "E-Mail", "Internet", "Logo", "Hinweis"];

const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

// Graph liefert den Entwurf je nach Weg als Text oder HTML.
export function htmlZuText(html) {
  return String(html)
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div|li|tr)>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&amp;/g, "&");
}

export function leseSignatur(text) {
  const f = {};
  const re = new RegExp(`^\\s*(${FELDER.join("|")}):[ \\t]*(.*)$`);
  for (const zeile of String(text || "").split(/\r?\n/)) {
    const m = zeile.match(re);
    if (m && m[2].trim()) f[m[1]] = m[2].trim();
  }
  return f.Name ? f : null;
}

const link = (href, text) => `<a href="${esc(href)}">${esc(text)}</a>`;
const SCHRIFT = "font-family:Aptos,Calibri,Arial,sans-serif;font-size:11pt";

// Nachbau von Michaels Outlook-Signatur: Strichlinie, fett-kursive Kopfzeile
// „Name – Motto“, Slogan, Adresse, Kontaktzeile mit •, Vertraulichkeitshinweis.
export function signaturHtml(f) {
  if (!f) return "";
  const p = (inhalt) => `<p style="margin:0;${SCHRIFT}">${inhalt}</p>`;
  const punkt = "<b><i> • </i></b>";
  const z = [p("- ".repeat(35).trim()), p("&nbsp;")];
  z.push(p(`<b><i>${esc(f.Motto ? `${f.Name} – ${f.Motto}` : f.Name)}</i></b>`));
  if (f.Slogan) z.push(p(esc(f.Slogan)));
  if (f.Adresse) z.push(p("&nbsp;"), p(esc(f.Adresse)));
  const kontakt = [];
  if (f.Telefon) kontakt.push(`Telefon: ${esc(f.Telefon)}`);
  if (f["E-Mail"]) kontakt.push(`E-Mail: ${link(`mailto:${f["E-Mail"]}`, f["E-Mail"])}`);
  if (f.Internet) kontakt.push(`Internet: ${link(/^https?:/.test(f.Internet) ? f.Internet : `https://${f.Internet}`, f.Internet)}`);
  if (kontakt.length) z.push(p(kontakt.join(punkt)));
  if (f.Logo && /^https:\/\//.test(f.Logo)) z.push(p("&nbsp;"), p(`<img src="${esc(f.Logo)}" alt="" width="120" style="border:0">`));
  if (f.Hinweis) z.push(p("&nbsp;"), p(esc(f.Hinweis).replace(/\s+(?=This e-?mail)/i, "<br>")));
  return `<div>${z.join("")}</div>`;
}

export function textZuHtml(text) {
  return `<div style="${SCHRIFT}">${esc(text).replace(/\r?\n/g, "<br>")}</div>`;
}

// Body für Graph: mit Signatur als HTML, sonst wie bisher als Text.
export function mitSignatur(text, sig) {
  return sig ? { contentType: "html", content: textZuHtml(text) + signaturHtml(sig) } : { contentType: "text", content: text };
}
