// Einheitliche Signatur für alle Kursplaner-Entwürfe an Anbieter.
// Die Daten stehen nicht im Repo (öffentlich), sondern im Entwurf
// "Kursplaner-Signatur" in Michaels Postfach, als Zeilen "Feld: Wert".
export const SIGNATUR_BETREFF = "Kursplaner-Signatur";
const FELDER = ["Name", "Motto", "Slogan", "Adresse", "Telefon", "E-Mail", "Internet", "Logo", "Hinweis"];
const TEAL = "#3a9a8f";

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

const link = (href, text) => `<a href="${esc(href)}" style="color:${TEAL};text-decoration:none">${esc(text)}</a>`;

export function signaturHtml(f) {
  if (!f) return "";
  const grau = "color:#555555;font-size:12px;line-height:17px";
  const zeilen = [];
  zeilen.push(`<div style="font-size:15px;font-weight:bold;color:#2f3d44">${esc(f.Name)}</div>`);
  if (f.Motto) zeilen.push(`<div style="font-size:13px;font-style:italic;color:#2f3d44">${esc(f.Motto)}</div>`);
  if (f.Slogan) zeilen.push(`<div style="font-size:13px;color:${TEAL};padding-bottom:6px">${esc(f.Slogan)}</div>`);
  const kontakt = [];
  if (f.Adresse) kontakt.push(esc(f.Adresse));
  if (f.Telefon) kontakt.push(link(`tel:${f.Telefon.replace(/[^\d+]/g, "")}`, f.Telefon));
  if (kontakt.length) zeilen.push(`<div style="${grau}">${kontakt.join(" · ")}</div>`);
  const netz = [];
  if (f["E-Mail"]) netz.push(link(`mailto:${f["E-Mail"]}`, f["E-Mail"]));
  if (f.Internet) netz.push(link(/^https?:/.test(f.Internet) ? f.Internet : `https://${f.Internet}`, f.Internet));
  if (netz.length) zeilen.push(`<div style="${grau}">${netz.join(" · ")}</div>`);
  const logo = f.Logo && /^https:\/\//.test(f.Logo)
    ? `<td valign="top" style="padding-right:12px"><img src="${esc(f.Logo)}" alt="" width="80" style="display:block;width:80px;height:auto;border:0"></td>`
    : "";
  const hinweis = f.Hinweis
    ? `<tr><td colspan="${logo ? 3 : 2}" style="padding-top:10px;color:#999999;font-size:10px;line-height:13px">${esc(f.Hinweis)}</td></tr>`
    : "";
  return (
    `<table cellpadding="0" cellspacing="0" border="0" style="margin-top:16px;font-family:Arial,Helvetica,sans-serif">` +
    `<tr>${logo}<td width="3" bgcolor="${TEAL}" style="background:${TEAL};width:3px">&nbsp;</td>` +
    `<td valign="top" style="padding-left:12px">${zeilen.join("")}</td></tr>${hinweis}</table>`
  );
}

export function textZuHtml(text) {
  return `<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px">${esc(text).replace(/\r?\n/g, "<br>")}</div>`;
}

// Body für Graph: mit Signatur als HTML, sonst wie bisher als Text.
export function mitSignatur(text, sig) {
  return sig ? { contentType: "html", content: textZuHtml(text) + signaturHtml(sig) } : { contentType: "text", content: text };
}
