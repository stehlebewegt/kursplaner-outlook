// Posteingang: Mails von SVE/PTM bekommen eine Kategorie (Farbe), und
// unbeantwortete Mails von Anbietern kommen in die Hinweise am Morgen.
// Erkannt wird über die Ansprechpartner (Outlook-Kontakte "Monatsbericht <Kürzel>").
// Öffentliche Ausgabe nur mit Zahlen.
import { holeKontakte } from "./monatsbericht.js";
import { absenderRegeln, anbieterVonAbsender, unbeantwortet, LETZTE_AKTION } from "../src/mails.js";

export async function ladeRegeln(graph, einstellungen, kontaktPostfaecher = []) {
  const kontakte = {};
  for (const a of einstellungen.anbieter) {
    const k = (await holeKontakte(graph, a.kuerzel, kontaktPostfaecher)) || [];
    kontakte[a.kuerzel] = k.map((x) => x.emailAddresses?.[0]?.address).filter(Boolean);
  }
  return absenderRegeln(kontakte);
}

const vorTagen = (n) => new Date(Date.now() - n * 86400000).toISOString().slice(0, 19) + "Z";

// Farbe für die Kategorien anlegen, falls es sie noch nicht gibt (ohne
// Berechtigung dafür erscheinen sie trotzdem, nur ohne Farbe).
const FARBEN = { SVE: "preset4", PTM: "preset7", "20 Vertretung andere Vereine": "preset8" };
export async function kategorienAnlegen(graph, namen) {
  try {
    const da = (await graph("GET", "/outlook/masterCategories")).value.map((k) => k.displayName);
    for (const n of namen.filter((n) => !da.includes(n)))
      await graph("POST", "/outlook/masterCategories", { displayName: n, color: FARBEN[n] || "preset0" });
  } catch {
    /* ohne MailboxSettings-Berechtigung eben ohne Farbe */
  }
}

let kategorienGeprueft = false;
export async function sortiere({ graph, regeln, einstellungen, ausgabe, tage = 2 }) {
  if (!einstellungen.mailKategorien || !regeln.length) return;
  const name = (kuerzel) => einstellungen.mailKategorien[kuerzel];
  if (!kategorienGeprueft) {
    await kategorienAnlegen(graph, Object.values(einstellungen.mailKategorien));
    kategorienGeprueft = true;
  }
  const mails = (
    await graph("GET", `/mailFolders/inbox/messages?$top=100&$select=from,categories&$filter=receivedDateTime ge ${vorTagen(tage)}`)
  ).value;
  let n = 0;
  for (const m of mails) {
    const k = anbieterVonAbsender(m.from?.emailAddress?.address, regeln);
    if (!k || !name(k) || (m.categories || []).includes(name(k))) continue;
    await graph("PATCH", `/messages/${m.id}`, { categories: [...(m.categories || []), name(k)] });
    n++;
  }
  ausgabe(`Posteingang: ${regeln.length} Absender-Regeln, ${n} Mails neu markiert`);
}

// Für die Hinweise: { kuerzel, absender, name, betreff, empfangen }
export async function offeneMails({ graph, regeln, jetztIso, einstellungen }) {
  if (!regeln.length) return [];
  const e = einstellungen.hinweise || {};
  const mails = (
    await graph(
      "GET",
      `/mailFolders/inbox/messages?$top=200&$select=from,subject,receivedDateTime&$filter=receivedDateTime ge ${vorTagen(e.mailBisTage ?? 14)}` +
        `&$expand=singleValueExtendedProperties($filter=id eq '${LETZTE_AKTION}')`
    )
  ).value.map((m) => ({
    id: m.id,
    absender: m.from?.emailAddress?.address,
    name: m.from?.emailAddress?.name || m.from?.emailAddress?.address || "",
    betreff: m.subject || "",
    empfangen: m.receivedDateTime,
    letzteAktion: Number(m.singleValueExtendedProperties?.[0]?.value) || null,
  }));
  return unbeantwortet(mails, regeln, jetztIso, { abTage: e.mailAbTage ?? 2, bisTage: e.mailBisTage ?? 14 });
}
