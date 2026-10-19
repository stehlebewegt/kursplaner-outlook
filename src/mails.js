// Mails von Anbietern erkennen: Absender-Domain eines Ansprechpartners
// (Outlook-Kontakt mit Kategorie "Monatsbericht <Kürzel>"). Bei Freemail-
// Adressen zählt nur die genaue Adresse, sonst die ganze Domain.
const FREEMAIL = /^(gmail\.com|googlemail\.com|web\.de|gmx\.(de|net|at|ch)|t-online\.de|outlook\.(com|de)|hotmail\.(com|de)|live\.(com|de)|yahoo\.(com|de)|icloud\.com|me\.com|aol\.(com|de)|freenet\.de|posteo\.de|mail\.de|arcor\.de)$/i;

// kontakte: { SVE: ["a@x.de", ...], PTM: [...] } → Liste von Regeln
export function absenderRegeln(kontakte) {
  const regeln = [];
  for (const [kuerzel, adressen] of Object.entries(kontakte)) {
    for (const a of adressen) {
      const adresse = a.toLowerCase().trim();
      const domain = adresse.split("@")[1];
      if (!domain) continue;
      regeln.push(FREEMAIL.test(domain) ? { kuerzel, adresse } : { kuerzel, domain });
    }
  }
  return regeln;
}

export function anbieterVonAbsender(absender, regeln) {
  const a = (absender || "").toLowerCase();
  const domain = a.split("@")[1] || "";
  return regeln.find((r) => (r.adresse ? r.adresse === a : r.domain === domain || domain.endsWith(`.${r.domain}`)))?.kuerzel || null;
}

// Unbeantwortete Mails: von einem Anbieter, älter als "abTage", jünger als
// "bisTage", weder beantwortet noch weitergeleitet (Outlook merkt sich das in
// der Eigenschaft PidTagLastVerbExecuted: 102/103 Antwort, 104 Weiterleitung).
export const LETZTE_AKTION = "Integer 0x1081";
export function unbeantwortet(mails, regeln, jetztIso, { abTage = 2, bisTage = 14 } = {}) {
  const jetzt = Date.parse(jetztIso);
  const alter = (m) => (jetzt - Date.parse(m.empfangen)) / 86400000;
  return mails
    .map((m) => ({ ...m, kuerzel: anbieterVonAbsender(m.absender, regeln) }))
    .filter((m) => m.kuerzel && alter(m) >= abTage && alter(m) <= bisTage && ![102, 103, 104].includes(m.letzteAktion))
    .sort((a, b) => (a.empfangen < b.empfangen ? -1 : 1));
}
