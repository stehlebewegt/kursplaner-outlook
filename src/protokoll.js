// Kursprotokoll im Termintext. Das Format bleibt das von Michaels Vorlage,
// damit es am iPhone lesbar und von Hand ergänzbar bleibt:
//   Offizielle Erstvertretung: für Eva
//   Thema: …
//   Material: …
//   Teilnehmer gesamt: 15
//   Anwesend: 10 / Entschuldigt: 4 / Unentschuldigt: 1
//   Vorkommnis: nein
//   Besonderheiten: … (darf mehrzeilig sein, steht deshalb am Schluss)

export const FELDER = [
  ["erstvertretung", "Offizielle Erstvertretung"],
  ["thema", "Thema"],
  ["material", "Material"],
  ["gesamt", "Teilnehmer gesamt"],
  ["anwesend", "Anwesend"],
  ["entschuldigt", "Entschuldigt"],
  ["unentschuldigt", "Unentschuldigt"],
  ["vorkommnis", "Vorkommnis"],
  ["besonderheiten", "Besonderheiten"],
];

const ZAHLEN = ["anwesend", "entschuldigt", "unentschuldigt"];

export function leseProtokoll(text = "") {
  const p = Object.fromEntries(FELDER.map(([k]) => [k, ""]));
  const zeilen = text.replace(/\r/g, "").split("\n");
  for (let i = 0; i < zeilen.length; i++) {
    const m = zeilen[i].match(/^\s*([^:]+):\s*(.*)$/);
    if (!m) continue;
    const feld = FELDER.find(([, label]) => label.toLowerCase() === m[1].trim().toLowerCase());
    if (!feld) continue;
    if (feld[0] === "besonderheiten") {
      p.besonderheiten = [m[2], ...zeilen.slice(i + 1)].join("\n").trim();
      break;
    }
    p[feld[0]] = m[2].trim();
  }
  return p;
}

// Gesamt = anwesend + entschuldigt + unentschuldigt, sobald eine Zahl da ist.
export function gesamt(p) {
  const zahlen = ZAHLEN.map((k) => p[k]).filter((v) => String(v).trim() !== "");
  if (!zahlen.length) return p.gesamt || "";
  return String(zahlen.reduce((s, v) => s + (parseInt(v, 10) || 0), 0));
}

export function schreibeProtokoll(p) {
  const w = { ...p, gesamt: gesamt(p) };
  const z = (k) => `${FELDER.find(([f]) => f === k)[1]}: ${w[k] ?? ""}`.trimEnd();
  return [
    z("erstvertretung"), "",
    z("thema"), "",
    z("material"), "",
    z("gesamt"), z("anwesend"), z("entschuldigt"), z("unentschuldigt"), "",
    z("vorkommnis"), "",
    z("besonderheiten"),
  ].join("\n");
}

// Material steht als "Matte, Stab" im Text.
export const teileMaterial = (text) => (text || "").split(/[,;]/).map((m) => m.trim()).filter(Boolean);

// Aus früheren Kursen: letzte Protokolle dieses Kurses, bekannte Themen und
// bekanntes Material. "termine" sind Kurse mit Text, "kursTitel" ohne Kürzel.
export function auswertungVerlauf(termine, kursTitel, vorDatum, standardMaterial = [], ohneKuerzel = (t) => t) {
  const mitProtokoll = termine
    .filter((t) => t.start.slice(0, 10) < vorDatum)
    .map((t) => ({ datum: t.start.slice(0, 10), titel: ohneKuerzel(t.titel).trim(), p: leseProtokoll(t.text) }))
    .filter((t) => t.p.thema)
    .sort((a, b) => (a.datum < b.datum ? 1 : -1));
  const letzte = mitProtokoll.filter((t) => t.titel === kursTitel.trim()).slice(0, 3);
  const zaehle = (liste) => {
    const n = new Map();
    for (const x of liste) n.set(x, (n.get(x) || 0) + 1);
    return [...n.keys()].sort((a, b) => n.get(b) - n.get(a) || a.localeCompare(b, "de"));
  };
  const themen = zaehle(mitProtokoll.map((t) => t.p.thema));
  // "Gymnastik Ball" und "Gymnastikball" sind dasselbe: die Schreibweise aus
  // der Grundliste gewinnt, sonst die zuerst gefundene.
  const schluessel = (m) => m.toLowerCase().replace(/[\s-]/g, "");
  const nachSchluessel = new Map();
  for (const m of [...standardMaterial, ...mitProtokoll.flatMap((t) => teileMaterial(t.p.material))]) {
    if (!nachSchluessel.has(schluessel(m))) nachSchluessel.set(schluessel(m), m);
  }
  const material = [...nachSchluessel.values()].sort((a, b) => a.localeCompare(b, "de"));
  return { letzte, themen, material };
}

// Ausfallgrund bei abgesagten Kursen: eigene erste Zeile im Termintext
// („Ausfallgrund: Halle zu“), der Rest (Protokoll, eigene Notizen) bleibt.
// Gründe wie im Reha-Cockpit; Urlaub und Ferien haben eigene Kürzel.
export const AUSFALLGRUENDE = ["Absage", "Halle zu", "Krank", "Feiertag", "Sonstiges"];
const GRUND_ZEILE = /^\s*Ausfallgrund:\s*(.*)$/im;
export const leseAusfallgrund = (text = "") => (text.match(GRUND_ZEILE)?.[1] || "").trim();
export function mitAusfallgrund(text = "", grund = "") {
  const rest = text.replace(/\r/g, "").replace(/^\s*Ausfallgrund:.*(\n|$)/im, "").replace(/^\n+/, "");
  return grund ? `Ausfallgrund: ${grund}\n${rest}`.trimEnd() : rest.trimEnd();
}
