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
