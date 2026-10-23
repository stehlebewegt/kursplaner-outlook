// Einstellungen, die Michael im Kursplaner selbst ändert. Sie liegen nicht im
// (öffentlichen) Code, sondern in seinem Postfach: versteckter Ordner
// „Kursplaner-Daten“, Nachricht „Kursplaner-Einstellungen“ mit JSON. So lesen
// Add-in, Kursplaner-Fenster und Automatik dieselben Werte.
// Was dort fehlt, kommt aus src/einstellungen.js.

export const ORDNER = "Kursplaner-Daten";
export const BETREFF = "Kursplaner-Einstellungen";

// Diese Felder darf der Kursplaner überschreiben.
export const ZAHLEN = ["vorbereitungMin", "nachbereitungMin", "pauseHeimfahrtMin", "maxLueckeAmStueckMin", "fahrzeitRundungMin", "erinnerungVorFahrtMin"];
export const BEARBEITBAR = [...ZAHLEN, "anbieter", "material", "vorlagen", "monatsbericht"];

export function mitEigenen(basis, eigene) {
  const e = { ...basis };
  if (!eigene || typeof eigene !== "object") return e;
  for (const k of ZAHLEN) if (Number.isFinite(eigene[k]) && eigene[k] >= 0) e[k] = eigene[k];
  if (Array.isArray(eigene.anbieter) && eigene.anbieter.length)
    e.anbieter = eigene.anbieter.filter((a) => a && a.kuerzel && a.name).map((a) => ({
      ...a,
      fahrzeitMin: Number(a.fahrzeitMin) || 0,
      kategorien: Array.isArray(a.kategorien) ? a.kategorien : [],
      ortErkennung: Array.isArray(a.ortErkennung) ? a.ortErkennung : [],
    }));
  if (Array.isArray(eigene.material)) e.material = eigene.material.filter(Boolean);
  if (Array.isArray(eigene.vorlagen)) e.vorlagen = eigene.vorlagen.filter((v) => v && v.id && v.knopf);
  if (eigene.monatsbericht && typeof eigene.monatsbericht === "object") e.monatsbericht = { ...basis.monatsbericht, ...eigene.monatsbericht };
  return e;
}

// Wendet die eigenen Werte auf das gemeinsame einstellungen-Objekt an
// (alle Module nutzen dasselbe Objekt).
export function wendeAn(einstellungen, eigene) {
  Object.assign(einstellungen, mitEigenen(einstellungen, eigene));
}

const ENTITAETEN = { "&quot;": '"', "&amp;": "&", "&lt;": "<", "&gt;": ">", "&#39;": "'", "&nbsp;": " " };
export function leseJson(text) {
  let t = String(text || "").trim();
  if (t.startsWith("<")) t = t.replace(/<br\s*\/?>/gi, "\n").replace(/<[^>]+>/g, "").replace(/&[a-z#0-9]+;/gi, (x) => ENTITAETEN[x] ?? x).trim();
  try {
    return JSON.parse(t);
  } catch {
    return null;
  }
}

// graph(methode, pfad, body): "/me" im Add-in/Fenster, "" in der Automatik
// (dort hängt graph() selbst /users/<postfach> davor).
async function ordnerVon(graph, me) {
  const { value } = await graph("GET", `${me}/mailFolders?includeHiddenFolders=true&$top=200&$select=id,displayName`);
  return value.find((f) => f.displayName === ORDNER) || null;
}

export async function holeEigene(graph, me = "") {
  const ordner = await ordnerVon(graph, me);
  if (!ordner) return null;
  const { value } = await graph("GET", `${me}/mailFolders/${ordner.id}/messages?$top=10&$select=id,subject,body`);
  const m = value.find((x) => x.subject === BETREFF);
  return m ? leseJson(m.body?.content) : null;
}

export async function speichereEigene(graph, daten, me = "") {
  let ordner = await ordnerVon(graph, me);
  ordner ??= await graph("POST", `${me}/mailFolders`, { displayName: ORDNER, isHidden: true });
  const { value } = await graph("GET", `${me}/mailFolders/${ordner.id}/messages?$top=10&$select=id,subject`);
  const m = value.find((x) => x.subject === BETREFF);
  const body = { contentType: "text", content: JSON.stringify(daten, null, 1) };
  if (m) await graph("PATCH", `${me}/messages/${m.id}`, { body });
  else await graph("POST", `${me}/mailFolders/${ordner.id}/messages`, { subject: BETREFF, body });
}
