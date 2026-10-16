// Fahrzeit ab der Heimadresse über OpenStreetMap (kostenlos, ohne Konto):
// Nominatim findet die Koordinaten, OSRM rechnet die Autofahrt.
// Ergebnis wird auf die nächste Viertelstunde aufgerundet.

// "Herr X (Geschäftlich), Bahnhofstr. 1, 88682 Salem" → erst alles, dann
// ohne die vorderen Teile, bis eine Adresse gefunden wird.
export function adressVarianten(ort) {
  const teile = ort.replace(/\n/g, ", ").split(",").map((t) => t.trim()).filter(Boolean);
  return teile.map((_, i) => teile.slice(i).join(", ")).filter((v) => /\d/.test(v) || v.split(" ").length > 1);
}

export const aufViertelstunde = (minuten) => Math.max(15, Math.ceil(minuten / 15) * 15);

// Nur echte Orte: keine Online-Termine und nichts Leeres.
export function hatAdresse(ort) {
  const o = (ort || "").trim();
  return o.length > 3 && !/teams|zoom|online|telefon|http/i.test(o);
}

const warte = (ms) => new Promise((r) => setTimeout(r, ms));

export function routenRechner({ heimadresse, maxMin = 240 }) {
  const koordinaten = new Map();
  let letzteAnfrage = 0;
  async function suche(text) {
    if (koordinaten.has(text)) return koordinaten.get(text);
    // Nominatim erlaubt höchstens eine Anfrage pro Sekunde.
    await warte(Math.max(0, 1100 - (Date.now() - letzteAnfrage)));
    letzteAnfrage = Date.now();
    const url = `https://nominatim.openstreetmap.org/search?format=json&limit=1&countrycodes=de&q=${encodeURIComponent(text)}`;
    const antwort = await fetch(url, { headers: { "User-Agent": "kursplaner-outlook (github.com/stehlebewegt/kursplaner-outlook)" } });
    const treffer = antwort.ok ? (await antwort.json())[0] : null;
    const k = treffer ? { lat: treffer.lat, lon: treffer.lon } : null;
    koordinaten.set(text, k);
    return k;
  }
  async function finde(ort) {
    for (const v of adressVarianten(ort)) {
      const k = await suche(v);
      if (k) return k;
    }
    return null;
  }
  async function route(a, b) {
    const url = `https://router.project-osrm.org/route/v1/driving/${a.lon},${a.lat};${b.lon},${b.lat}?overview=false`;
    const antwort = await fetch(url);
    if (!antwort.ok) return null;
    return (await antwort.json()).routes?.[0] ?? null;
  }
  let heim;
  async function fahrzeitMin(ort) {
    heim ??= await finde(heimadresse);
    const ziel = await finde(ort);
    if (!heim || !ziel) return null;
    const sekunden = (await route(heim, ziel))?.duration;
    if (!sekunden) return null;
    const minuten = sekunden / 60;
    // Praktisch zu Hause: keine Fahrt.
    if (minuten < 3) return 0;
    // Weiter als maxMin: vermutlich falsch gefundene Adresse, lieber nichts tun.
    return minuten > maxMin ? null : aufViertelstunde(minuten);
  }
  // Genaue Strecke für den Monatsbericht: { km, minuten } ohne Aufrunden auf
  // die Viertelstunde. null als Ort = zu Hause.
  fahrzeitMin.strecke = async (vonOrt, nachOrt) => {
    heim ??= await finde(heimadresse);
    const a = vonOrt === null ? heim : await finde(vonOrt);
    const b = nachOrt === null ? heim : await finde(nachOrt);
    if (!a || !b) return null;
    const r = await route(a, b);
    if (!r) return null;
    return { km: Math.round(r.distance / 100) / 10, minuten: Math.ceil(r.duration / 60) };
  };
  return fahrzeitMin;
}
