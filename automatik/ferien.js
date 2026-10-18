// Schulferien Baden-Württemberg aus OpenHolidays (öffentlich, offiziell gepflegt,
// neue Schuljahre kommen dort von selbst dazu). Ohne Antwort: keine Ferien,
// dann ändert sich auch nichts.
export async function holeFerien(von, bis, land = "DE-BW") {
  const url =
    `https://openholidaysapi.org/SchoolHolidays?countryIsoCode=DE&subdivisionCode=${land}` +
    `&languageIsoCode=DE&validFrom=${von}&validTo=${bis}`;
  const antwort = await fetch(url, { headers: { accept: "application/json" } });
  if (!antwort.ok) throw new Error(`Ferien-Abruf: ${antwort.status}`);
  return (await antwort.json()).map((f) => ({ von: f.startDate, bis: f.endDate, name: f.name?.[0]?.text || "" }));
}
