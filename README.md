# Kursplaner für Outlook

Outlook-Add-in für Michaels Kursplanung als Übungsleiter.

**Was es kann (Stand Schritt 2):**
- Im Anbieter-Kalender (SVE, PTM): „Ich gebe diesen Kurs“ / „Ich vertrete diesen Kurs (V)“ übernimmt den Kurs in den eigenen Kalender.
- Im eigenen Kalender: Kurs als Urlaub (U), Abgesagt (A), Ferien (F) markieren oder wieder auf „findet statt“ setzen.
- „Tag neu planen“: Anfahrt, Vorbereitung, Puffer, Nachbereitung und Rückfahrt für den ganzen Tag neu berechnen, mit Vorschau. Angefasst werden nur Blöcke, die das Add-in angelegt oder von den alten, von Hand angelegten übernommen hat.

## Aufbau
- `manifest.xml` – Beschreibung für Outlook (Button „Kursplaner“ in geöffneten Terminen)
- `taskpane.html`, `web/` – Oberfläche, Anmeldung und Kalenderzugriff über Microsoft Graph
- `src/einstellungen.js` – Anbieter, Fahrzeiten, Puffer, Kürzel, Kategorien
- `src/planer.js` – Berechnung der Blöcke eines Tages
- `src/aktionen.js` – was die Buttons tun
- `test/` – Tests mit echten Kurstagen aus Oktober 2026

Tests: `npm test` (Node 20+, keine Pakete nötig). Gehostet über GitHub Pages: https://stehlebewegt.github.io/kursplaner-outlook/

## Einrichten
1. App-Registrierung in Microsoft Entra (SPA, Umleitungs-URI `brk-multihub://stehlebewegt.github.io`, delegiert `Calendars.ReadWrite`, `Mail.ReadWrite`), IDs in `web/config.js` eintragen.
2. In Outlook: Add-ins abrufen → Meine Add-ins → Benutzerdefiniertes Add-in hinzufügen → Aus Datei → `manifest.xml`.
