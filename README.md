# Kursplaner für Outlook

Kern von „Tag neu planen“ (Schritt 2 im Plan).

- `src/einstellungen.js` – Anbieter, Fahrzeiten, Puffer, Kürzel, Kategorien (aus `planung/einstellungen.md`)
- `src/planer.js` – berechnet aus den Kursen eines Tages alle Blöcke für Fahrt, Vor- und Nachbereitung und vergleicht sie mit den vorhandenen
- `test/planer.test.js` – Tests mit echten Kurstagen aus Oktober 2026

Tests ausführen: `npm test` (Node 20 oder neuer, keine weiteren Pakete).

Offen: Outlook-Hülle (Manifest, Aufgabenbereich, Zugriff auf die Kalender über Microsoft Graph). Dafür nötig sind ein GitHub-Repository fürs Hosting und eine App-Registrierung in Microsoft 365.
