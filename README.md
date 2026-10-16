# Kursplaner für Outlook

Outlook-Add-in für Michaels Kursplanung als Übungsleiter.

**Was es kann (Stand Schritt 2):**
- Im Anbieter-Kalender (SVE, PTM): „Ich gebe diesen Kurs“ / „Ich vertrete diesen Kurs (V)“ übernimmt den Kurs in den eigenen Kalender.
- Im eigenen Kalender: Kurs als Urlaub (U), Abgesagt (A), Ferien (F) markieren oder wieder auf „findet statt“ setzen.
- „Tag neu planen“: Anfahrt, Vorbereitung, Puffer, Nachbereitung und Rückfahrt für den ganzen Tag neu berechnen, mit Vorschau. Angefasst werden nur Blöcke, die das Add-in angelegt oder von den alten, von Hand angelegten übernommen hat.

**Automatik (Schritt 3):** `automatik/lauf.js` läuft alle 15 Minuten (7–23 Uhr) in GitHub Actions und plant die nächsten 8 Wochen neu. So ziehen die Blöcke auch mit, wenn am iPhone ein Kurs verschoben, gelöscht oder mit (U)/(A)/(F) markiert wird. Braucht in Entra die Anwendungsberechtigung `Calendars.ReadWrite`, einen geheimen Clientschlüssel und in GitHub die Secrets `KURSPLANER_SECRET` und `KURSPLANER_POSTFACH`. Ohne die Variable `SCHREIBEN=1` zeigt sie nur an, was sie tun würde. Die Ausgabe ist öffentlich, deshalb nur Datum und Uhrzeit, keine Titel.

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
