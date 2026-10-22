// Michaels Einstellungen (Stand 16.10.2026), siehe planung/einstellungen.md.
// Später pflegt er das im Add-in selbst; bis dahin hier als Startwerte.
export const einstellungen = {
  anbieter: [
    {
      kuerzel: "SVE",
      name: "Sportverein Ettenkirch",
      kalender: "Alle SVE Kurse",
      ort: "SVE Sport- & Gesundheitscenter, Ettenkircher Straße 17, 88048 Friedrichshafen, Deutschland",
      ortErkennung: ["Ettenkircher", "SVE"],
      fahrzeitMin: 20,
      kategorien: ["10 SVE: RehaSport", "11 SVE: Vereinskurs"],
      kategorieVertretung: "19 Vertretung Kurs SVE",
      // An gesetzlichen Feiertagen (BW) keine Kurse: reguläre Kurse bekommen (F).
      feiertagFrei: true,
      empfaenger: "", // wird in Outlook hinterlegt, nicht im öffentlichen Code
    },
    {
      kuerzel: "PTM",
      name: "Physio-Team Markdorf",
      kalender: "Alle PTM Kurse",
      ort: "Physio-Team Markdorf, Ravensburgerstraße 36, 88677 Markdorf, Deutschland",
      ortErkennung: ["Ravensburger", "Physio-Team"],
      fahrzeitMin: 10,
      kategorien: ["01 PTM: PUR Kurs", "02 PTM: FLEX Kurs"],
      kategorieVertretung: "09 Vertretung Kurs PTM",
      // In den Schulferien (BW) keine Kurse: reguläre Kurse bekommen automatisch (F).
      ferienFrei: true,
      feiertagFrei: true,
      empfaenger: "", // wird in Outlook hinterlegt, nicht im öffentlichen Code
    },
  ],
  // Fahrzeit direkt zwischen zwei Anbietern; fehlt ein Paar, gilt die Anfahrt zum Ziel.
  // Werte aus OpenStreetMap, auf 5 Minuten aufgerundet (Probe strecken_probe).
  fahrzeitZwischen: { "SVE-PTM": 20 },
  vorbereitungMin: 15,
  nachbereitungMin: 15,
  // Lücke bis zu dieser Länge: Kurse gelten als "am Stück".
  maxLueckeAmStueckMin: 30,
  // Pause zwischen zwei Kursen: erst ab so vielen Minuten fährt Michael heim
  // (und nur, wenn die Zeit für Hin- und Rückweg reicht).
  pauseHeimfahrtMin: 120,
  kategorieFahrt: "90 🚗 Fahrzeit",
  kategoriePuffer: "Vor-/Nachbereitung",
  // Private Termine mit dieser Kategorie (z. B. Fortbildung über mehrere Tage):
  // Michael bleibt über Nacht, also keine Rückfahrt abends, wenn am nächsten
  // Tag wieder so ein Termin ist, und keine Anfahrt morgens nach so einem Tag.
  kategorieUebernachtung: "Übernachtung",
  // Schon vorhandene Fahrt-Blöcke bekommen den Google-Maps-Link nachträglich
  // (Michael hat am 18.10.2026 zugestimmt).
  routenLinkNachtragen: true,
  // Kursfahrten: tatsächliche Strecke (km, Minuten) aus OpenStreetMap in die
  // Notiz der Fahrt-Blöcke schreiben (Michaels Wunsch 22.10.2026).
  streckenNotiz: true,
  // Kursfahrten: berechnete Fahrzeit wird auf so viele Minuten aufgerundet
  // (Michael 22.10.2026: 5 Minuten).
  fahrzeitRundungMin: 5,
  // Erinnerung (Minuten) vor Anfahrten und Fahrten zwischen Anbietern; 0 = aus.
  erinnerungVorFahrtMin: 10,
  // Erinnerung bei Kursende (an der Nachbereitung): Protokoll ausfüllen.
  erinnerungProtokoll: true,
  // Ferien-Kürzel (F) automatisch setzen; false = nur anzeigen (Probelauf).
  ferienAutomatisch: true,
  // Entwürfe für Michael (nichts wird gesendet): morgens Hinweise zu fehlenden
  // Kursprotokollen und Überschneidungen, sonntags die Wochenvorschau.
  hinweise: { morgensAb: 6, protokollTage: 7, ueberschneidungTage: 14, wocheAb: 18, mailAbTage: 2, mailBisTage: 14 },
  // Mails von Anbietern bekommen im Posteingang diese Kategorie (Farbe).
  mailKategorien: { SVE: "SVE", PTM: "PTM" },
  // Fahrten mit dieser Kategorie kommen nicht in den Monatsbericht.
  kategorieNichtAbrechenbar: "Fahrtkosten nicht abrechenbar",
  // Material zum Anhaken im Kursprotokoll. Was im Protokoll neu dazukommt,
  // erscheint beim nächsten Mal automatisch mit.
  material: ["Airpad", "Gymnastikball", "Hanteln", "Hocker", "Matte", "Pilatesball", "Stab", "Theraband"],
  // Dauerkurse: Kurse mit "(kursleitung)" im Titel kommen aus den
  // Anbieter-Kalendern automatisch in den eigenen Kalender.
  // schreiben: false = die Automatik zeigt nur an, was sie tun würde.
  dauerkurse: { kursleitung: "Michael", schreiben: true },
  // Monatsbericht: an diesem Tag legt die Automatik pro Anbieter einen
  // E-Mail-Entwurf für den Vormonat an (wird nie selbst verschickt).
  // teilnehmer: true = Spalte mit Teilnehmerzahlen im Bericht.
  // {monat} und {anbieter} werden ersetzt.
  monatsbericht: {
    // Am Tag X kommt der laufende Monat (bis dahin geleistet, danach geplant)
    // plus Nachtrag zum Vormonat. {nachtrag} wird ersetzt, falls es einen gibt.
    tag: 15,
    nachtragText: " Am Ende stehen außerdem die Nachträge für {vormonat}, also was sich seit der letzten Auswertung noch getan hat.",
    teilnehmer: false,
    betreff: "Monatsauswertung {monat} – Michael Stehle ({anbieter})",
    // Ansprechpartner: Outlook-Kontakt mit Kategorie "Monatsbericht <Kürzel>"
    // (Vorname und E-Mail). {anrede} = "Anna" bzw. "Anna und Ben".
    text:
      "Hallo {anrede},\n\n" +
      "ich hoffe, bei dir ist alles gut! Im Anhang findest du meine Monatsauswertung für {monat}, " +
      "mit den Kursen, die ich bei euch bisher gegeben habe und noch geben werde, und den Fahrten dazu.{nachtrag}\n\n" +
      "Wenn etwas nicht passt oder du Fragen hast, melde dich gern jederzeit bei mir.\n\n" +
      "Herzliche Grüße und bis bald\nMichael Stehle",
    // Für mehrere Ansprechpartner oder wenn keiner hinterlegt ist.
    textGruppe:
      "Hallo {anrede},\n\n" +
      "ich hoffe, bei euch ist alles gut! Im Anhang findet ihr meine Monatsauswertung für {monat}, " +
      "mit den Kursen, die ich bei euch bisher gegeben habe und noch geben werde, und den Fahrten dazu.{nachtrag}\n\n" +
      "Wenn etwas nicht passt oder ihr Fragen habt, meldet euch gern jederzeit bei mir.\n\n" +
      "Herzliche Grüße und bis bald\nMichael Stehle",
  },
  // E-Mail-Vorlagen im Kursplaner (Bereich „E-Mail zum Kurs“). an: "anbieter" =
  // gemerkter Empfänger des Anbieters (wie bei der Monatsauswertung), "" = leer.
  vorlagen: [
    {
      id: "vertretung",
      knopf: "Vertretung anfragen",
      an: "",
      betreff: "Vertretung gesucht: {kurs} am {wochentag}, {kurzdatum}",
      text:
        "Hallo zusammen,\n\n" +
        "ich suche eine Vertretung für meinen Kurs {kurs} am {wochentag}, {datum}, von {zeit} ({anbieter}).\n\n" +
        "Hätte jemand von euch Zeit und Lust? Gebt mir gern kurz Bescheid.\n\n" +
        "Vielen Dank und viele Grüße\nMichael",
    },
    {
      id: "absage",
      knopf: "Kurs absagen",
      an: "anbieter",
      betreff: "Absage: {kurs} am {wochentag}, {kurzdatum}",
      text:
        "Hallo,\n\n" +
        "leider muss ich meinen Kurs {kurs} am {wochentag}, {datum}, von {zeit} absagen.\n\n" +
        "Kannst du bitte die Teilnehmer informieren? Danke dir!\n\n" +
        "Viele Grüße\nMichael",
    },
    {
      id: "krank",
      knopf: "Krank melden",
      an: "anbieter",
      betreff: "Krankmeldung: {kurs} am {wochentag}, {kurzdatum}",
      text:
        "Hallo,\n\n" +
        "ich bin leider krank und kann meinen Kurs {kurs} am {wochentag}, {datum}, von {zeit} nicht geben.\n\n" +
        "Ich melde mich, sobald ich wieder fit bin. Kannst du bitte die Teilnehmer informieren?\n\n" +
        "Viele Grüße\nMichael",
    },
  ],
  // Kürzel vor dem Titel. aktiv = Kurs findet statt (bekommt Fahrt und Puffer).
  kuerzel: {
    "": { bedeutung: "regulär", aktiv: true },
    V: { bedeutung: "Vertretung", aktiv: true },
    U: { bedeutung: "Urlaub", aktiv: false },
    A: { bedeutung: "abgesagt", aktiv: false },
    F: { bedeutung: "Ferien", aktiv: false },
    H: { bedeutung: "Hospitation", aktiv: true },
  },
};
