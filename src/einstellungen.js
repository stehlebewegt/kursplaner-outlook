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
      empfaenger: "", // wird in Outlook hinterlegt, nicht im öffentlichen Code
    },
    {
      kuerzel: "PTM",
      name: "Physio-Team Markdorf",
      kalender: "Alle PTM Kurse",
      ort: "Physio-Team Markdorf, Ravensburgerstraße 36, 88677 Markdorf, Deutschland",
      ortErkennung: ["Ravensburger", "Physio-Team"],
      fahrzeitMin: 15,
      kategorien: ["01 PTM: PUR Kurs", "02 PTM: FLEX Kurs"],
      kategorieVertretung: "09 Vertretung Kurs PTM",
      empfaenger: "", // wird in Outlook hinterlegt, nicht im öffentlichen Code
    },
  ],
  // Fahrzeit direkt zwischen zwei Anbietern; fehlt ein Paar, gilt die Anfahrt zum Ziel.
  fahrzeitZwischen: {},
  vorbereitungMin: 15,
  nachbereitungMin: 15,
  // Lücke bis zu dieser Länge: Kurse gelten als "am Stück".
  maxLueckeAmStueckMin: 30,
  // Pause zwischen zwei Kursen: erst ab so vielen Minuten fährt Michael heim
  // (und nur, wenn die Zeit für Hin- und Rückweg reicht).
  pauseHeimfahrtMin: 120,
  kategorieFahrt: "90 🚗 Fahrzeit",
  kategoriePuffer: "Vor-/Nachbereitung",
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
    tag: 15,
    teilnehmer: false,
    betreff: "Monatsauswertung {monat} – Michael Stehle ({anbieter})",
    text:
      "Hallo zusammen,\n\n" +
      "ich hoffe, bei euch ist alles gut! Im Anhang findet ihr meine Monatsauswertung für {monat}, " +
      "mit allen Kursen, die ich bei euch gegeben oder vertreten habe, und den Fahrten dazu.\n\n" +
      "Wenn etwas nicht passt oder ihr Fragen habt, meldet euch gern jederzeit bei mir.\n\n" +
      "Herzliche Grüße und bis bald\nMichael Stehle",
  },
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
