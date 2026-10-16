// IDs aus der App-Registrierung in Microsoft Entra ("Kursplaner").
// Keine Geheimnisse: diese IDs dürfen öffentlich sein.
export const config = {
  clientId: "HIER-ANWENDUNGS-ID",
  tenantId: "HIER-VERZEICHNIS-ID",
  scopes: ["Calendars.ReadWrite", "Mail.ReadWrite"],
  zeitzone: "W. Europe Standard Time",
};
