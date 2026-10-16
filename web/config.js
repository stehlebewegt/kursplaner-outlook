// IDs aus der App-Registrierung in Microsoft Entra ("Kursplaner").
// Keine Geheimnisse: diese IDs dürfen öffentlich sein.
export const config = {
  clientId: "41edec3a-0427-47db-80f2-f681342b02a9",
  tenantId: "1599c46c-7023-494d-92fa-ec2792bf681d",
  scopes: ["Calendars.ReadWrite", "Mail.ReadWrite"],
  zeitzone: "W. Europe Standard Time",
};
