// Nötig, damit Edge/Chrome den Kursplaner als App installieren. Kein
// Zwischenspeicher: es wird immer der aktuelle Stand geladen.
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", (e) => e.waitUntil(self.clients.claim()));
self.addEventListener("fetch", () => {});
