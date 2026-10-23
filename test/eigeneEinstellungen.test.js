import { test } from "node:test";
import assert from "node:assert/strict";
import { mitEigenen, leseJson, holeEigene, speichereEigene } from "../src/eigeneEinstellungen.js";
import { einstellungen } from "../src/einstellungen.js";

test("Eigene Werte überschreiben nur erlaubte Felder", () => {
  const e = mitEigenen(einstellungen, { vorbereitungMin: 20, heimadresse: "x", material: ["Ball"], monatsbericht: { tag: 10 }, vorbereitung: "kaputt" });
  assert.equal(e.vorbereitungMin, 20);
  assert.equal(e.heimadresse, einstellungen.heimadresse);
  assert.deepEqual(e.material, ["Ball"]);
  assert.equal(e.monatsbericht.tag, 10);
  assert.equal(e.monatsbericht.betreff, einstellungen.monatsbericht.betreff);
  assert.equal(mitEigenen(einstellungen, null).anbieter, einstellungen.anbieter);
});

test("JSON auch aus HTML-Text lesen", () => {
  assert.deepEqual(leseJson('<html><body><p>{&quot;a&quot;:1}</p></body></html>'), { a: 1 });
  assert.equal(leseJson("kein json"), null);
});

test("Speichern legt versteckten Ordner und Nachricht an, Lesen findet sie", async () => {
  const ordner = [];
  const nachrichten = [];
  const graph = async (m, pfad, body) => {
    if (m === "GET" && pfad.includes("/mailFolders?")) return { value: ordner };
    if (m === "POST" && pfad.endsWith("/mailFolders")) return ordner[ordner.push({ id: "o1", ...body }) - 1];
    if (m === "GET" && pfad.includes("/messages?")) return { value: nachrichten };
    if (m === "POST") return nachrichten.push({ id: "m1", ...body });
    if (m === "PATCH") return Object.assign(nachrichten[0], body);
  };
  await speichereEigene(graph, { vorbereitungMin: 20 }, "/me");
  assert.equal(ordner[0].isHidden, true);
  await speichereEigene(graph, { vorbereitungMin: 25 }, "/me");
  assert.equal(nachrichten.length, 1);
  assert.deepEqual(await holeEigene(graph, "/me"), { vorbereitungMin: 25 });
});
