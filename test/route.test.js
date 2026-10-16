import { test } from "node:test";
import assert from "node:assert/strict";
import { adressVarianten, aufViertelstunde, hatAdresse } from "../automatik/route.js";

test("Adressvarianten lassen vorne Namen weg", () => {
  assert.deepEqual(adressVarianten("Herr Kai Sallie (Geschäftlich), Bahnhofstr. 1, 88682 Salem, Deutschland"), [
    "Herr Kai Sallie (Geschäftlich), Bahnhofstr. 1, 88682 Salem, Deutschland",
    "Bahnhofstr. 1, 88682 Salem, Deutschland",
    "88682 Salem, Deutschland",
  ]);
});

test("Aufrunden auf die Viertelstunde, mindestens 15", () => {
  assert.equal(aufViertelstunde(4), 15);
  assert.equal(aufViertelstunde(15), 15);
  assert.equal(aufViertelstunde(16.2), 30);
});

test("Online-Termine haben keine Adresse", () => {
  assert.ok(hatAdresse("Podologie Bernhardt, Am Schlosssee 5, 88682 Salem"));
  assert.ok(!hatAdresse("Microsoft Teams-Besprechung"));
  assert.ok(!hatAdresse(""));
});
