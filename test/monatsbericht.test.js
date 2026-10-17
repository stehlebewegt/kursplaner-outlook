import { test } from "node:test";
import assert from "node:assert/strict";
import { anschreiben, vormonat } from "../automatik/monatsbericht.js";

const mb = { text: "du", textGruppe: "ihr" };
const k = (givenName, address) => ({ givenName, emailAddresses: [{ address }] });

test("Anrede aus Kontakten", () => {
  assert.deepEqual(anschreiben([], mb), { an: [], anrede: "zusammen", vorlage: "ihr" });
  assert.deepEqual(anschreiben([k("Anna", "a@x.de")], mb), { an: ["a@x.de"], anrede: "Anna", vorlage: "du" });
  assert.deepEqual(anschreiben([k("Anna", "a@x.de"), k("Ben", "b@x.de")], mb).anrede, "Anna und Ben");
  assert.deepEqual(anschreiben([k("Anna", "a@x.de"), k("Ben", "b@x.de"), k("Cem", "c@x.de")], mb).anrede, "Anna, Ben und Cem");
});

test("Vormonat", () => {
  assert.equal(vormonat("2026-11-15"), "2026-10");
  assert.equal(vormonat("2027-01-15"), "2026-12");
});
