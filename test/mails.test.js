import { test } from "node:test";
import assert from "node:assert/strict";
import { absenderRegeln, anbieterVonAbsender, unbeantwortet } from "../src/mails.js";

const regeln = absenderRegeln({ SVE: ["Info@Verein-Beispiel.de"], PTM: ["praxis.chefin@gmail.com"] });

test("Absender: Domain bei Firmen-Adressen, genaue Adresse bei Freemail", () => {
  assert.equal(anbieterVonAbsender("kollegin@verein-beispiel.de", regeln), "SVE");
  assert.equal(anbieterVonAbsender("x@mail.verein-beispiel.de", regeln), "SVE");
  assert.equal(anbieterVonAbsender("praxis.chefin@gmail.com", regeln), "PTM");
  assert.equal(anbieterVonAbsender("jemand@gmail.com", regeln), null);
});

test("Unbeantwortet: 2 bis 14 Tage alt, ohne Antwort", () => {
  const m = (id, absender, empfangen, letzteAktion) => ({ id, absender, empfangen, letzteAktion });
  const mails = [
    m("a", "kollegin@verein-beispiel.de", "2026-10-15T08:00:00Z"),
    m("b", "kollegin@verein-beispiel.de", "2026-10-15T08:00:00Z", 102),
    m("c", "kollegin@verein-beispiel.de", "2026-10-18T20:00:00Z"),
    m("d", "fremd@example.org", "2026-10-15T08:00:00Z"),
    m("e", "kollegin@verein-beispiel.de", "2026-09-01T08:00:00Z"),
  ];
  assert.deepEqual(unbeantwortet(mails, regeln, "2026-10-19T06:00:00Z").map((x) => x.id), ["a"]);
});
