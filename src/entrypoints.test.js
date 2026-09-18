import test from "node:test";
import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

// Entry points are never imported by the test suite, so a syntax error there
// (e.g. an unescaped backtick inside a template literal in the help text) would
// otherwise ship unnoticed.
const root = join(dirname(fileURLToPath(import.meta.url)), "..");

for (const file of ["cli.js", "scripts/ask-ticket.js"]) {
  test(`${file} parses`, () => {
    execFileSync(process.execPath, ["--check", join(root, file)], { stdio: "pipe" });
  });
}
