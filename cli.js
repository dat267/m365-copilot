#!/usr/bin/env node
// Bin entry for the m365-copilot CLI module.
//
//   m365-copilot auth              # sign in via a Playwright-driven browser
//   m365-copilot ask "prompt"      # one-shot answer from a temporary chat
//   m365-copilot ask -m claude …   # pick a model (m365-copilot ask --help)
//
// All command wiring lives in src/cli.js (testable without spawning a
// process); this file only instantiates it with the real collaborators and
// maps failures to exit codes. No interactive REPL — this CLI is one-shot.

import { CommanderError } from "commander";
import { buildProgram } from "./src/cli.js";

const program = buildProgram();

try {
  await program.parseAsync();
} catch (err) {
  if (err instanceof CommanderError) {
    // --help / invalid option: commander already printed the message.
    process.exit(err.exitCode);
  }
  console.error(`error: ${err.message}`);
  process.exit(1);
}
