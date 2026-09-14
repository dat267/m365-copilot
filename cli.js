#!/usr/bin/env node
// CLI for m365-copilot.
//
//   node cli.js "prompt"                    # one-shot
//   node cli.js --model=claude "prompt"     # pick a tone/model
//   node cli.js                             # interactive REPL (one conversation)
//
// Feed a large payload (e.g. a ticket export) alongside the instruction:
//
//   fsvc tickets show 10100 | node cli.js --context - "Draft a customer reply."
//   node cli.js --context ticket.md "Summarize this in 3 bullets."

import { readFile } from "node:fs/promises";
import { createInterface } from "node:readline/promises";
import { M365Session, ask, getAvailableModels } from "./src/index.js";
import { parseArgs, loadContext, buildPrompt } from "./src/prompt.js";

const { help, model, context: contextSpec, prompt, fresh } = parseArgs(process.argv.slice(2));

if (help) {
  console.log(`Usage: m365-copilot [--model=<id>] [--context <file|->] [--new] "prompt"

Models: ${getAvailableModels().join(", ")}

Options:
  --context <file|->    prepend a file (or stdin with -) to the prompt as data
                        (use with a data producer, e.g. \`fsvc tickets show 10100 | cli --context -\`)
  --new                 start a new conversation instead of resuming the saved one
                        (conversations are reused across runs by default; opt out
                        entirely with M365_NO_SESSION_PERSIST=1)

With no prompt, starts an interactive REPL (single conversation, streamed).
Env:
  M365_ACCESS_TOKEN     use a token copied from the browser (no browser needed, ~1h)
  M365_REFRESH_TOKEN    durable browser-copied token; persisted + auto-rotated
  M365_CLIENT_ID        override the first-party client id for the grant
  M365_DEBUG=1          log to ~/.config/m365-ask/debug.log
  M365_INSECURE=1       skip TLS verification (corporate MITM proxy)
  NODE_EXTRA_CA_CERTS   path to your corporate root CA (preferred)
  M365_NO_CODE_INTERPRETER=1  disable M365's server-side Python sandbox`);
  process.exit(0);
}

async function readStdin() {
  let data = "";
  process.stdin.setEncoding("utf8");
  for await (const chunk of process.stdin) data += chunk;
  return data;
}

const context = await loadContext(contextSpec, { readFile, readStdin });

if (prompt) {
  const text = await ask(buildPrompt(context, prompt), { model, fresh });
  process.stdout.write(text + "\n");
} else {
  const rl = createInterface({ input: process.stdin, output: process.stdout });
  const session = new M365Session({ model, fresh });
  console.error(`M365 Copilot [model=${model}]. Type a prompt; Ctrl-D to exit.`);
  if (context) {
    // Load the payload as the first turn so follow-ups share its context.
    const stream = await session.chat(buildPrompt(context, ""), { signal: AbortSignal.timeout(300_000) });
    for await (const delta of stream) process.stdout.write(delta);
    process.stdout.write("\n");
  }
  while (true) {
    let line;
    try {
      line = await rl.question("\n> ");
    } catch {
      break;
    }
    if (!line.trim()) continue;
    try {
      const stream = await session.chat(line, { signal: AbortSignal.timeout(300_000) });
      for await (const delta of stream) process.stdout.write(delta);
      process.stdout.write("\n");
    } catch (err) {
      console.error(`error: ${err.message}`);
    }
  }
  rl.close();
}
