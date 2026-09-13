#!/usr/bin/env node
// CLI for m365-ask.
//
//   node cli.js "prompt"                    # one-shot
//   node cli.js --model=claude "prompt"     # pick a tone/model
//   node cli.js                             # interactive REPL (one conversation)

import { createInterface } from "node:readline/promises";
import { M365Session, ask, getAvailableModels } from "./src/index.js";

const argv = process.argv.slice(2);

if (argv.includes("--help") || argv.includes("-h")) {
  console.log(`Usage: m365-ask [--model=<id>] "prompt"

Models: ${getAvailableModels().join(", ")}

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

let model = "m365-copilot";
const rest = [];
for (const a of argv) {
  if (a.startsWith("--model=")) model = a.slice("--model=".length);
  else rest.push(a);
}

if (rest.length > 0) {
  const text = await ask(rest.join(" "), { model });
  process.stdout.write(text + "\n");
} else {
  const rl = createInterface({ input: process.stdin, output: process.stdout });
  const session = new M365Session({ model });
  console.error(`M365 Copilot [model=${model}]. Type a prompt; Ctrl-D to exit.`);
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
