#!/usr/bin/env node
// Ask M365 Copilot about a Freshservice ticket — fetches the ticket from the
// Freshservice private API itself (no fsvc subprocess), then feeds the full
// conversation trace to Copilot as data.
//
//   node scripts/ask-ticket.js 10100 "Draft a concise customer reply."
//   node scripts/ask-ticket.js --follow 10100 "Summarize the issue and action items."
//   node scripts/ask-ticket.js --max-chars=40000 10100 "What is blocking this ticket?"
//
// Credentials/config are resolved exactly like fsvc (env vars first, then
// FSVC_CONFIG_FILE → ./fsvc.json → ~/.config/fsvc/fsvc.json):
//   subdomain / FSVC_SUBDOMAIN, itildesk-session / FSVC_ITILDESK_SESSION,
//   base-url / FSVC_BASE_URL.
//
// The payload is budget-truncated (head + tail) to stay under --max-chars, then
// loaded as the FIRST turn of one conversation so --follow reuses its context.

import { existsSync, readFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { createInterface } from "node:readline/promises";
import { M365Session } from "../src/index.js";
import { buildPrompt } from "../src/prompt.js";
import { parseTicketArgs, resolveConfig, fetchTicket, renderTicket, truncateContext, makeFreshserviceGet } from "../src/ticket.js";

const USAGE = 'usage: ask-ticket.js [--max-chars=N] [--follow] <ticket-id> "instruction"';

const { id, instruction, maxChars, follow } = parseTicketArgs(process.argv.slice(2));
if (!Number.isInteger(id) || id <= 0) {
  console.error(USAGE);
  process.exit(2);
}
if (!instruction) {
  console.error(`error: provide an instruction\n${USAGE}`);
  process.exit(2);
}

const configHome = process.env.XDG_CONFIG_HOME || join(homedir(), ".config");
const { baseUrl, session: freshserviceSession } = resolveConfig(process.env, {
  defaultPath: join(configHome, "fsvc", "fsvc.json"),
  exists: existsSync,
  readFile: (p) => readFileSync(p, "utf8"),
});

// The private API, session-cookie auth.
const get = makeFreshserviceGet({ baseUrl, session: freshserviceSession });

const { ticket, conversations } = await fetchTicket(id, get);
const raw = renderTicket(ticket, conversations);
const context = truncateContext(raw, maxChars);
if (context.length < raw.length) {
  console.error(`[ask-ticket] payload truncated ${raw.length} → ${context.length} chars (--max-chars=${maxChars})`);
}

const copilot = new M365Session();

async function turn(message, label) {
  const stream = await copilot.chat(message, { signal: AbortSignal.timeout(300_000) });
  for await (const delta of stream) process.stdout.write(delta);
  process.stdout.write("\n");
  const throttle = stream.throttle ? `${stream.throttle.current}/${stream.throttle.max}` : "?";
  console.error(`[ask-ticket] ${label} origin=${stream.contentOrigin ?? "?"} type=${stream.messageType ?? "Chat"} throttle=${throttle}`);
  return stream;
}

await turn(buildPrompt(context, instruction), "answer");

if (follow) {
  const rl = createInterface({ input: process.stdin, output: process.stdout });
  while (true) {
    let line;
    try {
      line = await rl.question("\n> ");
    } catch {
      break;
    }
    if (!line.trim()) continue;
    try {
      await turn(line, "follow-up");
    } catch (err) {
      console.error(`error: ${err.message}`);
    }
  }
  rl.close();
}
