#!/usr/bin/env node
// Ask M365 Copilot about a Freshservice ticket — fetches the ticket from the
// Freshservice private API itself (no fsvc subprocess), then feeds the full
// conversation trace to Copilot as data.
//
//   node scripts/ask-ticket.js 10100 "Draft a concise customer reply."
//   node scripts/ask-ticket.js --follow 10100 "Summarize the issue and action items."
//   node scripts/ask-ticket.js --max-chars=40000 10100 "What is blocking this ticket?"
//   node scripts/ask-ticket.js --new 10100 "Start a fresh conversation for this ticket."
//
// The M365 conversation is persisted and reused across runs by default; --new
// forces a fresh one (M365_NO_SESSION_PERSIST=1 disables persistence).
//
// Credentials/config are this repo's own — NOT fsvc's:
//   env:  FRESHSERVICE_SUBDOMAIN, FRESHSERVICE_SESSION, FRESHSERVICE_BASE_URL
//   file: $M365_FRESHSERVICE_CONFIG, else ./freshservice.json,
//         else ~/.config/m365-ask/freshservice.json
//   keys: { "subdomain": "acme", "session": "...", "baseUrl": "" }
//
// The session value is the browser's `_itildesk_session` cookie.
//
// The payload is budget-truncated (head + tail) to stay under --max-chars, then
// loaded as the FIRST turn of one conversation so --follow reuses its context.
//
// Ticket attachments are NOT downloaded. Their inventory (name, type, size,
// origin, signed URL) is appended to the prompt, so the model always sees the
// complete set. Images are additionally planned into batches of at most
// MAX_IMAGES_PER_MESSAGE — see README "Ticket attachments" for why and for the
// delivery strategy when a ticket has more than that.

import { existsSync, readFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { createInterface } from "node:readline/promises";
import { M365Session, MAX_IMAGES_PER_MESSAGE } from "../src/index.js";
import { buildPrompt } from "../src/prompt.js";
import { parseTicketArgs, resolveConfig, fetchTicket, renderTicket, truncateContext, makeFreshserviceGet, redactPII } from "../src/ticket.js";
import { planImageBatches, planAttachmentBatches } from "../src/attachments.js";

const USAGE = 'usage: ask-ticket.js [--max-chars=N] [--follow] [--new] <ticket-id> "instruction"';

const { id, instruction, maxChars, follow, fresh } = parseTicketArgs(process.argv.slice(2));
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
  defaultPath: join(configHome, "m365-ask", "freshservice.json"),
  exists: existsSync,
  readFile: (p) => readFileSync(p, "utf8"),
});

// The private API, session-cookie auth.
const get = makeFreshserviceGet({ baseUrl, session: freshserviceSession });

const { ticket, conversations, attachments } = await fetchTicket(id, get);
const raw = redactPII(renderTicket(ticket, conversations, { attachments }));
const context = truncateContext(raw, maxChars);
if (context.length < raw.length) {
  console.error(`[ask-ticket] payload truncated ${raw.length} → ${context.length} chars (--max-chars=${maxChars})`);
}

// Report how the attachments would be delivered. The attachment inventory is
// already in the prompt as text; this is about the payloads, which M365 caps at
// 3 per message — images and files SHARING that cap. See README
// "Ticket attachments". This script (unlike the PowerShell one) does not upload
// attachments, so the plan is informational.
if (attachments.length > 0) {
  const split = planImageBatches(attachments, { maxPerMessage: MAX_IMAGES_PER_MESSAGE });
  const batches = planAttachmentBatches(attachments, { maxPerMessage: MAX_IMAGES_PER_MESSAGE });
  const parts = [
    `${attachments.length} attachment(s): ${split.images.length} image(s), ${split.files.length} file(s)`,
  ];
  if (batches.length > 0) {
    parts.push(`${batches.length} attachment batch(es) at ≤${MAX_IMAGES_PER_MESSAGE}/message (images and files share the cap)`);
  }
  console.error(`[ask-ticket] ${parts.join("; ")}`);
  if (plan.files.length > 0) {
    console.error(
      `[ask-ticket] non-image files are listed in the prompt manifest only (no text extraction yet): ` +
        plan.files.map((f) => f.name).join(", "),
    );
  }
}

const copilot = new M365Session({ fresh });

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
