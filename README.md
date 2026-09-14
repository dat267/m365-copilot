# m365-copilot

Minimal programmatic **M365 Copilot text generation** — no proxy, no HTTP server.
Just auth + one WebSocket chat client, so you can prompt M365 Copilot from Node.

M365 Copilot has no public API and no API key, so this impersonates Microsoft's
own Office-web Copilot client (the undocumented SignalR/WebSocket "Sydney"
endpoint). You need a Microsoft 365 account with Copilot — nothing else.

## Install

```sh
npm install
npx playwright install chromium   # only for the one-time interactive sign-in
```

## Use

```sh
# One-shot
node cli.js "Summarize the key risks in this contract: ..."

# Pick a model (M365 selects models by a `tone` string)
node cli.js --model=claude "..."

# Interactive REPL (streams, reuses one conversation)
node cli.js

# Conversations are persisted and reused across runs; --new starts a fresh one
node cli.js --new "..."
```

First run opens a **visible browser** — sign in to Microsoft 365 once. The token
cache + browser profile live in `~/.config/m365-ask/`, so later runs are silent.

### No Playwright? Use a copied token

Playwright is only used to **obtain** a token and persist a refresh token — the
client itself just needs a token string. You can copy one out of an existing
signed-in browser session instead and skip browser automation entirely.

**Access token (simplest, ~1 hour).**

1. Open <https://m365.cloud.microsoft> and sign in.
2. DevTools (F12) → **Network** → filter **WS**.
3. Click the `…/m365Copilot/Chathub/…` WebSocket → **Headers** → **Request URL**.
4. Copy the `access_token=` value (URL-decode it — `%xx` escapes).
5. ```sh
   M365_ACCESS_TOKEN=<token> node cli.js "hi"
   ```

This is the token the real web client sends, so it's definitely available there.
It expires in ~1h and can't be refreshed — fine for a session, not for a service.

**Refresh token (durable — no browser after this).**

The refresh token is the credential worth copying: the client trades it for
fresh access tokens indefinitely and rotates it automatically.

- DevTools → **Network**, filter `token`, find the `POST` to
  `login.microsoftonline.com/…/oauth2/v2.0/token`; its JSON response contains
  `refresh_token`. (It may be issued from an iframe — check that frame's requests.)
- Or **Application → Storage → Local/Session Storage** for a key containing
  `refreshtoken`.

```sh
M365_REFRESH_TOKEN=<token> node cli.js "hi"
```

It's persisted to `~/.config/m365-ask/token.json` (with the **rotated** token),
so later runs need no env var. AAD rotates refresh tokens on every exchange —
keep one store, or you'll hold a stale one.

If the grant fails with a client mismatch, the browser token belongs to a
different first-party client: set `M365_CLIENT_ID` to the `client_id` from that
token request (also the `appid` claim of the access token).

Playwright remains the fallback: if the refresh token expires or is revoked,
`getToken()` opens the browser once more.

### As a library

```js
import { ask, M365Session } from "m365-copilot";

// one-shot
const text = await ask("Write a regex that matches ISO dates.");

// multi-turn — one conversation, so M365 keeps context and you don't burn a
// fresh conversation against the account's thread-rate limit
const session = new M365Session({ model: "m365-copilot" });
const stream = await session.chat("My name is Ada.");
for await (const delta of stream) process.stdout.write(delta);
console.log(await ask("What is my name?", { session })); // -> Ada
```

`M365Session` persists its `conversationId` (`~/.config/m365-ask/session.json`)
and resumes it on later runs, so separate invocations reuse one M365
conversation instead of burning a new one. Start fresh with
`session.newConversation()` or `new M365Session({ fresh: true })`; set
`M365_NO_SESSION_PERSIST=1` to disable persistence entirely.

`stream` is async-iterable (yields text deltas) and exposes after completion:

| getter | meaning |
|---|---|
| `fullText` | the complete answer |
| `hasContent` | whether the server returned anything |
| `messageType` | `"Disengaged"` when M365's safety filter fired (empty answer) |
| `contentOrigin` | e.g. `"DeepLeo"` (which backend answered) |
| `throttle` | `{ current, max }` — the per-conversation 600-message quota |
| `scores` | M365's own classifier scores (`dea_violation` rises before Disengaged) |

## Models

Selected by a server-validated `tone` string, not a model id. Run with `--help`
for the current list. Known-good: `m365-copilot`/`auto` (GPT-5 class),
`claude` (`Claude_Sonnet`, real Anthropic Claude, no agent needed),
`gpt-5.5`, `gpt-5.6-think-deeper`. Reasoning/`think-deeper` tones are slower.
The server rejects unknown tones, so if one stops working it's been retired.

## Work-network gotchas

**Corporate TLS inspection** (Zscaler/Netskope/etc.) makes the WebSocket fail
with `self-signed certificate in certificate chain`. Fix by trusting your corp
root CA:

```sh
NODE_EXTRA_CA_CERTS=/path/to/corp-root.pem node cli.js "hi"
```

Blunt fallback if you don't have the PEM (insecure — corp networks only):

```sh
M365_INSECURE=1 node cli.js "hi"
```

**Throttling.** M365 limits *conversations started* per unit time, and each
conversation is capped at ~600 messages. Conversations are reused across runs by
default, so `ask()` / `M365Session` do not start a new one per prompt — only
`{ fresh: true }`, `newConversation()`, or `--new` does. If turns start returning
empty, back off — it's account-level throttle, not a content problem.

**Disengaged.** Prompt-injection-shaped or very large tool-style prompts get a
`messageType:"Disengaged"` bot message with empty content — a refusal, not
throttling. `ask()` throws on it; rephrase or start a new conversation.

## Layout

```
cli.js          CLI (one-shot + REPL)
src/auth.js     MSAL PKCE + interactive sign-in + token cache
src/client.js   SignalR/WebSocket chat client (one turn per WS)
src/index.js    public API: ask() + M365Session
src/session-store.js  persisted default conversation (session.json)
src/prompt.js   CLI context/prompt assembly (--context, --new)
src/ticket.js   Freshservice ticket fetch/render + PII redaction
src/log.js      optional debug logging (M365_DEBUG=1)
scripts/        ask-ticket.js (repo imports) + ask-ticket-standalone.mjs (self-contained)
```

The ticket scripts redact email addresses and phone numbers from the ticket
before sending it to Copilot.

This is a trimmed extraction of `m365-copilot-proxy` (same authors' reverse
engineering); it keeps only the plain-chat path. No tool-calling, agents, or
image generation.
