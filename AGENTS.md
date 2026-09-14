# AGENTS.md

Guidance for AI agents (and humans) working in this repo.

## What this is

`m365-copilot` is a **minimal, standalone** Node library + CLI for programmatically
prompting **Microsoft 365 Copilot** for text generation. No proxy, no HTTP
server — just auth plus one SignalR/WebSocket chat client.

M365 Copilot has **no public API and no API key**, so this impersonates
Microsoft's own Office-web Copilot client against the undocumented "Sydney"
endpoint (`wss://substrate.office.com/m365Copilot/Chathub/...`). The user only
needs an M365 account with Copilot.

This is a **trimmed extraction** of the sibling repo `../m365-copilot-proxy`,
which did the original reverse engineering. That repo's
[`docs/m365-copilot-api.md`](../m365-copilot-proxy/docs/m365-copilot-api.md)
is the **source of truth for the protocol** — read it before changing anything
in `src/client.js` or `src/auth.js`. This project intentionally keeps only the
**plain-chat** path: no tool-calling, no Copilot Studio agents, no image
generation, no automated password/TOTP login.

## Operating principles (read first)

1. **Every M365 turn is quota and risk.** Account-level throttling tracks
   *conversations started* per unit time, and each conversation is capped at
   ~600 messages. Never fire concurrent requests. One `M365Session`
   (one `conversationId`) is persisted in `session.json` and reused across runs
   — a fresh conversation per prompt is what burns the thread budget; start one
   only explicitly (`{ fresh: true }`, `newConversation()`, `--new`). Space out
   test runs.

2. **An empty reply is usually NOT a bug and NOT always throttling.**
   - `messageType: "Disengaged"` → the safety filter refused (empty content).
     Rephrase or start a new conversation; retrying the same prompt just
     re-disengages and burns quota.
   - empty with `throttle` at max → quota.
   - empty with no throttle, returning fast → account degradation; **back off
     and wait**, don't loop.
   `ask()` throws on both Disengaged and empty so callers don't silently get `""`.

3. **Auth is the hard part; don't "simplify" it.** The `nativeclient` redirect
   is meant for embedded native hosts — a real browser follows it one hop
   further to `/common/wrongplace`, so the `?code=` exists only transiently.
   We capture it from the **navigation request**, not a settled URL. The token
   goes in the **WebSocket URL query string**, not a header. Node's native
   `WebSocket` does not work — it must be `ws` with a browser `Origin`/UA.

4. **`tone` selects the model and is server-validated.** Unknown tones error
   with `Failed to invoke 'Chat'`. Microsoft retires tones freely — e.g.
   `Gpt_Quick` (`"quick"`) worked historically but was rejected when last
   tested, so it is deliberately not offered. Only add tones confirmed live.

5. **Be scientific and minimal.** This is a reverse-engineering-derived client.
   Before adding a field/frame, confirm it against a live capture or the proxy
   docs. Prefer deleting code to adding it.

## Layout (plain ESM JavaScript — no build step)

| File | Role |
|---|---|
| `cli.js` | CLI: one-shot (`node cli.js "prompt"`) and interactive REPL |
| `src/auth.js` | MSAL PKCE, silent refresh, interactive sign-in, token cache, raw refresh-token grant, `decodeJwt` |
| `src/client.js` | `CopilotSession` — one WS turn (handshake, `Metrics` frame, frame dispatch, delta folding), and the `tone` map |
| `src/session-store.js` | persisted default conversation (`session.json`): id/turn-count resolution, load/save |
| `src/prompt.js` | CLI context/prompt assembly (`--context`, `--new`), no I/O |
| `src/ticket.js` | Freshservice ticket fetch/render + `redactPII` (used by `scripts/ask-ticket.js`) |
| `src/index.js` | Public API: `ask()` (one-shot) and `M365Session` (multi-turn, handles auth + reconnect) |
| `src/log.js` | Optional debug logging (`M365_DEBUG=1` → `~/.config/m365-ask/debug.log`) |
| `examples/` | Runnable examples |
| `scripts/` | `ask-ticket.js` (repo imports) and `ask-ticket-standalone.mjs` (Node built-ins only) |

ESM, `.js`-suffixed relative imports. No TypeScript, no bundler.

## Build & test

No build. Dependencies are `@azure/msal-node`, `playwright`, `ws`.

```sh
npm install
npx playwright install chromium   # once, only for the interactive sign-in browser
node cli.js --help
```

There is **no unit-test suite** — the protocol is only meaningfully testable
against the live API. Verify changes end-to-end (below).

## Running against real M365

Config/cache live in **`~/.config/m365-ask/`** (`msal-cache.json`,
`browser-profile/`, `session.json`) — deliberately separate from the proxy's
`~/.config/opencode-m365/`. Override with `M365_CONFIG_DIR`,
`M365_CACHE_FILE`, `M365_BROWSER_PROFILE`.

- **Playwright is optional.** It exists only to obtain a token. `getToken()`
  checks, in order: `M365_ACCESS_TOKEN` env → a still-valid `token.json` access
  token → a refresh-token grant (`M365_REFRESH_TOKEN` env or `token.json`) →
  MSAL silent refresh → interactive browser. So you can run this entirely on a
  browser-copied token and never install Chromium. See README "No Playwright?".
- **Refresh tokens ROTATE.** AAD returns a new one on every grant; `auth.js`
  persists it to `~/.config/m365-ask/token.json`. Never share one refresh token
  across two stores/processes — the stale copy stops working. (The old token
  keeps a short grace period, so a single extra use won't immediately break.)
- **Interactive auth** (when reached) opens a *visible* browser; sign in once.
  The persistent profile makes later runs SSO-silent. There is no
  password/TOTP-automation path here by design.
- **To test without a browser**, either set `M365_REFRESH_TOKEN` from the proxy's
  cache, or copy `~/.config/opencode-m365/msal-cache.json` to
  `~/.config/m365-ask/msal-cache.json` (MSAL silent path).

Verify end-to-end:

```sh
node cli.js "Reply with exactly the word: READY"     # expect: READY
node examples/multiturn.js                           # turn 2 must recall turn 1
```

## Gotchas to know before you "fix" something

- **Corporate TLS inspection** (Zscaler/Netskope/…) makes the WS upgrade fail
  with `self-signed certificate in certificate chain`. Prefer
  `NODE_EXTRA_CA_CERTS=/path/corp-root.pem`; `M365_INSECURE=1` disables
  verification on the WS only (blunt, corp-networks-only). **This is
  environmental, not a code bug** — don't "fix" it in the client.
- **Reasoning tones** (`*_Reasoning`, `think-deeper`) route through the
  `DeepLeo` pipeline and are slow. The default `magic` is the safe choice.
- **Streaming mixes deltas and full-text snapshots**, and the first token often
  arrives ONLY as a snapshot — so we fold both via `foldStreamText` and only
  ever emit a true prefix of the final answer. Don't "simplify" to plain delta
  concatenation; that drops the head of the response.
- **Code interpreter is enabled by default** on the agent-less path
  (`cwc_code_interpreter*` optionsSets) — M365 then writes and runs real Python
  server-side. Disable with `M365_NO_CODE_INTERPRETER=1`.
- **`VARIANTS`** (the WS-query feature-flag list) is cargo-culted from a
  captured real session; removing flags is untested. Keep the proven list.
- The auth still has the persistent-profile / anti-fingerprint defaults from the
  proxy. A wrong locale is cosmetic; changing a fingerprint that currently
  passes AAD's bot scoring is not worth the risk.

## Conventions

- Conventional Commits (`fix:`, `feat:`, `docs:`, `chore:`). No `Co-Authored-By`.
- Small focused files; handle errors explicitly.
- If you change protocol behaviour, update the proxy repo's
  `docs/m365-copilot-api.md` too — that's the canonical reference.
