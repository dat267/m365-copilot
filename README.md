# m365-copilot

Minimal programmatic **M365 Copilot text generation** — no proxy, no HTTP server.
Just auth + one WebSocket chat client, so you can prompt M365 Copilot from Node.

M365 Copilot has no public API and no API key, so this impersonates Microsoft's
own Office-web Copilot client (the undocumented SignalR/WebSocket "Sydney"
endpoint). You need a Microsoft 365 account with Copilot — nothing else.

## Install

From npm (when published):

```sh
npm install -g m365-copilot      # CLI on your PATH as `m365-copilot`
# or, as a library:
npm install m365-copilot         # import { ask, M365Session } from "m365-copilot"
# or one-off, nothing installed:
npx m365-copilot ask "..."
```

### Install straight from GitHub

No npm release needed — install the latest `main` directly:

```sh
npm install -g github:dat267/m365-copilot      # CLI on your PATH
# or as a dependency in another project:
npm install github:dat267/m365-copilot
# or one-off, nothing installed:
npx github:dat267/m365-copilot ask "..."
```

Private repo? Use SSH instead:

```sh
npm install -g git+ssh://git@github.com/dat267/m365-copilot.git
```

Pin a version with `github:dat267/m365-copilot#<tag-or-sha>` (e.g. `#v0.1.0`).
Note npm caches git installs; to pick up new `main` commits, re-run the
install with `--force` (or pin a tag/SHA). Tokens are **not** part of the
install — run `m365-copilot auth` once per machine (see below).

### Windows (no admin required)

Everything here is user-scope: npm's global prefix is `%APPDATA%\npm`, Chromium
caches to `%LOCALAPPDATA%\ms-playwright`, tokens live in
`%USERPROFILE%\.config\m365-copilot` — nothing touches system paths.

```powershell
# In a plain PowerShell window (no elevation):

# 1. Node LTS + git via scoop (user-space package manager; git is needed for
#    GitHub installs). Set-ExecutionPolicy is user-scope and needs no admin.
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned -Force
irm get.scoop.sh | iex
scoop install nodejs-lts git

# 2. Put npm's global bin dir on your user PATH (scoop's node doesn't add it)
[Environment]::SetEnvironmentVariable(
  "Path",
  [Environment]::GetEnvironmentVariable("Path", "User") + ";$env:APPDATA\npm",
  "User")
# ...then open a NEW terminal so the PATH change applies

# 3. Install the CLI straight from GitHub
npm install -g github:dat267/m365-copilot

# 4. Auth — Playwright is optional, only interactive sign-in needs it
npm install -g playwright
npx playwright install chromium
m365-copilot auth
```

No Chromium at all? Skip step 4 and set a browser-copied refresh token instead
(user env var, no admin):

```powershell
setx M365_REFRESH_TOKEN "<paste the token>"
```

Corporate TLS (Zscaler/Netskope) — also just a user env var:

```powershell
setx NODE_EXTRA_CA_CERTS "%USERPROFILE%\certs\corp-root.pem"
```

Prefer not to use scoop: manually unzip Node's `win-x64` zip and MinGit into
`%LOCALAPPDATA%` and add both to your user PATH — same result, same privileges.

### Pin a tag

Tag releases in the repo (`git tag v0.1.0 && git push origin v0.1.0`) and
install that tag for a stable deploy:

```sh
npm install -g github:dat267/m365-copilot#v0.1.0
```

### Updating

```sh
npm update -g m365-copilot       # when installed from npm
npm install -g --force github:dat267/m365-copilot   # when installed from main
```

Developing in this repo instead:

```sh
npm install
npx playwright install chromium   # only for the one-time interactive sign-in
```

Playwright is an **optional peer dependency**: only the interactive sign-in
(`m365-copilot auth`, or a first run with no token anywhere) needs it — and the
Chromium browser itself is a separate `npx playwright install chromium`
download. Without it, everything still runs as long as a token is supplied via
`M365_ACCESS_TOKEN` / `M365_REFRESH_TOKEN` (see "No Playwright?").

## Use

The CLI has exactly two commands (`node cli.js` is also installed as
`m365-copilot` via the package `bin`):

```sh
# One-time (or when tokens expire): sign in via a Playwright-driven browser
node cli.js auth

# Ask: one-shot answer from a TEMPORARY chat (no memory, never saved)
node cli.js ask "Summarize the key risks in this contract: ..."

# Pick a model (M365 selects models by a server-validated `tone` string)
node cli.js ask -m claude "..."

# Prepend a file (or stdin with -) as data around the prompt
node cli.js ask --context notes.md "Summarize this in 3 bullets."
cat notes.txt | node cli.js ask -c - "Draft a customer reply."

# Steer the answer with a system prompt
node cli.js ask --system persona.md "Draft a status update."
printf 'Be terse.' | node cli.js ask -s - "Is this outage handled?"
```

**System prompt**: `ask` prepends a system prompt to the message (M365 has no
system role). Resolution order: `--system <file|->` if given, else
`SYSTEM.md` in the config dir (`~/.config/m365-copilot/`, override with
`M365_CONFIG_DIR`) when that file exists, else none. The ticket scripts have
their own built-in system prompt and ignore this setting.

`ask` is always a temporary chat — M365 keeps no memory of it, it never appears
in your chat history, and nothing is persisted locally. There is no REPL; this
CLI is one-shot. For a persistent multi-turn conversation use the library
(`M365Session`).

First run opens a **visible browser** — sign in to Microsoft 365 once. The token
cache + browser profile live in `~/.config/m365-copilot/`, so later runs are silent.

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

It's persisted to `~/.config/m365-copilot/token.json` (with the **rotated** token),
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

`M365Session` persists its `conversationId` (`~/.config/m365-copilot/session.json`)
and resumes it on later runs, so separate invocations reuse one M365
conversation instead of burning a new one. Start fresh with
`session.newConversation()` or `new M365Session({ fresh: true })`; set
`M365_NO_SESSION_PERSIST=1` to disable persistence entirely.

**Temporary chat** (`new M365Session({ temporary: true })`; the CLI's `ask`
command is always temporary) sends `disableMemory=1` on the chat socket: the
server keeps no memory of the conversation and it never appears in your chat
history. Such a session also never reads or writes `session.json`.

`stream` is async-iterable (yields text deltas) and exposes after completion:

| getter | meaning |
|---|---|
| `fullText` | the complete answer |
| `hasContent` | whether the server returned anything |
| `messageType` | `"Disengaged"` when M365's safety filter fired (empty answer) |
| `contentOrigin` | e.g. `"DeepLeo"` (which backend answered) |
| `throttle` | `{ current, max }` — the per-conversation 600-message quota |
| `scores` | M365's own classifier scores (`dea_violation` rises before Disengaged) |

### Conversations & files (REST)

The web client also exposes plain REST endpoints on the same
`https://substrate.office.com/m365Copilot/` host, using the same Sydney token.
Both an `M365Session` method and a standalone function are available:

```js
const { listConversations, getConversation, deleteConversation, uploadFile, uploadImages } =
  await import("m365-copilot");

const { chats } = await listConversations({ token });      // newest first, 100/page
await getConversation({ token, conversationId });          // full history
await deleteConversation({ token, conversationIds: [id] }); // batch is supported
const { docId } = await uploadFile({                       // -> attach to a turn
  token, conversationId, data: bytes, fileName: "a.png", mimeType: "image/png",
});
// one message's images: at most MAX_IMAGES_PER_MESSAGE (3), enforced
const [{ docId: first }] = await uploadImages({
  token, conversationId, images: [{ data: bytes, fileName: "a.png" }],
});
```

On a session they pick up auth automatically and default to the current chat:

```js
const s = new M365Session();
for (const c of (await s.listConversations()).chats) console.log(c.chatName);
s.selectConversation(conversationId);   // resume an existing chat
await s.deleteConversation();           // defaults to the current conversation
await s.uploadImages([imgA, imgB, imgC]); // batch upload for one message
```

`selectConversation()` marks the conversation as *resumed*, so the next turn
correctly sends `isStartOfSession:false`.

**Upload limits.** At most **3 images per message** — `uploadImages()` enforces it
(`MAX_IMAGES_PER_MESSAGE`, overridable per call via `limit` since the real cap is
server-side and feature-flagged; the 3 is the observed value, not a protocol
constant). File size caps at **512 MB**, and an image the safety filter rejects
fails with **HTTP 417**. Uploads are sent sequentially — M365 throttles on
concurrent requests.

`uploadFile()` returns a `docId`; pass upload results as `chat(text, { attachments })`
to attach them to a turn (uploading alone does nothing).

## Freshservice credentials

`scripts/ask-ticket.js` talks to the Freshservice private API directly. It does
**not** use the `fsvc` tool and does not read fsvc's config. Configure it with
either env vars:

```sh
export FRESHSERVICE_SUBDOMAIN=acme
export FRESHSERVICE_SESSION=<the _itildesk_session cookie value>
# optional: FRESHSERVICE_BASE_URL=https://acme.freshservice.com
```

or a JSON file (env wins over file):

```json
{ "subdomain": "acme", "session": "...", "baseUrl": "" }
```

Looked up at `$M365_FRESHSERVICE_CONFIG`, else `./freshservice.json`, else
`~/.config/m365-copilot/freshservice.json`.

## Ticket attachments

Freshservice tickets carry images and files, and a ticket can easily have more
than M365 allows in one message (`MAX_IMAGES_PER_MESSAGE`, observed as 3).
`ask-ticket.js` handles this by **not** trying to inline everything:

1. **The manifest always goes in the prompt.** `fetchTicket()` collects
   attachments from the ticket, every conversation, and the
   `linked-attachments` endpoint, and `renderTicket()` appends a table of
   name / type / size / origin / signed URL. The model thus knows the complete
   set even when none of it is sent as an image.
2. **Images are planned, never dropped.** `planImageBatches()` splits images
   into groups of at most 3 and separates out non-image files, so a caller can
   deliver them across turns instead of silently losing the extras.

```js
import { planImageBatches } from "m365-copilot";
const { images, files, batches } = planImageBatches(attachments);
// 7 screenshots -> batches of [3, 3, 1]
```

### Strategy for more than 3 images

The 3-image cap applies to **vision payloads**, not to context in general, so
scale the other channels first:

| Attachment | Recommended delivery |
|---|---|
| Images and files (any mix) | Upload via `uploadImages()` / `uploadFileToCopilot()` and attach with `chat(..., { attachments })`. The cap is **shared — 3 per message** — and the scripts split a larger set across turns. |
| Nothing deliverable | Named in the manifest with its URL, so it is never a silent gap. |

> **Images** attach via `message.messageAnnotations` on the chat invocation
> (`uploadFile()` returns the `docId` for `id`). Uploading alone does nothing.
> Verified: an image reading "42" was uploaded and attached, and the model
> answered `42`.
>
> **Files (documents) use a different channel entirely.** They are uploaded into
> the user's OneDrive *Microsoft Copilot Chat Files* folder via Graph, then
> attached as a `LocalFile` annotation:
>
> ```
> GET  /me/drive/special/copilotuploads                              -> driveId
> POST /me/drive/special/copilotuploads:/<name>:/createUploadSession
> PUT  <uploadUrl>   (Content-Range required)                        -> driveItem
> ```
> ```json
> "messageAnnotations": [{
>   "id": "SPO_<base64url(siteId,webId,listId)>_<itemId>",
>   "text": "probe-file.txt",
>   "url": "https://<tenant>.sharepoint.com/.../Microsoft%20Copilot%20Chat%20Files/probe-file.txt",
>   "messageAnnotationType": "LocalFile"
> }]
> ```
>
> The three GUIDs are the ones packed inside the `b!` drive id.
> `uploadFileToCopilot()` does the upload; `toFileAnnotations()`/`spoId()` build
> the annotation. This needs a **Graph token** (`https://graph.microsoft.com/.default`);
> the Sydney token will not work. Verified live end to end.

### Per-message attachment budget

**3 attachments per message — images and files share one cap.** Verified against
the web client by attaching mixtures:

| Composition | Accepted |
|---|---|
| 1 image + 1 file | 2 |
| 1 image + 2 files | 3 |
| 1 image + 3 files | **3** (4th refused) |
| 3 images + any file | **3** (file refused) |

So batch them **together**, never three images plus three files. `MAX_IMAGES_PER_MESSAGE`
is the shared cap, enforced by `planImageBatches()`/`planAttachmentBatches()`.

### Freshservice ticket script

`scripts/ask-ticket.js` fetches the ticket from Freshservice's private API and
feeds the full conversation trace to Copilot as data. The payload is
budget-truncated (head + tail) to stay under `--max-chars`, then loaded as the
FIRST turn of one conversation so `--follow` reuses its context. Ticket
attachments are not downloaded — their inventory (name, type, size, origin,
signed URL) is appended to the prompt; images are additionally uploaded and
attached in batches under the shared 3-per-message cap above.

The ticket text is sent as-is — no PII or network redaction. By default the
script steers the model with a built-in system prompt for a plain-text
IT-support **ticket digest**; override it with `M365_TICKET_SYSTEM_PROMPT`.

```sh
node scripts/ask-ticket.js 10100 "Draft a concise customer reply."
node scripts/ask-ticket.js --follow 10100 "Summarize the issue and action items."
node scripts/ask-ticket.js --new 10100 "What is blocking this ticket?"
```

The M365 conversation is persisted and reused across runs by default; `--new`
forces a fresh one (`M365_NO_SESSION_PERSIST=1` disables persistence entirely).

Credentials come from `FRESHSERVICE_SUBDOMAIN`, `FRESHSERVICE_SESSION`,
`FRESHSERVICE_BASE_URL`, or a `freshservice.json` — no fsvc.

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

**Node 24+ shortcut:** if your OS already trusts the corp CA (it does — that's
why the browser works), `--use-system-ca` makes Node use the OS trust store and
avoids exporting a PEM at all. This also fixes the *REST* calls, which
`M365_INSECURE=1` does not cover:

```sh
node --use-system-ca cli.js "hi"
```

Blunt fallback if you don't have the PEM (insecure — corp networks only):

```sh
M365_INSECURE=1 node cli.js "hi"
```

**Throttling.** M365 limits *conversations started* per unit time, and each
conversation is capped at ~600 messages. Conversations are reused across runs by
default, so `ask()` / `M365Session` do not start a new one per prompt — only
`{ fresh: true }`, `newConversation()`, or `new M365Session({ fresh: true })`
does. If turns start returning
empty, back off — it's account-level throttle, not a content problem.

**Disengaged.** Prompt-injection-shaped or very large tool-style prompts get a
`messageType:"Disengaged"` bot message with empty content — a refusal, not
throttling. `ask()` throws on it; rephrase or start a new conversation.

## Layout

```
cli.js          bin: thin launcher — parses argv via commander, maps errors to exit codes
src/cli.js      command wiring: `auth` (Playwright sign-in) and `ask` (temporary chat, --model)
src/auth.js     MSAL PKCE + interactive sign-in + token cache
src/client.js   SignalR/WebSocket chat client (one turn per WS)
src/chat-api.js history/deletion/upload REST (GetChats, GetConversation, DeleteConversation, UploadFile)
src/attachments.js attachment planning: isImage, planImageBatches, manifest rendering
src/graph-upload.js  document upload to OneDrive copilotuploads + LocalFile annotations
src/index.js    public API: ask() + M365Session
src/session-store.js  persisted default conversation (session.json)
src/prompt.js   prompt assembly: context wrapping + loading (no arg parsing — commander owns that)
src/ticket.js   Freshservice ticket fetch/render
src/log.js      optional debug logging (M365_DEBUG=1)
scripts/        ask-ticket.js (Freshservice ticket → Copilot, repo imports)
```

The ticket text is sent to Copilot as-is — no redaction.

This is a trimmed extraction of `m365-copilot-proxy` (same authors' reverse
engineering); it keeps only the plain-chat path. No tool-calling, agents, or
image generation.
