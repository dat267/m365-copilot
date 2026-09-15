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

# Temporary chat: no server-side memory, and it does not appear in your history
node cli.js --temporary "..."
```

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

**Temporary chat** (`new M365Session({ temporary: true })`, or `--temporary`)
sends `disableMemory=1` on the chat socket: the server keeps no memory of the
conversation and it never appears in your chat history. Such a session also
never reads or writes `session.json`.

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
is the shared cap, enforced by `planImageBatches()`/`planAttachmentBatches()`. The
standalone PowerShell script does not upload images or files — see below.

### Standalone PowerShell, and long tickets

`scripts/ask-ticket-standalone.ps1` is the self-contained pwsh 7+ version. It uses
the **OS certificate store**, so corporate TLS-inspection roots work without
`NODE_EXTRA_CA_CERTS`.

It does **not truncate** a long ticket. It renders the ticket as ordered sections
(header, attachment manifest, each conversation, each inlined attachment), packs
them into as many messages as the per-message budget allows, and sends them as
successive turns of **one** M365 conversation, so the model holds the whole
ticket when it answers. Intermediate turns ask for `ACK` to keep replies cheap.

Attachments are **text only** here: every attachment is listed in the manifest, and
text-bearing ones (logs, csv, json, …) are inlined verbatim. Images and other
binaries are named but never uploaded (the Node API still supports uploads).

Redaction covers emails, IPs, MAC addresses, phone numbers and the usual Windows
network-config identifiers (host name, DNS suffixes, DHCPv6 IDs). If your org
blocks more, add regexes to `$Config.ExtraRedact`, e.g.
`@{ Pattern = '\bACME-[A-Z0-9]+\b'; Replacement = '[redacted-id]' }`.

```sh
# ask about a ticket (long ones split automatically); temporary chat by default
pwsh -File scripts/ask-ticket-standalone.ps1 "Draft a concise customer reply." 24613

# show the message split without spending any Copilot turns
pwsh -File scripts/ask-ticket-standalone.ps1 -Plan "x" 24613

# conversation control
pwsh -File scripts/ask-ticket-standalone.ps1 -ConversationId <guid> "..." 24613
pwsh -File scripts/ask-ticket-standalone.ps1 -LastConversation "..." 24613
pwsh -File scripts/ask-ticket-standalone.ps1 -Persist "..." 24613   # memory ON
```

**Conversation mode.** Every chat is **temporary by default** — the ChatHub URL
carries `disableMemory=1`, so M365 keeps no long-term memory of it and it never
appears in your history. `-Persist` turns memory back on. `-ConversationId`
targets a specific conversation and `-LastConversation` reuses the id the script
remembered from the previous run (`~/.config/m365-copilot/last-ticket-conversation.json`).
Every run records the id it used, so `-LastConversation` works next time; note
that resuming is only meaningful for a conversation that was **not** temporary.

Limits that drive the split (all overridable via `$Config` or env):

| Knob | Env | Default | Limit it enforces |
|---|---|---|---|
| `MaxTextChars` | `M365_TICKET_MAX_TEXT_CHARS` | 60000 | message text limit (per turn) |
| `MaxMessages` | `M365_TICKET_MAX_MESSAGES` | 60 | safety cap on turns |
| `MaxInlineFiles` | `M365_TICKET_MAX_INLINE_FILES` | 20 | attachment count limit |
| `MaxInlineFileBytes` | `M365_TICKET_MAX_INLINE_FILE_BYTES` | 262144 | attachment size limit |
| `MaxFileChars` | `M365_TICKET_MAX_FILE_CHARS` | 20000 | per-attachment inline budget |
| `ExtraRedact` | — | `@()` | extra regexes applied after the built-ins (org-specific identifiers) |

The caps are used to flag what would not fit. There is no separate "max
attachments" setting: how many can be delivered is derived from `MaxMessages`,
because **every attachment batch is also a turn** — with the default 60 turns
and 3 attachments per message that is a budget of 180, far beyond any real
ticket (the largest seen across 100 tickets was 4). It uploads them, attaches
each batch through `messageAnnotations`, and splits them across turns at
`MaxAttachmentsPerMessage` per turn.

Credentials come from `FRESHSERVICE_SUBDOMAIN`, `FRESHSERVICE_SESSION`,
`FRESHSERVICE_BASE_URL`, `M365_REFRESH_TOKEN`/`M365_ACCESS_TOKEN`, or `$Config` —
no fsvc, no repo imports.

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
src/chat-api.js history/deletion/upload REST (GetChats, GetConversation, DeleteConversation, UploadFile)
src/attachments.js attachment planning: isImage, planImageBatches, manifest rendering
src/graph-upload.js  document upload to OneDrive copilotuploads + LocalFile annotations
src/index.js    public API: ask() + M365Session
src/session-store.js  persisted default conversation (session.json)
src/prompt.js   CLI context/prompt assembly (--context, --new, --temporary)
src/ticket.js   Freshservice ticket fetch/render + PII redaction
src/log.js      optional debug logging (M365_DEBUG=1)
scripts/        ask-ticket.js (repo imports), ask-ticket-standalone.ps1 (self-contained, pwsh 7+)
```

The ticket scripts redact email addresses, IP addresses, MAC addresses, phone
numbers and Windows network-config identifiers (host name, DNS suffixes, DHCPv6
IDs) from the ticket before sending it to Copilot. `ask-ticket-standalone.ps1` (PowerShell 7+) is the
self-contained variant: it needs no Node and no repo imports, and uses the OS
certificate store, so no `NODE_EXTRA_CA_CERTS` is needed on a TLS-inspecting
corporate proxy.

This is a trimmed extraction of `m365-copilot-proxy` (same authors' reverse
engineering); it keeps only the plain-chat path. No tool-calling, agents, or
image generation.
