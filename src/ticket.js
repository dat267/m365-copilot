// Ticket-context helpers for the Freshservice → m365-copilot wrapper.
//
// Pure logic (no I/O): payload truncation, argument parsing, Markdown
// rendering, and the injected HTTP boundary that fetches a ticket. Tested with
// `node --test`.

import { renderAttachmentManifest } from "./attachments.js";

/** Marker inserted where an over-budget payload was elided. Constant length so
 *  the truncated result is exactly `maxChars` long. */
export const TRUNCATION_MARKER = "\n\n[...truncated...]\n\n";

/**
 * Normalises a Freshservice attachment into `{ name, contentType, size, url, source }`.
 *
 * The private `api/_` attachment object could not be confirmed from a capture
 * (the sampled tickets had none), so the well-known public-API field names are
 * accepted along with their common variants.
 */
export function normalizeAttachment(raw, source) {
  const a = raw ?? {};
  return {
    name: a.name ?? a.filename ?? a.file_name ?? "",
    contentType: a.content_type ?? a.contentType ?? a.mime_type ?? "",
    size: a.size ?? a.file_size ?? a.content_length ?? null,
    url: a.attachment_url ?? a.url ?? a.download_url ?? "",
    source,
  };
}

/** Fetches a ticket and all of its conversations. `get(path, query)` is the
 *  injected HTTP boundary; it returns parsed JSON. */export async function fetchTicket(id, get) {
  const { ticket } = await get(`tickets/${id}`);
  const conversations = [];
  const maxPages = 1000; // safety cap
  for (let page = 1; page <= maxPages; page++) {
    const res = await get(`tickets/${id}/conversations`, {
      per_page: "100",
      order_by: "created_at",
      order_type: "asc",
      page: String(page),
    });
    conversations.push(...(res.conversations ?? []));
    if (!res.meta?.has_next) break;
  }

  const attachments = [
    ...collectAttachments(ticket, "ticket"),
    ...conversations.flatMap((c) => collectAttachments(c, `conversation ${c.id}`)),
    ...(await fetchLinkedAttachments(id, get)),
  ];

  return { ticket, conversations, attachments };
}

function collectAttachments(container, source) {
  const raw = [...(container?.attachments ?? []), ...(container?.cloud_files ?? [])];
  return raw.map((a) => normalizeAttachment(a, source));
}

/** `linked-attachments` is a separate, newer endpoint — treat it as optional so
 *  an older Freshservice build (or a plan without it) doesn't fail the fetch. */
async function fetchLinkedAttachments(id, get) {
  try {
    const res = await get(`tickets/${id}/linked-attachments`);
    return (res?.attachments ?? []).map((a) => normalizeAttachment(a, "linked"));
  } catch {
    return [];
  }
}

const STATUS_NAMES = { 2: "Open", 3: "Pending", 4: "Resolved", 5: "Closed" };
const PRIORITY_NAMES = { 1: "Low", 2: "Medium", 3: "High", 4: "Urgent" };
const URGENCY_IMPACT_NAMES = { 1: "Low", 2: "Medium", 3: "High" };

const META_FIELDS = [
  ["Status", "status", "status_name"],
  ["Priority", "priority", "priority_name"],
  ["Urgency", "urgency", "urgency_name"],
  ["Impact", "impact", "impact_name"],
  ["Group", "group_id", "group_name"],
  ["Requester", "requester_id", "requester_name"],
  ["Responder", "responder_id", "responder_name"],
  ["Department", "department_id", "department_name"],
  ["Created", "created_at", ""],
  ["Updated", "updated_at", ""],
];

function fieldValue(obj, key) {
  if (!key) return "";
  const v = obj?.[key];
  if (v == null) return "";
  return typeof v === "object" ? JSON.stringify(v) : String(v);
}

function decodeEntities(s) {
  const named = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", nbsp: " " };
  return s.replace(/&(?:#(\d+)|#x([0-9a-f]+)|([a-z]+));/gi, (m, dec, hex, name) => {
    if (dec) return String.fromCodePoint(Number(dec));
    if (hex) return String.fromCodePoint(parseInt(hex, 16));
    return named[name.toLowerCase()] ?? m;
  });
}

/** Removes HTML tags and decodes entities. */
function stripHtml(s) {
  return decodeEntities(String(s ?? "").replace(/<[^>]*>/g, "")).trim();
}

function mapNumeric(key, val) {
  if (!/^\d+$/.test(val)) return val;
  const n = Number(val);
  if (key === "urgency" || key === "impact") return URGENCY_IMPACT_NAMES[n] ?? val;
  if (key === "priority") return PRIORITY_NAMES[n] ?? val;
  if (key === "status") return STATUS_NAMES[n] ?? val;
  return val;
}

function conversationAuthor(c) {
  const name = fieldValue(c?.user, "name");
  return name || fieldValue(c, "user_id");
}

/** Renders a ticket and its conversation trace as a Markdown document.
 *
 *  Attachment *contents* are not downloaded — but their inventory is listed so
 *  the model knows the complete set, including anything that cannot be
 *  delivered in one message (see ../README.md "Ticket attachments"). */
export function renderTicket(ticket, conversations, { attachments = [] } = {}) {
  const display = fieldValue(ticket, "display_id") || fieldValue(ticket, "id");
  const width = Math.max(...META_FIELDS.map(([label]) => label.length));
  const out = [`# Ticket #${display} — ${fieldValue(ticket, "subject")}`, ""];

  for (const [label, key, nameKey] of META_FIELDS) {
    const raw = fieldValue(ticket, nameKey) || fieldValue(ticket, key);
    out.push(`${label.padEnd(width)} : ${mapNumeric(key, raw)}`);
  }
  out.push("");

  const desc = stripHtml(fieldValue(ticket, "description_text") || fieldValue(ticket, "description"));
  if (desc) out.push(desc, "");

  out.push("## Conversations", "");
  const convs = conversations ?? [];
  if (convs.length === 0) out.push("(none)");
  for (const c of convs) {
    const dir = c.incoming === true ? "incoming" : "outgoing";
    const body = stripHtml(fieldValue(c, "body_text") || fieldValue(c, "body"));
    out.push(`### ${conversationAuthor(c)} (${dir}, ${fieldValue(c, "created_at")})`);
    out.push(body || "(no body)");
    out.push("");
  }

  out.push("", renderAttachmentManifest(attachments), "");
  return out.join("\n");
}

/** Resolves Freshservice credentials. This repo owns its own config — it does
 *  NOT read fsvc's (`fsvc.json` / `FSVC_*`); nothing here shells out to or
 *  depends on the fsvc tool.
 *
 *  Precedence: env vars, then the config file at `M365_FRESHSERVICE_CONFIG`,
 *  else `./freshservice.json`, else `io.defaultPath`.
 *
 *  Config file keys: `{ "subdomain": "acme", "session": "...", "baseUrl": "" }`.
 *  Env vars: `FRESHSERVICE_SUBDOMAIN`, `FRESHSERVICE_SESSION`,
 *  `FRESHSERVICE_BASE_URL`.
 *
 *  Returns `{ baseUrl, session }`. */
export function resolveConfig(env, io) {
  let file = {};
  const localPath = "freshservice.json";
  const path = env.M365_FRESHSERVICE_CONFIG || (io.exists(localPath) ? localPath : io.defaultPath);
  if (path && io.exists(path)) {
    file = JSON.parse(io.readFile(path));
  }
  const subdomain = env.FRESHSERVICE_SUBDOMAIN ?? file.subdomain;
  const session = env.FRESHSERVICE_SESSION ?? file.session;
  const override = env.FRESHSERVICE_BASE_URL ?? file.baseUrl;
  const derived = subdomain ? `https://${subdomain}.freshservice.com` : "";
  const baseUrl = override ? String(override).replace(/\/+$/, "") : derived;
  if (!baseUrl) {
    throw new Error(
      "no Freshservice base URL: set FRESHSERVICE_SUBDOMAIN (or FRESHSERVICE_BASE_URL), or subdomain/baseUrl in freshservice.json",
    );
  }
  if (!session) {
    throw new Error("no Freshservice session: set FRESHSERVICE_SESSION, or session in freshservice.json");
  }
  return { baseUrl, session };
}

/** Parses `ask-ticket` argv: `<ticket-id> <instruction...>`, plus
 *  `--max-chars=N` and `--follow`. */
export function parseTicketArgs(argv) {
  let maxChars = 60000;
  let follow = false;
  let fresh = false;
  const rest = [];
  for (const a of argv) {
    if (a.startsWith("--max-chars=")) maxChars = Number(a.slice("--max-chars=".length));
    else if (a === "--follow") follow = true;
    else if (a === "--new") fresh = true;
    else rest.push(a);
  }
  const [id, ...instruction] = rest;
  return { id: Number(id), instruction: instruction.join(" "), maxChars, follow, fresh };
}

/** Builds the injected `get(path, query)` used by `fetchTicket`, bound to a
 *  Freshservice base URL + session cookie. `fetchImpl` defaults to global
 *  fetch. */
export function makeFreshserviceGet({ baseUrl, session, fetchImpl = fetch }) {
  return async (path, query) => {
    const url = new URL(`${baseUrl}/api/_/${path}`);
    for (const [k, v] of Object.entries(query ?? {})) url.searchParams.set(k, v);
    const res = await fetchImpl(url, {
      headers: { Accept: "application/json", Cookie: `_itildesk_session=${session}` },
    });
    if (!res.ok) throw new Error(`Freshservice ${res.status} ${res.statusText} for ${path}`);
    return res.json();
  };
}

const EMAIL_RE = /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g;
// Ordered: international (+CC), parenthesised, NANP 3-3-4, then 0-prefixed
// national numbers.
//
// Deliberately does NOT include a bare `\d{10,15}` rule: Freshservice ids
// (ticket/requester/attachment) and signed-URL params such as `Expires=` are
// long digit runs, and redacting them corrupted attachment URLs and made
// requester ids render as "[redacted-phone]". Trade-off: a phone written with
// no separators and no leading + or 0 is not redacted.
const PHONE_PATTERNS = [
  /\+\d[\d \t().-]{5,}\d/g,
  /\(\d{3}\)[ .-]?\d{3}[ .-]\d{4}/g,
  /\b\d{3}[ .-]\d{3}[ .-]\d{4}\b/g,
  /\b0\d{1,3}[ .-]\d{3,4}[ .-]?\d{3,4}\b/g,
  /\b0\d{9,10}\b/g,
];

// IPv4 with 0-255 octets; the lookarounds keep it out of longer dotted runs
// (so a version like `v1.2.3.4` survives).
const IPV4_RE = /(?<![\w.])(?:25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(?:\.(?:25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}(?![\w.])/g;

// IPv6: full 8-group and `::`-compressed forms only. Deliberately omits the
// non-compressed shorthand, which would match HH:MM:SS times; the lookarounds
// keep identifiers like `std::vector` intact. A match is redacted only when it
// carries >=4 hex digits, so a bare `a::b` is left alone.
const IPV6_RE = /(?<![0-9A-Za-z:])(?:(?:[0-9a-fA-F]{1,4}:){7}[0-9a-fA-F]{1,4}|(?:[0-9a-fA-F]{1,4}:){1,7}:|(?:[0-9a-fA-F]{1,4}:){1,6}:[0-9a-fA-F]{1,4}|(?:[0-9a-fA-F]{1,4}:){1,5}(?::[0-9a-fA-F]{1,4}){1,2}|(?:[0-9a-fA-F]{1,4}:){1,4}(?::[0-9a-fA-F]{1,4}){1,3}|(?:[0-9a-fA-F]{1,4}:){1,3}(?::[0-9a-fA-F]{1,4}){1,4}|(?:[0-9a-fA-F]{1,4}:){1,2}(?::[0-9a-fA-F]{1,4}){1,5}|:(?:(?::[0-9a-fA-F]{1,4}){1,7}|:))(?![0-9A-Za-z:])/g;

/** Redacts email addresses, IP addresses and phone numbers so they are not
 *  sent to the model. */
export function redactPII(text) {
  let out = String(text ?? "").replace(EMAIL_RE, "[redacted-email]");
  out = out.replace(IPV4_RE, "[redacted-ip]");
  out = out.replace(IPV6_RE, (m) => (m.replace(/[^0-9a-fA-F]/g, "").length >= 4 ? "[redacted-ip]" : m));
  for (const re of PHONE_PATTERNS) out = out.replace(re, "[redacted-phone]");
  return out;
}

export function truncateContext(text, maxChars) {
  if (text.length <= maxChars) return text;
  if (maxChars <= TRUNCATION_MARKER.length) return text.slice(0, maxChars);
  const budget = maxChars - TRUNCATION_MARKER.length;
  const head = Math.ceil(budget * 0.6);
  const tail = budget - head;
  return text.slice(0, head) + TRUNCATION_MARKER + text.slice(text.length - tail);
}
