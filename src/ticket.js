// Ticket-context helpers for the Freshservice → m365-copilot wrapper.
//
// Pure logic (no I/O): payload truncation, argument parsing, Markdown
// rendering, and the injected HTTP boundary that fetches a ticket. Tested with
// `node --test`.

/** Marker inserted where an over-budget payload was elided. Constant length so
 *  the truncated result is exactly `maxChars` long. */
export const TRUNCATION_MARKER = "\n\n[...truncated...]\n\n";

/** Fetches a ticket and all of its conversations. `get(path, query)` is the
 *  injected HTTP boundary; it returns parsed JSON. */
export async function fetchTicket(id, get) {
  const { ticket } = await get(`tickets/${id}`);
  const conversations = [];
  const maxPages = 1000; // safety cap, mirrors fsvc's MaxPages
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
  return { ticket, conversations };
}

const STATUS_NAMES = { 2: "Open", 3: "Pending", 4: "Resolved", 5: "Closed" };
const PRIORITY_NAMES = { 1: "Low", 2: "Medium", 3: "High", 4: "Urgent" };
const URGENCY_IMPACT_NAMES = { 1: "Low", 2: "Medium", 3: "High" };

// label, key, nameKey — mirrors fsvc show.go ticketMetaFields.
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

/** Removes HTML tags and decodes entities, mirroring fsvc's stripHTML. */
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

/** Renders a ticket and its conversation trace as a Markdown document,
 *  mirroring `fsvc tickets show`. No images/attachments are downloaded — this
 *  text is fed to a model, not saved to disk. */
export function renderTicket(ticket, conversations) {
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
  return out.join("\n");
}

/** Resolves Freshservice credentials, mirroring fsvc/min:
 *  `FSVC_CONFIG_FILE` → `./fsvc.json` → `io.defaultPath`, with env vars
 *  overriding file values. Returns `{ baseUrl, session }`. */
export function resolveConfig(env, io) {
  let file = {};
  const path = env.FSVC_CONFIG_FILE || (io.exists("fsvc.json") ? "fsvc.json" : io.defaultPath);
  if (path && io.exists(path)) {
    file = JSON.parse(io.readFile(path));
  }
  const subdomain = env.FSVC_SUBDOMAIN ?? file.subdomain;
  const session = env.FSVC_ITILDESK_SESSION ?? file["itildesk-session"];
  const override = env.FSVC_BASE_URL ?? file["base-url"];
  const derived = subdomain ? `https://${subdomain}.freshservice.com` : "";
  const baseUrl = override ? String(override).replace(/\/+$/, "") : derived;
  if (!baseUrl) {
    throw new Error(
      "no Freshservice base URL: set subdomain (or base-url) in fsvc config, or FSVC_SUBDOMAIN/FSVC_BASE_URL",
    );
  }
  if (!session) {
    throw new Error(
      "no Freshservice session: set itildesk-session in fsvc config, or FSVC_ITILDESK_SESSION",
    );
  }
  return { baseUrl, session };
}

/** Parses `ask-ticket` argv: `<ticket-id> <instruction...>`, plus
 *  `--max-chars=N` and `--follow`. */
export function parseTicketArgs(argv) {
  let maxChars = 60000;
  let follow = false;
  const rest = [];
  for (const a of argv) {
    if (a.startsWith("--max-chars=")) maxChars = Number(a.slice("--max-chars=".length));
    else if (a === "--follow") follow = true;
    else rest.push(a);
  }
  const [id, ...instruction] = rest;
  return { id: Number(id), instruction: instruction.join(" "), maxChars, follow };
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

export function truncateContext(text, maxChars) {
  if (text.length <= maxChars) return text;
  if (maxChars <= TRUNCATION_MARKER.length) return text.slice(0, maxChars);
  const budget = maxChars - TRUNCATION_MARKER.length;
  const head = Math.ceil(budget * 0.6);
  const tail = budget - head;
  return text.slice(0, head) + TRUNCATION_MARKER + text.slice(text.length - tail);
}
