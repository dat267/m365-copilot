import test from "node:test";
import assert from "node:assert/strict";

import {
  truncateContext,
  fetchTicket,
  renderTicket,
  resolveConfig,
  parseTicketArgs,
  makeFreshserviceGet,
  TRUNCATION_MARKER,
} from "./ticket.js";

test("truncateContext leaves a payload under budget unchanged", () => {
  assert.equal(truncateContext("short ticket", 100), "short ticket");
});

test("truncateContext keeps the head and tail under the budget when over", () => {
  const text = "H".repeat(200) + "T".repeat(200);
  const out = truncateContext(text, 100);

  assert.ok(out.length <= 100, `expected <= 100 chars, got ${out.length}`);
  assert.ok(out.startsWith("H"), "kept the beginning of the payload");
  assert.ok(out.endsWith("T"), "kept the end of the payload");
  assert.ok(out.includes(TRUNCATION_MARKER), "marked the elision");
});

test("fetchTicket fetches the ticket and its conversations", async () => {
  const calls = [];
  const get = async (path, query) => {
    calls.push({ path, query });
    if (path === "tickets/10100") return { ticket: { id: 10100, subject: "Printer" } };
    return { conversations: [{ id: 1, body_text: "hi" }], meta: { has_next: false } };
  };

  const out = await fetchTicket(10100, get);

  assert.deepEqual(out, {
    ticket: { id: 10100, subject: "Printer" },
    conversations: [{ id: 1, body_text: "hi" }],
  });
  assert.equal(calls[0].path, "tickets/10100");
  assert.deepEqual(calls[1], {
    path: "tickets/10100/conversations",
    query: { per_page: "100", order_by: "created_at", order_type: "asc", page: "1" },
  });
});

test("fetchTicket follows meta.has_next across conversation pages", async () => {
  const get = async (path, query) => {
    if (path === "tickets/10100") return { ticket: {} };
    return Number(query.page) === 1
      ? { conversations: [{ id: 1 }], meta: { has_next: true } }
      : { conversations: [{ id: 2 }], meta: { has_next: false } };
  };

  const out = await fetchTicket(10100, get);

  assert.deepEqual(out.conversations, [{ id: 1 }, { id: 2 }]);
});

test("renderTicket renders metadata, description and conversation trace", () => {
  const ticket = {
    id: 10100,
    display_id: 10100,
    subject: "Printer not working",
    status: 2,
    status_name: "Open",
    priority: 2,
    priority_name: "Medium",
    urgency: 1,
    impact: 1,
    group_id: 4001,
    group_name: "Support",
    requester_id: 5001,
    requester_name: "Omar Saleh",
    responder_id: 3100,
    responder_name: "Nadia Rahman",
    department_id: 5100,
    department_name: "IT",
    created_at: "2026-08-01T10:00:00Z",
    updated_at: "2026-08-02T09:00:00Z",
    description: "<p>Printer jammed &amp; smoking</p>",
    description_text: "Printer jammed & smoking",
  };
  const conversations = [
    { id: 2, user_id: 2100, incoming: true, created_at: "2026-08-01T10:30:00Z", body_text: "<p>Please fix the printer</p>" },
    { id: 1, user_id: 3100, user: { name: "Nadia Rahman" }, incoming: false, created_at: "2026-08-01T11:00:00Z", body_text: "<p>Will do</p>" },
  ];

  const md = renderTicket(ticket, conversations);

  assert.ok(md.startsWith("# Ticket #10100 — Printer not working\n\n"), "title");
  assert.ok(md.includes("Status     : Open"), "status name");
  assert.ok(md.includes("Priority   : Medium"), "priority name");
  assert.ok(md.includes("Urgency    : Low"), "urgency numeric mapped to name");
  assert.ok(md.includes("Requester  : Omar Saleh"), "requester name");
  assert.ok(md.includes("Responder  : Nadia Rahman"), "responder name");
  assert.ok(md.includes("Department : IT"), "department name");
  assert.ok(md.includes("Printer jammed & smoking"), "description, entities decoded");
  assert.ok(md.includes("## Conversations"), "conversations heading");
  assert.ok(md.includes("### 2100 (incoming, 2026-08-01T10:30:00Z)"), "author falls back to id");
  assert.ok(md.includes("### Nadia Rahman (outgoing, 2026-08-01T11:00:00Z)"), "prefers nested user name");
  assert.ok(!md.includes("<p>"), "html stripped");
  assert.ok(md.indexOf("2026-08-01T10:30:00Z") < md.indexOf("2026-08-01T11:00:00Z"), "oldest first");
});

test("resolveConfig builds the base URL from env subdomain and session", () => {
  const env = { FSVC_SUBDOMAIN: "acme", FSVC_ITILDESK_SESSION: "sess" };
  const io = { defaultPath: "/home/u/.config/fsvc/fsvc.json", exists: () => false, readFile: () => "" };

  assert.deepEqual(resolveConfig(env, io), {
    baseUrl: "https://acme.freshservice.com",
    session: "sess",
  });
});

test("resolveConfig falls back to the config file named by FSVC_CONFIG_FILE", () => {
  const env = { FSVC_CONFIG_FILE: "/tmp/fsvc.json" };
  const io = {
    defaultPath: "/home/u/.config/fsvc/fsvc.json",
    exists: (p) => p === "/tmp/fsvc.json",
    readFile: () => JSON.stringify({ subdomain: "fileco", "itildesk-session": "filesess" }),
  };

  assert.deepEqual(resolveConfig(env, io), {
    baseUrl: "https://fileco.freshservice.com",
    session: "filesess",
  });
});

test("resolveConfig uses base-url and trims a trailing slash", () => {
  const env = { FSVC_SUBDOMAIN: "acme", FSVC_ITILDESK_SESSION: "s", FSVC_BASE_URL: "https://mock.local/" };
  const io = { defaultPath: "/nope", exists: () => false, readFile: () => "" };

  assert.equal(resolveConfig(env, io).baseUrl, "https://mock.local");
});

test("resolveConfig reads the default config path when there is no override", () => {
  const env = {};
  const io = {
    defaultPath: "/home/u/.config/fsvc/fsvc.json",
    exists: (p) => p === "/home/u/.config/fsvc/fsvc.json",
    readFile: () => JSON.stringify({ subdomain: "homeco", "itildesk-session": "homesess" }),
  };

  assert.deepEqual(resolveConfig(env, io), {
    baseUrl: "https://homeco.freshservice.com",
    session: "homesess",
  });
});

test("resolveConfig prefers a local fsvc.json over the default path", () => {
  const env = {};
  const io = {
    defaultPath: "/home/u/.config/fsvc/fsvc.json",
    exists: (p) => p === "fsvc.json" || p === "/home/u/.config/fsvc/fsvc.json",
    readFile: (p) =>
      JSON.stringify(
        p === "fsvc.json"
          ? { subdomain: "localco", "itildesk-session": "localsess" }
          : { subdomain: "homeco", "itildesk-session": "homesess" },
      ),
  };

  assert.deepEqual(resolveConfig(env, io), {
    baseUrl: "https://localco.freshservice.com",
    session: "localsess",
  });
});

test("resolveConfig throws when no session is configured", () => {
  const env = { FSVC_SUBDOMAIN: "acme" };
  const io = { defaultPath: "/nope", exists: () => false, readFile: () => "" };

  assert.throws(() => resolveConfig(env, io), /session/i);
});

test("resolveConfig throws when no base URL can be determined", () => {
  const env = { FSVC_ITILDESK_SESSION: "sess" };
  const io = { defaultPath: "/nope", exists: () => false, readFile: () => "" };

  assert.throws(() => resolveConfig(env, io), /subdomain|base.?url/i);
});

test("parseTicketArgs parses the id, instruction and default maxChars", () => {
  assert.deepEqual(parseTicketArgs(["10100", "draft", "a", "reply"]), {
    id: 10100,
    instruction: "draft a reply",
    maxChars: 60000,
    follow: false,
    fresh: false,
  });
});

test("parseTicketArgs honours --new", () => {
  assert.equal(parseTicketArgs(["--new", "10100", "hi"]).fresh, true);
});

test("parseTicketArgs honours --max-chars", () => {
  assert.equal(parseTicketArgs(["--max-chars=500", "10100", "hi"]).maxChars, 500);
});

test("parseTicketArgs honours --follow", () => {
  assert.equal(parseTicketArgs(["--follow", "10100", "hi"]).follow, true);
});

test("makeFreshserviceGet calls the private API with cookie auth and query params", async () => {
  let seen;
  const fetchImpl = async (url, opts) => {
    seen = { url: String(url), opts };
    return { ok: true, json: async () => ({ ticket: { id: 1 } }) };
  };

  const get = makeFreshserviceGet({ baseUrl: "https://acme.freshservice.com", session: "SESS", fetchImpl });
  const out = await get("tickets/10100", { per_page: "100" });

  assert.deepEqual(out, { ticket: { id: 1 } });
  assert.equal(seen.url, "https://acme.freshservice.com/api/_/tickets/10100?per_page=100");
  assert.equal(seen.opts.headers.Cookie, "_itildesk_session=SESS");
  assert.equal(seen.opts.headers.Accept, "application/json");
});

test("makeFreshserviceGet throws on a non-OK response", async () => {
  const fetchImpl = async () => ({ ok: false, status: 401, statusText: "Unauthorized" });
  const get = makeFreshserviceGet({ baseUrl: "https://x", session: "s", fetchImpl });

  await assert.rejects(() => get("tickets/1"), /401/);
});
