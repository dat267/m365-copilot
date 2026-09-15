import test from "node:test";
import assert from "node:assert/strict";

import {
  truncateContext,
  fetchTicket,
  renderTicket,
  resolveConfig,
  parseTicketArgs,
  makeFreshserviceGet,
  normalizeAttachment,
  redactPII,
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

test("normalizeAttachment reads the Freshservice attachment field names", () => {
  const a = normalizeAttachment(
    { id: 7, name: "Issue_Query.png", content_type: "image/png", size: 22065, attachment_url: "https://x/a" },
    "ticket",
  );

  assert.deepEqual(a, {
    name: "Issue_Query.png",
    contentType: "image/png",
    size: 22065,
    url: "https://x/a",
    source: "ticket",
  });
});

test("normalizeAttachment tolerates alternate field spellings", () => {
  const a = normalizeAttachment(
    { filename: "trace.log", mime_type: "text/plain", file_size: 900, url: "https://y/b" },
    "conversation 42",
  );

  assert.equal(a.name, "trace.log");
  assert.equal(a.contentType, "text/plain");
  assert.equal(a.size, 900);
  assert.equal(a.url, "https://y/b");
  assert.equal(a.source, "conversation 42");
});

test("normalizeAttachment falls back to canonical_url", () => {
  // The real private-API object carries both a signed `attachment_url` and a
  // `canonical_url`; the signed one wins, but canonical is the fallback.
  const signed = normalizeAttachment(
    { name: "a.pdf", attachment_url: "https://signed", canonical_url: "https://canon" },
    "ticket",
  );
  const canonOnly = normalizeAttachment({ name: "a.pdf", canonical_url: "https://canon" }, "ticket");

  assert.equal(signed.url, "https://signed");
  assert.equal(canonOnly.url, "https://canon");
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
    attachments: [],
  });
  assert.equal(calls[0].path, "tickets/10100");
  assert.deepEqual(calls[1], {
    path: "tickets/10100/conversations",
    query: { per_page: "100", order_by: "created_at", order_type: "asc", page: "1" },
  });
});

test("fetchTicket asks the private API to include attachment details", async () => {
  // The private API returned the `attachments` array only for a ticket GET that
  // carried the web client's `include` list; a bare GET omits it (verified
  // against a live capture), so no images/files were ever detected.
  const calls = [];
  const get = async (path, query) => {
    calls.push({ path, query });
    if (path === "tickets/10100") return { ticket: { id: 10100 } };
    return { conversations: [], meta: { has_next: false } };
  };

  await fetchTicket(10100, get);

  assert.equal(calls[0].path, "tickets/10100");
  assert.equal(calls[0].query?.include, "requester,stats,phone,feedback,ticket_status");
});

test("fetchTicket collects attachments from the ticket, conversations and linked attachments", async () => {
  const get = async (path) => {
    if (path === "tickets/10100") {
      return { ticket: { id: 10100, attachments: [{ name: "a.png", content_type: "image/png" }], cloud_files: [] } };
    }
    if (path === "tickets/10100/linked-attachments") {
      return { attachments: [{ name: "c.pdf", content_type: "application/pdf" }], meta: { count: 1 } };
    }
    return {
      conversations: [{ id: 5, attachments: [{ name: "b.log", content_type: "text/plain" }] }],
      meta: { has_next: false },
    };
  };

  const out = await fetchTicket(10100, get);

  assert.deepEqual(
    out.attachments.map((a) => [a.name, a.source]),
    [
      ["a.png", "ticket"],
      ["b.log", "conversation 5"],
      ["c.pdf", "linked"],
    ],
  );
});

test("fetchTicket still works when linked attachments are unavailable", async () => {
  const get = async (path) => {
    if (path === "tickets/10100") return { ticket: { id: 10100 } };
    if (path.endsWith("linked-attachments")) throw new Error("Freshservice 404 Not Found");
    return { conversations: [], meta: { has_next: false } };
  };

  const out = await fetchTicket(10100, get);

  assert.deepEqual(out.attachments, []);
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

test("renderTicket appends the attachment manifest", () => {
  const md = renderTicket({ id: 1, subject: "Printer" }, [], {
    attachments: [
      { name: "Issue_Query.png", contentType: "image/png", size: 22065, source: "ticket", url: "https://x/a" },
    ],
  });

  assert.match(md, /## Attachments/);
  assert.match(md, /Issue_Query\.png/);
  assert.ok(md.indexOf("## Conversations") < md.indexOf("## Attachments"), "manifest comes after the trace");
});

test("renderTicket notes the absence of attachments", () => {
  assert.match(renderTicket({ id: 1, subject: "s" }, []), /## Attachments\n\n\(none\)/);
});

test("resolveConfig builds the base URL from env subdomain and session", () => {
  const env = { FRESHSERVICE_SUBDOMAIN: "acme", FRESHSERVICE_SESSION: "sess" };
  const io = { defaultPath: "/home/u/.config/m365-copilot/freshservice.json", exists: () => false, readFile: () => "" };

  assert.deepEqual(resolveConfig(env, io), {
    baseUrl: "https://acme.freshservice.com",
    session: "sess",
  });
});

test("resolveConfig falls back to the config file named by M365_FRESHSERVICE_CONFIG", () => {
  const env = { M365_FRESHSERVICE_CONFIG: "/tmp/freshservice.json" };
  const io = {
    defaultPath: "/home/u/.config/m365-copilot/freshservice.json",
    exists: (p) => p === "/tmp/freshservice.json",
    readFile: () => JSON.stringify({ subdomain: "fileco", session: "filesess" }),
  };

  assert.deepEqual(resolveConfig(env, io), {
    baseUrl: "https://fileco.freshservice.com",
    session: "filesess",
  });
});

test("resolveConfig lets env override the config file", () => {
  const env = { FRESHSERVICE_SESSION: "envsess", M365_FRESHSERVICE_CONFIG: "/tmp/freshservice.json" };
  const io = {
    defaultPath: "/nope",
    exists: (p) => p === "/tmp/freshservice.json",
    readFile: () => JSON.stringify({ subdomain: "fileco", session: "filesess" }),
  };

  assert.deepEqual(resolveConfig(env, io), {
    baseUrl: "https://fileco.freshservice.com",
    session: "envsess",
  });
});

test("resolveConfig uses baseUrl and trims a trailing slash", () => {
  const env = {
    FRESHSERVICE_SUBDOMAIN: "acme",
    FRESHSERVICE_SESSION: "s",
    FRESHSERVICE_BASE_URL: "https://mock.local/",
  };
  const io = { defaultPath: "/nope", exists: () => false, readFile: () => "" };

  assert.equal(resolveConfig(env, io).baseUrl, "https://mock.local");
});

test("resolveConfig reads the default config path when there is no override", () => {
  const io = {
    defaultPath: "/home/u/.config/m365-copilot/freshservice.json",
    exists: (p) => p === "/home/u/.config/m365-copilot/freshservice.json",
    readFile: () => JSON.stringify({ subdomain: "homeco", session: "homesess" }),
  };

  assert.deepEqual(resolveConfig({}, io), {
    baseUrl: "https://homeco.freshservice.com",
    session: "homesess",
  });
});

test("resolveConfig prefers a local freshservice.json over the default path", () => {
  const io = {
    defaultPath: "/home/u/.config/m365-copilot/freshservice.json",
    exists: (p) => p === "freshservice.json" || p === "/home/u/.config/m365-copilot/freshservice.json",
    readFile: (p) =>
      JSON.stringify(
        p === "freshservice.json"
          ? { subdomain: "localco", session: "localsess" }
          : { subdomain: "homeco", session: "homesess" },
      ),
  };

  assert.deepEqual(resolveConfig({}, io), {
    baseUrl: "https://localco.freshservice.com",
    session: "localsess",
  });
});

test("resolveConfig throws when no session is configured", () => {
  const env = { FRESHSERVICE_SUBDOMAIN: "acme" };
  const io = { defaultPath: "/nope", exists: () => false, readFile: () => "" };

  assert.throws(() => resolveConfig(env, io), /session/i);
});

test("resolveConfig throws when no base URL can be determined", () => {
  const env = { FRESHSERVICE_SESSION: "sess" };
  const io = { defaultPath: "/nope", exists: () => false, readFile: () => "" };

  assert.throws(() => resolveConfig(env, io), /subdomain|base.?url/i);
});

test("resolveConfig does not read the fsvc config", () => {
  const io = {
    defaultPath: "/home/u/.config/m365-copilot/freshservice.json",
    exists: () => false,
    readFile: () => {
      throw new Error("should not read any config file");
    },
  };

  assert.throws(() => resolveConfig({}, io), /subdomain|base.?url/i);
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

test("redactPII replaces email addresses", () => {
  assert.equal(redactPII("Email omar.saleh@example.com now"), "Email [redacted-email] now");
});

test("redactPII replaces phone numbers in common formats", () => {
  const cases = [
    ["call +1 (555) 123-4567", "call [redacted-phone]"],
    ["call 555-123-4567", "call [redacted-phone]"],
    ["call (555) 123-4567", "call [redacted-phone]"],
    ["call +44 20 7946 0958", "call [redacted-phone]"],
    ["call 07123456789", "call [redacted-phone]"],
    ["call 020 7946 0958", "call [redacted-phone]"],
  ];
  for (const [input, expected] of cases) assert.equal(redactPII(input), expected, input);
});

test("redactPII replaces IPv4 addresses", () => {
  const cases = [
    ["from 192.168.1.10", "from [redacted-ip]"],
    ["host 10.0.0.0/8", "host [redacted-ip]/8"],
    ["public 203.0.113.7:8080", "public [redacted-ip]:8080"],
  ];
  for (const [input, expected] of cases) assert.equal(redactPII(input), expected, input);
});

test("redactPII replaces IPv6 addresses", () => {
  assert.equal(redactPII("connect to 2001:db8::1 now"), "connect to [redacted-ip] now");
  assert.equal(redactPII("loopback fe80::1"), "loopback [redacted-ip]");
  assert.equal(
    redactPII("full 2001:0db8:0000:0000:0000:0000:0000:0001"),
    "full [redacted-ip]",
  );
});

test("redactPII leaves non-IP dotted/colon runs alone", () => {
  for (const text of [
    "std::vector and ns::foo are not addresses",
    "at 12:30:00 today",
    "mac aa:bb:cc:dd:ee:ff",
    "version 1.2.3.400",
    "file v1.2.3.4",
  ]) {
    assert.equal(redactPII(text), text, text);
  }
});

test("redactPII leaves dates, ticket ids and references alone", () => {
  const text = "Created 2026-08-01T10:30:00Z, ticket #10100, ref INC0012345";
  assert.equal(redactPII(text), text);
});

test("redactPII leaves long numeric ids and attachment URLs intact", () => {
  const url =
    "https://acme.attachments.freshservice.com/data/helpdesk/attachments/production/21117119076/original/a.jpeg" +
    "?response-content-type=image/jpeg&Expires=1789502963&Signature=abc123";

  assert.equal(redactPII(url), url, "a signed attachment URL must survive intact");
  assert.equal(redactPII("Requester  : 21003608052"), "Requester  : 21003608052");
});
