import test from "node:test";
import assert from "node:assert/strict";

import {
  encodeFrame,
  createFrameParser,
  renderTicket,
  truncateContext,
  buildPrompt,
  foldStreamText,
  parseArgs,
  wsConnect,
  chooseRefreshToken,
  redactPII,
} from "./ask-ticket-standalone.mjs";

// Builds an unmasked server->client frame (client frames are masked, server's
// are not).
function serverFrame(payload, { opcode = 0x1, fin = true } = {}) {
  const body = Buffer.isBuffer(payload) ? payload : Buffer.from(payload);
  const len = body.length;
  let header;
  if (len < 126) {
    header = Buffer.from([(fin ? 0x80 : 0) | opcode, len]);
  } else if (len < 65536) {
    header = Buffer.alloc(4);
    header[0] = (fin ? 0x80 : 0) | opcode;
    header[1] = 126;
    header.writeUInt16BE(len, 2);
  } else {
    header = Buffer.alloc(10);
    header[0] = (fin ? 0x80 : 0) | opcode;
    header[1] = 127;
    header.writeBigUInt64BE(BigInt(len), 2);
  }
  return Buffer.concat([header, body]);
}

test("createFrameParser decodes an unmasked text frame", () => {
  const frames = [];
  const push = createFrameParser((fin, opcode, payload) => frames.push([fin, opcode, payload.toString()]));

  push(serverFrame("hello"));

  assert.deepEqual(frames, [[true, 0x1, "hello"]]);
});

test("frame codec round-trips a masked client frame", () => {
  const frames = [];
  const push = createFrameParser((_fin, _opcode, payload) => frames.push(payload.toString()));

  push(encodeFrame(Buffer.from("payload")));

  assert.deepEqual(frames, ["payload"]);
});

test("createFrameParser reassembles a fragmented message", () => {
  const frames = [];
  const push = createFrameParser((_fin, _opcode, payload) => frames.push(payload.toString()));

  push(Buffer.concat([serverFrame("hel", { fin: false }), serverFrame("lo", { opcode: 0x0, fin: true })]));

  assert.deepEqual(frames, ["hello"]);
});

test("createFrameParser tolerates byte-by-byte delivery", () => {
  const frames = [];
  const push = createFrameParser((_fin, _opcode, payload) => frames.push(payload.toString()));

  for (const byte of serverFrame("split")) push(Buffer.from([byte]));

  assert.deepEqual(frames, ["split"]);
});

test("createFrameParser surfaces ping frames", () => {
  const frames = [];
  const push = createFrameParser((_fin, opcode, payload) => frames.push([opcode, payload.toString()]));

  push(serverFrame("ping", { opcode: 0x9 }));

  assert.deepEqual(frames, [[0x9, "ping"]]);
});

test("encodeFrame uses the 16-bit and 64-bit length forms", () => {
  const lengths = [];
  const push = createFrameParser((_fin, _opcode, payload) => lengths.push(payload.length));

  push(encodeFrame(Buffer.from("x".repeat(300))));
  push(encodeFrame(Buffer.alloc(70000, 0x61)));

  assert.deepEqual(lengths, [300, 70000]);
});

test("renderTicket renders metadata and the conversation trace", () => {
  const ticket = {
    id: 10100, display_id: 10100, subject: "Printer not working",
    status: 2, status_name: "Open", priority: 2, priority_name: "Medium",
    urgency: 1, impact: 1, group_name: "Support", requester_name: "Omar Saleh",
    responder_name: "Nadia Rahman", department_name: "IT",
    created_at: "2026-08-01T10:00:00Z", updated_at: "2026-08-02T09:00:00Z",
    description_text: "Printer jammed & smoking",
  };
  const conversations = [
    { id: 2, user_id: 2100, incoming: true, created_at: "2026-08-01T10:30:00Z", body_text: "<p>Please fix</p>" },
    { id: 1, user_id: 3100, user: { name: "Nadia Rahman" }, created_at: "2026-08-01T11:00:00Z", body_text: "Will do" },
  ];

  const md = renderTicket(ticket, conversations);

  assert.ok(md.startsWith("# Ticket #10100 — Printer not working\n\n"));
  assert.ok(md.includes("Status     : Open"));
  assert.ok(md.includes("Urgency    : Low"));
  assert.ok(md.includes("## Conversations"));
  assert.ok(md.includes("### 2100 (incoming, 2026-08-01T10:30:00Z)"));
  assert.ok(md.includes("### Nadia Rahman (outgoing, 2026-08-01T11:00:00Z)"));
  assert.ok(!md.includes("<p>"), "html stripped");
});

test("truncateContext keeps head and tail under budget", () => {
  const out = truncateContext("H".repeat(200) + "T".repeat(200), 100);
  assert.ok(out.length <= 100);
  assert.ok(out.startsWith("H"));
  assert.ok(out.endsWith("T"));
  assert.equal(truncateContext("short", 100), "short");
});

test("buildPrompt wraps the context and appends the instruction", () => {
  const prompt = buildPrompt("DATA", "do it");
  assert.ok(prompt.includes("<<<CONTEXT\nDATA\nCONTEXT>>>"));
  assert.ok(prompt.endsWith("do it"));
});

test("foldStreamText handles snapshots and deltas", () => {
  assert.deepEqual(foldStreamText("", "hi"), { answer: "hi", emit: "hi" });
  assert.deepEqual(foldStreamText("hi", "hi there"), { answer: "hi there", emit: " there" });
  assert.deepEqual(foldStreamText("hi there", "hi"), { answer: "hi there", emit: null });
});

test("parseArgs takes the prompt then the ticket id", () => {
  assert.deepEqual(parseArgs(["draft a reply", "10100"]), { prompt: "draft a reply", id: 10100 });
});
test("wsConnect performs a real WebSocket handshake and exchanges frames", async () => {
  const { createServer } = await import("node:http");
  const { createHash } = await import("node:crypto");

  const received = [];
  const server = createServer();
  server.on("upgrade", (req, socket) => {
    assert.equal(req.headers.upgrade, "websocket");
    assert.equal(req.headers.origin, "https://example.test");
    assert.ok(req.headers["sec-websocket-key"], "client sent a key");
    const accept = createHash("sha1")
      .update(req.headers["sec-websocket-key"] + "258EAFA5-E914-47DA-95CA-C5AB0DC85B11")
      .digest("base64");
    socket.write(
      "HTTP/1.1 101 Switching Protocols\r\nUpgrade: websocket\r\nConnection: Upgrade\r\n" +
        `Sec-WebSocket-Accept: ${accept}\r\n\r\n`,
    );
    const push = createFrameParser((_fin, _opcode, payload) => {
      received.push(payload.toString());
      socket.write(serverFrame("world"));
    });
    socket.on("data", (chunk) => push(chunk));
  });
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const { port } = server.address();

  try {
    const ws = await wsConnect(`ws://127.0.0.1:${port}/chathub`, { Origin: "https://example.test" });
    const reply = await new Promise((resolve, reject) => {
      ws.on("message", resolve);
      ws.on("error", reject);
      ws.send("hello");
    });

    assert.deepEqual(received, ["hello"]);
    assert.equal(reply, "world");
    ws.close();
  } finally {
    server.close();
  }
});

test("chooseRefreshToken prefers env, then the rotated saved token, then config", () => {
  assert.equal(
    chooseRefreshToken({ env: { M365_REFRESH_TOKEN: "e" }, saved: { refreshToken: "s" }, config: { refreshToken: "c" } }),
    "e",
  );
  assert.equal(chooseRefreshToken({ saved: { refreshToken: "s" }, config: { refreshToken: "c" } }), "s");
  assert.equal(chooseRefreshToken({ config: { refreshToken: "c" } }), "c");
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

test("redactPII leaves dates, ticket ids and references alone", () => {
  const text = "Created 2026-08-01T10:30:00Z, ticket #10100, ref INC0012345";
  assert.equal(redactPII(text), text);
});

test("buildPrompt prepends an optional system prompt", () => {
  const prompt = buildPrompt("DATA", "do it", "You are terse.");
  assert.ok(prompt.startsWith("You are terse.\n\n"), "system prompt first");
  assert.ok(prompt.includes("<<<CONTEXT\nDATA\nCONTEXT>>>"));
  assert.ok(prompt.endsWith("do it"));
});

test("buildPrompt omits an empty or missing system prompt", () => {
  assert.ok(!buildPrompt("DATA", "do it", "").startsWith("\n"));
  assert.ok(buildPrompt("DATA", "do it").startsWith("The text between the CONTEXT"));
});
