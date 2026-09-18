import test from "node:test";
import assert from "node:assert/strict";

import { resolveSessionIds, loadSession, saveSession } from "./session-store.js";

test("resolveSessionIds reuses persisted ids by default", () => {
  const ids = resolveSessionIds({
    persisted: { sessionId: "s1", conversationId: "c1", turnCount: 3 },
    newId: () => "NEW",
  });

  assert.deepEqual(ids, { sessionId: "s1", conversationId: "c1", turnCount: 3 });
});
test("resolveSessionIds mints fresh ids for a new conversation", () => {
  let n = 0;
  const ids = resolveSessionIds({
    persisted: { sessionId: "s1", conversationId: "c1", turnCount: 3 },
    fresh: true,
    newId: () => `new-${++n}`,
  });

  assert.deepEqual(ids, { sessionId: "new-1", conversationId: "new-2", turnCount: 0 });
});

test("resolveSessionIds prefers explicit ids over persisted ones", () => {
  const ids = resolveSessionIds({
    persisted: { sessionId: "s1", conversationId: "c1", turnCount: 3 },
    conversationId: "explicit",
    newId: () => "NEW",
  });

  assert.equal(ids.sessionId, "s1");
  assert.equal(ids.conversationId, "explicit");
  assert.equal(ids.turnCount, 0, "an explicitly supplied conversation has an unknown turn count");
});

test("loadSession reads the persisted session", () => {
  const store = {
    enabled: true,
    path: "/tmp/session.json",
    readFile: (p) => {
      assert.equal(p, "/tmp/session.json");
      return JSON.stringify({ sessionId: "s1", conversationId: "c1", turnCount: 2 });
    },
  };

  assert.deepEqual(loadSession(store), { sessionId: "s1", conversationId: "c1", turnCount: 2 });
});

test("loadSession returns null when persistence is disabled", () => {
  const store = { enabled: false, path: "/tmp/session.json", readFile: () => "{}" };
  assert.equal(loadSession(store), null);
});

test("loadSession returns null when the file is missing or invalid", () => {
  const missing = { enabled: true, path: "/tmp/x", readFile: () => { throw new Error("ENOENT"); } };
  const invalid = { enabled: true, path: "/tmp/x", readFile: () => "not json" };

  assert.equal(loadSession(missing), null);
  assert.equal(loadSession(invalid), null);
});

test("saveSession writes the ids and turn count", () => {
  let written;
  const store = { enabled: true, path: "/tmp/session.json", writeFile: (p, data) => { written = [p, JSON.parse(data)]; } };

  saveSession(store, { sessionId: "s1", conversationId: "c1", turnCount: 4 });

  assert.deepEqual(written, ["/tmp/session.json", { sessionId: "s1", conversationId: "c1", turnCount: 4 }]);
});

test("saveSession does nothing when persistence is disabled", () => {
  let called = false;
  const store = { enabled: false, path: "/tmp/session.json", writeFile: () => { called = true; } };

  saveSession(store, { sessionId: "s1", conversationId: "c1", turnCount: 4 });

  assert.equal(called, false);
});
