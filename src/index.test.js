import test from "node:test";
import assert from "node:assert/strict";

import { M365Session } from "./index.js";

function fakeStore(persisted) {
  const writes = [];
  return {
    enabled: true,
    path: "/tmp/session.json",
    readFile: () => JSON.stringify(persisted),
    writeFile: (p, data) => writes.push(JSON.parse(data)),
    writes,
  };
}

test("M365Session resumes the persisted conversation by default", () => {
  const store = fakeStore({ sessionId: "s1", conversationId: "c1", turnCount: 3 });

  const s = new M365Session({ store });

  assert.equal(s.sessionId, "s1");
  assert.equal(s.conversationId, "c1");
  assert.equal(s.turnCount, 3);
});

test("M365Session starts fresh with fresh:true and persists the new ids", () => {
  const store = fakeStore({ sessionId: "s1", conversationId: "c1", turnCount: 3 });

  const s = new M365Session({ fresh: true, store });

  assert.notEqual(s.conversationId, "c1");
  assert.equal(s.turnCount, 0);
  assert.equal(store.writes.at(-1).conversationId, s.conversationId);
});

test("newConversation rotates the ids and persists them", () => {
  const store = fakeStore({ sessionId: "s1", conversationId: "c1", turnCount: 3 });
  const s = new M365Session({ store });

  s.newConversation();

  assert.notEqual(s.conversationId, "c1");
  assert.equal(s.turnCount, 0);
  assert.equal(store.writes.at(-1).conversationId, s.conversationId);
});

test("M365Session starts a new conversation when persistence is disabled", () => {
  const store = {
    enabled: false,
    path: "/tmp/session.json",
    readFile: () => {
      throw new Error("should not read");
    },
    writeFile: () => {},
  };

  const s = new M365Session({ store });

  assert.ok(s.conversationId);
  assert.equal(s.turnCount, 0);
});