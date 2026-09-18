import test from "node:test";
import assert from "node:assert/strict";

import { M365Session } from "./index.js";

const TOKEN = [
  Buffer.from(JSON.stringify({ alg: "none" })).toString("base64url"),
  Buffer.from(JSON.stringify({ oid: "OID-1", tid: "TID-9" })).toString("base64url"),
  "sig",
].join(".");

function jsonResponse(body, { status = 200 } = {}) {
  return {
    ok: status >= 200 && status < 300,
    status,
    json: async () => body,
    text: async () => JSON.stringify(body),
  };
}

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

test("selectConversation switches to an existing conversation and persists it", () => {
  const store = fakeStore({ sessionId: "s1", conversationId: "c1", turnCount: 3 });
  const s = new M365Session({ store });

  s.selectConversation("other");

  assert.equal(s.conversationId, "other");
  assert.equal(s.turnCount, 0, "turn count is unknown for a conversation we did not start");
  assert.equal(s.sessionId, "s1", "the transport session id is kept");
  assert.equal(store.writes.at(-1).conversationId, "other");
});

test("listConversations uses the session token and returns the chats", async () => {
  const calls = [];
  const fetchImpl = async (url, init) => {
    calls.push({ url, init });
    return jsonResponse({ chats: [{ conversationId: "c9", chatName: "Older chat" }] });
  };
  const store = fakeStore(null);
  const s = new M365Session({ store, token: TOKEN, fetchImpl });

  const result = await s.listConversations();

  assert.equal(calls[0].init.headers.Authorization, `Bearer ${TOKEN}`);
  assert.deepEqual(result.chats, [{ conversationId: "c9", chatName: "Older chat" }]);
});

test("deleteConversation defaults to the current conversation", async () => {
  const calls = [];
  const fetchImpl = async (url, init) => {
    calls.push({ url, init });
    return jsonResponse({ result: { value: "Success" } });
  };
  const s = new M365Session({ store: fakeStore({ sessionId: "s1", conversationId: "c1", turnCount: 1 }), token: TOKEN, fetchImpl });

  await s.deleteConversation();

  assert.equal(calls[0].url, "https://substrate.office.com/m365Copilot/DeleteConversation");
  assert.deepEqual(JSON.parse(calls[0].init.body).conversationIdsToDelete, ["c1"]);
});

test("getConversation loads the current conversation by default", async () => {
  const calls = [];
  const fetchImpl = async (url, init) => {
    calls.push({ url, init });
    return jsonResponse({ messages: [] });
  };
  const s = new M365Session({ store: fakeStore({ sessionId: "s1", conversationId: "c1", turnCount: 1 }), token: TOKEN, fetchImpl });

  await s.getConversation();

  const request = JSON.parse(new URL(calls[0].url).searchParams.get("request"));
  assert.equal(request.conversationId, "c1");
});

test("uploadFile uses the current conversation unless one is given", async () => {
  const calls = [];
  const fetchImpl = async (url, init) => {
    calls.push({ url, init });
    return jsonResponse({ docId: "d1" });
  };
  const s = new M365Session({ store: fakeStore({ sessionId: "s1", conversationId: "c1", turnCount: 1 }), token: TOKEN, fetchImpl });

  await s.uploadFile({ data: Buffer.from("x"), fileName: "a.txt" });
  await s.uploadFile({ conversationId: "c2", data: Buffer.from("x") });

  assert.equal(calls[0].init.body.get("conversationId"), "c1");
  assert.equal(calls[1].init.body.get("conversationId"), "c2");
});

test("uploadImages uploads a batch for the current conversation", async () => {
  const calls = [];
  const fetchImpl = async (url, init) => {
    calls.push(init);
    return jsonResponse({ docId: `d${calls.length}` });
  };
  const s = new M365Session({ store: fakeStore({ sessionId: "s1", conversationId: "c1", turnCount: 1 }), token: TOKEN, fetchImpl });

  const results = await s.uploadImages([{ data: Buffer.from("a") }, { data: Buffer.from("b") }]);

  assert.equal(calls.length, 2);
  assert.equal(calls[0].body.get("conversationId"), "c1");
  assert.deepEqual(results.map((r) => r.docId), ["d1", "d2"]);
});

test("uploadImages on a session enforces the per-message limit", async () => {
  const fetchImpl = async () => jsonResponse({ docId: "d" });
  const s = new M365Session({ store: fakeStore(null), token: TOKEN, fetchImpl });
  const four = Array.from({ length: 4 }, (_, i) => ({ data: Buffer.from(String(i)) }));

  await assert.rejects(() => s.uploadImages(four), /at most 3 images/);
});

test("the public API exposes the conversation REST helpers", async () => {
  const mod = await import("./index.js");
  for (const name of ["listConversations", "getConversation", "deleteConversation", "uploadFile"]) {
    assert.equal(typeof mod[name], "function", `${name} is exported`);
  }
});

test("the public API exposes the image limit and the batch upload", async () => {
  const mod = await import("./index.js");

  assert.equal(mod.MAX_IMAGES_PER_MESSAGE, 3);
  assert.equal(typeof mod.uploadImages, "function");
});

test("the public API exposes the attachment helpers", async () => {
  const mod = await import("./index.js");
  for (const name of ["isImage", "planImageBatches", "planAttachmentBatches", "renderAttachmentManifest"]) {
    assert.equal(typeof mod[name], "function", `${name} is exported`);
  }
});

test("the public API exposes the document-upload helpers", async () => {
  const mod = await import("./index.js");
  for (const name of ["uploadFileToCopilot", "toFileAnnotations", "spoId"]) {
    assert.equal(typeof mod[name], "function", `${name} is exported`);
  }
});

test("a temporary session does not resume or persist the saved conversation", () => {
  const store = fakeStore({ sessionId: "s1", conversationId: "c1", turnCount: 3 });

  const s = new M365Session({ temporary: true, store });

  assert.notEqual(s.conversationId, "c1", "must not resume the saved conversation");
  assert.equal(s.turnCount, 0);
  assert.equal(store.writes.length, 0, "temporary chats must not touch the saved session");
});

test("a temporary session records that it is temporary", () => {
  const s = new M365Session({ temporary: true, store: fakeStore(null) });
  assert.equal(s.temporary, true);
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