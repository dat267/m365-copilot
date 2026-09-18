import test from "node:test";
import assert from "node:assert/strict";

import { isFirstTurnOfNewConversation, buildChatUrl, toImageAnnotations } from "./client.js";

test("the first turn of a conversation we started is a start-of-session", () => {
  assert.equal(isFirstTurnOfNewConversation(0), true);
});

test("later turns are not start-of-session", () => {
  assert.equal(isFirstTurnOfNewConversation(3), false);
});

test("a resumed conversation is never start-of-session, even with no turns sent yet", () => {
  assert.equal(isFirstTurnOfNewConversation(0, { resumed: true }), false);
});

const urlArgs = {
  oid: "OID-1",
  tid: "TID-9",
  sessionId: "SID",
  conversationId: "CID",
  token: "TOK",
  requestId: "RID",
};

test("buildChatUrl targets the ChatHub and carries the conversation id", () => {
  const url = new URL(buildChatUrl(urlArgs));

  assert.equal(url.protocol, "wss:");
  assert.equal(url.host, "substrate.office.com");
  assert.equal(url.pathname, "/m365Copilot/Chathub/OID-1@TID-9");
  assert.equal(url.searchParams.get("ConversationId"), "CID");
  assert.equal(url.searchParams.get("X-SessionId"), "SID");
  assert.equal(url.searchParams.get("access_token"), "TOK");
});

test("buildChatUrl leaves memory enabled for a normal chat", () => {
  const url = new URL(buildChatUrl(urlArgs));

  assert.equal(url.searchParams.has("disableMemory"), false);
});

test("buildChatUrl marks a temporary chat with disableMemory=1", () => {
  const url = new URL(buildChatUrl({ ...urlArgs, temporary: true }));

  assert.equal(url.searchParams.get("disableMemory"), "1");
});

test("toImageAnnotations maps an UploadFile result to the captured wire shape", () => {
  const [a] = toImageAnnotations([{ docId: "0-ea-d12-1abf", fileName: "probe-42.png", fileType: ".png" }]);

  assert.deepEqual(a, {
    id: "0-ea-d12-1abf",
    messageAnnotationMetadata: {
      "@type": "File",
      annotationType: "File",
      fileType: "png",
      fileName: "probe-42.png",
    },
    messageAnnotationType: "ImageFile",
  });
});

test("toImageAnnotations skips uploads without a docId", () => {
  assert.deepEqual(toImageAnnotations([{ fileName: "x.png" }, null]), []);
});

test("toImageAnnotations returns nothing for no uploads", () => {
  assert.deepEqual(toImageAnnotations(), []);
});
