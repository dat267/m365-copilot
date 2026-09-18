import test from "node:test";
import assert from "node:assert/strict";

import {
  listConversations,
  getConversation,
  deleteConversation,
  uploadFile,
  uploadImages,
  MAX_IMAGES_PER_MESSAGE,
} from "./chat-api.js";

// A structurally valid (unsigned) token: the code only decodes the payload.
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

test("listConversations GETs GetChats and returns the chats", async () => {
  const calls = [];
  const fetchImpl = async (url, init) => {
    calls.push({ url, init });
    return jsonResponse({ chats: [{ conversationId: "c1", chatName: "Hello" }] });
  };

  const result = await listConversations({ token: TOKEN, fetchImpl });

  assert.equal(calls.length, 1);
  const { url, init } = calls[0];
  assert.equal(init.method, "GET");
  assert.ok(url.startsWith("https://substrate.office.com/m365Copilot/GetChats?request="));
  const request = JSON.parse(new URL(url).searchParams.get("request"));
  assert.equal(request.MaxReturnedChatsCount, 100);
  assert.deepEqual(result.chats, [{ conversationId: "c1", chatName: "Hello" }]);
});

test("getConversation GETs GetConversation for one id and returns the messages", async () => {
  const calls = [];
  const fetchImpl = async (url, init) => {
    calls.push({ url, init });
    return jsonResponse({ messages: [{ author: "user", text: "hi" }] });
  };

  const result = await getConversation({ token: TOKEN, conversationId: "c-42", fetchImpl });

  const { url, init } = calls[0];
  assert.equal(init.method, "GET");
  assert.ok(url.startsWith("https://substrate.office.com/m365Copilot/GetConversation?request="));
  assert.equal(JSON.parse(new URL(url).searchParams.get("request")).conversationId, "c-42");
  assert.deepEqual(result.messages, [{ author: "user", text: "hi" }]);
});

test("requests are authenticated and routed by anchor mailbox", async () => {
  const calls = [];
  const fetchImpl = async (url, init) => {
    calls.push({ url, init });
    return jsonResponse({ messages: [] });
  };

  await getConversation({ token: TOKEN, conversationId: "c-42", fetchImpl });

  const { headers } = calls[0].init;
  assert.equal(headers["Authorization"], `Bearer ${TOKEN}`);
  assert.equal(headers["X-AnchorMailbox"], "Oid:OID-1@TID-9");
  assert.equal(headers["X-Scenario"], "OfficeWebIncludedCopilot");
  assert.ok(headers["X-ClientRequestId"]);
});

test("deleteConversation POSTs DeleteConversation with the ids", async () => {
  const calls = [];
  const fetchImpl = async (url, init) => {
    calls.push({ url, init });
    return jsonResponse({ result: { value: "Success" } });
  };

  const result = await deleteConversation({ token: TOKEN, conversationIds: ["c1", "c2"], fetchImpl });

  const { url, init } = calls[0];
  assert.equal(url, "https://substrate.office.com/m365Copilot/DeleteConversation");
  assert.equal(init.method, "POST");
  const body = JSON.parse(init.body);
  assert.deepEqual(body.conversationIdsToDelete, ["c1", "c2"]);
  assert.ok(body.traceId);
  assert.equal(result.result.value, "Success");
});

test("deleteConversation accepts a single id", async () => {
  const calls = [];
  const fetchImpl = async (url, init) => {
    calls.push({ url, init });
    return jsonResponse({ result: { value: "Success" } });
  };

  await deleteConversation({ token: TOKEN, conversationId: "only", fetchImpl });

  assert.deepEqual(JSON.parse(calls[0].init.body).conversationIdsToDelete, ["only"]);
});

test("uploadFile POSTs multipart to UploadFile and returns the docId", async () => {
  const calls = [];
  const fetchImpl = async (url, init) => {
    calls.push({ url, init });
    return jsonResponse({
      fileName: "pic.png",
      docId: "0-ea-d1-90f960460d7c7569ac24c34483c0031a",
      fileUrl: "https://x/pic.png",
      result: { value: "Success" },
    });
  };

  const result = await uploadFile({
    token: TOKEN,
    conversationId: "c-42",
    data: Buffer.from("PNGDATA"),
    fileName: "pic.png",
    mimeType: "image/png",
    fetchImpl,
  });

  const { url, init } = calls[0];
  assert.equal(url, "https://substrate.office.com/m365Copilot/UploadFile");
  assert.equal(init.method, "POST");
  assert.ok(init.body instanceof FormData, "body is multipart FormData");
  assert.equal(init.body.get("scenario"), "UploadImage");
  assert.equal(init.body.get("conversationId"), "c-42");
  assert.equal(init.body.get("FileName"), "pic.png");
  assert.equal(
    init.body.get("FileBase64"),
    `data:image/png;base64,${Buffer.from("PNGDATA").toString("base64")}`,
  );
  assert.equal(result.docId, "0-ea-d1-90f960460d7c7569ac24c34483c0031a");
});

test("uploadFile sends the image-upload headers and lets fetch set the content type", async () => {
  const calls = [];
  const fetchImpl = async (url, init) => {
    calls.push({ url, init });
    return jsonResponse({ docId: "d" });
  };

  await uploadFile({ token: TOKEN, conversationId: "c", data: Buffer.from("x"), fetchImpl });

  const { headers } = calls[0].init;
  assert.equal(headers["Authorization"], `Bearer ${TOKEN}`);
  assert.equal(headers["X-AnchorMailbox"], "Oid:OID-1@TID-9");
  assert.equal(headers["X-Variants"], "feature.EnableImageSupportInUploadFile");
  assert.equal(headers["Content-Type"], undefined, "must not override the multipart boundary");
});

test("a failed request throws with the HTTP status and the server message", async () => {
  const fetchImpl = async () =>
    jsonResponse({ result: { value: "Error", message: "Chat not found" } }, { status: 404 });

  await assert.rejects(
    () => getConversation({ token: TOKEN, conversationId: "nope", fetchImpl }),
    /M365 GetConversation failed \(HTTP 404\).*Chat not found/,
  );
});

test("a failed upload throws rather than returning a body", async () => {
  const fetchImpl = async () => jsonResponse({ error: "nope" }, { status: 417 });

  await assert.rejects(
    () => uploadFile({ token: TOKEN, conversationId: "c", data: Buffer.from("x"), fetchImpl }),
    /HTTP 417/,
  );
});

test("uploadImages uploads each image and returns the results in order", async () => {
  const calls = [];
  const fetchImpl = async (url, init) => {
    calls.push(init);
    return jsonResponse({ docId: `doc-${calls.length}` });
  };
  const images = [
    { data: Buffer.from("a"), fileName: "a.png" },
    { data: Buffer.from("b"), fileName: "b.png" },
    { data: Buffer.from("c"), fileName: "c.png" },
  ];

  const results = await uploadImages({ token: TOKEN, conversationId: "c1", images, fetchImpl });

  assert.equal(calls.length, 3);
  assert.deepEqual(results.map((r) => r.docId), ["doc-1", "doc-2", "doc-3"]);
  assert.equal(calls[0].body.get("conversationId"), "c1");
  assert.equal(calls[0].body.get("FileName"), "a.png");
});

test("the observed per-message image limit is three", () => {
  assert.equal(MAX_IMAGES_PER_MESSAGE, 3);
});

test("uploadImages refuses more than the limit without uploading anything", async () => {
  let calls = 0;
  const fetchImpl = async () => {
    calls++;
    return jsonResponse({ docId: "d" });
  };
  const images = Array.from({ length: 4 }, (_, i) => ({ data: Buffer.from(String(i)), fileName: `${i}.png` }));

  await assert.rejects(
    () => uploadImages({ token: TOKEN, conversationId: "c", images, fetchImpl }),
    /at most 3 images/,
  );
  assert.equal(calls, 0, "must not upload any image once the limit is exceeded");
});

test("uploadImages accepts an override for when M365 changes the cap", async () => {
  const fetchImpl = async () => jsonResponse({ docId: "d" });
  const images = Array.from({ length: 4 }, (_, i) => ({ data: Buffer.from(String(i)) }));

  const results = await uploadImages({ token: TOKEN, conversationId: "c", images, fetchImpl, limit: 4 });

  assert.equal(results.length, 4);
});
