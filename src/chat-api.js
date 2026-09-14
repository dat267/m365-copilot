// M365 Copilot *history / file* REST APIs (the non-WebSocket half of the web
// client). Mined from a captured copilot.cloud.microsoft session; the protocol
// notes live in the sibling proxy repo's docs/m365-copilot-api.md.
//
// These calls reuse the SAME Sydney token as the chat WebSocket. Unlike the
// socket (token in the URL query) they take it as a normal `Authorization`
// header, plus an `X-AnchorMailbox` derived from the token's oid/tid.
//
// Everything takes an injectable `fetchImpl` so the request shape is testable
// without the network.

import { decodeJwt } from "./auth.js";

const BASE_URL = process.env.M365_SUBSTRATE_BASE || "https://substrate.office.com";
const SCENARIO = "OfficeWebIncludedCopilot";

/** `Oid:<oid>@<tid>` — M365's mailbox-routing header. */
export function anchorMailbox(token) {
  const { oid, tid } = decodeJwt(token);
  return `Oid:${oid}@${tid}`;
}

function jsonHeaders({ token, scenario, locale, clientRequestId }) {
  return {
    "Authorization": `Bearer ${token}`,
    "Content-Type": "application/json",
    "Accept-Language": locale,
    "X-ClientRequestId": clientRequestId,
    "X-Scenario": scenario,
    "X-AnchorMailbox": anchorMailbox(token),
    "Origin": "https://copilot.cloud.microsoft",
  };
}

async function finish(res, operation) {
  const text = await res.text();
  if (!res.ok) {
    let message = "";
    try {
      const j = JSON.parse(text);
      message =
        j?.result?.message ||
        j?.message ||
        j?.error?.message ||
        (typeof j?.error === "string" ? j.error : "") ||
        "";
    } catch {
      // non-JSON error body — the status is all we have
    }
    throw new Error(`M365 ${operation} failed (HTTP ${res.status})${message ? `: ${message}` : ""}`);
  }
  try {
    return JSON.parse(text);
  } catch {
    return {};
  }
}

/** GET {base}/m365Copilot/GetChats — the conversation list. */
export async function listConversations({ token, fetchImpl = fetch, baseUrl = BASE_URL, locale = "en-gb", scenario = SCENARIO, clientRequestId = crypto.randomUUID() } = {}) {
  const request = {
    source: "officeweb",
    traceId: crypto.randomUUID(),
    threadType: "bizchat",
    MaxReturnedChatsCount: 100,
    mergeWorkWebChats: true,
  };
  return getJson({ token, fetchImpl, baseUrl, locale, scenario, clientRequestId }, "/m365Copilot/GetChats", request, "GetChats");
}

/** GET {base}/m365Copilot/GetConversation — full message history of one chat. */
export async function getConversation({ token, conversationId, fetchImpl = fetch, baseUrl = BASE_URL, locale = "en-gb", scenario = SCENARIO, clientRequestId = crypto.randomUUID() } = {}) {
  const request = { conversationId, source: "officeweb", traceId: crypto.randomUUID() };
  return getJson({ token, fetchImpl, baseUrl, locale, scenario, clientRequestId }, "/m365Copilot/GetConversation", request, "GetConversation");
}

/** POST {base}/m365Copilot/DeleteConversation — batch delete (accepts many ids). */
export async function deleteConversation({ token, conversationIds, conversationId, fetchImpl = fetch, baseUrl = BASE_URL, locale = "en-gb", scenario = SCENARIO, clientRequestId = crypto.randomUUID() } = {}) {
  const ids = conversationIds ?? (conversationId ? [conversationId] : []);
  const body = { conversationIdsToDelete: ids, source: "officeweb", traceId: crypto.randomUUID() };
  const url = `${baseUrl}/m365Copilot/DeleteConversation`;
  const res = await fetchImpl(url, {
    method: "POST",
    headers: jsonHeaders({ token, scenario, locale, clientRequestId }),
    body: JSON.stringify(body),
  });
  return finish(res, "DeleteConversation");
}

/**
 * POST {base}/m365Copilot/UploadFile — upload one file/image and get back the
 * `docId` that a later chat turn attaches as a file annotation.
 *
 * @param {Buffer|Uint8Array|string} data raw bytes, or an existing `data:` URL
 * @param {string} [mimeType] required when `data` is raw bytes (default octet-stream)
 */
export async function uploadFile({
  token,
  conversationId,
  data,
  fileName,
  mimeType = "application/octet-stream",
  scenario = "UploadImage",
  fetchImpl = fetch,
  baseUrl = BASE_URL,
  locale = "en-gb",
  xScenario = SCENARIO,
} = {}) {
  const form = new FormData();
  form.append("scenario", scenario);
  form.append("conversationId", conversationId);
  form.append("FileBase64", toDataUrl(data, mimeType));
  if (fileName) form.append("FileName", fileName);

  const res = await fetchImpl(`${baseUrl}/m365Copilot/UploadFile`, {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${token}`,
      "Accept-Language": locale,
      "X-AnchorMailbox": anchorMailbox(token),
      "X-Scenario": xScenario,
      "X-Variants": "feature.EnableImageSupportInUploadFile",
      "Origin": "https://copilot.cloud.microsoft",
    },
    body: form,
  });
  return finish(res, "UploadFile");
}

/**
 * Observed M365 limit: at most 3 images can be attached to one message.
 *
 * The real cap is enforced server-side and is feature-flag driven (the web
 * client only ships the *message* for a count error, and the composer bundle
 * that supplies the number was not observable), so treat this as a default
 * guard rather than a protocol constant — override `limit` if it changes.
 */
export const MAX_IMAGES_PER_MESSAGE = 3;

/**
 * Upload the images for ONE message and return their results in order.
 * Sequential by design: M365 throttles on concurrent requests (AGENTS.md #1).
 *
 * @param {Array<{data: Buffer|Uint8Array|string, fileName?: string, mimeType?: string}>} images
 * @param {number} [limit] defaults to MAX_IMAGES_PER_MESSAGE
 */
export async function uploadImages({ images = [], limit = MAX_IMAGES_PER_MESSAGE, ...opts } = {}) {
  if (images.length > limit) {
    throw new Error(
      `M365 accepts at most ${limit} images per message (got ${images.length}). ` +
        `Send ${limit} here and put the rest in a follow-up message.`,
    );
  }
  const results = [];
  for (const image of images) {
    results.push(await uploadFile({ ...opts, ...image }));
  }
  return results;
}

function toDataUrl(data, mimeType) {
  if (typeof data === "string") {
    return data.startsWith("data:") ? data : `data:${mimeType};base64,${Buffer.from(data).toString("base64")}`;
  }
  return `data:${mimeType};base64,${Buffer.from(data).toString("base64")}`;
}

async function getJson({ token, fetchImpl, baseUrl, locale, scenario, clientRequestId }, path, request, operation) {
  const url = `${baseUrl}${path}?request=${encodeURIComponent(JSON.stringify(request))}`;
  const res = await fetchImpl(url, {
    method: "GET",
    headers: jsonHeaders({ token, scenario, locale, clientRequestId }),
  });
  return finish(res, operation);
}
