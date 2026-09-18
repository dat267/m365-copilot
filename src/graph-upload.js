// File (document) attachment for M365 Copilot — the OneDrive path.
//
// Images go through POST /m365Copilot/UploadFile. Files do NOT: the web client
// uploads them into the user's OneDrive "Microsoft Copilot Chat Files"
// (the `copilotuploads` special folder) via Graph, then attaches them to the
// chat message as a `LocalFile` annotation.
//
// Captured live (Playwright + WS capture). The sequence is:
//   GET  /me/drive/special/copilotuploads                       -> driveId
//   POST /me/drive/special/copilotuploads:/<name>:/createUploadSession
//   PUT  <uploadUrl>                                            -> driveItem
//   (attach) messageAnnotations += { id: SPO_<...>_<itemId>,
//                                    text: <name>, url: <webUrl>,
//                                    messageAnnotationType: "LocalFile" }
//
// The SPO_ id is `SPO_` + base64url("<siteId>,<webId>,<listId>") + "_" + itemId,
// where the three GUIDs are the ones packed into the `b!` drive id.
//
// Graph needs its own token (scope https://graph.microsoft.com/.default) — the
// Sydney token used for chat will not work here.

import { getGraphToken } from "./auth.js";

const GRAPH = "https://graph.microsoft.com/v1.0";

/** One 16-byte GUID from a `b!` drive id, in .NET mixed-endian order. */
function guidFromDriveId(buf, offset) {
  const le = (i, n) => Buffer.from(buf.subarray(i, i + n)).reverse().toString("hex");
  const be = (i, n) => buf.subarray(i, i + n).toString("hex");
  return `${le(offset, 4)}-${le(offset + 4, 2)}-${le(offset + 6, 2)}-${be(offset + 8, 2)}-${be(offset + 10, 6)}`;
}

/**
 * Builds the `id` the web client puts on a `LocalFile` annotation.
 * @param {string} driveId e.g. `b!ERERESIiMzNERFVVVVVVV...` (SharePoint `b!` id)
 * @param {string} itemId  the driveItem id
 */
export function spoId(driveId, itemId) {
  const raw = Buffer.from(String(driveId).replace(/^b!/, "").replace(/-/g, "+").replace(/_/g, "/"), "base64");
  if (raw.length < 48) throw new Error(`unexpected driveId: ${driveId}`);
  const ids = [0, 16, 32].map((o) => guidFromDriveId(raw, o)).join(",");
  return `SPO_${Buffer.from(ids, "utf8").toString("base64url")}_${itemId}`;
}

/**
 * Maps uploaded drive items to the `messageAnnotations` entries that attach
 * files to a turn.
 */
export function toFileAnnotations(items = []) {
  return items
    .filter((i) => i && i.itemId && i.driveId)
    .map((i) => ({
      id: spoId(i.driveId, i.itemId),
      text: i.fileName ?? "",
      url: i.webUrl ?? "",
      messageAnnotationType: "LocalFile",
    }));
}

async function graph(token, path, { method = "GET", body, fetchImpl = fetch } = {}) {
  const res = await fetchImpl(`${GRAPH}${path}`, {
    method,
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    ...(body ? { body: JSON.stringify(body) } : {}),
  });
  const text = await res.text();
  let json;
  try {
    json = text ? JSON.parse(text) : {};
  } catch {
    json = {};
  }
  if (!res.ok) {
    throw new Error(`Graph ${method} ${path.split("?")[0]} failed (HTTP ${res.status}): ${json.error?.message ?? text.slice(0, 200)}`);
  }
  return json;
}

/**
 * Upload one file into the Copilot uploads folder and return what the chat
 * annotation needs: `{ driveId, itemId, fileName, webUrl }`.
 *
 * Single PUT upload — fine for the small documents found on tickets. Files over
 * ~250 MB would need chunked ranges.
 */
export async function uploadFileToCopilot({
  fileName,
  data,
  mimeType = "application/octet-stream",
  token,
  fetchImpl = fetch,
} = {}) {
  const graphToken = token ?? (await getGraphToken());
  const folder = await graph(graphToken, "/me/drive/special/copilotuploads", { fetchImpl });
  const driveId = folder?.parentReference?.driveId;
  if (!driveId) throw new Error("could not resolve the copilotuploads drive");

  const session = await graph(
    graphToken,
    `/me/drive/special/copilotuploads:/${encodeURIComponent(fileName)}:/createUploadSession`,
    { method: "POST", body: { item: { "@microsoft.graph.conflictBehavior": "replace" } }, fetchImpl },
  );
  if (!session.uploadUrl) throw new Error("createUploadSession returned no uploadUrl");

  const bytes = Buffer.isBuffer(data) ? data : Buffer.from(data);
  // Upload-session PUTs require a Content-Range even for a single request.
  const put = await fetchImpl(session.uploadUrl, {
    method: "PUT",
    headers: {
      "Content-Type": mimeType,
      "Content-Length": String(bytes.length),
      "Content-Range": `bytes 0-${bytes.length - 1}/${bytes.length}`,
    },
    body: bytes,
  });
  const text = await put.text();
  if (!put.ok) throw new Error(`upload failed (HTTP ${put.status}): ${text.slice(0, 200)}`);
  const item = JSON.parse(text);

  return { driveId, itemId: item.id, fileName: item.name ?? fileName, webUrl: item.webUrl ?? "" };
}
