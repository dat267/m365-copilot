// Public API.
//
//   import { ask, M365Session } from "m365-copilot";
//
//   const text = await ask("Summarize this: ...");     // one-shot
//
//   const s = new M365Session();                       // multi-turn
//   for await (const delta of await s.chat("hi")) process.stdout.write(delta);

import { getToken, getTokenSilent, loginInteractive, refreshTokenGrant, decodeJwt } from "./auth.js";
import { defaultSessionStore, loadSession, saveSession, resolveSessionIds } from "./session-store.js";
import {
  listConversations as apiListConversations,
  getConversation as apiGetConversation,
  deleteConversation as apiDeleteConversation,
  uploadFile as apiUploadFile,
  uploadImages as apiUploadImages,
} from "./chat-api.js";
import {
  CopilotSession,
  foldStreamText,
  getToneForModel,
  getAvailableModels,
} from "./client.js";

/**
 * A high-level conversation: handles auth and reconnects, reusing one
 * conversationId so M365 keeps server-side context (and doesn't burn a new
 * conversation against the account's thread budget each turn).
 *
 * The conversation is PERSISTED (config dir `session.json`) and resumed by
 * default, so separate runs share it. Start a new one explicitly with
 * `newConversation()`, `{ fresh: true }`, or `M365_NO_SESSION_PERSIST=1` to
 * disable persistence entirely.
 */
export class M365Session {
  constructor({ model = "m365-copilot", fresh = false, sessionId, conversationId, temporary = false, store = defaultSessionStore(), token, fetchImpl, baseUrl } = {}) {
    // A temporary chat is ephemeral: never resume the saved conversation and
    // never persist this one (disabling the store makes saveSession a no-op).
    const persisted = temporary ? null : loadSession(store);
    const ids = resolveSessionIds({ persisted, fresh: fresh || temporary, sessionId, conversationId });
    this.model = model;
    this.sessionId = ids.sessionId;
    this.conversationId = ids.conversationId;
    this.turnCount = ids.turnCount;
    this.temporary = temporary;
    // A conversation we did not create in this instance is "resumed": its next
    // turn is never the start of a session (proxy docs §8).
    this.resumed = !fresh && !temporary && (persisted != null || conversationId != null);
    this._client = null;
    this._store = temporary ? { enabled: false, path: store.path } : store;
    // Overridable for tests / custom transports; otherwise real auth + fetch.
    this._token = token;
    this._fetchImpl = fetchImpl;
    this._baseUrl = baseUrl;
    saveSession(this._store, this);
  }

  /** Resolve the REST-call options: per-call > instance > real defaults. */
  async _apiOptions(opts = {}) {
    const merged = {
      ...opts,
      token: opts.token ?? this._token ?? (await getToken()),
      fetchImpl: opts.fetchImpl ?? this._fetchImpl ?? fetch,
    };
    const baseUrl = opts.baseUrl ?? this._baseUrl;
    if (baseUrl !== undefined) merged.baseUrl = baseUrl;
    return merged;
  }

  /** List this account's conversations, newest first. */
  async listConversations(opts) {
    return apiListConversations(await this._apiOptions(opts));
  }

  /** Full message history of one conversation (defaults to the current one). */
  async getConversation({ conversationId = this.conversationId, ...opts } = {}) {
    return apiGetConversation(await this._apiOptions({ ...opts, conversationId }));
  }

  /** Delete conversations (defaults to the current one). Pass `conversationIds` to batch. */
  async deleteConversation({ conversationIds, conversationId = this.conversationId, ...opts } = {}) {
    const ids = conversationIds ?? (conversationId ? [conversationId] : []);
    return apiDeleteConversation(await this._apiOptions({ ...opts, conversationIds: ids }));
  }

  /** Upload a file/image; returns `{ docId, fileUrl, ... }`. */
  async uploadFile({ conversationId = this.conversationId, ...rest } = {}) {
    return apiUploadFile(await this._apiOptions({ ...rest, conversationId }));
  }

  /**
   * Upload the images for ONE message (at most MAX_IMAGES_PER_MESSAGE, enforced).
   * @param {Array<{data: Buffer|Uint8Array|string, fileName?: string, mimeType?: string}>} images
   */
  async uploadImages(images, { conversationId = this.conversationId, ...rest } = {}) {
    return apiUploadImages(await this._apiOptions({ ...rest, conversationId, images }));
  }

  /**
   * Send a message. Returns an async-iterable stream (also with `.fullText`,
   * `.messageType`, `.throttle`, `.scores` getters after it completes).
   *
   * `attachments` are `uploadFile()`/`uploadImages()` results — they are attached
   * to the turn as image annotations (uploading alone does not attach them).
   */
  async chat(text, { signal, attachments } = {}) {
    const token = await getToken();
    const make = () =>
      new CopilotSession({
        sessionId: this.sessionId,
        conversationId: this.conversationId,
        turnCount: this.turnCount,
        resumed: this.resumed,
        temporary: this.temporary,
      });
    this._client ??= make();
    try {
      return this._persist(await this._client.chat(token, text, this.model, signal, attachments));
    } catch {
      // Stale/failed socket — reconnect with the SAME ids and retry once.
      this._client = make();
      return this._persist(await this._client.chat(token, text, this.model, signal, attachments));
    }
  }

  _persist(stream) {
    this.turnCount = this._client.turnCount;
    saveSession(this._store, this);
    return stream;
  }

  /** Drop context and start a fresh M365 conversation. */
  newConversation() {
    const ids = resolveSessionIds({ fresh: true });
    this.sessionId = ids.sessionId;
    this.conversationId = ids.conversationId;
    this.turnCount = 0;
    this.resumed = false;
    this._client = null;
    saveSession(this._store, this);
  }

  /**
   * Resume an EXISTING M365 conversation (e.g. one from `listConversations()`).
   * The next `chat()` reconnects to it and marks the turn as not-start-of-session.
   */
  selectConversation(conversationId) {
    this.conversationId = conversationId;
    this.turnCount = 0;
    this.resumed = true;
    this._client = null;
    saveSession(this._store, this);
  }
}

/**
 * One-shot convenience: send a prompt, return the full answer as a string.
 * Throws on Disengaged / empty responses so callers don't silently get "".
 *
 * @param {string} text
 * @param {{ model?: string, session?: M365Session, fresh?: boolean, signal?: AbortSignal }} [opts]
 *   `fresh` starts a new conversation (default: resume the persisted one).
 */
export async function ask(text, opts = {}) {
  const session = opts.session ?? new M365Session(opts);
  const stream = await session.chat(text, opts);
  for await (const _ of stream) {
    // drain to completion
  }
  if (stream.messageType === "Disengaged") {
    throw new Error(
      "M365 disengaged (its safety filter) — rephrase the prompt or start a new conversation.",
    );
  }
  if (!stream.hasContent) {
    throw new Error(
      "M365 returned no content (throttled, account degraded, or transient) — retry later.",
    );
  }
  return stream.fullText;
}

export {
  CopilotSession,
  getToken,
  getTokenSilent,
  loginInteractive,
  refreshTokenGrant,
  decodeJwt,
  foldStreamText,
  getToneForModel,
  getAvailableModels,
};

// Conversation history + file-upload REST helpers (usable without a session).
export {
  listConversations,
  getConversation,
  deleteConversation,
  uploadFile,
  uploadImages,
  MAX_IMAGES_PER_MESSAGE,
  anchorMailbox,
} from "./chat-api.js";

// Attachment planning/rendering (image batching under the per-message cap).
export { isImage, planImageBatches, planAttachmentBatches, renderAttachmentManifest } from "./attachments.js";

// Document (file) attachment via OneDrive + `LocalFile` annotations.
export { uploadFileToCopilot, toFileAnnotations, spoId } from "./graph-upload.js";
