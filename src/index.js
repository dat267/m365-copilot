// Public API.
//
//   import { ask, M365Session } from "m365-ask";
//
//   const text = await ask("Summarize this: ...");     // one-shot
//
//   const s = new M365Session();                       // multi-turn
//   for await (const delta of await s.chat("hi")) process.stdout.write(delta);

import { getToken, getTokenSilent, loginInteractive, refreshTokenGrant, decodeJwt } from "./auth.js";
import { defaultSessionStore, loadSession, saveSession, resolveSessionIds } from "./session-store.js";
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
  constructor({ model = "m365-copilot", fresh = false, sessionId, conversationId, store = defaultSessionStore() } = {}) {
    const ids = resolveSessionIds({ persisted: loadSession(store), fresh, sessionId, conversationId });
    this.model = model;
    this.sessionId = ids.sessionId;
    this.conversationId = ids.conversationId;
    this.turnCount = ids.turnCount;
    this._client = null;
    this._store = store;
    saveSession(store, this);
  }

  /**
   * Send a message. Returns an async-iterable stream (also with `.fullText`,
   * `.messageType`, `.throttle`, `.scores` getters after it completes).
   */
  async chat(text, { signal } = {}) {
    const token = await getToken();
    const make = () =>
      new CopilotSession({ sessionId: this.sessionId, conversationId: this.conversationId, turnCount: this.turnCount });
    this._client ??= make();
    try {
      return this._persist(await this._client.chat(token, text, this.model, signal));
    } catch {
      // Stale/failed socket — reconnect with the SAME ids and retry once.
      this._client = make();
      return this._persist(await this._client.chat(token, text, this.model, signal));
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
