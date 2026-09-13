// Public API.
//
//   import { ask, M365Session } from "m365-ask";
//
//   const text = await ask("Summarize this: ...");     // one-shot
//
//   const s = new M365Session();                       // multi-turn
//   for await (const delta of await s.chat("hi")) process.stdout.write(delta);

import { getToken, getTokenSilent, loginInteractive, refreshTokenGrant, decodeJwt } from "./auth.js";
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
 */
export class M365Session {
  constructor({ model = "m365-copilot" } = {}) {
    this.model = model;
    this.sessionId = crypto.randomUUID();
    this.conversationId = crypto.randomUUID();
    this._client = null;
  }

  /**
   * Send a message. Returns an async-iterable stream (also with `.fullText`,
   * `.messageType`, `.throttle`, `.scores` getters after it completes).
   */
  async chat(text, { signal } = {}) {
    const token = await getToken();
    const make = () => new CopilotSession({ sessionId: this.sessionId, conversationId: this.conversationId });
    this._client ??= make();
    try {
      return await this._client.chat(token, text, this.model, signal);
    } catch {
      // Stale/failed socket — reconnect with the SAME ids and retry once.
      this._client = make();
      return this._client.chat(token, text, this.model, signal);
    }
  }

  /** Drop context and start a fresh M365 conversation. */
  newConversation() {
    this.conversationId = crypto.randomUUID();
    this._client = null;
  }
}

/**
 * One-shot convenience: send a prompt, return the full answer as a string.
 * Throws on Disengaged / empty responses so callers don't silently get "".
 *
 * @param {string} text
 * @param {{ model?: string, session?: M365Session, signal?: AbortSignal }} [opts]
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
