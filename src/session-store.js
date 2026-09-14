// Persistent default conversation.
//
// Reusing one ConversationId across runs avoids burning the account's
// "conversations started" throttle (AGENTS.md principle #1). Explicitly start a
// new one with `M365Session.newConversation()`, `{ fresh: true }`, or CLI
// `--new`; disable persistence entirely with `M365_NO_SESSION_PERSIST=1`.
//
// `isStartOfSession` is true only on a conversation's first turn (proxy docs
// §8/§6), so the turn count is persisted too: a resumed conversation must send
// false.

import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { homedir } from "node:os";

const CONFIG_DIR = process.env.M365_CONFIG_DIR || join(homedir(), ".config", "m365-ask");
export const SESSION_FILE = process.env.M365_SESSION_FILE || join(CONFIG_DIR, "session.json");

/** Picks the ids to use: explicit > persisted > new. `fresh` always mints new. */
export function resolveSessionIds({ persisted, fresh = false, sessionId, conversationId, newId = () => crypto.randomUUID() } = {}) {
  if (fresh) return { sessionId: newId(), conversationId: newId(), turnCount: 0 };
  return {
    sessionId: sessionId ?? persisted?.sessionId ?? newId(),
    conversationId: conversationId ?? persisted?.conversationId ?? newId(),
    turnCount: sessionId || conversationId ? 0 : persisted?.turnCount ?? 0,
  };
}

export function loadSession(store) {
  if (!store.enabled) return null;
  try {
    return JSON.parse(store.readFile(store.path));
  } catch {
    return null;
  }
}

export function saveSession(store, { sessionId, conversationId, turnCount }) {
  if (!store.enabled) return;
  try {
    store.writeFile(store.path, JSON.stringify({ sessionId, conversationId, turnCount }, null, 2));
  } catch {
    // best effort
  }
}

/** The real filesystem store (config dir, `M365_NO_SESSION_PERSIST` opt-out). */
export function defaultSessionStore() {
  return {
    enabled: process.env.M365_NO_SESSION_PERSIST !== "1",
    path: SESSION_FILE,
    readFile: (p) => readFileSync(p, "utf8"),
    writeFile: (p, data) => {
      mkdirSync(dirname(p), { recursive: true });
      writeFileSync(p, data);
    },
  };
}