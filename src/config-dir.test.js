import test from "node:test";
import assert from "node:assert/strict";

import { defaultSessionStore } from "./session-store.js";

test("the default session store lives under ~/.config/m365-copilot", () => {
  const { path } = defaultSessionStore();

  assert.match(path, /[\\/]\.config[\\/]m365-copilot[\\/]session\.json$/);
});
