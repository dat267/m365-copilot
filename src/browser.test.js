import test from "node:test";
import assert from "node:assert/strict";

import { loadChromium } from "./browser.js";

test("loadChromium returns the playwright chromium object", async () => {
  const chromium = { launchPersistentContext() {} };
  assert.equal(await loadChromium(() => Promise.resolve({ chromium })), chromium);
});

test("loadChromium explains how to install playwright when it is missing", async () => {
  const missing = Object.assign(new Error("Cannot find package 'playwright'"), {
    code: "ERR_MODULE_NOT_FOUND",
  });

  await assert.rejects(
    loadChromium(() => Promise.reject(missing)),
    (err) => /playwright/.test(err.message) && /install/.test(err.message),
  );
});

test("loadChromium rethrows errors that are not a missing module", async () => {
  await assert.rejects(loadChromium(() => Promise.reject(new Error("boom"))), /boom/);
});
