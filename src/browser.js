// Playwright loader with a deployment-grade error message.
//
// Playwright is an optional peer dependency: only the interactive sign-in
// needs it, and even then the Chromium browser itself is a separate
// `npx playwright install chromium` download. So the package installs lean,
// and this module turns the raw ERR_MODULE_NOT_FOUND into instructions.

export async function loadChromium(loader = (m) => import(/* webpackIgnore: true */ m)) {
  try {
    const { chromium } = await loader("playwright");
    return chromium;
  } catch (err) {
    if (err?.code === "ERR_MODULE_NOT_FOUND") {
      throw new Error(
        "Playwright is required for browser sign-in but is not installed.\n" +
          "Run: npm i playwright && npx playwright install chromium\n" +
          "(Or avoid the browser entirely: set M365_ACCESS_TOKEN or M365_REFRESH_TOKEN — see README.)",
      );
    }
    throw err;
  }
}
