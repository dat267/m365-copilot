// Auth for M365 Copilot (Sydney) — MSAL PKCE with the `nativeclient` redirect.
//
// This is the only genuinely hard part of talking to M365 Copilot: there is no
// API key, so we impersonate Microsoft's own Office-web client.
//
// Flow:
//   1. getToken() tries a SILENT refresh from a cached MSAL token.
//   2. On failure it opens a VISIBLE browser; you sign in once (SSO/MFA).
//   3. The `?code=` is scraped from the navigation REQUEST to the nativeclient
//      redirect — the page itself bounces to /common/wrongplace, so waiting for
//      a settled URL misses it.
//   4. The code is exchanged (PKCE) and the token cache is persisted, so every
//      later run is silent.
//
// The persistent browser profile keeps the AAD session cookies, so subsequent
// sign-ins (after cache expiry) are usually SSO-silent too.

import * as msal from "@azure/msal-node";
import { readFileSync, writeFileSync, existsSync, mkdirSync } from "node:fs";
import { execSync } from "node:child_process";
import { join } from "node:path";
import { homedir } from "node:os";
import { createLogger } from "./log.js";

const log = createLogger("auth");

// Microsoft's first-party Office-web Copilot client. We did not register this.
// Overridable: if you copy a refresh token out of a browser that used a sibling
// first-party client, set M365_CLIENT_ID to that client's id.
const CLIENT_ID = process.env.M365_CLIENT_ID || "c0ab8ce9-e9a0-42e7-b064-33d422df41f1";
const AUTHORITY = "https://login.microsoftonline.com/common";
const REDIRECT_URI = "https://login.microsoftonline.com/common/oauth2/nativeclient";
const SCOPES = [
  "https://substrate.office.com/sydney/M365Chat.Read",
  "https://substrate.office.com/sydney/sydney.readwrite",
];
// Graph needs its own token. The first-party client is pre-consented for
// `/.default` (a per-scope request like Files.ReadWrite is rejected with
// AADSTS65002); `.default` returns Files.ReadWrite.All among others.
const GRAPH_SCOPES = ["https://graph.microsoft.com/.default"];

const CONFIG_DIR = process.env.M365_CONFIG_DIR || join(homedir(), ".config", "m365-ask");
const CACHE_FILE = process.env.M365_CACHE_FILE || join(CONFIG_DIR, "msal-cache.json");
const BROWSER_PROFILE_DIR = process.env.M365_BROWSER_PROFILE || join(CONFIG_DIR, "browser-profile");
// Raw refresh-token store (access token + rotated refresh token). See the
// "browser-copied token" path in getToken() — lets you run with no browser at all.
const TOKEN_FILE = process.env.M365_TOKEN_FILE || join(CONFIG_DIR, "token.json");

// A coherent, non-headless-looking UA. The default headless Chromium advertises
// "HeadlessChrome", a loud bot signal to AAD's risk engine.
const LOGIN_USER_AGENT =
  process.env.M365_LOGIN_UA ??
  "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36";

function resolveChromiumPath() {
  if (process.env.CHROMIUM_PATH) return process.env.CHROMIUM_PATH;
  if (process.platform === "win32") return undefined; // use Playwright's bundled browser
  for (const bin of ["chromium", "chromium-browser", "google-chrome", "chrome"]) {
    try {
      const found = execSync(`command -v ${bin}`, { stdio: ["ignore", "pipe", "ignore"] }).toString().trim();
      if (found) return found;
    } catch {
      // try next
    }
  }
  return undefined;
}

// --- MSAL cache persistence ---

let _app = null;
function getApp() {
  if (!_app) {
    _app = new msal.PublicClientApplication({ auth: { clientId: CLIENT_ID, authority: AUTHORITY } });
    if (existsSync(CACHE_FILE)) {
      try {
        _app.getTokenCache().deserialize(readFileSync(CACHE_FILE, "utf-8"));
      } catch {}
    }
  }
  return _app;
}

function saveCache(app) {
  try {
    mkdirSync(CONFIG_DIR, { recursive: true });
    writeFileSync(CACHE_FILE, app.getTokenCache().serialize());
  } catch {}
}

async function buildAuthUrlForScopes(app) {
  const cryptoProvider = new msal.CryptoProvider();
  const { verifier, challenge } = await cryptoProvider.generatePkceCodes();
  const authUrl = await app.getAuthCodeUrl({
    scopes: SCOPES,
    redirectUri: REDIRECT_URI,
    codeChallenge: challenge,
    codeChallengeMethod: "S256",
  });
  return { authUrl, verifier };
}

// --- Silent refresh ---

export async function getTokenSilent() {
  const app = getApp();
  const accounts = await app.getTokenCache().getAllAccounts();
  if (accounts.length === 0) return null;
  try {
    const result = await app.acquireTokenSilent({ scopes: SCOPES, account: accounts[0] });
    saveCache(app);
    return result.accessToken;
  } catch {
    return null;
  }
}

// --- Interactive (visible browser) login ---

export async function loginInteractive() {
  const { chromium } = await import("playwright");
  const app = getApp();
  const { authUrl, verifier } = await buildAuthUrlForScopes(app);

  const context = await chromium.launchPersistentContext(BROWSER_PROFILE_DIR, {
    headless: false, // a human has to see and drive this
    executablePath: resolveChromiumPath(),
    args: ["--no-sandbox", "--disable-dev-shm-usage", "--disable-blink-features=AutomationControlled"],
    userAgent: LOGIN_USER_AGENT,
    locale: "en-GB",
    timezoneId: "Europe/Copenhagen",
    viewport: { width: 1280, height: 800 },
  });
  await context.addInitScript(() => {
    Object.defineProperty(navigator, "webdriver", { get: () => undefined });
  });
  const page = context.pages()[0] ?? (await context.newPage());

  // The nativeclient redirect is meant for embedded native hosts to intercept;
  // a real browser follows it one hop further to /common/wrongplace, so the
  // ?code= only exists transiently. Capture it from the navigation request.
  const codePromise = new Promise((resolve) => {
    page.on("request", (req) => {
      const u = req.url();
      if (u.includes("/oauth2/nativeclient") && u.includes("code=")) {
        const c = new URL(u).searchParams.get("code");
        if (c) {
          log.info("Captured auth code from nativeclient redirect");
          resolve(c);
        }
      }
    });
  });

  try {
    console.error("[m365-ask] A browser window has opened — complete Microsoft sign-in there.");
    console.error("[m365-ask] Waiting up to 10 minutes...");
    await page.goto(authUrl, { waitUntil: "domcontentloaded" });

    let timer;
    const code = await Promise.race([
      codePromise,
      new Promise((_, rej) => {
        timer = setTimeout(() => rej(new Error("Timed out waiting for interactive sign-in")), 600_000);
      }),
    ]).finally(() => clearTimeout(timer));

    const result = await app.acquireTokenByCode({
      code,
      scopes: SCOPES,
      redirectUri: REDIRECT_URI,
      codeVerifier: verifier,
    });
    saveCache(app);
    log.info(`Interactive login succeeded as ${result.account?.username}`);
    return result.accessToken;
  } finally {
    await context.close();
  }
}

// --- Raw refresh-token grant (no MSAL, no browser) ---
//
// An access token alone expires in ~1h; a REFRESH token is the durable
// credential. If you copy one out of your browser (see README), this exchanges
// it for a fresh access token with a plain HTTPS POST. AAD ROTATES the refresh
// token on every call, so the newest one is persisted back to TOKEN_FILE —
// always use a single store, or you'll hold a stale token.

function loadTokenFile() {
  try {
    return JSON.parse(readFileSync(TOKEN_FILE, "utf-8"));
  } catch {
    return null;
  }
}

function saveTokenFile(obj) {
  try {
    mkdirSync(CONFIG_DIR, { recursive: true });
    writeFileSync(TOKEN_FILE, JSON.stringify(obj, null, 2));
  } catch {}
}

/** Exchange a refresh token for a fresh access token. Returns the rotated
 *  refresh token too — persist it. */
export async function refreshTokenGrant(refreshToken, scopes = SCOPES) {
  const body = new URLSearchParams({
    grant_type: "refresh_token",
    client_id: CLIENT_ID,
    refresh_token: refreshToken,
    scope: scopes.join(" "),
  });
  const res = await fetch("https://login.microsoftonline.com/common/oauth2/v2.0/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });
  const json = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(
      `refresh_token grant failed (HTTP ${res.status}): ${json.error_description || json.error || "unknown error"}`,
    );
  }
  return {
    accessToken: json.access_token,
    refreshToken: json.refresh_token ?? refreshToken, // AAD rotates these
    expiresAt: Date.now() + (json.expires_in ?? 3600) * 1000,
  };
}

// --- Public: get a usable token (silent, else interactive) ---

let inflight = null;
export function getToken() {
  return (inflight ??= doGetToken().finally(() => {
    inflight = null;
  }));
}

/**
 * A Graph access token, for the OneDrive/file-attachment path.
 * Rotates ONLY the refresh token in the shared token file, so the cached Sydney
 * access token is left intact.
 */
export async function getGraphToken() {
  if (process.env.M365_GRAPH_TOKEN) return process.env.M365_GRAPH_TOKEN;
  const saved = loadTokenFile();
  const refreshToken = process.env.M365_REFRESH_TOKEN || saved?.refreshToken;
  if (!refreshToken) {
    throw new Error("no refresh token available to mint a Graph token (set M365_REFRESH_TOKEN or M365_GRAPH_TOKEN)");
  }
  const t = await refreshTokenGrant(refreshToken, GRAPH_SCOPES);
  saveTokenFile({ accessToken: saved?.accessToken, expiresAt: saved?.expiresAt, refreshToken: t.refreshToken });
  return t.accessToken;
}

async function doGetToken() {
  // 1. An explicitly supplied access token (e.g. copied from the browser's
  //    WebSocket URL). Short-lived (~1h); no refresh possible from it.
  if (process.env.M365_ACCESS_TOKEN) {
    log.info("Using M365_ACCESS_TOKEN from env");
    return process.env.M365_ACCESS_TOKEN;
  }

  // 2. A previously saved access token that is still valid.
  const saved = loadTokenFile();
  if (saved?.accessToken && (saved.expiresAt ?? 0) - 60_000 > Date.now()) {
    log.info("Using cached access token");
    return saved.accessToken;
  }

  // 3. Refresh-token grant — browser-copied (env) or previously saved. This is
  //    the path that needs NO browser after the first token is supplied.
  const refreshToken = process.env.M365_REFRESH_TOKEN || saved?.refreshToken;
  if (refreshToken) {
    log.info("Refreshing access token via refresh_token grant");
    const t = await refreshTokenGrant(refreshToken);
    saveTokenFile(t);
    return t.accessToken;
  }

  // 4. Silent MSAL refresh from the interactive sign-in cache.
  const silent = await getTokenSilent();
  if (silent) {
    log.info("Token refreshed silently");
    return silent;
  }

  // 5. Visible browser sign-in (self-healing fallback).
  return loginInteractive();
}

/** Decode the JWT payload (oid/tid/exp). No signature verification — this is our
 *  own token, already validated by M365. */
export function decodeJwt(token) {
  const payload = token.split(".")[1];
  const padded = payload + "=".repeat((4 - (payload.length % 4)) % 4);
  return JSON.parse(Buffer.from(padded, "base64").toString());
}
