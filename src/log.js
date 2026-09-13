import { appendFileSync, mkdirSync } from "node:fs";
import { join } from "node:path";
import { homedir } from "node:os";

const LOG_DIR = process.env.M365_CONFIG_DIR || join(homedir(), ".config", "m365-ask");
const LOG_FILE = join(LOG_DIR, "debug.log");

const enabled = !!process.env.M365_DEBUG;

function write(level, component, ...args) {
  if (!enabled) return;
  const msg = args
    .map((a) => (typeof a === "string" ? a : JSON.stringify(a)))
    .join(" ");
  const line = `[${new Date().toISOString()}] [${level}] [${component}] ${msg}\n`;
  try {
    mkdirSync(LOG_DIR, { recursive: true });
    appendFileSync(LOG_FILE, line);
  } catch {
    // best effort
  }
}

export function createLogger(component) {
  return {
    info: (...args) => write("INFO", component, ...args),
    error: (...args) => write("ERROR", component, ...args),
    debug: (...args) => write("DEBUG", component, ...args),
  };
}
