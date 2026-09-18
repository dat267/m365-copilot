// Commander-based CLI wiring: `auth` and `ask`. No TUI, no REPL.
//
// buildProgram() takes its collaborators as injectable deps so tests can drive
// the real command parsing with fakes and no network. The bin (cli.js at the
// repo root) calls buildProgram() with the real ones.

import { Command, Option } from "commander";
import { readFile } from "node:fs/promises";
import { ask as realAsk, getAvailableModels as realModels, loginInteractive as realLogin } from "./index.js";
import { loadContext, buildPrompt } from "./prompt.js";

export function buildProgram({
  ask = realAsk,
  getAvailableModels = realModels,
  loginInteractive = realLogin,
  io = { readFile: (path) => readFile(path, "utf8"), readStdin },
  out = (s) => process.stdout.write(s),
  err = (s) => process.stderr.write(s),
} = {}) {
  const program = new Command();
  program
    .name("m365-copilot")
    .summary("Microsoft 365 Copilot from the command line")
    .exitOverride();

  program
    .command("auth")
    .description("sign in via a Playwright-driven browser and cache the tokens")
    .action(async () => {
      await loginInteractive();
      err("[m365-copilot] signed in; tokens cached.\n");
    });

  program
    .command("ask")
    .argument("[prompt...]", "the prompt (multiple words are joined)")
    .addOption(
      new Option("-m, --model <id>", "model to use")
        .choices(getAvailableModels())
        .default("m365-copilot"),
    )
    .option("-c, --context <file|->", "prepend a file (or stdin with -) to the prompt as data")
    .action(async (parts, options) => {
      const context = await loadContext(options.context, io);
      const answer = await ask(buildPrompt(context, parts.join(" ")), {
        model: options.model,
        temporary: true,
      });
      out(answer + "\n");
    });

  return program;
}

async function readStdin() {
  let data = "";
  process.stdin.setEncoding("utf8");
  for await (const chunk of process.stdin) data += chunk;
  return data;
}
