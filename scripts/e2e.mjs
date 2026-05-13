#!/usr/bin/env node
/**
 * E2E test: verifies that the Claude CLI can communicate with the gdrive-sa MCP server
 * and successfully call its tools against real Google Drive.
 *
 * Prerequisites:
 *   - `npm run build` has been run  (or run `npm run test:e2e` which builds first)
 *   - Claude CLI is installed and authenticated (`claude --version` works)
 *   - The service account has access to the configured root folder
 */

import { execSync, spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const SCRIPTS = dirname(fileURLToPath(import.meta.url));

let failed = false;

function check(label, fn) {
  process.stdout.write(`  ${label} ... `);
  try {
    const result = fn();
    console.log("✓");
    return result;
  } catch (err) {
    console.log("✗");
    console.error(`    ${err.message}\n`);
    failed = true;
    return null;
  }
}

function claude(prompt, allowedTools) {
  const args = ["-p", prompt, "--allowedTools", allowedTools];
  const result = spawnSync("claude", args, {
    cwd: SCRIPTS,
    encoding: "utf8",
    timeout: 90_000,
  });
  if (result.status !== 0) {
    throw new Error(result.stderr || `claude exited with code ${result.status}`);
  }
  if (!result.stdout.trim()) {
    throw new Error("Claude returned an empty response");
  }
  return result.stdout.trim();
}

console.log("\ngdrive-sa MCP — Claude CLI e2e\n");

// 1. Claude CLI present
check("claude CLI is installed", () => {
  const r = spawnSync("claude", ["--version"], { encoding: "utf8" });
  if (r.status !== 0) throw new Error("claude not found — install Claude CLI first");
});

// 2. Server is built
check("dist/index.js exists", () => {
  if (!existsSync(resolve(ROOT, "dist/index.js"))) {
    throw new Error("Run `npm run build` first");
  }
});

if (failed) process.exit(1);

// 3. list_folder
console.log("\n  [1/2] list_folder — listing root Drive folder (may take ~15s)...\n");
const listOutput = check(
  "Claude calls list_folder and returns results",
  () => claude(
    "Use the list_folder tool from the gdrive-sa MCP server to list the files in the root Google Drive folder. Reply with only the file names, one per line.",
    "mcp__gdrive-sa__list_folder"
  )
);

if (listOutput) {
  console.log("\n  Files found:\n");
  listOutput.split("\n").forEach(l => console.log("    " + l));
}

// 4. search
console.log("\n  [2/2] search — searching for all non-trashed files...\n");
const searchOutput = check(
  "Claude calls search and returns results",
  () => claude(
    "Use the search tool from the gdrive-sa MCP server to search for files with query 'trashed = false'. Reply with how many results were found.",
    "mcp__gdrive-sa__search"
  )
);

if (searchOutput) {
  console.log("\n  Response:", searchOutput);
}

console.log(failed ? "\n✗ Some e2e tests failed\n" : "\n✓ All e2e tests passed\n");
if (failed) process.exit(1);
