import { describe, it, expect } from "vitest";
import { spawnSync, spawn, type ChildProcessWithoutNullStreams } from "node:child_process";
import { fileURLToPath } from "node:url";
import { resolve, dirname } from "node:path";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "../..");
const DIST_INDEX = resolve(ROOT, "dist/index.js");

// ---------------------------------------------------------------------------
// Helpers for MCP stdio wire protocol (newline-delimited JSON)
// ---------------------------------------------------------------------------

function encodeMcpMessage(msg: object): Buffer {
  return Buffer.from(JSON.stringify(msg) + "\n", "utf8");
}

function readMcpResponse(child: ChildProcessWithoutNullStreams, timeoutMs = 4000): Promise<unknown> {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error("Timeout waiting for MCP response")), timeoutMs);
    let buf = "";

    function onData(chunk: Buffer) {
      buf += chunk.toString("utf8");
      const newline = buf.indexOf("\n");
      if (newline === -1) return;
      clearTimeout(timer);
      child.stdout.off("data", onData);
      try {
        resolve(JSON.parse(buf.slice(0, newline)));
      } catch (e) {
        reject(e);
      }
    }

    child.stdout.on("data", onData);
  });
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("CLI startup", () => {
  it("exits with code 1 when GOOGLE_APPLICATION_CREDENTIALS is not set", () => {
    const env = { ...process.env };
    delete env.GOOGLE_APPLICATION_CREDENTIALS;

    const result = spawnSync("node", [DIST_INDEX], { env, timeout: 5000 });

    expect(result.status).toBe(1);
    expect(result.stderr.toString()).toContain("GOOGLE_APPLICATION_CREDENTIALS");
  });

  it("starts and stays running when GOOGLE_APPLICATION_CREDENTIALS is set", async () => {
    // GoogleAuth is lazy — it only reads the key file on the first API call,
    // so a non-existent path is fine for startup testing.
    const child = spawn("node", [DIST_INDEX], {
      env: { ...process.env, GOOGLE_APPLICATION_CREDENTIALS: "/fake/sa.json" },
      stdio: ["pipe", "pipe", "pipe"],
    }) as ChildProcessWithoutNullStreams;

    let exited = false;
    child.on("close", () => { exited = true; });

    // Give enough time for an immediate crash to surface
    await new Promise((r) => setTimeout(r, 300));

    expect(exited).toBe(false);

    child.kill();
    await new Promise((r) => child.on("close", r));
  });

  it("completes MCP initialize handshake and reports server name", async () => {
    const child = spawn("node", [DIST_INDEX], {
      env: { ...process.env, GOOGLE_APPLICATION_CREDENTIALS: "/fake/sa.json" },
      stdio: ["pipe", "pipe", "pipe"],
    }) as ChildProcessWithoutNullStreams;

    const responsePromise = readMcpResponse(child);

    child.stdin.write(
      encodeMcpMessage({
        jsonrpc: "2.0",
        id: 1,
        method: "initialize",
        params: {
          protocolVersion: "2025-11-25",
          capabilities: {},
          clientInfo: { name: "test-client", version: "1.0.0" },
        },
      })
    );

    const response = await responsePromise;

    expect(response).toMatchObject({
      jsonrpc: "2.0",
      id: 1,
      result: {
        serverInfo: { name: "gdrive-sa", version: "1.0.0" },
        capabilities: { resources: {}, tools: {} },
      },
    });

    child.kill();
    await new Promise((r) => child.on("close", r));
  });
});
