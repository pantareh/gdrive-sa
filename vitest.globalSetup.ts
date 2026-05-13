import { execSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { dirname } from "node:path";

const ROOT = dirname(fileURLToPath(import.meta.url));

export async function setup() {
  execSync("npm run build", { cwd: ROOT, stdio: "inherit" });
}
