import { spawnSync } from "node:child_process";
import { closeSync, mkdirSync, openSync, rmSync } from "node:fs";
import { resolve } from "node:path";

const root = resolve(import.meta.dirname, "..");
const lockPath = resolve(root, ".cache/build.lock");
mkdirSync(resolve(root, ".cache"), { recursive: true });
let lock;
try {
  // Lock before cleaning dist so another build cannot remove our Wasm assets.
  lock = openSync(lockPath, "wx");
} catch (cause) {
  if (cause.code !== "EEXIST") throw cause;
  console.error("Another build is already running in this checkout. If all builds have stopped, remove .cache/build.lock and retry.");
  process.exit(1);
}
// shortcut: SIGKILL leaves this lock; remove it after stopping all builds.
process.once("exit", () => {
  closeSync(lock);
  rmSync(lockPath, { force: true });
});
process.once("SIGINT", () => process.exit(130));
process.once("SIGTERM", () => process.exit(143));

for (const step of ["generate:symbol-font-mappings", "clean", "build:core", "build:js"]) {
  const result = spawnSync(process.platform === "win32" ? "npm.cmd" : "npm", ["run", step], {
    cwd: root,
    stdio: "inherit",
    shell: process.platform === "win32",
  });
  if (result.error) throw result.error;
  if (result.status !== 0) {
    process.exitCode = result.status ?? 1;
    break;
  }
}
