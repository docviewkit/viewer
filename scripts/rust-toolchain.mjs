import { spawnSync } from "node:child_process";
import { delimiter, dirname } from "node:path";

// Use the same Cargo and sibling rustc for compilation and license metadata.
export function rustToolchain(root, environment = process.env) {
  let cargo = environment.CARGO;
  if (cargo === undefined) {
    const selected = spawnSync("rustup", ["which", "cargo"], {
      cwd: root, env: environment, encoding: "utf8",
    });
    cargo = selected.status === 0 && selected.stdout.trim() !== ""
      ? selected.stdout.trim()
      : process.platform === "win32" ? "cargo.exe" : "cargo";
  }
  const env = { ...environment };
  if (cargo.includes("/") || cargo.includes("\\")) {
    env.PATH = `${dirname(cargo)}${delimiter}${environment.PATH ?? ""}`;
  }
  return { cargo, env };
}
