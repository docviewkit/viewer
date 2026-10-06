import { createHash } from "node:crypto";
import { readdir, readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(fileURLToPath(new URL("..", import.meta.url)));
const files = (await readdir(resolve(root, "dist")))
  .filter((name) => name.endsWith(".d.ts"))
  .sort();
const hash = createHash("sha256");
for (const name of files) {
  hash.update(`${name}\0`);
  hash.update((await readFile(resolve(root, "dist", name), "utf8")).replaceAll("\r\n", "\n"));
  hash.update("\0");
}
const actual = { algorithm: "sha256", files, digest: hash.digest("hex") };
if (process.argv.includes("--print")) {
  console.log(JSON.stringify(actual, null, 2));
} else {
  const expected = JSON.parse(await readFile(resolve(root, "api-snapshot.json"), "utf8"));
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(`Public API snapshot changed. Review the declarations, then update api-snapshot.json deliberately.\nActual:\n${JSON.stringify(actual, null, 2)}`);
  }
  console.log(`Public API snapshot passed: ${files.length} declaration files, ${actual.digest}`);
}
