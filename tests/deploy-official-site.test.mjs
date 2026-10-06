import assert from "node:assert/strict";
import { mkdtemp, mkdir, readFile, readlink, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { resolve } from "node:path";
import { spawnSync } from "node:child_process";
import test from "node:test";

test("official-site deploy switches an exact release and preserves data", async (t) => {
  const root = await mkdtemp(resolve(tmpdir(), "docviewkit-deploy-"));
  const source = resolve(root, "source");
  const deploy = resolve(root, "home/apps/docviewkit");
  t.after(() => rm(root, { recursive: true, force: true }));

  await mkdir(resolve(source, "commercial/src"), { recursive: true });
  await mkdir(resolve(source, "dist"), { recursive: true });
  await mkdir(resolve(deploy, "uploads"), { recursive: true });
  await mkdir(resolve(deploy, "data"), { recursive: true });
  await writeFile(resolve(deploy, "data/persistent"), "keep");
  await writeFile(resolve(source, "commercial/src/server.mjs"), "");
  await writeFile(resolve(source, "commercial/package.json"), '{"type":"module"}\n');
  const run = (version) => spawnSync("bash", [
    "scripts/deploy-official-site.sh",
    `v${version}`,
    version,
    `docviewkit-official-site-${version}.tar.gz`,
  ], {
    cwd: resolve(import.meta.dirname, ".."),
    env: { ...process.env, HOME: resolve(root, "home"), DOCVIEWKIT_DEPLOY_ROOT: deploy },
    encoding: "utf8",
  });
  const archive = async (version) => {
    await writeFile(resolve(source, "dist/version.json"), `{"version":"${version}"}\n`);
    const path = resolve(deploy, "uploads", `docviewkit-official-site-${version}.tar.gz`);
    assert.equal(spawnSync("tar", ["-czf", path, "-C", source, "."]).status, 0);
  };

  await archive("0.2.41");
  let result = run("0.2.41");
  assert.equal(result.status, 0, result.stderr);
  assert.equal(await readlink(resolve(deploy, "current")), "releases/v0.2.41");
  assert.equal(await readFile(resolve(deploy, "server.js"), "utf8"), 'import("./current/commercial/src/server.mjs");\n');

  assert.equal(await readFile(resolve(deploy, "data/persistent"), "utf8"), "keep");
  await mkdir(resolve(deploy, "uploads"), { recursive: true });
  await archive("0.2.42");
  result = run("0.2.42");
  assert.equal(result.status, 0, result.stderr);
  assert.equal(await readlink(resolve(deploy, "current")), "releases/v0.2.42");
  assert.equal(await readFile(resolve(deploy, "data/persistent"), "utf8"), "keep");
});
