import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import {
  cp,
  mkdir,
  mkdtemp,
  readFile,
  rm,
  stat,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import { resolve } from "node:path";
import { pathToFileURL } from "node:url";
import test from "node:test";

import {
  prepareRelease,
  validateBundledFonts,
  validateLocalDependencies,
} from "../scripts/prepare-release.mjs";

const manifest = JSON.parse(await readFile(new URL("../package.json", import.meta.url), "utf8"));

test("bundled font gate rejects unreviewed, oversized, altered and unlicensed release assets", async (t) => {
  const stage = await mkdtemp(resolve(tmpdir(), "docviewkit-font-gate-"));
  t.after(() => rm(stage, { recursive: true, force: true }));
  await cp(new URL("../dist/", import.meta.url), stage, { recursive: true });
  await validateBundledFonts(stage);
  const extra = resolve(stage, "Unreviewed.woff2");
  await writeFile(extra, "unreviewed");
  await assert.rejects(validateBundledFonts(stage), /approved 2 families \/ 8 faces/);
  await rm(extra);
  const name = resolve(stage, "Carlito-Regular.ttf");
  const original = await readFile(name);
  await rm(name);
  await assert.rejects(validateBundledFonts(stage), /approved 2 families \/ 8 faces/);
  await writeFile(name, Buffer.alloc(3 * 1024 * 1024));
  await assert.rejects(validateBundledFonts(stage), /3 MiB total budget/);
  await writeFile(name, original);
  const caladeaPath = resolve(stage, "Caladea-Regular.ttf");
  const caladea = await readFile(caladeaPath);
  await writeFile(caladeaPath, Buffer.alloc(64 * 1024 + 1));
  await assert.rejects(validateBundledFonts(stage), /size budget: Caladea-Regular/);
  await writeFile(caladeaPath, caladea);
  const altered = Buffer.from(original);
  altered[altered.length - 1] ^= 1;
  await writeFile(name, altered);
  await assert.rejects(validateBundledFonts(stage), /reviewed upstream file: Carlito-Regular/);
  await writeFile(name, original);
  const licensePath = resolve(stage, "third-party-licenses/Carlito-1.103-LICENSE.txt");
  const license = await readFile(licensePath);
  await rm(licensePath);
  await assert.rejects(validateBundledFonts(stage), /license is missing/);
  await writeFile(licensePath, "OFL-1.1");
  await assert.rejects(validateBundledFonts(stage), /license differs/);
  await writeFile(licensePath, license);
  const noticesPath = resolve(stage, "THIRD_PARTY_NOTICES.md");
  const notices = await readFile(noticesPath, "utf8");
  await writeFile(noticesPath, notices.replace(/^.*\| Carlito \|.*\n/mu, ""));
  await assert.rejects(validateBundledFonts(stage), /attribution is missing: Carlito/);
  await writeFile(noticesPath, notices);
  await validateBundledFonts(stage);
});

test("release preparation projects one build into public Viewer, SDK and Pages without website source", async (t) => {
  const output = await mkdtemp(resolve(tmpdir(), "docviewkit-release-"));
  t.after(() => rm(output, { recursive: true, force: true }));

  const result = await prepareRelease({
    tag: `v${manifest.version}`,
    output,
    nativeVisualStatus: "not-certified",
  });
  assert.equal(result.version, manifest.version);
  assert.equal(result.npmTag, "latest");
  assert.equal(result.freeArchive, `docviewkit-viewer-${manifest.version}.tgz`);
  assert.equal(result.sdkArchive, `docviewkit-sdk-${manifest.version}.tgz`);
  assert.deepEqual(result.nativeVisualAcceptance, {
    policy: "native-visual-v1",
    status: "not-certified",
  });
  const releaseMetadata = JSON.parse(
    await readFile(resolve(output, "release.json"), "utf8"),
  );
  assert.deepEqual(releaseMetadata.nativeVisualAcceptance, result.nativeVisualAcceptance);

  for (const [family, version, licenseHash] of [
    ["Carlito", "1.103", "bbcf8ce9c8acc91355ab5d263ec9e37d498de2222336e7d42e589f791afe56c4"],
    ["Caladea", "1.002", "6f1041c12f758ed86d804acbcb54ad822d053fa15520184c28c3b8eabb8f66f6"],
  ]) {
    const sourceRoot = new URL(`../third_party/${family.toLowerCase()}/`, import.meta.url);
    const license = await readFile(new URL("LICENSE.txt", sourceRoot));
    assert.equal(createHash("sha256").update(license).digest("hex"), licenseHash);
    const assets = new Map([[`third-party-licenses/${family}-${version}-LICENSE.txt`, license]]);
    for (const variant of ["Regular", "Bold", "Italic", "BoldItalic"]) {
      const name = `${family}-${variant}.ttf`;
      assets.set(name, await readFile(new URL(name, sourceRoot)));
    }
    for (const [name, bytes] of assets) {
      for (const stage of ["free-package", `pages/sdk/v${manifest.version}`]) {
        assert.deepEqual(await readFile(resolve(output, stage, name)), bytes, `${stage}/${name}`);
      }
      for (const [archive, prefix] of [[result.freeArchive, "package/"], [result.sdkArchive, "package/dist/"]]) {
        assert.deepEqual(execFileSync("tar", ["-xOf", resolve(output, archive), `${prefix}${name}`]), bytes, `${archive}/${name}`);
      }
    }
  }

  const dingbatsLicense = await readFile(new URL("../third_party/pdf-standard-fonts/LICENSE_FOXIT", import.meta.url));
  const dingbatsLicensePath = "third-party-licenses/PDFium-FoxitDingbats-LICENSE.txt";
  for (const stage of ["free-package", `pages/sdk/v${manifest.version}`]) {
    assert.deepEqual(await readFile(resolve(output, stage, dingbatsLicensePath)), dingbatsLicense);
  }
  for (const [archive, prefix] of [[result.freeArchive, "package/"], [result.sdkArchive, "package/dist/"]]) {
    assert.deepEqual(execFileSync("tar", ["-xOf", resolve(output, archive), `${prefix}${dingbatsLicensePath}`]), dingbatsLicense);
  }

  const free = JSON.parse(await readFile(resolve(output, "free-package/package.json"), "utf8"));
  assert.equal(free.name, "@docviewkit/viewer");
  assert.equal(free.version, manifest.version);
  assert.equal(free.repository.url, "git+https://github.com/docviewkit/viewer.git");
  for (const keyword of [
    "document-viewer", "document-preview", "office-viewer", "pdf-viewer", "ofd-viewer",
    "word-viewer", "excel-viewer", "powerpoint-viewer", "frontend", "client-side",
    "web-component", "wasm", "webview", "electron", "tauri", "react", "vue", "angular",
    "doc", "docx", "xls", "xlsx", "ppt", "pptx", "pdf", "ofd", "odt", "ods", "odp",
    "pages", "numbers", "keynote", "wps", "et", "dps", "xps", "oxps", "rtf", "csv",
    "文档预览", "办公文档预览", "ofd预览",
  ]) {
    assert.ok(free.keywords.includes(keyword), `missing npm keyword: ${keyword}`);
  }
  assert.equal(new Set(free.keywords).size, free.keywords.length);
  assert.equal(free.exports["."].import, "./viewer.js");
  assert.equal(free.exports["./package.json"], "./package.json");
  assert.equal(free.license, "Apache-2.0");
  assert.equal(free.exports["./engine"].import, "./index.js");
  const engineApi = await import(pathToFileURL(resolve(output, "free-package/index.js")));
  const wasm = await readFile(resolve(output, "free-package/office-viewer-core.wasm"));
  const engine = await engineApi.createOfficeEngine({ wasm, execution: "inline" });
  try {
    const document = await engine.open(await readFile(new URL("./fixtures/visual-baseline.pptx", import.meta.url)));
    assert.equal(document.info.format, "pptx");
    assert.ok((await document.listObjects({ unitIndex: 0 })).length > 0);
    document.close();
  } finally {
    engine.close();
  }
  for (const entry of ["LICENSE", "NOTICE"]) {
    const bytes = await readFile(new URL(`../${entry}`, import.meta.url));
    for (const stage of ["free-package", "pages", `pages/sdk/v${manifest.version}`]) {
      assert.deepEqual(await readFile(resolve(output, stage, entry)), bytes, `${stage}/${entry}`);
    }
    for (const archive of [result.freeArchive, result.sdkArchive]) {
      assert.deepEqual(execFileSync("tar", ["-xOf", resolve(output, archive), `package/${entry}`]), bytes);
    }
  }
  const freeReadme = await readFile(resolve(output, "free-package/README.md"), "utf8");
  assert.equal(execFileSync("tar", ["-xOf", resolve(output, result.freeArchive), "package/README.md"], { encoding: "utf8" }), freeReadme);
  const sourceReadme = await readFile(new URL("../README.md", import.meta.url), "utf8");
  assert.equal(freeReadme, sourceReadme, "npm and GitHub must use the same product README");
  assert.equal(execFileSync("tar", ["-xOf", resolve(output, result.sdkArchive), "package/README.md"], { encoding: "utf8" }), sourceReadme);
  const chineseReadme = await readFile(new URL("../README.zh-CN.md", import.meta.url), "utf8");
  assert.ok(sourceReadme.includes("[简体中文](https://github.com/docviewkit/viewer/blob/main/README.zh-CN.md)"));
  assert.ok(chineseReadme.includes("[English](https://github.com/docviewkit/viewer/blob/main/README.md)"));
  const chineseBenefits = chineseReadme.indexOf("\n## 为什么选择 DocViewKit");
  const chineseScenarios = chineseReadme.indexOf("\n## 适用场景");
  assert.ok(chineseBenefits > 0 && chineseBenefits < chineseReadme.indexOf("\n```"));
  assert.ok(chineseScenarios > chineseBenefits && chineseScenarios < chineseReadme.indexOf("\n```"));
  for (const term of ["Apache-2.0", "OFD", "Electron", "Tauri", "WebView", "支持、定制和企业交付", "兼容性和渲染深度"]) assert.ok(chineseReadme.includes(term), term);
  assert.deepEqual(chineseReadme.match(/```[^\n]*\n[\s\S]*?```/gu), sourceReadme.match(/```[^\n]*\n[\s\S]*?```/gu), "translated quickstart must preserve the runnable integration example");
  for (const [surface, content] of [["GitHub", sourceReadme], ["npm", freeReadme]]) {
    const firstExample = content.indexOf("\n```");
    const benefits = content.indexOf("\n## Why DocViewKit");
    const scenarios = content.indexOf("\n## Where it fits");
    assert.ok(benefits > 0 && benefits < firstExample, `${surface} must explain product advantages before code`);
    assert.ok(scenarios > benefits && scenarios < firstExample, `${surface} must explain use cases before code`);
    assert.ok((content.match(/^```/gmu) ?? []).length <= 6, `${surface} must keep detailed engineering examples in documentation`);
    assert.match(content, /https:\/\/docviewkit\.com\/en\/demo\//u);
    assert.match(content, /https:\/\/docviewkit\.com\/docs\/supported-formats\//u);
    assert.match(content, /Apache-2\.0/u);
    assert.doesNotMatch(content, /\b(?:100% fidelity|pixel-perfect|zero latency)\b/iu);
  }
  assert.match(freeReadme, /OFD/);
  assert.match(freeReadme, /ofd-formats/);
  assert.match(freeReadme, /Electron, Tauri/);
  assert.match(free.description, /OFD/);
  assert.match(free.description, /Frontend-local/);

  for (const entry of ["viewer.js", "worker.js"]) {
    const source = await readFile(resolve(output, "free-package", entry), "utf8");
    const chunks = [...source.matchAll(/["']\.\/(runtime-[A-Z0-9]+\.js)["']/gu)]
      .map((match) => match[1]);
    assert.ok(chunks.length > 0, `${entry} must retain shared runtime imports`);
    for (const chunk of new Set(chunks)) {
      await stat(resolve(output, "free-package", chunk));
      await stat(resolve(output, `pages/sdk/v${manifest.version}`, chunk));
    }
  }

  const page = await readFile(resolve(output, "pages/index.html"), "utf8");
  const demo = await readFile(resolve(output, "pages/demo/demo.js"), "utf8");
  assert.match(page, new RegExp(`v${manifest.version.replaceAll(".", "\\.")}`, "u"));
  assert.match(demo, new RegExp(`const version = "${manifest.version.replaceAll(".", "\\.")}"`, "u"));
  assert.doesNotMatch(page, /__DOCVIEWKIT_/u);
  assert.doesNotMatch(demo, /__DOCVIEWKIT_/u);

  await stat(resolve(output, `pages/sdk/v${manifest.version}/viewer.js`));
  await assert.rejects(stat(resolve(output, "official-site")), { code: "ENOENT" });
  await assert.rejects(stat(new URL("../commercial/package.json", import.meta.url)), { code: "ENOENT" });

});

test("release preparation rejects a tag that differs from package and Cargo versions", async () => {
  await assert.rejects(
    prepareRelease({ tag: "v99.99.99", output: resolve(tmpdir(), "unused-docviewkit-release") }),
    /does not match/u,
  );
});

test("release preparation rejects an unknown native visual status", async () => {
  await assert.rejects(
    prepareRelease({
      tag: `v${manifest.version}`,
      output: resolve(tmpdir(), "unused-docviewkit-release"),
      nativeVisualStatus: "advisory",
    }),
    /Unsupported native visual acceptance status/u,
  );
});

test("release validation rejects untracked entry points from local dependencies", async (t) => {
  const root = await mkdtemp(resolve(tmpdir(), "docviewkit-local-dependency-"));
  t.after(() => rm(root, { recursive: true, force: true }));
  const dependency = resolve(root, "third_party/emf-converter");
  await mkdir(resolve(dependency, "dist"), { recursive: true });
  await writeFile(
    resolve(root, "package.json"),
    `${JSON.stringify({ devDependencies: { "emf-converter": "file:third_party/emf-converter" } })}\n`,
  );
  await writeFile(
    resolve(dependency, "package.json"),
    `${JSON.stringify({
      name: "emf-converter",
      main: "dist/index.js",
      types: "dist/index.d.ts",
    })}\n`,
  );
  await writeFile(resolve(dependency, "dist/index.js"), "export {};\n");
  await writeFile(resolve(dependency, "dist/index.d.ts"), "export {};\n");
  execFileSync("git", ["init", "-q"], { cwd: root });
  execFileSync(
    "git",
    ["add", "package.json", "third_party/emf-converter/package.json"],
    { cwd: root },
  );

  await assert.rejects(
    validateLocalDependencies(root),
    /emf-converter exports an untracked file: third_party\/emf-converter\/dist\/index\.js/u,
  );

  execFileSync(
    "git",
    ["add", "third_party/emf-converter/dist/index.js", "third_party/emf-converter/dist/index.d.ts"],
    { cwd: root },
  );
  await validateLocalDependencies(root);
});

test("Wasm builds use the project-selected Rust toolchain", async () => {
  const source = await readFile(new URL("../scripts/build-core.mjs", import.meta.url), "utf8");
  assert.doesNotMatch(source, /--toolchain[",\s]+stable/u);
  assert.match(source, /\["which", "cargo"\]/u);
  assert.match(source, /process\.env\.CARGO/u);
});

test("local server builds optimized Wasm", async () => {
  const source = await readFile(new URL("../scripts/build-core.mjs", import.meta.url), "utf8");
  const server = await readFile(new URL("../scripts/test-server.sh", import.meta.url), "utf8");
  assert.match(source, /"build",\s*"--release"/u);
  assert.doesNotMatch(source, /--debug|wasm32-unknown-unknown\/debug/u);
  assert.match(server, /npm run build\b/u);
  assert.doesNotMatch(server, /npm run build:debug/u);
  assert.equal(manifest.scripts["build:debug"], undefined);
  assert.equal(manifest.scripts["build:core:debug"], undefined);
});

test("CI installs browsers with the pinned Playwright package", async () => {
  const workflow = await readFile(new URL("../.github/workflows/ci.yml", import.meta.url), "utf8");
  assert.equal(
    manifest.scripts["test:browsers:install"],
    "node node_modules/playwright-core/cli.js install --with-deps chromium firefox webkit",
  );
  assert.match(workflow, /npm run test:browsers:install/u);
  assert.doesNotMatch(workflow, /npx playwright install/u);
});

test("automatic compilation is reserved for release tags", async () => {
  const ci = await readFile(new URL("../.github/workflows/ci.yml", import.meta.url), "utf8");
  const release = await readFile(new URL("../.github/workflows/release.yml", import.meta.url), "utf8");
  assert.match(ci, /^on:\n  workflow_dispatch:\s*$/mu);
  assert.doesNotMatch(ci, /^  (?:pull_request|push):/mu);
  assert.match(release, /^on:\n  push:\n    tags: \["v\*"\]/mu);
});

test("release workflow marks missing native goldens as not certified", async () => {
  const workflow = await readFile(new URL("../.github/workflows/release.yml", import.meta.url), "utf8");
  assert.match(workflow, /--validate-only --allow-missing/u);
  assert.match(workflow, /status=not-certified/u);
  assert.match(workflow, /Native visual acceptance: not certified/u);
  assert.match(workflow, /NATIVE_VISUAL_STATUS: \$\{\{ steps\.native_visual\.outputs\.status \}\}/u);
});

test("release workflow does not require unsupported private-repository attestations", async () => {
  const workflow = await readFile(new URL("../.github/workflows/release.yml", import.meta.url), "utf8");
  assert.match(workflow, /REPOSITORY_PRIVATE: \$\{\{ github\.event\.repository\.private \}\}/u);
  assert.match(workflow, /ENABLE_GITHUB_ATTESTATIONS: \$\{\{ vars\.ENABLE_GITHUB_ATTESTATIONS \}\}/u);
  assert.match(workflow, /status=unavailable/u);
  assert.match(workflow, /actions\/attest-build-provenance@v4/u);
  assert.match(workflow, /if: steps\.provenance\.outputs\.status == 'enabled'/u);
  assert.match(workflow, /SHA256SUMS and CycloneDX SBOM are attached/u);
  assert.match(workflow, /release-output\/docviewkit-viewer-\*\.tgz/u);
  assert.match(workflow, /release-output\/docviewkit-viewer-pages-\*\.tar\.gz/u);
});

test("same-repository release preserves published assets and dispatches complete drafts", async (t) => {
  const workflow = await readFile(new URL("../.github/workflows/release.yml", import.meta.url), "utf8");
  const root = await mkdtemp(resolve(tmpdir(), "docviewkit-release-workflow-"));
  t.after(() => rm(root, { recursive: true, force: true }));
  const calls = resolve(root, "gh-calls");
  const step = (name) => {
    const source = workflow.split(`      - name: ${name}\n`)[1]?.split(/\n      - |\n  [a-z]/u)[0];
    assert.ok(source, `missing workflow step: ${name}`);
    return source.split("        run: |\n")[1]
      .replace(/^          /gmu, "")
      .replace(/\$\{\{ steps\.provenance\.outputs\.status \}\}/gu, "enabled");
  };
  const run = async (name, draft) => {
    await writeFile(calls, "");
    execFileSync("bash", ["-e", "-c", `
      gh() {
        printf '%s\\n' "$*" >> "$GH_CALLS"
        if [ "$1 $2" = "release view" ]; then
          [ "$RELEASE_DRAFT" != missing ] || return 1
          case "$*" in
            *'--json isDraft'*) printf '%s\\n' "$RELEASE_DRAFT" ;;
            *'--json databaseId'*) printf '42\\n' ;;
          esac
        fi
      }
      ${step(name)}
    `], {
      env: { ...process.env, GH_CALLS: calls, GH_TOKEN: "test", GH_REPO: "docviewkit/viewer",
        GITHUB_REPOSITORY: "docviewkit/viewer",
        RELEASE_TAG: "v0.2.74", RELEASE_VERSION: "0.2.74", RELEASE_DRAFT: draft,
        NATIVE_VISUAL_STATUS: "not-certified" },
    });
    return readFile(calls, "utf8");
  };
  const draftStep = "Create or update source-tag draft release";
  const published = await run(draftStep, "false");
  assert.doesNotMatch(published, /release (?:create|edit|upload) /u);
  const created = await run(draftStep, "missing");
  assert.match(created, /release create v0\.2\.74 --draft --verify-tag/u);
  const resumed = await run(draftStep, "true");
  assert.match(resumed, /release edit v0\.2\.74/u);
  for (const name of ["viewer-0.2.74.tgz", "sdk-0.2.74.tgz", "viewer-pages-0.2.74.tar.gz",
    "sdk-0.2.74.cdx.json", "PUBLIC_SHA256SUMS", "SHA256SUMS", "release.json"]) {
    assert.ok(created.includes(name) && resumed.includes(name), `draft is missing ${name}`);
  }
  const dispatched = await run("Dispatch verified publication", "true");
  assert.match(dispatched, /api --method POST repos\/docviewkit\/viewer\/dispatches -f event_type=release-ready -F client_payload\[tag\]=v0\.2\.74 -F client_payload\[release_id\]=42/u);
  assert.doesNotMatch(dispatched, /release (?:create|edit|upload) /u);
  const skipped = await run("Dispatch verified publication", "false");
  assert.doesNotMatch(skipped, /api --method POST/u);
  assert.doesNotMatch(workflow, /DOCVIEWKIT_(?:APP_ID|APP_PRIVATE_KEY|RELEASE_TOKEN)|create-github-app-token|PUBLIC_REPOSITORY/u);
  assert.match(workflow, /if: github\.repository == 'docviewkit\/viewer'/u);
  assert.match(workflow, /actions\/upload-artifact@v7/u);
  const publish = await readFile(new URL("../.github/workflows/publish.yml", import.meta.url), "utf8");
  assert.match(publish, /repository_dispatch:\n    types: \[release-ready\]/u);
  assert.match(publish, /environment: npm-production/u);
  assert.match(publish, /id-token: write/u);
  assert.match(publish, /sha256sum -c PUBLIC_SHA256SUMS/u);
  assert.match(publish, /needs: \[publish, deploy-pages\]/u);
});

test("core workflows publish npm and Pages independently of website deployment", async () => {
  for (const path of ["ci.yml", "release.yml", "publish.yml"]) {
    const workflow = await readFile(new URL(`../.github/workflows/${path}`, import.meta.url), "utf8");
    assert.doesNotMatch(workflow, /commercial|official-site|DOCVIEWKIT_WEBSITE|DOCVIEWKIT_DEPLOY_WEBSITE|website-production/u);
  }
});

test("third-party license discovery is portable across case-sensitive file systems", async () => {
  const source = await readFile(
    new URL("../scripts/generate-third-party-notices.mjs", import.meta.url),
    "utf8",
  );
  assert.match(source, /readdir\(packageRoot, \{ withFileTypes: true \}\)/u);
  assert.match(source, /entry\.name\.toLowerCase\(\)/u);
  assert.match(source, /\["--no-default-features", "--features", "ofd-formats"\]/u);
});
