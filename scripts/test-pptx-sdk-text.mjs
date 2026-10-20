// Exact supplied SDK overview: compare title visibility and line counts with
// PowerPoint's export. Original overlap in the lower-right footnote is authored.
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createInterface } from "node:readline";
import { writeFile } from "node:fs/promises";
import { chromium, webkit } from "playwright-core";

const server = spawn(process.execPath, ["scripts/serve.mjs", "--port", "0"], { stdio: ["ignore", "pipe", "inherit"] });
try {
  const url = await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.once("exit", code => reject(new Error(`Server exited: ${code}`)));
    createInterface({ input: server.stdout }).on("line", line => {
      const match = line.match(/https?:\/\/\S+/u);
      if (match) resolve(match[0]);
    });
  });
  for (const [name, type] of Object.entries({ chromium, webkit })) {
    const browser = await type.launch({ headless: true });
    try {
      const page = await browser.newPage({ viewport: { width: 1400, height: 1000 } });
      const errors = [];
      page.on("pageerror", error => errors.push(error.message));
      page.on("console", message => { if (message.type() === "error") errors.push(message.text()); });
      if (process.env.SDK_TEXT_WASM) await page.route("**/dist/office-viewer-core.wasm", route => route.fulfill({ path: process.env.SDK_TEXT_WASM, contentType: "application/wasm" }));
      if (process.env.SDK_TEXT_FONT) await page.route("**/dist/font.js", route => route.fulfill({ path: process.env.SDK_TEXT_FONT, contentType: "text/javascript" }));
      await page.goto(`${url}examples/viewer.html?fixture=sdk-development-arrows.pptx`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
      assert.equal(await page.title(), "DocViewKit Viewer");
      await page.evaluate(() => document.querySelector("#viewer").reveal({ kind: "unit", unitIndex: 3 }));
      await page.waitForFunction(() => document.querySelector("#viewer").shadowRoot.querySelector('.continuous-page[data-unit-index="3"]')?.dataset.rendered === "true");
      const png = await page.evaluate(() => document.querySelector("#viewer").shadowRoot.querySelector('.continuous-page[data-unit-index="3"] canvas').toDataURL().split(",")[1]);
      await writeFile(`/tmp/sdk-text-${name}.png`, Buffer.from(png, "base64"));
      const checks = await page.evaluate(async () => {
        const { createOfficeEngine } = await import("/dist/engine.js");
        const engine = await createOfficeEngine({ execution: "inline", wasm: await (await fetch("/dist/office-viewer-core.wasm")).arrayBuffer() });
        const doc = await engine.open(new Uint8Array(await (await fetch("/tests/fixtures/sdk-development-arrows.pptx")).arrayBuffer()));
        const checks = [];
        try {
          for (const unitIndex of [3, 4]) {
            const objects = await doc.listObjects({ unitIndex });
            const frame = await doc.render({ unitIndex, scale: 1, includeTextFragments: true });
            try {
              const expected = unitIndex === 3
                ? [["2. system configuration", 1], ["Server Software", 1], ["Server Software\n(DN provides parts)", 2], ["Digikey SDK*", 1], ["Legend:\nDA: Device Agent; \nDNST: Denso China;\nDNJP: Denso JP", 4]]
                : [["3. Development and operation system", 1], ["Legend:\nDNST: Denso China;\nDNJP: Denso JP", 3]];
              for (const [text, lines] of expected) {
                const matches = objects.filter(object => object.text === text);
                if (!matches.length) throw new Error(`Missing original object: ${text}`);
                for (const object of matches) {
                  const fragments = frame.textFragments.filter(fragment => fragment.objectId === object.id && fragment.text.trim());
                  checks.push({ unitIndex, text, expectedLines: lines, lines: new Set(fragments.map(fragment => fragment.line)).size,
                    left: Math.min(...fragments.map(fragment => fragment.transform.e)) });
                }
              }
            } finally { frame.bitmap.close(); }
          }
          return checks;
        } finally { doc.close(); engine.close(); }
      });
      console.log(name, JSON.stringify({ checks, errors }));
      assert.deepEqual(errors, []);
      for (const check of checks) {
        assert.equal(check.lines, check.expectedLines, `${name}: ${check.text}`);
        if (/^[23]\./u.test(check.text)) assert.ok(check.left >= 15, `${name}: title is fully inside the page`);
      }
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
