import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { mkdir, writeFile } from 'node:fs/promises';
import { createInterface } from 'node:readline';
import { chromium, firefox, webkit } from 'playwright-core';

// Original Microsoft Open XML SDK corpus files, byte-for-byte copies.
const cases = [['auto', 431.6 * 4 / 3, 431.6, true], ['50', 512, 384, false], ['150', 1536, 1152, true]];
const server = spawn(process.execPath, ['scripts/serve.mjs', '--port', '0']);
const url = await new Promise((resolve, reject) => {
  createInterface({ input: server.stdout }).on('line', line => {
    const match = line.match(/http:\/\/\S+/); if (match) resolve(match[0]);
  });
  server.once('exit', code => reject(new Error(`Server exited: ${code}`)));
});
try {
  await mkdir('.cache/picture-watermark', { recursive: true });
  for (const name of (process.env.DOCVIEWKIT_BROWSERS ?? 'chromium,firefox,webkit').split(',')) {
    const browser = await ({ chromium, firefox, webkit })[name].launch({ timeout: 15000 });
    try {
      const page = await browser.newPage({ viewport: { width: 1100, height: 1250 } });
      const errors = [];
      page.on('pageerror', error => errors.push(error.message));
      page.on('console', message => { if (message.type() === 'error') errors.push(message.text()); });
      for (const [id, width, height, washout] of cases) {
        const fixture = `picture-watermark-${id}.docx`;
        await page.goto(`${url}examples/viewer.html?fixture=${fixture}&theme=light`);
        await page.waitForFunction(() => document.documentElement.dataset.ready === 'true', null, { timeout: 30000 });
        await page.screenshot({ path: `.cache/picture-watermark/${name}-${id}-viewer.png` });
        const actual = await page.evaluate(async ({ fixture, washout }) => {
          const { Core, DEFAULT_LIMITS } = await import('/dist/core.js');
          const { SceneRenderer } = await import('/dist/render.js');
          const core = await Core.create(await (await fetch('/dist/office-viewer-core.wasm')).arrayBuffer(), DEFAULT_LIMITS);
          const doc = core.open(new Uint8Array(await (await fetch(`/tests/fixtures/${fixture}`)).arrayBuffer()));
          const scene = doc.loadUnit(0)?.scene ?? doc.scene;
          const object = scene.objects.find(o => o.type === 'image' && o.source.part === 'word/header2.xml');
          const renderer = new SceneRenderer(scene.objects, DEFAULT_LIMITS);
          try {
            const unit = scene.info.units[0];
            const frame = await renderer.render(unit, { unitIndex: 0, viewport: { x: 0, y: 0, width: unit.width, height: unit.height }, background: '#ffffff' });
            const canvas = document.createElement('canvas'); canvas.width = frame.pixelWidth; canvas.height = frame.pixelHeight;
            const ctx = canvas.getContext('2d'); ctx.drawImage(frame.bitmap, 0, 0); frame.bitmap.close();
            let visual = object.visual;
            while (visual.visual) visual = visual.visual;
            const bitmap = await createImageBitmap(new Blob([visual.bytes], { type: visual.mediaType }));
            const source = document.createElement('canvas'); source.width = bitmap.width; source.height = bitmap.height;
            const sourceCtx = source.getContext('2d'); sourceCtx.drawImage(bitmap, 0, 0); bitmap.close();
            if (washout) {
              const pixels = sourceCtx.getImageData(0, 0, source.width, source.height);
              // Word's native PDF image gives RGB ~= min(255, 0.3 * RGB + 205.275).
              for (let i = 0; i < pixels.data.length; i++) if (i % 4 !== 3) pixels.data[i] = Math.min(255, .3 * pixels.data[i] + 205.275);
              sourceCtx.putImageData(pixels, 0, 0);
            }
            const expected = document.createElement('canvas'); expected.width = canvas.width; expected.height = canvas.height;
            const ec = expected.getContext('2d'); ec.fillStyle = '#fff'; ec.fillRect(0, 0, expected.width, expected.height);
            const b = object.bounds; ec.drawImage(source, b.x, b.y, b.width, b.height);
            const a = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
            const e = ec.getImageData(0, 0, canvas.width, canvas.height).data;
            let error = 0, count = 0;
            // Exclude the document's introductory text, retain image edges and page clipping.
            for (let y = 200; y < canvas.height; y += 3) for (let x = 0; x < canvas.width; x += 3) {
              const offset = (y * canvas.width + x) * 4;
              for (let c = 0; c < 3; c++) { error += Math.abs(a[offset + c] - e[offset + c]); count++; }
            }
            return { bounds: b, page: { width: unit.width, height: unit.height }, meanError: error / count, png: canvas.toDataURL() };
          } finally { renderer.close(); doc.close(); core.close(); }
        }, { fixture, washout });
        await writeFile(`.cache/picture-watermark/${name}-${id}.png`, Buffer.from(actual.png.split(',')[1], 'base64'));
        assert.ok(Math.abs(actual.bounds.width - width) < .01 && Math.abs(actual.bounds.height - height) < .01, `${fixture}: ${JSON.stringify(actual.bounds)}`);
        assert.ok(Math.abs(actual.bounds.x + width / 2 - actual.page.width / 2) < .01);
        assert.ok(Math.abs(actual.bounds.y + height / 2 - actual.page.height / 2) < .01);
        assert.ok(actual.meanError < 1, `${fixture}: pixel mean error ${actual.meanError}`);
        assert.deepEqual(errors, []);
        console.log(`${name}: ${fixture}: scale, center, washout, clipping passed; pixel MAE ${actual.meanError.toFixed(3)}`);
      }
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
