import assert from 'node:assert/strict';
import { chromium, firefox, webkit } from 'playwright-core';

for (const name of (process.env.DOCVIEWKIT_BROWSERS ?? 'chromium,firefox,webkit').split(',')) {
  const browser = await ({ chromium, firefox, webkit })[name].launch({ headless: true, timeout: 30_000 });
  try {
    const page = await browser.newPage({ viewport: { width: 1200, height: 900 } });
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.goto(`${process.env.VIEWER_URL ?? 'http://127.0.0.1:4173'}/examples/viewer.html?fixture=tdf157636.ppt`);
    await page.waitForFunction(() => document.documentElement.dataset.ready === 'true');
    const result = await page.evaluate(async () => {
      const { Core, DEFAULT_LIMITS } = await import('/dist/core.js');
      const { SceneRenderer } = await import('/dist/render.js');
      const core = await Core.create(await (await fetch('/dist/office-viewer-legacy-office.wasm')).arrayBuffer(), DEFAULT_LIMITS);
      const doc = core.open(new Uint8Array(await (await fetch('/tests/fixtures/tdf157636.ppt')).arrayBuffer()));
      const object = doc.scene.objects.find(object => object.type === 'image');
      const effect = object.visual;
      const bitmap = await createImageBitmap(new Blob([effect.visual.bytes], { type: effect.visual.mediaType }));
      const canvas = new OffscreenCanvas(bitmap.width, bitmap.height);
      const ctx = canvas.getContext('2d');
      ctx.drawImage(bitmap, 0, 0);
      const original = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
      const unit = { ...doc.scene.info.units[0], width: bitmap.width, height: bitmap.height };
      const renderer = new SceneRenderer([{ ...object, bounds: { x: 0, y: 0, width: bitmap.width, height: bitmap.height } }], DEFAULT_LIMITS);
      try {
        const frame = await renderer.render(unit, { unitIndex: 0, background: '#000000' });
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(frame.bitmap, 0, 0);
        const actual = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
        let nearWhite = 0, leftover = 0, dark = 0;
        for (let i = 0; i < original.length; i += 4) {
          const min = Math.min(...original.slice(i, i + 3));
          if (min >= 246 && min < 255) {
            nearWhite++;
            if (actual[i] > 240 && actual[i + 1] > 240 && actual[i + 2] > 240) leftover++;
          }
          if (Math.max(...original.slice(i, i + 3)) < 180) {
            dark++;
            for (let j = 0; j < 3; j++) if (Math.abs(actual[i + j] - original[i + j]) > 1) throw new Error('Logo detail changed');
          }
        }
        frame.bitmap.close();
        return { width: bitmap.width, height: bitmap.height, nearWhite, leftover, dark, diagnostics: frame.diagnostics };
      } finally { renderer.close(); bitmap.close(); doc.close(); }
    });
    await page.screenshot({ path: `/tmp/tdf157636-${name}.png` });
    console.log(name, result);
    assert.ok(result.nearWhite > 1000 && result.dark > 1000);
    assert.equal(result.leftover, 0, 'JPEG near-white color-key pixels must not leave white blocks');
    assert.deepEqual(result.diagnostics, []);
    assert.deepEqual(errors, []);
  } finally { await browser.close(); }
}
