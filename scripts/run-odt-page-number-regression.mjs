import assert from 'node:assert/strict';
import { mkdir } from 'node:fs/promises';
import { chromium, firefox, webkit } from 'playwright-core';

await mkdir('output/3923', { recursive: true });
for (const [name, type] of Object.entries({ chromium, firefox, webkit })) {
  if (process.env.BROWSER && process.env.BROWSER !== name) continue;
  const browser = await type.launch({ headless: true });
  try {
    const page = await browser.newPage({ viewport: { width: 1400, height: 1100 } });
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    page.on('console', message => { if (message.type() === 'error') errors.push(message.text()); });
    await page.addInitScript(() => {
      for (const prototype of [CanvasRenderingContext2D.prototype, OffscreenCanvasRenderingContext2D.prototype]) {
        const fillText = prototype.fillText;
        prototype.fillText = function(text, ...args) {
          (this.canvas.paintedText ??= []).push({ text, font: this.font });
          return fillText.call(this, text, ...args);
        };
        const drawImage = prototype.drawImage;
        prototype.drawImage = function(source, ...args) {
          if (source.paintedText) this.canvas.paintedText = [...source.paintedText];
          return drawImage.call(this, source, ...args);
        };
      }
      const createBitmap = globalThis.createImageBitmap;
      globalThis.createImageBitmap = async function(source, ...args) {
        const bitmap = await createBitmap(source, ...args);
        bitmap.paintedText = source.paintedText;
        return bitmap;
      };
      const transfer = OffscreenCanvas.prototype.transferToImageBitmap;
      OffscreenCanvas.prototype.transferToImageBitmap = function() {
        const bitmap = transfer.call(this);
        bitmap.paintedText = this.paintedText;
        return bitmap;
      };
    });
    // Inline execution makes the real Canvas text calls observable in this page.
    await page.route('**/examples/viewer.js', async route => {
      const response = await route.fetch();
      await route.fulfill({ response, body: (await response.text()).replace('engine: {', 'engine: { execution: "inline",') });
    });
    await page.goto(`${process.env.VIEWER_URL ?? 'http://127.0.0.1:4173'}/examples/viewer.html?fixture=oasis-3923-page-number-zero.odt`);
    await page.waitForFunction(() => document.documentElement.dataset.ready === 'true');
    const state = await page.locator('#viewer').evaluate(viewer => viewer.state);
    assert.equal(state.unitCount, 4, `${name}: physical page count`);
    for (let index = 0; index < 4; index++) {
      await page.locator('#viewer').evaluate((viewer, unitIndex) => viewer.reveal({ kind: 'unit', unitIndex }), index);
      const selector = `#viewer .continuous-page[data-unit-index="${index}"] canvas`;
      await page.waitForFunction(index => {
        const root = document.querySelector('#viewer').shadowRoot;
        return root.querySelector(`.continuous-page[data-unit-index="${index}"] canvas`)?.paintedText?.length > 0;
      }, index);
      const painted = await page.locator(selector).evaluate(canvas => canvas.paintedText);
      const number = String(index % 2);
      const header = painted.map(item => item.text).join('');
      assert.ok(header.includes(`Seite ${number} von 4`), `${name} page ${index}: ${JSON.stringify(painted)}`);
      if (index < 3) {
        assert.ok(painted.some(item => item.text === number && item.font.includes('53.333')), `${name}: large body number ${number}`);
      }
      assert.ok(painted.every(item => !/[\uE000\uE001]/u.test(item.text)), `${name}: no unresolved fields`);
      await page.locator(selector).screenshot({ path: `output/3923/${name}-${index}.png` });
    }
    assert.ok(state.info.units[0].height > state.info.units[0].width);
    assert.ok(state.info.units[2].width > state.info.units[2].height);
    assert.deepEqual(errors, [], `${name}: clean console`);
    console.log(`${name}: 4 pages, headers 0/1/0/1 of 4, body 0/1/0, portrait then landscape`);
  } finally { await browser.close(); }
}
