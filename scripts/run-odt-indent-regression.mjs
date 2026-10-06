import assert from 'node:assert/strict';
import { mkdir } from 'node:fs/promises';
import { chromium, firefox, webkit } from 'playwright-core';

await mkdir('output/3937', { recursive: true });
for (const [name, type] of Object.entries({ chromium, firefox, webkit })) {
  if (process.env.BROWSER && process.env.BROWSER !== name) continue;
  const browser = await type.launch({ headless: true, timeout: 30_000 });
  try {
    const page = await browser.newPage({ viewport: { width: 1400, height: 1100 } });
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    page.on('console', message => { if (message.type() === 'error') errors.push(message.text()); });
    await page.addInitScript(() => {
      globalThis.paintedLines = [];
      for (const prototype of [CanvasRenderingContext2D.prototype, OffscreenCanvasRenderingContext2D.prototype]) {
        const fillText = prototype.fillText;
        prototype.fillText = function(text, x, y, ...rest) {
          const point = this.getTransform().transformPoint({ x, y });
          const lines = globalThis.paintedLines;
          let line = lines.at(-1);
          if (!line || line.canvas !== this.canvas || Math.abs(line.y - point.y) > 0.1 || point.x < line.lastX) {
            line = { canvas: this.canvas, text: '', x: point.x, y: point.y };
            lines.push(line);
          }
          line.text += text;
          line.lastX = point.x;
          return fillText.call(this, text, x, y, ...rest);
        };
      }
    });
    await page.route('**/examples/viewer.js', async route => {
      const response = await route.fetch();
      await route.fulfill({ response, body: (await response.text()).replace('engine: {', 'engine: { execution: "inline",') });
    });
    await page.goto(`${process.env.VIEWER_URL ?? 'http://127.0.0.1:4173'}/examples/viewer.html?fixture=oasis-3937-background-border.odt`);
    await page.waitForFunction(() => document.documentElement.dataset.ready === 'true');
    await page.waitForFunction(() => globalThis.paintedLines.some(line => line.text.startsWith('Oder gehörten')));
    const lines = await page.evaluate(() => globalThis.paintedLines.map(({text, x, y}) => ({text, x, y})));
    const plain = lines.find(line => line.text.startsWith('Er hörte'));
    const index = lines.findIndex(line => line.text.startsWith('Oder gehörten'));
    const first = lines[index], second = lines[index + 1];
    console.log(name, JSON.stringify({ plain, first, second }));
    await page.locator('#viewer .continuous-page[data-unit-index="0"] canvas').screenshot({ path: `output/3937/${name}.png` });
    assert.ok(Math.abs(second.x - plain.x) < 0.5, `${name}: continuation line must align with the previous paragraph`);
    assert.ok(Math.abs(first.x - second.x - plain.x * 4.99 / 20) < 0.5, `${name}: first line must indent exactly 4.99 mm once`);
    assert.deepEqual(errors, [], `${name}: clean console`);
  } finally { await browser.close(); }
}
