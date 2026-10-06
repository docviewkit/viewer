import assert from 'node:assert/strict';
import { mkdir } from 'node:fs/promises';
import { chromium, firefox, webkit } from 'playwright-core';

await mkdir('output/comment001', { recursive: true });
for (const [name, type] of Object.entries({ chromium, firefox, webkit })) {
  if (process.env.BROWSER && process.env.BROWSER !== name) continue;
  const timeout = setTimeout(() => { console.error(`${name}: exceeded 60 seconds`); process.exit(1); }, 60_000);
  const browser = await type.launch({ headless: true, timeout: 20_000 });
  try {
    const page = await browser.newPage({ viewport: { width: 1200, height: 1200 } });
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    page.on('console', message => { if (message.type() === 'error') errors.push(message.text()); });
    await page.addInitScript(() => {
      globalThis.calendarText = [];
      for (const prototype of [CanvasRenderingContext2D.prototype, OffscreenCanvasRenderingContext2D.prototype]) {
        const original = prototype.fillText;
        prototype.fillText = function(text, x, y, ...rest) {
          globalThis.calendarText.push({ text, font: this.font, color: this.fillStyle, x, y });
          return original.call(this, text, x, y, ...rest);
        };
      }
    });
    await page.route('**/examples/viewer.js', async route => {
      const response = await route.fetch();
      await route.fulfill({ response, body: (await response.text()).replace('engine: {', 'engine: { execution: "inline",') });
    });
    await page.goto(`${process.env.VIEWER_URL ?? 'http://127.0.0.1:4173'}/examples/viewer.html?fixture=word-calendar-comments.docx`);
    await page.waitForFunction(() => document.documentElement.dataset.ready === 'true');
    await page.waitForFunction(() => globalThis.calendarText.some(run => run.text.includes('MAY')));
    const painted = await page.evaluate(() => globalThis.calendarText);
    for (const text of ['MAY', 'M', '1']) {
      const run = painted.find(run => run.text === text);
      assert.ok(run, `${name}: missing ${text}`);
      assert.equal(run.color, text === '1' ? '#000000' : '#4f81bd');
      const size = Number(run.font.match(/([\d.]+)px/)[1]);
      assert.ok(Math.abs(size - (text === 'MAY' ? 20 : text === 'M' ? 16 : 14) * 4 / 3) < 0.01);
      assert.ok((text === '1' ? /Calibri|Carlito/ : /Cambria|Caladea/).test(run.font), JSON.stringify(run));
    }
    assert.deepEqual(errors, []);
    await page.screenshot({ path: `output/comment001/${name}.png`, fullPage: true });
    console.log(`${name}: calendar header/date fonts and colors passed; clean console`);
  } finally {
    await browser.close();
    clearTimeout(timeout);
  }
}
