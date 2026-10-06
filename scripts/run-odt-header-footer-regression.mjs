import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { chromium, firefox, webkit } from 'playwright-core';

const url = process.env.VIEWER_URL ?? 'http://127.0.0.1:4173';
await mkdir('output/3789', { recursive: true });
for (const [name, browserType] of Object.entries({ chromium, firefox, webkit })) {
  if (process.env.BROWSER && process.env.BROWSER !== name) continue;
  const browser = await browserType.launch({ headless: true, timeout: 30_000 });
  try {
    const page = await browser.newPage({ viewport: { width: 1300, height: 1100 } });
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    page.on('console', message => { if (message.type() === 'error') errors.push(message.text()); });
    await page.goto(`${url}/examples/viewer.html?fixture=corpus-header-footer-first.odt`);
    await page.waitForFunction(() => document.documentElement.dataset.ready === 'true');
    const boxes = await page.locator('#viewer').evaluate(viewer => {
      const canvas = viewer.shadowRoot.querySelector('.continuous-page[data-unit-index="0"] canvas') ?? viewer.shadowRoot.querySelector('.surface');
      const { data, width, height } = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height);
      return [[50, 205, 50], [128, 0, 0]].map(color => {
        const points = [];
        for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
          const i = (y * width + x) * 4;
          if (data[i + 3] > 240 && color.every((v, c) => Math.abs(v - data[i + c]) < 8)) points.push([x, y]);
        }
        if (!points.length) return null;
        const xs = points.map(p => p[0]), ys = points.map(p => p[1]);
        const left = Math.min(...xs), right = Math.max(...xs), top = Math.min(...ys), bottom = Math.max(...ys);
        return { left, right, top, bottom, edges: [
          points.filter(([x]) => x <= left + 2).length,
          points.filter(([x]) => x >= right - 2).length,
          points.filter(([, y]) => y <= top + 2).length,
          points.filter(([, y]) => y >= bottom - 2).length,
        ] };
      });
    });
    const png = await page.locator('#viewer .continuous-page[data-unit-index="0"] canvas').evaluate(canvas => canvas.toDataURL());
    await writeFile(`output/3789/${name}.png`, Buffer.from(png.split(',')[1], 'base64'));
    assert.ok(boxes.every(box => box && box.edges.every(count => count > 100)), `${name}: all four header/footer edges must paint: ${JSON.stringify(boxes)}`);
    assert.deepEqual(errors, [], `${name}: clean browser console`);
    console.log(name, JSON.stringify(boxes));
  } finally { await browser.close(); }
}
