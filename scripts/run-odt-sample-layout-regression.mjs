import assert from 'node:assert/strict';
import { mkdir } from 'node:fs/promises';
import { chromium, firefox, webkit } from 'playwright-core';

const output = 'output/sample-odt-layout';
await mkdir(output, { recursive: true });
for (const [name, type] of Object.entries({ chromium, firefox, webkit })) {
  if (process.env.BROWSER && process.env.BROWSER !== name) continue;
  const browser = await type.launch({ headless: true, timeout: 30_000 });
  try {
    // Keep the whole page visible: a clipped viewport can leave the canvas bottom unpainted.
    const page = await browser.newPage({ viewport: { width: 1400, height: 2100 } });
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    page.on('console', message => { if (message.type() === 'error') errors.push(message.text()); });
    await page.addInitScript(() => {
      globalThis.painted = [];
      const ids = new WeakMap();
      let next = 0;
      for (const prototype of [CanvasRenderingContext2D.prototype, OffscreenCanvasRenderingContext2D.prototype]) {
        const original = prototype.fillText;
        prototype.fillText = function(text, x, y, ...rest) {
          if (!ids.has(this.canvas)) ids.set(this.canvas, next++);
          const transform = this.getTransform();
          const start = transform.transformPoint({ x, y });
          const end = transform.transformPoint({ x: x + this.measureText(text).width, y });
          const scale = Math.hypot(transform.a, transform.b);
          globalThis.painted.push({ canvas: ids.get(this.canvas), text, x: start.x / scale,
            y: start.y / scale, right: end.x / scale });
          return original.call(this, text, x, y, ...rest);
        };
      }
    });
    await page.route('**/examples/viewer.js', async route => {
      const response = await route.fetch();
      await route.fulfill({ response, body: (await response.text()).replace('engine: {', 'engine: { execution: "inline",') });
    });
    await page.goto(`${process.env.VIEWER_URL ?? 'http://127.0.0.1:4173'}/examples/viewer.html?fixture=file-sample_1MB.odt`);
    await page.waitForFunction(() => document.documentElement.dataset.ready === 'true');
    const count = await page.locator('#viewer').evaluate(viewer => viewer.state.unitCount);
    assert.equal(count, 5);
    for (let index = 0; index < count; index++) {
      await page.locator('#viewer').evaluate((viewer, unitIndex) => viewer.reveal({ kind: 'unit', unitIndex }), index);
      await page.locator(`#viewer .continuous-page[data-unit-index="${index}"] canvas`)
        .screenshot({ path: `${output}/${name}-${index}.png` });
    }
    const painted = await page.evaluate(() => globalThis.painted);
    const marker = painted.find(paint => paint.text === '•');
    const body = painted.find(paint => paint.canvas === marker?.canvas && Math.abs(paint.y - marker.y) < .1 && paint.text.startsWith('Maecenas'));
    assert.ok(marker && body);
    assert.ok(Math.abs(body.x - marker.x - 24) < .5, `${name}: authored list tab`);
    const lines = [];
    for (const paint of painted.filter(paint => paint.canvas === marker.canvas)) {
      let line = lines.at(-1);
      if (!line || Math.abs(line.y - paint.y) > .1 || paint.x < line.lastX - .1) {
        line = { text: '', y: paint.y, right: 0, lastX: 0 };
        lines.push(line);
      }
      line.text += paint.text;
      line.right = Math.max(line.right, paint.right);
      line.lastX = paint.x;
    }
    const first = lines.findIndex(line => line.text.startsWith('Vestibulum neque'));
    assert.ok(first >= 0);
    assert.ok(Math.abs(lines[first].right - lines[first + 1].right) < 1, `${name}: justified right edges`);
    const geometry = await page.evaluate(async () => {
      const { createOfficeEngine } = await import('/dist/engine.js');
      const engine = await createOfficeEngine({ execution: 'inline',
        formatPack: () => import('/dist/extended-formats.js').then(module => module.extendedFormatPack) });
      const document = await engine.open(new Uint8Array(await (await fetch('/tests/fixtures/file-sample_1MB.odt')).arrayBuffer()));
      try {
        const objects = await document.listObjects();
        const image = objects.find(object => object.type === 'image' && object.bounds.width > 600);
        const following = objects.find(object => object.unitIndex === image.unitIndex && object.text?.startsWith('Maecenas mauris'));
        const cell = text => objects.find(object => object.type === 'cell' && object.text === text);
        const body = objects.find(object => object.text?.startsWith('Vestibulum neque'));
        const next = objects.find(object => object.unitIndex === 0 && object.text?.startsWith('Maecenas mauris'));
        return { image, following, rows: [cell('1'), cell('2'), cell('3')].map(cell => cell.bounds.height), body, next };
      } finally { document.close(); engine.close(); }
    });
    assert.equal(geometry.image.unitIndex, 3);
    assert.ok(geometry.following.bounds.y >= geometry.image.bounds.y + geometry.image.bounds.height);
    assert.ok(Math.abs(geometry.rows[0] - geometry.rows[2]) < .1, `${name}: one-line table rows`);
    assert.ok(geometry.rows[1] > geometry.rows[0], `${name}: wrapped table row`);
    assert.ok(geometry.next.bounds.y > geometry.body.bounds.y + geometry.body.bounds.height, `${name}: measured body leaves paragraph spacing`);
    assert.deepEqual(errors, [], `${name}: clean console`);
    console.log(`${name}: 5 pages, justified body, 24px list tab, measured table rows, complete image on page 4`);
  } finally { await browser.close(); }
}
