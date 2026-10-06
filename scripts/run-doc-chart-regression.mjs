import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { chromium, firefox, webkit } from 'playwright-core';

const output = 'output/sample-doc';
await mkdir(output, { recursive: true });
for (const [name, type] of Object.entries({ chromium, firefox, webkit })) {
  if (process.env.BROWSER && process.env.BROWSER !== name) continue;
  const browser = await type.launch({ headless: true, timeout: 30_000 });
  try {
    const page = await browser.newPage({ viewport: { width: 1400, height: 1900 } });
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    page.on('console', message => { if (message.type() === 'error') errors.push(message.text()); });
    await page.goto(`${process.env.VIEWER_URL ?? 'http://127.0.0.1:4173'}/examples/viewer.html?fixture=file-sample_1MB.doc`);
    await page.waitForFunction(() => document.documentElement.dataset.ready === 'true');
    const geometry = await page.evaluate(async () => {
      const { createOfficeEngine } = await import('/dist/engine.js');
      const engine = await createOfficeEngine({ execution: 'inline',
        formatPack: () => import('/dist/extended-formats.js').then(module => module.extendedFormatPack) });
      const doc = await engine.open(new Uint8Array(await (await fetch('/tests/fixtures/file-sample_1MB.doc')).arrayBuffer()));
      try {
        return { info: doc.info, objects: await doc.listObjects(), diagnostics: doc.diagnostics() };
      } finally { doc.close(); engine.close(); }
    });
    const canvas = page.locator('#viewer .continuous-page[data-unit-index="0"] canvas');
    await canvas.screenshot({ path: `${output}/${name}-page-1.png` });
    const objects = geometry.objects.filter(object => object.id.startsWith('doc:chart:'));
    for (const label of ['Column 1', 'Column 2', 'Column 3', 'Row 1', 'Row 2', 'Row 3', 'Row 4']) {
      assert.ok(objects.some(object => object.text === label), `${name}: ${label}`);
    }
    const pixels = await canvas.evaluate((canvas, { objects, width, height }) => {
      const context = canvas.getContext('2d');
      return objects.filter(object => object.type === 'shape' && object.bounds.width > 10 && object.bounds.height > 10)
        .map(({ bounds }) => {
          const x = Math.floor((bounds.x + bounds.width / 2) * canvas.width / width);
          const y = Math.floor((bounds.y + bounds.height / 2) * canvas.height / height);
          return [...context.getImageData(x, y, 1, 1).data].slice(0, 3).join(',');
        });
    }, { objects, ...geometry.info.units[0] });
    for (const color of ['0,69,134', '255,66,14', '255,211,32']) {
      assert.equal(pixels.filter(pixel => pixel === color).length, 4, `${name}: four painted bars ${color}`);
    }
    assert.equal(geometry.objects.filter(object => object.type === 'image').length, 1, 'only the authored photograph remains');
    assert.equal(geometry.info.units.length, 5);
    assert.deepEqual(errors, [], `${name}: clean console`);
    await writeFile(`${output}/${name}.json`, JSON.stringify({ ...geometry, pixels, errors }, null, 2));
    console.log(`${name}: 12 painted bars, 7 labels, 5 pages, no chart preview, clean console`);
  } finally { await browser.close(); }
}
