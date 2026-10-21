import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { chromium, webkit } from 'playwright-core';

const output = process.env.NUMBERS_TEAM_OUTPUT ?? 'output/playwright/numbers-team';
await mkdir(output, { recursive: true });
const server = spawn(process.execPath, ['scripts/serve.mjs', '--port', '0'], { stdio: ['ignore', 'pipe', 'inherit'] });
try {
  const url = await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('Server startup timed out')), 10000);
    createInterface({ input: server.stdout }).on('line', line => {
      const match = line.match(/https?:\/\/\S+/u);
      if (match) { clearTimeout(timer); resolve(match[0]); }
    });
    server.once('exit', code => reject(new Error(`Server exited: ${code}`)));
  });
  for (const [name, type] of Object.entries({ chromium, webkit })) {
    const browser = await type.launch({ headless: true });
    try {
      const page = await browser.newPage({ viewport: { width: 1150, height: 1050 }, reducedMotion: 'reduce' });
      const errors = [];
      page.on('pageerror', error => errors.push(error.message));
      page.on('console', message => { if (message.type() === 'error') errors.push(message.text()); });
      if (process.env.NUMBERS_TEAM_BASELINE_WASM) {
        const bytes = await readFile(process.env.NUMBERS_TEAM_BASELINE_WASM);
        await page.route('**/dist/office-viewer-pdf.wasm', route => route.fulfill({ contentType: 'application/wasm', body: bytes }));
      }
      await page.goto(`${url}examples/viewer.html?filePicker=true`);
      await page.waitForFunction(() => document.querySelector('#viewer').state.status === 'ready');
      await page.locator('#viewer').evaluate(async v => {
        await v.open(await (await fetch('/tests/fixtures/numbers-team-budget.numbers')).arrayBuffer());
        await v.reveal({ kind: 'unit', unitIndex: 3 });
      });
      await page.waitForFunction(() => {
        const v = document.querySelector('#viewer'), c = v.shadowRoot.querySelector('.surface');
        const ratio = c.width / parseFloat(c.style.width);
        const pixel = c.getContext('2d').getImageData(Math.round(10 * ratio), Math.round(680 * ratio), 1, 1).data;
        return v.state.currentUnitIndex === 3 && pixel[2] > pixel[0] + 50;
      });
      const result = await page.evaluate(async () => {
        const { createOfficeEngine } = await import('/dist/engine.js');
        const { extendedFormatPack } = await import('/dist/extended-formats.js');
        const engine = await createOfficeEngine({ execution: 'inline', formatPack: async () => extendedFormatPack });
        const doc = await engine.open(await (await fetch('/tests/fixtures/numbers-team-budget.numbers')).arrayBuffer());
        try {
          const objects = await doc.listObjects();
          const cells = Object.fromEntries(objects.filter(o => o.type === 'cell').map(o => [o.id, o.text]));
          const frames = [];
          const pixelChecks = {};
          for (const unitIndex of [1, 3]) {
            const unit = doc.info.units[unitIndex];
            const frame = await doc.render({ unitIndex, scale: 1, pixelRatio: 2, includeTextFragments: true });
            try {
              const canvas = document.createElement('canvas'); canvas.width = frame.pixelWidth; canvas.height = frame.pixelHeight;
              canvas.getContext('2d').drawImage(frame.bitmap, 0, 0);
              frames.push({ unitIndex, data: canvas.toDataURL(), fragments: frame.textFragments.map(f => f.text) });
              if (unitIndex === 3) {
                const pixel = (x, y) => Array.from(canvas.getContext('2d').getImageData(Math.round(x * 2), Math.round(y * 2), 1, 1).data);
                const paid = row => objects.find(o => o.id === `numbers:3:10258:numbers:3:B${row}`).bounds;
                const darkPixels = row => {
                  const b = paid(row), x = b.x + b.width / 2, y = b.y + b.height / 2;
                  const pixels = canvas.getContext('2d').getImageData(Math.round((x - 9) * 2), Math.round((y - 10) * 2), 36, 40).data;
                  let dark = 0; for (let i = 0; i < pixels.length; i += 4) if (pixels[i + 3] && Math.max(pixels[i], pixels[i + 1], pixels[i + 2]) < 90) dark++;
                  return dark;
                };
                pixelChecks.pie = pixel(530, 365);
                pixelChecks.checked = darkPixels(2); pixelChecks.unchecked = darkPixels(4);
                pixelChecks.stripe = pixel(300, paid(3).y + 3);
                pixelChecks.plain = pixel(300, paid(4).y + 3);
                const border = objects.find(o => o.id === 'numbers:3:10256:numbers:3:border:horizontal:1:0:5');
                pixelChecks.separator = border ? pixel(border.bounds.x + 450, border.bounds.y) : null;
              }
            } finally { frame.bitmap.close(); }
          }
          return { cells, frames, pixelChecks, diagnostics: doc.diagnostics() };
        } finally { doc.close(); engine.close(); }
      });
      await page.screenshot({ path: `${output}/${name}-budget.png` });
      for (const frame of result.frames) {
        await writeFile(`${output}/${name}-direct-${frame.unitIndex}.png`, Buffer.from(frame.data.split(',')[1], 'base64'));
        delete frame.data;
      }
      await writeFile(`${output}/${name}-result.json`, JSON.stringify(result, null, 2));
      const failures = [];
      const check = (condition, message) => { if (!condition) failures.push(message); };
      check(result.cells['numbers:3:10256:numbers:3:B2'] === '¥50', 'currency ¥50');
      check(result.cells['numbers:3:10258:numbers:3:B14'] === '¥556', 'currency ¥556');
      check(result.cells['numbers:1:10032:numbers:1:C2'] === '9:00 AM', 'time 9:00 AM');
      check(result.cells['numbers:1:10032:numbers:1:A2'] === '10/25/20', 'date 10/25/20');
      check(result.pixelChecks.pie[2] > 200 && result.pixelChecks.pie[0] < 30, 'full native pie plot');
      check(result.pixelChecks.checked > result.pixelChecks.unchecked + 30, 'checked/unchecked vector pixels');
      check(result.pixelChecks.separator?.[2] > result.pixelChecks.separator?.[0] + 25, 'blue header separator');
      check(result.pixelChecks.stripe[0] < result.pixelChecks.plain[0] - 10, 'alternating row fill');
      check(result.frames[1].fragments.includes('¥50'), 'currency text fragment');
      check(!result.frames[1].fragments.some(text => /^(TRUE|FALSE)$/.test(text)), 'checkbox must not paint boolean text');
      await page.locator('#viewer').evaluate(async v => {
        for (const unitIndex of [0, 1, 2, 3]) await v.reveal({ kind: 'unit', unitIndex });
      });
      if (failures.length === 0) {
      await page.locator('#viewer .zoom-input').fill('200');
      await page.locator('#viewer .zoom-input').press('Enter');
      await page.waitForFunction(() => document.querySelector('#viewer').state.zoom === 2);
      await page.waitForFunction(() => {
        const c = document.querySelector('#viewer').shadowRoot.querySelector('.surface');
        const ratio = c.width / parseFloat(c.style.width);
        const x = 240.6667 * 2 * ratio, y = 394.5265 * 2 * ratio;
        const pixels = c.getContext('2d').getImageData(Math.round(x - 18 * ratio), Math.round(y - 20 * ratio), Math.round(36 * ratio), Math.round(40 * ratio)).data;
        let dark = 0; for (let i = 0; i < pixels.length; i += 4) if (pixels[i + 3] && Math.max(pixels[i], pixels[i + 1], pixels[i + 2]) < 90) dark++;
        return dark > 100;
      });
      await page.screenshot({ path: `${output}/${name}-zoom.png` });
      }
      assert.deepEqual(errors, [], `${name}: browser errors`);
      assert.deepEqual(failures, [], `${name}: ${JSON.stringify(result.pixelChecks)}`);
      console.log(`${name}: real Numbers currency, controls, dates, pie, separator, banded-row pixels and 200% zoom passed`);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
