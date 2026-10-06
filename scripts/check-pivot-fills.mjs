import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { mkdir } from 'node:fs/promises';
import { createInterface } from 'node:readline';
import { chromium, webkit, firefox } from 'playwright-core';

const cases = [
  ['Pivot.xlsx', 0, { A3: 'dce6f1', B3: 'dce6f1', A4: 'dce6f1', A5: 'ffffff', B5: 'ffffff' }],
  ['Pivot2.xlsx', 0, { A3: '366092', AB5: '366092', A6: '95b3d7', B7: '95b3d7',
    A8: 'dce6f1', T8: 'dce6f1', B9: 'ffffff', A35: 'ffffff', AB35: 'ffffff' }],
  ['Pivot2.xlsx', 1, { B5: 'dce6f1', C6: 'dce6f1', B7: 'ffffff', C83: 'ffffff' }],
  ['Pivot2.xlsx', 2, { B2: 'dce6f1', F3: 'dce6f1', B4: 'ffffff', D5: 'ffffff', B7: 'dce6f1' }],
];
const server = spawn(process.execPath, ['scripts/serve.mjs', '--port', '0']);
const url = await new Promise((resolve, reject) => {
  createInterface({ input: server.stdout }).on('line', line => {
    const match = line.match(/http:\/\/\S+/); if (match) resolve(match[0]);
  });
  server.once('exit', code => reject(new Error(`Server exited: ${code}`)));
});
try {
  await mkdir('.cache/pivot-fills', { recursive: true });
  for (const name of (process.env.DOCVIEWKIT_BROWSERS ?? 'chromium,webkit').split(',')) {
    const browser = await ({ chromium, webkit, firefox })[name].launch({ timeout: 15000 });
    try {
      const page = await browser.newPage({ viewport: { width: 1440, height: 1000 }, reducedMotion: 'reduce' });
      const errors = [];
      page.on('pageerror', error => errors.push(error.message));
      page.on('console', message => { if (message.type() === 'error') errors.push(message.text()); });
      for (const [fixture, unitIndex, expected] of cases) {
        await page.goto(`${url}examples/viewer.html?fixture=${fixture}`);
        await page.waitForFunction(() => document.documentElement.dataset.ready === 'true', null, { timeout: 15000 });
        const viewer = page.locator('docviewkit-viewer');
        if (unitIndex) {
          await viewer.evaluate((v, index) => new Promise(resolve => {
            v.addEventListener('docviewkit-diagnostic', resolve, { once: true });
            v.shadowRoot.querySelectorAll('.sheet-tab')[index].click();
          }), unitIndex);
        }
        await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
        await page.screenshot({ path: `.cache/pivot-fills/${name}-${fixture}-${unitIndex}.png` });
        const actual = await page.evaluate(async ({ fixture, unitIndex, addresses }) => {
          const { Core, DEFAULT_LIMITS } = await import('/dist/core.js');
          const { SceneRenderer } = await import('/dist/render.js');
          const core = await Core.create(await (await fetch('/dist/office-viewer-core.wasm')).arrayBuffer(), DEFAULT_LIMITS);
          const doc = core.open(new Uint8Array(await (await fetch(`/tests/fixtures/${fixture}`)).arrayBuffer()));
          const scene = doc.loadUnit(unitIndex).scene;
          const renderer = new SceneRenderer(scene.objects, DEFAULT_LIMITS);
          try {
            const colors = {};
            for (const address of addresses) {
              const cell = scene.objects.find(object => object.type === 'cell' && object.source.address === address);
              if (!cell) { colors[address] = 'missing'; continue; }
              const frame = await renderer.render(scene.info.units[unitIndex], { unitIndex, viewport: cell.bounds, background: '#ffffff' });
              const canvas = new OffscreenCanvas(frame.pixelWidth, frame.pixelHeight);
              const context = canvas.getContext('2d'); context.drawImage(frame.bitmap, 0, 0);
              const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data;
              const counts = new Map();
              for (let i = 0; i < pixels.length; i += 4) {
                const rgb = ((pixels[i] << 16) | (pixels[i + 1] << 8) | pixels[i + 2]).toString(16).padStart(6, '0');
                counts.set(rgb, (counts.get(rgb) ?? 0) + 1);
              }
              colors[address] = [...counts].sort((a, b) => b[1] - a[1])[0][0];
              frame.bitmap.close();
            }
            return colors;
          } finally { renderer.close(); doc.close(); core.close(); }
        }, { fixture, unitIndex, addresses: Object.keys(expected) });
        for (const [address, color] of Object.entries(expected)) {
          assert.match(actual[address], /^[0-9a-f]{6}$/, `${fixture}: missing ${address}`);
          // Excel and the shared HSL conversion can round theme channels by one.
          assert.ok([0, 2, 4].every(offset => Math.abs(parseInt(actual[address].slice(offset, offset + 2), 16)
            - parseInt(color.slice(offset, offset + 2), 16)) <= 1), `${name}: ${fixture} sheet ${unitIndex + 1} ${address}: ${actual[address]} != ${color}`);
        }
        assert.deepEqual(errors, []);
        console.log(`${name}: ${fixture} sheet ${unitIndex + 1} actual fill pixels passed`);
      }
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
