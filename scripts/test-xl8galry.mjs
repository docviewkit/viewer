import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { chromium, firefox, webkit } from 'playwright-core';

const output = process.env.XL8GALRY_OUTPUT ?? 'output/xl8galry/verified';
await mkdir(output, { recursive: true });
const server = spawn(process.execPath, ['scripts/serve.mjs', '--port', '0'], { stdio: ['ignore', 'pipe', 'inherit'] });
try {
  const url = await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('Viewer server timeout')), 10000);
    createInterface({ input: server.stdout }).on('line', line => {
      const match = line.match(/https?:\/\/\S+/u);
      if (match) { clearTimeout(timer); resolve(match[0]); }
    });
  });
  for (const name of (process.env.DOCVIEWKIT_BROWSERS ?? 'chromium,firefox,webkit').split(',')) {
    const browser = await ({ chromium, firefox, webkit })[name].launch({ headless: true, timeout: 30000 });
    try {
      const page = await browser.newPage({ viewport: { width: 1200, height: 800 }, reducedMotion: 'reduce' });
      const errors = [];
      page.on('pageerror', e => errors.push(e.message));
      page.on('console', e => { if (e.type() === 'error') errors.push(e.text()); });
      if (process.env.XL8GALRY_WASM) await page.route('**/dist/office-viewer-core.wasm', route => route.fulfill({ path: process.env.XL8GALRY_WASM, contentType: 'application/wasm' }));
      await page.goto(`${url}examples/viewer.html?fixture=xl8galry.xlsx`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === 'true');
      const info = await page.evaluate(() => document.querySelector('#viewer').state.info);
      assert.equal(info.units.length, 20);
      const results = [];
      for (const unit of info.units) {
        await page.evaluate(index => document.querySelector('#viewer').reveal({ kind: 'unit', unitIndex: index }), unit.index);
        await page.locator('.page-transition-outgoing').waitFor({ state: 'detached' });
        await page.screenshot({ path: `${output}/${name}-${String(unit.index + 1).padStart(2, '0')}.png` });
        const result = await page.evaluate(async index => {
          const { Core, DEFAULT_LIMITS } = await import('/dist/core.js');
          const { SceneRenderer } = await import('/dist/render.js');
          const core = await Core.create(await (await fetch('/dist/office-viewer-core.wasm')).arrayBuffer(), DEFAULT_LIMITS);
          const doc = core.open(new Uint8Array(await (await fetch('/tests/fixtures/xl8galry.xlsx')).arrayBuffer()));
          doc.loadUnit(index);
          const objects = doc.scene.objects.filter(o => o.unitIndex === index);
          if (index === 0) {
            const background = objects.find(o => { let v = o.visual; while (v.visual) v = v.visual; let p = v.fill; while (p?.paint) p = p.paint; return p?.kind === 'rect-gradient'; });
            if (!background) throw new Error('Sheet1 rectangular background missing');
            const backgroundRenderer = new SceneRenderer([background], DEFAULT_LIMITS);
            try {
              const frame = await backgroundRenderer.render(doc.scene.info.units[index], { unitIndex: index, scale: 1 });
              const canvas = new OffscreenCanvas(frame.bitmap.width, frame.bitmap.height);
              const context = canvas.getContext('2d');
              context.drawImage(frame.bitmap, 0, 0);
              const { x, y, width, height } = background.bounds;
              const pixel = context.getImageData(Math.floor(x + width * 0.9), Math.floor(y + height * 0.5), 1, 1).data;
              // Source stops are #339966 / #18472F; halfway from the top-right focus.
              if (Math.abs(pixel[1] - 112) > 2) throw new Error(`Sheet1 gradient midpoint too dark: ${[...pixel]}`);
              frame.bitmap.close();
            } finally { backgroundRenderer.close(); }
          }
          const renderer = new SceneRenderer(objects, DEFAULT_LIMITS);
          try {
            const frame = await renderer.render(doc.scene.info.units[index], { unitIndex: index, scale: 1 });
            if (index === 10) {
              const canvas = new OffscreenCanvas(frame.bitmap.width, frame.bitmap.height);
              const context = canvas.getContext('2d');
              context.drawImage(frame.bitmap, 0, 0);
              // Interior pixels at the two overlapping side faces in the real
              // exploded pie: the nearer Mountain/West faces must stay visible.
              const mountain = context.getImageData(184, 279, 1, 1).data;
              const west = context.getImageData(494, 268, 1, 1).data;
              if (!(mountain[0] > mountain[2] * 1.8 && Math.abs(mountain[1] - mountain[2]) <= 2)) throw new Error(`Pacific cuts through Mountain: ${[...mountain]}`);
              if (!(west[1] > west[0] * 1.15 && Math.abs(west[1] - west[2]) <= 2)) throw new Error(`East cuts through West: ${[...west]}`);
            }
            frame.bitmap.close();
            const visuals = objects.map(o => { let v = o.visual; while (v.visual) v = v.visual; return v; });
            return { count: objects.length, labels: objects.filter(o => o.text).map(o => o.text), paths: visuals.filter(v => v.geometry?.kind === 'path').length, patterns: visuals.filter(v => v.fill?.kind === 'pattern').map(v => v.fill.preset), diagnostics: frame.diagnostics };
          } finally { renderer.close(); doc.close(); core.close(); }
        }, unit.index);
        assert.ok(result.count >= 5, `${unit.name}: chart content`);
        assert.deepEqual(result.diagnostics, [], `${unit.name}: rendering diagnostics`);
        if (unit.index === 0) assert.ok(result.labels.includes('1189.679009'), 'General label precision');
        if (unit.index === 10) assert.ok(result.paths >= 24, '3D pie side faces');
        if (unit.index === 16) assert.ok(result.patterns.includes('pct75'), 'point pattern fill');
        assert.deepEqual(errors, [], `${unit.name}: browser errors`);
        results.push({ sheet: unit.name, ...result });
      }
      await writeFile(`${output}/${name}.json`, JSON.stringify(results, null, 2));
      console.log(`${name}: 20/20 sheets, no console errors or render diagnostics`);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
