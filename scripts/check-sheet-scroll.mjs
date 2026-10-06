import assert from 'node:assert/strict';
import { spawn, execFileSync } from 'node:child_process';
import { createInterface } from 'node:readline';
import { readFile, mkdir } from 'node:fs/promises';
import { chromium, firefox, webkit } from 'playwright-core';
import ts from 'typescript';

const fixture = process.env.SHEET_FIXTURE ?? 'China PAAK EPC evaluation.xlsx';
const ratio = Number(process.env.SHEET_PIXEL_RATIO ?? 1);
const server = spawn(process.execPath, ['scripts/serve.mjs', '--port', '0']);
const url = await new Promise((resolve, reject) => {
  createInterface({ input: server.stdout }).on('line', line => {
    const match = line.match(/http:\/\/\S+/);
    if (match) resolve(match[0]);
  });
  server.once('exit', code => reject(new Error(`Server exited: ${code}`)));
});
try {
  for (const name of (process.env.DOCVIEWKIT_BROWSERS ?? 'chromium,firefox,webkit').split(',')) {
    const browser = await ({ chromium, firefox, webkit })[name].launch({ timeout: 15000 });
    try {
      const page = await browser.newPage({ viewport: { width: 1000, height: 450 }, deviceScaleFactor: ratio });
      const errors = [];
      page.on('pageerror', error => errors.push(error.message));
      page.on('console', message => { if (message.type() === 'error') errors.push(message.text()); });
      if (process.env.SCROLL_BASELINE === '1' || process.env.SCROLL_SOURCE) {
        const source = process.env.SCROLL_SOURCE ? await readFile(process.env.SCROLL_SOURCE, 'utf8')
          : execFileSync('git', ['show', 'HEAD:src/viewer.ts'], { encoding: 'utf8' });
        const { outputText } = ts.transpileModule(source, { compilerOptions: { target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.ES2022 } });
        await page.route('**/dist/viewer.js', route => route.fulfill({ contentType: 'text/javascript', body: outputText }));
      }
      await page.addInitScript(() => {
        window.surfaceRequests = 0;
        window.allSurfaceRequests = 0;
        window.surfaceBitmaps = [];
        window.surfacePaints = 0;
        window.mainSurfacePaints = 0;
        const draw = CanvasRenderingContext2D.prototype.drawImage;
        CanvasRenderingContext2D.prototype.drawImage = function(...args) {
          if (this.canvas.classList.contains('surface')) window.mainSurfacePaints++;
          if (this.canvas.classList.contains('surface') || this.canvas.classList.contains('sheet-tile')) window.surfacePaints++;
          return draw.apply(this, args);
        };
        const workers = new WeakSet();
        const post = Worker.prototype.postMessage;
        Worker.prototype.postMessage = function(message, ...rest) {
          if (!workers.has(this)) {
            workers.add(this);
            this.addEventListener('message', event => {
              if (event.data.type === 'render' && event.data.ok) window.surfaceBitmaps.push(event.data.result.bitmap);
            });
          }
          if (message.type === 'render' && (message.supersedeKey === 'viewer-surface' || message.supersedeKey?.startsWith('viewer-sheet:'))) {
            window.allSurfaceRequests++;
            if (message.priority !== 'prefetch') window.surfaceRequests++;
          }
          if (window.slowSurface && message.type === 'render' && (message.supersedeKey === 'viewer-surface' || message.supersedeKey?.startsWith('viewer-sheet:'))) {
            setTimeout(() => post.call(this, message, ...rest), 120);
            return;
          }
          return post.call(this, message, ...rest);
        };
      });
      await page.goto(`${url}examples/viewer.html?fixture=${encodeURIComponent(fixture)}`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === 'true');
      const viewer = page.locator('docviewkit-viewer');
      const initial = await viewer.evaluate(v => v.shadowRoot.querySelector('.surface').toDataURL());
      const comparePixels = () => viewer.evaluate(async (v, fixture) => {
        const { createOfficeEngine } = await import('/dist/engine.js');
        const engine = await createOfficeEngine();
        const doc = await engine.open(new Uint8Array(await (await fetch(`/tests/fixtures/${encodeURIComponent(fixture)}`)).arrayBuffer()));
        try {
          const canvas = v.shadowRoot.querySelector('.surface');
          const stage = v.shadowRoot.querySelector('.stage');
          const request = { unitIndex: 0, scale: v.state.zoom, pixelRatio: devicePixelRatio,
            viewport: { x: (parseFloat(stage.style.left) - 48) / v.state.zoom, y: (parseFloat(stage.style.top) - 24) / v.state.zoom, width: parseFloat(canvas.style.width) / v.state.zoom, height: parseFloat(canvas.style.height) / v.state.zoom },
            background: '#ffffff', includeTextFragments: true, sheetSizes: { rows: [], columns: [] } };
          const direct = await doc.render(request);
          const reference = document.createElement('canvas');
          reference.width = direct.pixelWidth; reference.height = direct.pixelHeight;
          const context = reference.getContext('2d'); context.drawImage(direct.bitmap, 0, 0);
          const expected = context.getImageData(0, 0, reference.width, reference.height).data;
          const actual = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data;
          let changed = 0, maxDelta = 0, edgeChanges = 0;
          for (let i = 0; i < actual.length; i++) {
            const delta = Math.abs(actual[i] - expected[i]);
            const x = Math.floor(i / 4) % canvas.width, y = Math.floor(i / 4 / canvas.width);
            // A viewport can clip the endpoint of a grid stroke differently from
            // cropping a larger render. Compare every internal seam; report the outermost pixel separately.
            if (x === 0 || y === 0 || x === canvas.width - 1 || y === canvas.height - 1) {
              if (delta) edgeChanges++;
            } else {
              if (delta) changed++;
              maxDelta = Math.max(maxDelta, delta);
            }
          }
          direct.bitmap.close();
          return { changed, maxDelta, channels: actual.length, edgeChanges };
        } finally { doc.close(); engine.close(); }
      }, fixture);
      const parity = await comparePixels();
      assert.ok(parity.maxDelta <= 1, `${name}: tile seams changed actual document pixels: ${JSON.stringify(parity)}`);

      await page.waitForTimeout(250);
      await viewer.evaluate(v => v.shadowRoot.querySelector('.workspace').scrollTop = 550);
      await page.waitForFunction(() => parseFloat(document.querySelector('docviewkit-viewer').shadowRoot.querySelector('.stage').style.top) > 200);
      const requests = await page.evaluate(() => window.surfaceRequests);
      await viewer.evaluate(async v => {
        v.shadowRoot.querySelector('.workspace').scrollTop = 0;
        await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      });
      const returned = await viewer.evaluate(v => ({ image: v.shadowRoot.querySelector('.surface').toDataURL(), requests: window.surfaceRequests }));
      assert.equal(returned.requests, requests, `${name}: revisiting a rendered viewport started another worker render`);
      assert.equal(returned.image, initial, `${name}: cached viewport was not restored within two animation frames`);
      await viewer.evaluate(async v => {
        v.shadowRoot.querySelector('.workspace').scrollTop = 5;
        await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      });
      assert.equal(await page.evaluate(() => window.surfaceRequests), requests, `${name}: nearby position missed cached coverage`);
      const initialSheetHeight = await viewer.evaluate(v => parseFloat(v.shadowRoot.querySelector('.stage-wrap').style.height));
      const zoom = viewer.locator('.zoom-input');
      await zoom.fill('125');
      await zoom.press('Enter');
      await page.waitForFunction(previous => window.surfaceRequests > previous, requests);
      await page.waitForFunction(height => parseFloat(document.querySelector('docviewkit-viewer').shadowRoot.querySelector('.stage-wrap').style.height) > height * 1.2, initialSheetHeight);
      const moving = await viewer.evaluate(async v => {
        const workspace = v.shadowRoot.querySelector('.workspace');
        window.slowSurface = true;
        const beforeRequests = window.allSurfaceRequests;
        const beforePaints = window.surfacePaints;
        const uncovered = [], intervals = [];
        let lastTime = performance.now();
        for (let index = 1; index <= 80; index++) {
          // Continuous motion followed by an immediate reversal, while a worker
          // result takes 120ms. Inspect every displayed frame, not just settling.
          workspace.scrollTop = (index <= 40 ? index : 80 - index) * 30;
          workspace.scrollLeft = (index <= 40 ? index : 80 - index) * 45;
          const time = await new Promise(resolve => requestAnimationFrame(resolve));
          intervals.push(time - lastTime); lastTime = time;
          const area = workspace.getBoundingClientRect();
          const surfaces = [...v.shadowRoot.querySelectorAll('.surface, .sheet-tile')].map(c => c.getBoundingClientRect());
          let missing = 0;
          for (let y = area.top + 30; y < area.bottom - 20; y += 20) {
            for (let x = area.left + 55; x < area.right - 20; x += 40) {
              if (!surfaces.some(r => x >= r.left && x < r.right && y >= r.top && y < r.bottom)) missing++;
            }
          }
          if (missing) uncovered.push({ index, missing, top: workspace.scrollTop });
        }
        window.slowSurface = false;
        return { requests: window.allSurfaceRequests - beforeRequests, paints: window.surfacePaints - beforePaints, uncovered, p95: intervals.sort((a, b) => a - b)[75] };
      });
      assert.ok(moving.paints > 0, `${name}: no frames painted during continuous scrolling`);
      assert.equal(moving.uncovered.length, 0, `${name}: content disappeared during dragging: ${JSON.stringify(moving)}`);
      console.log(`${name} ${fixture} DPR ${ratio}: dragging ${JSON.stringify(moving)}`);
      assert.ok(moving.p95 < 50, `${name}: main-thread drag frames stalled: ${moving.p95}ms`);
      assert.ok(moving.requests <= 40, `${name}: continuous scrolling flooded the slow renderer: ${JSON.stringify(moving)}`);
      await page.waitForTimeout(300);
      await viewer.evaluate(v => {
        const workspace = v.shadowRoot.querySelector('.workspace');
        workspace.scrollLeft = 3300; workspace.scrollTop = 700;
      });
      await page.waitForFunction(() => {
        const root = document.querySelector('docviewkit-viewer').shadowRoot;
        const workspace = root.querySelector('.workspace').getBoundingClientRect();
        const surface = root.querySelector('.surface').getBoundingClientRect();
        return surface.left <= workspace.left + 48 && surface.top <= workspace.top + 24
          && surface.right >= workspace.right - 16 && surface.bottom >= workspace.bottom - 16;
      });
      const scrolledParity = await comparePixels();
      // Fractional device transforms can differ by two 8-bit antialiasing levels
      // between independently rasterized viewports; integer-scale paths stay within one.
      assert.ok(scrolledParity.maxDelta <= (Number.isInteger(ratio) ? 1 : 2), `${name}: independently rendered tiles changed pixels: ${JSON.stringify(scrolledParity)}`);

      await mkdir('.cache/sheet-scroll', { recursive: true });
      await page.screenshot({ path: `.cache/sheet-scroll/${name}-${ratio}-${fixture}.png` });
      for (const percent of [140, 150, 160, 170, 180, 190, 200, 210, 220]) {
        await zoom.fill(String(percent));
        await zoom.press('Enter');
        await page.waitForFunction(value => document.querySelector('docviewkit-viewer').state.zoom === value / 100, percent);
      }
      await page.waitForFunction(height => parseFloat(document.querySelector('docviewkit-viewer').shadowRoot.querySelector('.stage-wrap').style.height) > height * 2.1, initialSheetHeight);
      await viewer.evaluate(v => { const w = v.shadowRoot.querySelector('.workspace'); w.scrollLeft = w.scrollTop = 1e9; });
      await page.waitForFunction(() => {
        const root = document.querySelector('docviewkit-viewer').shadowRoot;
        const w = root.querySelector('.workspace').getBoundingClientRect();
        const c = root.querySelector('.surface').getBoundingClientRect();
        return c.left <= w.left + 48 && c.top <= w.top + 24 && c.right >= w.right - 16 && c.bottom >= w.bottom - 16;
      });
      await page.waitForTimeout(100);
      const idlePaints = await page.evaluate(() => window.mainSurfacePaints);
      await page.waitForTimeout(150);
      assert.equal(await page.evaluate(() => window.mainSurfacePaints), idlePaints, `${name}: sheet edge repainted while idle`);
      const retained = await page.evaluate(() => ({
        count: window.surfaceBitmaps.filter(bitmap => bitmap.width !== 0).length,
        pixels: [...document.querySelector('docviewkit-viewer').shadowRoot.querySelectorAll('.sheet-tile')].reduce((sum, canvas) => sum + canvas.width * canvas.height, 0),
      }));
      assert.ok(retained.count <= 8 && retained.pixels <= 32 * 1024 * 1024, `${name}: sheet cache exceeded its budget: ${JSON.stringify(retained)}`);
      const tabs = viewer.locator('.sheet-tab');
      if (await tabs.count() > 1) {
        await tabs.nth(1).click();
        await page.waitForFunction(() => document.querySelector('docviewkit-viewer').state.currentUnitIndex === 1);
        await tabs.nth(0).click();
        await page.waitForFunction(() => document.querySelector('docviewkit-viewer').state.currentUnitIndex === 0);
      }
      await viewer.evaluate(v => v.close());
      assert.equal(await viewer.evaluate(v => v.shadowRoot.querySelectorAll('.sheet-tile').length), 0);
      assert.equal(await page.evaluate(() => window.surfaceBitmaps.filter(bitmap => bitmap.width !== 0).length), 0, `${name}: document close leaked cached bitmaps`);
      assert.deepEqual(errors, []);
      console.log(`${name}: return scroll restored identical pixels within two frames, zero new render requests, clean console`);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
