import assert from 'node:assert/strict';
import { writeFileSync } from 'node:fs';
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { chromium, firefox, webkit } from 'playwright-core';

const server = spawn(process.execPath, ['scripts/serve.mjs', '--port', '0'], { stdio: ['ignore', 'pipe', 'inherit'] });
try {
  const url = await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('Server startup timed out')), 10000);
    createInterface({ input: server.stdout }).on('line', line => {
      const match = line.match(/https?:\/\/\S+/u);
      if (match) { clearTimeout(timer); resolve(match[0]); }
    });
  });
  for (const name of (process.env.DOCVIEWKIT_BROWSERS ?? 'chromium,firefox,webkit').split(',')) {
    const browser = await ({ chromium, firefox, webkit })[name].launch({ headless: true, timeout: 30000 });
    try {
      const page = await browser.newPage({ viewport: { width: 1200, height: 1000 } });
      if (process.env.OFD_CLIP_WASM) {
        await page.route('**/office-viewer-ofd.wasm', route => route.fulfill({ path: process.env.OFD_CLIP_WASM, contentType: 'application/wasm' }));
      }
      const errors = [];
      page.on('pageerror', e => errors.push(e.message));
      page.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
      await page.goto(`${url}examples/viewer.html?fixture=ofdrw-y.ofd`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === 'true');
      const pages = page.locator('#viewer').locator('.continuous-page');
      await pages.nth(2).scrollIntoViewIfNeeded();
      await page.waitForFunction(() => document.querySelector('#viewer').shadowRoot.querySelectorAll('.continuous-page')[2]?.dataset.rendered === 'true');
      const pixels = await pages.nth(2).evaluate(element => {
        const canvas = element.querySelector('canvas');
        const { width: w, height: h } = canvas;
        // Original page 3: diagram at x=9.6..132 mm, y=29.3..120 mm on a 204x264 mm page.
        const data = canvas.getContext('2d').getImageData(w * .045, h * .11, w * .61, h * .35).data;
        let cyan = 0;
        for (let i = 0; i < data.length; i += 4) {
          if (data[i] < 100 && data[i + 1] > 130 && data[i + 2] > 140 && data[i + 3] > 200) cyan++;
        }
        return cyan;
      });
      const png = await pages.nth(2).locator('canvas').evaluate(canvas => canvas.toDataURL('image/png').split(',')[1]);
      writeFileSync(`/tmp/ofd-y-${name}-${process.env.OFD_CLIP_STAGE ?? 'after'}.png`, Buffer.from(png, 'base64'));
      assert.ok(pixels > 1000, `${name}: missing cyan network diagram (${pixels} pixels)`);
      assert.deepEqual(errors, []);
      console.log(`${name}: diagram visible (${pixels} cyan pixels), no console errors`);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
