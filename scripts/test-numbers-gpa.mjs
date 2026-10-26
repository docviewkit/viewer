import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { chromium, webkit } from 'playwright-core';

const output = process.env.NUMBERS_GPA_OUTPUT ?? 'output/playwright/numbers-gpa';
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
      const page = await browser.newPage({ viewport: { width: 1200, height: 900 }, reducedMotion: 'reduce' });
      const errors = [];
      page.on('pageerror', error => errors.push(error.message));
      page.on('console', message => { if (message.type() === 'error') errors.push(message.text()); });
      if (process.env.NUMBERS_GPA_BASELINE_WASM) {
        const bytes = await readFile(process.env.NUMBERS_GPA_BASELINE_WASM);
        await page.route('**/dist/office-viewer-pdf.wasm', route => route.fulfill({ contentType: 'application/wasm', body: bytes }));
      }
      await page.goto(`${url}examples/viewer.html?filePicker=true`);
      await page.waitForFunction(() => document.querySelector('#viewer').state.status === 'ready');
      await page.locator('#viewer').evaluate(async viewer => {
        await viewer.open(await (await fetch('/tests/fixtures/numbers-gpa.numbers')).arrayBuffer());
      });
      const result = await page.evaluate(async () => {
        const { createOfficeEngine } = await import('/dist/engine.js');
        const { extendedFormatPack } = await import('/dist/extended-formats.js');
        const engine = await createOfficeEngine({ execution: 'inline', formatPack: async () => extendedFormatPack });
        const doc = await engine.open(await (await fetch('/tests/fixtures/numbers-gpa.numbers')).arrayBuffer());
        try {
          const viewport = { x: 380, y: 120, width: 370, height: 220 };
          const frame = await doc.render({ unitIndex: 0, viewport, scale: 3, pixelRatio: 1, includeTextFragments: true });
          try {
            const canvas = document.createElement('canvas'); canvas.width = frame.pixelWidth; canvas.height = frame.pixelHeight;
            canvas.getContext('2d').drawImage(frame.bitmap, 0, 0);
            const markers = (await doc.listObjects()).filter(o => o.id.includes(':line-marker-'));
            const pixel = (x, y) => Array.from(canvas.getContext('2d').getImageData(Math.floor((x - viewport.x) * 3), Math.floor((y - viewport.y) * 3), 1, 1).data);
            const groups = Object.values(Object.groupBy(frame.textFragments, f => f.objectId));
            return { image: canvas.toDataURL(), markers: markers.map(o => o.bounds),
              centers: markers.map(o => pixel(o.bounds.x + o.bounds.width / 2, o.bounds.y + o.bounds.height / 2)),
              rings: markers.map(o => pixel(o.bounds.x + o.bounds.width / 2, o.bounds.y + 0.5)),
              labels: groups.map(fs => fs.map(f => f.text).join(' ').replace(/\s+/gu, ' ').trim()) };
          } finally { frame.bitmap.close(); }
        } finally { doc.close(); engine.close(); }
      });
      await writeFile(`${output}/${name}-chart.png`, Buffer.from(result.image.split(',')[1], 'base64'));
      delete result.image;
      await writeFile(`${output}/${name}-result.json`, JSON.stringify(result, null, 2));
      await page.screenshot({ path: `${output}/${name}-viewer.png` });
      const failures = [];
      if (result.markers.length !== 4) failures.push('missing four line circles');
      result.markers.forEach((b, i) => {
        const expectedX = (317 + i * 206 / 3) * 4 / 3;
        const expectedY = (121 + (4 - (i === 0 ? 3.67857142857143 : 0)) / 4 * 102) * 4 / 3;
        if (Math.abs(b.x + b.width / 2 - expectedX) > .01 || Math.abs(b.y + b.height / 2 - expectedY) > .01) failures.push(`wrong circle position ${i}`);
      });
      if (!result.centers.every(p => p.slice(0, 3).every(v => v > 240))) failures.push('missing white circle interiors');
      if (!result.rings.every(p => p[2] > p[0] + 80 && p[2] > p[1] + 80)) failures.push('missing blue circle outlines');
      for (const text of ['GPA by Term', 'Term 1', 'Term 2', 'Term 3', 'Term 4', '0', '1', '2', '3', '4']) {
        if (!result.labels.includes(text)) failures.push(`missing label ${text}`);
      }
      assert.deepEqual(errors, [], `${name}: console errors`);
      assert.deepEqual(failures, [], `${name}: ${JSON.stringify(result)}`);
      console.log(`${name}: original GPA line circles, white interiors, blue outlines and both axes passed`);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
