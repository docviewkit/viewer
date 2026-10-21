import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { chromium, webkit } from 'playwright-core';

const output = process.env.NUMBERS_RETIREMENT_OUTPUT ?? 'output/playwright/numbers-retirement';
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
      const page = await browser.newPage({ viewport: { width: 1200, height: 1100 }, reducedMotion: 'reduce' });
      const errors = [];
      page.on('pageerror', error => errors.push(error.message));
      page.on('console', message => { if (message.type() === 'error') errors.push(message.text()); });
      if (process.env.NUMBERS_RETIREMENT_BASELINE_WASM) {
        const bytes = await readFile(process.env.NUMBERS_RETIREMENT_BASELINE_WASM);
        await page.route('**/dist/office-viewer-pdf.wasm', route => route.fulfill({ contentType: 'application/wasm', body: bytes }));
      }
      await page.goto(`${url}examples/viewer.html?filePicker=true`);
      await page.waitForFunction(() => document.querySelector('#viewer').state.status === 'ready');
      await page.locator('#viewer').evaluate(async viewer => {
        await viewer.open(await (await fetch('/tests/fixtures/numbers-retirement.numbers')).arrayBuffer());
      });
      const result = await page.evaluate(async () => {
        const { createOfficeEngine } = await import('/dist/engine.js');
        const { extendedFormatPack } = await import('/dist/extended-formats.js');
        const engine = await createOfficeEngine({ execution: 'inline', formatPack: async () => extendedFormatPack });
        const doc = await engine.open(await (await fetch('/tests/fixtures/numbers-retirement.numbers')).arrayBuffer());
        try {
          const viewport = { x: 0, y: 500, width: 500, height: 260 };
          const frame = await doc.render({ unitIndex: 0, viewport, scale: 2, pixelRatio: 1, includeTextFragments: true });
          try {
            const canvas = document.createElement('canvas'); canvas.width = frame.pixelWidth; canvas.height = frame.pixelHeight;
            canvas.getContext('2d').drawImage(frame.bitmap, 0, 0);
            const fragments = frame.textFragments.map(f => f.text);
            const objects = await doc.listObjects();
            const table = objects.find(o => o.type === 'table' && o.id.endsWith('-1139670'));
            const cell = objects.find(o => o.id === 'numbers:0:1139670:numbers:0:C1');
            const pixel = (x, y) => Array.from(canvas.getContext('2d').getImageData(Math.floor((x - viewport.x) * 2), Math.floor((y - viewport.y) * 2), 1, 1).data);
            const blank = pixel(450, 700);
            const cellFill = pixel(cell.bounds.x + cell.bounds.width - 4, cell.bounds.y + cell.bounds.height - 4);
            const outside = await doc.render({ unitIndex: 0, viewport: { x: 1500, y: 1500, width: 32, height: 32 }, scale: 1, pixelRatio: 1 });
            let outsideFill;
            try {
              const c = document.createElement('canvas'); c.width = c.height = 32;
              c.getContext('2d').drawImage(outside.bitmap, 0, 0);
              outsideFill = Array.from(c.getContext('2d').getImageData(15, 15, 1, 1).data);
            } finally { outside.bitmap.close(); }
            return { image: canvas.toDataURL(), fragments, table: table.bounds, blank, cellFill, outsideFill,
              hidden: objects.filter(o => o.unitIndex === 0 && /Present value of needed funds|First year savings contribution|Real dollar contribution at year 0/.test(o.text ?? '')).map(o => o.text) };
          } finally { frame.bitmap.close(); }
        } finally { doc.close(); engine.close(); }
      });
      await writeFile(`${output}/${name}-goals.png`, Buffer.from(result.image.split(',')[1], 'base64'));
      delete result.image;
      await writeFile(`${output}/${name}-result.json`, JSON.stringify(result, null, 2));
      await page.screenshot({ path: `${output}/${name}-viewer.png` });
      const failures = [];
      if (result.hidden.length) failures.push('hidden helper rows rendered');
      if (Math.abs(result.table.height - 2 * 21.659973 * 4 / 3) > .01) failures.push('hidden rows consume height');
      for (const key of ['blank', 'cellFill', 'outsideFill']) if (JSON.stringify(result[key]) !== JSON.stringify([255, 251, 239, 255])) failures.push(`incorrect ${key}`);
      if (!result.fragments.includes('¥117') || !result.fragments.includes('¥2,389,605')) failures.push('visible currency values missing');
      assert.deepEqual(errors, [], `${name}: console errors`);
      assert.deepEqual(failures, [], `${name}: ${JSON.stringify(result)}`);
      console.log(`${name}: hidden helper rows, compact geometry, sheet and transparent-cell background pixels passed`);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
