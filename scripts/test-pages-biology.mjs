import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { chromium, webkit, firefox } from 'playwright-core';

const output = process.env.PAGES_BIOLOGY_OUTPUT ?? 'output/playwright/pages-biology';
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
  const failures = [];
  const browsers = process.env.PAGES_BIOLOGY_FIREFOX === '1' ? { firefox } : { chromium, webkit };
  for (const [name, type] of Object.entries(browsers)) {
    const browser = await type.launch({ headless: true });
    try {
      const page = await browser.newPage({ viewport: { width: 1500, height: 1050 }, reducedMotion: 'reduce' });
      const errors = [];
      page.on('pageerror', error => errors.push(error.message));
      page.on('console', message => { if (message.type() === 'error') errors.push(message.text()); });
      if (process.env.PAGES_BIOLOGY_BASELINE_WASM) {
        const bytes = await readFile(process.env.PAGES_BIOLOGY_BASELINE_WASM);
        await page.route('**/dist/office-viewer-pdf.wasm', route => route.fulfill({ contentType: 'application/wasm', body: bytes }));
      }
      await page.goto(`${url}examples/viewer.html?filePicker=true`);
      await page.waitForFunction(() => document.querySelector('#viewer').state.status === 'ready');
      await page.locator('#viewer').evaluate(async viewer => {
        await viewer.open(await (await fetch('/tests/fixtures/pages-biology.pages')).arrayBuffer());
      });
      const result = await page.evaluate(async () => {
        const { createOfficeEngine } = await import('/dist/engine.js');
        const { extendedFormatPack } = await import('/dist/extended-formats.js');
        const engine = await createOfficeEngine({ formatPack: async () => extendedFormatPack });
        const doc = await engine.open(await (await fetch('/tests/fixtures/pages-biology.pages')).arrayBuffer());
        try {
          const objects = await doc.listObjects();
          const frame = await doc.render({ unitIndex: 0, scale: 2, pixelRatio: 1, includeTextFragments: true });
          try {
            const canvas = document.createElement('canvas'); canvas.width = frame.pixelWidth; canvas.height = frame.pixelHeight;
            const ctx = canvas.getContext('2d'); ctx.drawImage(frame.bitmap, 0, 0);
            const titleObject = objects.find(o => o.text === 'THE STUDY OF CELL BIOLOGY');
            const title = frame.textFragments.filter(f => f.objectId === titleObject.id);
            const pixels = Array.from(ctx.getImageData(4,4,1,1).data);
            const otherPages = [];
            for (const unitIndex of [1,2]) {
              const other = await doc.render({ unitIndex, scale: 1, pixelRatio: 1 });
              try {
                const c = document.createElement('canvas'); c.width = other.pixelWidth; c.height = other.pixelHeight;
                const context = c.getContext('2d'); context.drawImage(other.bitmap,0,0);
                otherPages.push({ unitIndex, pixels: Array.from(context.getImageData(2,2,1,1).data), image: c.toDataURL() });
              } finally { other.bitmap.close(); }
            }
            return { image: canvas.toDataURL(), pixels, otherPages,
              title: title.map(f=>({text:f.text,font:f.font,y:f.transform.f})),
              fonts: objects.filter(o=>o.fontRuns?.length).map(o=>({text:o.text?.slice(0,80),runs:o.fontRuns})) };
          } finally { frame.bitmap.close(); }
        } finally { doc.close(); engine.close(); }
      });
      await writeFile(`${output}/${name}-page.png`, Buffer.from(result.image.split(',')[1], 'base64'));
      delete result.image;
      for (const other of result.otherPages) {
        await writeFile(`${output}/${name}-page-${other.unitIndex+1}.png`, Buffer.from(other.image.split(',')[1], 'base64'));
        delete other.image;
      }
      await writeFile(`${output}/${name}-result.json`, JSON.stringify(result, null, 2));
      await page.screenshot({ path: `${output}/${name}-viewer.png` });
      try {
        assert.deepEqual(errors, [], `${name}: console errors`);
        assert.deepEqual(result.pixels, [56,62,78,255], `${name}: section background`);
        assert.deepEqual(result.otherPages.map(p=>p.pixels), [[56,62,78,255],[255,255,255,255]], `${name}: section fill isolation`);
        const titleLines = [...new Set(result.title.map(f=>f.y))].map(y=>result.title.filter(f=>f.y===y).map(f=>f.text).join('').trim());
        assert.deepEqual(titleLines, ['THE','STUDY','OF CELL','BIOLOGY'], `${name}: title line breaks`);
        assert.ok(result.fonts.filter(o=>o.runs.some(r=>r.authoredFamily.startsWith('AvenirNext'))).every(o=>o.runs.every(r=>r.source === 'browser')), `${name}: Avenir faces retained`);
        console.log(`${name}: three page backgrounds, complete title lines and Avenir faces verified`);
      } catch (error) { failures.push(`${name}: ${error.message}; title=${JSON.stringify(result.title.map(f=>f.text))}`); }
    } finally { await browser.close(); }
  }
  assert.deepEqual(failures, []);
} finally { server.kill(); }
