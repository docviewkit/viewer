import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { chromium, webkit } from 'playwright-core';

const output = process.env.PAGES_BROCHURE_OUTPUT ?? 'output/playwright/pages-brochure';
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
      const page = await browser.newPage({ viewport: { width: 1500, height: 1050 }, reducedMotion: 'reduce' });
      const errors = [];
      page.on('pageerror', error => errors.push(error.message));
      page.on('console', message => { if (message.type() === 'error') errors.push(message.text()); });
      if (process.env.PAGES_BROCHURE_BASELINE_WASM) {
        const bytes = await readFile(process.env.PAGES_BROCHURE_BASELINE_WASM);
        await page.route('**/dist/office-viewer-pdf.wasm', route => route.fulfill({ contentType: 'application/wasm', body: bytes }));
      }
      await page.goto(`${url}examples/viewer.html?filePicker=true`);
      await page.waitForFunction(() => document.querySelector('#viewer').state.status === 'ready');
      await page.locator('#viewer').evaluate(async viewer => {
        await viewer.open(await (await fetch('/tests/fixtures/pages-brochure.pages')).arrayBuffer());
      });
      const result = await page.evaluate(async () => {
        const { createOfficeEngine } = await import('/dist/engine.js');
        const { extendedFormatPack } = await import('/dist/extended-formats.js');
        const engine = await createOfficeEngine({ formatPack: async () => extendedFormatPack });
        const doc = await engine.open(await (await fetch('/tests/fixtures/pages-brochure.pages')).arrayBuffer());
        try {
          const objects = await doc.listObjects();
          const frame = await doc.render({ unitIndex: 0, scale: 2, pixelRatio: 1, includeTextFragments: true });
          try {
            const canvas = document.createElement('canvas'); canvas.width = frame.pixelWidth; canvas.height = frame.pixelHeight;
            const ctx = canvas.getContext('2d'); ctx.drawImage(frame.bitmap, 0, 0);
            const headings = frame.textFragments.filter(f => f.objectId.endsWith('-33816') && f.text.includes('Subheadin'));
            const headingInk = headings.map(f => {
              const x = Math.floor(f.transform.e*2), y = Math.floor((f.transform.f - f.height)*2);
              const pixels = ctx.getImageData(x,y,Math.ceil(f.width*2),Math.ceil(f.height*4)).data;
              let ink = 0; for(let i=0;i<pixels.length;i+=4) if(pixels[i]<100 && pixels[i+1]<100 && pixels[i+2]<100) ink++;
              return ink;
            });
            return { image: canvas.toDataURL(), controlObjects: objects.filter(o => o.text?.includes('\x04')).map(o=>o.id),
              controlFragments: frame.textFragments.filter(f=>f.text.includes('\x04')).map(f=>f.text),
              headings: headings.map(f=>({text:f.text,font:f.font})), headingInk,
              body: frame.textFragments.filter(f => f.objectId.endsWith('-33816') && ((f.start >= 25 && f.end <= 220) || f.start >= 240)).map(f=>({text:f.text,font:f.font})) };
          } finally { frame.bitmap.close(); }
        } finally { doc.close(); engine.close(); }
      });
      await writeFile(`${output}/${name}-page.png`, Buffer.from(result.image.split(',')[1], 'base64'));
      delete result.image;
      await writeFile(`${output}/${name}-result.json`, JSON.stringify(result, null, 2));
      await page.screenshot({ path: `${output}/${name}-viewer.png` });
      const failures = [];
      if (result.controlObjects.length || result.controlFragments.length) failures.push('invisible story terminator exposed as text');
      if (result.headings.length !== 2 || result.headings.some(f=>!/^bold\b|\b700\b/u.test(f.font))) failures.push('two headings must use bold fonts');
      if (result.headingInk.some(count=>count<150)) failures.push('heading ink missing');
      if (result.body.length<2 || result.body.some(f=>/^bold\b|\b700\b/u.test(f.font))) failures.push('body text must stay regular');
      assert.deepEqual(errors, [], `${name}: console errors`);
      assert.deepEqual(failures, [], `${name}: ${JSON.stringify(result)}`);
      console.log(`${name}: hidden story marker omitted; two bold headings and regular body verified in rendered fonts and pixels`);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
