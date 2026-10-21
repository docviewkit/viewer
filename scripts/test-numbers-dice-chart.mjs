import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { chromium, webkit } from 'playwright-core';

const output = process.env.NUMBERS_DICE_OUTPUT ?? 'output/playwright/numbers-dice';
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
      if (process.env.NUMBERS_DICE_BASELINE_WASM) {
        const bytes = await readFile(process.env.NUMBERS_DICE_BASELINE_WASM);
        await page.route('**/dist/office-viewer-pdf.wasm', route => route.fulfill({ contentType: 'application/wasm', body: bytes }));
      }
      await page.goto(`${url}examples/viewer.html?filePicker=true`);
      await page.waitForFunction(() => document.querySelector('#viewer').state.status === 'ready');
      await page.locator('#viewer').evaluate(async viewer => {
        await viewer.open(await (await fetch('/tests/fixtures/numbers-dice-lab.numbers')).arrayBuffer());
      });
      const result = await page.evaluate(async () => {
        const { createOfficeEngine } = await import('/dist/engine.js');
        const { extendedFormatPack } = await import('/dist/extended-formats.js');
        const engine = await createOfficeEngine({ execution: 'inline', formatPack: async () => extendedFormatPack });
        const doc = await engine.open(await (await fetch('/tests/fixtures/numbers-dice-lab.numbers')).arrayBuffer());
        try {
          const viewport = { x: 250, y: 530, width: 470, height: 350 };
          const frame = await doc.render({ unitIndex: 0, viewport, scale: 2, pixelRatio: 1, includeTextFragments: true });
          try {
            const canvas = document.createElement('canvas'); canvas.width = frame.pixelWidth; canvas.height = frame.pixelHeight;
            canvas.getContext('2d').drawImage(frame.bitmap, 0, 0);
            const plot = { x: 244 * 4 / 3, y: 421.6 * 4 / 3, width: 266 * 4 / 3, height: 202 * 4 / 3 };
            const fragments = frame.textFragments;
            const seriesLabels = fragments.filter(f => f.transform.f > plot.y + plot.height && f.transform.f < plot.y + plot.height + 30).map(f => f.text);
            const ticks = fragments.filter(f => /^-?\d+(?:\.\d+)?$/.test(f.text) && f.transform.e < plot.x && f.width < 30 && f.transform.f >= plot.y - 10 && f.transform.f <= plot.y + plot.height + 10).map(f => f.text);
            const categoryWidth = plot.width / 11;
            const barWidth = categoryWidth / 1.7;
            const x = plot.x + categoryWidth * 4 + barWidth * .35, y = plot.y;
            const pixel = (px, py) => Array.from(canvas.getContext('2d').getImageData(Math.floor((px - viewport.x) * 2), Math.floor((py - viewport.y) * 2), 1, 1).data);
            const corner = pixel(x + 1, y + 1), center = pixel(x + barWidth / 2, y + 2);
            return { image: canvas.toDataURL(), seriesLabels, ticks, corner, center,
              title: fragments.some(f => f.text === 'Total' && fragments.some(g => g.objectId === f.objectId && g.text === 'Occurrences')), fragments };
          } finally { frame.bitmap.close(); }
        } finally { doc.close(); engine.close(); }
      });
      await writeFile(`${output}/${name}-chart.png`, Buffer.from(result.image.split(',')[1], 'base64'));
      delete result.image;
      await writeFile(`${output}/${name}-result.json`, JSON.stringify(result, null, 2));
      await page.screenshot({ path: `${output}/${name}-viewer.png` });
      const failures = [];
      if (JSON.stringify(result.seriesLabels) !== JSON.stringify(Array.from({ length: 11 }, (_, i) => String(i + 2)))) failures.push('missing series labels 2..12');
      if (Math.min(...result.corner.slice(0, 3)) < 180 || result.center[2] < result.center[0] + 50) failures.push('missing rounded cap pixels');
      if (JSON.stringify(result.ticks) !== JSON.stringify(['0', '1', '2'])) failures.push('authored major tick count');
      if (!result.title) failures.push('missing vertical axis title');
      assert.deepEqual(errors, [], `${name}: console errors`);
      assert.deepEqual(failures, [], `${name}: ${JSON.stringify({ labels: result.seriesLabels, ticks: result.ticks, corner: result.corner, center: result.center })}`);
      console.log(`${name}: native rounded caps, series labels 2..12, ticks and axis title passed`);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
