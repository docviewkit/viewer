import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { chromium, webkit } from 'playwright-core';

const output = process.env.NUMBERS_CORRELATION_OUTPUT ?? 'output/playwright/numbers-correlation';
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
      if (process.env.NUMBERS_CORRELATION_BASELINE_WASM) {
        const bytes = await readFile(process.env.NUMBERS_CORRELATION_BASELINE_WASM);
        await page.route('**/dist/office-viewer-pdf.wasm', route => route.fulfill({ contentType: 'application/wasm', body: bytes }));
      }
      await page.goto(`${url}examples/viewer.html?filePicker=true`);
      await page.waitForFunction(() => document.querySelector('#viewer').state.status === 'ready');
      await page.locator('#viewer').evaluate(async viewer => {
        await viewer.open(await (await fetch('/tests/fixtures/numbers-correlation.numbers')).arrayBuffer());
      });
      const result = await page.evaluate(async () => {
        const { createOfficeEngine } = await import('/dist/engine.js');
        const { extendedFormatPack } = await import('/dist/extended-formats.js');
        const engine = await createOfficeEngine({ execution: 'inline', formatPack: async () => extendedFormatPack });
        const doc = await engine.open(await (await fetch('/tests/fixtures/numbers-correlation.numbers')).arrayBuffer());
        try {
          const viewport = { x: 300, y: 550, width: 440, height: 450 };
          const frame = await doc.render({ unitIndex: 0, viewport, scale: 2, pixelRatio: 1, includeTextFragments: true });
          try {
            const canvas = document.createElement('canvas'); canvas.width = frame.pixelWidth; canvas.height = frame.pixelHeight;
            canvas.getContext('2d').drawImage(frame.bitmap, 0, 0);
            const objects = await doc.listObjects();
            const markers = objects.filter(o => o.id.includes('scatter-marker'));
            const trendlines = objects.filter(o => o.id.includes('scatter-trendline'));
            const plot = { x: 280 * 4 / 3, y: 447 * 4 / 3, width: 246.628067 * 4 / 3, height: 238 * 4 / 3 };
            const pixel = (x, y) => Array.from(canvas.getContext('2d').getImageData(Math.floor((x - viewport.x) * 2), Math.floor((y - viewport.y) * 2), 1, 1).data);
            const map = (x,y) => [plot.x + x / 70 * plot.width, plot.y + (8-y) / 8 * plot.height];
            const points = [[55.5,6],[61.2,6.2],[69,7.4],[66.5,7]];
            const markerPixels = points.map(p => pixel(...map(...p)));
            const meanX = points.reduce((s,p) => s+p[0],0)/4, meanY = points.reduce((s,p) => s+p[1],0)/4;
            const slope = points.reduce((s,p) => s+(p[0]-meanX)*(p[1]-meanY),0) / points.reduce((s,p) => s+(p[0]-meanX)**2,0);
            const trendPixel = pixel(...map(63, meanY + slope*(63-meanX)));
            const groups = Object.values(Object.groupBy(frame.textFragments, f => f.objectId));
            const text = fs => fs.map(f => f.text).join(' ').replace(/\s+/gu, ' ').trim();
            const texts = groups.map(text);
            const xBounds = fs => fs.flatMap(f => [[0,0],[f.width,0],[0,f.height],[f.width,f.height]].map(([x,y]) => f.transform.a*x+f.transform.c*y+f.transform.e));
            const title = groups.find(fs => text(fs) === 'Hand Size (in)');
            const ticks = groups.filter(fs => ['1.6','3.2','4.8','6.4'].includes(text(fs)));
            const labelGap = title && ticks.length ? Math.min(...ticks.flatMap(xBounds)) - Math.max(...xBounds(title)) : -1;
            const layoutFrame = await doc.render({ unitIndex: 0, viewport: {x:0,y:0,width:740,height:550}, scale:1, pixelRatio:1, includeTextFragments:true });
            let dataTitleX = -1, separator = false, instructionFragment;
            try {
              const c = document.createElement('canvas'); c.width = 740; c.height = 550;
              const ctx = c.getContext('2d'); ctx.drawImage(layoutFrame.bitmap,0,0);
              dataTitleX = layoutFrame.textFragments.find(f => f.text === 'Data')?.transform.e ?? -1;
              const instructions = layoutFrame.textFragments.find(f => f.text.startsWith('Instructions'));
              instructionFragment = instructions;
              if (instructions) for(let y = Math.floor(instructions.transform.f)-30; y < instructions.transform.f-2; y++) {
                const pixels = ctx.getImageData(340,y,350,1).data;
                let dark = 0; for(let x=0;x<pixels.length;x+=4) if(pixels[x]<200 && pixels[x+1]<200 && pixels[x+2]<200) dark++;
                if(dark>330) separator = true;
              }
            } finally { layoutFrame.bitmap.close(); }
            return { image: canvas.toDataURL(), markers: markers.map(o => o.bounds), trendlines: trendlines.length,
              dataTitleX, separator, instructionFragment, labelGap, connections: objects.filter(o => o.id.includes('scatter-line')).length, markerPixels, trendPixel, texts,
              diagnostics: doc.info.diagnostics ?? [], plot, points };
          } finally { frame.bitmap.close(); }
        } finally { doc.close(); engine.close(); }
      });
      await writeFile(`${output}/${name}-chart.png`, Buffer.from(result.image.split(',')[1], 'base64'));
      delete result.image;
      await writeFile(`${output}/${name}-result.json`, JSON.stringify(result, null, 2));
      await page.screenshot({ path: `${output}/${name}-viewer.png` });
      const failures = [];
      if (result.markers.length !== 4) failures.push('missing four XY markers or spurious blank points');
      result.markers.forEach((b,i) => {
        const [x,y] = result.points[i];
        if (Math.abs(b.x+b.width/2 - result.plot.x - x/70*result.plot.width) > .01
          || Math.abs(b.y+b.height/2 - result.plot.y - (8-y)/8*result.plot.height) > .01) failures.push(`wrong XY marker ${i}`);
      });
      const blue = p => p[2] > p[0]+70 && p[1] > p[0]+30;
      if (!result.markerPixels.every(blue)) failures.push('missing plus symbol pixels');
      if (result.trendlines !== 1 || !blue(result.trendPixel)) failures.push('missing linear regression pixels');
      if (result.dataTitleX < 0 || result.dataTitleX > 5) failures.push('table title is not left aligned');
      if (!result.separator) failures.push('missing instruction separator pixels');
      if (result.labelGap < 1) failures.push('vertical title overlaps numeric ticks');
      if (result.connections !== 0) failures.push('unexpected point connection lines');
      for (const text of ['Results','Height (in)','Hand Size (in)','0','14','28','42','56','70','1.6','3.2','4.8','6.4','8']) {
        if (!result.texts.includes(text)) failures.push(`missing axis/title ${text}`);
      }
      assert.deepEqual(errors, [], `${name}: console errors`);
      assert.deepEqual(failures, [], `${name}: ${JSON.stringify(result)}`);
      console.log(`${name}: real scatter XY positions, plus pixels, linear trendline, numeric axes and titles passed`);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
