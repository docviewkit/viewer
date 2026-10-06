import assert from 'node:assert/strict';
import { spawn, execFileSync } from 'node:child_process';
import { readFile, mkdir, writeFile } from 'node:fs/promises';
import { createInterface } from 'node:readline';
import { build } from 'esbuild';
import { chromium, webkit, firefox } from 'playwright-core';

const baseline = process.env.SHADOW_BASELINE === '1';
const worker = baseline ? (await build({
  entryPoints: ['src/worker.ts'], bundle: true, write: false, format: 'esm', target: 'es2022',
  plugins: [{ name: 'before-shadow-fix', setup(builder) {
    builder.onLoad({ filter: /\/src\/render\.ts$/ }, () => ({
      contents: execFileSync('git', ['show', 'HEAD:src/render.ts'], { encoding: 'utf8' }),
      loader: 'ts', resolveDir: `${process.cwd()}/src`,
    }));
  } }],
})).outputFiles[0].text : await readFile('dist/worker.js', 'utf8');
const guard = `
let allocatedPixels = 0;
const NativeCanvas = OffscreenCanvas;
for (const axis of ['width', 'height']) {
  const descriptor = Object.getOwnPropertyDescriptor(NativeCanvas.prototype, axis);
  Object.defineProperty(NativeCanvas.prototype, axis, { ...descriptor, set(value) {
    allocatedPixels += value * this[axis === 'width' ? 'height' : 'width'];
    if (allocatedPixels > 256 * 1024 * 1024) throw new Error('Shadow scratch canvases exceeded 256M cumulative pixels');
    descriptor.set.call(this, value);
  } });
}
globalThis.OffscreenCanvas = new Proxy(NativeCanvas, { construct(Type, args) {
  allocatedPixels += args[0] * args[1];
  if (allocatedPixels > 256 * 1024 * 1024) throw new Error('Shadow scratch canvases exceeded 256M cumulative pixels');
  return Reflect.construct(Type, args);
} });
const post = globalThis.postMessage.bind(globalThis);
globalThis.postMessage = (message, ...args) => {
  if (message.type === 'render' && message.ok) {
    console.log('shadow-allocation-pixels:' + allocatedPixels);
    allocatedPixels = 0;
  }
  return post(message, ...args);
};
`;
const server = spawn(process.execPath, ['scripts/serve.mjs', '--port', '0']);
const url = await new Promise((resolve, reject) => {
  createInterface({ input: server.stdout }).on('line', line => {
    const match = line.match(/http:\/\/\S+/);
    if (match) resolve(match[0]);
  });
  server.once('exit', code => reject(new Error(`Server exited: ${code}`)));
});
try {
  for (const name of (process.env.DOCVIEWKIT_BROWSERS ?? 'chromium,webkit').split(',')) {
    const browser = await ({ chromium, webkit, firefox })[name].launch({ timeout: 15000 });
    try {
      const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
      const errors = [], allocations = [];
      page.on('pageerror', error => errors.push(error.message));
      page.on('console', message => {
        if (message.type() === 'error') errors.push(message.text());
        if (message.text().startsWith('shadow-allocation-pixels:')) allocations.push(Number(message.text().split(':')[1]));
      });
      await page.route('**/dist/worker.js', route => route.fulfill({ contentType: 'text/javascript', body: guard + worker }));
      const start = Date.now();
      await page.goto(`${url}examples/viewer.html?fixture=OlapPivotA3.xlsx`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === 'true'
        || document.documentElement.dataset.error !== undefined, null, { timeout: 15000 });
      assert.deepEqual(errors, [], `${name}: original workbook failed to render`);
      const openedMs = Date.now() - start;
      assert.ok(allocations.length > 0 && allocations[0] < 256 * 1024 * 1024);
      const viewer = page.locator('docviewkit-viewer');
      await viewer.locator('[data-interaction-mode="text"]').click();
      const content = await viewer.evaluate(v => ({
        status: v.state.status,
        cells: v.shadowRoot.querySelectorAll('.text-layer-fragment').length,
        width: v.shadowRoot.querySelector('.surface').width,
      }));
      assert.equal(content.status, 'ready');
      assert.ok(content.cells > 50 && content.width > 1000, JSON.stringify(content));
      await viewer.evaluate(v => v.shadowRoot.querySelector('.workspace').scrollLeft = 5000);
      await page.waitForFunction(() => parseFloat(document.querySelector('docviewkit-viewer').shadowRoot.querySelector('.stage').style.left) > 4000);
      await viewer.evaluate(v => v.shadowRoot.querySelector('.workspace').scrollLeft = 0);
      await page.waitForFunction(() => parseFloat(document.querySelector('docviewkit-viewer').shadowRoot.querySelector('.stage').style.left) < 100);
      await mkdir('.cache/olap-pivot', { recursive: true });
      await page.screenshot({ path: `.cache/olap-pivot/${name}.png` });
      await viewer.evaluate(v => v.shadowRoot.querySelector('.workspace').scrollLeft = 1700);
      await page.waitForFunction(() => parseFloat(document.querySelector('docviewkit-viewer').shadowRoot.querySelector('.stage').style.left) > 1000);
      await page.screenshot({ path: `.cache/olap-pivot/${name}-icons-viewer.png` });
      assert.deepEqual(errors, []);
      await viewer.evaluate(v => v.close());
      const parity = await page.evaluate(async () => {
        const { Core, DEFAULT_LIMITS } = await import('/dist/core.js');
        const { SceneRenderer } = await import('/dist/render.js');
        const core = await Core.create(await (await fetch('/dist/office-viewer-core.wasm')).arrayBuffer(), DEFAULT_LIMITS);
        const doc = core.open(new Uint8Array(await (await fetch('/tests/fixtures/OlapPivotA3.xlsx')).arrayBuffer()));
        const scene = doc.loadUnit(0).scene;
        const cell = scene.objects.find(object => object.type === 'cell' && object.source.address === 'P4');
        const renderer = new SceneRenderer(scene.objects, DEFAULT_LIMITS);
        const frame = await renderer.render(scene.info.units[0], { unitIndex: 0, scale: 2,
          viewport: { x: cell.bounds.x, y: cell.bounds.y - 19, width: cell.bounds.width, height: 190 }, background: '#ffffff' });
        const canvas = document.createElement('canvas');
        canvas.width = frame.pixelWidth; canvas.height = frame.pixelHeight;
        const context = canvas.getContext('2d'); context.drawImage(frame.bitmap, 0, 0);
        const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data;
        for (const [row, color] of [[4, [22, 163, 74]], [5, [250, 204, 21]], [6, [250, 204, 21]],
          [7, [250, 204, 21]], [8, [217, 48, 37]], [9, [217, 48, 37]], [10, [22, 163, 74]]]) {
          const top = (row - 3) * 38;
          const ink = [];
          for (let y = top + 1; y < top + 37; y++) for (let x = 0; x < canvas.width; x++) {
            const offset = (y * canvas.width + x) * 4;
            if (color.every((value, channel) => Math.abs(pixels[offset + channel] - value) <= 1)) ink.push([x, y]);
          }
          const xs = ink.map(point => point[0]), ys = ink.map(point => point[1]);
          if (ink.length < 100 || Math.min(...xs) < 4 || Math.max(...xs) > 36
            || Math.abs((Math.min(...ys) + Math.max(...ys)) / 2 - (top + 18.5)) > 1)
            throw new Error(`P${row}: traffic-light pixels must be left-aligned and vertically centered`);
          let textRight = 0;
          for (let y = top + 2; y < top + 36; y++) for (let x = 40; x < canvas.width; x++) {
            const offset = (y * canvas.width + x) * 4;
            if (pixels[offset] < 80 && pixels[offset + 1] < 80 && pixels[offset + 2] < 80) textRight = Math.max(textRight, x);
          }
          // Ink excludes the trailing glyph's side bearing (notably "1" in WebKit).
          if (textRight < canvas.width - 10 || textRight > canvas.width - 3)
            throw new Error(`P${row}: number must stay aligned to the right edge (${textRight}/${canvas.width})`);
        }
        const iconsPng = canvas.toDataURL('image/png').split(',')[1];
        frame.bitmap.close(); renderer.close();
        const legend = scene.objects.filter(object => object.source.part === 'xl/charts/chart1.xml' && object.text?.startsWith('Measure1 -'));
        const chart = scene.objects.find(object => object.numericId === legend[0]?.parentNumericId);
        if (legend.length !== 6 || !chart || legend.some(({ bounds: b }) =>
          b.x < chart.bounds.x || b.y < chart.bounds.y || b.x + b.width > chart.bounds.x + chart.bounds.width
          || b.y + b.height > chart.bounds.y + chart.bounds.height)) throw new Error('Legend overflowed the actual chart');
        const categories = scene.objects.filter(object => object.source.part === 'xl/charts/chart1.xml' && object.text?.startsWith('Level1Item'));
        if (categories.length < 10 || categories.length > 20) throw new Error('Category labels were not thinned');
        const original = scene.objects.find(object => object.visual.kind === 'advanced-effect' && object.visual.outerShadow);
        const results = [];
        for (const format of ['xlsx', 'docx', 'pptx']) for (const variant of [0, 1, 2]) {
          const object = { ...original, parentNumericId: undefined, parentId: undefined, source: { ...original.source, format } };
          if (variant === 1) object.visual = { ...original.visual, outerShadow: {
            ...original.visual.outerShadow, scaleX: -1.2, scaleY: 0.7, skewX: 20, skewY: -10, offsetX: 25, offsetY: -8,
          } };
          const request = { unitIndex: 0, viewport: variant === 2 ? { x: 94, y: 260, width: 50, height: 70 }
            : { x: 0, y: 0, width: 640, height: 480 }, scale: variant === 1 ? 1.3 : 1, background: '#ffffff' };
          const pixels = [];
          for (const cropped of [false, true]) {
            // An identity wrapper preserves the same line but uses the unchanged
            // general shadow path as the pixel oracle for the bounded fast path.
            const input = cropped ? object : { ...object, visual: { ...object.visual, visual: {
              kind: 'layer', transform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }, opacity: 1, visual: object.visual.visual,
            } } };
            const renderer = new SceneRenderer([input], DEFAULT_LIMITS);
            const frame = await renderer.render(scene.info.units[0], request);
            const canvas = new OffscreenCanvas(frame.pixelWidth, frame.pixelHeight);
            const context = canvas.getContext('2d');
            context.drawImage(frame.bitmap, 0, 0);
            pixels.push(context.getImageData(0, 0, canvas.width, canvas.height).data);
            frame.bitmap.close(); renderer.close();
          }
          let maxDelta = 0, changed = 0;
          for (let index = 0; index < pixels[0].length; index++) {
            const delta = Math.abs(pixels[0][index] - pixels[1][index]);
            maxDelta = Math.max(maxDelta, delta); if (delta) changed++;
          }
          results.push({ format, variant, maxDelta, changed });
        }
        doc.close(); core.close();
        return { results, iconsPng };
      });
      await writeFile(`.cache/olap-pivot/${name}-icons.png`, Buffer.from(parity.iconsPng, 'base64'));
      console.log(`${name}: traffic-light placement passed; shadow pixel comparisons ${JSON.stringify(parity.results)}`);
      assert.ok(parity.results.every(result => result.maxDelta <= 1), `${name}: shadow pixels changed: ${JSON.stringify(parity.results)}`);
      console.log(`${name}: OlapPivotA3 opened in ${openedMs}ms; first-render canvas allocation ${allocations[0]} pixels; horizontal return scroll passed`);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
