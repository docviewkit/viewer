import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { chromium, firefox, webkit } from 'playwright-core';

const server = spawn(process.execPath, ['scripts/serve.mjs', '--port', '0'], { stdio: ['ignore', 'pipe', 'inherit'] });
try {
  const url = await new Promise((resolve, reject) => {
    server.once('error', reject);
    server.once('exit', code => reject(new Error(`Server exited: ${code}`)));
    createInterface({ input: server.stdout }).on('line', line => {
      const match = line.match(/https?:\/\/\S+/u);
      if (match) resolve(match[0]);
    });
  });
  for (const name of (process.env.DOCVIEWKIT_BROWSERS ?? 'chromium,firefox,webkit').split(',')) {
    const browser = await ({ chromium, firefox, webkit })[name].launch({ headless: true });
    const timeout = setTimeout(() => {
      console.error(`${name}: extrusion regression exceeded 90 seconds`);
      void browser.close();
    }, 90_000);
    try {
      const page = await browser.newPage();
      const errors = [];
      page.on('pageerror', error => errors.push(error.message));
      if (process.env.EXTRUSION_BASELINE) await page.route('**/dist/render-baseline.js', route =>
        route.fulfill({ path: process.env.EXTRUSION_BASELINE, contentType: 'text/javascript' }));
      await page.goto(`${url}examples/inspector.html`);
      const results = await page.evaluate(async baseline => {
        const { Core, DEFAULT_LIMITS } = await import('/dist/core.js');
        const { SceneRenderer } = await import('/dist/render.js');
        const Before = baseline ? (await import('/dist/render-baseline.js')).SceneRenderer : undefined;
        const { registerFonts, collectFontRequests } = await import('/dist/font.js');
        const core = await Core.create(await (await fetch('/dist/office-viewer-core.wasm')).arrayBuffer(), DEFAULT_LIMITS);
        const results = [];
        const filteredCanvas = 'filter' in new OffscreenCanvas(1, 1).getContext('2d');
        try {
          for (const file of ['Text_withExtrusion_200chars.pptx', 'Text_withEffects_100chars.pptx', 'image-3d-rotation.pptx', 'smartart-orgchart-3d.pptx']) {
            const doc = core.open(new Uint8Array(await (await fetch(`/tests/fixtures/${file}`)).arrayBuffer()));
            doc.loadUnit(0);
            const fonts = await registerFonts([], undefined, undefined, undefined, [], [],
              { requestedFaces: collectFontRequests(doc.scene.objects) });
            try {
              const requests = [{ pixelRatio: 1 }, { pixelRatio: 2 }];
              if (file === 'Text_withExtrusion_200chars.pptx') requests.push({ pixelRatio: 1, scale: 1.25,
                viewport: { x: 150, y: 150, width: 650, height: 450 } });
              for (const request of requests) {
                const { pixelRatio } = request;
                const render = async (Renderer, objects = doc.scene.objects) => {
                  const Native = OffscreenCanvas;
                  let pixels = 0;
                  globalThis.OffscreenCanvas = class extends Native {
                    constructor(w, h) { super(w, h); pixels += w * h; }
                  };
                  const renderer = new Renderer(objects, DEFAULT_LIMITS, fonts);
                  try {
                    const start = performance.now();
                    const frame = await renderer.render(doc.scene.info.units[0], { unitIndex: 0, ...request });
                    const canvas = new Native(frame.bitmap.width, frame.bitmap.height);
                    const context = canvas.getContext('2d');
                    context.drawImage(frame.bitmap, 0, 0); frame.bitmap.close();
                    const data = context.getImageData(0, 0, canvas.width, canvas.height).data;
                    return { ms: performance.now() - start, pixels, data };
                  } finally { renderer.close(); globalThis.OffscreenCanvas = Native; }
                };
                const measure = async Renderer => {
                  const samples = [];
                  for (let i = 0; i < 3; i++) samples.push(await render(Renderer));
                  return samples.sort((a, b) => a.ms - b.ms)[1];
                };
                const before = Before ? await measure(Before) : undefined;
                const after = await measure(SceneRenderer);
                let extrusionInk;
                if (file === 'Text_withExtrusion_200chars.pptx') {
                  const objects = structuredClone(doc.scene.objects);
                  const strip = visual => {
                    if (visual.threeD) visual.threeD.extrusionHeight = 0;
                    if (visual.visual) strip(visual.visual);
                  };
                  objects.forEach(object => strip(object.visual));
                  const flat = await render(SceneRenderer, objects);
                  extrusionInk = after.data.reduce((count, value, i) => count + (Math.abs(value - flat.data[i]) > 10), 0);
                }
                let differences = 0, maxDelta = 0;
                if (before) for (let i = 0; i < after.data.length; i++) {
                  const delta = Math.abs(after.data[i] - before.data[i]);
                  if (delta) differences++;
                  maxDelta = Math.max(maxDelta, delta);
                }
                // WebKit can vary a few antialiased channels even against itself.
                // Accept only values reproduced by another unchanged render.
                let unexplained = differences;
                if (differences && Before && file !== 'Text_withExtrusion_200chars.pptx') {
                  const repeated = await render(Before);
                  unexplained = 0;
                  for (let i = 0; i < after.data.length; i++) {
                    if (after.data[i] < Math.min(before.data[i], repeated.data[i])
                      || after.data[i] > Math.max(before.data[i], repeated.data[i])) unexplained++;
                  }
                }
                results.push({ file, ...request, ms: after.ms, pixels: after.pixels, extrusionInk,
                  beforeMs: before?.ms, beforePixels: before?.pixels, differences, maxDelta, unexplained, filteredCanvas });
              }
            } finally { fonts.close(); doc.close(); }
          }
        } finally { core.close(); }
        return results;
      }, !!process.env.EXTRUSION_BASELINE);
      console.log(name, JSON.stringify(results));
      for (const result of results) {
        assert.equal(result.unexplained, 0, `${name} ${result.file} @${result.pixelRatio}: unchanged pixels`);
        if (result.extrusionInk !== undefined) assert.ok(result.extrusionInk > 1_000, 'authored extrusion remains visible');
        if (result.filteredCanvas && result.file === 'Text_withExtrusion_200chars.pptx') {
          assert.ok(result.pixels < 15_000_000 * result.pixelRatio ** 2,
            `${name}: extrusion must not allocate a full slide per depth layer (${result.pixels} pixels)`);
        }
      }
      assert.deepEqual(errors, []);
    } finally { clearTimeout(timeout); await browser.close(); }
  }
} finally { server.kill('SIGTERM'); }
