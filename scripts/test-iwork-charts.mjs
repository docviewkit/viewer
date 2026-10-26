import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { chromium, webkit } from 'playwright-core';
import { Core, DEFAULT_LIMITS } from '../dist/core.js';

const output = process.env.IWORK_CHART_OUTPUT ?? 'output/iwork-chart-coverage/regression';
const baseline = process.env.IWORK_CHART_BASELINE;
const matrix = await jsonMatrix();
async function jsonMatrix() { return JSON.parse(await readFile('tests/fixtures/iwork-chart-coverage.json', 'utf8')); }
const incomplete = matrix.rows.filter(r => r.scope === 'static' && Object.values(r.hosts).some(h => h.status !== 'passed' || h.nativeVisual !== 'passed')).map(r => r.name);
const incompleteFeatures = matrix.requiredFeatures.filter(f => f.status !== 'passed').map(f => f.id);
const failures = [];
const checks = [];
const check = (name, passed) => { checks.push({ name, passed }); if (!passed) failures.push(name); };
const sha256 = bytes => createHash('sha256').update(bytes).digest('hex');
const json = async path => JSON.parse(await readFile(path, 'utf8'));
const fixtures = [
  ['iwork-charting-basics.numbers', 11],
  ['iwork-extended-charts.key', 4],
  ['iwork-imported-charts.pages', 4],
  ['iwork-extended-charts.pages', 1],
  ['iwork-static-charts.key', 17],
  ['iwork-manual-axis.key', 1],
];
await mkdir(output, { recursive: true });
for (const fixture of matrix.fixtures) assert.equal(sha256(await readFile(`tests/fixtures/${fixture.file}`)), fixture.sha256, `${fixture.file} changed after coverage review`);
// The native PDF remains the oracle. This checks its source and artifact chain,
// then compares extracted native vector geometry, independently of our renderer.
for (const host of ['numbers', 'keynote', 'pages', 'manual-axis']) {
  const folder = `tests/fixtures/iwork-chart-oracles/${host}`;
  const reference = await json(`${folder}/reference.json`);
  assert.equal(reference.oracleSuite, 'apple-iwork');
  assert.equal(reference.referenceApplication.version, '15.4');
  assert.equal(sha256(await readFile(`tests/fixtures/${reference.source.file}`)), reference.source.sha256);
  assert.equal(sha256(await readFile(`${folder}/reference.pdf`)), reference.pdfs[0].sha256);
}
const native = await json('tests/fixtures/iwork-chart-oracles/numbers/geometry.json');
assert.equal(sha256(await readFile(`tests/fixtures/${native.fixture}`)), native.fixtureSha256);
assert.equal(sha256(await readFile('tests/fixtures/iwork-chart-oracles/numbers/reference.pdf')), native.pdfSha256);
const unwrap = visual => visual.visual ? unwrap(visual.visual) : visual;
const core = await Core.create(await readFile(baseline ?? 'dist/office-viewer-pdf.wasm'), DEFAULT_LIMITS);
try {
  for (const [file, count] of fixtures) {
    const doc = core.open(await readFile(`tests/fixtures/${file}`));
    try {
      doc.loadAll();
      const scene = doc.scene;
      check(`${file}: chart count ${count}`, scene.objects.filter(o => o.source?.kind === 'chart').length === count);
      check(`${file}: native objects only`, !scene.objects.some(o => /preview/u.test(o.source?.kind ?? '')));
      if (file === 'iwork-static-charts.key') {
        check(`${file}: eight explicit 3D approximations`, scene.diagnostics.filter(d => d.message.startsWith('IWORK_CHART_3D_APPROXIMATE:')).length === 8);
      }
      if (file === 'iwork-manual-axis.key') {
        const oracle = await json('tests/fixtures/iwork-chart-oracles/manual-axis/geometry.json');
        const labels = scene.objects.map(o => unwrap(o.visual).runs?.map(r => r.text).join('')).filter(Boolean);
        check('Keynote: authored manual range and tick precision', oracle.labels.every(t => labels.includes(t)));
        const bars = scene.objects.filter(o => { const v = unwrap(o.visual); return o.source?.kind === 'chart-series' && v.kind === 'painted-shape' && v.fill.kind === 'solid' && o.bounds.width > 10 && o.bounds.height > 10; });
        check('Keynote: manual-axis bar geometry matches native PDF', bars.length === 3 && bars.every((o, i) => {
          const b = o.bounds; const [x, y, w, h] = oracle.bars[i]; return [b.x-x, b.y-y, b.x+b.width-x-w, b.y+b.height-y-h].every(v => Math.abs(v) < 1.5);
        }));
      }
      if (file !== native.fixture) continue;
      for (const radar of native.radar) {
        const shape = scene.objects.find(o => o.id.endsWith(`:radar-series-${radar.series}`));
        const visual = shape && unwrap(shape.visual);
        const points = visual?.geometry?.commands?.filter(p => ['moveTo', 'lineTo'].includes(p.kind)).map(p => [p.x + shape.bounds.x, p.y + shape.bounds.y]) ?? [];
        check(`Numbers: native radar series ${radar.series} vertices`, radar.points.every(p => points.some(a => Math.hypot(a[0] - p[0], a[1] - p[1]) < 1)));
      }
      for (const [i, bubble] of native.bubbles.entries()) {
        const shape = scene.objects.find(o => o.id.endsWith(`:bubble-0-${i}`));
        const b = shape?.bounds;
        check(`Numbers: native bubble ${i} center and radius`, b !== undefined
          && Math.hypot(b.x + b.width / 2 - bubble.center[0], b.y + b.height / 2 - bubble.center[1]) < 1.5
          && Math.abs(b.width / 2 - bubble.radius) < 1.5);
      }
      check('Numbers: bubble has no unauthored connection lines', !scene.objects.some(o => o.unitIndex === 7 && /:scatter-line-/u.test(o.id)));
      const markers = scene.objects.filter(o => o.unitIndex === 5 && /:line-marker-/u.test(o.id));
      check('Numbers: two-axis line retains rainfall series and native centers', markers.length === native.dualMarkers.length && markers.every((o, i) =>
        Math.hypot(o.bounds.x + o.bounds.width / 2 - native.dualMarkers[i][0], o.bounds.y + o.bounds.height / 2 - native.dualMarkers[i][1]) < 1.5));
    } finally { doc.close(); }
  }
} finally { core.close(); }

const server = spawn(process.execPath, ['scripts/serve.mjs', '--port', '0'], { stdio: ['ignore', 'pipe', 'inherit'] });
try {
  const url = await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('Server startup timed out')), 10000);
    createInterface({ input: server.stdout }).on('line', line => {
      const match = line.match(/https?:\/\/\S+/u);
      if (match) { clearTimeout(timer); resolve(match[0]); }
    });
    server.once('exit', code => { clearTimeout(timer); reject(new Error(`Server exited: ${code}`)); });
  });
  for (const [name, browserType] of Object.entries({ chromium, webkit })) {
    const browser = await browserType.launch({ headless: true });
    try {
      const page = await browser.newPage({ viewport: { width: 1400, height: 1000 }, reducedMotion: 'reduce' });
      const errors = [];
      page.on('pageerror', e => errors.push(e.message));
      page.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
      if (baseline) {
        const bytes = await readFile(baseline);
        await page.route('**/dist/office-viewer-pdf.wasm', r => r.fulfill({ contentType: 'application/wasm', body: bytes }));
      }
      await page.goto(`${url}examples/viewer.html?filePicker=true`);
      await page.waitForFunction(() => document.querySelector('#viewer').state.status === 'ready');
      for (const [file, count] of fixtures) {
        await page.locator('#viewer').evaluate(async (viewer, file) => {
          await viewer.open(await (await fetch(`/tests/fixtures/${file}`)).arrayBuffer());
        }, file);
        const result = await page.evaluate(async file => {
          const { createOfficeEngine } = await import('/dist/engine.js');
          const { extendedFormatPack } = await import('/dist/extended-formats.js');
          const engine = await createOfficeEngine({ execution: 'inline', formatPack: async () => extendedFormatPack });
          const doc = await engine.open(await (await fetch(`/tests/fixtures/${file}`)).arrayBuffer());
          try {
            const objects = await doc.listObjects();
            const charts = objects.filter(o => o.source.kind === 'chart');
            const captures = [];
            for (const chart of charts.filter(c => !file.endsWith('.numbers') || [2, 5, 7].includes(c.unitIndex))) {
              const b = chart.bounds;
              const viewport = { x: Math.max(0, b.x - 60), y: Math.max(0, b.y - 40), width: b.width + 150, height: b.height + 120 };
              const frame = await doc.render({ unitIndex: chart.unitIndex, viewport, scale: 1.5, pixelRatio: 1, includeTextFragments: true });
              try {
                const canvas = document.createElement('canvas'); canvas.width = frame.pixelWidth; canvas.height = frame.pixelHeight;
                const ctx = canvas.getContext('2d'); ctx.drawImage(frame.bitmap, 0, 0);
                const pixels = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
                let coloredPixels = 0;
                for (let i = 0; i < pixels.length; i += 4) if (pixels[i + 3] > 100 && Math.max(pixels[i], pixels[i + 1], pixels[i + 2]) - Math.min(pixels[i], pixels[i + 1], pixels[i + 2]) > 40) coloredPixels++;
                captures.push({ unit: chart.unitIndex, image: canvas.toDataURL(), viewport, coloredPixels, labels: frame.textFragments.map(t => t.text).join('') });
              } finally { frame.bitmap.close(); }
            }
            return { charts, captures };
          } finally { doc.close(); engine.close(); }
        }, file);
        check(`${name} ${file}: public API chart count`, result.charts.length === count);
        for (const capture of result.captures) {
          await writeFile(`${output}/${name}-${file}-${capture.unit}.png`, Buffer.from(capture.image.split(',')[1], 'base64'));
          delete capture.image;
          check(`${name} ${file} unit ${capture.unit}: chart color rasterized`, capture.coloredPixels > 100);
          if (file === native.fixture) {
            const labels = { 2: ['Nutty', 'Roasted', 'Spicy', 'Floral', 'Fruity'], 5: ['Jan', 'Dec', '90'], 7: ['Salespeople', 'Units Sold'] }[capture.unit];
            check(`${name} Numbers unit ${capture.unit}: authored labels`, labels.every(label => capture.labels.includes(label)));
          }
        }
        await writeFile(`${output}/${name}-${file}.json`, `${JSON.stringify(result, null, 2)}\n`);
        if (file !== 'iwork-static-charts.key') await page.screenshot({ path: `${output}/${name}-${file}-viewer.png` });
      }
      check(`${name}: clean console`, errors.length === 0);
      if (errors.length) failures.push(...errors);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
await writeFile(`${output}/report.json`, `${JSON.stringify({ baseline: baseline ?? null, checks, failures, incomplete, incompleteFeatures }, null, 2)}\n`);
console.log(`iWork charts: ${checks.filter(c => c.passed).length}/${checks.length} checks passed; ${output}/report.json`);
assert.deepEqual(failures, [], 'iWork real-file/native geometry/browser regression');

console.log(`Static completeness: ${incomplete.length} chart types and ${incompleteFeatures.length} feature groups remain unaccepted. See docs/research/IWORK-CHART-COVERAGE.md.`);
if (process.argv.includes('--require-complete')) assert.deepEqual([...incomplete, ...incompleteFeatures], [], 'iWork static chart completeness gate');
