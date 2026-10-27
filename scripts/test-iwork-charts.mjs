import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFile, writeFile, mkdir, readdir } from 'node:fs/promises';
import { spawn, spawnSync } from 'node:child_process';
import { createInterface } from 'node:readline';
import { chromium, webkit } from 'playwright-core';
import { Core, DEFAULT_LIMITS } from '../dist/core.js';
import { evaluateAccuracy } from '../dist/accuracy.js';
import { NATIVE_VISUAL_POLICY, decodePng, nativeVisualRegions } from './native-visual-policy.mjs';

const output = process.env.IWORK_CHART_OUTPUT ?? 'output/iwork-chart-coverage/regression';
const baseline = process.env.IWORK_CHART_BASELINE;
const rendererBaseline = process.env.IWORK_CHART_RENDERER_BASELINE;
const onlyFixture = process.env.IWORK_CHART_FIXTURE;
const requireComplete = process.argv.includes('--require-complete');
assert.ok(!requireComplete || onlyFixture === undefined, 'Completeness requires the entire fixture corpus');
const fullVisual = requireComplete || process.argv.includes('--native-visual');
const nativeVisual = [];
const nativePngs = new Map();
const nativeViewports = new Map();
const matrix = await jsonMatrix();
async function jsonMatrix() { return JSON.parse(await readFile('tests/fixtures/iwork-chart-coverage.json', 'utf8')); }
const incomplete = matrix.rows.filter(r => r.scope === 'static' && Object.values(r.hosts).some(h => h.status !== 'passed' || h.nativeVisual !== 'passed')).map(r => r.name);
const incompleteFeatures = matrix.requiredFeatures.filter(f => f.status !== 'passed').map(f => f.id);
const failures = [];
const checks = [];
const check = (name, passed) => { checks.push({ name, passed }); if (!passed) failures.push(name); };
const sha256 = bytes => createHash('sha256').update(bytes).digest('hex');
const json = async path => JSON.parse(await readFile(path, 'utf8'));
const labelOracleFolders = {'iwork-chart-formatting.key':'formatting','iwork-chart-auto-axis.key':'auto-axis','iwork-chart-label-boundaries.key':'axis-labels','iwork-chart-radar-labels.key':'radar-labels','iwork-chart-radar-value-axis.key':'radar-value-axis','iwork-chart-single-sided-types.key':'single-sided-types','iwork-chart-cross-axis.key':'cross-axis','iwork-chart-cross-formats.key':'cross-formats','auto-count-heldout.key':'auto-count','auto-count.key':'auto-count-controls','one-sided-currency.key':'one-sided-currency','one-sided-number-two.key':'one-sided-number-two','3d-signed-area.key':'3d-signed-area','stack-cross-axis.key':'stack-cross-axis','stack-opposite-signs.key':'stack-opposite-signs','stack-sparse-negative.key':'stack-sparse-negative','stack-cross-boundaries.key':'stack-cross-boundaries','stack-cross-boundaries-2d.key':'stack-cross-boundaries-2d','axis-tick-boundaries.key':'axis-tick-boundaries','axis-tick-horizontal.key':'axis-tick-horizontal','axis-tick-logarithmic.key':'axis-tick-logarithmic','axis-tick-logarithmic-ranges.key':'axis-tick-logarithmic-ranges'};
const worldDiffuseFixtures = ['xp','xm','yp','ym','zp','zm','stacked-xm'].map(n=>`3d-world-diffuse-${n}.key`);
const worldDiffuseProbes = new Map();
const defaultBarLightingFixtures = ['3d-default-lighting.key','3d-default-lighting-stack-source.key'];

const fixtures = [
  ['iwork-charting-basics.numbers', 11],
  ['iwork-extended-charts.key', 4],
  ['iwork-imported-charts.pages', 4],
  ['iwork-extended-charts.pages', 1],
  ['iwork-static-charts.key', 17],
  ['iwork-manual-axis.key', 1],
  ['iwork-chart-boundaries.key', 8],
  ['iwork-chart-formatting.key', 10],
  ['iwork-chart-auto-axis.key', 31],
  ['iwork-chart-label-boundaries.key', 6],
  ['iwork-chart-radar-grid.key', 3],
  ['iwork-chart-radar-labels.key', 3],
  ['iwork-chart-radar-value-axis.key', 2],
  ['iwork-chart-single-sided-types.key', 11],
  ['iwork-chart-cross-axis.key', 28],
  ['iwork-chart-cross-formats.key', 4],
  ['auto-count-heldout.key', 16],
  ['auto-count.key', 4],
  ['one-sided-currency.key', 9],
  ['one-sided-number-two.key', 9],
  ['iwork-chart-3d-clipped.key', 4],
  ['3d-default-lighting.key', 2],
  ['3d-default-lighting-stack-source.key', 2],
  ['3d-light-off-unique.key', 2],
  ['3d-center-intensity-four.key', 2],
  ['3d-six-default-lighting.key', 12],
  ['3d-six-implicit-lighting.key', 12],
  ...worldDiffuseFixtures.map(file=>[file,2]),
  ['iwork-chart-3d-series-gap.key', 7],
  ['iwork-chart-3d-sparse-area.key', 3],
  ['3d-zero-area.key', 5],
  ['3d-flat-area.key', 2],
  ['3d-signed-area.key', 5],
  ['stack-cross-axis.key', 12],
  ['stack-opposite-signs.key', 6],
  ['stack-sparse-negative.key', 4],
  ['stack-cross-boundaries.key', 5],
  ['stack-cross-boundaries-2d.key', 5],
  ['axis-tick-boundaries.key', 5],
  ['axis-tick-horizontal.key', 3],
  ['axis-tick-logarithmic.key', 2],
  ['axis-tick-logarithmic-ranges.key', 5],
  ['iwork-chart-edge-cases.key', 10],
  ['iwork-static-matrix.key', 22],
  ['iwork-static-matrix.numbers', 22],
  ['iwork-static-matrix-inline.pages', 22],
  ['iwork-static-matrix-floating.pages', 22],
  ['numbers-multiple-sheets-charts.numbers', 2],
  ['numbers-dice-lab.numbers', 2],
  ['numbers-team-budget.numbers', 2],
  ['numbers-retirement.numbers', 1],
  ['chart-original.numbers', 1],
].filter(([file]) => onlyFixture === undefined || file === onlyFixture || (onlyFixture==='3d-six-implicit-lighting.key'&&file==='3d-six-default-lighting.key'));
assert.ok(fixtures.length > 0, 'Unknown iWork chart fixture filter');
await mkdir(output, { recursive: true });
for (const fixture of matrix.fixtures) assert.equal(sha256(await readFile(`tests/fixtures/${fixture.file}`)), fixture.sha256, `${fixture.file} changed after coverage review`);
// The native PDF remains the oracle. This checks its source and artifact chain,
// then compares extracted native vector geometry, independently of our renderer.
for (const host of ['numbers', 'keynote', 'pages', 'manual-axis', 'boundaries', 'existing-multiple', 'existing-dice', 'existing-team', 'existing-retirement', 'existing-original', 'static-keynote', 'formatting', 'auto-axis', 'axis-labels', 'radar-grid', '3d-clipped', '3d-default-lighting', '3d-default-lighting-stack-source', '3d-light-off-unique', '3d-center-intensity-four', '3d-six-default-lighting', '3d-six-implicit-lighting', ...worldDiffuseFixtures.map(file=>file.slice(0,-4)), '3d-series-gap', '3d-sparse-area', '3d-zero-area', '3d-flat-area', '3d-signed-area', 'stack-cross-axis', 'stack-opposite-signs', 'stack-sparse-negative', 'stack-cross-boundaries', 'stack-cross-boundaries-2d', 'axis-tick-boundaries', 'axis-tick-horizontal', 'axis-tick-logarithmic', 'axis-tick-logarithmic-ranges', 'radar-labels', 'radar-value-axis', 'single-sided-types', 'cross-axis', 'cross-formats', 'auto-count', 'auto-count-controls', 'one-sided-currency', 'one-sided-number-two', 'edge-cases', 'matrix-keynote', 'matrix-inline-pages', 'matrix-floating-pages', 'matrix-numbers']) {
  const folder = `tests/fixtures/iwork-chart-oracles/${host}`;
  const reference = await json(`${folder}/reference.json`);
  if (onlyFixture !== undefined && reference.source.file !== onlyFixture && !(onlyFixture==='3d-six-implicit-lighting.key'&&reference.source.file==='3d-six-default-lighting.key')) continue;
  assert.equal(reference.oracleSuite, 'apple-iwork');
  assert.equal(reference.referenceApplication.version, '15.4');
  assert.equal(sha256(await readFile(`tests/fixtures/${reference.source.file}`)), reference.source.sha256);
  assert.equal(sha256(await readFile(`${folder}/reference.pdf`)), reference.pdfs[0].sha256);
  if (host==='static-keynote') {
    const geometry=await json(`${folder}/geometry.json`);
    assert.equal(geometry.threeDimensionalPie.nativePageSha256,reference.pages[geometry.threeDimensionalPie.unit].sha256);
  }
  if (['formatting', 'auto-axis', 'axis-labels', 'radar-grid', '3d-clipped', '3d-default-lighting', '3d-default-lighting-stack-source', '3d-light-off-unique', '3d-center-intensity-four', '3d-six-default-lighting', '3d-six-implicit-lighting', ...worldDiffuseFixtures.map(file=>file.slice(0,-4)), '3d-series-gap', '3d-sparse-area', '3d-zero-area', '3d-flat-area', '3d-signed-area', 'stack-cross-axis', 'stack-opposite-signs', 'stack-sparse-negative', 'stack-cross-boundaries', 'stack-cross-boundaries-2d', 'axis-tick-boundaries', 'axis-tick-horizontal', 'axis-tick-logarithmic', 'axis-tick-logarithmic-ranges', 'radar-labels', 'radar-value-axis', 'single-sided-types', 'cross-axis', 'cross-formats', 'auto-count', 'auto-count-controls', 'one-sided-currency', 'one-sided-number-two', 'radar-grid', 'edge-cases', 'existing-original', 'matrix-floating-pages'].includes(host)) {
    const geometry = await json(`${folder}/geometry.json`);
    assert.equal(geometry.fixtureSha256, reference.source.sha256);
    assert.equal(geometry.pdfSha256, reference.pdfs[0].sha256);
    if (['axis-labels','single-sided-types','cross-axis'].includes(host)) for (const unit of geometry.units.filter(u=>u.nativePageSha256)) {
      assert.equal(unit.nativePageSha256, reference.pages[unit.unit].sha256);
      assert.equal(unit.nativePageSha256, sha256(await readFile(`${folder}/page-${String(unit.unit+1).padStart(4,'0')}.png`)));
    }
  }
}
if (fullVisual) {
  // Numbers uses actual 100% sheet captures, independently of print pagination.
  if (onlyFixture === undefined || onlyFixture === 'iwork-static-matrix.numbers') {
    const folder = 'tests/fixtures/iwork-chart-oracles/matrix-numbers/viewports';
    const reference = await json(`${folder}/reference.json`);
    assert.equal(reference.fixtureSha256, sha256(await readFile(`tests/fixtures/${reference.fixture}`)));
    assert.equal(reference.sourceCopyBeforeAndAfterSha256, reference.fixtureSha256);
    assert.equal(reference.nativeZoomPercent, 100);
    assert.equal(reference.engineScale, reference.nativeDevicePixelRatio / (4/3));
    const [px,py,pw,ph] = reference.sourcePointViewport;
    const [cx,cy,cw,ch] = reference.cropWindowPixels;
    assert.deepEqual([cx,cy,cw,ch], [
      reference.nativeSheetOriginWindowPixels[0]+px*reference.nativeDevicePixelRatio,
      reference.nativeSheetOriginWindowPixels[1]+py*reference.nativeDevicePixelRatio,
      pw*reference.nativeDevicePixelRatio,ph*reference.nativeDevicePixelRatio,
    ]);
    for (const [index,key] of ['x','y','width','height'].entries())
      assert.ok(Math.abs(reference.engineViewport[key]-reference.sourcePointViewport[index]*4/3)<1e-9);
    assert.equal(reference.units.length, 22);
    const pages = new Map(), frames = new Map();
    for (const unit of reference.units) {
      const bytes = await readFile(`${folder}/${unit.file}`);
      assert.equal(sha256(bytes), unit.sha256);
      assert.equal(sha256(await readFile(`${folder}/${unit.rawWindowFile}`)), unit.rawWindowSha256);
      const ax = await readFile(`${folder}/${unit.axFile}`);
      assert.equal(sha256(ax), unit.axSha256);
      assert.match(ax.toString(), /menu button Zoom[^\n]*Value: 100%/u);
      assert.ok(ax.toString().includes(`layout area Description: ${unit.sheetName},`));
      assert.ok(!ax.toString().includes('layout item (selected)'));
      const image = decodePng(bytes, `Numbers native viewport ${unit.unit}`);
      assert.deepEqual([image.width,image.height], [cw,ch]);
      pages.set(unit.unit, image);
      frames.set(unit.unit, {...reference.engineViewport, scale:reference.engineScale, sheetName:unit.sheetName});
    }
    nativePngs.set(reference.fixture, pages);
    nativeViewports.set(reference.fixture, frames);
  }
  for (const host of ['static-keynote', 'matrix-keynote', 'matrix-inline-pages', 'matrix-floating-pages']) {
    const folder = `tests/fixtures/iwork-chart-oracles/${host}`;
    const reference = await json(`${folder}/reference.json`);
    if (onlyFixture !== undefined && reference.source.file !== onlyFixture) continue;
    const raster = `${output}/native-${host}`;
    await mkdir(raster, { recursive: true });
    const run = spawnSync('pdftoppm', ['-png', '-r', '96', '-cropbox', '-hide-annotations', '-freetype', 'yes', '-aa', 'yes', '-aaVector', 'yes', '-thinlinemode', 'none', '-forcenum', `${folder}/reference.pdf`, `${raster}/page`], { encoding: 'utf8' });
    assert.equal(run.status, 0, `Native PDF rasterization failed: ${run.stderr}`);
    const pages = new Map();
    for (const file of await readdir(raster)) {
      const match = file.match(/^page-(\d+)\.png$/u);
      if (!match) continue;
      const unit = Number(match[1])-1;
      const bytes = await readFile(`${raster}/${file}`);
      assert.equal(sha256(bytes), reference.pages[unit].sha256, 'Native raster drift: use the pinned rasterizer and flags');
      pages.set(unit, decodePng(bytes, `${reference.source.file} native unit ${unit}`));
    }
    assert.equal(pages.size, reference.pages.length);
    nativePngs.set(reference.source.file, pages);
  }
}
const native = await json('tests/fixtures/iwork-chart-oracles/numbers/geometry.json');
assert.equal(sha256(await readFile(`tests/fixtures/${native.fixture}`)), native.fixtureSha256);
assert.equal(sha256(await readFile('tests/fixtures/iwork-chart-oracles/numbers/reference.pdf')), native.pdfSha256);
const unwrap = visual => visual.visual ? unwrap(visual.visual) : visual;
const distanceToSegment = (p, a, b) => {
  const dx=b[0]-a[0], dy=b[1]-a[1];
  const t=Math.max(0,Math.min(1,((p[0]-a[0])*dx+(p[1]-a[1])*dy)/(dx*dx+dy*dy || 1)));
  return Math.hypot(p[0]-a[0]-t*dx,p[1]-a[1]-t*dy);
};
const closeBoundaries = (a, b) => [ [a,b], [b,a] ].every(([from,to]) =>
  from.every(p=>to.some((q,i)=>i>0 && distanceToSegment(p,to[i-1],q)<1.5)));
const rasterRgbAt = (image,x,y,radius,accept) => {
  for(let yy=Math.max(0,Math.floor(y-radius));yy<=Math.min(image.height-1,Math.ceil(y+radius));yy++)
    for(let xx=Math.max(0,Math.floor(x-radius));xx<=Math.min(image.width-1,Math.ceil(x+radius));xx++) {
      const offset=(yy*image.width+xx)*4;
      if(accept(image.data.subarray(offset,offset+3)))return true;
    }
  return false;
};
const chromaticMask = image => {
  const mask=new Uint8Array(image.width*image.height);
  for(let p=0;p<mask.length;p++) {
    const rgb=image.data.subarray(p*4,p*4+3);
    mask[p]=image.data[p*4+3]>100 && Math.max(...rgb)-Math.min(...rgb)>40?1:0;
  }
  return mask;
};
const coloredConvexCoverage = image => {
  const mask=chromaticMask(image),points=[];
  let colored=0;
  for(let y=0;y<image.height;y++) {
    let left=image.width,right=-1;
    for(let x=0;x<image.width;x++)if(mask[y*image.width+x]){left=Math.min(left,x);right=x;colored++;}
    if(right>=0)points.push([left,y],[right+1,y],[left,y+1],[right+1,y+1]);
  }
  points.sort((a,b)=>a[0]-b[0]||a[1]-b[1]);
  const cross=(a,b,c)=>(b[0]-a[0])*(c[1]-a[1])-(b[1]-a[1])*(c[0]-a[0]);
  const chain=points=>{const result=[];for(const p of points){while(result.length>1&&cross(result.at(-2),result.at(-1),p)<=0)result.pop();result.push(p);}return result;};
  const hull=[...chain(points).slice(0,-1),...chain([...points].reverse()).slice(0,-1)];
  const area=Math.abs(hull.reduce((sum,a,i)=>{const b=hull[(i+1)%hull.length];return sum+a[0]*b[1]-a[1]*b[0];},0))/2;
  return area>0?colored/area:0;
};
const rasterComponents = (image, relativeXCrop, mode='chromatic') => {
  const mask=mode==='dark'?Uint8Array.from({length:image.width*image.height},(_,p)=>image.data[p*4+3]>100&&Math.max(...image.data.subarray(p*4,p*4+3))<40?1:0):chromaticMask(image);
  const queue=new Uint32Array(mask.length);
  if(relativeXCrop) {
    let left=image.width,right=0;
    for(let p=0;p<mask.length;p++)if(mask[p]){left=Math.min(left,p%image.width);right=Math.max(right,p%image.width);}
    const [low,high]=relativeXCrop.map(t=>Math.round(left+(right-left)*t));
    for(let p=0;p<mask.length;p++)if(p%image.width<low||p%image.width>=high)mask[p]=0;
  }
  let count=0;
  for(let start=0;start<mask.length;start++) {
    if(!mask[start])continue;
    let head=0,tail=1;queue[0]=start;mask[start]=0;
    while(head<tail){const p=queue[head++],x=p%image.width,y=Math.floor(p/image.width);
      for(let dy=-1;dy<=1;dy++)for(let dx=-1;dx<=1;dx++) {
        if(x+dx<0||x+dx>=image.width||y+dy<0||y+dy>=image.height)continue;
        const q=p+dy*image.width+dx;if(mask[q]){mask[q]=0;queue[tail++]=q;}
      }
    }
    if(tail>=(mode==='dark'?5000:16))count++;
  }
  return count;
};
const chart3dRasterFeatures = (image, valleyWindow=[.2,.45]) => {
  const blue=new Uint32Array(image.width),red=new Uint32Array(image.width),green=new Uint32Array(image.width),redTop=new Uint32Array(image.width).fill(image.height);
  let blueTop=image.height,blueBottom=-1;
  for(let p=0;p<image.width*image.height;p++) {
    const [r,g,b]=image.data.subarray(p*4,p*4+3);
    if(image.data[p*4+3]<=100)continue;
    if(b-r>40&&b-g>15){blue[p%image.width]++;blueTop=Math.min(blueTop,Math.floor(p/image.width));blueBottom=Math.max(blueBottom,Math.floor(p/image.width));}
    if(g-r>15&&g-b>40)green[p%image.width]++;
    if(r-b>40&&r-g>15){red[p%image.width]++;redTop[p%image.width]=Math.min(redTop[p%image.width],Math.floor(p/image.width));}
  }
  const left=blue.findIndex(n=>n>0),right=blue.findLastIndex(n=>n>0);
  const [low,high]=valleyWindow.map(t=>Math.round(left+(right-left)*t));
  const redPixels=red.reduce((sum,n)=>sum+n,0);
  const redLeft=red.findIndex(n=>n>0),redRight=red.findLastIndex(n=>n>0);
  const [x0,x1]=[.05,.45].map(t=>Math.round(redLeft+(redRight-redLeft)*t));
  const points=Array.from({length:Math.max(0,x1-x0)},(_,i)=>[x0+i,redTop[x0+i]]).filter(([x])=>red[x]);
  const mx=points.reduce((sum,[x])=>sum+x,0)/points.length,my=points.reduce((sum,[,y])=>sum+y,0)/points.length;
  const slope=points.reduce((sum,[x,y])=>sum+(x-mx)*(y-my),0)/points.reduce((sum,[x])=>sum+(x-mx)**2,0);
  const rms=Math.sqrt(points.reduce((sum,[x,y])=>sum+(y-my-slope*(x-mx))**2,0)/points.length);
  const bluePixels=blue.reduce((sum,n)=>sum+n,0);
  return {blueBodyCoverage:bluePixels>0?bluePixels/((right-left+1)*(blueBottom-blueTop+1)):0,greenPresent:green.reduce((sum,n)=>sum+n,0)>=16,bluePresent:blue.reduce((sum,n)=>sum+n,0)>=16,redPresent:redPixels>=16,redTopIsStraight:rms<1,
    redBandHasArea:redPixels>=blue.reduce((sum,n)=>sum+n,0)*.5,
    blueValleyNearZero:Math.min(...blue.slice(low,high))/Math.max(...blue)<.1,
    redStartsAtFirstCategory:redPixels>=16&&(red.findIndex(n=>n>0)-left)/(right-left)<.05};
};
const pathBoundary = (object) => {
  const points=[]; let previous, first;
  const point=(x,y)=>[x+object.bounds.x,y+object.bounds.y];
  for (const c of unwrap(object.visual).geometry.commands??[]) {
    if (c.kind==='moveTo') { previous=point(c.x,c.y); first=previous; points.push(previous); }
    else if(c.kind==='lineTo') { previous=point(c.x,c.y); points.push(previous); }
    else if(c.kind==='closePath') { if(first) points.push(first); previous=first; }
    else if(c.kind==='bezierCurveTo' && previous) {
      const a=previous,b=point(c.cp1x,c.cp1y),d=point(c.x,c.y),cc=point(c.cp2x,c.cp2y);
      for(let i=1;i<=32;i++) {
        const t=i/32;
        points.push([0,1].map(k=>(1-t)**3*a[k]+3*(1-t)**2*t*b[k]+3*(1-t)*t*t*cc[k]+t**3*d[k]));
      }
      previous=d;
    }
  }
  return points;
};
const solidComponents = faces => {
  const components=[];
  for(const face of faces) {
    const points=new Set(pathBoundary(face).map(p=>p.map(v=>v.toFixed(3)).join(',')));
    const touching=components.filter(c=>[...points].some(p=>c.has(p)));
    for(const c of touching){for(const p of c)points.add(p);components.splice(components.indexOf(c),1);}
    components.push(points);
  }
  return components.length;
};
const core = await Core.create(await readFile(baseline ?? 'dist/office-viewer-pdf.wasm'), DEFAULT_LIMITS);
try {
  for (const [file, count] of fixtures) {
    const doc = core.open(await readFile(`tests/fixtures/${file}`));
    try {
      doc.loadAll();
      const scene = doc.scene;
      for (const [unit,frame] of nativeViewports.get(file) ?? [])
        check(`Numbers native viewport ${unit}: source sheet identity`, scene.info.units[unit]?.name===frame.sheetName);
      if (file === 'iwork-static-matrix-floating.pages') {
        const oracle = await json('tests/fixtures/iwork-chart-oracles/matrix-floating-pages/geometry.json');
        for (const unit of oracle.units.slice(0,2)) {
          const bars = scene.objects.filter(o=>o.unitIndex===unit.unit && o.source?.kind==='chart-series')
            .filter(o=>{const v=unwrap(o.visual);const rgb=[v.fill.color>>>24,(v.fill.color>>>16)&255,(v.fill.color>>>8)&255];return v.kind==='painted-shape' && v.geometry==='rectangle' && v.fill.kind==='solid' && Math.max(...rgb)-Math.min(...rgb)>51;});
          check(`Pages floating ${unit.unit}: native bar count`, bars.length===unit.shapes.length);
          for (const [i,{bounds:[x,y,w,h]}] of unit.shapes.entries())
            check(`Pages floating ${unit.unit}: native bar boundary ${i}`, bars.some(({bounds:b})=>
              [b.x-x,b.y-y,b.x+b.width-x-w,b.y+b.height-y-h].every(v=>Math.abs(v)<1.5)));
        }
      }
      if (file === 'iwork-chart-edge-cases.key') {
        const oracle = await json('tests/fixtures/iwork-chart-oracles/edge-cases/geometry.json');
        for (const unit of oracle.units) {
          const children = scene.objects.filter(o=>o.unitIndex===unit.unit && o.source?.kind==='chart-series');
          const shapes = children.filter(o=>{const v=unwrap(o.visual);return v.kind==='painted-shape' && v.fill.kind==='solid';});
          check(`Keynote edges ${unit.unit}: native nonempty fill count`, shapes.length===unit.shapes.length);
          if (unit.unit!==0) {
            for (const [i, {bounds:[x,y,w,h]}] of unit.shapes.entries()) {
              check(`Keynote edges ${unit.unit}: native fill boundary ${i}`, shapes.some(({bounds:b})=>
                [b.x-x,b.y-y,b.x+b.width-x-w,b.y+b.height-y-h].every(v=>Math.abs(v)<1.5)));
            }
          }
          if (unit.unit===0) {
            const labels=children.flatMap(o=>unwrap(o.visual).runs?.map(r=>r.text)??[]);
            check('Keynote negative pie: native hidden data labels', labels.every(text=>unit.labels.includes(text)));
            for(const [i,shape] of unit.shapes.entries())
              check(`Keynote negative pie: native curved boundary ${i}`, shapes.some(o=>closeBoundaries(shape.boundary,pathBoundary(o))));
          }
        }
      }
      if (Object.hasOwn(labelOracleFolders,file)) {
        const oracle = await json(`tests/fixtures/iwork-chart-oracles/${labelOracleFolders[file]}/geometry.json`);
        for (const unit of oracle.units) {
          if(unit.threeDimensional)check(`${file} ${unit.unit}: actual 3D scene path`,scene.objects.some(o=>o.unitIndex===unit.unit&&o.id.includes(':3d-face-')));
          const labels = scene.objects.filter(o=>o.unitIndex===unit.unit).map(o=>unwrap(o.visual).runs?.map(r=>r.text).join('')).filter(Boolean);
          const axisLabels=labels.filter(text=>!(unit.categoryLabels??[]).includes(text)&&!(['iwork-chart-radar-labels.key','iwork-chart-radar-value-axis.key'].includes(file)?['Manual Axis 5 to 50']:['Currency','Scientific','Prefix suffix','Fractions','Manual Axis 5 to 50','A','B','C']).includes(text));
          check(`Keynote ${file} ${unit.unit}: exact native axis labels`, JSON.stringify(axisLabels.sort())===JSON.stringify([...unit.labels].sort()));
          if (unit.blackSegments) {
            const lines=[];
            for (const o of scene.objects.filter(o=>o.unitIndex===unit.unit)) {
              const v=unwrap(o.visual);
              if(v.kind!=='painted-shape'||v.stroke?.color!==255||v.geometry?.kind!=='path')continue;
              let previous;
              for(const c of v.geometry.commands) {
                if(c.kind==='moveTo')previous=[o.bounds.x+c.x,o.bounds.y+c.y];
                else if(c.kind==='lineTo'&&previous){const point=[o.bounds.x+c.x,o.bounds.y+c.y];lines.push([previous,point]);previous=point;}
              }
            }
            check(`${file} ${unit.unit}: native black segment count`,lines.length===unit.blackSegments.length);
            const matches=(a,b)=>a.every((p,i)=>Math.hypot(p[0]-b[i][0],p[1]-b[i][1])<.01);
            for(const [index,line] of unit.blackSegments.entries())
              check(`${file} ${unit.unit}: exact native black segment ${index}`,lines.some(actual=>matches(line,actual)||matches(line,[...actual].reverse())));
          }
          if (unit.gridSegments) {
            const lines=[];
            for(const o of scene.objects.filter(o=>o.unitIndex===unit.unit)){
              const v=unwrap(o.visual),color=v.stroke?.color;
              if(v.kind!=='painted-shape'||v.stroke?.kind!=='solid'||color===255||v.geometry?.kind!=='path')continue;
              const rgb=[color>>>24,(color>>>16)&255,(color>>>8)&255];
              if(Math.max(...rgb)!==Math.min(...rgb)||rgb[0]===255)continue;
              let p;
              for(const c of v.geometry.commands){
                const q=[o.bounds.x+c.x,o.bounds.y+c.y];
                if(c.kind==='moveTo')p=q;
                else if(c.kind==='lineTo'&&p){lines.push([p,q]);p=q;}
              }
            }
            check(`${file} ${unit.unit}: native gray grid segment count`,lines.length===unit.gridSegments.length);
            const matches=(a,b)=>a.every((p,i)=>Math.hypot(p[0]-b[i][0],p[1]-b[i][1])<.01);
            for(const [index,line] of unit.gridSegments.entries())
              check(`${file} ${unit.unit}: exact native gray grid ${index}`,lines.some(actual=>matches(line,actual)||matches(line,[...actual].reverse())));
          }
          if (unit.areaBoundaries) {
            const areas=scene.objects.filter(o=>o.unitIndex===unit.unit&&o.source?.kind==='chart-series'&&unwrap(o.visual).geometry?.kind==='path');
            const sample=points=>points.flatMap((p,i)=>i===0?[p]:[1/3,2/3,1].map(t=>points[i-1].map((v,k)=>v+(p[k]-v)*t)));
            for (const [index,boundary] of unit.areaBoundaries.entries())
              check(`${file} ${unit.unit}: native area boundary ${index}`,areas.some(o=>closeBoundaries(sample(pathBoundary(o)),sample(boundary))));
          }
          if (unit.nativeBarCount !== undefined) {
            const bars=scene.objects.filter(o=>o.unitIndex===unit.unit && o.source?.kind==='chart-series')
              .filter(o=>{const v=unwrap(o.visual);return v.kind==='painted-shape' && v.geometry==='rectangle' && v.fill.kind==='solid';});
            check(`${file} ${unit.unit}: native visible bar count`,bars.length===unit.nativeBarCount);
          }
          if (unit.bars) {
            const bars=scene.objects.filter(o=>o.unitIndex===unit.unit && o.source?.kind==='chart-series')
              .filter(o=>{const v=unwrap(o.visual);return v.kind==='painted-shape' && v.geometry==='rectangle' && v.fill.kind==='solid';});
            check(`Keynote hidden minimum ${unit.unit}: native bar count`, bars.length===unit.bars.length);
            for (const [index,[x,y,w,h]] of unit.bars.entries())
              check(`Keynote hidden minimum ${unit.unit}: native bar bounds ${index}`, bars.some(({bounds:b})=>
                [b.x-x,b.y-y,b.x+b.width-x-w,b.y+b.height-y-h].every(v=>Math.abs(v)<1.5)));
          }
        }
      }
      check(`${file}: chart count ${count}`, scene.objects.filter(o => o.source?.kind === 'chart').length === count);
      const census=matrix.fixtures.find(f=>f.file===file)?.nativeArchiveCensus;
      if (census) {
        const charts=scene.objects.filter(o=>o.source?.kind==='chart');
        check(`${file}: all 22 native static types occur once`, census.length===22 &&
          matrix.rows.filter(r=>r.scope==='static').every(r=>census.filter(c=>c.type===r.type).length===1));
        check(`${file}: native archive census maps to emitted chart parents`, charts.length===census.length &&
          census.every(c=>charts.filter(o=>o.source.component===c.component).length===1));
      }
      check(`${file}: native objects only`, !scene.objects.some(o => /preview/u.test(o.source?.kind ?? '')));
      if (file === 'iwork-static-charts.key') {
        const pie=(await json('tests/fixtures/iwork-chart-oracles/static-keynote/geometry.json')).threeDimensionalPie;
        const labels=scene.objects.filter(o=>o.unitIndex===pie.unit).flatMap(o=>unwrap(o.visual).runs?.map(r=>r.text)??[]);
        check('Keynote 3D pie: native percentage labels retained',pie.labels.every(text=>labels.includes(text)));
        const pieStyle=(await json('tests/fixtures/iwork-chart-oracles/static-keynote/geometry.json')).defaultPieLabels;
        const runs=scene.objects.filter(o=>o.unitIndex===pieStyle.unit).flatMap(o=>unwrap(o.visual).runs??[]);
        check('Keynote 2D pie: native themed percentage font and color',pieStyle.labels.every(text=>runs.some(r=>r.text===text && r.fontFamily===pieStyle.font && Math.abs(r.fontSize-pieStyle.fontSizeCss)<.001 && r.color===pieStyle.colorRgba)));


        check(`${file}: eight explicit 3D approximations`, scene.diagnostics.filter(d => d.message.startsWith('IWORK_CHART_3D_APPROXIMATE:')).length === 8);
        check('Keynote: native axes carry device stroke adjustment', scene.objects.some(o => o.unitIndex===0 && o.visual.kind==='stroke-style' && o.visual.style.strokeAdjust===true));
        const oracle = await json('tests/fixtures/iwork-chart-oracles/static-keynote/geometry.json');
        for (const unit of oracle.units) {
          const shapes = scene.objects.filter(o => o.unitIndex === unit.unit && o.source?.kind === 'chart-series').map(o => unwrap(o.visual));
          check(`Keynote static ${unit.unit}: native theme series colors`, unit.colors.every(c => shapes.some(v => [v.fill,v.stroke].some(p => p?.kind === 'solid' && [24,16,8].every((shift,k) => Math.abs(((p.color >>> shift)&255)/255-c[k])<.015)))));
          const labels=shapes.flatMap(v => v.runs?.map(r => r.text) ?? []);
          check(`Keynote static ${unit.unit}: native data selection and legend visibility`, !labels.includes('Alpha') && !labels.includes('Beta') && unit.labels.every(t=>labels.includes(t)));
        }

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
      if (file === 'numbers-dice-lab.numbers') {
        const oracle = await json('tests/fixtures/iwork-chart-oracles/existing-dice/geometry.json');
        const title = scene.objects.find(o => unwrap(o.visual).runs?.some(r => r.text === oracle.title));
        check('Numbers dice: native donut title center', title !== undefined && Math.abs(title.bounds.x + title.bounds.width/2 - oracle.titleCenterX)<1);
      }
      if (file === 'numbers-team-budget.numbers') {
        const oracle = await json('tests/fixtures/iwork-chart-oracles/existing-team/geometry.json');
        const labels = scene.objects.filter(o => o.unitIndex === oracle.unit && o.source?.kind === 'chart-series').flatMap(o => unwrap(o.visual).runs?.map(r => r.text) ?? []);
        check('Numbers team: native integer axis precision', oracle.axisLabels.every(t => labels.includes(t)));
      }
      if (file === 'numbers-retirement.numbers') {
        const oracle = await json('tests/fixtures/iwork-chart-oracles/existing-retirement/geometry.json');
        const areas = scene.objects.filter(o => o.source?.kind === 'chart-series' && o.unitIndex === oracle.unit && o.bounds.width>200 && o.bounds.height>100).map(o => unwrap(o.visual)).filter(v => v.kind === 'painted-shape' && v.fill.kind === 'solid');
        check('Numbers retirement: native opaque area fill', areas.length>0 && areas.every(v => (v.fill.color & 255)/255 === oracle.areaOpacity));
      }
      if (file === 'numbers-multiple-sheets-charts.numbers') {
        const oracle = await json('tests/fixtures/iwork-chart-oracles/existing-multiple/geometry.json');
        const title = scene.objects.find(o => unwrap(o.visual).runs?.some(r => r.text === oracle.title));
        check('Numbers: donut title is centered on native plot, independent of axis formatting', title !== undefined && Math.abs(title.bounds.x + title.bounds.width/2 - oracle.titleCenterX)<1);
      }
      if (['numbers-multiple-sheets-charts.numbers', 'chart-original.numbers'].includes(file)) {
        const host = file === 'chart-original.numbers' ? 'existing-original' : 'existing-multiple';
        const oracle = await json(`tests/fixtures/iwork-chart-oracles/${host}/geometry.json`);
        const shapes = scene.objects.filter(o=>o.unitIndex===0 && o.source?.kind==='chart-series' && unwrap(o.visual).kind==='painted-shape');
        for (const [i, [x,y,w,h]] of oracle.bars.entries()) {
          check(`${host}: native quantized bar ${i}`, shapes.some(({bounds:b})=>[b.x-x,b.y-y,b.width-w,b.height-h].every(v=>Math.abs(v)<.001)));
        }
      }
      if(['iwork-chart-radar-labels.key','iwork-chart-radar-value-axis.key'].includes(file)) {
        const oracle=await json(`tests/fixtures/iwork-chart-oracles/${labelOracleFolders[file]}/geometry.json`);
        for(const unit of oracle.units) {
          const axisLines=scene.objects.filter(o=>o.unitIndex===unit.unit&&o.source?.kind==='chart-series')
            .filter(o=>{const v=unwrap(o.visual);return v.kind==='painted-shape'&&v.stroke.kind==='solid'&&v.stroke.color!==255&&v.geometry.commands?.length===2;});
          check(`Keynote Radar value axis ${unit.unit}: native line count`,axisLines.length===unit.valueAxisLines.length);
          for(const [i,line] of unit.valueAxisLines.entries())check(`Keynote Radar value axis ${unit.unit}: native boundary ${i}`,
            axisLines.some(o=>closeBoundaries(pathBoundary(o),line.points)&&Math.abs(unwrap(o.visual).strokeWidth-line.width)<.01));
          const spokes=scene.objects.filter(o=>o.unitIndex===unit.unit&&o.source?.kind==='chart-series')
            .filter(o=>{const v=unwrap(o.visual);return v.kind==='painted-shape'&&v.fill.kind==='none'&&v.stroke.kind==='solid'&&v.stroke.color===255;});
          check(`Keynote Radar spokes ${unit.unit}: native line count`,spokes.length===unit.spokes.length);
          for(const [i,spoke] of unit.spokes.entries())check(`Keynote Radar spokes ${unit.unit}: native boundary ${i}`,
            spokes.some(o=>closeBoundaries(pathBoundary(o),spoke.points)&&Math.abs(unwrap(o.visual).strokeWidth-spoke.width)<.01));
        }
      }
      if(file==='3d-zero-area.key') {
        const faces=unit=>scene.objects.filter(o=>o.unitIndex===unit&&o.id.includes(':3d-face-'));
        check('Keynote zero 3D area: no degenerate plates',faces(0).length===0);
        for(const [unit,hue] of [[1,'blue'],[3,'green']])check(`Keynote zero stacked layer ${unit}: no ${hue} plates`,
          faces(unit).every(o=>{const color=unwrap(o.visual).fill.color,r=color>>>24,g=(color>>>16)&255,b=(color>>>8)&255;
            return hue==='blue'?!(b-r>40&&b-g>15):!(g-r>15&&g-b>40);}));
        const geometry=unit=>faces(unit).map(o=>JSON.stringify(pathBoundary(o))).sort();
        check('Keynote zero stacked layer: series order preserves visible geometry',
          JSON.stringify(geometry(1))===JSON.stringify(geometry(3)));
      }
      if(file==='iwork-chart-3d-series-gap.key') {
        const faces=unit=>scene.objects.filter(o=>o.unitIndex===unit && o.id.includes(':3d-face-'));
        for(const [unit,count] of [1,2,2,1,2,2,1].entries())
          check(`Keynote 3D series gap ${unit}: native solid contact`,solidComponents(faces(unit))===count);
        for(const [a,b] of [[1,2],[4,5]])
          check(`Keynote 3D series gap ${a}/${b}: separately authored gap geometry`,
            JSON.stringify(faces(a).map(pathBoundary))!==JSON.stringify(faces(b).map(pathBoundary)));
      }
      if(worldDiffuseFixtures.includes(file)) {
        const oracle=(await json(`tests/fixtures/iwork-chart-oracles/${file.slice(0,-4)}/geometry.json`)).units[1];
        const faces=scene.objects.filter(o=>o.unitIndex===1&&o.id.includes(':3d-face-'));check(`${file}: six visible native faces`,faces.length===6);
        const probes=[];
        for(const sample of oracle.faceColors)for(const index of {top:[0,1],front:[3,4],right:[2,5]}[sample.face]) {
          const face=faces[index],c=unwrap(face.visual).fill.color,rgb=[c>>>24,(c>>>16)&255,(c>>>8)&255];
          check(`${file}: native ${sample.face} ${index} world diffuse RGB`,rgb.every((v,k)=>Math.abs(v-sample.rgb[k])<=NATIVE_VISUAL_POLICY.pixelChannel));
          const points=pathBoundary(face);probes.push({face:sample.face,index,x:points.reduce((v,p)=>v+p[0],0)/points.length,y:points.reduce((v,p)=>v+p[1],0)/points.length,rgb:sample.rgb});
        }
        worldDiffuseProbes.set(file,probes);
      }
      if(file==='3d-six-implicit-lighting.key') {
        const comparison=core.open(await readFile('tests/fixtures/3d-six-default-lighting.key'));comparison.loadAll();
        try {
          const faces=(s,unit)=>s.objects.filter(o=>o.unitIndex===unit&&o.id.includes(':3d-face-')).map(o=>({bounds:o.bounds,visual:o.visual}));
          for(let unit=0;unit<6;unit++)check(`Native implicit Default type ${unit}: equal explicit Default faces`,JSON.stringify(faces(scene,unit))===JSON.stringify(faces(comparison.scene,unit+6)));
        } finally {comparison.close();}
      }
      if(file==='3d-six-default-lighting.key') {
        const oracle=await json('tests/fixtures/iwork-chart-oracles/3d-six-default-lighting/geometry.json');
        const faces=unit=>scene.objects.filter(o=>o.unitIndex===unit&&o.id.includes(':3d-face-')).map(o=>({bounds:o.bounds,visual:o.visual}));
        for(const pair of oracle.pairs)
          check(`Native ${pair.name}: ${pair.nativeIdentical?'preserve identical Default faces':'preserve authored Glossy versus Default'}`,(JSON.stringify(faces(pair.implicit))===JSON.stringify(faces(pair.explicit)))===pair.nativeIdentical);
      }
      if(file==='3d-center-intensity-four.key') {
        const oracle=(await json('tests/fixtures/iwork-chart-oracles/3d-center-intensity-four/geometry.json')).units[1];
        const expected=oracle.faceColors.find(s=>s.face==='right').rgb;
        const caps=scene.objects.filter(o=>o.unitIndex===1&&/\:3d-face-(2|5)$/.test(o.id));
        check('Native fourfold point intensity: two end caps remain visible',caps.length===2);
        for(const [i,o] of caps.entries()) {
          const c=unwrap(o.visual).fill.color,rgb=[c>>>24,(c>>>16)&255,(c>>>8)&255];
          check(`Native fourfold point intensity: cap ${i} diffuse saturation`,rgb.every((v,k)=>Math.abs(v-expected[k])<=NATIVE_VISUAL_POLICY.pixelChannel));
        }
      }
      if(file==='3d-light-off-unique.key') {
        const faces=scene.objects.filter(o=>o.unitIndex===1&&o.id.includes(':3d-face-'));
        const colors=faces.map(o=>unwrap(o.visual).fill?.color).filter(c=>c!==undefined)
          .map(c=>[c>>>24,(c>>>16)&255,(c>>>8)&255]);
        check('Native disabled lights: six visible opaque faces remain',faces.length===6);
        check('Native disabled lights: faces stay dark and achromatic',colors.length===6&&colors.every(c=>Math.max(...c)<40&&Math.max(...c)===Math.min(...c)));
        const oracle=(await json('tests/fixtures/iwork-chart-oracles/3d-light-off-unique/geometry.json')).units[1];
        for(const sample of oracle.faceColors)
          check(`Native disabled lights: ${sample.face} ambient color`,colors.length===6&&colors.every(c=>c.every((v,i)=>Math.abs(v-sample.rgb[i])<=NATIVE_VISUAL_POLICY.pixelChannel)));
      }
      if(defaultBarLightingFixtures.includes(file)) {
        const faces=unit=>scene.objects.filter(o=>o.unitIndex===unit&&o.id.includes(':3d-face-'))
          .map(o=>({bounds:o.bounds,visual:o.visual}));
        check(`${file}: native implicit and explicit Default have equal faces`,JSON.stringify(faces(0))===JSON.stringify(faces(1)));
        for(const unit of [0,1]) {
          const colors=faces(unit).map(o=>unwrap(o.visual).fill?.color).filter(c=>c!==undefined)
            .map(c=>[c>>>24,(c>>>16)&255,(c>>>8)&255]);
          check(`${file} ${unit}: native dark and chromatic faces`,colors.some(c=>Math.max(...c)-Math.min(...c)<40)&&colors.some(c=>Math.max(...c)-Math.min(...c)>40));
          const oracle=(await json(`tests/fixtures/iwork-chart-oracles/${file.slice(0,-4)}/geometry.json`)).units[unit];
          for(const sample of oracle.faceColors)
            check(`${file} ${unit}: native ${sample.face} diffuse color`,colors.some(c=>c.every((v,i)=>Math.abs(v-sample.rgb[i])<=NATIVE_VISUAL_POLICY.pixelChannel)));
        }
      }
      if(file==='iwork-chart-3d-clipped.key') {
        const oracle=await json('tests/fixtures/iwork-chart-oracles/3d-clipped/geometry.json');
        for(const unit of oracle.units) {
          const faces=scene.objects.filter(o=>o.unitIndex===unit.unit && o.id.includes(':3d-face-'));
          check(`Keynote clipped 3D ${unit.unit}: native independent bars`,solidComponents(faces)===unit.coloredComponents);
        }
      }
      if (file === 'iwork-chart-radar-grid.key') {
        const oracle=await json('tests/fixtures/iwork-chart-oracles/radar-grid/geometry.json');
        for(const unit of oracle.units) {
          const grids=scene.objects.filter(o=>o.unitIndex===unit.unit && o.source?.kind==='chart-series')
            .filter(o=>{const v=unwrap(o.visual);return v.kind==='painted-shape' && v.fill.kind==='none' && v.stroke.kind==='solid' && (v.geometry==='ellipse'||(v.geometry.commands?.length??0)>2);});
          check(`Keynote radar grid ${unit.unit}: native grid count`,grids.length===unit.grids.length);
          for(const [i,expected] of unit.grids.entries()) {
            check(`Keynote radar grid ${unit.unit}: native boundary ${i}`,grids.some(o=>{
              const v=unwrap(o.visual),g=v.geometry,b=o.bounds;
              const actual=g==='ellipse'?Array.from({length:129},(_,i)=>[b.x+b.width/2+b.width/2*Math.cos(i/128*2*Math.PI),b.y+b.height/2+b.height/2*Math.sin(i/128*2*Math.PI)]):pathBoundary(o);
              return closeBoundaries(actual,expected.boundary) && Math.abs(v.strokeWidth-expected.width)<.01;
            }));
          }
        }
      }
      if (file === 'iwork-chart-boundaries.key') {
        const oracle = await json('tests/fixtures/iwork-chart-oracles/boundaries/geometry.json');
        for (const unit of oracle.units) {
          const children = scene.objects.filter(o => o.unitIndex === unit.unit && o.source?.kind === 'chart-series');
          const labels = children.flatMap(o => unwrap(o.visual).runs?.map(r => r.text) ?? []);
          check(`Keynote boundary ${unit.unit}: native ticks and categories`, unit.labels.every(t => labels.includes(t)));
          const shapes = children.map(o => ({ bounds: o.bounds, visual: unwrap(o.visual) })).filter(o => o.visual.kind === 'painted-shape');
          const points = shapes.flatMap(o => o.visual.geometry.commands?.filter(p => ['moveTo', 'lineTo'].includes(p.kind)).map(p => [p.x + o.bounds.x, p.y + o.bounds.y]) ?? []);
          check(`Keynote boundary ${unit.unit}: finite paths`, shapes.every(o => (o.visual.geometry.commands ?? []).filter(p => ['moveTo', 'lineTo'].includes(p.kind)).every(p => Number.isFinite(p.x) && Number.isFinite(p.y))));
          for (const [i, drawing] of unit.drawings.entries()) {
            const expectedPoints = drawing.points;
            if (expectedPoints.length) {
              const matching = shapes.filter(o => {
                const paint = drawing.filled ? o.visual.fill : o.visual.stroke;
                if (paint.kind !== 'solid') return false;
                const rgb = [24, 16, 8].map(shift => ((paint.color >>> shift) & 255) / 255);
                return rgb.every((v, k) => Math.abs(v - drawing.color[k]) < .04);
              });
              check(`Keynote boundary ${unit.unit}: native path ${i}`, matching.some(o => {
                const actual = o.visual.geometry.commands?.filter(p => ['moveTo', 'lineTo'].includes(p.kind)).map(p => [p.x + o.bounds.x, p.y + o.bounds.y]) ?? [];
                let previous, start;
                const mids=[];
                for (const command of o.visual.geometry.commands ?? []) {
                  const p = command.kind === 'closePath' ? start : [command.x+o.bounds.x,command.y+o.bounds.y];
                  if (command.kind === 'moveTo') start=p;
                  else if (previous && p) mids.push([(previous[0]+p[0])/2,(previous[1]+p[1])/2]);
                  previous=p;
                }
                const edges=expectedPoints.filter((_,i) => i%2===0).map((a,i) => [a,expectedPoints[i*2+1]]);
                if (drawing.filled) edges.push([expectedPoints.at(-1),expectedPoints[0]]);
                return expectedPoints.every(p => actual.some(a => Math.hypot(a[0]-p[0], a[1]-p[1]) < 1.5))
                  && mids.every(p => edges.some(([a,b]) => distanceToSegment(p,a,b)<1.5))
                  && (!drawing.filled || Math.abs((o.visual.fill.color & 255)/255-drawing.opacity)<.01);
              }));
            } else {
              check(`Keynote boundary ${unit.unit}: native rectangle ${i}`, shapes.some(o => {
                const [x,y,w,h] = drawing.bounds; const b=o.bounds;
                return [b.x-x,b.y-y,b.x+b.width-x-w,b.y+b.height-y-h].every(v => Math.abs(v)<1.5);
              }));
            }
          }
        }
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
      let explicitDefaultCaptures;
      page.on('pageerror', e => errors.push(e.message));
      page.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
      if (baseline) {
        const bytes = await readFile(baseline);
        await page.route('**/dist/office-viewer-pdf.wasm', r => r.fulfill({ contentType: 'application/wasm', body: bytes }));
      }
      if (rendererBaseline) {
        const bytes = await readFile(rendererBaseline);
        await page.route('**/dist/render.js', r => r.fulfill({contentType:'text/javascript',body:bytes}));
      }
      await page.goto(`${url}examples/viewer.html?filePicker=true`);
      await page.waitForFunction(() => document.querySelector('#viewer').state.status === 'ready');
      for (const [file, count] of fixtures) {
        await page.locator('#viewer').evaluate(async (viewer, file) => {
          await viewer.open(await (await fetch(`/tests/fixtures/${file}`)).arrayBuffer());
        }, file);
        const nativeFrames = Object.fromEntries([...(nativePngs.get(file) ?? [])].map(([unit,image])=>[unit,
          nativeViewports.get(file)?.get(unit) ?? {x:0,y:0,width:image.width,height:image.height,scale:1}]));
        if(worldDiffuseFixtures.includes(file)||defaultBarLightingFixtures.includes(file)||file==='3d-light-off-unique.key'||file==='3d-center-intensity-four.key'||file==='3d-six-default-lighting.key'||file==='3d-six-implicit-lighting.key') {
          const r=await json(`tests/fixtures/iwork-chart-oracles/${file.slice(0,-4)}/reference.json`);
          for(const [unit,p] of r.pages.entries())nativeFrames[unit]={x:0,y:0,width:p.width,height:p.height,scale:1};
        }
        if(['axis-tick-boundaries.key','axis-tick-horizontal.key','axis-tick-logarithmic.key','axis-tick-logarithmic-ranges.key'].includes(file)) {
          const oracle=await json(`tests/fixtures/iwork-chart-oracles/${labelOracleFolders[file]}/geometry.json`);
          for(const unit of oracle.units)nativeFrames[unit.unit]=unit.nativeFrame;
        }
        const result = await page.evaluate(async ({file, nativeFrames}) => {
          const { createOfficeEngine } = await import('/dist/engine.js');
          const { extendedFormatPack } = await import('/dist/extended-formats.js');
          const engine = await createOfficeEngine({ execution: 'inline', formatPack: async () => extendedFormatPack });
          const doc = await engine.open(await (await fetch(`/tests/fixtures/${file}`)).arrayBuffer());
          try {
            const objects = await doc.listObjects();
            const charts = objects.filter(o => o.source.kind === 'chart');
            const captures = [];
            for (const chart of charts) {
              const b = chart.bounds;
              const nativeFrame = nativeFrames[chart.unitIndex];
              const viewport = nativeFrame ? {x:nativeFrame.x,y:nativeFrame.y,width:nativeFrame.width,height:nativeFrame.height} : { x: Math.max(0, b.x - 60), y: Math.max(0, b.y - 40), width: b.width + 150, height: b.height + 120 };
              let frame;
              try {
                frame = await doc.render({ unitIndex: chart.unitIndex, viewport, scale: nativeFrame?.scale ?? 1.5, pixelRatio: 1, includeTextFragments: true });
              } catch (error) {
                captures.push({unit:chart.unitIndex,viewport,error:{code:error.code,message:error.message}});
                continue;
              }
              try {
                const canvas = document.createElement('canvas'); canvas.width = frame.pixelWidth; canvas.height = frame.pixelHeight;
                const ctx = canvas.getContext('2d'); ctx.drawImage(frame.bitmap, 0, 0);
                const pixels = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
                let coloredPixels = 0;
                for (let i = 0; i < pixels.length; i += 4) if (pixels[i + 3] > 100 && Math.max(pixels[i], pixels[i + 1], pixels[i + 2]) - Math.min(pixels[i], pixels[i + 1], pixels[i + 2]) > 40) coloredPixels++;
                const objectLabels = new Map();
                for (const fragment of frame.textFragments)
                  objectLabels.set(fragment.objectId,(objectLabels.get(fragment.objectId)??'')+fragment.text);
                captures.push({ unit: chart.unitIndex, image: canvas.toDataURL(), viewport, coloredPixels,
                  labels: frame.textFragments.map(t => t.text).join(''), objectLabels:[...objectLabels.values()] });
              } finally { frame.bitmap.close(); }
            }
            return { charts, captures };
          } finally { doc.close(); engine.close(); }
        }, {file, nativeFrames});
        check(`${name} ${file}: public API chart count`, result.charts.length === count);
        if(file==='3d-six-default-lighting.key')explicitDefaultCaptures=result.captures.map(c=>c.image);
        if(file==='3d-six-implicit-lighting.key') {
          const folder='tests/fixtures/iwork-chart-oracles/3d-six-implicit-lighting',oracle=await json(`${folder}/geometry.json`),native=await json(`${folder}/reference.json`),otherNative=await json('tests/fixtures/iwork-chart-oracles/3d-six-default-lighting/reference.json');
          assert.equal(oracle.comparisonFixture,'3d-six-default-lighting.key');const comparison=explicitDefaultCaptures;assert.ok(comparison,'explicit Default control captures required');
          for(const pair of oracle.pairs) {
            assert.equal(native.pages[pair.implicit].sha256,otherNative.pages[pair.explicit].sha256,'native missing/explicit Default must be identical');
            check(`${name} implicit Default ${pair.name}: preserve exact explicit Default raster`,result.captures[pair.implicit]?.image===comparison[pair.explicit]);
          }
        }
        if(file==='3d-six-default-lighting.key') {
          const folder='tests/fixtures/iwork-chart-oracles/3d-six-default-lighting',r=await json(`${folder}/reference.json`),oracle=await json(`${folder}/geometry.json`);
          for(const pair of oracle.pairs) {
            assert.equal(r.pages[pair.implicit].sha256===r.pages[pair.explicit].sha256,pair.nativeIdentical);
            check(`${name} native ${pair.name}: preserve Default raster equivalence boundary`,(result.captures[pair.implicit]?.image===result.captures[pair.explicit]?.image)===pair.nativeIdentical);
          }
        }
        if(defaultBarLightingFixtures.includes(file)) {
          const r=await json(`tests/fixtures/iwork-chart-oracles/${file.slice(0,-4)}/reference.json`);
          assert.equal(r.pages[0].sha256,r.pages[1].sha256,'Native Default captures must be byte-identical');
          check(`${name} ${file}: native Default roundtrip preserves exact raster`,result.captures[0]?.image===result.captures[1]?.image);
        }
        for (const capture of result.captures) {
          check(`${name} ${file} unit ${capture.unit}: public render request`, capture.error===undefined);
          if (capture.error) continue;
          await writeFile(`${output}/${name}-${file}-${capture.unit}.png`, Buffer.from(capture.image.split(',')[1], 'base64'));
          if (file === 'iwork-static-charts.key' && capture.unit===13) {
            const pie=(await json('tests/fixtures/iwork-chart-oracles/static-keynote/geometry.json')).threeDimensionalPie;
            check(`${name} 3D pie: native percentage labels rendered`,pie.labels.every(text=>capture.labels.includes(text)));
          }
          if(file==='iwork-chart-radar-value-axis.key') {
            const folder='tests/fixtures/iwork-chart-oracles/radar-value-axis';
            const oracle=await json(`${folder}/geometry.json`),reference=await json(`${folder}/reference.json`);
            const unit=oracle.units[capture.unit],nativePage=reference.pages[capture.unit];
            const bytes=await readFile(`${folder}/page-${String(capture.unit+1).padStart(4,'0')}.png`);
            assert.equal(sha256(bytes),nativePage.sha256);
            const actual=decodePng(Buffer.from(capture.image.split(',')[1],'base64'),'Radar value axis browser');
            const expected=decodePng(bytes,'Radar value axis Apple');
            for(const [kind,lines] of [['value axis',unit.valueAxisLines],['radial lines',unit.spokes]]) {
              const accept=kind==='value axis'?rgb=>rgb.every(c=>Math.abs(c-136)<8):rgb=>rgb.every(c=>c<80);
              for(const [i,line] of lines.entries()) {
                const [a,b]=line.points,[x,y]=[0,1].map(k=>a[k]+(b[k]-a[k])*.8);
                assert.equal(rasterRgbAt(expected,x,y,2,accept),true,`Apple Radar ${capture.unit}/${kind}/${i}`);
                check(`${name} Radar ${capture.unit}: native ${kind} raster probe ${i}`,
                  rasterRgbAt(actual,(x-capture.viewport.x)*1.5,(y-capture.viewport.y)*1.5,3,accept));
              }
            }
          }
          if(file==='iwork-chart-radar-labels.key') {
            const folder='tests/fixtures/iwork-chart-oracles/radar-labels';
            const oracle=await json(`${folder}/geometry.json`),reference=await json(`${folder}/reference.json`);
            const unit=oracle.units[capture.unit],nativePage=reference.pages[capture.unit];
            const bytes=await readFile(`${folder}/page-${String(capture.unit+1).padStart(4,'0')}.png`);
            assert.equal(sha256(bytes),nativePage.sha256);
            const actual=decodePng(Buffer.from(capture.image.split(',')[1],'base64'),'Radar browser');
            const expected=decodePng(bytes,'Radar Apple');
            const hasBlack=(image,x,y,radius)=>rasterRgbAt(image,x,y,radius,rgb=>rgb.every(c=>c<80));
            for(const [i,spoke] of (unit.spokes.length?unit.spokes:oracle.units[1].spokes).entries()) {
              const [a,b]=spoke.points,[x,y]=[0,1].map(k=>a[k]+(b[k]-a[k])*.8);
              const nativeBlack=hasBlack(expected,x,y,2);
              assert.equal(nativeBlack,unit.spokes.length>0,`Apple Radar spoke ${capture.unit}/${i}`);
              check(`${name} Radar spokes ${capture.unit}: native raster probe ${i}`,
                hasBlack(actual,(x-capture.viewport.x)*1.5,(y-capture.viewport.y)*1.5,3)===nativeBlack);
            }
          }
          if(['axis-tick-boundaries.key','axis-tick-horizontal.key','axis-tick-logarithmic.key','axis-tick-logarithmic-ranges.key'].includes(file)) {
            const folder=`tests/fixtures/iwork-chart-oracles/${labelOracleFolders[file]}`;
            const unit=(await json(`${folder}/geometry.json`)).units[capture.unit];
            const bytes=await readFile(`${folder}/page-${String(capture.unit+1).padStart(4,'0')}.png`);
            assert.equal(sha256(bytes),unit.nativePageSha256);
            const expected=decodePng(bytes,'Apple axis ticks'),actual=decodePng(Buffer.from(capture.image.split(',')[1],'base64'),'axis ticks browser');
            assert.deepEqual([actual.width,actual.height],[expected.width,expected.height]);
            for(const [index,line] of unit.tickSegments.entries()) {
              const left=Math.max(0,Math.floor(Math.min(...line.map(p=>p[0])))-2),right=Math.min(expected.width,Math.ceil(Math.max(...line.map(p=>p[0])))+2);
              const top=Math.max(0,Math.floor(Math.min(...line.map(p=>p[1])))-2),bottom=Math.min(expected.height,Math.ceil(Math.max(...line.map(p=>p[1])))+2);
              let equal=true;
              for(let y=top;y<bottom;y++)for(let x=left;x<right;x++) {
                const black=image=>image.data.subarray((y*image.width+x)*4,(y*image.width+x)*4+3).every(c=>c<80);
                if(black(expected)!==black(actual))equal=false;
              }
              check(`${name} axis ticks ${capture.unit}: exact native black-ink patch ${index}`,equal);
            }
            for(const [index,line] of (unit.gridSegments??[]).entries()) {
              const [a,b]=line,p=a.map((v,k)=>v+(b[k]-v)*.01);
              const left=Math.floor(p[0])-2,right=Math.ceil(p[0])+2,top=Math.floor(p[1])-2,bottom=Math.ceil(p[1])+2;
              let equal=true,nativeInk=0;
              const gray=(image,x,y)=>{
                const rgb=image.data.subarray((y*image.width+x)*4,(y*image.width+x)*4+3);
                return Math.max(...rgb)-Math.min(...rgb)<3&&Math.min(...rgb)>80&&Math.max(...rgb)<230;
              };
              for(let y=top;y<bottom;y++)for(let x=left;x<right;x++) {
                nativeInk+=Number(gray(expected,x,y));
                if(gray(expected,x,y)!==gray(actual,x,y))equal=false;
              }
              assert.ok(nativeInk>0,`Native gray grid probe ${capture.unit}/${index} must contain ink`);
              check(`${name} axis ticks ${capture.unit}: exact native gray grid patch ${index}`,equal);
            }
          }
          if(file==='stack-sparse-negative.key'&&capture.unit>=2) {
            const folder='tests/fixtures/iwork-chart-oracles/stack-sparse-negative';
            const unit=(await json(`${folder}/geometry.json`)).units[capture.unit];
            const bytes=await readFile(`${folder}/page-${String(capture.unit+1).padStart(4,'0')}.png`);
            assert.equal(sha256(bytes),unit.nativePageSha256);
            const native=chart3dRasterFeatures(decodePng(bytes,'Apple sparse negative stack'),[.55,.75]);
            const actual=chart3dRasterFeatures(decodePng(Buffer.from(capture.image.split(',')[1],'base64'),'sparse negative stack browser'),[.55,.75]);
            for (const feature of ['greenPresent','blueValleyNearZero']) {
              assert.equal(native[feature],true,`Native sparse negative stack ${capture.unit}: ${feature}`);
              check(`${name} sparse negative stack ${capture.unit}: native ${feature}`,actual[feature]===native[feature]);
            }
          }
          if(file==='3d-signed-area.key') {
            const folder='tests/fixtures/iwork-chart-oracles/3d-signed-area';
            const unit=(await json(`${folder}/geometry.json`)).units[capture.unit];
            const bytes=await readFile(`${folder}/page-${String(capture.unit+1).padStart(4,'0')}.png`);
            assert.equal(sha256(bytes),unit.nativePageSha256);
            const native=chart3dRasterFeatures(decodePng(bytes,'Apple signed area'));
            const actual=chart3dRasterFeatures(decodePng(Buffer.from(capture.image.split(',')[1],'base64'),'signed area browser'));
            assert.ok(Math.abs(native.blueBodyCoverage-unit.blueBodyCoverage)<1e-12);
            assert.equal(native.greenPresent,unit.greenPresent);
            check(`${name} signed 3D area ${capture.unit}: native front body retained`,actual.blueBodyCoverage>=unit.blueBodyCoverage*.9);
            check(`${name} signed 3D area ${capture.unit}: native green presence`,actual.greenPresent===unit.greenPresent);
          }
          if(file==='3d-flat-area.key') {
            const folder='tests/fixtures/iwork-chart-oracles/3d-flat-area';
            const unit=(await json(`${folder}/geometry.json`)).units[capture.unit];
            const bytes=await readFile(`${folder}/page-${String(capture.unit+1).padStart(4,'0')}.png`);
            assert.equal(sha256(bytes),unit.nativePageSha256);
            const native=decodePng(bytes,'Apple flat area'),actual=decodePng(Buffer.from(capture.image.split(',')[1],'base64'),'flat area browser');
            assert.equal(rasterComponents(native),unit.coloredComponents);
            check(`${name} flat 3D ${capture.unit}: native independent solid`,rasterComponents(actual)===unit.coloredComponents);
            if(Object.hasOwn(unit,'convexSilhouette')) {
              assert.ok(Math.abs(coloredConvexCoverage(native)-unit.convexCoverage)<1e-12);
              assert.equal(coloredConvexCoverage(native)>=.995,unit.convexSilhouette);
              check(`${name} flat 3D area: native convex silhouette`,(coloredConvexCoverage(actual)>=.995)===unit.convexSilhouette);
            }
          }
          if(file==='3d-zero-area.key') {
            const folder='tests/fixtures/iwork-chart-oracles/3d-zero-area';
            const unit=(await json(`${folder}/geometry.json`)).units[capture.unit];
            const bytes=await readFile(`${folder}/page-${String(capture.unit+1).padStart(4,'0')}.png`);
            assert.equal(sha256(bytes),unit.nativePageSha256);
            const native=decodePng(bytes,'Apple zero area'),actual=decodePng(Buffer.from(capture.image.split(',')[1],'base64'),'zero area browser');
            assert.equal(rasterComponents(native),unit.coloredComponents);
            assert.equal(chart3dRasterFeatures(native).bluePresent,unit.bluePresent);
            assert.equal(chart3dRasterFeatures(native).greenPresent,unit.greenPresent);
            check(`${name} zero 3D area ${capture.unit}: native colored regions`,rasterComponents(actual)===unit.coloredComponents);
            check(`${name} zero 3D area ${capture.unit}: native blue presence`,chart3dRasterFeatures(actual).bluePresent===unit.bluePresent);
            check(`${name} zero 3D area ${capture.unit}: native green presence`,chart3dRasterFeatures(actual).greenPresent===unit.greenPresent);
            if(capture.unit===0)check(`${name} zero 3D area: empty native volume`,capture.coloredPixels===unit.coloredPixels);
          }
          if(file==='iwork-chart-3d-sparse-area.key') {
            const folder='tests/fixtures/iwork-chart-oracles/3d-sparse-area';
            const unit=(await json(`${folder}/geometry.json`)).units[capture.unit];
            const bytes=await readFile(`${folder}/page-${String(capture.unit+1).padStart(4,'0')}.png`);
            assert.equal(sha256(bytes),unit.nativePageSha256);
            const native=chart3dRasterFeatures(decodePng(bytes,'Apple sparse 3D'));
            const actual=chart3dRasterFeatures(decodePng(Buffer.from(capture.image.split(',')[1],'base64'),'sparse 3D browser'));
            for(const property of ['redPresent','blueValleyNearZero','redStartsAtFirstCategory','redBandHasArea','redTopIsStraight'])if(Object.hasOwn(unit,property)) {
              assert.equal(native[property],unit[property],`Apple sparse 3D ${capture.unit}: ${property}`);
              check(`${name} sparse 3D ${capture.unit}: native ${property}`,actual[property]===unit[property]);
            }
          }
          if(file==='iwork-chart-3d-series-gap.key') {
            const folder='tests/fixtures/iwork-chart-oracles/3d-series-gap';
            const unit=(await json(`${folder}/geometry.json`)).units[capture.unit];
            const bytes=await readFile(`${folder}/page-${String(capture.unit+1).padStart(4,'0')}.png`);
            assert.equal(sha256(bytes),unit.nativePageSha256);
            assert.equal(rasterComponents(decodePng(bytes,'Apple series gap'),unit.relativeChromaticXCrop),unit.coloredComponents);
            check(`${name} 3D series gap ${capture.unit}: native visible contact`,
              rasterComponents(decodePng(Buffer.from(capture.image.split(',')[1],'base64'),'series gap browser'),unit.relativeChromaticXCrop)===unit.coloredComponents);
          }
          if(worldDiffuseFixtures.includes(file)&&capture.unit===1) {
            const raster=decodePng(Buffer.from(capture.image.split(',')[1],'base64'),'world diffuse surface interiors');
            for(const probe of worldDiffuseProbes.get(file)) {
              const rgb=[[],[],[]];for(let y=Math.floor(probe.y)-1;y<=Math.floor(probe.y)+1;y++)for(let x=Math.floor(probe.x)-1;x<=Math.floor(probe.x)+1;x++)for(let k=0;k<3;k++)rgb[k].push(raster.data[(y*raster.width+x)*4+k]);
              check(`${name} ${file}: native ${probe.face} ${probe.index} diffuse raster`,rgb.every((values,k)=>Math.abs(values.sort((a,b)=>a-b)[4]-probe.rgb[k])<=NATIVE_VISUAL_POLICY.pixelChannel));
            }
          }
          if(file==='3d-center-intensity-four.key'&&capture.unit===1) {
            const oracle=(await json('tests/fixtures/iwork-chart-oracles/3d-center-intensity-four/geometry.json')).units[1];
            const expected=oracle.faceColors.find(s=>s.face==='right').rgb;
            const raster=decodePng(Buffer.from(capture.image.split(',')[1],'base64'),'fourfold native diffuse cap');let matching=0;
            for(let p=0;p<raster.data.length;p+=4)if(expected.every((v,k)=>Math.abs(v-raster.data[p+k])<=NATIVE_VISUAL_POLICY.pixelChannel))matching++;
            check(`${name} fourfold point intensity: native saturated diffuse color rasterized`,matching>=16);
          }
          if(defaultBarLightingFixtures.includes(file)) {
            const folder=`tests/fixtures/iwork-chart-oracles/${file.slice(0,-4)}`;
            const unit=(await json(`${folder}/geometry.json`)).units[capture.unit];
            const bytes=await readFile(`${folder}/page-${String(capture.unit+1).padStart(4,'0')}.png`);
            assert.equal(sha256(bytes),unit.nativePageSha256);
            assert.equal(rasterComponents(decodePng(bytes,'Native Default lighting')),unit.coloredComponents);
            check(`${name} ${file} ${capture.unit}: native independently colored bars`,rasterComponents(decodePng(Buffer.from(capture.image.split(',')[1],'base64'),'Default lighting browser'))===unit.coloredComponents);
            const raster=decodePng(Buffer.from(capture.image.split(',')[1],'base64'),'Default diffuse colors');
            for(const sample of unit.faceColors) {
              let matching=0;
              for(let p=0;p<raster.data.length;p+=4)
                if(sample.rgb.every((v,i)=>Math.abs(v-raster.data[p+i])<=NATIVE_VISUAL_POLICY.pixelChannel))matching++;
              check(`${name} ${file} ${capture.unit}: native ${sample.face} diffuse color rasterized`,matching>=16);
            }
          }
          if(file==='iwork-chart-3d-clipped.key') {
            const folder='tests/fixtures/iwork-chart-oracles/3d-clipped';
            const oracle=await json(`${folder}/geometry.json`),unit=oracle.units[capture.unit];
            const bytes=await readFile(`${folder}/page-${String(capture.unit+1).padStart(4,'0')}.png`);
            assert.equal(sha256(bytes),unit.nativePageSha256);
            assert.equal(rasterComponents(decodePng(bytes,'Apple clipped 3D')),unit.coloredComponents);
            check(`${name} clipped 3D ${capture.unit}: native independent bars rasterized`,
              rasterComponents(decodePng(Buffer.from(capture.image.split(',')[1],'base64'),'clipped 3D browser'))===unit.coloredComponents);
          }
          if(file==='iwork-chart-radar-grid.key') {
            const folder='tests/fixtures/iwork-chart-oracles/radar-grid';
            const oracle=await json(`${folder}/geometry.json`),unit=oracle.units[capture.unit];
            const reference=await json(`${folder}/reference.json`),nativePage=reference.pages[capture.unit];
            const bytes=await readFile(`${folder}/page-${String(capture.unit+1).padStart(4,'0')}.png`);
            assert.equal(sha256(bytes),nativePage.sha256);
            const actual=decodePng(Buffer.from(capture.image.split(',')[1],'base64'),'Radar browser');
            const expected=decodePng(bytes,'Radar Apple');
            const hasGray=(image,x,y,radius)=>rasterRgbAt(image,x,y,radius,
              rgb=>Math.max(...rgb)-Math.min(...rgb)<8 && rgb.every(c=>c>=100 && c<=200));
            for(const [i,[x,y]] of unit.gridProbes.entries()) {
              const nativeGray=hasGray(expected,x,y,2);
              assert.equal(nativeGray,capture.unit!==2,`Apple Radar probe ${capture.unit}/${i}`);
              check(`${name} Radar grid ${capture.unit}: native raster probe ${i}`,
                hasGray(actual,(x-capture.viewport.x)*1.5,(y-capture.viewport.y)*1.5,3)===nativeGray);
            }
          }
          if (Object.hasOwn(labelOracleFolders,file)) {
            const oracle = await json(`tests/fixtures/iwork-chart-oracles/${labelOracleFolders[file]}/geometry.json`);
            const labels = capture.objectLabels.filter(text=>!(oracle.units[capture.unit].categoryLabels??[]).includes(text)&&!(['iwork-chart-radar-labels.key','iwork-chart-radar-value-axis.key'].includes(file)?['Manual Axis 5 to 50']:['Currency','Scientific','Prefix suffix','Fractions','Manual Axis 5 to 50','A','B','C']).includes(text));
            check(`${name} ${file} ${capture.unit}: exact native labels through public renderer`, JSON.stringify(labels.sort())===JSON.stringify([...oracle.units[capture.unit].labels].sort()));
          }
          if (nativePngs.has(file)) {
            const actual = decodePng(Buffer.from(capture.image.split(',')[1], 'base64'), `${name} static slide ${capture.unit}`);
            const expected = nativePngs.get(file).get(capture.unit);
            const unitIndex = capture.unit;
            const units = [{index:unitIndex,type:file.endsWith('.pages')?'page':'slide',width:expected.width,height:expected.height}];
            const objects = nativeVisualRegions(unitIndex,expected.width,expected.height);
            const report = evaluateAccuracy({fileId:`${name}-${file}-${unitIndex}`,corpusClass:'minimal',
              expected:{units,objects,visuals:[{unitIndex,...expected}]},
              actual:{units,objects,visuals:[{unitIndex,...actual}]},policy:NATIVE_VISUAL_POLICY});
            if (file === 'iwork-static-charts.key' && unitIndex === 0) {
              for (const [label,rect] of [['grid',[200,130,600,6]],['axis',[131,200,6,400]],['bar coverage',[480,796,90,5]]]) {
                const [x,y,w,h]=rect;
                let same=true;
                for(let row=y;row<y+h;row++) for(let col=x;col<x+w;col++) {
                  const offset=(row*expected.width+col)*4;
                  for(let c=0;c<4;c++) if(Math.abs(expected.data[offset+c]-actual.data[offset+c])>8) same=false;
                }
                check(`${name} static slide 0: native ${label} raster boundaries`,same);
              }
            }
            nativeVisual.push({file,browser:name,unit:unitIndex,passed:report.passed,visualSimilarity:report.visualSimilarity,failureReasons:report.failureReasons});
            check(`${name} ${file} unit ${unitIndex}: full native visual policy`, report.passed);
          }
          if(file==='3d-light-off-unique.key'&&capture.unit===1) {
            const folder='tests/fixtures/iwork-chart-oracles/3d-light-off-unique';
            const oracle=(await json(`${folder}/geometry.json`)).units[1];
            const bytes=await readFile(`${folder}/page-0002.png`);assert.equal(sha256(bytes),oracle.nativePageSha256);
            const native=decodePng(bytes,'Native disabled lights');
            assert.equal(oracle.coloredPixels,0);assert.ok(rasterComponents(native,undefined,'dark')>0);
            check(`${name} disabled lights: dark volumes remain visible`,rasterComponents(decodePng(Buffer.from(capture.image.split(',')[1],'base64'),'disabled lighting browser'),undefined,'dark')>0);
          }
          delete capture.image;
          const nativeEmptyBars = ['auto-count-heldout.key','auto-count.key'].includes(file)
            && (await json(`tests/fixtures/iwork-chart-oracles/${labelOracleFolders[file]}/geometry.json`)).units[capture.unit].nativeBarCount===0;
          const nativeDarkBars = file==='3d-light-off-unique.key'&&capture.unit===1;
          const nativeEmptyArea = file==='3d-zero-area.key'
            && (await json('tests/fixtures/iwork-chart-oracles/3d-zero-area/geometry.json')).units[capture.unit].coloredPixels===0;
          if(nativeDarkBars) {
            check(`${name} disabled lights: no chromatic face remains`,capture.coloredPixels===0);
          } else if (nativeEmptyArea) {
            check(`${name} ${file} unit ${capture.unit}: native empty area has no colored volume`,capture.coloredPixels===0);
          } else if (nativeEmptyBars || (file === 'iwork-chart-auto-axis.key' && capture.unit === 15)) {
            check(`${name} ${file} unit ${capture.unit}: native empty range has no colored bars`, capture.coloredPixels === 0);
          } else if (!(worldDiffuseFixtures.includes(file)&&capture.unit===1)) {
            // These native signed-axis controls can legitimately be dark or
            // weakly chromatic; exact per-face RGB probes above cover their ink.
            check(`${name} ${file} unit ${capture.unit}: chart color rasterized`, capture.coloredPixels > 100);
          }
          if (file === native.fixture) {
            const labels = { 2: ['Nutty', 'Roasted', 'Spicy', 'Floral', 'Fruity'], 5: ['Jan', 'Dec', '90'], 7: ['Salespeople', 'Units Sold'] }[capture.unit];
            if (labels) check(`${name} Numbers unit ${capture.unit}: authored labels`, labels.every(label => capture.labels.includes(label)));
          }
        }
        await writeFile(`${output}/${name}-${file}.json`, `${JSON.stringify(result, null, 2)}\n`);
        if (nativeViewports.has(file)) {
          const viewport = nativeViewports.get(file).get(16);
          const worker = await page.evaluate(async ({file,viewport}) => {
            const {createOfficeEngine} = await import('/dist/engine.js');
            const {extendedFormatPack} = await import('/dist/extended-formats.js');
            const engine = await createOfficeEngine({execution:'worker',formatPack:async()=>extendedFormatPack});
            let doc;
            try {
              doc = await engine.open(await (await fetch(`/tests/fixtures/${file}`)).arrayBuffer());
              const {scale,sheetName,...rect} = viewport;
              const frame = await doc.render({unitIndex:16,viewport:rect,scale,pixelRatio:1,includeTextFragments:true});
              try {
                return {width:frame.pixelWidth,height:frame.pixelHeight,objects:frame.renderedObjectCount,labels:frame.textFragments.map(f=>f.text).join('')};
              } finally {frame.bitmap.close();}
            } catch(error) {return {error:error.message};}
            finally {doc?.close();engine.close();}
          }, {file,viewport});
          check(`${name} Numbers native signed viewport: worker transport`, worker.width===1800 && worker.height===960 && worker.objects>10 && worker.labels?.includes('Type 14'));
          await writeFile(`${output}/${name}-native-viewport-worker.json`, `${JSON.stringify(worker,null,2)}\n`);
        }
        if (file !== 'iwork-static-charts.key') await page.screenshot({ path: `${output}/${name}-${file}-viewer.png` });
      }
      check(`${name}: clean console`, errors.length === 0);
      if (errors.length) failures.push(...errors);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
const missingNativeVisualWitnesses = matrix.rows.filter(r=>r.scope==='static').flatMap(row=>
  Object.entries(row.hosts).filter(([,host])=>!host.witnesses.some(w=>w.units.some(unit=>
    ['chromium','webkit'].every(browser=>nativeVisual.some(v=>
      v.file===w.fixture && v.unit===unit && v.browser===browser && v.passed)))))
    .map(([host])=>`${row.name}: ${host}`));
await writeFile(`${output}/report.json`, `${JSON.stringify({ baseline: baseline ?? null, rendererBaseline:rendererBaseline??null, checks, failures, incomplete, incompleteFeatures, nativeVisual, missingNativeVisualWitnesses }, null, 2)}\n`);
console.log(`iWork charts: ${checks.filter(c => c.passed).length}/${checks.length} checks passed; ${output}/report.json`);
assert.deepEqual(failures, [], 'iWork real-file/native geometry/browser regression');

console.log(`Static completeness: ${incomplete.length} chart types and ${incompleteFeatures.length} feature groups remain unaccepted. See docs/research/IWORK-CHART-COVERAGE.md.`);
if (requireComplete) assert.deepEqual([...incomplete, ...incompleteFeatures, ...missingNativeVisualWitnesses], [], 'iWork static chart completeness requires fresh native visual witnesses for every type and host');
