import assert from 'node:assert/strict';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { Core, DEFAULT_LIMITS } from '../dist/core.js';
import { createZip } from '../tests/zip-fixture.mjs';
import { readZipEntries } from './accuracy-metamorphic.mjs';
import { chromium, firefox, webkit } from 'playwright-core';

// Native reference: PowerPoint 16.113.2, 96 dpi PDF export; source SHA-256
// 73b93656401b6781d84e58b88072179de94150947b46a3b05c8e50eaa6152bdb.
// All hosts carry the exact picture XML and GIF from the reported Microsoft corpus file.
const original = await readFile(new URL('../tests/fixtures/image-3d-rotation.pptx', import.meta.url));
const parts = Object.fromEntries(readZipEntries(original).map(({ name, data }) => [name, data]));
const slide = new TextDecoder().decode(parts['ppt/slides/slide1.xml']);
const picture = slide.match(/<p:pic>[\s\S]*?<\/p:pic>/u)[0];
const media = parts['ppt/media/image1.gif'];
assert.ok(media);
const ns = 'http://schemas.openxmlformats.org/';
const rels = target => `<Relationships xmlns="${ns}package/2006/relationships"><Relationship Id="rId3" Type="${ns}officeDocument/2006/relationships/image" Target="${target}"/></Relationships>`;
const contentTypes = (main, type) => `<Types xmlns="${ns}package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="gif" ContentType="image/gif"/><Override PartName="/${main}" ContentType="application/vnd.openxmlformats-officedocument.${type}.main+xml"/></Types>`;
const rootRel = target => `<Relationships xmlns="${ns}package/2006/relationships"><Relationship Id="rId1" Type="${ns}officeDocument/2006/relationships/officeDocument" Target="${target}"/></Relationships>`;
function packageFor(host, pic) {
  if (host === 'pptx') return createZip({ ...parts, 'ppt/slides/slide1.xml': slide.replace(picture, pic) });
  if (host === 'docx') return createZip({
    '[Content_Types].xml': contentTypes('word/document.xml', 'wordprocessingml.document'),
    '_rels/.rels': rootRel('word/document.xml'),
    'word/document.xml': `<w:document xmlns:w="${ns}wordprocessingml/2006/main" xmlns:wp="${ns}drawingml/2006/wordprocessingDrawing" xmlns:a="${ns}drawingml/2006/main" xmlns:pic="${ns}drawingml/2006/picture" xmlns:r="${ns}officeDocument/2006/relationships"><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="1676400" cy="2209800"/><wp:docPr id="7" name="Image_Small.gif"/><a:graphic><a:graphicData uri="${ns}drawingml/2006/picture">${pic.replaceAll('p:', 'pic:')}</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p><w:sectPr/></w:body></w:document>`,
    'word/_rels/document.xml.rels': rels('media/image1.gif'), 'word/media/image1.gif': media,
  });
  return createZip({
    '[Content_Types].xml': contentTypes('xl/workbook.xml', 'spreadsheetml.sheet').replace('</Types>', `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`),
    '_rels/.rels': rootRel('xl/workbook.xml'),
    'xl/workbook.xml': `<workbook xmlns="${ns}spreadsheetml/2006/main" xmlns:r="${ns}officeDocument/2006/relationships"><sheets><sheet name="Pictures" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${ns}package/2006/relationships"><Relationship Id="rId1" Type="${ns}officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': `<worksheet xmlns="${ns}spreadsheetml/2006/main" xmlns:r="${ns}officeDocument/2006/relationships"><sheetData/><drawing r:id="rId1"/></worksheet>`,
    'xl/worksheets/_rels/sheet1.xml.rels': `<Relationships xmlns="${ns}package/2006/relationships"><Relationship Id="rId1" Type="${ns}officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
    'xl/drawings/drawing1.xml': `<xdr:wsDr xmlns:xdr="${ns}drawingml/2006/spreadsheetDrawing" xmlns:a="${ns}drawingml/2006/main" xmlns:r="${ns}officeDocument/2006/relationships"><xdr:absoluteAnchor><xdr:pos x="3733800" y="2758281"/><xdr:ext cx="1676400" cy="2209800"/>${pic.replaceAll('p:', 'xdr:')}<xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>`,
    'xl/drawings/_rels/drawing1.xml.rels': rels('../media/image1.gif'), 'xl/media/image1.gif': media,
  });
}
const flat = picture.replace(/<a:scene3d>[\s\S]*?<\/a:scene3d>/u, '');
const crop = pic => pic.replace('<a:stretch>', '<a:srcRect l="20000" t="10000" r="10000" b="5000"/><a:stretch>');
const shape = (pic, preset) => pic.replace('prst="rect"', `prst="${preset}"`);
const custom = pic => pic.replace('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>', '<a:custGeom><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="50" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:lnTo><a:pt x="0" y="100"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>');
const cases = {
  original: picture, flat,
  flipH: flat.replace('<a:xfrm>', '<a:xfrm flipH="1">'),
  flipV: flat.replace('<a:xfrm>', '<a:xfrm flipV="1">'),
  rotate90: flat.replace('<a:xfrm>', '<a:xfrm rot="5400000">'),
  crop: crop(flat),
  cropPercent: crop(flat).replace('l="20000" t="10000" r="10000" b="5000"','l="20%" t="10%" r="10%" b="5%"'),
  cameraZoom: picture.replace('prst="perspectiveHeroicExtremeLeftFacing"','prst="perspectiveHeroicExtremeLeftFacing" zoom="75000"'),
  cameraFov: picture.replace('prst="perspectiveHeroicExtremeLeftFacing"','prst="perspectiveHeroicExtremeLeftFacing" fov="1800000"'),
  padding: flat.replace('<a:stretch>', '<a:srcRect l="-20000" t="-10000" r="-10000" b="-5000"/><a:stretch>'),
  ellipse: shape(flat, 'ellipse'), triangle: shape(flat, 'triangle'), custom: custom(flat),
  adjusted: shape(flat, 'roundRect').replace('<a:avLst/>', '<a:avLst><a:gd name="adj" fmla="val 40000"/></a:avLst>'),
  combined: crop(shape(picture, 'triangle')).replace('<a:xfrm>', '<a:xfrm rot="2100000" flipH="1">'),
  explicitFront: picture.replace('prst="perspectiveHeroicExtremeLeftFacing"/>', 'prst="perspectiveHeroicExtremeLeftFacing"><a:rot lat="0" lon="0" rev="0"/></a:camera>'),
};
// Exercise all 62 camera presets, including their zero-depth legacy/oblique planes.
const presets = [
  ...['BottomDown','BottomUp','LeftDown','LeftUp','RightDown','RightUp','TopDown','TopUp'].map(n=>'isometric'+n),
  ...[1,2,3,4].flatMap(n=>['Left','Right',n<3?'Top':'Bottom'].map(side=>`isometricOffAxis${n}${side}`)),
  ...['legacyOblique','legacyPerspective'].flatMap(prefix=>['TopLeft','Top','TopRight','Left','Front','Right','BottomLeft','Bottom','BottomRight'].map(n=>prefix+n)),
  ...['TopLeft','Top','TopRight','Left','Right','BottomLeft','Bottom','BottomRight'].map(n=>'oblique'+n),
  'orthographicFront',
  ...['Front','Left','Right','Above','Below','AboveLeftFacing','AboveRightFacing','ContrastingLeftFacing','ContrastingRightFacing','HeroicLeftFacing','HeroicRightFacing','HeroicExtremeLeftFacing','HeroicExtremeRightFacing','Relaxed','RelaxedModerately'].map(n=>'perspective'+n),
];
assert.equal(presets.length,62);
for (const preset of presets) cases[preset] = picture.replace('perspectiveHeroicExtremeLeftFacing',preset);
const scenes = [];
const core = await Core.create(await readFile(process.env.PICTURE_WASM ?? new URL('../dist/office-viewer-core.wasm', import.meta.url)), DEFAULT_LIMITS);
const failures = [];
try {
  for (const [name, pic] of Object.entries(cases)) for (const host of ['pptx', 'docx', 'xlsx']) {
    const bytes = packageFor(host, pic);
    const doc = core.open(bytes);
    try {
      doc.loadUnit(0);
      const object = doc.scene.objects.find(o => o.type === 'image');
      assert.ok(object, `${host}/${name}: picture missing`);
      const layers = [];
      for (let visual = object.visual; visual; visual = visual.visual) layers.push(visual);
      const camera = layers.find(v => v.threeD)?.threeD;
      if (name === 'original' && camera?.cameraPreset !== 'perspectiveHeroicExtremeLeftFacing') failures.push(`${host}: lost 3D camera`);
      if (name === 'original' && Math.abs((camera?.cameraLongitude ?? 0) - 34.5) > .001) failures.push(`${host}: lost preset angles`);
      if (name === 'crop' && !layers.some(v => (v.cropLeft ?? v.fill?.cropLeft) > .19)) failures.push(`${host}: lost crop`);
      if (['triangle', 'custom', 'adjusted'].includes(name) && !layers.some(v => typeof (v.clip ?? v.geometry) === 'object')) failures.push(`${host}: lost ${name} clipping`);
      if (['flipH', 'flipV', 'rotate90', 'combined'].includes(name) && !layers.some(v => v.kind === 'layer')) failures.push(`${host}: lost ${name} transform`);
      scenes.push({ name, host, bytes: Buffer.from(bytes).toString('base64') });
    } finally { doc.close(); }
  }
} finally { core.close(); }
console.log('Adapter regressions:', failures.length ? failures : `${scenes.length} passed`);
assert.deepEqual(failures, []);
if (process.env.PICTURE_PARSE_ONLY) process.exit(0);
await mkdir('output/picture-transforms', { recursive: true });
const server = spawn(process.execPath, ['scripts/serve.mjs', '--port', '0'], { stdio: ['ignore', 'pipe', 'inherit'] });
try {
  const url = await new Promise((resolve, reject) => {
    server.once('error', reject); server.once('exit', code => reject(new Error(`server ${code}`)));
    createInterface({ input: server.stdout }).on('line', line => { const match = line.match(/https?:\/\/\S+/u); if (match) resolve(match[0]); });
  });
  for (const name of (process.env.DOCVIEWKIT_BROWSERS ?? 'chromium,firefox,webkit').split(',')) {
    console.log(`${name}: launching`);
    const browser = await ({chromium, firefox, webkit})[name].launch({headless:true, timeout:30000});
    try {
      const page = await browser.newPage({ viewport:{width:1200,height:900} });
      const timer=setTimeout(()=>void page.close().catch(()=>{}),90000); timer.unref();
      const errors=[];
      page.on('pageerror',e=>errors.push(e.message));
      page.on('console',m=>{if(m.type()==='error')errors.push(m.text());});
      await page.goto(`${url}examples/viewer.html?fixture=image-3d-rotation.pptx`);
      await page.waitForFunction(()=>document.documentElement.dataset.ready==='true');
      console.log(`${name}: viewer ready`);
      await page.screenshot({path:`output/picture-transforms/${name}-viewer.png`});
      const results = await page.evaluate(async scenes => {
        const {Core,DEFAULT_LIMITS}=await import('/dist/core.js');
        const {SceneRenderer}=await import('/dist/render.js');
        const core=await Core.create(await(await fetch('/dist/office-viewer-core.wasm')).arrayBuffer(),DEFAULT_LIMITS);
        const results = {};
        try { for (const {name,host,bytes} of scenes) {
          const doc=core.open(Uint8Array.from(atob(bytes),c=>c.charCodeAt(0))); doc.loadUnit(0);
          const object=structuredClone(doc.scene.objects.find(o=>o.type==='image'));
          const dx=200-object.bounds.x,dy=200-object.bounds.y;
          for(let v=object.visual;v;v=v.visual) if(v.kind==='layer') {
            const t=v.transform;t.e+=dx-t.a*dx-t.c*dy;t.f+=dy-t.b*dx-t.d*dy;
          }
          object.bounds.x=200;object.bounds.y=200;object.unitIndex=0;delete object.parentNumericId;
          const renderer=new SceneRenderer([object],DEFAULT_LIMITS);
          try {
            const frame=await renderer.render({type:'slide',index:0,id:'unit:0',name:'Picture',width:700,height:700},{unitIndex:0,scale:1});
            const canvas=new OffscreenCanvas(700,700),ctx=canvas.getContext('2d');ctx.drawImage(frame.bitmap,0,0);frame.bitmap.close();
            const pixels=ctx.getImageData(0,0,700,700).data;
            let count=0,x0=700,y0=700,x1=0,y1=0,hash=0;
            // Color silhouette excludes the canvas background and remains stable across hosts.
            const mask=[];
            for(let y=0;y<700;y++) for(let x=0;x<700;x++) {
              const i=(y*700+x)*4;
              if(pixels[i+3]>127 && pixels[i+2]<150 && pixels[i]>180 && pixels[i+1]>150){count++;x0=Math.min(x0,x);y0=Math.min(y0,y);x1=Math.max(x1,x);y1=Math.max(y1,y);mask.push(y*700+x);hash=(hash*31+x+y*700)|0;}
            }
            results[`${host}/${name}`]={count,bounds:[x0,y0,x1,y1],hash,mask,diagnostics:frame.diagnostics};
          } finally {renderer.close();doc.close();}
        }} finally {core.close();}
        const native = await createImageBitmap(await(await fetch('/tests/fixtures/image-3d-rotation-reference.png')).blob());
        const canvas = new OffscreenCanvas(700,700), ctx = canvas.getContext('2d');
        ctx.drawImage(native,200-3733800/9525,200-2758281/9525);native.close();
        const pixels=ctx.getImageData(0,0,700,700).data, mask=[];
        for(let y=180;y<470;y++)for(let x=180;x<410;x++) {
          const i=(y*700+x)*4;
          if(pixels[i+3]>127&&pixels[i+2]<150&&pixels[i]>180&&pixels[i+1]>150)mask.push(y*700+x);
        }
        results.native={mask};
        return results;
      }, scenes);
      const native=new Set(results.native.mask);delete results.native;
      const actual=new Set(results['pptx/original'].mask);
      const intersection=[...actual].filter(p=>native.has(p)).length;
      const iou=intersection/(actual.size+native.size-intersection);
      assert.ok(iou>.94,`${name}: native PowerPoint silhouette overlap ${iou}`);
      for(const [key,value] of Object.entries(results)) {assert.ok(value.count>100,`${name}/${key}: missing picture`);assert.deepEqual(value.diagnostics,[]);}
      for(const variant of Object.keys(cases)) {
        const reference=results[`pptx/${variant}`];
        for(const host of ['docx','xlsx']) {
          const actual=results[`${host}/${variant}`];
          assert.ok(Math.abs(reference.count-actual.count)<Math.max(20,reference.count*.01),`${name}/${host}/${variant}: silhouette differs from PPTX`);
          assert.ok(reference.bounds.every((value,i)=>Math.abs(value-actual.bounds[i])<=1),`${name}/${host}/${variant}: bounds differ`);
        }
      }
      const get=key=>results[`pptx/${key}`];
      assert.notEqual(get('flat').hash,get('original').hash,'real camera must change pixels');
      assert.equal(get('flat').hash,get('explicitFront').hash,'explicit zero camera angles override the preset');
      assert.notEqual(get('flat').hash,get('crop').hash,'source crop changes mapped image');
      assert.equal(get('cropPercent').hash,get('crop').hash,'strict and transitional crop percentages agree');
      assert.ok(get('cameraZoom').count<get('original').count*.65,'authored camera zoom scales the projected image');
      assert.notEqual(get('cameraFov').hash,get('original').hash,'authored field of view changes the perspective');
      assert.ok(get('padding').count<get('flat').count*.85,'negative crop adds empty padding');
      assert.ok(get('triangle').count<get('flat').count*.9,'preset clip removes exterior pixels');
      assert.ok(Math.abs(get('triangle').count-get('custom').count)<50,'custom and preset triangle clip agree');
      for(const variant of ['flipH','flipV','rotate90']) assert.notEqual(get('flat').hash,get(variant).hash,`${variant} must transform the image`);
      assert.deepEqual(errors,[]);
      clearTimeout(timer);
      for(const result of Object.values(results))delete result.mask;
      await writeFile(`output/picture-transforms/${name}.json`,JSON.stringify(results,null,2));
      console.log(`${name}: ${scenes.length} picture combinations passed; native silhouette overlap ${(iou*100).toFixed(2)}%; clean console`);
    } finally {await browser.close();}
  }
} finally {server.kill();}
