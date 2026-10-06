import { chromium, firefox, webkit } from 'playwright-core';
import assert from 'node:assert/strict';
import { mkdir } from 'node:fs/promises';
await mkdir('output/playwright/smartart-3d', {recursive:true});
const deadline = setTimeout(() => { console.error('SmartArt browser regression timed out'); process.exit(1); }, 60_000);
const name = process.env.BROWSER ?? 'chromium';
const browser = await ({chromium, firefox, webkit}[name]).launch({headless:true});
try {
 const page = await browser.newPage({viewport:{width:1200,height:1000},deviceScaleFactor:1});
 const errors=[];page.on('pageerror',e=>errors.push(e.message));page.on('console',m=>{if(['error','warning'].includes(m.type())) errors.push(m.text());});
 await page.goto('http://127.0.0.1:4173/examples/viewer.html?fixture=smartart-orgchart-3d.pptx');
 await page.waitForFunction(()=>document.documentElement.dataset.ready==='true');
 await page.waitForFunction(()=>[...document.querySelector('#viewer').shadowRoot.querySelectorAll('canvas')].some(c=>c.width>900));
 await page.screenshot({path:`output/playwright/smartart-3d/${process.env.STAGE??'after'}-${name}.png`});
 const pixels = await page.evaluate(async () => {
   const { Core, DEFAULT_LIMITS } = await import('/dist/core.js');
   const { SceneRenderer } = await import('/dist/render.js');
   const core = await Core.create(await (await fetch('/dist/office-viewer-core.wasm')).arrayBuffer(), DEFAULT_LIMITS);
   const doc = core.open(new Uint8Array(await (await fetch('/tests/fixtures/smartart-orgchart-3d.pptx')).arrayBuffer()));
   let renderer;
   try {
     doc.loadUnit(0);
     const node = doc.scene.objects.find(o => o.text === 'Top');
     renderer = new SceneRenderer([node], DEFAULT_LIMITS);
     const frame = await renderer.render({type:'slide',index:0,id:'smartart',name:'SmartArt',width:960,height:720}, {unitIndex:0});
     const canvas = new OffscreenCanvas(frame.bitmap.width, frame.bitmap.height);
     const ctx = canvas.getContext('2d', {willReadFrequently:true}); ctx.drawImage(frame.bitmap,0,0); frame.bitmap.close();
     const {x,y,width:w,height:h} = node.bounds;
     const sample = (a,b) => [...ctx.getImageData(Math.round(x+a),Math.round(y+b),1,1).data];
     const pixels = {left:sample(7,h/2), right:sample(w-7,h/2), face:sample(22,h/2),
       top:sample(22,15), bottom:sample(22,h-15), shadow:sample(w/2,h+2), background:sample(w/2,h+15)};
     renderer.close();
     renderer = new SceneRenderer([{...node, visual:{...node.visual,
       threeD:{...node.visual.threeD, cameraRevolution:90}}}], DEFAULT_LIMITS);
     const rotated = await renderer.render({type:'slide',index:0,id:'rotated',name:'Rotated',width:960,height:720}, {unitIndex:0});
     ctx.clearRect(0,0,canvas.width,canvas.height);ctx.drawImage(rotated.bitmap,0,0);rotated.bitmap.close();
     return {...pixels, rotatedRim:sample(w/2,h/2+w/2-7), rotatedFace:sample(w/2,h/2+w/2-22)};
   } finally {renderer?.close();doc.close();core.close();}
 });
 assert.ok(pixels.left[1] > pixels.face[1] + 30, `left bevel highlight missing: ${JSON.stringify(pixels)}`);
 assert.ok(pixels.right[1] > pixels.face[1] + 30, `right bevel highlight missing: ${JSON.stringify(pixels)}`);
 assert.ok(pixels.top[2] > pixels.bottom[2] + 20, 'authored gradient missing');
 assert.ok(pixels.background[0] > pixels.shadow[0] + 15, 'outer shadow missing');
 assert.ok(pixels.rotatedRim[1] > pixels.rotatedFace[1] + 15, `rotated bevel projected twice: ${JSON.stringify(pixels)}`);
 assert.deepEqual(errors, []);
 console.log(JSON.stringify({browser:name,errors,pixels}));
} finally {await browser.close();clearTimeout(deadline);}
