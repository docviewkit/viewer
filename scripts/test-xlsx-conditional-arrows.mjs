import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { chromium, firefox, webkit } from 'playwright-core';
const server = spawn(process.execPath, ['scripts/serve.mjs', '--port', '0'], {stdio: ['ignore','pipe','inherit']});
try {
  const url = await new Promise(resolve => createInterface({input: server.stdout}).on('line', line => { const m = line.match(/https?:\/\/\S+/); if(m) resolve(m[0]); }));
  for (const name of (process.env.DOCVIEWKIT_BROWSERS ?? 'chromium,firefox,webkit').split(',')) {
    const browser = await ({chromium,firefox,webkit})[name].launch({headless:true,timeout:30000});
    try {
      const page = await browser.newPage({viewport:{width:1000,height:700}});
      const errors = [];
      page.on('pageerror', e => errors.push(e.message));
      page.on('console', m => {if(m.type()==='error') errors.push(m.text());});
      if(process.env.ARROWS_WASM) await page.route('**/dist/office-viewer-core.wasm', r => r.fulfill({path:process.env.ARROWS_WASM,contentType:'application/wasm'}));
      await page.goto(`${url}examples/viewer.html?fixture=tdf162948.xlsx`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === 'true');
      const result = await page.evaluate(async () => {
        const {Core,DEFAULT_LIMITS} = await import('/dist/core.js');
        const {SceneRenderer} = await import('/dist/render.js');
        const core = await Core.create(await (await fetch('/dist/office-viewer-core.wasm')).arrayBuffer(),DEFAULT_LIMITS);
        const doc = core.open(new Uint8Array(await (await fetch('/tests/fixtures/tdf162948.xlsx')).arrayBuffer()));
        doc.loadUnit(0);
        const objects = doc.scene.objects;
        const result = {cells:objects.filter(o=>o.type==='cell'), icons:objects.filter(o=>o.type==='shape')};
        const renderer = new SceneRenderer(objects, DEFAULT_LIMITS);
        try {
          const frame = await renderer.render(doc.scene.info.units[0], {unitIndex:0,scale:1});
          const canvas = new OffscreenCanvas(frame.bitmap.width,frame.bitmap.height);
          const context = canvas.getContext('2d'); context.drawImage(frame.bitmap,0,0); frame.bitmap.close();
          const pixels = context.getImageData(0,0,canvas.width,canvas.height).data;
          result.yellow = 0; result.green = 0; result.diagnostics = frame.diagnostics;
          for(let i=0;i<pixels.length;i+=4) {
            if(pixels[i]===255 && pixels[i+1]===192 && pixels[i+2]===0) result.yellow++;
            if(pixels[i]===112 && pixels[i+1]===173 && pixels[i+2]===71) result.green++;
          }
        } finally {renderer.close();doc.close();core.close();}
        return result;
      });
      await page.screenshot({path:`output/playwright/tdf162948-${name}${process.env.ARROWS_WASM ? '-before' : '-after'}.png`});
      assert.equal(result.icons.length,8);
      result.icons.forEach((icon,i)=>{
        const cell=result.cells[i];
        assert.ok(Math.abs(icon.bounds.x-cell.bounds.x-2)<0.1,'icon is at left edge');
        assert.ok(Math.abs(icon.bounds.y+icon.bounds.height/2-cell.bounds.y-cell.bounds.height/2)<0.1,'icon is vertically centered');
        assert.equal(icon.visual.geometry.kind,'path','solid vector arrow');
        assert.equal(icon.visual.fill,i===1?0x70ad47ff:0xffc000ff,'A1 is yellow, B1 green');
        const tip=icon.visual.geometry.commands[0];
        assert.ok(i===1?tip.y<2&&tip.x>5:tip.x>12&&tip.y>5,'A1 points right, B1 up');
      });
      assert.ok(result.yellow > 100 && result.green > 20, 'arrow fills reach rendered pixels');
      assert.deepEqual(result.diagnostics,[]);
      assert.deepEqual(errors,[]);
      console.log(`${name}: 8 arrows, directions, placement, vector styles and clean console passed`);
    } finally {await browser.close();}
  }
} finally {server.kill();}
