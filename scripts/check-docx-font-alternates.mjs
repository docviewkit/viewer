import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { mkdir, writeFile } from 'node:fs/promises';
import { createInterface } from 'node:readline';
import { chromium, firefox, webkit } from 'playwright-core';

const server = spawn(process.execPath, ['scripts/serve.mjs', '--port', '0']);
const url = await new Promise((resolve, reject) => {
  createInterface({ input: server.stdout }).on('line', line => {
    const match = line.match(/http:\/\/\S+/); if (match) resolve(match[0]);
  });
  server.once('exit', code => reject(new Error(`Server exited: ${code}`)));
});
try {
  await mkdir('.cache/complex0-repair', { recursive: true });
  for (const name of (process.env.DOCVIEWKIT_BROWSERS ?? 'chromium,firefox,webkit').split(',')) {
    const browser = await ({ chromium, firefox, webkit })[name].launch({ timeout: 15000 });
    try {
      const page = await browser.newPage({ viewport: { width: 1100, height: 1250 } });
      const errors = [];
      page.on('pageerror', error => errors.push(error.message));
      page.on('console', message => { if (message.type() === 'error') errors.push(message.text()); });
      await page.goto(`${url}examples/viewer.html?fixture=complex0.docx&theme=light`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === 'true', null, { timeout: 60000 });
      await page.screenshot({ path: `.cache/complex0-repair/${name}-viewer.png` });
      for (const execution of ['inline', 'worker']) {
        const result = await page.evaluate(async execution => {
          const { createOfficeEngine } = await import('/dist/engine.js');
          const { fontShorthand } = await import('/dist/font.js');
          const bytes = await (await fetch('/tests/fixtures/complex0.docx')).arrayBuffer();
          const engine = await createOfficeEngine({ execution });
          let doc;
          try {
            doc = await engine.open(bytes);
            const objects = await doc.listObjects({ unitIndex: 0, textOnly: true });
            const runs = objects.flatMap(object => (object.fontRuns ?? []).map(run => ({ ...run, text: object.text?.slice(run.start, run.end) })));
            const alternates = runs.filter(run => run.authoredFamily === 'Bremen Bd BT');
            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');
            ctx.font = fontShorthand(alternates[0]?.renderedFamily ?? '', 16, false, false);
            const actualWidth = ctx.measureText('This column of text is 2.75 inches').width;
            ctx.font = fontShorthand('Courier New', 16, false, false);
            const expectedWidth = ctx.measureText('This column of text is 2.75 inches').width;
            const frame = await doc.render({ unitIndex: 0, pixelRatio: 1, background: '#ffffff' });
            canvas.width = frame.pixelWidth; canvas.height = frame.pixelHeight;
            ctx.drawImage(frame.bitmap, 0, 0); frame.bitmap.close();
            const firstPage = canvas.toDataURL();
            let controls, webControls;
            for (const unit of doc.info.units) {
              const objects = await doc.listObjects({ unitIndex: unit.index });
              if (objects.some(object => object.text?.includes("background sound"))) {
                const region = objects.filter(object => object.source?.paragraphIndex >= 657 && object.source?.paragraphIndex <= 665);
                const top = Math.min(...region.map(object => object.bounds.y)) - 10;
                const bottom = Math.max(...region.map(object => object.bounds.y + object.bounds.height)) + 10;
                const frame = await doc.render({ unitIndex: unit.index, pixelRatio: 1, background: "#ffffff",
                  viewport: { x: 110, y: top, width: 600, height: bottom - top } });
                canvas.width = frame.pixelWidth; canvas.height = frame.pixelHeight;
                ctx.drawImage(frame.bitmap, 0, 0); frame.bitmap.close();
                webControls = canvas.toDataURL();
              }
              if (!objects.some(object => object.text?.includes("This button should"))) continue;
              const region = objects.filter(object => object.source?.paragraphIndex >= 668 && object.source?.paragraphIndex <= 671);
              const top = Math.min(...region.map(object => object.bounds.y)) - 10;
              const bottom = Math.max(...region.map(object => object.bounds.y + object.bounds.height)) + 10;
              const frame = await doc.render({ unitIndex: unit.index, pixelRatio: 1, includeTextFragments: true, background: "#ffffff",
                viewport: { x: 110, y: top, width: 600, height: bottom - top } });
              canvas.width = frame.pixelWidth; canvas.height = frame.pixelHeight;
              ctx.drawImage(frame.bitmap, 0, 0); frame.bitmap.close();
              controls = { objects, fragments: frame.textFragments, png: canvas.toDataURL() };
            }
            let matrix;
            for (const unit of doc.info.units) {
              const objects = await doc.listObjects({ unitIndex: unit.index });
              const equation = objects.find(object => object.text === '1 2 3\n2 3 1\n3 1 2');
              if (!equation) continue;
              const caption = objects.find(object => object.id === equation.parentId);
              const fractionEquation = objects.find(object => object.text === 'x = 1⁄2');
              const fractionCaption = objects.find(object => object.id === fractionEquation?.parentId);
              const rendered = await doc.render({ unitIndex: unit.index, pixelRatio: 1, includeTextFragments: true, background: '#ffffff',
                viewport: { x: 110, y: equation.bounds.y - 10, width: 596,
                  height: fractionCaption.bounds.y + fractionCaption.bounds.height - equation.bounds.y + 20 } });
              canvas.width = rendered.pixelWidth; canvas.height = rendered.pixelHeight;
              ctx.drawImage(rendered.bitmap, 0, 0); rendered.bitmap.close();
              matrix = { equation, caption, fragments: rendered.textFragments.filter(fragment => fragment.objectId === caption?.id
                || fragment.objectId.startsWith(`${equation.id}:part:`)), fractionCaption,
                fractionFragments: rendered.textFragments.filter(fragment => fragment.objectId === fractionCaption?.id),
                png: canvas.toDataURL() };
              break;
            }
            return { units: doc.info.units.length, alternates, actualWidth, expectedWidth,
              diagnostics: [...doc.diagnostics(), ...frame.diagnostics], png: firstPage, matrix, controls, webControls };
          } finally { doc?.close(); engine.close(); }
        }, execution);
        await writeFile(`.cache/complex0-repair/${name}-${execution}.png`, Buffer.from(result.png.split(',')[1], 'base64'));
        delete result.png;
        await writeFile(`.cache/complex0-repair/${name}-${execution}-controls.png`, Buffer.from(result.controls.png.split(',')[1], 'base64'));
        delete result.controls.png;
        await writeFile(`.cache/complex0-repair/${name}-${execution}-web-controls.png`, Buffer.from(result.webControls.split(',')[1], 'base64'));
        delete result.webControls;
        await writeFile(`.cache/complex0-repair/${name}-${execution}-matrix.png`, Buffer.from(result.matrix.png.split(',')[1], 'base64'));
        delete result.matrix.png;
        await writeFile(`.cache/complex0-repair/${name}-${execution}.json`, JSON.stringify(result, null, 2));
        assert.ok(result.alternates.length >= 4, 'The original two-column paragraph must retain its regular, bold, italic and bold-italic runs');
        for (const run of result.alternates) {
          assert.match(run.renderedFamily, /Courier/iu, `${name}/${execution}: ${JSON.stringify(run)}`);
          assert.equal(run.source, 'fallback');
        }
        assert.ok(Math.abs(result.actualWidth - result.expectedWidth) < .01, 'Layout and rendering must use Courier New advances');
        assert.equal(result.units, 10, 'Courier New matches the explicit-name control; remaining decorative fonts still use approximations');
        assert.ok(result.diagnostics.some(d => d.code === 'FONT_METRICS_APPLIED'));
        assert.ok(result.diagnostics.some(d => d.code === 'FONT_SUBSTITUTED' && d.message.includes('Bremen Bd BT') && d.message.includes('Courier')));
        const captionStart = result.matrix.fragments.find(fragment => fragment.objectId === result.matrix.caption.id && fragment.text.startsWith('3'));
        assert.equal(captionStart?.line, 0, 'The matrix description starts beside its middle row');
        assert.ok(Math.abs(captionStart.transform.f - result.matrix.caption.bounds.y) < 1,
          'An authored equation baseline must not receive paragraph leading a second time');
        const fractionText = result.matrix.fractionFragments.find(fragment => fragment.text.trim());
        assert.ok(fractionText && Math.abs(fractionText.transform.f - result.matrix.fractionCaption.bounds.y) < 1,
          'The fraction description shares its authored baseline without extra leading');
        assert.ok(!result.controls.objects.some(object => /^(Top|Bottom) of Form$/u.test(object.text ?? "")), "Hidden form markers stay invisible");
        for (const index of [669, 671]) {
          const objects = result.controls.objects.filter(object => object.source?.paragraphIndex === index);
          const image = objects.find(object => object.type === "image");
          const fragment = result.controls.fragments.find(fragment => objects.some(object => object.id === fragment.objectId) && fragment.text.trim());
          assert.ok(fragment && fragment.transform.e >= image.bounds.x + image.bounds.width - .01, "The caption starts after its inline control");
        }
        console.log(`${name}/${execution}: declared Courier New, font advances, matrix/fraction baselines, inline controls and hidden form markers passed`);
      }
      assert.deepEqual(errors, []);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
