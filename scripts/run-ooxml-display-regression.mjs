import assert from 'node:assert/strict';
import { mkdir, readFile } from 'node:fs/promises';
import { createZip } from '../tests/zip-fixture.mjs';
import { readZipEntries } from './accuracy-metamorphic.mjs';
import { chromium, firefox, webkit } from 'playwright-core';

await mkdir('output/ooxml-display', { recursive: true });
for (const [name, type] of Object.entries({ chromium, firefox, webkit })) {
  if (process.env.BROWSER && process.env.BROWSER !== name) continue;
  const deadline = setTimeout(() => {
    console.error(`${name}: browser regression exceeded 90 seconds`);
    process.exit(1);
  }, 90_000);
  const browser = await type.launch({ headless: true });
  try {
    const page = await browser.newPage({ viewport: { width: 1400, height: 1100 } });
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    page.on('console', message => { if (message.type() === 'error') errors.push(message.text()); });
    await page.addInitScript(() => {
      globalThis.paintedFills = [];
      globalThis.paintedSegments = [];
      globalThis.printImages = [];
      const toBlob = HTMLCanvasElement.prototype.toBlob;
      HTMLCanvasElement.prototype.toBlob = function(...args) {
        globalThis.printImages.push({ width: this.width, height: this.height, text: this.paintedText ?? [] });
        return toBlob.apply(this, args);
      };
      addEventListener('DOMContentLoaded', () => new MutationObserver(() => {
        for (const frame of document.querySelectorAll('iframe')) {
          frame.contentWindow.print = () => { globalThis.printed = true; };
        }
      }).observe(document.body, { childList: true }));
      for (const prototype of [CanvasRenderingContext2D.prototype, OffscreenCanvasRenderingContext2D.prototype]) {
        for (const name of ['fill', 'fillRect']) {
          const original = prototype[name];
          prototype[name] = function(...args) {
            globalThis.paintedFills.push(this.fillStyle);
            return original.apply(this, args);
          };
        }
        for (const name of ['moveTo', 'lineTo']) {
          const original = prototype[name];
          prototype[name] = function(x, y) {
            if (name === 'lineTo' && this.lastMove) globalThis.paintedSegments.push([this.lastMove, [x, y]]);
            this.lastMove = [x, y];
            return original.call(this, x, y);
          };
        }
        const fillText = prototype.fillText;
        prototype.fillText = function(text, ...args) {
          (this.canvas.paintedText ??= []).push({ text, font: this.font, x: args[0], y: args[1] });
          return fillText.call(this, text, ...args);
        };
        const drawImage = prototype.drawImage;
        prototype.drawImage = function(source, ...args) {
          if (source.paintedText) this.canvas.paintedText = [...(this.canvas.paintedText ?? []), ...source.paintedText];
          return drawImage.call(this, source, ...args);
        };
      }
      const createBitmap = globalThis.createImageBitmap;
      globalThis.createImageBitmap = async function(source, ...args) {
        const bitmap = await createBitmap(source, ...args);
        bitmap.paintedText = source.paintedText;
        return bitmap;
      };
      const transfer = OffscreenCanvas.prototype.transferToImageBitmap;
      OffscreenCanvas.prototype.transferToImageBitmap = function() {
        const bitmap = transfer.call(this);
        bitmap.paintedText = this.paintedText;
        return bitmap;
      };
    });
    // Inline execution makes the real Canvas text calls observable in this page.
    await page.route('**/examples/viewer.js', async route => {
      const response = await route.fetch();
      await route.fulfill({ response, body: (await response.text()).replace('engine: {', 'engine: { execution: "inline",') });
    });
    for (const [fixture, expected] of [
      ['ooxml-minor-ticks.xlsx', ['2', '4']],
      ['ooxml-calendar-axis.xlsx', ['1/1/2019', '2/1/2019', '3/1/2019', '4/1/2019', '5/1/2019']],
      ['ooxml-reversed-categories.docx', ['Concursuri', 'Saloane']],
      ['ooxml-gutter-left.docx', ['Half in gutter']],
      ['ooxml-gutter-top.docx', ['He heard quiet']],
      ['ooxml-gutter-right.docx', ['hello']],
      ['ooxml-mirror-margins.docx', ['Lorem']],
      ['ooxml-ruby.docx', ['きもん', 'ほうがく', 'ぎょうし']],
      ['ooxml-display-units.docx', ['Billions', '2E-9']],
      ['ooxml-multilevel-axis.docx', ['Categoria 1', 'Categoria 2', 'Categoria 3', 'Categoria 4', '2011', '2012']],
      ['ooxml-nested-alternate.docx', ['ABC', 'PQR']],
      ['ooxml-conditional-priority.xlsx', ['ABC', 'BAC']],
    ]) {
      await page.goto(`${process.env.VIEWER_URL ?? 'http://127.0.0.1:4173'}/examples/viewer.html?fixture=${fixture}`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === 'true');
      await page.waitForFunction(() => [...document.querySelector('#viewer').shadowRoot.querySelectorAll('canvas')]
        .some(canvas => canvas.paintedText?.length > 0));
      const painted = await page.locator('#viewer').evaluate(viewer =>
        [...viewer.shadowRoot.querySelectorAll('canvas')].flatMap(canvas => canvas.paintedText ?? []));
      const text = painted.map(item => item.text).join('');
      for (const token of expected) assert.ok(text.includes(token), `${name} ${fixture}: missing ${token}: ${text}`);
      if (fixture === 'ooxml-minor-ticks.xlsx') {
        assert.ok(await page.evaluate(() => globalThis.paintedSegments.filter(([a, b]) => Math.abs(Math.abs(a[0] - b[0]) - 3) < .02 && Math.abs(a[1] - b[1]) < .02).length >= 12), `${name}: minor ticks must reach Canvas paths`);
      }
      if (fixture === 'ooxml-mirror-margins.docx') {
        const first = painted.find(item => item.text.includes('Lorem'));
        await page.locator('#viewer').locator('[data-action="next"]').click();
        await page.waitForFunction(() => [...document.querySelector('#viewer').shadowRoot.querySelectorAll('canvas')].some(canvas => canvas.paintedText?.some(item => item.text.includes('Lorem') && Math.abs(item.x - 280) < .1)));
        const second = await page.locator('#viewer').evaluate(viewer => [...viewer.shadowRoot.querySelectorAll('canvas')].flatMap(canvas => canvas.paintedText ?? []).find(item => item.text.includes('Lorem') && Math.abs(item.x - 280) < .1));
        assert.ok(second.x > first.x + 100, `${name}: even-page body uses mirrored left margin: ${first.x} -> ${second.x}`);
      }
      if (fixture === 'ooxml-reversed-categories.docx') {
        const first = painted.find(item => item.text.includes('Concursuri'));
        const last = painted.find(item => item.text.includes('Saloane'));
        assert.ok(first.y < last.y, `${name}: reversed categories painted in authored order`);
      }
      if (fixture === 'ooxml-conditional-priority.xlsx') {
        assert.ok(await page.evaluate(() => globalThis.paintedFills.some(color => ['rgba(255, 0, 0, 1)', '#ff0000'].includes(color))),
          `${name}: real conditional-format cells must be red`);
      }
      await page.locator('#viewer').screenshot({ path: `output/ooxml-display/${name}-${fixture}.png` });
      assert.deepEqual(errors, [], `${name}: clean console`);
      console.log(`${name}: ${fixture}: visible labels and clean console passed`);
    }
    {
      const parts = Object.fromEntries(readZipEntries(await readFile('tests/fixtures/ooxml-minor-ticks.xlsx'))
        .map(({ name, data }) => [name, data]));
      for (const part of ['[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml',
        'xl/_rels/workbook.xml.rels', 'xl/worksheets/sheet1.xml']) {
        const xml = new TextDecoder().decode(parts[part]).replace('encoding="UTF-8"', 'encoding="UTF-16"');
        parts[part] = Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(xml, 'utf16le')]);
      }
      await page.route('**/tests/fixtures/ooxml-minor-ticks.xlsx', route =>
        route.fulfill({ body: Buffer.from(createZip(parts)), contentType: 'application/octet-stream' }));
      await page.goto(`${process.env.VIEWER_URL ?? 'http://127.0.0.1:4173'}/examples/viewer.html?fixture=ooxml-minor-ticks.xlsx`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === 'true');
      await page.waitForFunction(() => [...document.querySelector('#viewer').shadowRoot.querySelectorAll('canvas')]
        .some(canvas => canvas.paintedText?.some(item => item.text === 'col1')));
      const text = await page.locator('#viewer').evaluate(viewer =>
        [...viewer.shadowRoot.querySelectorAll('canvas')].flatMap(canvas => canvas.paintedText ?? []).map(item => item.text).join(''));
      assert.ok(text.includes('col2'), `${name}: UTF-16 worksheet text reaches Canvas`);
      assert.deepEqual(errors, [], `${name}: UTF-16 workbook console`);
      console.log(`${name}: derived UTF-16 workbook: visible cells and clean console passed`);
      await page.unroute('**/tests/fixtures/ooxml-minor-ticks.xlsx');
    }
    for (const [carrier, part, token] of [
      ['ooxml-nested-alternate.docx', 'word/document.xml', 'ABC'],
      ['corpus-chart-wall.pptx', 'ppt/slides/slide1.xml', 'Anne'],
    ]) {
      const parts = Object.fromEntries(readZipEntries(await readFile(`tests/fixtures/${carrier}`))
        .map(({ name, data }) => [name, data]));
      const xml = new TextDecoder().decode(parts[part]).replace('encoding="UTF-8"', 'encoding="UTF-16"');
      parts[part] = Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(xml, 'utf16le')]);
      await page.route(`**/tests/fixtures/${carrier}`, route =>
        route.fulfill({ body: Buffer.from(createZip(parts)), contentType: 'application/octet-stream' }));
      await page.goto(`${process.env.VIEWER_URL ?? 'http://127.0.0.1:4173'}/examples/viewer.html?fixture=${carrier}`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === 'true');
      await page.waitForFunction(token => [...document.querySelector('#viewer').shadowRoot.querySelectorAll('canvas')]
        .some(canvas => canvas.paintedText?.some(item => item.text.includes(token))), token);
      assert.deepEqual(errors, [], `${name} ${carrier}: UTF-16 main XML console`);
      console.log(`${name}: derived UTF-16 ${carrier}: visible text and clean console passed`);
      await page.unroute(`**/tests/fixtures/${carrier}`);
    }
    {
      const parts = Object.fromEntries(readZipEntries(await readFile('tests/fixtures/ooxml-omml-slide.pptx'))
        .map(({ name, data }) => [name, data]));
      const xml = new TextDecoder().decode(parts['ppt/slides/slide1.xml']);
      const start = xml.indexOf('<m:oMath xmlns:m=');
      const contentStart = xml.indexOf('>', start) + 1;
      const end = xml.indexOf('</m:oMath>', contentStart);
      const fraction = '<m:f><m:num><m:r><m:t xml:space="preserve">a</m:t></m:r></m:num><m:den><m:r><m:t xml:space="preserve">b</m:t></m:r></m:den></m:f>';
      parts['ppt/slides/slide1.xml'] = xml.slice(0, contentStart) + fraction + xml.slice(end);
      await page.route('**/tests/fixtures/ooxml-omml-slide.pptx', route =>
        route.fulfill({ body: Buffer.from(createZip(parts)), contentType: 'application/octet-stream' }));
      await page.goto(`${process.env.VIEWER_URL ?? 'http://127.0.0.1:4173'}/examples/viewer.html?fixture=ooxml-omml-slide.pptx`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === 'true');
      await page.waitForFunction(() => [...document.querySelector('#viewer').shadowRoot.querySelectorAll('canvas')]
        .some(canvas => canvas.paintedText?.some(item => item.text === 'b')));
      const math = await page.locator('#viewer').evaluate(viewer =>
        [...viewer.shadowRoot.querySelectorAll('canvas')].flatMap(canvas => canvas.paintedText ?? [])
          .filter(item => item.text === 'a' || item.text === 'b'));
      assert.ok(math.some(item => item.text === 'a'), `${name}: missing numerator`);
      assert.ok(math.find(item => item.text === 'a').y < math.find(item => item.text === 'b').y,
        `${name}: fraction remains flattened in Canvas`);
      assert.deepEqual(errors, [], `${name}: PPTX OMML console`);
      await page.locator('#viewer').screenshot({ path: `output/ooxml-display/${name}-omml-fraction.png` });
      console.log(`${name}: derived real OMML fraction: stacked Canvas text and clean console passed`);
    }
    const calendarChart = readZipEntries(await readFile('tests/fixtures/ooxml-calendar-axis.xlsx'))
      .find(entry => entry.name === 'xl/charts/chart1.xml').data;
    for (const carrier of ['ooxml-display-units.docx', 'corpus-chart-wall.pptx']) {
      const parts = Object.fromEntries(readZipEntries(await readFile(`tests/fixtures/${carrier}`))
        .map(({ name, data }) => [name, data]));
      parts[Object.keys(parts).find(part => /^(word|ppt)\/charts\/chart1\.xml$/u.test(part))] = calendarChart;
      await page.route(`**/tests/fixtures/${carrier}`, route => route.fulfill({ body: Buffer.from(createZip(parts)), contentType: 'application/octet-stream' }));
      await page.goto(`${process.env.VIEWER_URL ?? 'http://127.0.0.1:4173'}/examples/viewer.html?fixture=${carrier}`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === 'true');
      await page.waitForFunction(() => [...document.querySelector('#viewer').shadowRoot.querySelectorAll('canvas')]
        .some(canvas => canvas.paintedText?.some(item => item.text.includes('2/1/2019'))));
      const painted = await page.locator('#viewer').evaluate(viewer =>
        [...viewer.shadowRoot.querySelectorAll('canvas')].flatMap(canvas => canvas.paintedText ?? []).map(item => item.text).join(''));
      assert.ok(painted.includes('5/1/2019'), `${name} ${carrier}: calendar labels reach Canvas`);
      assert.deepEqual(errors, [], `${name} ${carrier}: clean console`);
      console.log(`${name}: derived calendar ChartML in ${carrier}: visible labels and clean console passed`);
    }
    // Keep the real workbook's styles/geometry; add visible text to its empty
    // authored title row and a second page so the print consumer is observable.
    const parts = Object.fromEntries(readZipEntries(await readFile('tests/fixtures/ooxml-print-titles.xlsx')).map(({ name, data }) => [name, data]));
    parts['xl/worksheets/sheet1.xml'] = new TextDecoder().decode(parts['xl/worksheets/sheet1.xml'])
      .replace('<c r="B5" s="12"/>', '<c r="B5" s="12" t="inlineStr"><is><t>REPEAT_ROW_5</t></is></c>')
      .replace(/fitToPage="1"/gu, 'fitToPage="0"')
      .replace('<pageSetup ', '<pageSetup scale="50" ')
      .replace('</worksheet>', '<rowBreaks count="1" manualBreakCount="1"><brk id="10" max="16383" man="1"/></rowBreaks></worksheet>');
    await page.route('**/tests/fixtures/ooxml-print-titles.xlsx', route => route.fulfill({ body: Buffer.from(createZip(parts)), contentType: 'application/octet-stream' }));
    await page.goto(`${process.env.VIEWER_URL ?? 'http://127.0.0.1:4173'}/examples/viewer.html?fixture=ooxml-print-titles.xlsx`);
    await page.waitForFunction(() => document.documentElement.dataset.ready === 'true');
    await page.locator('#viewer').locator('[data-action="print"]').click();
    await page.waitForFunction(() => globalThis.printed === true);
    const images = await page.evaluate(() => globalThis.printImages);
    assert.equal(images.length, 2, `${name}: real manual break produces two print pages`);
    for (const image of images) assert.ok(image.text.map(item => item.text).join('').includes('REPEAT_ROW_5'), `${name}: title row painted on every print page`);
    await page.locator('iframe').evaluate(frame => { frame.style.cssText = 'position:fixed;left:0;top:0;width:1100px;height:1050px;border:0;z-index:999999'; });
    await page.locator('iframe').screenshot({ path: `output/ooxml-display/${name}-print-titles.png` });
    assert.deepEqual(errors, [], `${name}: print console`);
    console.log(`${name}: actual viewer print: two pages, repeated title and clean console passed`);
  } finally { await browser.close(); clearTimeout(deadline); }
}
