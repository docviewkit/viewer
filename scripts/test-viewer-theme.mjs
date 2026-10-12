// Run after npm run build:js. DOCVIEWKIT_BROWSERS optionally selects engines.
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { readZipEntries, writeZipEntries } from "./accuracy-metamorphic.mjs";
import { spawn } from "node:child_process";
import { createInterface } from "node:readline";
import { chromium, webkit } from "playwright-core";

async function contentScreenshot(page, surface) {
  await surface.waitFor({ state: "visible" });
  const bounds = await surface.boundingBox(), workspace = await page.locator(".workspace").boundingBox();
  const viewport = page.viewportSize();
  const x = Math.max(0, bounds.x, workspace.x), y = Math.max(0, bounds.y, workspace.y);
  // Tall continuous canvases extend behind chrome; compare only their visible document region.
  return page.screenshot({ clip: { x, y,
    width: Math.min(viewport.width, bounds.x + bounds.width, workspace.x + workspace.width) - x,
    height: Math.min(viewport.height, bounds.y + bounds.height, workspace.y + workspace.height) - y,
  } });
}

async function assertDarkContent(page, light, dark, label, authoredFills = false) {
  const counts = await page.evaluate(async ([light, dark]) => {
    const decode = async (data) => {
      const bytes = Uint8Array.from(atob(data), (character) => character.charCodeAt(0));
      const bitmap = await createImageBitmap(new Blob([bytes], { type: "image/png" }));
      const canvas = document.createElement("canvas");
      canvas.width = bitmap.width;
      canvas.height = bitmap.height;
      const context = canvas.getContext("2d");
      context.drawImage(bitmap, 0, 0);
      bitmap.close();
      return context.getImageData(0, 0, canvas.width, canvas.height);
    };
    const { data: before, width, height } = await decode(light);
    const { data: after, width: darkWidth, height: darkHeight } = await decode(dark);
    if (width !== darkWidth || height !== darkHeight) throw new Error("Theme changed document dimensions");
    let paper = 0, darkPaper = 0, ink = 0, lightInk = 0;
    // Fractional canvas edges blend with surrounding chrome; compare the document interior.
    for (let y = 1; y < height - 1; y++) for (let x = 1; x < width - 1; x++) {
      const i = (y * width + x) * 4;
      if (before[i] >= 245 && before[i + 1] >= 245 && before[i + 2] >= 245) {
        paper++;
        if (after[i] < 40 && after[i + 1] < 40 && after[i + 2] < 40) darkPaper++;
      }
      if (before[i] < 100 && before[i + 1] < 100 && before[i + 2] < 100) {
        ink++;
        if ((after[i] + after[i + 1] + after[i + 2]) / 3 > 140) lightInk++;
      }
    }
    return { paper, darkPaper, ink, lightInk };
  }, [light.toString("base64"), dark.toString("base64")]);
  assert.ok(counts.paper > 100 && (authoredFills ? counts.darkPaper > 100 : counts.darkPaper / counts.paper > .98),
    `${label}: white document content must become dark: ${JSON.stringify(counts)}`);
  // Thumbnail downsampling can leave no pixels below the dark-ink threshold.
  assert.ok(counts.ink === 0 || (authoredFills ? counts.lightInk > 0 : counts.lightInk / counts.ink > .98),
    `${label}: dark document text must remain readable: ${JSON.stringify(counts)}`);
}

async function setReadingTheme(page, viewer, theme, adapted) {
  const before = await viewer.evaluate(v => v.shadowRoot.querySelector(v.state.pageMode === "continuous"
    ? '.continuous-page[data-rendered="true"] canvas' : '.stage .surface')?.toDataURL());
  await viewer.evaluate((v, theme) => v.setAttribute("theme", theme), theme);
  if (adapted) await page.waitForFunction(([theme, before]) => {
    const viewer = document.querySelector("#viewer");
    const canvas = viewer.shadowRoot.querySelector(viewer.state.pageMode === "continuous"
      ? '.continuous-page[data-rendered="true"] canvas' : '.stage .surface');
    return viewer.state.theme === theme && canvas?.width > 1
      && canvas.dataset.preview !== "true" && canvas.toDataURL() !== before;
  }, [theme, before]);
}

// Real DOCX structure: automatic overrides, explicit colors, a white paragraph and an embedded image.
const readingEntries = readZipEntries(readFileSync("tests/fixtures/inline-images-textbox.docx"));
const documentEntry = readingEntries.find(entry => entry.name === "word/document.xml");
const originalXml = new TextDecoder().decode(documentEntry.data);
const drawing = originalXml.match(/<w:drawing>[\s\S]*?<\/w:drawing>/u)?.[0];
assert.ok(drawing, "reading regression needs a real embedded image");
const paragraph = (text, color = "auto", properties = "") => `<w:p>${properties}<w:r><w:rPr><w:sz w:val="36"/><w:color w:val="${color}"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
documentEntry.data = new TextEncoder().encode(originalXml.replace(/<w:body>[\s\S]*<\/w:body>/u,
  `<w:body>${paragraph("Auto ink")}${paragraph("Explicit red", "FF0000")}${paragraph("Explicit black", "000000")}${paragraph("White box", "auto", '<w:pPr><w:shd w:val="clear" w:fill="FFFFFF"/></w:pPr>')}<w:p><w:r>${drawing}</w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr></w:body>`));
const readingBytes = [...writeZipEntries(readingEntries)];

const packageBytes = entries => [...writeZipEntries(Object.entries({ "styles.xml": '<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:styles/></office:document-styles>', ...entries, "META-INF/manifest.xml": `<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="${entries.mimetype}"/></manifest:manifest>` }).map(([name, xml]) => ({ name, data: new TextEncoder().encode(xml), compression: name === "mimetype" ? "store" : "deflate" })))];
const sheetEntries = readZipEntries(readFileSync("tests/fixtures/basicspreadsheet-custom-styles.xlsx"));
const sheetXml = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${["Auto ink", "Explicit red", "Explicit black", "White box"].map((text, index) => `<row r="${index + 1}" ht="30" customHeight="1"><c r="A${index + 1}" s="${index}" t="inlineStr"><is><t>${text}</t></is></c></row>`).join("")}</sheetData></worksheet>`;
for (const entry of sheetEntries) {
  if (entry.name === "xl/worksheets/sheet1.xml") entry.data = new TextEncoder().encode(sheetXml);
  if (entry.name === "xl/styles.xml") entry.data = new TextEncoder().encode(`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="3"><font><sz val="14"/><color theme="1"/><name val="Arial"/></font><font><sz val="14"/><color rgb="FFFF0000"/><name val="Arial"/></font><font><sz val="14"/><color rgb="FF000000"/><name val="Arial"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/></patternFill></fill></fills><borders count="1"><border/></borders><cellXfs count="4">${[0, 1, 2, 0].map((font, index) => `<xf fontId="${font}" fillId="${index === 3 ? 1 : 0}" borderId="0" numFmtId="0"/>`).join("")}</cellXfs></styleSheet>`);
}
const odfNamespaces = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"';
const odfContent = (sheet) => `<office:document-content ${odfNamespaces}><office:automatic-styles><style:style style:name="auto" style:family="${sheet ? "table-cell" : "paragraph"}"><style:text-properties fo:color="#FF0000" style:use-window-font-color="true"/></style:style><style:style style:name="red" style:family="${sheet ? "table-cell" : "paragraph"}"><style:text-properties fo:color="#FF0000"/></style:style><style:style style:name="black" style:family="${sheet ? "table-cell" : "paragraph"}"><style:text-properties fo:color="#000000"/></style:style><style:style style:name="white" style:family="${sheet ? "table-cell" : "paragraph"}">${sheet ? '<style:table-cell-properties fo:background-color="#FFFFFF"/>' : '<style:text-properties fo:background-color="#FFFFFF"/>'}</style:style></office:automatic-styles><office:body><office:${sheet ? "spreadsheet" : "text"}>${sheet ? '<table:table table:name="Reading">' : ""}${["Auto ink", "Explicit red", "Explicit black", "White box"].map((text, index) => {
  const style = ["auto", "red", "black", "white"][index];
  return sheet ? `<table:table-row><table:table-cell table:style-name="${style}" office:value-type="string"><text:p>${text}</text:p></table:table-cell></table:table-row>` : `<text:p text:style-name="${style}">${text}</text:p>`;
}).join("")}${sheet ? '</table:table>' : ""}</office:${sheet ? "spreadsheet" : "text"}></office:body></office:document-content>`;
const semanticCases = [
  { format: "docx", bytes: readingBytes, image: true },
  { format: "xlsx", bytes: [...writeZipEntries(sheetEntries)], sheet: true },
  { format: "odt", bytes: packageBytes({ mimetype: "application/vnd.oasis.opendocument.text", "content.xml": odfContent(false) }) },
  { format: "ods", bytes: packageBytes({ mimetype: "application/vnd.oasis.opendocument.spreadsheet", "content.xml": odfContent(true) }), sheet: true },
  { format: "rtf", bytes: [...new TextEncoder().encode(String.raw`{\rtf1\ansi{\colortbl;\red255\green0\blue0;\red0\green0\blue0;}\fs36\cf0 Auto ink\par\cf1 Explicit red\par\cf2 Explicit black\par}`)] },
];

async function verifySemanticColors(page) {
  for (const execution of ["inline", "worker"]) for (const testCase of semanticCases) {
  console.log(`semantic: ${testCase.format}/${execution}`);
  const checks = await page.evaluate(async ({ bytes, format, sheet, image, execution }) => {
    const { createOfficeEngine } = await import("/dist/engine.js");
    const engine = await createOfficeEngine({ execution,
      formatPack: () => import("/dist/extended-formats.js").then(module => module.extendedFormatPack) });
    const doc = await engine.open(new Uint8Array(bytes));
    const objects = await doc.listObjects();
    const frames = [];
    const failures = [];
    const check = (condition, label) => { if (!condition) failures.push(label); };
    try {
      for (const colorScheme of ["light", "dark", "light"]) {
        const frame = await doc.render({ unitIndex: 0, colorScheme, scale: 1, includeTextFragments: true, ...(sheet ? { viewport: { x: 0, y: 0, width: 320, height: 240 } } : {}) });
        const canvas = new OffscreenCanvas(frame.pixelWidth, frame.pixelHeight);
        const ctx = canvas.getContext("2d"); ctx.drawImage(frame.bitmap, 0, 0);
        frames.push({ data: ctx.getImageData(0, 0, canvas.width, canvas.height).data, width: canvas.width, fragments: frame.textFragments });
        frame.bitmap.close();
      }
      let invalidColorScheme = false;
      try { await doc.render({ unitIndex: 0, colorScheme: "sepia" }); }
      catch (error) { invalidColorScheme = error.code === "INVALID_RENDER_REQUEST"; }
      check(invalidColorScheme, "unknown color scheme is rejected");
      const [light, dark, restored] = frames;
      const paper = (light.data.length / 4 - light.width * 5 - 5) * 4;
      check(light.data[paper] === 255 && dark.data[paper] === 26, "default paper changes to dark");
      check(light.data.every((value, index) => value === restored.data[index]), "light render and cache restore exact original pixels");
      for (const label of ["Auto ink", "Explicit red", "Explicit black", ...(format === "rtf" ? [] : ["White box"])]) {
        const object = objects.find(object => object.text === label);
        const fragment = light.fragments.find(fragment => fragment.objectId === object?.id && fragment.text.trim());
        check(fragment !== undefined, `${label} reaches real text layout`);
        if (!fragment) continue;
        let pixels = 0, correct = 0;
        // Glyph edges blend with the new paper; compare fully covered authored-color pixels.
        for (let y = Math.max(0, Math.floor(fragment.transform.f)); y < Math.ceil(fragment.transform.f + fragment.height); y++) {
          for (let x = Math.max(0, Math.floor(fragment.transform.e)); x < Math.ceil(fragment.transform.e + fragment.width); x++) {
            const i = (y * light.width + x) * 4;
            const isInk = label === "Explicit red" ? light.data[i] === 255 && light.data[i + 1] === 0 && light.data[i + 2] === 0
              : light.data[i] === 0 && light.data[i + 1] === 0 && light.data[i + 2] === 0;
            if (!isInk) continue;
            pixels++;
            if (label === "Auto ink" ? dark.data[i] > 200 && dark.data[i + 1] > 200
              : [0, 1, 2].every(channel => light.data[i + channel] === dark.data[i + channel])) correct++;
          }
        }
        check(pixels > 10 && correct / pixels > .98, `${label} follows its color semantics (${correct}/${pixels})`);
      }
      const images = objects.filter(object => object.type === "image");
      if (image) check(images.length > 0, "real embedded image is parsed");
      for (const image of images) {
        let pixels = 0, correct = 0;
        const { x, y, width, height } = image.bounds;
        for (let py = Math.ceil(y) + 1; py < Math.min(light.data.length / 4 / light.width, Math.floor(y + height) - 1); py++) {
          for (let px = Math.ceil(x) + 1; px < Math.min(light.width, Math.floor(x + width) - 1); px++) {
            const i = (py * light.width + px) * 4;
            if (Math.max(...light.data.slice(i, i + 3)) - Math.min(...light.data.slice(i, i + 3)) < 50
              && Math.max(...light.data.slice(i, i + 3)) !== 0) continue;
            pixels++;
            if ([0, 1, 2].every(channel => light.data[i + channel] === dark.data[i + channel])) correct++;
          }
        }
        check(pixels > 10 && correct / pixels > .98, `embedded image preserves colors (${correct}/${pixels})`);
      }
      // Explicit custom media and ordinary print/export renders keep original ink and paper.
      const original = await doc.render({ unitIndex: 0, background: "#ffffff", ...(sheet ? { viewport: { x: 0, y: 0, width: 320, height: 240 } } : {}) });
      const originalCanvas = new OffscreenCanvas(original.pixelWidth, original.pixelHeight);
      const originalContext = originalCanvas.getContext("2d"); originalContext.drawImage(original.bitmap, 0, 0);
      const originalData = originalContext.getImageData(0, 0, originalCanvas.width, originalCanvas.height).data;
      check(originalData.every((value, index) => value === light.data[index]), "ordinary print/export request preserves original pixels");
      original.bitmap.close();
      return failures;
    } finally { doc.close(); engine.close(); }
  }, { ...testCase, execution });
  assert.deepEqual(checks, [], `${testCase.format}/${execution}: semantic reading colors`);
  }
}

async function verifyChartColors(page) {
  const result = await page.evaluate(async () => {
    const { createOfficeEngine } = await import("/dist/engine.js");
    const engine = await createOfficeEngine({ execution: "worker" });
    const doc = await engine.open(await (await fetch("/tests/fixtures/chart-empty-title-original.xlsx")).arrayBuffer());
    try {
      const drawings = (await doc.listObjects({ unitIndex: 0 })).filter(object => object.source.kind === "drawing");
      if (!drawings.length) return "real spreadsheet chart was not parsed";
      const x = Math.min(...drawings.map(object => object.bounds.x));
      const y = Math.min(...drawings.map(object => object.bounds.y));
      const width = Math.max(...drawings.map(object => object.bounds.x + object.bounds.width)) - x;
      const height = Math.max(...drawings.map(object => object.bounds.y + object.bounds.height)) - y;
      const pixels = [];
      for (const colorScheme of ["light", "dark"]) {
        const frame = await doc.render({ unitIndex: 0, colorScheme, viewport: { x, y, width, height } });
        const canvas = new OffscreenCanvas(frame.pixelWidth, frame.pixelHeight);
        const ctx = canvas.getContext("2d"); ctx.drawImage(frame.bitmap, 0, 0);
        pixels.push(ctx.getImageData(0, 0, canvas.width, canvas.height).data); frame.bitmap.close();
      }
      // Fractional outer borders blend with the reading medium; chart contents keep their pixels.
      for (let py = 2; py < Math.ceil(height) - 2; py++) for (let px = 2; px < Math.ceil(width) - 2; px++) {
        const i = (py * Math.ceil(width) + px) * 4;
        if ([0, 1, 2, 3].some(channel => pixels[0][i + channel] !== pixels[1][i + channel])) return "spreadsheet chart pixels changed";
      }
      return "";
    } finally { doc.close(); engine.close(); }
  });
  assert.equal(result, "", "charts retain authored colors");
}

const server = spawn(process.execPath, ["scripts/serve.mjs", "--port", "0"], { stdio: ["ignore", "pipe", "inherit"] });
try {
  const url = await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error("Server startup timeout")), 10_000);
    createInterface({ input: server.stdout }).on("line", (line) => {
      const match = line.match(/https?:\/\/[^\s]+/u);
      if (match) { clearTimeout(timer); resolve(match[0]); }
    });
  });
  for (const [name, engine] of Object.entries({ chromium, webkit })) {
    if (process.env.DOCVIEWKIT_BROWSERS && !process.env.DOCVIEWKIT_BROWSERS.split(",").includes(name)) continue;
    console.log(`${name}: starting`);
    const browser = await engine.launch({ timeout: 20_000 });
    try {
      const page = await browser.newPage({ viewport: { width: 1440, height: 900 }, colorScheme: "light", reducedMotion: "reduce" });
      page.setDefaultTimeout(20_000);
      const errors = [];
      page.on("pageerror", (error) => errors.push(error.message));
      page.on("console", (message) => { if (message.type() === "error") errors.push(message.text()); });
      await page.goto(`${url}examples/viewer.html?fixture=visual-baseline.pptx&locale=zh-CN`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
      assert.equal(await page.title(), "DocViewKit Viewer");
      await verifySemanticColors(page);
      await verifyChartColors(page);
      const viewer = page.locator("docviewkit-viewer");
      const button = viewer.locator('[data-action="theme"]');
      assert.equal(await button.count(), 1, `${name}: missing theme control`);
      const theme = () => viewer.evaluate((v) => v.state.theme);
      const pixels = () => viewer.locator(".stage canvas").first().evaluate((c) => c.toDataURL());
      const before = await pixels();
      const surface = viewer.locator(".stage canvas").first();
      const light = await contentScreenshot(page, surface);
      await page.screenshot({ path: `/tmp/viewer-theme-${name}-light.png` });
      await page.emulateMedia({ colorScheme: "dark" });
      await page.waitForFunction(() => document.querySelector("#viewer").state.theme === "dark");
      await button.click(); // auto -> light
      assert.equal(await theme(), "light");
      await button.focus();
      await page.keyboard.press("Enter"); // light -> dark
      assert.equal(await theme(), "dark");
      assert.match(await button.getAttribute("aria-label"), /深色/);
      assert.equal(await viewer.evaluate((v) => getComputedStyle(v).getPropertyValue("--dv-accent").trim()), "#6aa7ff");
      await page.waitForFunction(() => getComputedStyle(document.querySelector("#viewer").shadowRoot.querySelector(".toolbar")).backgroundColor === "rgb(34, 37, 43)");
      assert.ok(await pixels() === before, "display theme must preserve source pixels");
      await page.screenshot({ path: `/tmp/viewer-theme-${name}-dark.png` });
      assert.equal(await surface.evaluate((c) => getComputedStyle(c).filter), "none", `${name}: slides must preserve authored colors in dark mode`);
      await page.emulateMedia({ media: "print" });
      await page.waitForFunction(() => getComputedStyle(document.querySelector("#viewer").shadowRoot.querySelector(".surface")).filter === "none");
      assert.equal(await surface.evaluate((c) => getComputedStyle(c).filter), "none", "printing must retain authored colors");
      await page.emulateMedia({ media: "screen" });
      await page.emulateMedia({ colorScheme: "light" });
      assert.equal(await theme(), "dark", "explicit theme overrides system");
      for (const width of [1440, 930, 720, 640, 601, 600, 541, 540, 480, 375, 320]) {
        await page.setViewportSize({ width, height: 900 });
        await page.evaluate(() => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))));
        assert.equal(await button.isVisible(), true);
        const overlap = await viewer.evaluate((v) => {
          const toolbar = v.shadowRoot.querySelector(".toolbar").getBoundingClientRect();
          const controls = [...v.shadowRoot.querySelectorAll(".toolbar button, .toolbar .zoom-value")]
            .map((e) => e.getBoundingClientRect()).filter((r) => r.width > 0);
          return controls.some((r, i) => r.left < toolbar.left || r.right > toolbar.right
            || controls.slice(i + 1).some((s) => r.left < s.right && s.left < r.right));
        });
        assert.equal(overlap, false, `${name}: toolbar overlap at ${width}px`);
      }
      await page.setViewportSize({ width: 375, height: 812 });
      await page.waitForTimeout(500); // Let the resized document frame finish rendering before visual capture.
      await page.screenshot({ path: `/tmp/viewer-theme-${name}-mobile.png` });
      await button.click(); // dark -> auto
      assert.equal(await theme(), "light");
      await page.emulateMedia({ colorScheme: "dark" });
      await page.waitForFunction(() => document.querySelector("#viewer").state.theme === "dark");
      await viewer.evaluate((v) => { v.removeAttribute("theme"); v.config = { ...v.config, theme: "light", locale: "en" }; });
      assert.equal(await theme(), "light");
      assert.match(await button.getAttribute("aria-label"), /Light/);
      await button.click();
      assert.equal(await theme(), "dark");
      await viewer.evaluate((v) => v.close());
      assert.equal(await button.isEnabled(), true);
      await button.click();
      assert.equal(await theme(), "dark"); // auto, OS dark, even without a document
      await page.setViewportSize({ width: 1440, height: 900 });
      for (const fixture of ["word-footnote-page-break.docx", "railway-invoice-gbk.pdf", "basicspreadsheet-custom-styles.xlsx"]) {
        console.log(`${name}: ${fixture}`);
        await page.goto(`${url}examples/viewer.html?fixture=${fixture}&theme=light`);
        await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
        if (fixture.endsWith(".docx")) {
          await viewer.evaluate((v) => v.config = { ...v.config, navigation: "visible" });
          await viewer.locator('[data-action="navigation-thumbnails"]').click();
          const thumbnail = viewer.locator('.unit-button[data-unit-index="0"] canvas');
          await page.waitForFunction(() => document.querySelector("#viewer").shadowRoot
            .querySelector('.unit-button[data-unit-index="0"] canvas')?.dataset.rendered === "true");
          const lightThumbnail = await thumbnail.screenshot({ path: `/tmp/viewer-theme-${name}-thumbnail-light.png` });
          await setReadingTheme(page, viewer, "dark", true);
          await page.waitForFunction(() => {
            const canvas = document.querySelector("#viewer").shadowRoot.querySelector('.unit-button[data-unit-index="0"] canvas');
            return canvas?.dataset.rendered === "true" && canvas.getContext("2d").getImageData(0, 0, 1, 1).data[0] === 26;
          });
          const darkThumbnail = await thumbnail.screenshot({ path: `/tmp/viewer-theme-${name}-thumbnail-dark.png` });
          const thumbnailInk = await thumbnail.evaluate(canvas => {
            const { data } = canvas.getContext("2d").getImageData(0, 0, canvas.width, canvas.height);
            let paper = 0, ink = 0;
            for (let i = 0; i < data.length; i += 4) {
              if (data[i] === 26 && data[i + 1] === 26 && data[i + 2] === 26) paper++;
              else if (data[i] > 40 && data[i] < 240) ink++;
            }
            return { paper, ink };
          });
          // Tiny glyph hinting differs from resampling a page, especially in WebKit.
          assert.ok(thumbnailInk.paper > 100 && thumbnailInk.ink > 20, `${name}: thumbnail must have dark paper and readable ink`);
          await page.screenshot({ path: `/tmp/viewer-theme-${name}-document.png` });
          await setReadingTheme(page, viewer, "light", true);
        }
        for (let pass = 0; pass < 2; pass++) {
          const mode = await viewer.evaluate((v) => v.state.pageMode);
          const surface = viewer.locator(mode === "continuous"
            ? '.continuous-page[data-rendered="true"] canvas' : ".stage .surface").first();
          const light = await contentScreenshot(page, surface);
          const source = await surface.evaluate((c) => c.toDataURL());
          const adapted = !fixture.endsWith(".pdf");
          await setReadingTheme(page, viewer, "dark", adapted);
          if (adapted) await assertDarkContent(page, light, await contentScreenshot(page, surface), `${name}: ${fixture}/${mode}`, fixture.endsWith(".xlsx"));
          assert.equal(await surface.evaluate((c) => c.toDataURL()) === source, !adapted, "semantic content changes only reading defaults");
          const filters = await viewer.evaluate((v) => [...v.shadowRoot.querySelectorAll(
            ".stage canvas, .sheet-tile, .continuous-page canvas, .thumbnail canvas")]
            .map((c) => getComputedStyle(c).filter));
          assert.ok(filters.every((filter) => filter === "none"), `${name}: whole-canvas filter in ${fixture}`);
          if (fixture.endsWith(".xlsx")) {
            assert.equal(await viewer.locator(".stage-wrap").evaluate((e) => getComputedStyle(e).backgroundColor), "rgb(26, 26, 26)");
            assert.equal(await viewer.locator(".sheet-header-corner").evaluate((e) => getComputedStyle(e).backgroundColor), "rgb(29, 32, 37)");
          }
          await setReadingTheme(page, viewer, "light", adapted);
          assert.ok((await contentScreenshot(page, surface)).equals(light), `${name}: light theme must restore ${fixture}`);
          if (mode !== "continuous") break;
          await viewer.locator('[data-action="page-mode"]:visible').click();
          await page.waitForFunction(() => {
            const viewer = document.querySelector("#viewer"), surface = viewer.shadowRoot.querySelector(".surface");
            return viewer.state.pageMode === "single" && surface.width > 1
              && surface.dataset.preview !== "true" && surface.getBoundingClientRect().width > 1;
          });
        }
      }
      assert.deepEqual(errors, []);
      console.log(`${name}: semantic reading colors, authored image/slide/PDF colors, original rendering, page modes, sheet tiles/headers, theme cycle, system/config, keyboard, 11 widths, console PASS`);
    } finally { await browser.close(); }
  }
} finally { server.kill(); }
