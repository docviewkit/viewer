import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const html = await readFile(new URL("../examples/visual-harness.html", import.meta.url), "utf8");
const script = await readFile(new URL("../examples/visual-harness.js", import.meta.url), "utf8");
const server = await readFile(new URL("../scripts/serve.mjs", import.meta.url), "utf8");
const serverManager = await readFile(new URL("../scripts/test-server.sh", import.meta.url), "utf8");
const manifest = JSON.parse(await readFile(new URL("../package.json", import.meta.url), "utf8"));
const nativeCaptureRunner = await readFile(new URL("../scripts/capture-native-visual-suite.mjs", import.meta.url), "utf8");
const pdfPreviewRunner = await readFile(new URL("../scripts/run-pdf-preview-suite.mjs", import.meta.url), "utf8");
const pdfPreviewReference = await readFile(new URL("../scripts/render-pdf-preview-reference.swift", import.meta.url), "utf8");

test("visual harness renders a content-addressed local fixture without application chrome", () => {
  assert.match(html, /<canvas id="surface"/);
  assert.match(html, /id="candidate-export"/);
  assert.match(html, /id="candidate-exports"/);
  assert.doesNotMatch(html, /toolbar|commandbar|inspector/iu);
  assert.match(script, /new URLSearchParams\(location\.search\)/);
  assert.match(script, /document\.body\.dataset\.state\s*=\s*"pass"/);
  assert.match(script, /officeDocument\.render/);
  assert.match(script, /officeDocument\.listObjects\(\{ unitIndex: index, textOnly: true \}\)/);
  assert.match(script, /TEXT_ATTENTION_PADDING_PX\s*=\s*4/);
  assert.match(script, /MAX_TEXT_ATTENTION_CANDIDATES\s*=\s*4_096/);
  assert.match(script, /MAX_TEXT_ATTENTION_REGIONS\s*=\s*256/);
  assert.match(script, /MAX_TEXT_ATTENTION_PIXELS\s*=\s*250_000/);
  assert.match(script, /status:\s*bounded\s*\?\s*"bounded"\s*:\s*"complete"/);
  assert.doesNotMatch(script, /white\s*\/\s*total/);
  assert.match(script, /officeDocument\.info\.units\.map/);
  assert.match(script, /renderedUnits/);
  assert.match(script, /nonWhitePixelCount/);
  assert.match(script, /dataset\.diagnostics/);
  assert.match(script, /pixelRatio:\s*1/);
  assert.match(script, /surface\.toDataURL\("image\/png"\)/);
  assert.match(script, /exportAll requires all=1/);
  assert.match(script, /exportedUnits\.push\(\{ index, dataUrl: rendered\.dataUrl \}\)/);
  assert.match(script, /exportLink\.dataset\.unitIndex = String\(index\)/);
  assert.match(script, /parameters\.get\("fontManifest"\)/);
  assert.match(script, /parameters\.get\("sheetRange"\)/);
  assert.match(script, /parameters\.has\("password"\)/);
  assert.match(script, /password\.length > 1_024/);
  assert.match(script, /viewportForSheetRange/);
  assert.match(script, /fixtureSha256/);
  assert.match(script, /unitType:\s*selectedUnit\.type/);
  assert.match(script, /font manifest must contain a fonts array/);
  assert.match(script, /isSafeFixturePath/);
  assert.match(script, /encodePath\(fixture\)/);
  assert.match(script, /encodeURIComponent\(entry\.file\)/);
  assert.match(script, /entry\.family\.trim\(\) !== entry\.family/);
  assert.match(script, /FONT_STYLES\.has\(style\)/);
  assert.match(script, /weight < 1 \|\| weight > 1000/);
  assert.match(script, /FONT_STRETCHES\.has\(stretch\)/);
  assert.match(script, /formatPack:\s*\(\) => import\("\.\.\/dist\/extended-formats\.js"\)/);
  assert.match(script, /fontManifest === null/);
  assert.match(script, /NotoSansHans-Regular\.otf/);
  for (const extension of [
    "pptx", "ppt", "xlsx", "xls", "docx", "doc", "rtf", "key", "pages", "numbers",
    "odp", "fodp", "ods", "fods", "odt", "wps", "et", "dps", "pdf", "xps", "oxps",
  ]) {
    assert.match(script, new RegExp(`\\b${extension}\\b`, "u"));
  }
});

test("PDF Preview visual gate keeps the native oracle separate from OfficeViewer parsing", () => {
  assert.match(nativeCaptureRunner, /\(\?:DocViewKit Viewer\|Office Viewer Inspector\)/);
  assert.match(pdfPreviewRunner, /\(\?:DocViewKit Viewer\|Office Viewer Inspector\)/);
  assert.match(pdfPreviewRunner, /globalThis\.visualHarness\.renderUnit\(index\)/);
  assert.match(pdfPreviewRunner, /minTolerantPixelRatio:\s*0\.98/);
  assert.match(pdfPreviewRunner, /pixelChannel:\s*8/);
  assert.match(pdfPreviewRunner, /PDF_PIXEL_RADIUS\s*=\s*3/);
  assert.match(pdfPreviewRunner, /pixelRadius:\s*PDF_PIXEL_RADIUS/);
  assert.match(pdfPreviewRunner, /visual-object-region/);
  assert.match(pdfPreviewRunner, /minObjectRegionSimilarity:\s*0\.65/);
  assert.match(pdfPreviewRunner, /code === "FONT_LOAD_FAILED"/);
  assert.match(pdfPreviewRunner, /expectedObservationJson/);
  assert.match(pdfPreviewRunner, /attention\.status === "bounded"/);
  assert.match(pdfPreviewRunner, /attention\.padding < PDF_PIXEL_RADIUS/);
  assert.match(pdfPreviewRunner, /render-pdf-preview-reference\.swift/);
  assert.match(pdfPreviewReference, /PDFDocument\(url: input\)/);
  assert.match(pdfPreviewReference, /page\.draw\(with: \.cropBox, to: context\)/);
});

test("development server exposes only the configured fixture and QA-font directories needed by visual QA", () => {
  assert.match(server, /Location: "\/examples\/viewer\.html\?fixture=visual-baseline\.pptx"/);
  assert.match(server, /"dist\/viewer\.js"/);
  assert.match(server, /const fixtureRoot = resolve\(values\.get\("--fixture-root"\)/);
  assert.match(server, /const fontRoot = resolve\(values\.get\("--font-root"\)/);
  assert.match(server, /\["\/tests\/fixtures\/",\s*fixtureRoot\]/);
  assert.match(server, /\["\/tests\/fonts\/",\s*fontRoot\]/);
  assert.match(server, /node_modules\/@embedpdf\/fonts-sc\/fonts/);
  assert.match(server, /\["\.ttf",\s*"font\/ttf"\]/);
  assert.match(server, /\["\.woff2",\s*"font\/woff2"\]/);
  assert.match(server, /\["\.pptx",\s*"application\/vnd\.openxmlformats-officedocument\.presentationml\.presentation"\]/);
  assert.match(server, /\["\.docx",\s*"application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.document"\]/);
  assert.match(server, /\["\.xlsx",\s*"application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet"\]/);
  assert.match(server, /\["\.ppt",\s*"application\/vnd\.ms-powerpoint"\]/);
  assert.match(server, /\["\.doc",\s*"application\/msword"\]/);
  assert.match(server, /\["\.xls",\s*"application\/vnd\.ms-excel"\]/);
  assert.match(server, /\["\.key",\s*"application\/vnd\.apple\.keynote"\]/);
  assert.match(server, /\["\.pages",\s*"application\/vnd\.apple\.pages"\]/);
  assert.match(server, /\["\.numbers",\s*"application\/vnd\.apple\.numbers"\]/);
  assert.match(server, /\["\.wps",\s*"application\/vnd\.kingsoft\.writer"\]/);
  assert.match(server, /\["\.et",\s*"application\/vnd\.kingsoft\.spreadsheets"\]/);
  assert.match(server, /\["\.dps",\s*"application\/vnd\.kingsoft\.presentation"\]/);
  assert.match(server, /\["\.pdf",\s*"application\/pdf"\]/);
  assert.match(server, /media-src\s+'self'\s+blob:/);
  assert.match(server, /values\.get\("--host"\) \?\? "127\.0\.0\.1"/);
  assert.equal(manifest.scripts.inspect, "node scripts/serve.mjs --host 0.0.0.0");
  assert.match(serverManager, /VIEWER_HOST="\$\{VIEWER_HOST:-\$\{HOST:-0\.0\.0\.0\}\}"/);
  assert.match(server, /port < 0/);
  assert.match(server, /server\.address\(\)/);
});
