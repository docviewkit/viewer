import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const runner = await readFile(
  new URL("../scripts/run-pdfjs-reference-corpus.mjs", import.meta.url),
  "utf8",
);

test("PDF.js reference corpus runner cannot silently omit PDFs", () => {
  assert.match(runner, /entry\.name\.toLowerCase\(\)\.endsWith\("\.pdf"\)/u);
  assert.match(runner, /discoveredPdfs:\s*allFiles\.length/u);
  assert.match(runner, /selectedPdfs:\s*entries\.length/u);
  assert.match(runner, /pendingPdfs:\s*entries\.length\s*-\s*files_\.length/u);
  assert.match(runner, /Path escapes the PDF root/u);
});

test("PDF.js reference corpus runner pins reproducible rendering semantics", () => {
  assert.match(runner, /pdfjsCommit:\s*pdfjsCommit\(\)/u);
  assert.match(runner, /scale:\s*1/u);
  assert.match(runner, /devicePixelRatio:\s*1/u);
  assert.match(runner, /background:\s*"#ffffff"/u);
  assert.match(runner, /isEvalSupported:\s*false/u);
  assert.match(runner, /host:\s*"chromium"/u);
  assert.match(runner, /options\.host === "chromium"/u);
  assert.match(runner, /oracleHost:\s*options\.host/u);
  assert.match(runner, /Math\.sumPrecise \?\?=/u);
  assert.match(runner, /browserLikeCanvasContext/u);
  assert.match(runner, /if \(property === "font"\) return true/u);
  assert.match(runner, /chromiumPdfjsRender/u);
  assert.match(runner, /oracleHost:\s*"chromium"/u);
  assert.match(runner, /enableXfa:\s*true/u);
  assert.match(runner, /await pdfPage\.getXfa\(\)/u);
  assert.match(runner, /XfaLayer\.render\(\{/u);
  assert.match(runner, /locator\("#pdfjs-xfa-capture"\)\.screenshot/u);
});

test("PDF.js failures are isolated and routed to Acrobat rather than accepted", () => {
  assert.match(runner, /--max-old-space-size=1024/u);
  assert.match(runner, /child\.kill\("SIGKILL"\)/u);
  assert.match(runner, /code:\s*signal === "SIGKILL" \? "PDFJS_TIMEOUT"/u);
  assert.match(runner, /status:\s*"pdfjs-failed"/u);
  assert.match(runner, /oracle:\s*"acrobat-required"/u);
  assert.match(runner, /writeJsonAtomic\(ledgerPath/u);
  assert.match(runner, /standardFixturePasswords = \["password"\]/u);
});
