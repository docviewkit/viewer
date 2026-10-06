import { createOfficeEngine } from "../dist/engine.js";

const parameters = new URLSearchParams(location.search);
const fixture = parameters.get("fixture") ?? "visual-baseline.pptx";
const unitIndex = Number(parameters.get("unit") ?? "0");
const scale = Number(parameters.get("scale") ?? "1");
const renderAll = parameters.get("all") === "1";
const exportAll = parameters.get("exportAll") === "1";
const fontManifest = parameters.get("fontManifest");
const sheetRange = parameters.get("sheetRange");
const password = parameters.has("password") ? parameters.get("password") : undefined;
const surface = document.querySelector("#surface");
const status = document.querySelector("#status");
const candidateExport = document.querySelector("#candidate-export");
const candidateExports = document.querySelector("#candidate-exports");

const FONT_STYLES = new Set(["normal", "italic", "oblique"]);
const FONT_STRETCHES = new Set([
  "ultra-condensed",
  "extra-condensed",
  "condensed",
  "semi-condensed",
  "normal",
  "semi-expanded",
  "expanded",
  "extra-expanded",
  "ultra-expanded",
]);

function isSafeFileName(value, extensions) {
  return typeof value === "string"
    && value.length > 0
    && value.length <= 160
    && !value.startsWith(".")
    && !value.includes("..")
    && !/[\\/\0]/u.test(value)
    && /^[A-Za-z0-9][A-Za-z0-9._ -]*$/u.test(value)
    && extensions.some((extension) => value.toLowerCase().endsWith(extension));
}

function isSafeFixturePath(value) {
  return typeof value === "string"
    && value.length > 0
    && value.length <= 320
    && !value.startsWith("/")
    && !value.includes("\\")
    && !value.includes("\0")
    && value.split("/").every((segment) => segment !== "" && segment !== "." && segment !== ".."
      && /^[A-Za-z0-9][A-Za-z0-9._ #()-]*$/u.test(segment))
    && /\.(?:pptx|pptm|ppsx|ppsm|potx|potm|ppt|odp|otp|fodp|dps|xlsx|xlsm|xltx|xltm|xls|ods|ots|fods|et|docx|docm|dotx|dotm|doc|odt|ott|wps|rtf|key|pages|numbers|pdf|xps|oxps)$/iu.test(value);
}

function encodePath(value) {
  return value.split("/").map(encodeURIComponent).join("/");
}

function columnIndex(label) {
  let index = 0;
  for (const character of label) index = index * 26 + character.charCodeAt(0) - 64;
  return index - 1;
}

function axisOffset(axis, index) {
  let offset = index * axis.defaultSize;
  for (const span of axis.spans) {
    if (span.start >= index) break;
    const covered = Math.min(index - 1, span.end) - span.start + 1;
    if (covered > 0) offset += covered * (span.size - axis.defaultSize);
  }
  return offset;
}

function viewportForSheetRange(unit, value) {
  const match = /^([A-Z]{1,3})([1-9][0-9]{0,6}):([A-Z]{1,3})([1-9][0-9]{0,6})$/u.exec(value ?? "");
  if (match === null) throw new Error("sheetRange must be an uppercase bounded A1 range such as A1:H40");
  const startColumn = columnIndex(match[1]);
  const startRow = Number(match[2]) - 1;
  const endColumn = columnIndex(match[3]);
  const endRow = Number(match[4]) - 1;
  if (startColumn > endColumn || startRow > endRow || endColumn >= 16_384 || endRow >= 1_048_576) {
    throw new Error("sheetRange is outside the spreadsheet grid or has reversed bounds");
  }
  const x = axisOffset(unit.columnAxis, startColumn);
  const y = axisOffset(unit.rowAxis, startRow);
  const right = axisOffset(unit.columnAxis, endColumn + 1);
  const bottom = axisOffset(unit.rowAxis, endRow + 1);
  const viewport = { x, y, width: right - x, height: bottom - y };
  if (viewport.width <= 0 || viewport.height <= 0) throw new Error("sheetRange resolves to an empty viewport");
  return viewport;
}

const TEXT_ATTENTION_PADDING_PX = 4;
const MAX_TEXT_ATTENTION_CANDIDATES = 4_096;
const MAX_TEXT_ATTENTION_REGIONS = 256;
const MAX_TEXT_ATTENTION_PIXELS = 250_000;

function textAttentionRegions(objects, viewport, renderScale, width, height, unitIndex) {
  const candidates = [];
  let sourceTextObjectCount = 0;
  let visibleTextObjectCount = 0;
  let omittedCandidateCount = 0;
  for (const object of objects) {
    if (typeof object.text !== "string"
      || object.text.trim() === ""
      || ![
        object.bounds.x,
        object.bounds.y,
        object.bounds.width,
        object.bounds.height,
      ].every(Number.isFinite)
      || object.bounds.width <= 0
      || object.bounds.height <= 0) continue;
    sourceTextObjectCount += 1;
    const candidate = {
      firstId: object.id,
      left: (object.bounds.x - (viewport?.x ?? 0)) * renderScale,
      top: (object.bounds.y - (viewport?.y ?? 0)) * renderScale,
      right: (object.bounds.x + object.bounds.width - (viewport?.x ?? 0)) * renderScale,
      bottom: (object.bounds.y + object.bounds.height - (viewport?.y ?? 0)) * renderScale,
    };
    if (![candidate.left, candidate.top, candidate.right, candidate.bottom].every(Number.isFinite)
      || candidate.right <= 0
      || candidate.bottom <= 0
      || candidate.left >= width
      || candidate.top >= height) continue;
    visibleTextObjectCount += 1;
    if (candidates.length === MAX_TEXT_ATTENTION_CANDIDATES) {
      omittedCandidateCount += 1;
      continue;
    }
    candidates.push(candidate);
  }
  candidates.sort((left, right) => left.top - right.top || left.left - right.left);
  const lines = [];
  for (const candidate of candidates) {
    let line;
    for (let index = lines.length - 1; index >= Math.max(0, lines.length - 8); index -= 1) {
      const current = lines[index];
      const overlap = Math.min(current.bottom, candidate.bottom) - Math.max(current.top, candidate.top);
      const minHeight = Math.min(current.bottom - current.top, candidate.bottom - candidate.top);
      const gap = Math.max(0, candidate.left - current.right, current.left - candidate.right);
      if (overlap >= minHeight * 0.6 && gap <= Math.max(32, minHeight * 2)) {
        line = current;
        break;
      }
    }
    if (line === undefined) {
      lines.push({ ...candidate });
    } else {
      line.left = Math.min(line.left, candidate.left);
      line.top = Math.min(line.top, candidate.top);
      line.right = Math.max(line.right, candidate.right);
      line.bottom = Math.max(line.bottom, candidate.bottom);
    }
  }
  const regions = [];
  let pixelCount = 0;
  let omittedRegionCount = 0;
  for (const line of lines) {
    const left = Math.max(0, Math.floor(line.left) - TEXT_ATTENTION_PADDING_PX);
    const top = Math.max(0, Math.floor(line.top) - TEXT_ATTENTION_PADDING_PX);
    const right = Math.min(width, Math.ceil(line.right) + TEXT_ATTENTION_PADDING_PX);
    const bottom = Math.min(height, Math.ceil(line.bottom) + TEXT_ATTENTION_PADDING_PX);
    if (right <= left || bottom <= top) continue;
    const pixels = (right - left) * (bottom - top);
    if (regions.length === MAX_TEXT_ATTENTION_REGIONS
      || pixelCount + pixels > MAX_TEXT_ATTENTION_PIXELS) {
      omittedRegionCount += 1;
      continue;
    }
    const sourceKey = `text-region:${line.firstId}`;
    regions.push({
      id: sourceKey,
      sourceKey,
      type: "text-region",
      unitIndex,
      bounds: { x: left, y: top, width: right - left, height: bottom - top },
    });
    pixelCount += pixels;
  }
  const bounded = omittedCandidateCount !== 0 || omittedRegionCount !== 0;
  return {
    objects: regions,
    attention: {
      status: bounded ? "bounded" : "complete",
      padding: TEXT_ATTENTION_PADDING_PX,
      sourceTextObjectCount,
      visibleTextObjectCount,
      candidateCount: candidates.length,
      mergedRegionCount: lines.length,
      regionCount: regions.length,
      pixelCount,
      omittedCandidateCount,
      omittedRegionCount,
      limits: {
        candidates: MAX_TEXT_ATTENTION_CANDIDATES,
        regions: MAX_TEXT_ATTENTION_REGIONS,
        pixels: MAX_TEXT_ATTENTION_PIXELS,
      },
    },
  };
}

async function sha256(bytes) {
  const digest = await crypto.subtle.digest("SHA-256", bytes);
  return `sha256:${[...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, "0")).join("")}`;
}

async function loadManifestFonts(manifestFile) {
  if (!isSafeFileName(manifestFile, [".json"])) {
    throw new Error("fontManifest must be a safe local JSON filename");
  }
  const manifestResponse = await fetch(`/tests/fonts/${encodeURIComponent(manifestFile)}`, { cache: "no-store" });
  if (!manifestResponse.ok) throw new Error(`font manifest fetch failed with HTTP ${manifestResponse.status}`);
  const manifest = await manifestResponse.json();
  if (manifest === null || typeof manifest !== "object" || !Array.isArray(manifest.fonts)) {
    throw new Error("font manifest must contain a fonts array");
  }
  if (manifest.fonts.length === 0 || manifest.fonts.length > 256) {
    throw new Error("font manifest must contain from 1 through 256 fonts");
  }
  return Promise.all(manifest.fonts.map(async (entry, index) => {
    if (entry === null || typeof entry !== "object") throw new Error(`fonts[${index}] must be an object`);
    if (typeof entry.family !== "string" || entry.family.trim() !== entry.family
      || entry.family.length === 0 || entry.family.length > 128) {
      throw new Error(`fonts[${index}].family is invalid`);
    }
    if (!isSafeFileName(entry.file, [".otf", ".ttf", ".woff", ".woff2"])) {
      throw new Error(`fonts[${index}].file must be a safe local font filename`);
    }
    const style = entry.style ?? "normal";
    if (!FONT_STYLES.has(style)) throw new Error(`fonts[${index}].style is invalid`);
    const weight = entry.weight ?? 400;
    if (!Number.isInteger(weight) || weight < 1 || weight > 1000) {
      throw new Error(`fonts[${index}].weight must be an integer from 1 through 1000`);
    }
    const stretch = entry.stretch ?? "normal";
    if (!FONT_STRETCHES.has(stretch)) throw new Error(`fonts[${index}].stretch is invalid`);
    const fontResponse = await fetch(`/tests/fonts/${encodeURIComponent(entry.file)}`, { cache: "no-store" });
    if (!fontResponse.ok) throw new Error(`fonts[${index}] fetch failed with HTTP ${fontResponse.status}`);
    return {
      family: entry.family,
      bytes: await fontResponse.arrayBuffer(),
      style,
      weight,
      stretch,
    };
  }));
}

function fail(cause) {
  const message = cause instanceof Error ? cause.stack ?? cause.message : String(cause);
  document.body.dataset.state = "error";
  status.textContent = message;
  document.title = "FAIL - OfficeViewer visual harness";
  globalThis.visualHarness = Object.freeze({ state: "error", message });
}

let engine;
let officeDocument;
let retainedDocument = false;
try {
  if (!isSafeFixturePath(fixture)) {
    throw new Error("fixture must be a supported local Office filename");
  }
  if (!Number.isInteger(unitIndex) || unitIndex < 0) throw new Error("unit must be a non-negative integer");
  if (!Number.isFinite(scale) || scale <= 0 || scale > 4) throw new Error("scale must be in (0, 4]");
  if (exportAll && !renderAll) throw new Error("exportAll requires all=1");
  const response = await fetch(`/tests/fixtures/${encodePath(fixture)}`, { cache: "no-store" });
  if (!response.ok) throw new Error(`fixture fetch failed with HTTP ${response.status}`);
  const bytes = await response.arrayBuffer();
  const fixtureSha256 = await sha256(bytes);
  let fonts;
  if (fontManifest === null) {
    const fontResponse = await fetch("/tests/fonts/NotoSansHans-Regular.otf", { cache: "no-store" });
    if (!fontResponse.ok) throw new Error(`font fetch failed with HTTP ${fontResponse.status}`);
    fonts = [{ family: "Noto Sans S Chinese", bytes: await fontResponse.arrayBuffer() }];
  } else {
    fonts = await loadManifestFonts(fontManifest);
  }
  engine = await createOfficeEngine({
    fonts,
    limits: {
      renderPixels: 256_000_000,
      imagePixels: 256_000_000,
      totalImagePixels: 512_000_000,
    },
    formatPack: () => import("../dist/extended-formats.js").then(({ extendedFormatPack }) => extendedFormatPack),
  });
  if (password !== undefined && password.length > 1_024) throw new Error("password is too long");
  officeDocument = await engine.open(bytes, {
    timeoutMs: 60_000,
    ...(password === undefined ? {} : { password }),
  });
  if (unitIndex >= officeDocument.info.units.length) throw new Error("unit exceeds the document unit count");
  const selectedUnit = officeDocument.info.units[unitIndex];
  if (sheetRange !== null && selectedUnit.type !== "sheet") throw new Error("sheetRange is valid only for a sheet unit");
  if (selectedUnit.type === "sheet" && sheetRange === null) {
    throw new Error("sheet fixtures require a deterministic sheetRange such as A1:H40");
  }
  const unitIndices = renderAll ? officeDocument.info.units.map(({ index }) => index) : [unitIndex];
  const renderedUnits = [];
  const exportedUnits = [];
  let nonWhitePixelCount = 0;
  const renderUnit = async (index) => {
    if (!Number.isInteger(index) || index < 0 || index >= officeDocument.info.units.length) {
      throw new RangeError("renderUnit index is outside the document");
    }
    const unit = officeDocument.info.units[index];
    const viewport = unit.type === "sheet" ? viewportForSheetRange(unit, sheetRange) : undefined;
    const documentObjects = await officeDocument.listObjects({ unitIndex: index, textOnly: true });
    const frame = await officeDocument.render({
      unitIndex: index,
      scale,
      pixelRatio: 1,
      background: "#ffffff",
      ...(viewport === undefined ? {} : { viewport }),
    });
    try {
      const rendered = {
        index,
        unitType: unit.type,
        unitName: unit.name,
        width: frame.pixelWidth,
        height: frame.pixelHeight,
        renderedObjectCount: frame.renderedObjectCount,
        ...(viewport === undefined ? {} : { viewport, sheetRange }),
      };
      surface.width = frame.pixelWidth;
      surface.height = frame.pixelHeight;
      surface.style.width = `${frame.pixelWidth}px`;
      surface.style.height = `${frame.pixelHeight}px`;
      const context = surface.getContext("2d", { alpha: false });
      if (context === null) throw new Error("Canvas 2D is unavailable");
      context.drawImage(frame.bitmap, 0, 0);
      const pixels = context.getImageData(0, 0, surface.width, surface.height).data;
      let unitNonWhitePixelCount = 0;
      for (let offset = 0; offset < pixels.length; offset += 4) {
        if (pixels[offset] < 250 || pixels[offset + 1] < 250 || pixels[offset + 2] < 250) {
          unitNonWhitePixelCount += 1;
        }
      }
      const dataUrl = surface.toDataURL("image/png");
      candidateExport.href = dataUrl;
      const attentionEvidence = textAttentionRegions(
        documentObjects,
        viewport,
        scale,
        surface.width,
        surface.height,
        index,
      );
      return {
        rendered,
        dataUrl,
        nonWhitePixelCount: unitNonWhitePixelCount,
        objects: attentionEvidence.objects,
        attention: attentionEvidence.attention,
        diagnostics: [...officeDocument.diagnostics(), ...frame.diagnostics],
      };
    } finally {
      frame.bitmap.close();
    }
  };
  for (const index of unitIndices) {
    const rendered = await renderUnit(index);
    renderedUnits.push(rendered.rendered);
    nonWhitePixelCount += rendered.nonWhitePixelCount;
    if (exportAll) {
      exportedUnits.push({ index, dataUrl: rendered.dataUrl });
      const exportLink = document.createElement("a");
      exportLink.dataset.unitIndex = String(index);
      exportLink.download = `candidate-${String(index + 1).padStart(2, "0")}.png`;
      exportLink.href = rendered.dataUrl;
      candidateExports.append(exportLink);
    }
  }
  document.body.dataset.state = "pass";
  document.body.dataset.unitCount = String(officeDocument.info.units.length);
  document.body.dataset.renderedUnitCount = String(renderedUnits.length);
  document.body.dataset.emptyUnitCount = String(renderedUnits.filter(({ renderedObjectCount }) => renderedObjectCount === 0).length);
  document.body.dataset.nonWhitePixelCount = String(nonWhitePixelCount);
  status.dataset.diagnostics = JSON.stringify(officeDocument.diagnostics().slice(0, 100));
  status.textContent = `PASS ${officeDocument.info.format} rendered=${renderedUnits.length}/${officeDocument.info.units.length} unit=${unitIndex} ${surface.width}x${surface.height}`;
  document.title = "PASS - OfficeViewer visual harness";
  globalThis.visualHarness = Object.freeze({
    state: "pass",
    fixture,
    fixtureSha256,
    format: officeDocument.info.format,
    unitIndex,
    unitType: selectedUnit.type,
    unitName: selectedUnit.name,
    unitCount: officeDocument.info.units.length,
    renderedUnits,
    width: surface.width,
    height: surface.height,
    diagnostics: officeDocument.diagnostics(),
    renderUnit: async (index) => {
      const rendered = await renderUnit(index);
      return Object.freeze({
        ...rendered.rendered,
        dataUrl: rendered.dataUrl,
        nonWhitePixelCount: rendered.nonWhitePixelCount,
        objects: rendered.objects,
        attention: rendered.attention,
        diagnostics: rendered.diagnostics,
      });
    },
    ...(selectedUnit.type === "sheet" ? {
      sheetRange,
      viewport: viewportForSheetRange(selectedUnit, sheetRange),
    } : {}),
    exportedUnits,
  });
  retainedDocument = true;
  addEventListener("pagehide", () => {
    officeDocument?.close();
    engine?.close();
    officeDocument = undefined;
    engine = undefined;
  }, { once: true });
} catch (cause) {
  fail(cause);
} finally {
  if (!retainedDocument) {
    officeDocument?.close();
    engine?.close();
  }
}
