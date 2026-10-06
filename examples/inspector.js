import { createOfficeEngine } from "../dist/engine.js";
import { dominantDocumentFontSize } from "../dist/font.js";
import { sheetAutoFitColumnWidth, sheetPrintPages } from "../dist/render.js";
import { OfficeEngineError } from "../dist/types.js";

const MAX_INPUT_BYTES = 128 * 1024 * 1024;
const MAX_RENDER_PIXELS = 32_000_000;
const OPEN_TIMEOUT_MS = 60_000;
const MAX_VISIBLE_DIAGNOSTICS = 500;
const MAX_VISIBLE_TEXT = 4_000;
const MAX_TEXT_LAYER_OBJECTS = 20_000;
const MAX_DIAGNOSTIC_OBJECTS = 50_000;
const SHEET_SCROLL_THROTTLE_MS = 80;
const SHEET_VIEWPORT_OVERSCAN_RATIO = 0.5;
const SURFACE_PADDING = 28;
const SHEET_ROW_HEADER_WIDTH = 48;
const SHEET_COLUMN_HEADER_HEIGHT = 24;
const MIN_COLUMN_WIDTH = 24;
const MIN_ROW_HEIGHT = 12;
const THUMBNAIL_MAX_WIDTH = 180;
const THUMBNAIL_MAX_HEIGHT = 255;
const THUMBNAIL_CACHE_PIXELS = 12_000_000;
const CONTINUOUS_CACHE_PIXELS = 40_000_000;
const CONTINUOUS_SMOOTH_SCROLL_VIEWPORTS = 2;
const PRINT_RENDER_LONG_EDGE = 1600;

function element(id) {
  const value = document.getElementById(id);
  if (value === null) throw new Error(`Inspector element #${id} is missing`);
  return value;
}

const ui = {
  app: element("app"),
  clearButton: element("clear-button"),
  dropZone: element("drop-zone"),
  fileInput: element("file-input"),
  status: element("status"),
  fileName: element("file-name"),
  fileSize: element("file-size"),
  documentFormat: element("document-format"),
  documentKind: element("document-kind"),
  previousUnit: element("previous-unit"),
  unitSelect: element("unit-select"),
  nextUnit: element("next-unit"),
  searchForm: element("search-form"),
  searchInput: element("search-input"),
  searchPrevious: element("search-previous"),
  searchNext: element("search-next"),
  searchStatus: element("search-status"),
  fitView: element("fit-view"),
  continuousButton: element("continuous-button"),
  slideshowButton: element("slideshow-button"),
  printButton: element("print-button"),
  fullscreenButton: element("fullscreen-button"),
  unitNavigationTitle: element("unit-navigation-title"),
  unitCount: element("unit-count"),
  unitList: element("unit-list"),
  viewer: element("viewer"),
  emptyState: element("empty-state"),
  surfaceWrap: element("surface-wrap"),
  surfaceStage: element("surface-stage"),
  surface: element("surface"),
  frozenTop: element("frozen-top"),
  frozenLeft: element("frozen-left"),
  frozenCorner: element("frozen-corner"),
  sheetColumnHeaders: element("sheet-column-headers"),
  sheetRowHeaders: element("sheet-row-headers"),
  sheetPageBreaks: element("sheet-page-breaks"),
  sheetHeaderCorner: element("sheet-header-corner"),
  sheetResizeGuide: element("sheet-resize-guide"),
  textLayer: element("text-layer"),
  mediaLayer: element("media-layer"),
  continuousView: element("continuous-view"),
  slideshowControls: element("slideshow-controls"),
  slideshowPrevious: element("slideshow-previous"),
  slideshowPosition: element("slideshow-position"),
  slideshowNext: element("slideshow-next"),
  slideshowExit: element("slideshow-exit"),
  selectionBox: element("selection-box"),
  speakerNotesPanel: element("speaker-notes-panel"),
  speakerNotes: element("speaker-notes"),
  objectId: element("object-id"),
  objectType: element("object-type"),
  objectBounds: element("object-bounds"),
  objectPixels: element("object-pixels"),
  objectText: element("object-text"),
  fontSelectedText: element("font-selected-text"),
  fontAuthoredFamily: element("font-authored-family"),
  fontRenderedFamily: element("font-rendered-family"),
  fontSource: element("font-source"),
  fontListStatus: element("font-list-status"),
  fontList: element("font-list"),
  sourceType: element("source-type"),
  sourcePart: element("source-part"),
  sourceQuality: element("source-quality"),
  sourceFields: element("source-fields"),
  ancestorList: element("ancestor-list"),
  documentTab: element("document-tab"),
  renderTab: element("render-tab"),
  diagnosticBody: element("diagnostic-body"),
  inspector: element("inspector"),
  inspectorToggle: element("inspector-toggle"),
  inspectorClose: element("inspector-close"),
  unitStatus: element("unit-status"),
  diagnosticStatus: element("diagnostic-status"),
  zoomOut: element("zoom-out"),
  zoomRange: element("zoom-range"),
  zoomIn: element("zoom-in"),
  zoomValue: element("zoom-value"),
};

const state = {
  engine: undefined,
  sourceDocument: undefined,
  officeDocument: undefined,
  openController: undefined,
  documentDiagnostics: [],
  renderDiagnostics: [],
  diagnosticObjects: undefined,
  diagnosticTab: "document",
  sessionRevision: 0,
  renderRevision: 0,
  surfaceContextLost: false,
  hitRevision: 0,
  displayedFrame: undefined,
  searchResults: [],
  searchIndex: -1,
  searchRevision: 0,
  searchTimer: undefined,
  viewportAnchor: undefined,
  sheetViewportOrigin: undefined,
  sheetScrollTimer: undefined,
  sheetSizes: new Map(),
  sheetResize: undefined,
  continuous: false,
  continuousRevision: 0,
  continuousObserver: undefined,
  continuousNavigationObserver: undefined,
  continuousPages: new Map(),
  mediaUrls: new Set(),
  continuousPixels: 0,
  retainContinuousPages: false,
  thumbnailRevision: 0,
  thumbnailObserver: undefined,
  thumbnailRecords: new Map(),
  thumbnailPixels: 0,
  documentFonts: [],
  fontSelectionFrame: undefined,
  fontGestureItem: undefined,
  zoomMode: "fit",
  slideshow: false,
  slideshowZoomMode: undefined,
  disposed: false,
};

function text(target, value) {
  target.textContent = value;
}

function truncate(value, maximum = MAX_VISIBLE_TEXT) {
  if (value.length <= maximum) return value;
  return `${value.slice(0, maximum)}\n…（已截断 ${value.length - maximum} 个字符）`;
}

function formatBytes(bytes) {
  if (!Number.isFinite(bytes) || bytes < 0) return "-";
  if (bytes < 1024) return `${bytes} B`;
  const units = ["KiB", "MiB", "GiB"];
  let value = bytes / 1024;
  let unit = units[0];
  for (let index = 1; index < units.length && value >= 1024; index += 1) {
    value /= 1024;
    unit = units[index];
  }
  let fractionDigits = 2;
  if (value >= 100) fractionDigits = 0;
  else if (value >= 10) fractionDigits = 1;
  return `${value.toFixed(fractionDigits)} ${unit}`;
}

function formatNumber(value) {
  if (!Number.isFinite(value)) return "-";
  return Number.isInteger(value) ? String(value) : value.toFixed(2).replace(/\.00$/, "");
}

function formatRect(rectangle) {
  return `x ${formatNumber(rectangle.x)}, y ${formatNumber(rectangle.y)}, w ${formatNumber(rectangle.width)}, h ${formatNumber(rectangle.height)}`;
}

function setPageState(pageState) {
  document.body.dataset.state = pageState;
  ui.app.dataset.state = pageState;
}

function setDetectedFormat(format) {
  document.body.dataset.detectedFormat = format;
  ui.app.dataset.detectedFormat = format;
}

function setDocumentKind(kind) {
  document.body.dataset.documentKind = kind;
  ui.app.dataset.documentKind = kind;
}

function setDocumentOpen(open) {
  ui.app.dataset.documentOpen = String(open);
}

function setStatus(message, status = "ready", explicitPageState) {
  ui.status.dataset.state = status;
  text(ui.status, message);
  let pageState = explicitPageState;
  if (pageState === undefined) {
    if (status === "loading") pageState = "loading";
    else if (status === "error") pageState = "error";
    else if (state.officeDocument === undefined) pageState = "idle";
    else pageState = "ready";
  }
  setPageState(pageState);
}

function errorMessage(cause) {
  if (cause instanceof OfficeEngineError) return `${cause.code}: ${cause.message}`;
  if (cause instanceof Error) return cause.message;
  return "发生未知错误";
}

function errorDiagnostics(cause) {
  return cause instanceof OfficeEngineError ? cause.diagnostics : [];
}

function unsupportedDocumentLabel(file) {
  const name = file.name.trim();
  if (/\.html?$/iu.test(name) || file.type.trim().toLowerCase() === "text/html") return "HTML/HTM";
  if (/\.txt$/iu.test(name)) return "TXT";
  return undefined;
}

function unitLabel(unit) {
  if (unit.type === "slide") return `幻灯片 ${unit.index + 1} · ${unit.name}`;
  if (unit.type === "sheet") {
    return `工作表 ${unit.index + 1} · ${unit.name} · ${unit.rows} 行 × ${unit.columns} 列`;
  }
  return `页面 ${unit.index + 1} · ${unit.name}`;
}

function navigationPresentation(kind) {
  const presentations = {
    presentation: { mode: "thumbnails", heading: "幻灯片", noun: "幻灯片" },
    spreadsheet: { mode: "sheets", heading: "工作表", noun: "工作表" },
    text: { mode: "pages", heading: "页面", noun: "页面" },
  };
  return presentations[kind] ?? { mode: "units", heading: "导航", noun: "项目" };
}

function selectedUnit() {
  const officeDocument = state.officeDocument;
  if (officeDocument === undefined) return undefined;
  const unitIndex = Number(ui.unitSelect.value);
  if (!Number.isInteger(unitIndex)) return undefined;
  const unit = officeDocument.info.units[unitIndex];
  return unit?.index === unitIndex ? unit : undefined;
}

function sheetOverrides(unit) {
  let overrides = state.sheetSizes.get(unit.index);
  if (overrides === undefined) {
    overrides = { rows: new Map(), columns: new Map() };
    state.sheetSizes.set(unit.index, overrides);
  }
  return overrides;
}

function sheetSizeRequest(unit) {
  if (unit.type !== "sheet" || unit.rows === 0 || unit.columns === 0) return undefined;
  const overrides = sheetOverrides(unit);
  return {
    rows: [...overrides.rows].map(([index, size]) => ({ index, size })),
    columns: [...overrides.columns].map(([index, size]) => ({ index, size })),
  };
}

function axisLayout(axis, count, overrides) {
  const baseSize = (index) => axis.spans.find((span) => span.start <= index && index <= span.end)?.size
    ?? axis.defaultSize;
  const baseOffset = (index) => {
    let offset = index * axis.defaultSize;
    for (const span of axis.spans) {
      if (span.start >= index) break;
      const covered = Math.min(index - 1, span.end) - span.start + 1;
      if (covered > 0) offset += covered * (span.size - axis.defaultSize);
    }
    return offset;
  };
  const offset = (index) => {
    let value = baseOffset(index);
    for (const [overrideIndex, size] of overrides) {
      if (overrideIndex < index) value += size - baseSize(overrideIndex);
    }
    return value;
  };
  const size = (index) => overrides.get(index) ?? baseSize(index);
  const indexAt = (value) => {
    if (count === 0) return -1;
    let low = 0;
    let high = count;
    while (low < high) {
      const middle = Math.floor((low + high + 1) / 2);
      if (offset(middle) <= value) low = middle;
      else high = middle - 1;
    }
    return Math.min(count - 1, Math.max(0, low));
  };
  const unmap = (value) => {
    if (value <= 0) return value;
    const total = offset(count);
    const baseTotal = baseOffset(count);
    if (value >= total) return baseTotal + value - total;
    const index = indexAt(value);
    const targetSize = size(index);
    return baseOffset(index) + (targetSize <= 0 ? 0 : (value - offset(index)) * baseSize(index) / targetSize);
  };
  const map = (value) => {
    if (value <= 0) return value;
    const baseTotal = baseOffset(count);
    if (value >= baseTotal) return offset(count) + value - baseTotal;
    let low = 0;
    let high = count;
    while (low < high) {
      const middle = Math.floor((low + high + 1) / 2);
      if (baseOffset(middle) <= value) low = middle;
      else high = middle - 1;
    }
    const originalSize = baseSize(low);
    return offset(low) + (originalSize <= 0 ? 0 : (value - baseOffset(low)) * size(low) / originalSize);
  };
  return { count, offset, size, indexAt, map, unmap, total: offset(count) };
}

function sheetLayout(unit) {
  const overrides = sheetOverrides(unit);
  const rows = axisLayout(unit.rowAxis, unit.rows, overrides.rows);
  const columns = axisLayout(unit.columnAxis, unit.columns, overrides.columns);
  return {
    rows: unit.rows === 0 ? { ...rows, total: unit.height } : rows,
    columns: unit.columns === 0 ? { ...columns, total: unit.width } : columns,
  };
}

function columnLabel(index) {
  let label = "";
  for (let value = index + 1; value > 0; value = Math.floor((value - 1) / 26)) {
    label = String.fromCharCode(65 + (value - 1) % 26) + label;
  }
  return label;
}

function updateSpeakerNotes(unit) {
  const notes = unit?.type === "slide" ? unit.speakerNotes : undefined;
  const paragraphs = unit?.type === "slide" ? unit.speakerNoteParagraphs : undefined;
  ui.speakerNotesPanel.hidden = unit?.type !== "slide";
  ui.speakerNotes.replaceChildren();
  if (unit?.type !== "slide") return;
  if (paragraphs === undefined || paragraphs.length === 0) {
    text(ui.speakerNotes, notes === undefined || notes.trim() === "" ? "此幻灯片没有演讲者讲稿。" : truncate(notes));
    return;
  }
  const fragment = document.createDocumentFragment();
  let remaining = MAX_VISIBLE_TEXT;
  for (const paragraph of paragraphs) {
    if (remaining <= 0) break;
    const block = document.createElement("div");
    block.className = "speaker-notes-paragraph";
    if (paragraph.align === "center") block.style.textAlign = "center";
    else if (paragraph.align === "end") block.style.textAlign = "end";
    else if (paragraph.align === "start") block.style.textAlign = "start";
    else block.style.textAlign = "justify";
    block.style.marginLeft = `${paragraph.marginLeft}px`;
    block.style.marginRight = `${paragraph.marginRight}px`;
    block.style.textIndent = `${paragraph.firstLineIndent}px`;
    block.style.marginTop = `${paragraph.spaceBefore}px`;
    block.style.marginBottom = `${paragraph.spaceAfter}px`;
    if (paragraph.lineHeight > 0) block.style.lineHeight = `${paragraph.lineHeight}px`;
    for (const run of paragraph.runs) {
      if (remaining <= 0) break;
      const value = run.text.slice(0, remaining);
      remaining -= value.length;
      const span = document.createElement("span");
      span.style.fontFamily = cssFontFamily(run.fontFamily);
      span.style.fontSize = `${run.fontSize}px`;
      span.style.color = packedRgba(run.color);
      span.style.fontWeight = run.bold ? "700" : "400";
      span.style.fontStyle = run.italic ? "italic" : "normal";
      span.style.textDecorationLine = [run.underline && "underline", run.strikethrough && "line-through"]
        .filter(Boolean)
        .join(" ") || "none";
      span.style.letterSpacing = `${run.letterSpacing}px`;
      span.style.verticalAlign = `${run.baselineShift}px`;
      if ((run.highlight & 0xff) !== 0) span.style.backgroundColor = packedRgba(run.highlight);
      text(span, value);
      block.append(span);
    }
    fragment.append(block);
  }
  if (remaining <= 0) fragment.append(document.createTextNode("\n…（已截断）"));
  ui.speakerNotes.replaceChildren(fragment);
}

function packedRgba(value) {
  const red = (value >>> 24) & 0xff;
  const green = (value >>> 16) & 0xff;
  const blue = (value >>> 8) & 0xff;
  const alpha = (value & 0xff) / 255;
  return `rgba(${red}, ${green}, ${blue}, ${alpha})`;
}

function resetHit() {
  state.hitRevision += 1;
  ui.selectionBox.dataset.visible = "false";
  text(ui.objectId, "-");
  text(ui.objectType, "-");
  text(ui.objectBounds, "-");
  text(ui.objectPixels, "-");
  text(ui.objectText, "-");
  text(ui.sourceType, "-");
  text(ui.sourcePart, "-");
  text(ui.sourceQuality, "-");
  text(ui.sourceFields, "-");
  ui.ancestorList.replaceChildren();
  const empty = document.createElement("li");
  empty.className = "ancestor-empty";
  text(empty, "-");
  ui.ancestorList.append(empty);
}

const FONT_SOURCE_LABELS = {
  embedded: "文档内嵌",
  host: "宿主提供",
  browser: "浏览器本地",
  fallback: "近似回退",
};
const GENERIC_FONT_FAMILIES = new Set([
  "serif", "sans-serif", "monospace", "cursive", "fantasy", "system-ui",
  "ui-serif", "ui-sans-serif", "ui-monospace", "ui-rounded",
]);

function resetFontSelection() {
  if (state.fontSelectionFrame !== undefined) cancelAnimationFrame(state.fontSelectionFrame);
  state.fontSelectionFrame = undefined;
  state.fontGestureItem = undefined;
  text(ui.fontSelectedText, "请在文档中选择文字");
  text(ui.fontAuthoredFamily, "-");
  text(ui.fontRenderedFamily, "-");
  text(ui.fontSource, "-");
}

function selectedFontRuns(selection) {
  if (selection.rangeCount === 0 || selection.isCollapsed) return [];
  const range = selection.getRangeAt(0);
  const fontRun = (node) => (node instanceof Element ? node : node.parentElement)?.closest("[data-authored-family]");
  const start = fontRun(range.startContainer);
  const end = fontRun(range.endContainer);
  const fragment = range.cloneContents();
  const candidates = [start, ...fragment.querySelectorAll("[data-authored-family]"), end];
  const runs = new Map();
  for (const candidate of candidates) {
    if (!(candidate instanceof HTMLElement)) continue;
    const authoredFamily = candidate.dataset.authoredFamily;
    const renderedFamily = candidate.dataset.renderedFamily;
    const source = candidate.dataset.fontSource;
    if (authoredFamily === undefined || renderedFamily === undefined || source === undefined) continue;
    runs.set(`${authoredFamily}\0${renderedFamily}\0${source}`, { authoredFamily, renderedFamily, source });
  }
  return [...runs.values()];
}

function itemFontRuns(item) {
  const runs = new Map();
  for (const candidate of item.querySelectorAll("[data-authored-family]")) {
    if (!(candidate instanceof HTMLElement)) continue;
    const authoredFamily = candidate.dataset.authoredFamily;
    const renderedFamily = candidate.dataset.renderedFamily;
    const source = candidate.dataset.fontSource;
    if (authoredFamily === undefined || renderedFamily === undefined || source === undefined) continue;
    runs.set(`${authoredFamily}\0${renderedFamily}\0${source}`, { authoredFamily, renderedFamily, source });
  }
  return [...runs.values()];
}

function displayedFontFamily(run) {
  return run.source === "fallback" ? run.renderedFamily : run.authoredFamily;
}

function renderFontFamilies(target, runs, label) {
  const unique = new Map();
  for (const run of runs) {
    const name = label(run);
    const key = `${name}\0${run.renderedFamily}`;
    if (!unique.has(key)) unique.set(key, { name, family: run.renderedFamily });
  }
  target.replaceChildren();
  let index = 0;
  for (const { name, family } of unique.values()) {
    if (index !== 0) target.append(document.createTextNode(" / "));
    const preview = document.createElement("span");
    preview.className = "font-family-preview";
    preview.style.fontFamily = cssFontFamily(family);
    text(preview, name);
    target.append(preview);
    index += 1;
  }
  if (index === 0) text(target, "-");
}

function showFontDetails(selectedText, runs) {
  text(ui.fontSelectedText, truncate(selectedText, 300));
  renderFontFamilies(ui.fontAuthoredFamily, runs, (run) => run.authoredFamily);
  renderFontFamilies(ui.fontRenderedFamily, runs, displayedFontFamily);
  text(ui.fontSource, [...new Set(runs.map((run) => FONT_SOURCE_LABELS[run.source] ?? run.source))].join(" / "));
}

function updateSelectedFont() {
  state.fontSelectionFrame = undefined;
  const selection = document.getSelection();
  if (selection === null || selection.isCollapsed || selection.rangeCount === 0) {
    const item = state.fontGestureItem;
    const runs = item?.isConnected === true ? itemFontRuns(item) : [];
    if (item !== undefined && runs.length > 0) {
      showFontDetails(item.textContent ?? "", runs);
      return;
    }
    resetFontSelection();
    return;
  }
  const range = selection.getRangeAt(0);
  const container = range.commonAncestorContainer instanceof Element
    ? range.commonAncestorContainer
    : range.commonAncestorContainer.parentElement;
  if (container === null || container.closest(".text-layer") === null) {
    resetFontSelection();
    return;
  }
  const runs = selectedFontRuns(selection);
  if (runs.length === 0) {
    resetFontSelection();
    return;
  }
  showFontDetails(selection.toString(), runs);
}

function queueSelectedFontUpdate() {
  if (state.fontSelectionFrame !== undefined) cancelAnimationFrame(state.fontSelectionFrame);
  state.fontSelectionFrame = requestAnimationFrame(updateSelectedFont);
}

function normalizedFontFamily(value) {
  return value.trim().replace(/^(['"])(.*)\1$/u, "$2");
}

function cssFontFamily(value) {
  const family = normalizedFontFamily(value);
  return GENERIC_FONT_FAMILIES.has(family.toLowerCase()) ? family : JSON.stringify(family);
}

function renderDocumentFonts() {
  const entries = new Map();
  const add = (family, source) => {
    const normalized = normalizedFontFamily(family);
    if (normalized.length === 0) return;
    const sources = entries.get(normalized) ?? new Set();
    sources.add(source);
    entries.set(normalized, sources);
  };
  for (const run of state.documentFonts) {
    add(run.source === "fallback" ? run.renderedFamily : run.authoredFamily, FONT_SOURCE_LABELS[run.source] ?? run.source);
  }
  ui.fontList.replaceChildren();
  const sorted = [...entries].sort(([left], [right]) => left.localeCompare(right, undefined, { sensitivity: "base" }));
  if (sorted.length === 0) {
    const empty = document.createElement("li");
    empty.className = "font-list-empty";
    text(empty, state.officeDocument === undefined ? "当前文档没有字体信息" : "当前文档没有文字字体信息");
    ui.fontList.append(empty);
  } else {
    const fragment = document.createDocumentFragment();
    for (const [family, sources] of sorted) {
      const item = document.createElement("li");
      const name = document.createElement("span");
      name.className = "font-list-name code-value";
      text(name, family);
      const source = document.createElement("span");
      source.className = "font-list-source";
      text(source, [...sources].join(" / "));
      item.append(name, source);
      fragment.append(item);
    }
    ui.fontList.append(fragment);
  }
  text(ui.fontListStatus, state.officeDocument === undefined ? "尚未打开文档" : `${sorted.length} 个文档字体`);
}

async function updateDocumentFonts(officeDocument, expectedSession) {
  try {
    const objects = await officeDocument.listObjects({
      unitIndex: selectedUnit()?.index ?? 0,
      textOnly: true,
      limit: 100_000,
    });
    if (officeDocument !== state.officeDocument || expectedSession !== state.sessionRevision) return;
    state.documentFonts = objects.flatMap((object) => object.fontRuns ?? []);
    renderDocumentFonts();
  } catch (cause) {
    if (officeDocument !== state.officeDocument || expectedSession !== state.sessionRevision) return;
    state.documentFonts = [];
    text(ui.fontListStatus, `文档字体读取失败：${errorMessage(cause)}`);
  }
}

function clearCanvas() {
  state.renderRevision += 1;
  state.displayedFrame = undefined;
  const context = ui.surface.getContext("2d");
  context?.clearRect(0, 0, ui.surface.width, ui.surface.height);
  ui.surface.width = 0;
  ui.surface.height = 0;
  ui.surface.style.width = "0px";
  ui.surface.style.height = "0px";
  ui.surfaceStage.style.width = "0px";
  ui.surfaceStage.style.height = "0px";
  ui.surfaceStage.style.left = "";
  ui.surfaceStage.style.top = "";
  ui.surfaceWrap.style.width = "";
  ui.surfaceWrap.style.height = "";
  ui.surfaceWrap.dataset.sheet = "false";
  ui.sheetColumnHeaders.replaceChildren();
  ui.sheetRowHeaders.replaceChildren();
  ui.sheetPageBreaks.replaceChildren();
  ui.sheetResizeGuide.dataset.visible = "false";
  clearFrozenSurfaces();
  ui.surfaceWrap.dataset.continuous = "false";
  ui.textLayer.replaceChildren();
  clearMediaLayer();
  ui.continuousView.replaceChildren();
  ui.continuousView.hidden = true;
  ui.surfaceStage.hidden = false;
  state.continuousObserver?.disconnect();
  state.continuousObserver = undefined;
  state.continuousNavigationObserver?.disconnect();
  state.continuousNavigationObserver = undefined;
  state.continuousPages.clear();
  state.continuousPixels = 0;
  state.retainContinuousPages = false;
  state.continuousRevision += 1;
  ui.surfaceWrap.dataset.visible = "false";
  ui.emptyState.hidden = false;
  resetFontSelection();
  resetHit();
}

function clearMediaLayer() {
  for (const player of ui.mediaLayer.querySelectorAll("audio, video")) player.pause();
  ui.mediaLayer.replaceChildren();
  for (const url of state.mediaUrls) URL.revokeObjectURL(url);
  state.mediaUrls.clear();
}

function handleSurfaceContextLost() {
  if (state.surfaceContextLost) return;
  state.surfaceContextLost = true;
  state.renderRevision += 1;
  state.displayedFrame = undefined;
  ui.textLayer.replaceChildren();
  clearMediaLayer();
  clearFrozenSurfaces();
  resetFontSelection();
  resetHit();
  if (state.officeDocument !== undefined) {
    setStatus("浏览器正在恢复文档画布；恢复后将自动重新渲染。", "warning");
  }
}

function handleSurfaceContextRestored() {
  if (!state.surfaceContextLost) return;
  state.surfaceContextLost = false;
  if (state.officeDocument === undefined || state.disposed) return;
  setStatus("文档画布已恢复，正在重新渲染…", "loading");
  void renderCurrentUnit(state.sessionRevision);
}

function renderMediaLayer(mediaItems, displayed) {
  clearMediaLayer();
  const scaleX = displayed.cssWidth / displayed.viewport.width;
  const scaleY = displayed.cssHeight / displayed.viewport.height;
  const fragment = document.createDocumentFragment();
  for (const media of mediaItems) {
    const { bounds, transform } = media;
    const url = URL.createObjectURL(new Blob([media.bytes], { type: media.mediaType }));
    state.mediaUrls.add(url);
    const left = (bounds.x - displayed.viewport.x) * scaleX;
    const top = (bounds.y - displayed.viewport.y) * scaleY;
    const localE = (transform.a * bounds.x + transform.c * bounds.y + transform.e - bounds.x) * scaleX;
    const localF = (transform.b * bounds.x + transform.d * bounds.y + transform.f - bounds.y) * scaleY;
    const matrix = [
      transform.a,
      transform.b * scaleY / scaleX,
      transform.c * scaleX / scaleY,
      transform.d,
      localE,
      localF,
    ];
    if (media.kind === "video") {
      const player = document.createElement("video");
      player.className = "embedded-media";
      player.src = url;
      player.controls = true;
      player.playsInline = true;
      player.preload = "metadata";
      player.setAttribute("aria-label", `嵌入式视频 ${media.objectId}`);
      player.style.left = `${left}px`;
      player.style.top = `${top}px`;
      player.style.width = `${bounds.width * scaleX}px`;
      player.style.height = `${bounds.height * scaleY}px`;
      player.style.transform = `matrix(${matrix.join(",")})`;
      fragment.append(player);
      continue;
    }
    const player = document.createElement("audio");
    player.src = url;
    player.preload = "metadata";
    player.hidden = true;
    const button = document.createElement("button");
    button.type = "button";
    button.className = "embedded-media";
    button.setAttribute("aria-label", `播放嵌入式音频 ${media.objectId}`);
    button.style.left = `${left}px`;
    button.style.top = `${top}px`;
    button.style.width = `${bounds.width * scaleX}px`;
    button.style.height = `${bounds.height * scaleY}px`;
    button.style.transform = `matrix(${matrix.join(",")})`;
    text(button, "▶");
    button.addEventListener("click", () => {
      if (player.paused) void player.play().catch(() => undefined);
      else player.pause();
    });
    player.addEventListener("play", () => {
      text(button, "❚❚");
      button.setAttribute("aria-label", `暂停嵌入式音频 ${media.objectId}`);
    });
    const showPlay = () => {
      text(button, "▶");
      button.setAttribute("aria-label", `播放嵌入式音频 ${media.objectId}`);
    };
    player.addEventListener("pause", showPlay);
    player.addEventListener("ended", showPlay);
    fragment.append(player, button);
  }
  ui.mediaLayer.append(fragment);
}

function frozenSurfaces() {
  return [ui.frozenTop, ui.frozenLeft, ui.frozenCorner];
}

function clearFrozenSurfaces() {
  ui.surfaceWrap.dataset.frozen = "false";
  for (const canvas of frozenSurfaces()) {
    const context = canvas.getContext("2d");
    context?.clearRect(0, 0, canvas.width, canvas.height);
    canvas.width = 0;
    canvas.height = 0;
    canvas.style.width = "0px";
    canvas.style.height = "0px";
    canvas.style.left = "";
    canvas.style.top = "";
    canvas.dataset.visible = "false";
  }
}

function updateSearchControls() {
  const count = state.searchResults.length;
  const active = count > 0 && state.searchIndex >= 0;
  ui.searchPrevious.disabled = !active;
  ui.searchNext.disabled = !active;
  text(ui.searchStatus, active ? `${state.searchIndex + 1} / ${count}` : `0 / ${count}`);
}

function resetSearch(options = {}) {
  state.searchRevision += 1;
  if (state.searchTimer !== undefined) clearTimeout(state.searchTimer);
  state.searchTimer = undefined;
  state.searchResults = [];
  state.searchIndex = -1;
  state.viewportAnchor = undefined;
  if (options.clearQuery !== false) ui.searchInput.value = "";
  if (options.disable === true) ui.searchInput.disabled = true;
  updateSearchControls();
}

function closeActiveSession() {
  if (leaveSlideshowState() && document.fullscreenElement === ui.viewer) {
    void document.exitFullscreen().catch(() => undefined);
  }
  state.openController?.abort();
  state.openController = undefined;
  state.thumbnailRevision += 1;
  state.thumbnailObserver?.disconnect();
  state.thumbnailObserver = undefined;
  state.thumbnailRecords.clear();
  state.thumbnailPixels = 0;
  const officeDocument = state.officeDocument;
  const sourceDocument = state.sourceDocument;
  state.officeDocument = undefined;
  state.sourceDocument = undefined;
  state.diagnosticObjects = undefined;
  officeDocument?.close();
  if (sourceDocument !== officeDocument) sourceDocument?.close();
  if (state.sheetScrollTimer !== undefined) clearTimeout(state.sheetScrollTimer);
  state.sheetScrollTimer = undefined;
  state.sheetViewportOrigin = undefined;
  state.sheetSizes.clear();
  state.sheetResize = undefined;
  state.documentFonts = [];
  renderDocumentFonts();
  resetSearch({ disable: true });
  clearCanvas();
}

function resetUnitControls() {
  ui.unitSelect.replaceChildren();
  const option = document.createElement("option");
  option.value = "";
  text(option, "-");
  ui.unitSelect.append(option);
  ui.unitSelect.disabled = true;
  ui.previousUnit.disabled = true;
  ui.nextUnit.disabled = true;
  ui.fitView.disabled = true;
  ui.continuousButton.disabled = true;
  ui.continuousButton.setAttribute("aria-pressed", "false");
  state.continuous = false;
  ui.slideshowButton.hidden = true;
  ui.slideshowButton.disabled = true;
  ui.slideshowControls.hidden = true;
  ui.viewer.dataset.playing = "false";
  ui.printButton.disabled = true;
  ui.searchInput.disabled = true;
  ui.zoomOut.disabled = true;
  ui.zoomRange.disabled = true;
  ui.zoomIn.disabled = true;
  ui.unitList.replaceChildren();
  ui.unitList.dataset.navigationMode = "empty";
  text(ui.unitNavigationTitle, "导航");
  text(ui.unitCount, "0 项");
  text(ui.unitStatus, "未加载文档");
  state.zoomMode = "fit";
  state.sheetViewportOrigin = undefined;
  ui.zoomRange.value = "100";
  text(ui.zoomValue, "适应");
}

function unitButtons() {
  return [...ui.unitList.querySelectorAll(".unit-button")];
}

function updateUnitNavigation(options = {}) {
  const buttons = unitButtons();
  const selected = ui.unitSelect.selectedIndex;
  for (const [index, button] of buttons.entries()) {
    const isSelected = index === selected;
    button.setAttribute("aria-selected", String(isSelected));
    button.tabIndex = isSelected ? 0 : -1;
  }
  if (options.focus === true) buttons[selected]?.focus();
}

function selectContinuousUnit(unitIndex) {
  if (!state.continuous) return;
  const position = [...ui.unitSelect.options]
    .findIndex((option) => Number(option.value) === unitIndex);
  if (position < 0 || ui.unitSelect.selectedIndex === position) return;
  ui.unitSelect.selectedIndex = position;
  updateUnitButtons();
  const button = ui.unitList.querySelector(`[data-unit-index="${unitIndex}"]`);
  button?.scrollIntoView({ block: "nearest" });
}

function revealContinuousPage(record, block = "start") {
  const viewerBounds = ui.viewer.getBoundingClientRect();
  const pageBounds = record.element.getBoundingClientRect();
  let distance = 0;
  if (pageBounds.bottom < viewerBounds.top) {
    distance = viewerBounds.top - pageBounds.bottom;
  } else if (pageBounds.top > viewerBounds.bottom) {
    distance = pageBounds.top - viewerBounds.bottom;
  }
  const smoothLimit = Math.max(1, ui.viewer.clientHeight) * CONTINUOUS_SMOOTH_SCROLL_VIEWPORTS;
  record.element.scrollIntoView({
    block,
    behavior: distance <= smoothLimit ? "smooth" : "auto",
  });
}

function updateUnitButtons() {
  const hasUnits = state.officeDocument !== undefined && ui.unitSelect.options.length > 0;
  const selected = ui.unitSelect.selectedIndex;
  ui.previousUnit.disabled = !hasUnits || selected <= 0;
  ui.nextUnit.disabled = !hasUnits || selected < 0 || selected >= ui.unitSelect.options.length - 1;
  ui.fitView.disabled = !hasUnits;
  ui.zoomOut.disabled = !hasUnits;
  ui.zoomRange.disabled = !hasUnits;
  ui.zoomIn.disabled = !hasUnits;
  const slideshowAvailable = hasUnits && selectedUnit()?.type === "slide";
  ui.slideshowButton.hidden = !slideshowAvailable;
  ui.slideshowButton.disabled = !slideshowAvailable;
  ui.slideshowPrevious.disabled = !slideshowAvailable || selected <= 0;
  ui.slideshowNext.disabled = !slideshowAvailable || selected >= ui.unitSelect.options.length - 1;
  text(ui.slideshowPosition, slideshowAvailable
    ? `${selected + 1} / ${ui.unitSelect.options.length}`
    : "0 / 0");
  updateUnitNavigation();
  if (!hasUnits) {
    text(ui.unitStatus, "未加载文档");
    return;
  }
  const unit = selectedUnit();
  const presentation = navigationPresentation(state.officeDocument.info.kind);
  text(ui.unitStatus, unit === undefined
    ? `${ui.unitSelect.options.length} ${presentation.noun}`
    : `${unit.index + 1} / ${ui.unitSelect.options.length} · ${presentation.noun}`);
}

function populateUnits(units) {
  state.thumbnailObserver?.disconnect();
  state.thumbnailObserver = undefined;
  state.thumbnailRecords.clear();
  state.thumbnailPixels = 0;
  ui.unitSelect.replaceChildren();
  ui.unitList.replaceChildren();
  const kind = state.officeDocument?.info.kind ?? "";
  const presentation = navigationPresentation(kind);
  ui.unitList.dataset.navigationMode = presentation.mode;
  text(ui.unitNavigationTitle, presentation.heading);
  text(ui.unitCount, `${units.length} ${presentation.noun}`);
  for (const [position, unit] of units.entries()) {
    const option = document.createElement("option");
    option.value = String(unit.index);
    text(option, unitLabel(unit));
    ui.unitSelect.append(option);

    const button = document.createElement("button");
    button.className = "unit-button";
    button.type = "button";
    button.dataset.unitIndex = String(unit.index);
    button.setAttribute("role", "option");
    button.setAttribute("aria-label", unitLabel(unit));
    button.setAttribute("aria-selected", "false");
    button.tabIndex = -1;

    const order = document.createElement("span");
    order.className = "unit-order";
    text(order, String(unit.index + 1));
    const card = document.createElement("span");
    card.className = "unit-card";
    if (presentation.mode !== "sheets" && presentation.mode !== "flow") {
      const preview = document.createElement("canvas");
      preview.className = "unit-preview";
      preview.width = 1;
      preview.height = 1;
      preview.style.aspectRatio = `${unit.width} / ${unit.height}`;
      preview.dataset.thumbnailState = "pending";
      preview.setAttribute("aria-hidden", "true");
      card.append(preview);
      state.thumbnailRecords.set(unit.index, {
        unit,
        canvas: preview,
        rendering: false,
        visible: false,
        pixels: 0,
        lastUsed: 0,
      });
    }
    const name = document.createElement("span");
    name.className = "unit-name";
    text(name, unit.name || `${presentation.noun} ${unit.index + 1}`);
    const meta = document.createElement("span");
    meta.className = "unit-meta";
    text(meta, unit.type === "sheet"
      ? `${unit.rows} 行 × ${unit.columns} 列`
      : `${formatNumber(unit.width)} × ${formatNumber(unit.height)}`);
    card.append(name, meta);
    button.append(order, card);
    button.addEventListener("click", () => selectUnitPosition(position));
    ui.unitList.append(button);
  }
  const enabled = units.length > 0;
  ui.unitSelect.disabled = !enabled;
  ui.searchInput.disabled = !enabled;
  ui.printButton.disabled = !enabled;
  ui.continuousButton.disabled = !enabled || kind !== "text";
  if (enabled) ui.unitSelect.selectedIndex = 0;
  updateUnitButtons();
}

function updateCurrentUnitPreview(unitIndex) {
  const preview = ui.unitList.querySelector(`[data-unit-index="${unitIndex}"] .unit-preview`);
  if (!(preview instanceof HTMLCanvasElement) || ui.surface.width === 0 || ui.surface.height === 0) return;
  const record = state.thumbnailRecords.get(unitIndex);
  if (record !== undefined && record.pixels === 0) {
    const ratio = Math.min(Math.max(window.devicePixelRatio || 1, 1), 2);
    const scale = thumbnailScale(record.unit) * ratio;
    preview.width = Math.max(1, Math.round(record.unit.width * scale));
    preview.height = Math.max(1, Math.round(record.unit.height * scale));
  }
  const context = preview.getContext("2d");
  if (context === null) return;
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, preview.width, preview.height);
  const scale = Math.min(preview.width / ui.surface.width, preview.height / ui.surface.height);
  const width = ui.surface.width * scale;
  const height = ui.surface.height * scale;
  context.drawImage(ui.surface, (preview.width - width) / 2, (preview.height - height) / 2, width, height);
  preview.dataset.thumbnailState = "ready";
  if (record !== undefined) {
    state.thumbnailPixels -= record.pixels;
    record.pixels = preview.width * preview.height;
    record.lastUsed = performance.now();
    state.thumbnailPixels += record.pixels;
    evictThumbnailPixels();
  }
}

function thumbnailScale(unit) {
  return Math.min(
    1,
    THUMBNAIL_MAX_WIDTH / unit.width,
    THUMBNAIL_MAX_HEIGHT / unit.height,
  );
}

async function renderUnitThumbnail(officeDocument, unit, expectedSession, expectedRevision) {
  const record = state.thumbnailRecords.get(unit.index);
  const preview = record?.canvas;
  if (!(preview instanceof HTMLCanvasElement)
    || record.rendering
    || preview.dataset.thumbnailState === "ready") return;
  record.rendering = true;
  preview.dataset.thumbnailState = "rendering";
  let frame;
  let retry = false;
  try {
    frame = await officeDocument.render({
      unitIndex: unit.index,
      scale: thumbnailScale(unit),
      pixelRatio: Math.min(Math.max(window.devicePixelRatio || 1, 1), 2),
      background: "#ffffff",
    }, {
      priority: "prefetch",
      supersedeKey: `thumbnail:${unit.index}`,
    });
    if (expectedSession !== state.sessionRevision
      || expectedRevision !== state.thumbnailRevision
      || officeDocument !== state.officeDocument) return;
    preview.width = frame.pixelWidth;
    preview.height = frame.pixelHeight;
    const context = preview.getContext("2d");
    if (context === null) throw new Error("浏览器无法创建缩略图 Canvas 2D 上下文");
    context.clearRect(0, 0, frame.pixelWidth, frame.pixelHeight);
    context.drawImage(frame.bitmap, 0, 0);
    preview.dataset.thumbnailState = "ready";
    state.thumbnailPixels -= record.pixels;
    record.pixels = frame.pixelWidth * frame.pixelHeight;
    record.lastUsed = performance.now();
    state.thumbnailPixels += record.pixels;
    evictThumbnailPixels();
  } catch (cause) {
    if (expectedSession !== state.sessionRevision
      || expectedRevision !== state.thumbnailRevision
      || officeDocument !== state.officeDocument) return;
    if (cause instanceof OfficeEngineError && cause.code === "OPERATION_ABORTED") {
      preview.dataset.thumbnailState = "pending";
      retry = record.visible;
    } else {
      preview.dataset.thumbnailState = "error";
      preview.title = `缩略图渲染失败：${errorMessage(cause)}`;
    }
  } finally {
    record.rendering = false;
    frame?.bitmap.close();
  }
  if (retry
    && expectedSession === state.sessionRevision
    && expectedRevision === state.thumbnailRevision
    && officeDocument === state.officeDocument) {
    void renderUnitThumbnail(officeDocument, unit, expectedSession, expectedRevision);
  }
}

function releaseThumbnail(record) {
  if (record.rendering || record.visible || record.pixels === 0) return;
  state.thumbnailPixels -= record.pixels;
  record.pixels = 0;
  record.canvas.width = 1;
  record.canvas.height = 1;
  record.canvas.dataset.thumbnailState = "pending";
}

function evictThumbnailPixels() {
  if (state.thumbnailPixels <= THUMBNAIL_CACHE_PIXELS) return;
  const selected = selectedUnit()?.index;
  const candidates = [...state.thumbnailRecords.values()]
    .filter((record) => !record.visible && record.unit.index !== selected && record.pixels > 0)
    .sort((left, right) => left.lastUsed - right.lastUsed);
  for (const record of candidates) {
    releaseThumbnail(record);
    if (state.thumbnailPixels <= THUMBNAIL_CACHE_PIXELS) break;
  }
}

function renderUnitThumbnails(officeDocument, expectedSession) {
  const expectedRevision = ++state.thumbnailRevision;
  state.thumbnailObserver?.disconnect();
  if (typeof IntersectionObserver === "function") {
    state.thumbnailObserver = new IntersectionObserver((entries) => {
      if (expectedSession !== state.sessionRevision
        || expectedRevision !== state.thumbnailRevision
        || officeDocument !== state.officeDocument) return;
      for (const entry of entries) {
        const unitIndex = Number(entry.target.closest("[data-unit-index]")?.dataset.unitIndex);
        const record = state.thumbnailRecords.get(unitIndex);
        if (record === undefined) continue;
        record.visible = entry.isIntersecting;
        if (entry.isIntersecting) {
          record.lastUsed = performance.now();
          void renderUnitThumbnail(officeDocument, record.unit, expectedSession, expectedRevision);
        }
      }
      evictThumbnailPixels();
    }, { root: ui.unitList, rootMargin: "320px 0px", threshold: 0.01 });
    for (const unit of officeDocument.info.units.filter((unit) => unit.type !== "sheet")) {
      const record = state.thumbnailRecords.get(unit.index);
      if (record !== undefined) state.thumbnailObserver.observe(record.canvas);
    }
    return;
  }
  const selected = selectedUnit()?.index ?? 0;
  for (const index of [selected, selected + 1]) {
    const record = state.thumbnailRecords.get(index);
    if (record !== undefined) void renderUnitThumbnail(officeDocument, record.unit, expectedSession, expectedRevision);
  }
}

function diagnosticLocation(diagnostic) {
  const values = [];
  if (diagnostic.part !== undefined) values.push(diagnostic.part);
  if (diagnostic.objectId !== undefined) values.push(diagnostic.objectId);
  const unitIndex = Number(diagnostic.details?.unitIndex);
  const unit = Number.isInteger(unitIndex) && unitIndex >= 0
    ? state.officeDocument?.info.units[unitIndex]
    : undefined;
  if (unit !== undefined) values.push(unitLabel(unit));
  return values.length === 0 ? "-" : values.join(" · ");
}

function diagnosticDetails(diagnostic) {
  const summary = `${diagnostic.fidelity} / ${diagnostic.phase}`;
  if (diagnostic.details === undefined || Object.keys(diagnostic.details).length === 0) return summary;
  return `${summary}\n${truncate(JSON.stringify(diagnostic.details, null, 2), 1_200)}`;
}

function diagnosticTarget(diagnostic) {
  const unitIndex = Number(diagnostic.details?.unitIndex);
  if (!Number.isInteger(unitIndex) || unitIndex < 0) return undefined;
  const values = ["x", "y", "width", "height"].map((key) => Number(diagnostic.details?.[key]));
  const bounds = values.every(Number.isFinite) && values[2] >= 0 && values[3] >= 0
    ? { x: values[0], y: values[1], width: values[2], height: values[3] }
    : undefined;
  return { unitIndex, bounds };
}

function diagnosticObjectId(diagnostic) {
  if (diagnostic.objectId !== undefined) return diagnostic.objectId;
  const objectId = diagnostic.details?.objectId;
  return typeof objectId === "string" && objectId.length > 0 ? objectId : undefined;
}

function normalizedPart(value) {
  return value.replace(/^\/+|\\/gu, "/");
}

function diagnosticSourceHints(diagnostic) {
  const details = diagnostic.details ?? {};
  const aliases = {
    address: "address",
    column: "column",
    drawingId: "drawingId",
    elementId: "elementId",
    paragraphId: "paragraphId",
    paragraphIndex: "paragraphIndex",
    path: "path",
    recordOffset: "recordOffset",
    row: "row",
    shapeId: "shapeId",
    sheetName: "sheetName",
    tableName: "tableName",
  };
  return Object.entries(aliases)
    .filter(([detailKey]) => details[detailKey] !== undefined)
    .map(([detailKey, sourceKey]) => [sourceKey, details[detailKey]]);
}

function sourceMatchesHints(source, hints) {
  return hints.every(([key, value]) => source[key] !== undefined
    && String(source[key]) === String(value));
}

async function resolveDiagnosticTarget(officeDocument, diagnostic) {
  const objectId = diagnosticObjectId(diagnostic);
  if (objectId !== undefined) {
    const object = await officeDocument.getObject(objectId);
    if (object !== undefined) return { object, unitIndex: object.unitIndex, bounds: object.bounds };
  }
  const explicit = diagnosticTarget(diagnostic);
  if (explicit !== undefined) return explicit;
  if (diagnostic.part === undefined) return undefined;
  const part = normalizedPart(diagnostic.part);
  state.diagnosticObjects ??= officeDocument.listObjects({ limit: MAX_DIAGNOSTIC_OBJECTS });
  const objects = await state.diagnosticObjects;
  let matches = objects.filter((object) => normalizedPart(object.source.part) === part);
  if (matches.length === 0) return undefined;
  const hints = diagnosticSourceHints(diagnostic);
  if (hints.length > 0) {
    const hinted = matches.filter((object) => sourceMatchesHints(object.source, hints));
    if (hinted.length > 0) matches = hinted;
  }
  const units = new Set(matches.map((object) => object.unitIndex));
  if (units.size !== 1) return undefined;
  const unitIndex = matches[0].unitIndex;
  return matches.length === 1
    ? { object: matches[0], unitIndex, bounds: matches[0].bounds }
    : { unitIndex };
}

async function focusDiagnostic(diagnostic) {
  const officeDocument = state.officeDocument;
  if (officeDocument === undefined) return;
  const target = await resolveDiagnosticTarget(officeDocument, diagnostic);
  if (target === undefined || officeDocument !== state.officeDocument) {
    setStatus(`该诊断只有文件级位置，无法可靠映射到单个页面：${diagnostic.message}`, "warning");
    return;
  }
  const object = target.object;
  const position = [...ui.unitSelect.options]
    .findIndex((option) => Number(option.value) === target.unitIndex);
  if (position < 0) return;
  state.viewportAnchor = target.bounds === undefined ? undefined : {
    unitIndex: target.unitIndex,
    bounds: target.bounds,
  };
  ui.unitSelect.selectedIndex = position;
  updateUnitButtons();
  if (state.continuous
    && target.bounds !== undefined
    && object === undefined) {
    setContinuousMode(false);
  }
  if (state.continuous && officeDocument.info.kind === "text") {
    await renderContinuousPage(target.unitIndex);
    if (officeDocument !== state.officeDocument) return;
    const record = state.continuousPages.get(target.unitIndex);
    if (record === undefined) return;
    revealContinuousPage(record, "center");
    if (object !== undefined) {
      const ancestors = await ancestorsForObject(officeDocument, object);
      showObjectDetails({ object, ancestors });
      const match = [...record.textLayer.querySelectorAll("[data-object-id]")]
        .find((item) => item.dataset.objectId === object.id);
      match?.scrollIntoView({ block: "center", inline: "center", behavior: "smooth" });
    }
  } else {
    await renderCurrentUnit();
    if (officeDocument !== state.officeDocument) return;
    const displayed = state.displayedFrame;
    if (displayed === undefined || displayed.unitIndex !== target.unitIndex) return;
    if (object !== undefined) {
      const ancestors = await ancestorsForObject(officeDocument, object);
      if (displayed !== state.displayedFrame) return;
      showHit({ object, ancestors }, displayed);
    } else if (target.bounds !== undefined) {
      const synthetic = {
        id: `diagnostic:${target.unitIndex}`,
        type: "shape",
        unitIndex: target.unitIndex,
        bounds: target.bounds,
        hidden: false,
        source: { format: officeDocument.info.format, kind: "diagnostic", part: diagnostic.part ?? "-", mapping: "approximate" },
      };
      showHit({ object: synthetic, ancestors: [] }, displayed);
    } else {
      resetHit();
    }
    (target.bounds === undefined ? ui.surfaceStage : ui.selectionBox)
      .scrollIntoView({ block: "center", inline: "center", behavior: "smooth" });
  }
  setStatus(`已定位诊断：${diagnostic.message}`, "ready");
}

function renderDiagnosticTable() {
  const diagnostics = state.diagnosticTab === "document"
    ? state.documentDiagnostics
    : state.renderDiagnostics;
  ui.diagnosticBody.replaceChildren();
  if (diagnostics.length === 0) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.className = "diagnostic-empty";
    cell.colSpan = 5;
    text(cell, "无诊断信息");
    row.append(cell);
    ui.diagnosticBody.append(row);
    return;
  }

  for (const diagnostic of diagnostics.slice(0, MAX_VISIBLE_DIAGNOSTICS)) {
    const row = document.createElement("tr");
    const locatable = diagnosticObjectId(diagnostic) !== undefined
      || diagnosticTarget(diagnostic) !== undefined
      || diagnostic.part !== undefined;
    row.dataset.locatable = String(locatable);
    if (locatable) {
      row.tabIndex = 0;
      row.setAttribute("aria-label", `定位诊断：${diagnostic.message}`);
      row.addEventListener("click", () => void focusDiagnostic(diagnostic));
      row.addEventListener("keydown", (event) => {
        if (event.key !== "Enter" && event.key !== " ") return;
        event.preventDefault();
        void focusDiagnostic(diagnostic);
      });
    }
    const severity = document.createElement("td");
    severity.className = `severity-${diagnostic.severity}`;
    severity.dataset.label = "级别";
    text(severity, diagnostic.severity);
    const code = document.createElement("td");
    code.dataset.label = "代码";
    const codeValue = document.createElement("code");
    text(codeValue, truncate(diagnostic.code, 200));
    code.append(codeValue);
    const message = document.createElement("td");
    message.dataset.label = "描述";
    text(message, truncate(diagnostic.message));
    const location = document.createElement("td");
    location.dataset.label = "位置";
    text(location, truncate(diagnosticLocation(diagnostic), 1_000));
    const details = document.createElement("td");
    details.className = "code-value";
    details.dataset.label = "保真度 / 阶段";
    text(details, diagnosticDetails(diagnostic));
    row.append(severity, code, message, location, details);
    ui.diagnosticBody.append(row);
  }

  if (diagnostics.length > MAX_VISIBLE_DIAGNOSTICS) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.className = "diagnostic-empty";
    cell.colSpan = 5;
    text(cell, `仅显示前 ${MAX_VISIBLE_DIAGNOSTICS} 条，另有 ${diagnostics.length - MAX_VISIBLE_DIAGNOSTICS} 条未展开`);
    row.append(cell);
    ui.diagnosticBody.append(row);
  }
}

function updateDiagnosticTabs() {
  text(ui.documentTab, `文档（${state.documentDiagnostics.length}）`);
  text(ui.renderTab, `渲染（${state.renderDiagnostics.length}）`);
  const documentSelected = state.diagnosticTab === "document";
  ui.documentTab.setAttribute("aria-selected", String(documentSelected));
  ui.renderTab.setAttribute("aria-selected", String(!documentSelected));
  ui.documentTab.tabIndex = documentSelected ? 0 : -1;
  ui.renderTab.tabIndex = documentSelected ? -1 : 0;
  text(ui.diagnosticStatus, `诊断 ${state.documentDiagnostics.length + state.renderDiagnostics.length}`);
  renderDiagnosticTable();
}

function selectDiagnosticTab(tab, options = {}) {
  state.diagnosticTab = tab;
  updateDiagnosticTabs();
  if (options.focus === true) {
    (tab === "document" ? ui.documentTab : ui.renderTab).focus();
  }
}

function resetDocumentMetadata() {
  text(ui.fileName, "未选择文件");
  text(ui.fileSize, "-");
  text(ui.documentFormat, "-");
  text(ui.documentKind, "-");
  setDetectedFormat("");
  setDocumentKind("");
  ui.fileName.title = "未选择文件";
}

function clearAll() {
  state.sessionRevision += 1;
  closeActiveSession();
  state.documentDiagnostics = [];
  state.renderDiagnostics = [];
  state.diagnosticTab = "document";
  ui.fileInput.value = "";
  resetDocumentMetadata();
  resetUnitControls();
  updateDiagnosticTabs();
  ui.clearButton.disabled = true;
  setDocumentOpen(false);
  if (state.engine === undefined) {
    setStatus("正在初始化 WebAssembly 引擎…", "loading");
  } else {
    setStatus("引擎已就绪，请选择一个 Office 或 PDF 文件。", "ready", "idle");
  }
}

function requestedSheetScale(unit) {
  if (state.zoomMode !== "fit") return Number(ui.zoomRange.value) / 100;
  const layout = sheetLayout(unit);
  const baseline = {
    width: Math.min(layout.columns.total, Math.max(1_200, ui.viewer.clientWidth - SURFACE_PADDING * 2)),
    height: Math.min(layout.rows.total, Math.max(800, ui.viewer.clientHeight - SURFACE_PADDING * 2)),
  };
  return fitScale(baseline);
}

function sheetViewport(unit, scale) {
  if (unit.type !== "sheet") return undefined;
  const layout = sheetLayout(unit);
  const visibleWidth = ui.viewer.clientWidth / scale;
  const visibleHeight = ui.viewer.clientHeight / scale;
  const width = Math.min(
    layout.columns.total,
    Math.max(1_200, visibleWidth * (1 + SHEET_VIEWPORT_OVERSCAN_RATIO)),
  );
  const height = Math.min(
    layout.rows.total,
    Math.max(800, visibleHeight * (1 + SHEET_VIEWPORT_OVERSCAN_RATIO)),
  );
  const sourceAnchor = state.viewportAnchor?.unitIndex === unit.index ? state.viewportAnchor.bounds : undefined;
  const anchor = sourceAnchor === undefined ? undefined : {
    x: layout.columns.map(sourceAnchor.x),
    y: layout.rows.map(sourceAnchor.y),
    width: layout.columns.map(sourceAnchor.x + sourceAnchor.width) - layout.columns.map(sourceAnchor.x),
    height: layout.rows.map(sourceAnchor.y + sourceAnchor.height) - layout.rows.map(sourceAnchor.y),
  };
  const origin = state.sheetViewportOrigin?.unitIndex === unit.index
    ? state.sheetViewportOrigin
    : undefined;
  const x = Math.max(0, Math.min(
    layout.columns.total - width,
    anchor === undefined ? origin?.x ?? 0 : anchor.x + anchor.width / 2 - width / 2,
  ));
  const y = Math.max(0, Math.min(
    layout.rows.total - height,
    anchor === undefined ? origin?.y ?? 0 : anchor.y + anchor.height / 2 - height / 2,
  ));
  return {
    x,
    y,
    width,
    height,
  };
}

function queueSheetViewportRender() {
  const unit = selectedUnit();
  const displayed = state.displayedFrame;
  if (unit?.type !== "sheet" || displayed === undefined || displayed.unitIndex !== unit.index) return;
  positionFrozenSurfaces(displayed.viewport, displayed.cssWidth / displayed.viewport.width);
  const scale = displayed.cssWidth / displayed.viewport.width;
  if (!Number.isFinite(scale) || scale <= 0) return;
  const layout = sheetLayout(unit);
  positionSheetHeaders(unit, displayed.viewport, scale);
  const visibleWidth = Math.min(layout.columns.total, ui.viewer.clientWidth / scale);
  const visibleHeight = Math.min(layout.rows.total, ui.viewer.clientHeight / scale);
  const visible = {
    x: Math.max(0, Math.min(
      layout.columns.total - visibleWidth,
      (ui.viewer.scrollLeft - SURFACE_PADDING - SHEET_ROW_HEADER_WIDTH) / scale,
    )),
    y: Math.max(0, Math.min(
      layout.rows.total - visibleHeight,
      (ui.viewer.scrollTop - SURFACE_PADDING - SHEET_COLUMN_HEADER_HEIGHT) / scale,
    )),
    width: visibleWidth,
    height: visibleHeight,
  };
  const marginX = Math.max(0, (displayed.viewport.width - visible.width) / 4);
  const marginY = Math.max(0, (displayed.viewport.height - visible.height) / 4);
  const covered = visible.x >= displayed.viewport.x + (displayed.viewport.x <= 0 ? 0 : marginX)
    && visible.y >= displayed.viewport.y + (displayed.viewport.y <= 0 ? 0 : marginY)
    && visible.x + visible.width <= displayed.viewport.x + displayed.viewport.width
      - (displayed.viewport.x + displayed.viewport.width >= layout.columns.total ? 0 : marginX)
    && visible.y + visible.height <= displayed.viewport.y + displayed.viewport.height
      - (displayed.viewport.y + displayed.viewport.height >= layout.rows.total ? 0 : marginY);
  if (covered) return;
  state.viewportAnchor = undefined;
  state.sheetViewportOrigin = {
    unitIndex: unit.index,
    x: Math.max(0, visible.x - (displayed.viewport.width - visible.width) / 2),
    y: Math.max(0, visible.y - (displayed.viewport.height - visible.height) / 2),
  };
  if (state.sheetScrollTimer !== undefined
    || state.renderRevision !== displayed.renderRevision) return;
  void renderCurrentUnit();
  state.sheetScrollTimer = setTimeout(() => {
    state.sheetScrollTimer = undefined;
    queueSheetViewportRender();
  }, SHEET_SCROLL_THROTTLE_MS);
}

function fitScale(bounds) {
  const inset = state.slideshow ? 0 : 56;
  const availableWidth = Math.max(1, ui.viewer.clientWidth - inset);
  const availableHeight = Math.max(1, ui.viewer.clientHeight - inset);
  const scale = Math.min(availableWidth / bounds.width, availableHeight / bounds.height);
  return state.slideshow ? scale : Math.min(1, scale);
}

function requestedScale(bounds) {
  if (state.zoomMode === "fit") return fitScale(bounds);
  const percent = Number(ui.zoomRange.value);
  return Number.isFinite(percent) && percent > 0 ? percent / 100 : 1;
}

function showZoom(scale, mode = state.zoomMode) {
  const percent = Math.max(25, Math.min(200, Math.round(scale * 100)));
  ui.zoomRange.value = String(percent);
  text(ui.zoomValue, mode === "fit" ? `${percent}%` : `${ui.zoomRange.value}%`);
}

function setZoomPercent(percent) {
  const value = Math.max(25, Math.min(200, Math.round(percent)));
  state.zoomMode = "custom";
  ui.zoomRange.value = String(value);
  text(ui.zoomValue, `${value}%`);
}

function applyZoomDelta(delta) {
  if (state.officeDocument === undefined) return;
  setZoomPercent(Number(ui.zoomRange.value) + delta);
  void renderCurrentUnit();
}

function fitCurrentUnit() {
  if (state.officeDocument === undefined) return;
  state.zoomMode = "fit";
  text(ui.zoomValue, "适应");
  void renderCurrentUnit();
}

async function toggleFullscreen() {
  try {
    if (document.fullscreenElement === null) await ui.app.requestFullscreen();
    else await document.exitFullscreen();
  } catch (cause) {
    setStatus(`切换全屏失败：${errorMessage(cause)}`, "error");
  }
}

function leaveSlideshowState() {
  if (!state.slideshow) return false;
  state.slideshow = false;
  ui.viewer.dataset.playing = "false";
  ui.slideshowControls.hidden = true;
  state.zoomMode = state.slideshowZoomMode ?? "fit";
  state.slideshowZoomMode = undefined;
  return true;
}

async function startSlideshow() {
  if (selectedUnit()?.type !== "slide" || state.slideshow) return;
  state.slideshowZoomMode = state.zoomMode;
  state.slideshow = true;
  state.zoomMode = "fit";
  ui.viewer.dataset.playing = "true";
  ui.slideshowControls.hidden = false;
  updateUnitButtons();
  try {
    await ui.viewer.requestFullscreen();
    ui.viewer.focus();
    await renderCurrentUnit();
  } catch (cause) {
    leaveSlideshowState();
    void renderCurrentUnit();
    setStatus(`播放幻灯片失败：${errorMessage(cause)}`, "error");
  }
}

async function stopSlideshow(options = {}) {
  if (!leaveSlideshowState()) return;
  if (options.exitFullscreen !== false && document.fullscreenElement === ui.viewer) {
    try {
      await document.exitFullscreen();
    } catch (cause) {
      setStatus(`退出幻灯片播放失败：${errorMessage(cause)}`, "error");
    }
  }
  if (state.officeDocument !== undefined) void renderCurrentUnit();
  if (options.restoreFocus !== false && !ui.slideshowButton.hidden) ui.slideshowButton.focus();
}

async function printDocument() {
  const officeDocument = state.officeDocument;
  if (officeDocument === undefined || ui.printButton.disabled) return;
  ui.printButton.disabled = true;
  const printFrame = document.createElement("iframe");
  printFrame.title = "打印文档";
  printFrame.style.cssText = "position:fixed;right:100%;bottom:100%;width:1px;height:1px;border:0";
  document.body.append(printFrame);
  const urls = [];
  const clear = () => {
    for (const url of urls) URL.revokeObjectURL(url);
    printFrame.remove();
  };
  let printStarted = false;
  try {
    const target = printFrame.contentDocument;
    if (target === null) throw new Error("浏览器无法创建打印文档");
    const style = target.createElement("style");
    style.textContent = `
      @page { margin: 0; }
      html, body { margin: 0; background: #ffffff; }
      section { display: flex; width: 100%; box-sizing: border-box; justify-content: center; break-after: page; break-inside: avoid-page; }
      section:last-child { break-after: auto; }
      img { display: block; width: auto; height: auto; max-width: 99vw; max-height: 99vh; object-fit: contain; }
    `;
    target.head.append(style);
    for (const unit of officeDocument.info.units) {
      if (officeDocument !== state.officeDocument) throw new DOMException("Superseded", "AbortError");
      const sheetSizes = unit.type === "sheet" ? sheetSizeRequest(unit) : undefined;
      const printPages = unit.type === "sheet" ? sheetPrintPages(unit, sheetSizes) : [undefined];
      for (const sheetPage of printPages) {
        const page = target.createElement("section");
        if (sheetPage !== undefined) {
          page.style.width = `${sheetPage.paper.width}px`;
          page.style.height = `${sheetPage.paper.height}px`;
          page.style.padding = `${sheetPage.margins.top}px ${sheetPage.margins.right}px ${sheetPage.margins.bottom}px ${sheetPage.margins.left}px`;
        }
        let frame;
        try {
          const width = sheetPage?.viewport.width ?? unit.width;
          const height = sheetPage?.viewport.height ?? unit.height;
          const scale = Math.max(.01, Math.min(2, PRINT_RENDER_LONG_EDGE / Math.max(1, width, height)));
          frame = await officeDocument.render({
            unitIndex: unit.index,
            ...(sheetPage === undefined ? {} : { viewport: sheetPage.viewport }),
            scale,
            pixelRatio: 1,
            background: "#ffffff",
            ...(sheetSizes === undefined ? {} : { sheetSizes }),
          }, { priority: "interactive", supersedeKey: "print" });
          const canvas = target.createElement("canvas");
          canvas.width = frame.pixelWidth;
          canvas.height = frame.pixelHeight;
          const context = canvas.getContext("2d");
          if (context === null) throw new Error("浏览器无法创建打印 Canvas 2D 上下文");
          context.drawImage(frame.bitmap, 0, 0);
          const blob = await new Promise((resolve, reject) => canvas.toBlob((value) => {
            if (value === null) reject(new Error("打印图像编码失败"));
            else resolve(value);
          }, "image/png"));
          const url = URL.createObjectURL(blob);
          urls.push(url);
          const image = target.createElement("img");
          image.alt = "";
          image.src = url;
          if (sheetPage !== undefined) {
            image.style.width = `${sheetPage.viewport.width * sheetPage.scale}px`;
            image.style.height = `${sheetPage.viewport.height * sheetPage.scale}px`;
          }
          await image.decode();
          page.append(image);
        } catch (cause) {
          if (cause instanceof DOMException && cause.name === "AbortError") throw cause;
          page.textContent = `${unitLabel(unit)}打印渲染失败：${errorMessage(cause)}`;
        } finally {
          frame?.bitmap.close();
        }
        target.body.append(page);
      }
    }
    const printWindow = printFrame.contentWindow;
    if (printWindow === null) throw new Error("浏览器无法创建打印窗口");
    printWindow.addEventListener("afterprint", clear, { once: true });
    printStarted = true;
    try {
      printWindow.focus();
      printWindow.print();
    } catch (cause) {
      printStarted = false;
      throw cause;
    }
  } catch (cause) {
    if (!(cause instanceof DOMException && cause.name === "AbortError")) {
      setStatus(`打印失败：${errorMessage(cause)}`, "error");
    }
  } finally {
    if (!printStarted) clear();
    if (officeDocument === state.officeDocument) ui.printButton.disabled = false;
  }
}

function updateFullscreenButton() {
  const active = document.fullscreenElement === ui.app;
  ui.fullscreenButton.setAttribute("aria-pressed", String(active));
  text(ui.fullscreenButton, active ? "退出全屏" : "全屏");
}

function handleFullscreenChange() {
  updateFullscreenButton();
  if (state.slideshow && document.fullscreenElement !== ui.viewer) {
    void stopSlideshow({ exitFullscreen: false });
  }
}

function renderPixelEstimate(bounds, scale, pixelRatio) {
  return Math.ceil(bounds.width * scale * pixelRatio) * Math.ceil(bounds.height * scale * pixelRatio);
}

function frozenPaneRequests(unit, viewport) {
  if (unit.type !== "sheet") return [];
  const layout = sheetLayout(unit);
  const frozenWidth = unit.frozenColumns > 0
    ? Math.min(layout.columns.offset(unit.frozenColumns), layout.columns.total, viewport.width)
    : 0;
  const frozenHeight = unit.frozenRows > 0
    ? Math.min(layout.rows.offset(unit.frozenRows), layout.rows.total, viewport.height)
    : 0;
  const requests = [];
  if (frozenHeight > 0) {
    requests.push({
      pane: "top",
      viewport: { x: viewport.x, y: 0, width: viewport.width, height: frozenHeight },
    });
  }
  if (frozenWidth > 0) {
    requests.push({
      pane: "left",
      viewport: { x: 0, y: viewport.y, width: frozenWidth, height: viewport.height },
    });
  }
  if (frozenWidth > 0 && frozenHeight > 0) {
    requests.push({
      pane: "corner",
      viewport: { x: 0, y: 0, width: frozenWidth, height: frozenHeight },
    });
  }
  return requests;
}

function frozenSurface(pane) {
  if (pane === "top") return ui.frozenTop;
  if (pane === "left") return ui.frozenLeft;
  return ui.frozenCorner;
}

function positionFrozenSurfaces(viewport, scale) {
  if (!Number.isFinite(scale) || scale <= 0) return;
  const stageLeft = SURFACE_PADDING + SHEET_ROW_HEADER_WIDTH + viewport.x * scale;
  const stageTop = SURFACE_PADDING + SHEET_COLUMN_HEADER_HEIGHT + viewport.y * scale;
  ui.frozenTop.style.left = `${stageLeft}px`;
  ui.frozenTop.style.top = `${Math.max(stageTop, ui.viewer.scrollTop)}px`;
  ui.frozenLeft.style.left = `${Math.max(stageLeft, ui.viewer.scrollLeft)}px`;
  ui.frozenLeft.style.top = `${stageTop}px`;
  ui.frozenCorner.style.left = `${Math.max(stageLeft, ui.viewer.scrollLeft)}px`;
  ui.frozenCorner.style.top = `${Math.max(stageTop, ui.viewer.scrollTop)}px`;
}

async function renderFrozenPaneFrames(
  officeDocument,
  unit,
  viewport,
  scale,
  pixelRatio,
) {
  const requests = frozenPaneRequests(unit, viewport);
  const results = await Promise.allSettled(requests.map(async (request) => ({
    request,
    frame: await officeDocument.render({
      unitIndex: unit.index,
      viewport: request.viewport,
      scale,
      pixelRatio,
      background: "#ffffff",
      sheetSizes: sheetSizeRequest(unit),
    }, {
      priority: "visible",
      supersedeKey: `frozen:${request.pane}`,
    }),
  })));
  const frames = results
    .filter((result) => result.status === "fulfilled")
    .map((result) => result.value);
  const failed = results.find((result) => result.status === "rejected");
  if (failed !== undefined) {
    for (const { frame } of frames) frame.bitmap.close();
    throw failed.reason;
  }
  return {
    frames,
    diagnostics: frames.flatMap(({ frame }) => frame.diagnostics),
    renderedObjectCount: frames.reduce((total, { frame }) => total + frame.renderedObjectCount, 0),
  };
}

function drawFrozenPaneFrames(frozen, viewport, scale) {
  clearFrozenSurfaces();
  for (const { request, frame } of frozen.frames) {
    const target = frozenSurface(request.pane);
    const context = target.getContext("2d");
    if (context === null) throw new Error("浏览器无法创建冻结窗格 Canvas 2D 上下文");
    target.width = frame.pixelWidth;
    target.height = frame.pixelHeight;
    target.style.width = `${frame.viewport.width * scale}px`;
    target.style.height = `${frame.viewport.height * scale}px`;
    context.clearRect(0, 0, frame.pixelWidth, frame.pixelHeight);
    context.drawImage(frame.bitmap, 0, 0);
    target.dataset.visible = "true";
  }
  if (frozen.frames.length === 0) return;
  ui.surfaceWrap.dataset.frozen = "true";
  positionFrozenSurfaces(viewport, scale);
}

function safeVisibleLink(value) {
  const candidate = value.trim();
  if (candidate.length === 0 || candidate.length > 2_048 || /\s/u.test(candidate)) return undefined;
  try {
    const parsed = new URL(candidate);
    return ["http:", "https:", "mailto:", "tel:"].includes(parsed.protocol)
      ? parsed.href
      : undefined;
  } catch {
    return undefined;
  }
}

function headerItem(axis, index, start, size, label) {
  const item = document.createElement("div");
  item.className = "sheet-header-item";
  if (axis === "column") {
    item.style.left = `${start}px`;
    item.style.width = `${size}px`;
    item.style.height = "100%";
  } else {
    item.style.top = `${start}px`;
    item.style.width = "100%";
    item.style.height = `${size}px`;
  }
  text(item, label);
  const resizer = document.createElement("span");
  resizer.className = "sheet-header-resizer";
  resizer.dataset.axis = axis;
  resizer.dataset.index = String(index);
  resizer.tabIndex = 0;
  resizer.setAttribute("role", "separator");
  resizer.setAttribute("aria-label", axis === "column" ? `调整 ${label} 列宽` : `调整第 ${label} 行高`);
  resizer.setAttribute("aria-orientation", axis === "column" ? "vertical" : "horizontal");
  item.append(resizer);
  return item;
}

function renderSheetHeaders(unit, viewport, scale) {
  if (unit.type !== "sheet") return;
  const layout = sheetLayout(unit);
  const columns = document.createDocumentFragment();
  const startColumn = layout.columns.indexAt(viewport.x);
  const endColumn = layout.columns.indexAt(viewport.x + viewport.width);
  for (let index = startColumn; index >= 0 && index <= endColumn; index += 1) {
    columns.append(headerItem(
      "column",
      index,
      (layout.columns.offset(index) - viewport.x) * scale,
      layout.columns.size(index) * scale,
      columnLabel(index),
    ));
  }
  const rows = document.createDocumentFragment();
  const startRow = layout.rows.indexAt(viewport.y);
  const endRow = layout.rows.indexAt(viewport.y + viewport.height);
  for (let index = startRow; index >= 0 && index <= endRow; index += 1) {
    rows.append(headerItem(
      "row",
      index,
      (layout.rows.offset(index) - viewport.y) * scale,
      layout.rows.size(index) * scale,
      String(index + 1),
    ));
  }
  ui.sheetColumnHeaders.replaceChildren(columns);
  ui.sheetRowHeaders.replaceChildren(rows);
  ui.sheetColumnHeaders.style.width = `${viewport.width * scale}px`;
  ui.sheetRowHeaders.style.height = `${viewport.height * scale}px`;
  positionSheetHeaders(unit, viewport, scale);
}

function renderSheetPageBreaks(unit, layout, scale) {
  ui.sheetPageBreaks.replaceChildren();
  ui.sheetPageBreaks.hidden = unit.printSettings?.viewMode === undefined;
  if (ui.sheetPageBreaks.hidden) return;
  ui.sheetPageBreaks.style.left = `${SURFACE_PADDING + SHEET_ROW_HEADER_WIDTH}px`;
  ui.sheetPageBreaks.style.top = `${SURFACE_PADDING + SHEET_COLUMN_HEADER_HEIGHT}px`;
  ui.sheetPageBreaks.style.width = `${layout.columns.total * scale}px`;
  ui.sheetPageBreaks.style.height = `${layout.rows.total * scale}px`;
  const pages = sheetPrintPages(unit, sheetSizeRequest(unit));
  const columns = new Set(pages.flatMap(({ viewport }) => [viewport.x, viewport.x + viewport.width]));
  const rows = new Set(pages.flatMap(({ viewport }) => [viewport.y, viewport.y + viewport.height]));
  for (const [axis, boundaries, total] of [
    ["column", columns, layout.columns.total],
    ["row", rows, layout.rows.total],
  ]) {
    for (const boundary of boundaries) {
      if (boundary <= 0 || boundary >= total) continue;
      const line = document.createElement("span");
      line.className = "sheet-page-break";
      line.dataset.axis = axis;
      if (axis === "column") line.style.left = `${boundary * scale}px`;
      else line.style.top = `${boundary * scale}px`;
      ui.sheetPageBreaks.append(line);
    }
  }
}

function positionSheetHeaders(unit, viewport, scale) {
  if (unit.type !== "sheet") return;
  const stageLeft = SHEET_ROW_HEADER_WIDTH + viewport.x * scale;
  const stageTop = SHEET_COLUMN_HEADER_HEIGHT + viewport.y * scale;
  ui.sheetColumnHeaders.style.transform = `translate3d(${stageLeft}px, 0, 0)`;
  ui.sheetRowHeaders.style.transform = `translate3d(0, ${stageTop}px, 0)`;
  ui.sheetHeaderCorner.style.transform = "translate3d(0, 0, 0)";
}

function commitSheetSize(axis, index, size) {
  const unit = selectedUnit();
  if (unit?.type !== "sheet") return;
  const overrides = sheetOverrides(unit);
  (axis === "column" ? overrides.columns : overrides.rows).set(index, size);
  state.viewportAnchor = undefined;
  state.sheetViewportOrigin = undefined;
  void renderCurrentUnit();
}

function beginSheetResize(event, target) {
  const unit = selectedUnit();
  const displayed = state.displayedFrame;
  const axis = target.dataset.axis;
  const index = Number(target.dataset.index);
  if (unit?.type !== "sheet" || displayed?.unitIndex !== unit.index
    || (axis !== "row" && axis !== "column") || !Number.isInteger(index)) return;
  event.preventDefault();
  target.setPointerCapture(event.pointerId);
  const scale = displayed.cssWidth / displayed.viewport.width;
  const layout = sheetLayout(unit);
  const startSize = (axis === "column" ? layout.columns : layout.rows).size(index);
  state.sheetResize = {
    axis,
    index,
    pointerId: event.pointerId,
    start: axis === "column" ? event.clientX : event.clientY,
    startSize,
    size: startSize,
    scale,
  };
  ui.sheetResizeGuide.dataset.axis = axis;
  ui.sheetResizeGuide.dataset.visible = "true";
  updateSheetResize(event);
}

function updateSheetResize(event) {
  const resize = state.sheetResize;
  if (resize === undefined || event.pointerId !== resize.pointerId) return;
  const coordinate = resize.axis === "column" ? event.clientX : event.clientY;
  const minimum = resize.axis === "column" ? MIN_COLUMN_WIDTH : MIN_ROW_HEIGHT;
  resize.size = Math.max(minimum, resize.startSize + (coordinate - resize.start) / resize.scale);
  const wrapBounds = ui.surfaceWrap.getBoundingClientRect();
  if (resize.axis === "column") {
    ui.sheetResizeGuide.style.width = "1px";
    ui.sheetResizeGuide.style.left = `${event.clientX - wrapBounds.left}px`;
    ui.sheetResizeGuide.style.top = `${ui.viewer.scrollTop}px`;
    ui.sheetResizeGuide.style.height = `${ui.viewer.clientHeight}px`;
  } else {
    ui.sheetResizeGuide.style.height = "1px";
    ui.sheetResizeGuide.style.left = `${ui.viewer.scrollLeft}px`;
    ui.sheetResizeGuide.style.top = `${event.clientY - wrapBounds.top}px`;
    ui.sheetResizeGuide.style.width = `${ui.viewer.clientWidth}px`;
  }
}

function finishSheetResize(event) {
  const resize = state.sheetResize;
  if (resize === undefined || event.pointerId !== resize.pointerId) return;
  state.sheetResize = undefined;
  ui.sheetResizeGuide.dataset.visible = "false";
  commitSheetSize(resize.axis, resize.index, resize.size);
}

async function autoFitSheetSize(event, target) {
  const unit = selectedUnit();
  const axis = target.dataset.axis;
  const index = Number(target.dataset.index);
  if (unit?.type !== "sheet" || (axis !== "row" && axis !== "column") || !Number.isInteger(index)) return;
  event.preventDefault();
  const overrides = sheetOverrides(unit);
  if (axis === "row") {
    overrides.rows.delete(index);
  } else {
    const officeDocument = state.officeDocument;
    if (officeDocument === undefined) return;
    const sourceColumns = axisLayout(unit.columnAxis, unit.columns, new Map());
    const x = sourceColumns.offset(index);
    const sourceWidth = sourceColumns.size(index);
    try {
      const objects = await officeDocument.listObjects({
        unitIndex: unit.index,
        types: ["cell"],
        textOnly: true,
        viewport: { x, y: 0, width: sourceWidth, height: unit.height },
        limit: 20_000,
      });
      if (officeDocument !== state.officeDocument || unit !== selectedUnit()) return;
      const context = document.createElement("canvas").getContext("2d");
      if (context === null) return;
      const cells = objects
        .filter((object) => object.text !== undefined && object.bounds.width <= sourceWidth + 0.01)
        .map((object) => ({
          text: object.text,
          width: sourceWidth,
          wrap: object.wrapText ?? false,
          ...(object.fontRuns === undefined ? {} : { fontRuns: object.fontRuns }),
        }));
      if (cells.length === 0) overrides.columns.delete(index);
      else overrides.columns.set(index, sheetAutoFitColumnWidth(context, cells, MIN_COLUMN_WIDTH));
    } catch {
      if (officeDocument !== state.officeDocument) return;
      overrides.columns.delete(index);
    }
  }
  state.viewportAnchor = undefined;
  state.sheetViewportOrigin = undefined;
  void renderCurrentUnit();
}

function overlaps(left, right) {
  return left.x < right.x + right.width
    && left.x + left.width > right.x
    && left.y < right.y + right.height
    && left.y + left.height > right.y;
}

function displayedObjectBounds(object, displayed) {
  const unit = displayed.officeDocument.info.units[displayed.unitIndex];
  if (unit?.type !== "sheet") return object.bounds;
  const layout = sheetLayout(unit);
  const x = layout.columns.map(object.bounds.x);
  const y = layout.rows.map(object.bounds.y);
  return {
    x,
    y,
    width: layout.columns.map(object.bounds.x + object.bounds.width) - x,
    height: layout.rows.map(object.bounds.y + object.bounds.height) - y,
  };
}

async function renderTextLayer(
  displayed,
  target = ui.textLayer,
  isCurrent = () => displayed === state.displayedFrame && displayed.officeDocument === state.officeDocument,
) {
  const officeDocument = displayed.officeDocument;
  const unit = officeDocument.info.units[displayed.unitIndex];
  let objectViewport = displayed.viewport;
  if (unit?.type === "sheet") {
    const layout = sheetLayout(unit);
    const right = layout.columns.unmap(displayed.viewport.x + displayed.viewport.width);
    const bottom = layout.rows.unmap(displayed.viewport.y + displayed.viewport.height);
    const x = layout.columns.unmap(displayed.viewport.x);
    const y = layout.rows.unmap(displayed.viewport.y);
    objectViewport = { x, y, width: right - x, height: bottom - y };
  }
  const objects = await officeDocument.listObjects({
    unitIndex: displayed.unitIndex,
    types: ["text-box", "paragraph", "shape", "cell"],
    textOnly: true,
    viewport: objectViewport,
    limit: MAX_TEXT_LAYER_OBJECTS,
  });
  if (!isCurrent()) return;
  const fragment = document.createDocumentFragment();
  const visibleObjects = new Map();
  const scaleX = displayed.cssWidth / displayed.viewport.width;
  const scaleY = displayed.cssHeight / displayed.viewport.height;
  for (const object of objects) {
    const bounds = displayedObjectBounds(object, displayed);
    if (!overlaps(bounds, displayed.viewport)) continue;
    const value = object.text;
    if (value === undefined || value.length === 0) continue;
    const key = `${bounds.x}:${bounds.y}:${bounds.width}:${bounds.height}:${value}`;
    const existing = visibleObjects.get(key);
    if (existing === undefined
      || (existing.object.fontRuns?.length ?? 0) === 0 && (object.fontRuns?.length ?? 0) > 0) {
      visibleObjects.set(key, { object, bounds });
    }
  }
  for (const { object, bounds } of visibleObjects.values()) {
    const value = object.text;
    if (value === undefined || value.length === 0) continue;
    const href = safeVisibleLink(value);
    const item = document.createElement(href === undefined ? "span" : "a");
    item.className = `text-layer-item${href === undefined ? "" : " text-layer-link"}`;
    item.dataset.objectId = object.id;
    const textLayout = object.textLayout;
    const positioned = textLayout !== undefined
      && textLayout.orientation === "horizontal"
      && textLayout.rotationDegrees === 0
      && textLayout.columnCount === 1;
    const leftInset = positioned ? textLayout.insetLeft + textLayout.marginLeft : 0;
    const rightInset = positioned ? textLayout.insetRight + textLayout.marginRight : 0;
    const topInset = positioned ? textLayout.insetTop : 0;
    const bottomInset = positioned ? textLayout.insetBottom : 0;
    item.style.left = `${(bounds.x + leftInset - displayed.viewport.x) * scaleX}px`;
    item.style.top = `${(bounds.y + topInset - displayed.viewport.y) * scaleY}px`;
    item.style.width = `${Math.max(1, (bounds.width - leftInset - rightInset) * scaleX)}px`;
    item.style.height = `${Math.max(1, (bounds.height - topInset - bottomInset) * scaleY)}px`;
    if (positioned) {
      item.style.justifyContent = textLayout.verticalAlign === "center"
        ? "center" : textLayout.verticalAlign === "bottom" ? "flex-end" : "flex-start";
      item.style.whiteSpace = textLayout.wrap ? "pre-wrap" : "pre";
      item.style.overflowWrap = textLayout.wrap ? "anywhere" : "normal";
      item.style.overflow = textLayout.horizontalOverflow === "overflow"
        && textLayout.verticalOverflow === "overflow" ? "visible" : "hidden";
    }
    const fontRuns = object.fontRuns ?? [];
    const fontSize = dominantDocumentFontSize(fontRuns);
    const renderedFontSize = fontSize === undefined
      ? Math.max(6, Math.min(20, bounds.height * scaleY / 1.2))
      : Math.max(1, fontSize * scaleY);
    item.style.fontSize = `${renderedFontSize}px`;
    const content = document.createElement("span");
    content.className = "text-layer-content";
    const textAlign = positioned ? textLayout.textAlign : undefined;
    const justified = textAlign !== undefined && !["start", "center", "end"].includes(textAlign);
    if (positioned && textLayout.firstLineIndent !== 0) {
      content.style.textIndent = `${textLayout.firstLineIndent * scaleX}px`;
    }
    if (positioned && textLayout.lineHeight !== undefined) {
      const lineHeight = textLayout.lineHeight * scaleY;
      content.style.lineHeight = `${lineHeight}px`;
      const baselineOffset = Math.max(0, lineHeight - renderedFontSize * 1.2) / 2;
      if (baselineOffset > 0) content.style.transform = `translateY(${baselineOffset}px)`;
    }
    if (positioned && !textLayout.wrap) content.style.width = "max-content";
    else if (textAlign !== undefined) {
      content.style.width = justified ? "100%" : "fit-content";
      content.style.maxWidth = "100%";
    }
    if (textAlign !== undefined) {
      content.style.alignSelf = justified ? (textLayout?.wrap === true ? "stretch" : "start") : textAlign;
      content.style.textAlign = justified ? "justify" : textAlign;
    }
    let cursor = 0;
    for (const run of fontRuns) {
      if (run.start < cursor || run.end <= run.start || run.end > value.length) continue;
      if (run.start > cursor) content.append(document.createTextNode(value.slice(cursor, run.start)));
      const span = document.createElement("span");
      span.className = "text-font-run";
      span.dataset.authoredFamily = run.authoredFamily;
      span.dataset.renderedFamily = run.renderedFamily;
      span.dataset.fontSource = run.source;
      span.style.fontFamily = cssFontFamily(run.renderedFamily);
      if (run.fontSize !== undefined) span.style.fontSize = `${Math.max(1, run.fontSize * scaleY)}px`;
      if (run.bold !== undefined) span.style.fontWeight = run.bold ? "700" : "400";
      if (run.italic !== undefined) span.style.fontStyle = run.italic ? "italic" : "normal";
      text(span, value.slice(run.start, run.end));
      content.append(span);
      cursor = run.end;
    }
    if (cursor < value.length) content.append(document.createTextNode(value.slice(cursor)));
    item.append(content);
    if (href !== undefined) {
      item.href = href;
      item.target = "_blank";
      item.rel = "noopener noreferrer";
      item.setAttribute("aria-label", `打开链接：${value}`);
    }
    fragment.append(item);
  }
  target.replaceChildren(fragment);
}

function placeRenderedSurface(unit, viewport, scale, cssWidth, cssHeight) {
  const sheet = unit.type === "sheet";
  ui.surfaceWrap.dataset.sheet = String(sheet);
  ui.surfaceStage.style.width = `${cssWidth}px`;
  ui.surfaceStage.style.height = `${cssHeight}px`;
  ui.speakerNotesPanel.style.width = unit.type === "slide" ? `${cssWidth}px` : "";
  if (!sheet) {
    ui.surfaceWrap.style.width = "";
    ui.surfaceWrap.style.height = "";
    ui.surfaceStage.style.left = "";
    ui.surfaceStage.style.top = "";
    return;
  }
  const maximumCssDimension = 16_000_000;
  const layout = sheetLayout(unit);
  const virtualWidth = Math.min(maximumCssDimension, layout.columns.total * scale + SURFACE_PADDING * 2 + SHEET_ROW_HEADER_WIDTH);
  const virtualHeight = Math.min(maximumCssDimension, layout.rows.total * scale + SURFACE_PADDING * 2 + SHEET_COLUMN_HEADER_HEIGHT);
  ui.surfaceWrap.style.width = `${Math.max(ui.viewer.clientWidth, virtualWidth)}px`;
  ui.surfaceWrap.style.height = `${Math.max(ui.viewer.clientHeight, virtualHeight)}px`;
  ui.surfaceStage.style.left = `${SURFACE_PADDING + SHEET_ROW_HEADER_WIDTH + viewport.x * scale}px`;
  ui.surfaceStage.style.top = `${SURFACE_PADDING + SHEET_COLUMN_HEADER_HEIGHT + viewport.y * scale}px`;
  renderSheetHeaders(unit, viewport, scale);
  renderSheetPageBreaks(unit, layout, scale);
}

function continuousPageScale(unit) {
  if (state.zoomMode !== "fit") return Number(ui.zoomRange.value) / 100;
  return Math.min(1, Math.max(0.1, (ui.viewer.clientWidth - 96) / unit.width));
}

function releaseContinuousPage(record) {
  if (record.visible || record.rendering || !record.rendered) return;
  state.continuousPixels -= record.pixels;
  record.pixels = 0;
  record.rendered = false;
  record.canvas.width = 1;
  record.canvas.height = 1;
  record.textLayer.replaceChildren();
  record.element.dataset.rendered = "false";
}

function evictContinuousPixels() {
  if (state.retainContinuousPages || state.continuousPixels <= CONTINUOUS_CACHE_PIXELS) return;
  const candidates = [...state.continuousPages.values()]
    .filter((record) => !record.visible && record.rendered)
    .sort((left, right) => left.lastUsed - right.lastUsed);
  for (const record of candidates) {
    releaseContinuousPage(record);
    if (state.continuousPixels <= CONTINUOUS_CACHE_PIXELS) break;
  }
}

async function renderContinuousPage(unitIndex, revision = state.continuousRevision) {
  const officeDocument = state.officeDocument;
  const record = state.continuousPages.get(unitIndex);
  const unit = officeDocument?.info.units[unitIndex];
  if (officeDocument === undefined
    || unit?.type !== "page"
    || record === undefined
    || record.rendered
    || record.rendering
    || !state.continuous
    || revision !== state.continuousRevision) return;
  record.rendering = true;
  let frame;
  try {
    const pixelRatio = Math.min(Math.max(window.devicePixelRatio || 1, 1), 2);
    frame = await officeDocument.render({
      unitIndex,
      scale: record.scale,
      pixelRatio,
      background: "#ffffff",
    }, {
      priority: "visible",
      supersedeKey: `continuous:${unitIndex}`,
    });
    if (revision !== state.continuousRevision
      || !state.continuous
      || officeDocument !== state.officeDocument) return;
    record.canvas.width = frame.pixelWidth;
    record.canvas.height = frame.pixelHeight;
    const context = record.canvas.getContext("2d");
    if (context === null) throw new Error("浏览器无法创建连续页面 Canvas 2D 上下文");
    context.drawImage(frame.bitmap, 0, 0);
    const displayed = {
      officeDocument,
      unitIndex,
      viewport: frame.viewport,
      cssWidth: unit.width * record.scale,
      cssHeight: unit.height * record.scale,
      pixelWidth: frame.pixelWidth,
      pixelHeight: frame.pixelHeight,
    };
    await renderTextLayer(
      displayed,
      record.textLayer,
      () => revision === state.continuousRevision
        && state.continuous
        && officeDocument === state.officeDocument,
    );
    if (revision !== state.continuousRevision || !state.continuous) return;
    record.rendered = true;
    state.continuousPixels -= record.pixels;
    record.pixels = frame.pixelWidth * frame.pixelHeight;
    record.lastUsed = performance.now();
    state.continuousPixels += record.pixels;
    record.element.dataset.rendered = "true";
    evictContinuousPixels();
    if (frame.diagnostics.length > 0) {
      state.renderDiagnostics = [...state.renderDiagnostics, ...frame.diagnostics];
      updateDiagnosticTabs();
    }
  } catch (cause) {
    if (revision !== state.continuousRevision || !state.continuous) return;
    if (!(cause instanceof OfficeEngineError && cause.code === "OPERATION_ABORTED")) {
      record.element.dataset.error = "true";
      record.element.setAttribute("aria-label", `${unitLabel(unit)}渲染失败：${errorMessage(cause)}`);
    }
  } finally {
    record.rendering = false;
    frame?.bitmap.close();
  }
}

async function buildContinuousPages(expectedSession = state.sessionRevision) {
  const officeDocument = state.officeDocument;
  if (officeDocument === undefined
    || officeDocument.info.kind !== "text"
    || !state.continuous
    || expectedSession !== state.sessionRevision) return;
  const revision = ++state.continuousRevision;
  state.renderRevision += 1;
  state.displayedFrame = undefined;
  state.continuousObserver?.disconnect();
  state.continuousNavigationObserver?.disconnect();
  state.continuousObserver = undefined;
  state.continuousNavigationObserver = undefined;
  state.continuousPages.clear();
  state.continuousPixels = 0;
  ui.continuousView.replaceChildren();
  ui.textLayer.replaceChildren();
  ui.selectionBox.dataset.visible = "false";
  ui.surfaceStage.hidden = true;
  ui.continuousView.hidden = false;
  ui.surfaceWrap.dataset.sheet = "false";
  ui.surfaceWrap.dataset.continuous = "true";
  ui.surfaceWrap.style.width = "";
  ui.surfaceWrap.style.height = "";
  ui.surfaceWrap.dataset.visible = "true";
  ui.emptyState.hidden = true;
  state.renderDiagnostics = [];
  updateDiagnosticTabs();

  const fragment = document.createDocumentFragment();
  for (const unit of officeDocument.info.units) {
    if (unit.type !== "page") continue;
    const scale = continuousPageScale(unit);
    const page = document.createElement("section");
    page.className = "continuous-page";
    page.dataset.unitIndex = String(unit.index);
    page.dataset.rendered = "false";
    page.style.width = `${unit.width * scale}px`;
    page.style.height = `${unit.height * scale}px`;
    page.setAttribute("aria-label", unitLabel(unit));
    const canvas = document.createElement("canvas");
    canvas.width = 1;
    canvas.height = 1;
    canvas.setAttribute("aria-label", `${unitLabel(unit)}渲染结果`);
    const textLayer = document.createElement("div");
    textLayer.className = "text-layer";
    textLayer.setAttribute("aria-label", `${unitLabel(unit)}可选择文本层`);
    page.append(canvas, textLayer);
    state.continuousPages.set(unit.index, {
      element: page,
      canvas,
      textLayer,
      scale,
      rendered: false,
      rendering: false,
      visible: false,
      pixels: 0,
      lastUsed: 0,
    });
    fragment.append(page);
  }
  ui.continuousView.append(fragment);

  if (typeof IntersectionObserver === "function") {
    state.continuousObserver = new IntersectionObserver((entries) => {
      for (const entry of entries) {
        const unitIndex = Number(entry.target.dataset.unitIndex);
        const record = state.continuousPages.get(unitIndex);
        if (record === undefined) continue;
        record.visible = entry.isIntersecting;
        if (entry.isIntersecting) {
          record.lastUsed = performance.now();
          void renderContinuousPage(unitIndex, revision);
        }
      }
      evictContinuousPixels();
    }, { root: ui.viewer, rootMargin: "400px 0px", threshold: 0.01 });
    for (const record of state.continuousPages.values()) state.continuousObserver.observe(record.element);

    state.continuousNavigationObserver = new IntersectionObserver((entries) => {
      if (revision !== state.continuousRevision || !state.continuous) return;
      const entry = entries.find((candidate) => candidate.isIntersecting);
      const unitIndex = Number(entry?.target.dataset.unitIndex);
      if (Number.isInteger(unitIndex)) selectContinuousUnit(unitIndex);
    }, { root: ui.viewer, rootMargin: "-45% 0px -45% 0px", threshold: 0.01 });
    for (const record of state.continuousPages.values()) {
      state.continuousNavigationObserver.observe(record.element);
    }
  }
  const selected = selectedUnit()?.index ?? 0;
  const selectedRecord = state.continuousPages.get(selected);
  const nextRecord = state.continuousPages.get(selected + 1);
  if (selectedRecord !== undefined) selectedRecord.visible = true;
  if (nextRecord !== undefined) nextRecord.visible = true;
  await Promise.all([
    renderContinuousPage(selected, revision),
    renderContinuousPage(selected + 1, revision),
  ]);
  if (revision !== state.continuousRevision || !state.continuous) return;
  state.continuousPages.get(selected)?.element.scrollIntoView({ block: "start" });
  showZoom(continuousPageScale(officeDocument.info.units[selected]), state.zoomMode);
  setStatus(`连续页面已就绪 · ${state.continuousPages.size} 页按视口加载。`, "ready");
}

function setContinuousMode(enabled) {
  const supported = state.officeDocument?.info.kind === "text";
  state.continuous = supported && enabled;
  ui.continuousButton.setAttribute("aria-pressed", String(state.continuous));
  text(ui.continuousButton, state.continuous ? "单页模式" : "连续页面");
  if (state.continuous) void buildContinuousPages();
  else {
    state.continuousRevision += 1;
    state.continuousObserver?.disconnect();
    state.continuousObserver = undefined;
    state.continuousNavigationObserver?.disconnect();
    state.continuousNavigationObserver = undefined;
    state.continuousPages.clear();
    state.continuousPixels = 0;
    ui.continuousView.replaceChildren();
    ui.continuousView.hidden = true;
    ui.surfaceStage.hidden = false;
    ui.surfaceWrap.dataset.continuous = "false";
    void renderCurrentUnit();
  }
}

async function renderCurrentUnit(expectedSession = state.sessionRevision) {
  const officeDocument = state.officeDocument;
  const unit = selectedUnit();
  if (officeDocument === undefined || unit === undefined
    || expectedSession !== state.sessionRevision || state.surfaceContextLost) return;
  updateSpeakerNotes(unit);
  if (state.continuous && officeDocument.info.kind === "text") {
    await buildContinuousPages(expectedSession);
    return;
  }

  const renderRevision = ++state.renderRevision;
  const preserveDisplayedSheet = unit.type === "sheet"
    && state.displayedFrame?.officeDocument === officeDocument
    && state.displayedFrame.unitIndex === unit.index;
  if (!preserveDisplayedSheet) {
    state.displayedFrame = undefined;
    ui.textLayer.replaceChildren();
    resetHit();
    setStatus(`正在渲染${unitLabel(unit)}…`, "loading");
  }
  state.renderDiagnostics = [];
  updateDiagnosticTabs();
  const scale = unit.type === "sheet" ? requestedSheetScale(unit) : requestedScale(unit);
  const viewport = sheetViewport(unit, scale);
  const renderBounds = viewport ?? { x: 0, y: 0, width: unit.width, height: unit.height };
  const pixelRatio = Math.min(Math.max(window.devicePixelRatio || 1, 1), 2);
  const frozenRequests = frozenPaneRequests(unit, renderBounds);
  const estimatedPixels = renderPixelEstimate(renderBounds, scale, pixelRatio)
    + frozenRequests.reduce(
      (total, request) => total + renderPixelEstimate(request.viewport, scale, pixelRatio),
      0,
    );
  if (!Number.isSafeInteger(estimatedPixels) || estimatedPixels > MAX_RENDER_PIXELS) {
    clearCanvas();
    setStatus(
      `当前缩放预计产生 ${estimatedPixels.toLocaleString()} 像素，超过 ${MAX_RENDER_PIXELS.toLocaleString()} 像素限制；请选择“适应窗口”。`,
      "error",
    );
    return;
  }

  let frame;
  let frozen;
  try {
    frame = await officeDocument.render({
      unitIndex: unit.index,
      ...(viewport === undefined ? {} : { viewport }),
      scale,
      pixelRatio,
      background: "#ffffff",
      ...(unit.type === "sheet" ? { sheetSizes: sheetSizeRequest(unit) } : {}),
    }, {
      priority: "interactive",
      supersedeKey: "surface",
    });
    if (expectedSession !== state.sessionRevision
      || renderRevision !== state.renderRevision
      || officeDocument !== state.officeDocument) {
      return;
    }
    frozen = await renderFrozenPaneFrames(
      officeDocument,
      unit,
      frame.viewport,
      scale,
      pixelRatio,
    );
    if (expectedSession !== state.sessionRevision
      || renderRevision !== state.renderRevision
      || officeDocument !== state.officeDocument) {
      return;
    }

    const context = ui.surface.getContext("2d");
    if (context === null) throw new Error("浏览器无法创建 Canvas 2D 上下文");
    if (context.isContextLost?.() === true) {
      handleSurfaceContextLost();
      return;
    }
    ui.surface.width = frame.pixelWidth;
    ui.surface.height = frame.pixelHeight;
    const cssWidth = frame.viewport.width * scale;
    const cssHeight = frame.viewport.height * scale;
    ui.surface.style.width = `${cssWidth}px`;
    ui.surface.style.height = `${cssHeight}px`;
    placeRenderedSurface(unit, frame.viewport, scale, cssWidth, cssHeight);
    context.clearRect(0, 0, frame.pixelWidth, frame.pixelHeight);
    context.drawImage(frame.bitmap, 0, 0);
    if (context.isContextLost?.() === true) {
      handleSurfaceContextLost();
      return;
    }
    drawFrozenPaneFrames(frozen, frame.viewport, scale);
    if (preserveDisplayedSheet) {
      ui.textLayer.replaceChildren();
      resetHit();
    }
    showZoom(scale);

    state.displayedFrame = {
      sessionRevision: expectedSession,
      renderRevision,
      officeDocument,
      unitIndex: unit.index,
      viewport: frame.viewport,
      cssWidth,
      cssHeight,
      pixelWidth: frame.pixelWidth,
      pixelHeight: frame.pixelHeight,
    };
    renderMediaLayer(frame.media, state.displayedFrame);
    await renderTextLayer(state.displayedFrame);
    if (expectedSession !== state.sessionRevision
      || renderRevision !== state.renderRevision
      || officeDocument !== state.officeDocument) {
      return;
    }
    state.renderDiagnostics = [...frame.diagnostics, ...frozen.diagnostics];
    ui.emptyState.hidden = true;
    ui.surfaceWrap.dataset.visible = "true";
    updateCurrentUnitPreview(unit.index);
    updateDiagnosticTabs();
    setStatus(
      `${unitLabel(unit)}渲染完成 · 视口 ${formatRect(frame.viewport)} · ${frame.pixelWidth} × ${frame.pixelHeight} 像素 · 成功绘制 ${frame.renderedObjectCount + frozen.renderedObjectCount} 个对象`,
      state.renderDiagnostics.some((diagnostic) => diagnostic.severity === "error" || diagnostic.severity === "fatal")
        ? "warning"
        : "ready",
    );
    queueSheetViewportRender();
  } catch (cause) {
    if (expectedSession !== state.sessionRevision
      || renderRevision !== state.renderRevision
      || officeDocument !== state.officeDocument) {
      return;
    }
    clearCanvas();
    state.renderDiagnostics = errorDiagnostics(cause);
    selectDiagnosticTab("render");
    setStatus(`渲染失败：${errorMessage(cause)}`, "error");
  } finally {
    frame?.bitmap.close();
    for (const { frame: frozenFrame } of frozen?.frames ?? []) frozenFrame.bitmap.close();
  }
}

function sourceFields(source) {
  const fields = {};
  for (const [key, value] of Object.entries(source)) {
    if (key === "format" || key === "kind" || key === "part" || key === "mapping") continue;
    fields[key] = value;
  }
  return Object.keys(fields).length === 0 ? "-" : JSON.stringify(fields, null, 2);
}

function showObjectDetails(hit) {
  const object = hit.object;
  text(ui.objectId, object.id);
  text(ui.objectType, object.type);
  text(ui.objectBounds, formatRect(object.bounds));
  text(ui.objectText, object.text === undefined || object.text.length === 0 ? "-" : truncate(object.text));
  text(ui.sourceType, `${object.source.format} / ${object.source.kind}`);
  text(ui.sourcePart, object.source.part);
  text(ui.sourceQuality, object.source.mapping);
  text(ui.sourceFields, truncate(sourceFields(object.source)));

  ui.ancestorList.replaceChildren();
  if (hit.ancestors.length === 0) {
    const empty = document.createElement("li");
    empty.className = "ancestor-empty";
    text(empty, "无祖先对象");
    ui.ancestorList.append(empty);
  } else {
    for (const ancestor of hit.ancestors) {
      const item = document.createElement("li");
      text(item, `${ancestor.type} · ${ancestor.id}`);
      ui.ancestorList.append(item);
    }
  }
}

function showHit(hit, displayed) {
  const object = hit.object;
  const bounds = displayedObjectBounds(object, displayed);
  showObjectDetails(hit);
  const pixelBounds = {
    x: (bounds.x - displayed.viewport.x) * displayed.pixelWidth / displayed.viewport.width,
    y: (bounds.y - displayed.viewport.y) * displayed.pixelHeight / displayed.viewport.height,
    width: bounds.width * displayed.pixelWidth / displayed.viewport.width,
    height: bounds.height * displayed.pixelHeight / displayed.viewport.height,
  };
  text(ui.objectPixels, formatRect(pixelBounds));

  const left = (bounds.x - displayed.viewport.x) * displayed.cssWidth / displayed.viewport.width;
  const top = (bounds.y - displayed.viewport.y) * displayed.cssHeight / displayed.viewport.height;
  const width = bounds.width * displayed.cssWidth / displayed.viewport.width;
  const height = bounds.height * displayed.cssHeight / displayed.viewport.height;
  ui.selectionBox.style.left = `${left}px`;
  ui.selectionBox.style.top = `${top}px`;
  ui.selectionBox.style.width = `${width}px`;
  ui.selectionBox.style.height = `${height}px`;
  ui.selectionBox.dataset.visible = "true";
}

async function ancestorsForObject(officeDocument, object) {
  const ancestors = [];
  const visited = new Set([object.id]);
  let parentId = object.parentId;
  while (parentId !== undefined && ancestors.length < 128 && !visited.has(parentId)) {
    visited.add(parentId);
    const parent = await officeDocument.getObject(parentId);
    if (parent === undefined) break;
    ancestors.unshift(parent);
    parentId = parent.parentId;
  }
  return ancestors;
}

async function focusSearchResult(index, expectedRevision = state.searchRevision) {
  const officeDocument = state.officeDocument;
  const count = state.searchResults.length;
  if (officeDocument === undefined || count === 0 || expectedRevision !== state.searchRevision) return;
  const normalized = (index % count + count) % count;
  const result = state.searchResults[normalized];
  state.searchIndex = normalized;
  state.viewportAnchor = { unitIndex: result.object.unitIndex, bounds: result.object.bounds };
  updateSearchControls();

  const position = [...ui.unitSelect.options]
    .findIndex((option) => Number(option.value) === result.object.unitIndex);
  if (position < 0) return;
  ui.unitSelect.selectedIndex = position;
  updateUnitButtons();
  if (state.continuous && officeDocument.info.kind === "text") {
    await renderContinuousPage(result.object.unitIndex);
    if (expectedRevision !== state.searchRevision || officeDocument !== state.officeDocument) return;
    const record = state.continuousPages.get(result.object.unitIndex);
    if (record === undefined) return;
    revealContinuousPage(record, "center");
    for (const item of ui.continuousView.querySelectorAll('[data-search-current="true"]')) {
      item.dataset.searchCurrent = "false";
    }
    const match = [...record.textLayer.querySelectorAll("[data-object-id]")]
      .find((item) => item.dataset.objectId === result.object.id);
    if (match !== undefined) match.dataset.searchCurrent = "true";
    const ancestors = await ancestorsForObject(officeDocument, result.object);
    if (expectedRevision !== state.searchRevision) return;
    showObjectDetails({ object: result.object, ancestors });
    text(ui.objectPixels, "连续页面");
    setStatus(
      `搜索“${ui.searchInput.value}”：第 ${normalized + 1} / ${count} 个对象，${result.ranges.length} 处匹配。`,
      "ready",
    );
    return;
  }
  await renderCurrentUnit();
  if (expectedRevision !== state.searchRevision || officeDocument !== state.officeDocument) return;
  const displayed = state.displayedFrame;
  if (displayed === undefined || displayed.unitIndex !== result.object.unitIndex) return;
  const ancestors = await ancestorsForObject(officeDocument, result.object);
  if (expectedRevision !== state.searchRevision || displayed !== state.displayedFrame) return;
  showHit({ object: result.object, ancestors }, displayed);
  ui.selectionBox.scrollIntoView({ block: "center", inline: "center", behavior: "smooth" });
  setStatus(
    `搜索“${ui.searchInput.value}”：第 ${normalized + 1} / ${count} 个对象，${result.ranges.length} 处匹配。`,
    "ready",
  );
}

async function performSearch() {
  const officeDocument = state.officeDocument;
  const query = ui.searchInput.value;
  const revision = ++state.searchRevision;
  state.searchResults = [];
  state.searchIndex = -1;
  state.viewportAnchor = undefined;
  updateSearchControls();
  if (officeDocument === undefined || query.length === 0) {
    resetHit();
    return;
  }
  try {
    const results = await officeDocument.searchText({ query, limit: 1_000 });
    if (revision !== state.searchRevision || officeDocument !== state.officeDocument) return;
    state.searchResults = results;
    state.searchIndex = results.length === 0 ? -1 : 0;
    updateSearchControls();
    if (results.length === 0) {
      resetHit();
      setStatus(`未找到“${query}”。`, "ready");
      return;
    }
    await focusSearchResult(0, revision);
  } catch (cause) {
    if (revision !== state.searchRevision) return;
    setStatus(`搜索失败：${errorMessage(cause)}`, "error");
  }
}

function queueSearch() {
  if (state.searchTimer !== undefined) clearTimeout(state.searchTimer);
  state.searchResults = [];
  state.searchIndex = -1;
  state.viewportAnchor = undefined;
  updateSearchControls();
  state.searchTimer = setTimeout(() => {
    state.searchTimer = undefined;
    void performSearch();
  }, 160);
}

function stepSearch(offset) {
  if (state.searchResults.length === 0) {
    void performSearch();
    return;
  }
  void focusSearchResult(state.searchIndex + offset);
}

async function hitTest(event) {
  const displayed = state.displayedFrame;
  if (displayed === undefined) {
    setStatus("当前画布尚未完成渲染，暂时无法执行命中测试。", "warning");
    return;
  }
  if (displayed.officeDocument !== state.officeDocument) {
    setStatus("文档已切换，请等待当前画布完成渲染后再试。", "warning");
    return;
  }
  const bounds = ui.surface.getBoundingClientRect();
  if (bounds.width <= 0 || bounds.height <= 0) return;
  let x = displayed.viewport.x
    + (event.clientX - bounds.left) * displayed.viewport.width / bounds.width;
  let y = displayed.viewport.y
    + (event.clientY - bounds.top) * displayed.viewport.height / bounds.height;
  const unit = displayed.officeDocument.info.units[displayed.unitIndex];
  if (unit?.type === "sheet") {
    const layout = sheetLayout(unit);
    x = layout.columns.unmap(x);
    y = layout.rows.unmap(y);
  }
  const hitRevision = ++state.hitRevision;
  try {
    const hits = await displayed.officeDocument.hitTest({
      unitIndex: displayed.unitIndex,
      x,
      y,
      limit: 16,
    });
    if (hitRevision !== state.hitRevision
      || displayed !== state.displayedFrame
      || displayed.sessionRevision !== state.sessionRevision) {
      return;
    }
    if (hits.length === 0) {
      resetHit();
      setStatus(`坐标 ${formatNumber(x)}, ${formatNumber(y)} 未命中文档对象。`, "ready");
      return;
    }
    showHit(hits[0], displayed);
    setStatus(
      `坐标 ${formatNumber(x)}, ${formatNumber(y)} 命中 ${hits.length} 个对象；当前显示最上层对象 ${hits[0].object.id}。`,
      "ready",
    );
  } catch (cause) {
    if (hitRevision !== state.hitRevision || displayed !== state.displayedFrame) return;
    resetHit();
    setStatus(`命中测试失败：${errorMessage(cause)}`, "error");
  }
}

async function openFile(file) {
  const sessionRevision = ++state.sessionRevision;
  closeActiveSession();
  setDocumentOpen(true);
  state.documentDiagnostics = [];
  state.renderDiagnostics = [];
  state.diagnosticTab = "document";
  resetUnitControls();
  updateDiagnosticTabs();
  text(ui.fileName, file.name || "未命名文件");
  ui.fileName.title = file.name || "未命名文件";
  text(ui.fileSize, formatBytes(file.size));
  text(ui.documentFormat, "识别中…");
  text(ui.documentKind, "-");
  setDetectedFormat("");
  setDocumentKind("");
  ui.clearButton.disabled = false;

  const unsupportedLabel = unsupportedDocumentLabel(file);
  if (unsupportedLabel !== undefined) {
    text(ui.documentFormat, "-");
    text(ui.documentKind, "-");
    setDetectedFormat("");
    setDocumentKind("");
    setStatus(`${unsupportedLabel} 文档不受支持。`, "error");
    return;
  }

  if (file.size > MAX_INPUT_BYTES) {
    text(ui.documentFormat, "-");
    setStatus(
      `文件大小为 ${formatBytes(file.size)}，超过检查页的 128 MiB 输入限制。`,
      "error",
    );
    return;
  }
  if (state.engine === undefined) {
    text(ui.documentFormat, "-");
    setStatus("引擎尚未就绪，请稍后重试。", "error");
    return;
  }

  const controller = new AbortController();
  state.openController = controller;
  setStatus(`正在读取并解析 ${file.name || "未命名文件"}…`, "loading");
  let candidate;
  let layoutDocument;
  try {
    const bytes = await file.arrayBuffer();
    if (sessionRevision !== state.sessionRevision || controller.signal.aborted) return;
    let password;
    for (let attempt = 0; attempt < 3; attempt += 1) {
      try {
        candidate = await state.engine.open(bytes, {
          signal: controller.signal,
          timeoutMs: OPEN_TIMEOUT_MS,
          fileName: file.name,
          ...(password === undefined ? {} : { password }),
        });
        break;
      } catch (cause) {
        if (cause?.code !== "PDF_PASSWORD_REQUIRED" && cause?.code !== "PDF_PASSWORD_INCORRECT") throw cause;
        const supplied = window.prompt(
          cause.code === "PDF_PASSWORD_INCORRECT" ? "密码错误，请重试：" : "请输入 PDF 密码：",
        );
        if (supplied === null) throw cause;
        password = supplied;
      }
    }
    if (candidate === undefined) throw new Error("PDF 密码重试次数过多");
    if (sessionRevision !== state.sessionRevision || controller.signal.aborted) {
      candidate.close();
      return;
    }
    state.sourceDocument = candidate;
    state.diagnosticObjects = undefined;
    text(ui.documentFormat, candidate.info.format.toUpperCase());
    text(ui.documentKind, candidate.info.kind);
    setDetectedFormat(candidate.info.format);
    setDocumentKind(candidate.info.kind);
    layoutDocument = candidate;
    if (sessionRevision !== state.sessionRevision || controller.signal.aborted) {
      if (layoutDocument !== candidate) layoutDocument.close();
      candidate.close();
      return;
    }
    state.officeDocument = layoutDocument;
    state.documentDiagnostics = layoutDocument.diagnostics();
    populateUnits(layoutDocument.info.units);
    updateDiagnosticTabs();
    void updateDocumentFonts(layoutDocument, sessionRevision);
    if (layoutDocument.info.units.length === 0) {
      setStatus("文档已识别，但没有可渲染单元。请检查文档诊断。", "warning");
      return;
    }
    await renderCurrentUnit(sessionRevision);
    void renderUnitThumbnails(layoutDocument, sessionRevision);
  } catch (cause) {
    if (layoutDocument !== undefined && layoutDocument !== candidate) layoutDocument.close();
    candidate?.close();
    if (sessionRevision !== state.sessionRevision || controller.signal.aborted) return;
    state.officeDocument = undefined;
    state.sourceDocument = undefined;
    state.documentDiagnostics = errorDiagnostics(cause);
    text(ui.documentFormat, "-");
    text(ui.documentKind, "-");
    setDetectedFormat("");
    setDocumentKind("");
    resetUnitControls();
    updateDiagnosticTabs();
    setStatus(`打开失败：${errorMessage(cause)}`, "error");
  } finally {
    if (state.openController === controller) state.openController = undefined;
  }
}

async function openFixtureFromQuery() {
  const fixture = new URLSearchParams(location.search).get("fixture");
  if (fixture === null) return false;
  if (fixture.startsWith(".") || fixture.length > 160 || /[\\/\0]/u.test(fixture)
    || !/\.(?:pptx|pptm|ppsx|ppsm|potx|potm|odp|otp|fodp|dps|xlsx|xlsm|xltx|xltm|ods|ots|fods|et|docx|docm|dotx|dotm|odt|ott|wps|csv|rtf|pdf)$/iu.test(fixture)) {
    throw new Error("fixture 必须是受支持的本地 Office/WPS/PDF 测试文件名");
  }
  const response = await fetch(`/tests/fixtures/${encodeURIComponent(fixture)}`, { cache: "no-store" });
  if (!response.ok) throw new Error(`fixture 加载失败：HTTP ${response.status}`);
  await openFile(new File([await response.arrayBuffer()], fixture, { type: "application/octet-stream" }));
  return true;
}

function selectAdjacentUnit(offset) {
  const nextIndex = ui.unitSelect.selectedIndex + offset;
  selectUnitPosition(nextIndex);
}

function selectUnitPosition(position, options = {}) {
  if (position < 0 || position >= ui.unitSelect.options.length || state.officeDocument === undefined) return;
  state.viewportAnchor = undefined;
  state.sheetViewportOrigin = undefined;
  ui.unitSelect.selectedIndex = position;
  updateUnitButtons();
  if (options.focus === true) updateUnitNavigation({ focus: true });
  if (state.continuous) {
    const unitIndex = Number(ui.unitSelect.value);
    void renderContinuousPage(unitIndex).then(() => {
      const record = state.continuousPages.get(unitIndex);
      if (record !== undefined) revealContinuousPage(record);
    });
    return;
  }
  ui.viewer.scrollTo({ left: 0, top: 0 });
  void renderCurrentUnit();
}

function handleUnitListKeydown(event) {
  const buttons = unitButtons();
  if (buttons.length === 0) return;
  const current = Math.max(0, buttons.indexOf(document.activeElement));
  let next = current;
  if (event.key === "ArrowDown" || event.key === "ArrowRight") next = Math.min(buttons.length - 1, current + 1);
  else if (event.key === "ArrowUp" || event.key === "ArrowLeft") next = Math.max(0, current - 1);
  else if (event.key === "Home") next = 0;
  else if (event.key === "End") next = buttons.length - 1;
  else return;
  event.preventDefault();
  selectUnitPosition(next, { focus: true });
}

function handleDiagnosticTabKeydown(event) {
  if (!["ArrowLeft", "ArrowRight", "Home", "End"].includes(event.key)) return;
  event.preventDefault();
  const selectDocument = event.key === "ArrowLeft" || event.key === "Home";
  selectDiagnosticTab(selectDocument ? "document" : "render", { focus: true });
}

function setInspectorExpanded(expanded, options = {}) {
  ui.app.dataset.inspectorExpanded = String(expanded);
  ui.inspectorToggle.setAttribute("aria-expanded", String(expanded));
  ui.inspectorToggle.setAttribute("aria-pressed", String(expanded));
  text(ui.inspectorToggle, expanded ? "隐藏检查器" : "显示检查器");
  ui.inspector.setAttribute("aria-hidden", String(!expanded));
  ui.inspector.inert = !expanded;
  if (!expanded && (options.restoreFocus === true || ui.inspector.contains(document.activeElement))) {
    ui.inspectorToggle.focus();
  }
}

function openFilePicker() {
  if (!ui.fileInput.disabled) ui.fileInput.click();
}

ui.fileInput.addEventListener("change", () => {
  const file = ui.fileInput.files?.[0];
  ui.fileInput.value = "";
  if (file !== undefined) void openFile(file);
});

ui.dropZone.addEventListener("click", (event) => {
  if (event.target instanceof Element && event.target.closest("label") !== null) return;
  openFilePicker();
});

ui.dropZone.addEventListener("keydown", (event) => {
  if (event.key !== "Enter" && event.key !== " ") return;
  event.preventDefault();
  openFilePicker();
});

for (const type of ["dragenter", "dragover"]) {
  ui.app.addEventListener(type, (event) => {
    event.preventDefault();
    if (state.engine !== undefined) {
      ui.dropZone.dataset.active = "true";
      ui.app.dataset.dragActive = "true";
    }
    if (event.dataTransfer !== null) event.dataTransfer.dropEffect = "copy";
  });
}

ui.app.addEventListener("dragleave", (event) => {
  if (!(event.relatedTarget instanceof Node) || !ui.app.contains(event.relatedTarget)) {
    ui.dropZone.dataset.active = "false";
    ui.app.dataset.dragActive = "false";
  }
});

ui.app.addEventListener("drop", (event) => {
  event.preventDefault();
  ui.dropZone.dataset.active = "false";
  ui.app.dataset.dragActive = "false";
  if (state.engine === undefined) return;
  const file = event.dataTransfer?.files[0];
  if (file !== undefined) void openFile(file);
});

ui.clearButton.addEventListener("click", clearAll);
ui.previousUnit.addEventListener("click", () => selectAdjacentUnit(-1));
ui.nextUnit.addEventListener("click", () => selectAdjacentUnit(1));
ui.unitSelect.addEventListener("change", () => {
  state.viewportAnchor = undefined;
  state.sheetViewportOrigin = undefined;
  updateUnitButtons();
  if (state.continuous) {
    const unitIndex = Number(ui.unitSelect.value);
    void renderContinuousPage(unitIndex).then(() => {
      const record = state.continuousPages.get(unitIndex);
      if (record !== undefined) revealContinuousPage(record);
    });
    return;
  }
  ui.viewer.scrollTo({ left: 0, top: 0 });
  void renderCurrentUnit();
});
ui.unitList.addEventListener("keydown", handleUnitListKeydown);
ui.searchForm.addEventListener("submit", (event) => {
  event.preventDefault();
  stepSearch(1);
});
ui.searchInput.addEventListener("input", queueSearch);
ui.searchInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    stepSearch(event.shiftKey ? -1 : 1);
  } else if (event.key === "Escape") {
    event.preventDefault();
    resetSearch({ clearQuery: true, disable: false });
    resetHit();
  }
});
ui.searchPrevious.addEventListener("click", () => stepSearch(-1));
ui.searchNext.addEventListener("click", () => stepSearch(1));
ui.fitView.addEventListener("click", fitCurrentUnit);
ui.continuousButton.addEventListener("click", () => setContinuousMode(!state.continuous));
ui.slideshowButton.addEventListener("click", () => void startSlideshow());
ui.slideshowPrevious.addEventListener("click", () => selectAdjacentUnit(-1));
ui.slideshowNext.addEventListener("click", () => selectAdjacentUnit(1));
ui.slideshowExit.addEventListener("click", () => void stopSlideshow());
ui.printButton.addEventListener("click", () => void printDocument());
ui.fullscreenButton.addEventListener("click", () => void toggleFullscreen());
ui.zoomOut.addEventListener("click", () => applyZoomDelta(-25));
ui.zoomIn.addEventListener("click", () => applyZoomDelta(25));
ui.zoomRange.addEventListener("input", () => setZoomPercent(Number(ui.zoomRange.value)));
ui.zoomRange.addEventListener("change", () => void renderCurrentUnit());
ui.surface.addEventListener("contextlost", handleSurfaceContextLost);
ui.surface.addEventListener("contextrestored", handleSurfaceContextRestored);
ui.surface.addEventListener("click", (event) => void hitTest(event));
ui.textLayer.addEventListener("click", (event) => {
  if (event.target instanceof HTMLAnchorElement) return;
  void hitTest(event);
});
for (const headers of [ui.sheetColumnHeaders, ui.sheetRowHeaders]) {
  headers.addEventListener("pointerdown", (event) => {
    const target = event.target instanceof Element ? event.target.closest(".sheet-header-resizer") : null;
    if (target instanceof HTMLElement) beginSheetResize(event, target);
  });
  headers.addEventListener("dblclick", (event) => {
    const target = event.target instanceof Element ? event.target.closest(".sheet-header-resizer") : null;
    if (target instanceof HTMLElement) void autoFitSheetSize(event, target);
  });
  headers.addEventListener("keydown", (event) => {
    const target = event.target instanceof Element ? event.target.closest(".sheet-header-resizer") : null;
    const unit = selectedUnit();
    if (!(target instanceof HTMLElement) || unit?.type !== "sheet") return;
    const axis = target.dataset.axis;
    let direction = 0;
    if (axis === "column") {
      if (event.key === "ArrowLeft") direction = -1;
      else if (event.key === "ArrowRight") direction = 1;
    } else if (event.key === "ArrowUp") {
      direction = -1;
    } else if (event.key === "ArrowDown") {
      direction = 1;
    }
    if (direction === 0) return;
    event.preventDefault();
    const index = Number(target.dataset.index);
    const layout = sheetLayout(unit);
    const current = (axis === "column" ? layout.columns : layout.rows).size(index);
    const minimum = axis === "column" ? MIN_COLUMN_WIDTH : MIN_ROW_HEIGHT;
    commitSheetSize(axis, index, Math.max(minimum, current + direction * (event.shiftKey ? 10 : 2)));
  });
}
window.addEventListener("pointermove", updateSheetResize);
window.addEventListener("pointerup", finishSheetResize);
window.addEventListener("pointercancel", finishSheetResize);
ui.viewer.addEventListener("scroll", queueSheetViewportRender, { passive: true });
ui.documentTab.addEventListener("click", () => selectDiagnosticTab("document"));
ui.renderTab.addEventListener("click", () => selectDiagnosticTab("render"));
ui.documentTab.addEventListener("keydown", handleDiagnosticTabKeydown);
ui.renderTab.addEventListener("keydown", handleDiagnosticTabKeydown);
ui.inspectorToggle.addEventListener("click", () => {
  setInspectorExpanded(ui.app.dataset.inspectorExpanded !== "true", { restoreFocus: true });
});
ui.inspectorClose.addEventListener("click", () => setInspectorExpanded(false, { restoreFocus: true }));
document.addEventListener("pointerdown", (event) => {
  state.fontGestureItem = event.target instanceof Element
    ? event.target.closest(".text-layer-item") ?? undefined
    : undefined;
  if (state.fontGestureItem === undefined) resetFontSelection();
});
document.addEventListener("pointerup", (event) => {
  const item = event.target instanceof Element
    ? event.target.closest(".text-layer-item") ?? undefined
    : undefined;
  if (item !== undefined) state.fontGestureItem = item;
  queueSelectedFontUpdate();
});
document.addEventListener("selectionchange", queueSelectedFontUpdate);
document.addEventListener("fullscreenchange", handleFullscreenChange);

const narrowInspectorQuery = window.matchMedia("(max-width: 1080px)");
if (narrowInspectorQuery.matches) setInspectorExpanded(false);
narrowInspectorQuery.addEventListener("change", (event) => {
  if (event.matches) setInspectorExpanded(false);
});

document.addEventListener("keydown", (event) => {
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "o") {
    event.preventDefault();
    openFilePicker();
    return;
  }
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "f") {
    event.preventDefault();
    if (!ui.searchInput.disabled) {
      ui.searchInput.focus();
      ui.searchInput.select();
    }
    return;
  }
  const target = event.target;
  if (target instanceof HTMLInputElement
    || target instanceof HTMLSelectElement
    || target instanceof HTMLTextAreaElement
    || (target instanceof HTMLElement && target.isContentEditable)) {
    return;
  }
  if (state.slideshow) {
    if (event.key === "ArrowLeft" || event.key === "ArrowUp" || event.key === "PageUp") {
      event.preventDefault();
      selectAdjacentUnit(-1);
    } else if (event.key === "ArrowRight" || event.key === "ArrowDown" || event.key === "PageDown"
      || (event.key === " " && !(target instanceof HTMLButtonElement))) {
      event.preventDefault();
      selectAdjacentUnit(1);
    } else if (event.key === "Home") {
      event.preventDefault();
      selectUnitPosition(0);
    } else if (event.key === "End") {
      event.preventDefault();
      selectUnitPosition(ui.unitSelect.options.length - 1);
    }
    return;
  }
  if (event.key === "PageUp") {
    event.preventDefault();
    selectAdjacentUnit(-1);
  } else if (event.key === "PageDown") {
    event.preventDefault();
    selectAdjacentUnit(1);
  } else if (event.key === "Escape" && ui.app.dataset.inspectorExpanded === "true") {
    setInspectorExpanded(false, { restoreFocus: true });
  } else if (event.key === "+" || event.key === "=") {
    event.preventDefault();
    applyZoomDelta(25);
  } else if (event.key === "-") {
    event.preventDefault();
    applyZoomDelta(-25);
  }
});

let resizeFrame;
const resizeObserver = new ResizeObserver(() => {
  if (state.zoomMode !== "fit" || state.officeDocument === undefined) return;
  const unit = selectedUnit();
  const displayed = state.displayedFrame;
  if (unit !== undefined && displayed !== undefined && displayed.unitIndex === unit.index) {
    const nextScale = unit.type === "sheet" ? requestedSheetScale(unit) : requestedScale(unit);
    const displayedScale = displayed.cssWidth / displayed.viewport.width;
    if (Math.abs(nextScale - displayedScale) < 0.001) return;
  }
  if (resizeFrame !== undefined) cancelAnimationFrame(resizeFrame);
  resizeFrame = requestAnimationFrame(() => {
    resizeFrame = undefined;
    void renderCurrentUnit();
  });
});
resizeObserver.observe(ui.viewer);

window.addEventListener("pagehide", (event) => {
  if (event.persisted || state.disposed) return;
  state.disposed = true;
  state.sessionRevision += 1;
  closeActiveSession();
  resizeObserver.disconnect();
  if (resizeFrame !== undefined) cancelAnimationFrame(resizeFrame);
  state.engine?.close();
  state.engine = undefined;
});

async function initialize() {
  renderDocumentFonts();
  try {
    const engine = await createOfficeEngine({
      formatPack: () => import("../dist/extended-formats.js")
        .then(({ extendedFormatPack }) => extendedFormatPack),
      limits: {
        inputBytes: MAX_INPUT_BYTES,
        renderPixels: MAX_RENDER_PIXELS,
      },
    });
    if (state.disposed) {
      engine.close();
      return;
    }
    state.engine = engine;
    ui.fileInput.disabled = false;
    ui.dropZone.dataset.disabled = "false";
    setStatus("引擎已就绪，请选择一个 Office 或 PDF 文件。", "ready", "idle");
    try {
      await openFixtureFromQuery();
    } catch (cause) {
      setStatus(`自动加载测试文件失败：${errorMessage(cause)}`, "error");
    }
  } catch (cause) {
    setStatus(`引擎初始化失败：${errorMessage(cause)}`, "error");
  }
}

void initialize();
