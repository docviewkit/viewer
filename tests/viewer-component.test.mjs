import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const source = await readFile(new URL("../src/viewer.ts", import.meta.url), "utf8");
const html = await readFile(new URL("../examples/viewer.html", import.meta.url), "utf8");
const script = await readFile(new URL("../examples/viewer.js", import.meta.url), "utf8");
const browserMatrix = await readFile(
  new URL("../scripts/run-browser-matrix.mjs", import.meta.url),
  "utf8",
);
const packageJson = JSON.parse(await readFile(new URL("../package.json", import.meta.url), "utf8"));

test("package exposes the framework-independent Viewer entry", () => {
  assert.deepEqual(packageJson.exports["./viewer"], {
    types: "./dist/viewer.d.ts",
    import: "./dist/viewer.js",
  });
  assert.deepEqual(packageJson.sideEffects, ["./dist/viewer.js"]);
  assert.match(source, /class DocViewKitViewerElement extends HTMLElement/);
  assert.match(source, /customElements\.define\(tagName, DocViewKitViewerElement\)/);
  assert.match(html, /<docviewkit-viewer\b/);
});

test("public surface stays small and source location is unified", () => {
  assert.match(source, /async open\(\s*source:/);
  assert.match(source, /options: ViewerOpenOptions/);
  assert.match(source, /async reveal\(target: ViewerTarget\)/);
  assert.match(source, /async close\(\): Promise<void>/);
  assert.match(source, /destroy\(\): void/);
  assert.match(source, /readonly kind: "unit"/);
  assert.match(source, /readonly kind: "object"/);
  assert.match(source, /readonly kind: "region"/);
  assert.match(source, /readonly kind: "source"/);
  assert.doesNotMatch(source, /goToPage|goToSlide|goToCell/);
});

test("Viewer owns rendering while the host owns document acquisition", () => {
  assert.match(source, /createOfficeEngine/);
  assert.match(source, /document\.render/);
  assert.match(source, /document\.searchText/);
  assert.match(source, /document\.getObject/);
  assert.match(script, /fetch\(`\/tests\/fixtures/);
  assert.match(script, /viewer\.open\(file\)/);
  assert.match(source, /PDF_PASSWORD_REQUIRED/);
  assert.match(script, /fileInput\.addEventListener\("change"/);
  assert.match(script, /viewer\.addEventListener\("drop"/);
  assert.match(html, /slot="toolbar-start"/);
  assert.doesNotMatch(source, /<input[^>]+type=["']file/i);
  assert.doesNotMatch(source, /fetch\(/);
  assert.match(source, /readonly watermark\?: string/);
  assert.match(source, /watermark !== undefined \? \{ watermark \} : \{\}/);
  assert.equal((source.match(/\.\.\.this\.#watermarkRequest\(\)/g) ?? []).length, 6);
});

test("Viewer search invalidates stale results before its debounced request", () => {
  assert.match(source, /\.search-input"\)\.addEventListener\("input", \(event\) => \{[\s\S]*?#searchQuery = \(event\.currentTarget as HTMLInputElement\)\.value\.trim\(\);[\s\S]*?#searchResults = \[\];[\s\S]*?#searchIndex = -1;/);
  assert.match(source, /for \(const action of \["previous-result", "next-result"\]\)[\s\S]*?button\.disabled = this\.#searchResults\.length === 0/);
  assert.match(browserMatrix, /did not complete document search/);
  assert.match(browserMatrix, /did not hide document search/);
});

test("Viewer transfers its Blob buffer and refreshes it only for a password retry", () => {
  assert.match(source, /source instanceof Blob \? \{ transferInput: true \} : \{\}/);
  assert.match(source, /source instanceof Blob && bytes instanceof ArrayBuffer && bytes\.byteLength === 0[\s\S]*bytes = await source\.arrayBuffer\(\)/);
});

test("reference Viewer supports direct file drop unless the host disables it", () => {
  assert.match(script, /const fileDropEnabled = parameters\.get\("fileDrop"\) !== "false"/);
  assert.match(script, /for \(const eventName of \["dragenter", "dragover", "dragleave", "drop"\]\) \{[\s\S]*?event\.preventDefault\(\)/);
  assert.match(script, /if \(fileDropEnabled\) \{[\s\S]*?viewer\.addEventListener\("drop"/);
});

test("reference Viewer shows its file picker only when explicitly enabled", () => {
  assert.match(html, /class="open-button"[^>]* hidden/);
  assert.match(script, /const filePickerEnabled = parameters\.get\("filePicker"\) === "true"/);
  assert.match(script, /openButton\.hidden = !filePickerEnabled/);
});

test("Viewer requests PDF passwords through a masked, cleared dialog", () => {
  assert.match(source, /class="password-input" type="password"/);
  assert.match(source, /autocomplete="current-password"/);
  assert.match(source, /dialog\.showModal\(\)/);
  assert.match(source, /input\.value = ""/);
  assert.doesNotMatch(script, /window\.prompt\(/);
});

test("localization is icon-first without losing accessible names", () => {
  assert.match(source, /interface ViewerMessages/);
  assert.match(source, /messages\?: Readonly<Record<string, Partial<ViewerMessages>>>/);
  assert.match(source, /new Intl\.Locale/);
  assert.match(source, /button\.ariaLabel = label/);
  assert.match(source, /button\.title = label/);
  assert.match(source, /shell\.dir = isRtl/);
  assert.match(source, /--dv-icon-size: 20px/);
  assert.match(source, /--dv-icon-stroke: 1\.5/);
  assert.doesNotMatch(source, /[\u{1F300}-\u{1FAFF}]/u);
});

test("theme defaults to browser preference and keeps the document surface white", () => {
  assert.match(source, /type ViewerTheme = "auto" \| "light" \| "dark"/);
  assert.match(source, /matchMedia\("\(prefers-color-scheme: dark\)"\)/);
  assert.match(source, /data-resolved-theme="dark"/);
  assert.match(source, /color-scheme: dark/);
  assert.match(source, /\.stage \{[\s\S]*?background: #ffffff/);
  assert.match(source, /background: "#ffffff"/);
});

test("bounded customization includes feature flags, tokens, parts, and responsive layout", () => {
  assert.match(source, /interface ViewerFeatures/);
  assert.match(source, /features\?: Partial<ViewerFeatures>/);
  assert.match(source, /--dv-accent/);
  assert.match(source, /part="toolbar"/);
  assert.match(source, /part="navigation"/);
  assert.match(source, /part="document"/);
  assert.match(source, /slot name="toolbar-start"/);
  assert.match(source, /slot name="toolbar-end"/);
  assert.match(source, /initialPageMode\?: ViewerPageMode/);
  assert.match(source, /readonly pageMode: ViewerPageMode/);
  assert.match(source, /pageMode: true/);
  assert.match(source, /interactionModeSwitcher: false/);
  assert.match(source, /diagnostics: false/);
  assert.match(source, /diagnostics: \["diagnostics"\]/);
  assert.match(source, /interactionMode: "object"/);
  assert.match(source, /hyperlinks: true/);
  assert.match(source, /class="text-layer"/);
  assert.match(source, /document\.listObjects/);
  assert.match(source, /document\.hitTest/);
  assert.match(source, /docviewkit-objectselect/);
  assert.match(source, /part="continuous"/);
  assert.match(source, /container: docviewkit \/ inline-size/);
  assert.match(source, /@container docviewkit \(max-width: 720px\)/);
  assert.match(source, /prefers-reduced-motion: reduce/);
  assert.match(source, /#previewSingleLayout/);
  assert.match(source, /requestAnimationFrame/);
  assert.doesNotMatch(source, /singleLayoutTimer|continuousLayoutTimer/);
  assert.match(source, /compactLayout !== this\.#compactLayout/);
  assert.match(source, /44px|40px/);
});

test("Viewer chrome keeps stable heights while retaining compact mobile touch targets", () => {
  assert.match(source, /--dv-toolbar-height: 44px/);
  assert.match(source, /--dv-status-height: 28px/);
  assert.match(source, /\.icon-button \{[\s\S]*?flex: 0 0 32px;[\s\S]*?width: 32px;[\s\S]*?height: 32px/);
  assert.match(source, /\.toolbar \.icon-button \{ flex-basis: 40px; width: 40px; height: 40px; \}/);
  assert.doesNotMatch(source, /max-width: 720px[\s\S]{0,180}--dv-(?:toolbar|status)-height/);
  assert.doesNotMatch(source, /max-width: 480px[\s\S]{0,180}--dv-status-height/);
});

test("zoom percentage accepts bounded manual input", () => {
  assert.match(source, /const MIN_ZOOM = \.25;[\s\S]*const MAX_ZOOM = 4;/);
  assert.match(source, /class="zoom-input" type="number" min="\$\{MIN_ZOOM \* 100\}"[\s\S]*max="\$\{MAX_ZOOM \* 100\}" step="1"/);
  assert.match(source, /zoomInput\.addEventListener\("change", \(\) => this\.#commitZoomInput\(\)\)/);
  assert.match(source, /event\.key === "Enter"[\s\S]*this\.#commitZoomInput\(\)/);
  assert.match(source, /event\.key === "Escape"[\s\S]*this\.#syncZoomInput\(\)/);
  assert.match(source, /#commitZoomInput\(\)[\s\S]*Number\.isFinite\(input\.valueAsNumber\)[\s\S]*this\.#setZoom\(input\.valueAsNumber \/ 100, false\)/);
  assert.match(source, /this\.#root\.activeElement !== zoomInput\) this\.#syncZoomInput\(\)/);
  assert.match(source, /Math\.min\(MAX_ZOOM, Math\.max\(MIN_ZOOM/);
});

test("Viewer exposes exactly three mutually exclusive content interaction modes", () => {
  assert.match(source, /type ViewerInteractionMode = "object" \| "display" \| "text"/);
  assert.match(source, /readonly interactionMode: ViewerInteractionMode/);
  assert.match(source, /this\.#features\.interactionMode !== "text"/);
  assert.match(source, /this\.#features\.interactionMode !== "object"/);
  assert.match(source, /\.text-layer-item \{[\s\S]*?user-select: text/);
  assert.match(source, /\.text-layer-item::selection/);
  assert.match(source, /@deprecated Use interactionMode: "text"/);
  assert.match(source, /legacyTextSelection === true[\s\S]*?\? "text"[\s\S]*?legacyObjectSelection === false[\s\S]*?\? "display"/);
});

test("Viewer optionally exposes its three interaction modes as one segmented control", () => {
  assert.match(source, /readonly interactionModeSwitcher: boolean/);
  assert.equal((source.match(/<button class="interaction-option"[^>]*data-action="interaction-mode"/g) ?? []).length, 3);
  for (const mode of ["object", "display", "text"]) {
    assert.match(source, new RegExp(`data-interaction-mode="${mode}"`));
  }
  assert.match(source, /\.interaction-switcher"\)\.hidden = !this\.#features\.interactionModeSwitcher/);
  assert.match(source, /#setInteractionMode\(value: string \| undefined\)[\s\S]*features: \{ \.\.\.this\.#config\.features, interactionMode: value \}/);
  assert.match(source, /button\.setAttribute\("aria-pressed", String\(button\.dataset\.interactionMode === this\.#features\.interactionMode\)\)/);
  assert.match(source, /button\.querySelector<HTMLElement>\("span"\)!\.textContent = label/);
  assert.match(source, /\.interaction-option \.icon \{ display: block; width: 18px; height: 18px; \}/);
  assert.match(source, /\.interaction-option span \{ display: none; \}/);
  assert.match(source, /min-width: 601px[\s\S]*\.interaction-switcher[\s\S]*left: 50%[\s\S]*translateX\(-50%\)/);
  assert.match(source, /max-width: 540px[\s\S]*\.toolbar \.interaction-switcher \{ display: none; \}/);
  assert.match(source, /max-width: 480px[\s\S]*\.toolbar \[data-action="fullscreen"\]/);
  assert.match(source, /\.divider\.secondary:not\(:has\(~ \[data-action="print"\]:not\(\[hidden\]\), ~ \[data-action="fullscreen"\]:not\(\[hidden\]\)\)\)/);
  assert.match(script, /interactionModeSwitcher: parameters\.get\("interactionModeSwitcher"\) !== "false"/);
  assert.match(script, /print: parameters\.get\("print"\) !== "false"/);
  assert.match(script, /fullscreen: parameters\.get\("fullscreen"\) !== "false"/);
});

test("safe document hyperlinks and internal page links are independently configurable", () => {
  assert.match(source, /readonly hyperlinks: boolean/);
  assert.match(source, /this\.#features\.hyperlinks/);
  assert.match(source, /candidate\.trigger === "click"/);
  assert.match(source, /action\.kind === "command" && action\.action === "navigate"/);
  assert.match(source, /action\.kind === "command" && action\.action === "reveal"/);
  assert.match(source, /this\.reveal\(\{ kind: "object", objectId: action\.target \}\)/);
  assert.match(source, /\["http:", "https:", "mailto:", "tel:"\]/);
  assert.match(source, /window\.open\(target, "_blank", "noopener,noreferrer"\)/);
});

test("Apache-2.0 Viewer exposes customization without runtime licensing", () => {
  assert.match(source, /license: OPEN_SOURCE_LICENSE/);
  assert.doesNotMatch(source, /#refreshLicense|#applyLicenseState|license-brand|license-overlay|data-license-part/);
  assert.match(source, /readonly license: LicenseState/);
});

test("open source minimal mode leaves only the rendering workspace visible", () => {
  assert.match(source, /readonly minimal\?: boolean/);
  assert.match(source, /this\.#config\.minimal === true/);
  assert.match(source, /\.shell"\)\.dataset\.minimal = String\(enabled\)/);
  assert.match(source, /\.shell\[data-minimal="true"\] \{[^}]*display: block;[^}]*border: 0;/s);
  assert.match(source, /\.shell\[data-minimal="true"\] :is\(\.toolbar, \.formula-bar, \.navigation, \.navigation-resizer, \.sheet-tabs, \.statusbar, \.diagnostic-panel\) \{ display: none; \}/);
  assert.match(source, /\.shell\[data-minimal="true"\] \.workspace \{ width: 100%; height: 100%; \}/);
  assert.match(browserMatrix, /did not isolate the rendering workspace in minimal mode/);
});

test("single and continuous page modes share lazy rendering and navigation state", () => {
  assert.match(source, /data-action="page-mode"/);
  assert.match(source, /const DEFAULT_PAGE_MODE: ViewerPageMode = "continuous"/);
  assert.match(source, /#pageMode: ViewerPageMode = DEFAULT_PAGE_MODE/);
  assert.match(source, /this\.#config\.initialPageMode \?\? DEFAULT_PAGE_MODE/);
  assert.match(source, /#buildContinuousPages/);
  assert.match(source, /rootMargin: "400px 0px"/);
  assert.match(source, /rootMargin: "-45% 0px -45% 0px"/);
  assert.match(source, /CONTINUOUS_CACHE_PIXELS = 40_000_000/);
  assert.match(source, /CONTINUOUS_SMOOTH_SCROLL_VIEWPORTS = 2/);
  assert.match(source, /this\.#syncContinuousUnit/);
  assert.match(source, /unit\.type !== "sheet"/);
  assert.match(source, /record\.canvas\.width = 1/);
  assert.match(source, /record\.visible \|\| Math\.abs\(unitIndex - this\.#unitIndex\) <= 1/);
  assert.match(source, /record\.renderedScale - record\.scale/);
});

test("continuous rendering does not add page numbers to document content", () => {
  assert.doesNotMatch(source, /continuous-page-number/);
});

test("zoomed continuous pages remain horizontally reachable", () => {
  assert.match(source, /\.continuous-view \{[^}]*width: max-content;[^}]*min-width: 100%;/s);
});

test("object selections are reprojected from document bounds after zoom", () => {
  assert.match(source, /#selectionBounds: Rect \| undefined/);
  assert.match(source, /#selectionUnitIndex = -1/);
  assert.match(source, /#showSelection\(bounds: Rect, reveal = false, objectId\?: string, text\?: string\): void \{[\s\S]*?this\.#selectionBounds = copyRect\(bounds\)/);
  assert.match(source, /if \(reveal && selection\?\.hidden === false\)[\s\S]*scrollIntoView/);
  assert.match(source, /#showSelection\(result\.object\.bounds, true, result\.object\.id, result\.object\.text\)/);
  assert.match(source, /if \(region !== undefined\) this\.#showSelection\(region, true, object\?\.id, object\?\.text\)/);
  assert.match(source, /this\.#showSelection\(hit\.object\.bounds, false, hit\.object\.id, hit\.object\.text\)/);
  assert.match(source, /#queueContinuousLayout\(preserveAnchor = false\): void \{[\s\S]*?this\.#refreshObjectHighlights\(\)/);
  assert.match(source, /this\.#rendered = \{ frame, unit, scale \};[\s\S]*?this\.#refreshObjectHighlights\(\)/);
  assert.match(source, /#refreshObjectHighlight\(kind: "selection" \| "hover"[\s\S]*?bounds\.width \* record\.scale/);
});

test("object interaction previews hover targets and supports direct deselection", () => {
  assert.match(source, /class="object-hover" hidden/);
  assert.match(source, /page\.append\(canvas, textLayer, hover, selection\)/);
  assert.match(source, /workspace\.addEventListener\("pointermove", \(event\) => this\.#queueObjectHover\(event\)/);
  assert.match(source, /workspace\.addEventListener\("pointerleave", \(\) => this\.#clearObjectHover\(\)\)/);
  assert.match(source, /event\.pointerType === "touch"[\s\S]*this\.#clearObjectHover\(\)/);
  assert.match(source, /this\.#hoverTimer = setTimeout\([\s\S]*40\);/);
  assert.match(source, /#updateObjectHover\([\s\S]*document\.hitTest\(\{ \.\.\.point, limit: 16 \}\)/);
  assert.match(source, /if \(point === undefined\) \{[\s\S]*this\.#hideSelection\(\)/);
  assert.match(source, /hit\.object\.id === this\.#selectionObjectId[\s\S]*this\.#hideSelection\(\)[\s\S]*this\.#showObjectHover\(hit/);
  assert.match(source, /event\.key === "Escape" && this\.#selectionBounds !== undefined[\s\S]*this\.#hideSelection\(\)/);
});

test("selected text objects support copy and read-only cut", () => {
  assert.match(source, /class="object-clipboard" tabindex="-1" aria-hidden="true"/);
  assert.match(source, /for \(const type of \["copy", "cut"\] as const\)[\s\S]*#copySelectedObject/);
  assert.match(source, /#copySelectedObject\(event: ClipboardEvent\): void \{[\s\S]*interactionMode !== "object"[\s\S]*setData\("text\/plain", this\.#selectionText\)[\s\S]*event\.preventDefault\(\)/);
  assert.match(source, /clipboard\.value = hit\.object\.text \?\? ""[\s\S]*clipboard\.focus\(\{ preventScroll: true \}\)[\s\S]*clipboard\.select\(\)/);
  assert.match(source, /#hideSelection\(\): void \{[\s\S]*this\.#selectionText = undefined[\s\S]*clipboard\.value = ""/);
});

test("single-page rendering warms adjacent PDF pages at prefetch priority", () => {
  assert.match(source, /#prefetchAdjacentPages\(/);
  assert.match(source, /priority: "prefetch"/);
  assert.match(source, /viewer-page-prefetch:/);
  assert.match(source, /frame\.bitmap\.close\(\)/);
  assert.match(source, /setTimeout\(\(\) => \{[\s\S]*?#prefetchAdjacentPages\(document, unit\.index\);[\s\S]*?\}, 80\)/);
  assert.match(source, /clearTimeout\(this\.#pagePrefetchTimer\)/);
});

test("spreadsheets use Excel-like formula bar, headers, resizing, and bottom sheet tabs", () => {
  assert.match(source, /data-spreadsheet/);
  assert.match(source, /class="formula-bar"[^>]*role="group"/);
  assert.match(source, /class="cell-address">A1</);
  assert.match(source, /source\.formula \?\? object\?\.text/);
  assert.match(source, /class="sheet-column-headers sheet-header-axis"/);
  assert.match(source, /class="sheet-row-headers sheet-header-axis"/);
  assert.match(source, /class="sheet-tabs"[^>]*role="tablist"/);
  assert.match(source, /\.shell\[data-spreadsheet="true"\] \.navigation[^}]*display: none/s);
  assert.match(source, /\.shell\[data-spreadsheet="true"\] \.stage-wrap\s*\{[^}]*justify-content:\s*start[^}]*align-content:\s*start/s);
  assert.match(source, /#sheetHeaderItem\("column"/);
  assert.match(source, /#sheetHeaderItem\("row"/);
  assert.match(source, /const size = layout\.columns\.size\(index\)[\s\S]*?if \(size <= 0\) continue[\s\S]*?#sheetHeaderItem\("column"/);
  assert.match(source, /const size = layout\.rows\.size\(index\)[\s\S]*?if \(size <= 0\) continue[\s\S]*?#sheetHeaderItem\("row"/);
  assert.match(source, /role", "separator"/);
  assert.match(source, /setPointerCapture\(event\.pointerId\)/);
  assert.match(source, /addEventListener\("selectstart"[\s\S]*?event\.preventDefault\(\)/);
  assert.match(source, /#updateSheetResize\(event: PointerEvent\)[\s\S]*?event\.preventDefault\(\)/);
  assert.match(source, /\.sheet-header-axis, \.sheet-header-corner[\s\S]*?-webkit-user-select: none/);
  assert.match(source, /\.shell\[data-sheet-resizing="true"\] \.workspace,[\s\S]*?-webkit-user-select: none/);
  assert.match(source, /#beginSheetResize\(event: PointerEvent\): void \{[\s\S]*?#clearNativeSelection\(\)[\s\S]*?dataset\.sheetResizing = "true"/);
  assert.match(source, /#updateSheetResize\(event: PointerEvent\): void \{[\s\S]*?#clearNativeSelection\(\)/);
  assert.match(source, /if \(resize\.axis === "column"\) \{[\s\S]*?guide\.style\.width = "1px"[\s\S]*?\} else \{[\s\S]*?guide\.style\.height = "1px"/);
  assert.match(source, /#finishSheetResize\(event\?: PointerEvent\): void \{[\s\S]*?#clearNativeSelection\(\)[\s\S]*?delete this\.#element<HTMLElement>\("\.shell"\)\.dataset\.sheetResizing/);
  assert.match(source, /#clearNativeSelection\(\): void \{[\s\S]*?document\.getSelection\(\)[\s\S]*?#root as ShadowRoot[\s\S]*?getSelection/);
  assert.match(source, /sheetSizes: this\.#sheetSizeRequest\(unit\)/);
  assert.match(source, /addEventListener\("dblclick",[\s\S]*?#autoFitSheetSize\(event\)/);
  assert.match(source, /#autoFitSheetSize\(event: Event\): Promise<void>[\s\S]*?sheetAutoFitColumnWidth\(context, cells, MIN_COLUMN_WIDTH\)/);
  assert.match(source, /#autoFitSheetSize\(event: Event\): Promise<void>[\s\S]*?sheetAutoFitRowHeight\(context, cells, sourceHeight\)/);
  assert.match(source, /\.sheet-tab\[aria-selected="true"\]/);
  assert.match(source, /initialZoom === undefined && document\.info\.kind === "spreadsheet"/);
  assert.match(source, /this\.#document === undefined \|\| \(!this\.#fit && this\.#rendered\?\.unit\.type !== "sheet"\)/);
});

test("spreadsheet scrolling keeps an overscanned frame until the replacement is ready", () => {
  assert.match(source, /SHEET_VIEWPORT_OVERSCAN_RATIO\s*=\s*0\.5/);
  assert.match(source, /#sheetVisibleViewport\(/);
  assert.match(source, /#sheetRenderViewport\(/);
  const queueStart = source.indexOf("  #queueSheetScroll(): void {");
  const queue = source.slice(
    queueStart,
    source.indexOf("#setNavigationWidth", queueStart),
  );
  assert.match(queue, /const visible = this\.#sheetVisibleViewport\(unit, rendered\.scale\)/);
  assert.match(source, /const marginX = Math\.max\(0, \(frame\.width - visible\.width\) \/ 4\)/);
  assert.match(queue, /if \(this\.#sheetFrameCoversViewport\(unit, frame, visible\)\) return/);
});

test("navigation thumbnails preserve the rendered viewport aspect ratio", () => {
  const thumbnailRenderer = source.slice(
    source.indexOf("async #renderThumbnail"),
    source.indexOf("#toggleNavigation", source.indexOf("async #renderThumbnail")),
  );
  assert.match(source, /--dv-thumbnail-aspect/);
  assert.match(source, /--dv-thumbnail-max-width/);
  assert.match(thumbnailRenderer, /viewport\.width \/ Math\.max\(1, viewport\.height\)/);
  assert.match(thumbnailRenderer, /120 \/ Math\.max\(1, unit\.width\), 96 \/ Math\.max\(1, unit\.height\)/);
  assert.match(source, /\.thumbnail canvas \{[^}]*width: 100%;[^}]*height: 100%;[^}]*min-width: 0;[^}]*min-height: 0;/);
  assert.match(source, /box-shadow: inset 0 0 0 1px #cfd4dc/);
  assert.doesNotMatch(source, /\.thumbnail \{[^}]*min-height:/);
  assert.doesNotMatch(thumbnailRenderer, /canvas\.style\.width =/);
  assert.doesNotMatch(thumbnailRenderer, /canvas\.style\.height =/);
});

test("navigation prefers a real document outline and can switch to thumbnails", () => {
  assert.match(source, /this\.#navigationMode = \(this\.#info\?\.outline\.length \?\? 0\) > 0 \? "outline" : "thumbnails"/);
  assert.match(source, /data-action="navigation-outline"/);
  assert.match(source, /data-action="navigation-thumbnails"/);
  assert.match(source, /\.navigation-switcher \{[^}]*display: flex;[^}]*justify-content: flex-start;/s);
  assert.match(source, /\.navigation-mode \{[^}]*width: 32px;[^}]*height: 28px;[^}]*border-inline-start-width: 0;/s);
  assert.match(source, /\.navigation-mode:first-child \{[^}]*border-inline-start-width: 1px;/s);
  assert.match(source, /\.navigation-mode:focus-visible \{[^}]*outline: 2px solid var\(--dv-focus\);/s);
  assert.match(source, /@media \(pointer: coarse\) \{[\s\S]*?\.navigation-mode \{[\s\S]*?width: 40px;[\s\S]*?height: 40px;/);
  assert.doesNotMatch(source, /class="navigation-mode"[^>]*>[\s\S]*?<span>/);
  assert.match(source, /#renderOutline\(\)/);
  assert.match(source, /button\.style\.setProperty\("--dv-outline-level", String\(item\.level\)\)/);
  assert.match(source, /this\.#outlineSelection = index;[\s\S]*this\.#selectUnit\(item\.unitIndex\)/);
  assert.match(source, /switcher\.hidden = !hasOutline/);
  assert.match(source, /units\.length \?\? 0\) > 1 \|\| \(this\.#info\?\.outline\.length \?\? 0\) > 0/);
});

test("navigation keeps its layout animation without laying out every page per frame", () => {
  assert.match(source, /grid-template-columns var\(--dv-motion-layout\) var\(--dv-ease-out\)/);
  assert.match(source, /opacity var\(--dv-motion-layout\) var\(--dv-ease-out\)/);
  assert.match(source, /transform var\(--dv-motion-layout\) var\(--dv-ease-out\)/);
  assert.match(source, /\.shell\[data-navigation="false"\] \{ grid-template-columns: 0 0 minmax\(0, 1fr\); \}/);
  assert.match(source, /\.shell\[data-navigation="false"\] \.navigation \{[^}]*transform: translateX\(-12px\);/s);
  assert.doesNotMatch(source, /\.shell\[data-navigation="false"\] \.navigation[^}]*display: none/s);
  assert.match(source, /if \(this\.#navigationTransitioning\) \{[\s\S]*this\.#previewContinuousNavigationLayout\(\)[\s\S]*this\.#previewSingleLayout\(\)[\s\S]*return;/);
  assert.match(source, /#previewContinuousNavigationLayout\(\): void \{[\s\S]*\.continuous-view"\)\.style\.width = `\$\{workspace\.clientWidth\}px`[\s\S]*this\.#unitIndex - 1[\s\S]*this\.#unitIndex \+ 1[\s\S]*record\.element\.style\.width/);
  assert.match(source, /#queueContinuousLayout\(preserveAnchor = false\)[\s\S]*anchorElement[\s\S]*\.continuous-view"\)\.style\.removeProperty\("width"\)[\s\S]*workspace\.scrollTop \+= settledRect\.top - anchorRect\.top/);
  assert.match(source, /const finish = \(\): void => \{[\s\S]*this\.#navigationTransitioning = false;[\s\S]*this\.#queueContinuousLayout\(true\)/);
  assert.match(source, /event\.propertyName !== "grid-template-columns"[\s\S]*finish\(\)/);
});

test("printing prepares every document unit independently of the screen page mode", () => {
  assert.match(source, /else if \(action === "print"\) void this\.#printDocument\(\)/);
  assert.match(source, /#printDocument\(\): Promise<void> \{[\s\S]*for \(const unit of info\.units\)[\s\S]*document\.render\(\{[\s\S]*unitIndex: unit\.index/);
  assert.match(source, /sheetPrintPages\(unit, sheetSizes\)[\s\S]*for \(const sheetPage of printPages\)[\s\S]*sheetPage\?\.fragments[\s\S]*for \(const fragment of fragments\)[\s\S]*viewport: fragment\.viewport/);
  assert.match(source, /#printDocument\(\): Promise<void> \{[\s\S]*canvas\.toBlob[\s\S]*URL\.createObjectURL[\s\S]*printFrame\.contentWindow[\s\S]*printWindow\.print\(\)/);
  assert.match(source, /document\.body\.append\(printFrame\)[\s\S]*printWindow\.addEventListener\("afterprint",[\s\S]*URL\.revokeObjectURL/);
  assert.match(source, /@page \{ margin: 0; \}[\s\S]*\.print-page img \{[^}]*max-width: 99vw;[^}]*max-height: 99vh;/s);
  assert.doesNotMatch(source, /\bwindow\.print\(\)/);
});

test("visible Viewer thumbnails retry after an interrupted prefetch", () => {
  const thumbnailRenderer = source.slice(
    source.indexOf("async #renderThumbnail"),
    source.indexOf("#toggleNavigation", source.indexOf("async #renderThumbnail")),
  );
  assert.match(source, /preview\.dataset\.visible = String\(entry\.isIntersecting\)/);
  assert.match(thumbnailRenderer, /OPERATION_ABORTED[\s\S]*retry\s*=/);
  assert.match(thumbnailRenderer,
    /this\.#thumbnailRequests\.delete\(unitIndex\)[\s\S]*if \(retry && document === this\.#document\)/);
  assert.match(thumbnailRenderer, /void this\.#renderThumbnail\(unitIndex\)/);
});

test("focused navigation thumbnails use vertical arrows and remain visible", () => {
  assert.match(source,
    /target\.closest\("\.unit-button, \.outline-button"\) !== null[\s\S]*event\.key === "ArrowUp"[\s\S]*event\.key === "ArrowDown"/);
  assert.match(source, /this\.#stepNavigationUnit\(event\.key === "ArrowDown" \? 1 : -1\)/);
  assert.match(source, /button\.tabIndex = current \? 0 : -1/);
  assert.match(source, /focus\(\{ preventScroll: true \}\)/);
  assert.match(source, /navigation\.scrollTop \+= buttonBounds\.(?:top|bottom) - navigationBounds\.(?:top|bottom)/);
});

test("desktop navigation width is pointer and keyboard adjustable", () => {
  assert.match(source, /class="navigation-resizer"[^>]*role="separator"[^>]*aria-orientation="vertical"/);
  assert.match(source, /\.navigation-resizer \{[^}]*cursor: col-resize;[^}]*touch-action: none;/s);
  assert.match(source, /#beginNavigationResize\(event as PointerEvent\)/);
  assert.match(source, /#updateNavigationResize\(event as PointerEvent\)/);
  assert.match(source, /#finishNavigationResize\(event as PointerEvent\)/);
  assert.match(source, /target\.classList\.contains\("navigation-resizer"\)[\s\S]*event\.key === "ArrowLeft"/);
  assert.match(source, /--dv-navigation-width/);
});

test("single-page wheel gestures settle on one page until a new gesture", () => {
  assert.match(source, /addEventListener\("wheel", \(event\) => this\.#handleWheel\(event\), \{ passive: false \}\)/);
  assert.match(source, /workspace\.scrollTop >= maxScrollTop - 2/);
  assert.match(source, /workspace\.scrollTop <= 2/);
  assert.match(source, /Math\.abs\(this\.#wheelDelta\) < threshold/);
  assert.match(source, /this\.#wheelPageLocked = true/);
  assert.match(source, /if \(this\.#wheelPageLocked\) \{\s*event\.preventDefault\(\);\s*this\.#scheduleWheelReset\(\);\s*return;/);
  assert.match(source, /this\.#wheelResetTimer = setTimeout\(\(\) => this\.#resetWheelGesture\(\), 260\)/);
  assert.match(source, /this\.#selectUnit\(this\.#unitIndex \+ direction\)/);
  assert.match(source, /event\.ctrlKey/);
  assert.doesNotMatch(source, /wheelCooldownUntil|wheelGestureNavigated|wheelUnlockAt|wheelLastMagnitude|newImpulse/);
});

test("spreadsheet wheel gestures only scroll the sheet viewport", () => {
  const start = source.indexOf("\n  #handleWheel(event: WheelEvent)");
  const wheelHandler = source.slice(start, source.indexOf("\n  #scheduleWheelReset(): void", start));
  assert.match(wheelHandler,
    /this\.#info\?\.units\[this\.#unitIndex\]\?\.type === "sheet"[\s\S]*this\.#resetWheelGesture\(\);\s*return;/);
});

test("pinch gestures zoom the Viewer without scaling the browser page", () => {
  assert.match(source, /if \(event\.ctrlKey\) \{\s*event\.preventDefault\(\)/);
  assert.match(source, /this\.#handlePinchZoom\(event\)/);
  assert.match(source, /Math\.exp\(-delta \* \.01\)/);
  assert.match(source, /const zoom = this\.#setZoom\(target, false, event, true\)/);
  assert.match(source, /zoom !== undefined[\s\S]*this\.#previewPinchZoom\(zoom, event[\s\S]*this\.#schedulePinchRender\(\)/);
  assert.match(source, /stage\.style\.transformOrigin =/);
  assert.match(source, /#setZoom\([\s\S]*value: number,[\s\S]*snap = true,[\s\S]*focalPoint\?/);
});

test("pinch gestures defer high-resolution rendering until the gesture settles", () => {
  const continuousLayout = source.slice(
    source.indexOf("#queueContinuousLayout(preserveAnchor = false): void"),
    source.indexOf("async #buildContinuousPages", source.indexOf("#queueContinuousLayout(preserveAnchor = false): void")),
  );
  assert.match(source, /const PINCH_RENDER_DELAY = 120/);
  assert.match(source, /#pinchZoomTimer: ReturnType<typeof setTimeout> \| undefined/);
  assert.match(source, /#schedulePinchRender\(\): void \{[\s\S]*setTimeout\(\(\) => this\.#commitPinchZoom\(\), PINCH_RENDER_DELAY\)/);
  assert.match(source, /#commitPinchZoom\(\): void \{[\s\S]*this\.#queueContinuousLayout\(\)[\s\S]*this\.#queueSingleLayout\(\)/);
  assert.match(source, /#setZoom\([\s\S]*deferRender = false[\s\S]*if \(!deferRender\) \{[\s\S]*this\.#queueContinuousLayout\(\)/);
  assert.match(source, /#previewPinchZoom\([\s\S]*this\.#applyContinuousZoomAnchor\(\);\s*return;/);
  assert.match(source, /#previewPinchZoom\([\s\S]*this\.#previewSinglePageScale\(targetScale\)[\s\S]*this\.#applySingleZoomAnchor\(rendered\.frame, targetScale\);\s*return;/);
  assert.match(source, /#previewSinglePageScale\(targetScale: number\): void \{[\s\S]*canvas\.style\.width =[\s\S]*stage\.style\.width =/);
  assert.match(source, /#animatePinchSettle\([\s\S]*duration: 160[\s\S]*cubic-bezier\(\.22, 1, \.36, 1\)/);
  assert.doesNotMatch(source, /#previewPinchZoom\([\s\S]*this\.#applyContinuousZoomAnchor\(true\)/);
  assert.doesNotMatch(continuousLayout, /this\.#applyContinuousZoomAnchor\(\)/);
  assert.match(browserMatrix, /highResolutionDeadline = performance\.now\(\) \+ 3_000/);
  assert.match(browserMatrix, /canvas\.width \/ pageRect\.width >= window\.devicePixelRatio \* \.9/);
});

test("continuous cache eviction never discards the current or rendering page", () => {
  const eviction = source.slice(
    source.indexOf("#evictContinuousPages"),
    source.indexOf("#clearContinuousPages", source.indexOf("#evictContinuousPages")),
  );
  assert.match(eviction, /unitIndex !== this\.#unitIndex/);
  assert.match(eviction, /!record\.rendering/);
});

test("Viewer stops zooming before the render pixel limit without entering an error state", () => {
  const setZoom = source.slice(
    source.indexOf("#canRenderAtZoom"),
    source.indexOf("#fitView", source.indexOf("#canRenderAtZoom")),
  );
  const renderCurrent = source.slice(
    source.indexOf("async #renderCurrent"),
    source.indexOf("#queueSingleLayout", source.indexOf("async #renderCurrent")),
  );
  assert.match(source, /#renderPixelLimit = DEFAULT_LIMITS\.renderPixels/);
  assert.match(source, /engineOptions\.limits\?\.renderPixels \?\? DEFAULT_LIMITS\.renderPixels/);
  assert.match(setZoom, /pixels > this\.#renderPixelLimit/);
  assert.match(setZoom, /zoom > this\.#zoom && !this\.#canRenderAtZoom\(zoom\)/);
  assert.match(renderCurrent, /cause\.code === "RENDER_PIXEL_LIMIT"[\s\S]*this\.#zoom = this\.#rendered\.scale/);
  assert.match(renderCurrent, /this\.#emitState\(\);\s*return;\s*}\s*this\.#setStatus\("error"/);
});

test("Viewer keeps the pointer document point anchored while zooming", () => {
  const zoomAnchor = source.slice(
    source.indexOf("#captureZoomAnchor"),
    source.indexOf("#fitView", source.indexOf("#captureZoomAnchor")),
  );
  assert.match(source, /interface ZoomAnchor/);
  assert.match(zoomAnchor, /focalPoint\?\.clientX \?\? workspaceRect\.left \+ workspaceRect\.width \/ 2/);
  assert.doesNotMatch(zoomAnchor, /focalPoint\?\.clientY/);
  assert.match(zoomAnchor, /this\.#zoomAnchor = deferRender[\s\S]*:\s*this\.#captureZoomAnchor\(zoom, focalPoint\)/);
  assert.match(source, /this\.#applySingleZoomAnchor\(frame, scale\)/);
  assert.match(source, /this\.#applyContinuousZoomAnchor\(\)/);
  assert.match(source, /if \(!continuing\) this\.#pinchVerticalAnchor = this\.#capturePinchVerticalAnchor\(event\)/);
  assert.match(source, /#capturePinchVerticalAnchor\(event: WheelEvent\)[\s\S]*documentY: \(clientY - rect\.top\) \/ record\.scale/);
  assert.match(zoomAnchor, /this\.#pinchVerticalAnchor\?\.unitIndex \?\? this\.#unitIndex/);
  assert.match(zoomAnchor, /workspace\.scrollWidth <= workspace\.clientWidth[\s\S]*workspace\.scrollLeft = 0/);
  assert.match(zoomAnchor, /workspace\.scrollLeft \+=/);
  assert.match(zoomAnchor, /workspace\.scrollTop \+= rect\.top \+ anchor\.documentY \* record\.scale - anchor\.clientY/);
  assert.doesNotMatch(source, /page-top|pageTop/);
});

test("single-page navigation animates in the requested direction", () => {
  assert.match(source, /transitionDirection = Math\.sign\(unitIndex - this\.#unitIndex\)/);
  assert.match(source, /this\.#renderCurrent\(transitionDirection\)/);
  assert.match(source, /className = "page-transition-outgoing"/);
  assert.match(source, /currentCanvas\.before\(canvas\)/);
  assert.match(source, /translateY\(\$\{direction \* 6\}px\)/);
  assert.match(source, /translateY\(\$\{-direction \* 4\}px\)/);
  assert.match(source, /prefers-reduced-motion: reduce/);
  assert.match(source, /this\.#pageAnimations = animations/);
  assert.match(source, /this\.#outgoingPage = outgoing/);
  assert.match(source, /this\.#info\?\.format !== "pdf"/);
  assert.doesNotMatch(source, /translateY\(\$\{direction \* 14\}px\) scale\(\.995\)/);
});

test("Viewer cleans up document, bitmap, thumbnail, and engine resources", () => {
  assert.match(source, /this\.#document\?\.close\(\)/);
  assert.match(source, /if \(frame !== undefined\) frame\.bitmap\.close\(\)/);
  assert.match(source, /frame\?\.bitmap\.close\(\)/);
  assert.match(source, /this\.#thumbnailRequests\.clear\(\)/);
  assert.match(source, /this\.#engine\?\.close\(\)/);
  assert.match(source, /this\.#openController\?\.abort\(\)/);
  assert.match(source, /this\.#continuousObserver\?\.disconnect\(\)/);
  assert.match(source, /this\.#continuousNavigationObserver\?\.disconnect\(\)/);
  assert.match(source, /document\.removeEventListener\("fullscreenchange", this\.#handleFullscreenChange\)/);
});

test("Viewer treats first-render failure as an open failure", () => {
  const renderCurrent = source.slice(
    source.indexOf("async #renderCurrent"),
    source.indexOf("#queueSingleLayout", source.indexOf("async #renderCurrent")),
  );
  assert.match(renderCurrent, /catch \(cause\)[\s\S]*throw cause;/);
});

test("Viewer shares one in-flight Engine creation across concurrent opens", () => {
  assert.match(source, /#enginePromise:/);
  assert.match(source, /#getEngine\(\): Promise<OfficeEngine>/);
  assert.doesNotMatch(source, /#engine \?\?= await createOfficeEngine/);
});

test("Viewer thumbnails release ImageBitmaps after copying them to canvas", () => {
  const thumbnailRenderer = source.slice(
    source.indexOf("async #renderThumbnail"),
    source.indexOf("#toggleNavigation", source.indexOf("async #renderThumbnail")),
  );
  assert.match(thumbnailRenderer, /this\.#drawThumbnail\(unitIndex, frame\.bitmap, frame\.viewport\)[\s\S]*finally \{\s*frame\?\.bitmap\.close\(\)/);
  assert.doesNotMatch(source, /#thumbnailFrames/);
});

test("Viewer bounds long-document navigation DOM", () => {
  assert.match(source, /const NAVIGATION_WINDOW_SIZE = 40/);
  assert.match(source, /units\.slice\(start, end\)/);
  assert.match(source, /aria-setsize/);
  assert.match(source, /#renderNavigationWindow\(center/);
});

test("Viewer retains overlapping thumbnail canvases while shifting the navigation window", () => {
  const navigationRenderer = source.slice(
    source.indexOf("#renderNavigationWindow(center"),
    source.indexOf("async #renderThumbnail", source.indexOf("#renderNavigationWindow(center")),
  );
  assert.match(navigationRenderer, /const retainedButtons = new Map\(/);
  assert.match(navigationRenderer, /list\.querySelectorAll<HTMLElement>\("\.unit-button"\)/);
  assert.match(navigationRenderer, /const retained = retainedButtons\.get\(unit\.index\)/);
  assert.match(navigationRenderer, /fragment\.append\(retained\);\s*continue;/);
});
