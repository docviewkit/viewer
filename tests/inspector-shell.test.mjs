import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const html = await readFile(new URL("../examples/inspector.html", import.meta.url), "utf8");
const script = await readFile(new URL("../examples/inspector.js", import.meta.url), "utf8");

test("inspector imports resolve before document initialization", async () => {
  for (const [, names, path] of script.matchAll(/import\s*\{([^}]+)\}\s*from\s*"([^"]+)"/gu)) {
    const module = await import(new URL(path, new URL("../examples/inspector.js", import.meta.url)));
    for (const name of names.split(",").map((name) => name.trim()).filter(Boolean)) {
      assert.ok(name in module, `${path} must export ${name}`);
    }
  }
});

function startTag(id) {
  const match = html.match(new RegExp(`<[^>]+\\bid=["']${id}["'][^>]*>`, "i"));
  assert.ok(match, `expected #${id} in inspector.html`);
  return match[0];
}

test("inspector exposes an accessible full-height Office viewer shell", () => {
  assert.match(html, /<a[^>]+class="skip-link"[^>]+href="#viewer"/);
  assert.match(html, /<header[^>]+class="titlebar"/);
  assert.match(html, /<nav[^>]+class="commandbar"[^>]+aria-label="文档命令"/);
  assert.match(html, /<nav[^>]+id="unit-navigation"[^>]+aria-label="文档导航"/);
  assert.match(html, /<main[^>]+id="viewer"/);
  assert.match(html, /<aside[^>]+id="inspector"/);
  assert.match(html, /<footer[^>]+class="statusbar"/);
  assert.match(html, /min-height:\s*100dvh/);
});

test("navigation and inspector controls publish their ARIA state", () => {
  assert.match(startTag("unit-list"), /role="listbox"/);
  assert.match(startTag("inspector-toggle"), /aria-controls="inspector"/);
  assert.match(startTag("inspector-toggle"), /aria-expanded="true"/);
  assert.match(startTag("zoom-range"), /type="range"/);
  assert.match(startTag("zoom-range"), /aria-label="缩放"/);
  assert.match(startTag("document-tab"), /aria-controls="diagnostic-panel"/);
  assert.match(startTag("render-tab"), /tabindex="-1"/);
});

test("script switches navigation presentation by document kind", () => {
  assert.match(script, /function\s+navigationPresentation\s*\(/);
  assert.match(script, /presentation[\s\S]*幻灯片/);
  assert.match(script, /spreadsheet[\s\S]*工作表/);
  assert.match(script, /text[\s\S]*页面/);
  assert.match(script, /ui\.app\.dataset\.documentKind/);
  assert.match(script, /ui\.unitList\.dataset\.navigationMode/);
});

test("keyboard behavior covers unit navigation, diagnostic tabs, opening, and zoom", () => {
  assert.match(script, /function\s+handleUnitListKeydown\s*\(/);
  assert.match(script, /ArrowDown/);
  assert.match(script, /ArrowRight/);
  assert.match(script, /Home/);
  assert.match(script, /End/);
  assert.match(script, /PageUp/);
  assert.match(script, /PageDown/);
  assert.match(script, /event\.key\.toLowerCase\(\)\s*===\s*"o"/);
  assert.match(script, /function\s+handleDiagnosticTabKeydown\s*\(/);
  assert.match(script, /function\s+setInspectorExpanded\s*\(/);
});

test("loaded state hides the welcome uploader without removing file access", () => {
  assert.match(html, /\.app\[data-document-open="true"\]\s+\.welcome-drop\s*\{[^}]*display:\s*none/);
  assert.match(startTag("open-file-button"), /for="file-input"/);
  assert.match(script, /function\s+setDocumentOpen\s*\(/);
  assert.match(script, /setDocumentOpen\(true\)/);
  assert.match(script, /setDocumentOpen\(false\)/);
});

test("viewer passes the original filename to lazy format dispatch", () => {
  assert.match(script, /fileName:\s*file\.name/);
});

test("every strict JavaScript element lookup has matching markup", () => {
  const ids = [...script.matchAll(/element\("([^"]+)"\)/g)].map((match) => match[1]);
  assert.ok(ids.length > 0);
  for (const id of ids) startTag(id);
});

test("shell avoids embedded SVG and emoji artwork", () => {
  assert.doesNotMatch(html, /<svg\b|data:image\/svg\+xml/i);
  assert.doesNotMatch(html, /[\u{1F300}-\u{1FAFF}]/u);
});

test("viewer exposes document text search with keyboard navigation", () => {
  assert.match(html, /role="search"/);
  assert.match(startTag("search-input"), /type="search"/);
  assert.match(startTag("search-input"), /aria-label="搜索文档"/);
  assert.match(startTag("search-previous"), /aria-label="上一个搜索结果"/);
  assert.match(startTag("search-next"), /aria-label="下一个搜索结果"/);
  assert.match(startTag("search-status"), /aria-live="polite"/);
  assert.match(script, /\.searchText\s*\(/);
  assert.match(script, /event\.key\.toLowerCase\(\)\s*===\s*"f"/);
  assert.match(script, /function\s+focusSearchResult\s*\(/);
});

test("viewer exposes an accessible fullscreen command", () => {
  assert.match(startTag("fullscreen-button"), /aria-label="切换全屏"/);
  assert.match(script, /requestFullscreen\s*\(/);
  assert.match(script, /document\.exitFullscreen\s*\(/);
  assert.match(script, /fullscreenchange/);
});

test("slide documents expose a fullscreen slideshow mode", () => {
  assert.match(startTag("slideshow-button"), /aria-label="播放幻灯片"/);
  assert.match(startTag("slideshow-button"), /\bhidden\b/);
  assert.match(startTag("slideshow-controls"), /aria-label="幻灯片播放控制"/);
  assert.match(startTag("slideshow-previous"), /aria-label="上一张幻灯片"/);
  assert.match(startTag("slideshow-next"), /aria-label="下一张幻灯片"/);
  assert.match(startTag("slideshow-exit"), /aria-label="退出幻灯片播放"/);
  assert.match(script, /selectedUnit\(\)\?\.type\s*===\s*"slide"/);
  assert.match(script, /function\s+startSlideshow\s*\(/);
  assert.match(script, /ui\.viewer\.requestFullscreen\s*\(/);
  assert.match(script, /function\s+stopSlideshow\s*\(/);
  assert.match(script, /function\s+closeActiveSession\s*\(\)[\s\S]*leaveSlideshowState\(\)/);
  assert.match(script, /state\.slideshow[\s\S]*ArrowRight[\s\S]*PageDown/);
  assert.match(html, /\.viewer\[data-playing="true"\][\s\S]*background:\s*#111318/);
});

test("viewer prints every document unit in an isolated frame", () => {
  assert.match(startTag("print-button"), /aria-label="打印文档"/);
  assert.match(script, /function\s+printDocument\s*\([\s\S]*for \(const unit of officeDocument\.info\.units\)/);
  assert.match(script, /sheetPrintPages\(unit, sheetSizes\)[\s\S]*for \(const sheetPage of printPages\)[\s\S]*viewport: sheetPage\.viewport/);
  assert.match(script, /printFrame\.contentWindow[\s\S]*printWindow\.print\s*\(/);
  assert.doesNotMatch(script, /\bwindow\.print\s*\(/);
  assert.match(script, /function\s+printDocument\s*\(/);
  assert.match(script, /URL\.createObjectURL\(blob\)/);
  assert.match(script, /URL\.revokeObjectURL\(url\)/);
});

test("viewer builds a selectable text layer and safe visible hyperlinks", () => {
  assert.match(startTag("text-layer"), /aria-label="可选择文本层"/);
  assert.match(html, /\.text-layer-item\s*\{[\s\S]*?user-select:\s*text/);
  assert.match(script, /\.listObjects\s*\(/);
  assert.match(script, /function\s+renderTextLayer\s*\(/);
  assert.match(script, /function\s+safeVisibleLink\s*\(/);
  assert.match(script, /rel\s*=\s*"noopener noreferrer"/);
});

test("viewer exposes local user-controlled playback for embedded media", () => {
  assert.match(startTag("media-layer"), /aria-label="嵌入式媒体播放层"/);
  assert.match(script, /function\s+renderMediaLayer\s*\(/);
  assert.match(script, /URL\.createObjectURL\(new Blob\(\[media\.bytes\]/);
  assert.match(script, /player\.controls\s*=\s*true/);
  assert.match(script, /player\.playsInline\s*=\s*true/);
  assert.match(script, /player\.preload\s*=\s*"metadata"/);
  assert.match(script, /fragment\.append\(player,\s*button\)/);
  assert.doesNotMatch(script, /button\.append\(player\)/);
  assert.match(script, /URL\.revokeObjectURL\(url\)/);
  assert.doesNotMatch(script, /player\.autoplay\s*=\s*true/);
});

test("font inspector reports authored and rendered fonts for selected text", () => {
  assert.match(html, /id="font-selected-text"[^>]*>请在文档中选择文字</);
  assert.match(html, /id="font-authored-family"/);
  assert.match(html, /id="font-rendered-family"/);
  assert.match(html, /id="font-source"/);
  assert.match(script, /document\.addEventListener\("selectionchange"/);
  assert.match(script, /object\.fontRuns/);
  assert.match(script, /run\.authoredFamily/);
  assert.match(script, /run\.renderedFamily/);
  assert.match(script, /function\s+displayedFontFamily\s*\(/);
  assert.match(script, /run\.source\s*===\s*"fallback"\s*\?\s*run\.renderedFamily\s*:\s*run\.authoredFamily/);
  assert.match(script, /function\s+renderFontFamilies\s*\(/);
  assert.match(script, /preview\.style\.fontFamily\s*=\s*cssFontFamily\(family\)/);
  assert.match(html, /\.font-family-preview\s*\{/);
  assert.match(script, /function\s+itemFontRuns\s*\(/);
  assert.match(script, /state\.fontGestureItem/);
  assert.match(script, /document\.addEventListener\("pointerup"/);
});

test("selectable text deduplication prefers the object carrying font runs", () => {
  assert.match(script, /const\s+visibleObjects\s*=\s*new Map/);
  assert.match(script, /existing\.object\.fontRuns/);
  assert.match(script, /object\.fontRuns/);
  assert.match(script, /visibleObjects\.set\(key,\s*\{ object, bounds \}\)/);
});

test("font inspector lists only fonts resolved for the current document", () => {
  assert.match(startTag("font-list"), /aria-label="当前文档字体列表"/);
  assert.match(startTag("font-list-status"), /aria-live="polite"/);
  assert.match(html, /id="document-fonts-heading"[^>]*>文档字体</);
  assert.match(script, /state\.documentFonts/);
  assert.match(script, /object\.fontRuns/);
  assert.doesNotMatch(html, /读取系统字体|font-list-button/);
  assert.doesNotMatch(script, /queryLocalFonts|document\.fonts|systemFonts|readSystemFonts/);
});

test("sheets use a virtual scroll surface and bounded rerendering", () => {
  assert.match(script, /function\s+requestedSheetScale\s*\(/);
  assert.match(script, /function\s+sheetViewport\s*\(unit,\s*scale\)/);
  assert.match(script, /ui\.viewer\.clientWidth\s*\/\s*scale/);
  assert.match(script, /ui\.viewer\.clientHeight\s*\/\s*scale/);
  assert.match(script, /SHEET_VIEWPORT_OVERSCAN_RATIO\s*=\s*0\.5/);
  assert.match(script, /visibleWidth\s*\*\s*\(1\s*\+\s*SHEET_VIEWPORT_OVERSCAN_RATIO\)/);
  assert.match(script, /visibleHeight\s*\*\s*\(1\s*\+\s*SHEET_VIEWPORT_OVERSCAN_RATIO\)/);
  assert.match(script, /const\s+visibleWidth\s*=\s*Math\.min\(layout\.columns\.total,\s*ui\.viewer\.clientWidth\s*\/\s*scale\)/);
  assert.match(script, /const\s+visibleHeight\s*=\s*Math\.min\(layout\.rows\.total,\s*ui\.viewer\.clientHeight\s*\/\s*scale\)/);
  assert.match(script, /function\s+queueSheetViewportRender\s*\(/);
  assert.match(script, /ui\.viewer\.addEventListener\("scroll"/);
  assert.match(script, /state\.sheetViewportOrigin/);
  assert.match(script, /state\.renderRevision\s*!==\s*displayed\.renderRevision/);
  assert.match(script, /preserveDisplayedSheet/);
  assert.match(html, /surface-wrap\[data-sheet="true"\]\s*\{[^}]*background:\s*#fdfefe/s);
  const queueStart = script.indexOf("function queueSheetViewportRender()");
  const queueEnd = script.indexOf("\nfunction fitScale", queueStart);
  const queue = script.slice(queueStart, queueEnd);
  assert.ok(queue.indexOf("void renderCurrentUnit();") < queue.indexOf("setTimeout("));
});

test("sheet headers expose virtual row and column labels with accessible resizers", () => {
  assert.match(startTag("sheet-column-headers"), /aria-label="列标题"/);
  assert.match(startTag("sheet-row-headers"), /aria-label="行标题"/);
  assert.match(script, /function\s+renderSheetHeaders\s*\(/);
  assert.match(script, /role",\s*"separator"/);
  assert.match(script, /function\s+beginSheetResize\s*\(/);
  assert.match(script, /function\s+finishSheetResize\s*\(/);
  assert.match(script, /sheetSizes:\s*sheetSizeRequest\(unit\)/);
  assert.match(script, /resize\.axis\s*===\s*"column"[\s\S]*?style\.width\s*=\s*"1px"/);
  assert.match(script, /else\s*\{[\s\S]*?style\.height\s*=\s*"1px"/);
  assert.match(script, /function\s+autoFitSheetSize\s*\([\s\S]*?sheetAutoFitColumnWidth\(context, cells, MIN_COLUMN_WIDTH\)/);
  assert.match(script, /addEventListener\("dblclick",[\s\S]*?autoFitSheetSize\(event, target\)/);
});

test("sheet headers stay compositor-sticky while the native scroller moves", () => {
  assert.match(html, /\.sheet-header-axis,\s*\.sheet-header-corner\s*\{[^}]*position:\s*sticky/s);
  const positionStart = script.indexOf("function positionSheetHeaders");
  const positionEnd = script.indexOf("\nfunction commitSheetSize", positionStart);
  const position = script.slice(positionStart, positionEnd);
  assert.match(position, /const\s+stageLeft\s*=\s*SHEET_ROW_HEADER_WIDTH\s*\+\s*viewport\.x\s*\*\s*scale/);
  assert.match(position, /const\s+stageTop\s*=\s*SHEET_COLUMN_HEADER_HEIGHT\s*\+\s*viewport\.y\s*\*\s*scale/);
  assert.match(position, /sheetColumnHeaders\.style\.transform\s*=\s*`translate3d\(\$\{stageLeft\}px, 0, 0\)`/);
  assert.match(position, /sheetRowHeaders\.style\.transform\s*=\s*`translate3d\(0, \$\{stageTop\}px, 0\)`/);
  assert.match(position, /sheetHeaderCorner\.style\.transform\s*=\s*"translate3d\(0, 0, 0\)"/);
  assert.doesNotMatch(position, /style\.(?:left|top)\s*=/);
});

test("sheet header gutters mask scrolled worksheet content", () => {
  assert.match(startTag("sheet-header-mask-top"), /aria-hidden="true"/);
  assert.match(startTag("sheet-header-mask-left"), /aria-hidden="true"/);
  assert.match(html, /\.sheet-header-mask\s*\{[^}]*position:\s*sticky[^}]*background:\s*#fdfefe/s);
  assert.match(html, /\.sheet-header-mask\s*\{[^}]*z-index:\s*6/s);
  assert.match(html, /\.sheet-header-axis,\s*\.sheet-header-corner\s*\{[^}]*z-index:\s*5/s);
  assert.match(html, /\.sheet-header-mask-top\s*\{[^}]*height:\s*var\(--surface-inset\)/s);
  assert.match(html, /\.sheet-header-mask-left\s*\{[^}]*width:\s*var\(--surface-inset\)/s);
});

test("background sheet rerenders keep the displayed status stable", () => {
  const renderStart = script.indexOf("async function renderCurrentUnit");
  const preserveStart = script.indexOf("const preserveDisplayedSheet", renderStart);
  const diagnosticsReset = script.indexOf("state.renderDiagnostics = [];", preserveStart);
  const loadingStatus = script.indexOf("setStatus(`正在渲染${unitLabel(unit)}…`, \"loading\");", preserveStart);
  assert.ok(loadingStatus > preserveStart && loadingStatus < diagnosticsReset);
});

test("background sheet rerenders commit the main and frozen canvases together", () => {
  const renderStart = script.indexOf("async function renderCurrentUnit");
  const frozenFramesReady = script.indexOf("await renderFrozenPaneFrames(", renderStart);
  const mainCanvasCommit = script.indexOf("ui.surface.width = frame.pixelWidth;", renderStart);
  const frozenCanvasCommit = script.indexOf("drawFrozenPaneFrames(", mainCanvasCommit);
  assert.ok(frozenFramesReady > renderStart && frozenFramesReady < mainCanvasCommit);
  assert.ok(frozenCanvasCommit > mainCanvasCommit);
});

test("main canvas automatically rerenders after its 2D context is restored", () => {
  assert.match(script, /function\s+handleSurfaceContextLost\s*\(/);
  assert.match(script, /function\s+handleSurfaceContextRestored\s*\(/);
  assert.match(script, /ui\.surface\.addEventListener\("contextlost",\s*handleSurfaceContextLost\)/);
  assert.match(script, /ui\.surface\.addEventListener\("contextrestored",\s*handleSurfaceContextRestored\)/);
  const lostStart = script.indexOf("function handleSurfaceContextLost");
  const restoredStart = script.indexOf("function handleSurfaceContextRestored", lostStart);
  const listenersStart = script.indexOf('ui.surface.addEventListener("contextlost"', restoredStart);
  const lost = script.slice(lostStart, restoredStart);
  const restored = script.slice(restoredStart, listenersStart);
  assert.match(lost, /state\.surfaceContextLost\s*=\s*true/);
  assert.match(lost, /state\.renderRevision\s*\+=\s*1/);
  assert.match(lost, /clearFrozenSurfaces\s*\(\s*\)/);
  assert.doesNotMatch(lost, /preventDefault\s*\(/);
  assert.match(restored, /state\.surfaceContextLost\s*=\s*false/);
  assert.match(restored, /void\s+renderCurrentUnit\s*\(/);
});

test("virtual sheets composite authored frozen rows and columns", () => {
  assert.match(script, /function\s+frozenPaneRequests\s*\(/);
  assert.match(script, /function\s+renderFrozenPaneFrames\s*\(/);
  assert.match(script, /function\s+drawFrozenPaneFrames\s*\(/);
  assert.match(script, /unit\.frozenRows/);
  assert.match(script, /unit\.frozenColumns/);
  assert.match(script, /layout\.columns\.offset\(unit\.frozenColumns\)/);
  assert.match(script, /layout\.rows\.offset\(unit\.frozenRows\)/);
  assert.match(script, /Promise\.allSettled\(requests\.map/);
});

test("text documents expose a lazy continuous-page mode", () => {
  assert.match(startTag("continuous-button"), /aria-pressed="false"/);
  assert.match(startTag("continuous-view"), /aria-label="连续文档"/);
  assert.match(script, /IntersectionObserver/);
  assert.match(script, /function\s+setContinuousMode\s*\(/);
  assert.match(script, /function\s+renderContinuousPage\s*\(/);
  assert.match(script, /CONTINUOUS_CACHE_PIXELS/);
  assert.match(script, /function\s+evictContinuousPixels\s*\(/);
  assert.match(script, /canvas\.width\s*=\s*1/);
  assert.doesNotMatch(html, /continuous-page-number/);
  assert.doesNotMatch(script, /continuous-page-number/);
});

test("continuous scrolling selects and reveals the current page thumbnail", () => {
  assert.match(script, /function\s+selectContinuousUnit\s*\(unitIndex\)/);
  assert.match(script, /Number\(option\.value\)\s*===\s*unitIndex/);
  assert.match(script, /ui\.unitSelect\.selectedIndex\s*=\s*position/);
  assert.match(script, /button\?\.scrollIntoView\(\{\s*block:\s*"nearest"\s*\}\)/);
  assert.match(script, /state\.continuousNavigationObserver\s*=\s*new IntersectionObserver/);
  assert.match(script, /selectContinuousUnit\(unitIndex\)/);
  assert.match(script, /rootMargin:\s*"-45% 0px -45% 0px"/);
});

test("continuous page jumps skip long smooth-scroll animations", () => {
  assert.match(script, /const\s+CONTINUOUS_SMOOTH_SCROLL_VIEWPORTS\s*=\s*2/);
  assert.match(script, /function\s+revealContinuousPage\s*\(/);
  assert.match(script, /pageBounds\.bottom\s*<\s*viewerBounds\.top/);
  assert.match(script, /pageBounds\.top\s*>\s*viewerBounds\.bottom/);
  assert.match(script, /distance\s*<=\s*smoothLimit\s*\?\s*"smooth"\s*:\s*"auto"/);
  assert.match(script, /revealContinuousPage\(record/);
});

test("locatable diagnostics navigate to their page and highlighted bounds", () => {
  assert.match(script, /function\s+diagnosticTarget\s*\(/);
  assert.match(script, /function\s+diagnosticObjectId\s*\(/);
  assert.match(script, /function\s+resolveDiagnosticTarget\s*\(/);
  assert.match(script, /function\s+focusDiagnostic\s*\(/);
  assert.match(script, /officeDocument\.getObject\(objectId\)/);
  assert.match(script, /officeDocument\.listObjects\(\{\s*limit:\s*MAX_DIAGNOSTIC_OBJECTS\s*\}\)/);
  assert.match(script, /normalizedPart\(object\.source\.part\)\s*===\s*part/);
  assert.match(script, /function\s+diagnosticSourceHints\s*\(/);
  assert.match(script, /function\s+sourceMatchesHints\s*\(/);
  assert.match(script, /row\.dataset\.locatable\s*=\s*String\(locatable\)/);
  assert.match(script, /row\.addEventListener\("click"/);
  assert.match(script, /event\.key\s*!==\s*"Enter"\s*&&\s*event\.key\s*!==\s*" "/);
  assert.match(html, /\.diagnostic-table tr\[data-locatable="true"\]/);
});

test("navigation thumbnails render only near the visible list and evict old pixels", () => {
  assert.match(script, /function\s+renderUnitThumbnail\s*\(/);
  assert.match(script, /function\s+renderUnitThumbnails\s*\(/);
  assert.match(script, /state\.thumbnailObserver\s*=\s*new IntersectionObserver/);
  assert.match(script, /root:\s*ui\.unitList/);
  assert.match(script, /THUMBNAIL_CACHE_PIXELS/);
  assert.match(script, /function\s+evictThumbnailPixels\s*\(/);
  assert.match(script, /void\s+renderUnitThumbnails\(layoutDocument,\s*sessionRevision\)/);
  assert.match(script, /preview\.dataset\.thumbnailState\s*=\s*"ready"/);
});

test("visible navigation thumbnails retry after an interrupted prefetch", () => {
  const renderStart = script.indexOf("async function renderUnitThumbnail");
  const renderEnd = script.indexOf("\nfunction releaseThumbnail", renderStart);
  const render = script.slice(renderStart, renderEnd);
  assert.match(render, /OPERATION_ABORTED[\s\S]*retry\s*=\s*record\.visible/);
  assert.match(render, /record\.rendering\s*=\s*false[\s\S]*if\s*\(retry/);
  assert.match(render, /expectedSession\s*===\s*state\.sessionRevision/);
  assert.match(render, /expectedRevision\s*===\s*state\.thumbnailRevision/);
  assert.match(render, /void\s+renderUnitThumbnail\(officeDocument,\s*unit,\s*expectedSession,\s*expectedRevision\)/);
});

test("browser QA can load only a whitelisted local fixture from the query string", () => {
  assert.match(script, /function\s+openFixtureFromQuery\s*\(/);
  assert.match(script, /new URLSearchParams\(location\.search\)/);
  assert.match(script, /\/tests\/fixtures\//);
  assert.match(script, /encodeURIComponent\(fixture\)/);
  for (const extension of ["pptx", "ppsx", "potx", "odp", "otp", "fodp", "xlsx", "xltx", "ods", "ots", "fods", "docx", "dotx", "odt", "ott", "csv", "rtf", "pdf"]) {
    assert.match(script, new RegExp(`\\b${extension}\\b`, "u"));
    assert.match(startTag("file-input"), new RegExp(`\\.${extension}(?:,|\")`, "u"));
  }
  assert.doesNotMatch(startTag("file-input"), /\.html|\.htm|\.txt/iu);
  assert.doesNotMatch(script, /\|html\|htm\|/u);
  assert.doesNotMatch(script, /\|txt\|/u);
});

test("viewer rejects HTML files that bypass the native picker through drag and drop", () => {
  assert.match(script, /function\s+unsupportedDocumentLabel\s*\(/);
  assert.match(script, /file\.type\.trim\(\)\.toLowerCase\(\)\s*===\s*"text\/html"/u);
  assert.match(script, /unsupportedDocumentLabel\(file\)/u);
  assert.match(script, /return "HTML\/HTM"/u);
  assert.match(script, /\$\{unsupportedLabel\} 文档不受支持/u);
});

test("viewer rejects TXT files that bypass the native picker through drag and drop", () => {
  assert.match(script, /\.txt\$/u);
  assert.match(script, /return "TXT"/u);
  assert.match(script, /\$\{unsupportedLabel\} 文档不受支持/u);
});

test("narrow command bar scrolls without collapsing command groups into each other", () => {
  assert.match(html, /@media\s*\(max-width:\s*720px\)[\s\S]*?\.commandbar\s*\{[^}]*justify-content:\s*flex-start[^}]*overflow-x:\s*auto/iu);
  assert.match(html, /@media\s*\(max-width:\s*720px\)[\s\S]*?\.command-group\s*\{[^}]*flex:\s*0\s+0\s+auto[^}]*min-width:\s*max-content/iu);
});
