import { spawn } from "node:child_process";
import { createInterface } from "node:readline";
import { chromium, firefox, webkit } from "playwright-core";

const engines = { chromium, firefox, webkit };
const requested = (process.env.DOCVIEWKIT_BROWSERS ?? "chromium,firefox,webkit")
  .split(",")
  .map((name) => name.trim())
  .filter(Boolean);
for (const name of requested) {
  if (!(name in engines)) throw new Error(`Unsupported browser: ${name}`);
}
const deviceScaleFactor = Number(process.env.DOCVIEWKIT_DEVICE_SCALE_FACTOR ?? 1);
if (!Number.isFinite(deviceScaleFactor) || deviceScaleFactor < 1 || deviceScaleFactor > 4) {
  throw new Error("DOCVIEWKIT_DEVICE_SCALE_FACTOR must be between 1 and 4");
}

const server = spawn(process.execPath, ["scripts/serve.mjs", "--port", "0"], {
  cwd: process.cwd(),
  stdio: ["ignore", "pipe", "inherit"],
});
const lines = createInterface({ input: server.stdout });
const url = await new Promise((resolve, reject) => {
  const timeout = setTimeout(() => reject(new Error("Viewer server did not start")), 10_000);
  lines.on("line", (line) => {
    const match = line.match(/https?:\/\/[^\s]+/u);
    if (match === null) return;
    clearTimeout(timeout);
    resolve(match[0]);
  });
  server.once("exit", (code) => reject(new Error(`Viewer server exited before startup (${code})`)));
});

try {
  for (const name of requested) {
    const executablePath = name === "chromium" ? process.env.DOCVIEWKIT_CHROMIUM_EXECUTABLE : undefined;
    const browser = await engines[name].launch({ headless: true, proxy: { server: url }, ...(executablePath === undefined ? {} : { executablePath }) });
    try {
      const page = await browser.newPage({ viewport: { width: 1440, height: 900 }, deviceScaleFactor });
      const errors = [];
      page.on("console", (message) => {
        if (message.type() === "error") errors.push(`console: ${message.text()}`);
      });
      page.on("pageerror", (error) => errors.push(`page: ${error.message}`));
      // A real proxy covers worker imports in Firefox, which page request routing misses.
      // Keep a production hostname so a legacy localhost exception cannot hide a license gate.
      await page.goto("http://customer.example.test/examples/viewer.html?fixture=visual-baseline.pptx");
      try {
        await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
      } catch (cause) {
        const state = await page.locator("docviewkit-viewer").evaluate(viewer => ({
          error: document.documentElement.dataset.error,
          status: viewer.state.status,
        }));
        throw new Error(`${name} production-origin load failed: ${JSON.stringify({ ...state, errors })}`, { cause });
      }
      const openSource = await page.locator("docviewkit-viewer").evaluate(async (viewer) => {
        const bytes = await (await fetch("/tests/fixtures/visual-baseline.pptx")).arrayBuffer();
        const { createOfficeEngine } = await import("/dist/index.js");
        const engine = await createOfficeEngine({ execution: "inline", license: { token: "invalid", publicKey: "", origin: location.origin } });
        let objectCount;
        try {
          const document = await engine.open(bytes);
          objectCount = (await document.listObjects({ unitIndex: 0 })).length;
          document.close();
        } finally {
          engine.close();
        }
        viewer.config = { ...viewer.config, license: { token: "expired", publicKey: "" }, minimal: true };
        const minimal = getComputedStyle(viewer.shadowRoot.querySelector(".toolbar")).display === "none";
        const slots = [...viewer.shadowRoot.querySelectorAll("slot")].every(slot => !slot.hidden);
        viewer.config = { ...viewer.config, minimal: false, watermark: "APACHE REGRESSION" };
        await viewer.open(new Uint8Array(bytes));
        const pixels = () => {
          const canvas = viewer.shadowRoot.querySelector(".stage canvas");
          return canvas.getContext("2d").getImageData(0, 0, canvas.width, canvas.height).data;
        };
        const marked = pixels();
        viewer.config = { ...viewer.config, watermark: undefined };
        await viewer.open(new Uint8Array(bytes));
        const plain = pixels();
        let changedPixels = 0;
        for (let index = 0; index < Math.min(marked.length, plain.length); index += 4) {
          if (marked[index] !== plain[index] || marked[index + 1] !== plain[index + 1] || marked[index + 2] !== plain[index + 2]) changedPixels++;
        }
        return { objectCount, minimal, slots, changedPixels, state: viewer.state.license.status, branded: viewer.shadowRoot.querySelector(".license-branding") !== null };
      });
      if (openSource.objectCount <= 0 || !openSource.minimal || !openSource.slots || openSource.changedPixels < 100
        || openSource.state !== "open-source" || openSource.branded) {
        throw new Error(`${name} restricted Apache-2.0 capabilities at a production origin: ${JSON.stringify(openSource)}`);
      }
      await page.goto(`${url}examples/viewer.html?fixture=Photo%20Formats%20-%20CGM-O12-XL-Pictures.xlsx`);
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
      await page.locator(".sheet-tab").nth(3).click();
      const glow = await page.evaluate(async () => {
        const { Core, DEFAULT_LIMITS } = await import("/dist/core.js");
        const { SceneRenderer } = await import("/dist/render.js");
        const core = await Core.create(await (await fetch("/dist/office-viewer-core.wasm")).arrayBuffer(), DEFAULT_LIMITS);
        const document = core.open(new Uint8Array(await (await fetch("/tests/fixtures/Photo%20Formats%20-%20CGM-O12-XL-Pictures.xlsx")).arrayBuffer()));
        let renderer;
        try {
          document.loadUnit(3);
          const objects = document.scene.objects.filter(object => object.unitIndex === 3 && object.visual.glow !== undefined);
          if (objects.length !== 1 || objects[0].visual.glow.radius !== 24
            || (objects[0].visual.glow.color & 255) !== 102
            || objects[0].visual.threeD?.cameraPreset !== "isometricLeftDown") {
            throw new Error("Sheet4 lost its authored 24px, 40% glow or 3D camera");
          }
          renderer = new SceneRenderer(objects, DEFAULT_LIMITS);
          // Isolate the real picture at a fixed scale, so grid/theme/DPR cannot affect pixel checks.
          const frame = await renderer.render({ type: "slide", index: 3, id: "glow", name: "Sheet4", width: 1100, height: 1000 }, { unitIndex: 3 });
          const canvas = new OffscreenCanvas(frame.bitmap.width, frame.bitmap.height);
          const context = canvas.getContext("2d");
          context.drawImage(frame.bitmap, 0, 0);
          frame.bitmap.close();
          const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data;
          const redHalo = (left, top, width, height) => {
            let total = 0;
            for (let y = top; y < top + height; y++) {
              for (let x = left; x < left + width; x++) {
                const offset = (y * canvas.width + x) * 4;
                total += Math.max(0, pixels[offset] - pixels[offset + 1]);
              }
            }
            return total / (width * height);
          };
          return { perimeter: redHalo(302, 200, 6, 50), arrow: redHalo(612, 380, 8, 40) };
        } finally {
          renderer?.close();
          document.close();
          core.close();
        }
      });
      // Before the fix: perimeter ≈5 (missing), arrow ≈142 (three accumulated shadows).
      // Linux WebKit paints the repaired perimeter at 19; require a visible halo across engines.
      if (glow.perimeter <= 10 || glow.arrow <= 10 || glow.arrow >= 60) {
        throw new Error(`${name} Sheet4 glow lost thin lines or over-saturated solids: ${JSON.stringify(glow)}`);
      }
      for (const extension of ["ods", "xlsx", "xls"]) {
        for (const theme of ["light", "dark"]) {
          await page.goto(`${url}examples/viewer.html?fixture=oasis-2173-tab-color.${extension}&theme=${theme}`);
          await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
          for (const selected of [0, 1]) {
            await page.locator(".sheet-tab").nth(selected).click();
            await page.waitForFunction((index) => document.querySelector("#viewer").shadowRoot
              .querySelectorAll(".sheet-tab")[index].getAttribute("aria-selected") === "true", selected);
            const tabs = await page.locator(".sheet-tab").evaluateAll(elements => elements.map(tab => ({
              name: tab.textContent,
              color: getComputedStyle(tab, "::after").backgroundColor,
              visible: tab.getBoundingClientRect().height > 0,
            })));
            if (tabs.length !== 2 || tabs[0].name !== "Bar" || tabs[1].name !== "Bubble"
              || tabs[0].color !== "rgb(255, 0, 0)" || tabs[1].color !== "rgb(0, 0, 255)"
              || tabs.some(tab => !tab.visible)) {
              throw new Error(`${name} ${extension} ${theme} lost worksheet tab colors: ${JSON.stringify(tabs)}`);
            }
          }
        }
      }
      await page.goto(`${url}examples/viewer.html?fixture=visual-baseline.pptx`, { waitUntil: "domcontentloaded" });
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true", null, { timeout: 30_000 });
      if (await page.locator(".open-button").isVisible()) throw new Error(`${name} showed the file picker by default`);
      const state = await page.locator("docviewkit-viewer").evaluate((viewer) => {
        const toolbarBounds = viewer.shadowRoot?.querySelector(".toolbar")?.getBoundingClientRect();
        const switcherBounds = viewer.shadowRoot?.querySelector(".interaction-switcher")?.getBoundingClientRect();
        return {
          canvasCount: viewer.shadowRoot?.querySelectorAll("canvas").length ?? 0,
          width: viewer.getBoundingClientRect().width,
          height: viewer.getBoundingClientRect().height,
          licenseStatus: viewer.state.license.status,
          brandingVisible: viewer.shadowRoot?.querySelector(".license-branding")?.hidden === false,
          overlayHidden: viewer.shadowRoot?.querySelector(".license-overlay") == null,
          diagnosticsHidden: viewer.shadowRoot?.querySelector('[data-action="diagnostics"]')?.hidden === true,
          interactionSwitcherVisible: viewer.shadowRoot?.querySelector(".interaction-switcher")?.hidden === false,
          interactionCenterOffset: toolbarBounds === undefined || switcherBounds === undefined
            ? Number.POSITIVE_INFINITY
            : Math.abs((switcherBounds.left + switcherBounds.right - toolbarBounds.left - toolbarBounds.right) / 2),
        };
      });
      if (state.canvasCount === 0 || state.width <= 0 || state.height <= 0) {
        throw new Error(`${name} rendered an empty viewer: ${JSON.stringify(state)}`);
      }
      if (state.licenseStatus !== "open-source"
        || state.brandingVisible
        || !state.overlayHidden
        || !state.diagnosticsHidden
        || !state.interactionSwitcherVisible
        || state.interactionCenterOffset > 1
      ) {
        throw new Error(`${name} did not expose unbranded open source capabilities: ${JSON.stringify(state)}`);
      }
      for (const width of [720, 640, 601, 600, 541, 540, 480, 375]) {
        await page.setViewportSize({ width, height: 900 });
        const layout = await page.locator("docviewkit-viewer").evaluate((viewer) => {
          const root = viewer.shadowRoot;
          const toolbar = root.querySelector(".toolbar").getBoundingClientRect();
          const switcher = root.querySelector(".interaction-switcher").getBoundingClientRect();
          const controls = [...root.querySelectorAll(".toolbar button, .toolbar .zoom-value")]
            .filter((button) => button.getBoundingClientRect().width > 0)
            .map((button) => button.getBoundingClientRect());
          return {
            zoomOutVisible: root.querySelector('[data-action="zoom-out"]').getBoundingClientRect().width > 0,
            zoomValueVisible: root.querySelector(".zoom-value").getBoundingClientRect().width > 0,
            switcherVisible: switcher.width > 0,
            centerOffset: Math.abs((switcher.left + switcher.right - toolbar.left - toolbar.right) / 2),
            rightGap: toolbar.right - controls.at(-1).right,
            overlap: controls.some((bounds, index) => index > 0 && bounds.left < controls[index - 1].right),
          };
        });
        if (layout.rightGap > 16 || layout.rightGap < 0 || layout.overlap
          || !layout.zoomOutVisible || !layout.zoomValueVisible
          || (width >= 541 && !layout.switcherVisible)
          || (width >= 601 && layout.centerOffset > 1)) {
          throw new Error(`${name} crowded narrow toolbar at ${width}px: ${JSON.stringify(layout)}`);
        }
      }
      await page.setViewportSize({ width: 1440, height: 900 });
      const minimalMode = await page.locator("docviewkit-viewer").evaluate((viewer) => {
        viewer.config = { ...viewer.config, minimal: true };
        const shell = viewer.shadowRoot?.querySelector(".shell");
        const workspace = viewer.shadowRoot?.querySelector(".workspace");
        if (!(shell instanceof HTMLElement) || !(workspace instanceof HTMLElement)) return undefined;
        const hidden = [".toolbar", ".formula-bar", ".navigation", ".navigation-resizer", ".sheet-tabs", ".statusbar", ".diagnostic-panel"]
          .every((selector) => getComputedStyle(viewer.shadowRoot.querySelector(selector)).display === "none");
        const shellBounds = shell.getBoundingClientRect();
        const workspaceBounds = workspace.getBoundingClientRect();
        viewer.config = { ...viewer.config, minimal: false };
        return {
          hidden,
          workspaceWidth: workspaceBounds.width,
          workspaceHeight: workspaceBounds.height,
          shellWidth: shellBounds.width,
          shellHeight: shellBounds.height,
        };
      });
      if (minimalMode === undefined || !minimalMode.hidden
        || Math.abs(minimalMode.workspaceWidth - minimalMode.shellWidth) > 1
        || Math.abs(minimalMode.workspaceHeight - minimalMode.shellHeight) > 1) {
        throw new Error(`${name} did not isolate the rendering workspace in minimal mode: ${JSON.stringify(minimalMode)}`);
      }
      await page.locator("docviewkit-viewer").evaluate((viewer) => {
        viewer.shadowRoot?.querySelector('[data-action="search"]')?.click();
        viewer.config = { ...viewer.config };
      });
      await page.setViewportSize({ width: 930, height: 900 });
      const compactSearchLayout = await page.locator("docviewkit-viewer").evaluate((viewer) => {
        const search = viewer.shadowRoot?.querySelector(".search");
        const switcher = viewer.shadowRoot?.querySelector(".interaction-switcher");
        const searchAction = viewer.shadowRoot?.querySelector('[data-action="search"]');
        if (!(search instanceof HTMLElement) || !(switcher instanceof HTMLElement)) return undefined;
        const searchBounds = search.getBoundingClientRect();
        const switcherBounds = switcher.getBoundingClientRect();
        return {
          searchRight: searchBounds.right,
          switcherLeft: switcherBounds.left,
          switcherVisibility: getComputedStyle(switcher).visibility,
          searchActionHidden: searchAction instanceof HTMLElement && searchAction.hidden,
          searchActionDisplay: searchAction instanceof HTMLElement ? getComputedStyle(searchAction).display : undefined,
        };
      });
      if (compactSearchLayout === undefined
        || compactSearchLayout.switcherVisibility !== "visible"
        || !compactSearchLayout.searchActionHidden
        || compactSearchLayout.searchActionDisplay !== "none"
        || compactSearchLayout.searchRight > compactSearchLayout.switcherLeft) {
        throw new Error(`${name} overlapped search and interaction controls: ${JSON.stringify(compactSearchLayout)}`);
      }
      await page.setViewportSize({ width: 640, height: 900 });
      const narrowSearchLayout = await page.locator("docviewkit-viewer").evaluate((viewer) => {
        const search = viewer.shadowRoot?.querySelector(".search");
        const switcher = viewer.shadowRoot?.querySelector(".interaction-switcher");
        return {
          searchVisible: search instanceof HTMLElement && getComputedStyle(search).display !== "none",
          searchWidth: search?.getBoundingClientRect().width,
          switcherDisplay: switcher instanceof HTMLElement ? getComputedStyle(switcher).display : undefined,
        };
      });
      if (!narrowSearchLayout.searchVisible || narrowSearchLayout.switcherDisplay !== "none"
        || narrowSearchLayout.searchWidth > 300) {
        throw new Error(`${name} did not prioritize search on a narrow toolbar: ${JSON.stringify(narrowSearchLayout)}`);
      }
      await page.setViewportSize({ width: 1440, height: 900 });
      const searchInput = page.locator("docviewkit-viewer .search-input");
      await searchInput.focus();
      const searchFocus = await searchInput.evaluate((input) => {
        const search = input.closest(".search");
        return {
          inputOutlineStyle: getComputedStyle(input).outlineStyle,
          searchBoxShadow: search instanceof HTMLElement ? getComputedStyle(search).boxShadow : "none",
        };
      });
      if (searchFocus.inputOutlineStyle !== "none" || searchFocus.searchBoxShadow === "none") {
        throw new Error(`${name} showed nested search focus rings: ${JSON.stringify(searchFocus)}`);
      }
      await searchInput.fill("视觉基线");
      await page.waitForFunction(
        () => document.querySelector("docviewkit-viewer")?.state.search.total === 1,
        null,
        { timeout: 30_000 },
      );
      const searchState = await page.locator("docviewkit-viewer").evaluate((viewer) => ({
        ...viewer.state.search,
        position: viewer.shadowRoot?.querySelector(".search-position")?.textContent,
        previousDisabled: viewer.shadowRoot?.querySelector('[data-action="previous-result"]')?.disabled,
        nextDisabled: viewer.shadowRoot?.querySelector('[data-action="next-result"]')?.disabled,
      }));
      if (searchState.query !== "视觉基线"
        || searchState.current !== 1
        || searchState.total !== 1
        || searchState.position !== "1 of 1"
        || searchState.previousDisabled
        || searchState.nextDisabled) {
        throw new Error(`${name} did not complete document search: ${JSON.stringify(searchState)}`);
      }
      await page.waitForFunction(
        () => document.querySelector("docviewkit-viewer")?.shadowRoot?.querySelector(".search-highlight-match") !== null,
        null,
        { timeout: 30_000 },
      );
      const searchHighlight = await page.locator("docviewkit-viewer").evaluate((viewer) => {
        const match = viewer.shadowRoot?.querySelector(".search-highlight-match");
        const fragment = match?.closest(".search-highlight-fragment");
        const selection = viewer.shadowRoot?.querySelector(".selection");
        const matchBounds = match?.getBoundingClientRect();
        const fragmentBounds = fragment?.getBoundingClientRect();
        return {
          text: match?.textContent,
          matchWidth: matchBounds?.width ?? 0,
          fragmentWidth: fragmentBounds?.width ?? 0,
          objectSelectionHidden: selection instanceof HTMLElement && selection.hidden,
        };
      });
      if (searchHighlight.text !== "视觉基线"
        || searchHighlight.matchWidth <= 0
        || searchHighlight.matchWidth > searchHighlight.fragmentWidth
        || !searchHighlight.objectSelectionHidden) {
        throw new Error(`${name} did not highlight the exact matching text: ${JSON.stringify(searchHighlight)}`);
      }
      await searchInput.fill("不存在的搜索词");
      const invalidatedSearch = await page.locator("docviewkit-viewer").evaluate((viewer) => ({
        ...viewer.state.search,
        position: viewer.shadowRoot?.querySelector(".search-position")?.textContent,
        previousDisabled: viewer.shadowRoot?.querySelector('[data-action="previous-result"]')?.disabled,
        nextDisabled: viewer.shadowRoot?.querySelector('[data-action="next-result"]')?.disabled,
      }));
      if (invalidatedSearch.query !== "不存在的搜索词"
        || invalidatedSearch.current !== 0
        || invalidatedSearch.total !== 0
        || invalidatedSearch.position !== "0 / 0"
        || !invalidatedSearch.previousDisabled
        || !invalidatedSearch.nextDisabled) {
        throw new Error(`${name} retained stale document search results: ${JSON.stringify(invalidatedSearch)}`);
      }
      const hiddenSearch = await page.locator("docviewkit-viewer").evaluate((viewer) => {
        viewer.config = {
          ...viewer.config,
          features: { ...viewer.config.features, search: false },
        };
        return {
          ...viewer.state.search,
          buttonHidden: viewer.shadowRoot?.querySelector('[data-action="search"]')?.hidden,
          formHidden: viewer.shadowRoot?.querySelector(".search")?.hidden,
        };
      });
      if (!hiddenSearch.buttonHidden
        || !hiddenSearch.formHidden
        || hiddenSearch.query !== ""
        || hiddenSearch.current !== 0
        || hiddenSearch.total !== 0) {
        throw new Error(`${name} did not hide document search: ${JSON.stringify(hiddenSearch)}`);
      }
      await page.locator("docviewkit-viewer").evaluate((viewer) => {
        viewer.config = {
          ...viewer.config,
          features: { ...viewer.config.features, search: true },
        };
      });
      const configuredDiagnosticsVisible = await page.locator("docviewkit-viewer").evaluate((viewer) => {
        viewer.config = {
          ...viewer.config,
          features: { ...viewer.config.features, diagnostics: true },
        };
        return viewer.shadowRoot?.querySelector('[data-action="diagnostics"]')?.hidden === false;
      });
      if (!configuredDiagnosticsVisible) {
        throw new Error(`${name} did not show diagnostics after explicit configuration`);
      }
      await page.locator("docviewkit-viewer").evaluate((viewer) => {
        viewer.config = {
          ...viewer.config,
          features: { ...viewer.config.features, interactionMode: "text" },
        };
      });
      await page.waitForFunction(
        () => Array.from(
          document.querySelector("docviewkit-viewer")?.shadowRoot?.querySelectorAll(".text-layer-item") ?? [],
        ).some((item) => (item.textContent ?? "").trim().length > 0),
        null,
        { timeout: 30_000 },
      );
      const textInteraction = await page.locator("docviewkit-viewer").evaluate((viewer) => {
        const item = Array.from(viewer.shadowRoot?.querySelectorAll(".text-layer-item") ?? [])
          .find((candidate) => candidate instanceof HTMLElement
            && (candidate.textContent ?? "").trim().length > 0);
        if (!(item instanceof HTMLElement)) return undefined;
        const range = document.createRange();
        range.selectNodeContents(item);
        const selection = viewer.shadowRoot?.getSelection?.() ?? document.getSelection();
        selection?.removeAllRanges();
        selection?.addRange(range);
        const style = getComputedStyle(item);
        const bounds = range.getBoundingClientRect();
        const result = {
          text: selection?.toString() || range.toString() || item.textContent || "",
          selectable: style.getPropertyValue("user-select") === "text"
            || style.getPropertyValue("-webkit-user-select") === "text",
          x: bounds.left + bounds.width / 2,
          y: bounds.top + bounds.height / 2,
        };
        selection?.removeAllRanges();
        return result;
      });
      if (textInteraction === undefined || textInteraction.text.length === 0 || !textInteraction.selectable) {
        throw new Error(`${name} did not expose selectable text in text interaction mode`);
      }
      const titleSelectionLayout = await page.locator("docviewkit-viewer").evaluate((viewer) => {
        const root = viewer.shadowRoot;
        const surface = root?.querySelector(".stage .surface");
        const item = Array.from(root?.querySelectorAll(".text-layer-item") ?? [])
          .find((candidate) => candidate.textContent === "OfficeViewer 视觉基线");
        const run = item?.querySelector("[data-font-run]");
        if (!(surface instanceof HTMLElement) || !(item instanceof HTMLElement) || !(run instanceof HTMLElement)) {
          return undefined;
        }
        const surfaceBounds = surface.getBoundingClientRect();
        const scale = surfaceBounds.width / 960;
        const range = document.createRange();
        range.selectNodeContents(item);
        const selectionBounds = range.getBoundingClientRect();
        const renderedBounds = [...item.querySelectorAll(".text-layer-fragment")]
          .map((fragment) => fragment.getBoundingClientRect());
        return {
          documentLeft: (selectionBounds.left - surfaceBounds.left) / scale,
          selectionWidth: selectionBounds.width / scale,
          renderedWidth: (Math.max(...renderedBounds.map((bounds) => bounds.right))
            - Math.min(...renderedBounds.map((bounds) => bounds.left))) / scale,
          fontWeight: getComputedStyle(run).fontWeight,
          layoutSource: item.dataset.layoutSource,
          fragmentCount: item.querySelectorAll(".text-layer-fragment").length,
        };
      });
      if (titleSelectionLayout === undefined
        || Math.abs(titleSelectionLayout.documentLeft - 69.6) > 0.1
        || Math.abs(titleSelectionLayout.selectionWidth - titleSelectionLayout.renderedWidth) > 1
        || Number(titleSelectionLayout.fontWeight) < 700
        || titleSelectionLayout.layoutSource !== "render"
        || titleSelectionLayout.fragmentCount < 1) {
        throw new Error(`${name} text selection ignored authored text-box layout: ${JSON.stringify(titleSelectionLayout)}`);
      }
      await page.locator("docviewkit-viewer").evaluate((viewer, point) => {
        viewer.dataset.objectSelectCount = "0";
        viewer.addEventListener("docviewkit-objectselect", () => {
          viewer.dataset.objectSelectCount = String(Number(viewer.dataset.objectSelectCount ?? 0) + 1);
        });
        viewer.config = {
          ...viewer.config,
          features: { ...viewer.config.features, interactionMode: "object" },
        };
        const target = viewer.shadowRoot?.elementFromPoint(point.x, point.y);
        target?.dispatchEvent(new MouseEvent("click", {
          bubbles: true,
          composed: true,
          clientX: point.x,
          clientY: point.y,
        }));
      }, textInteraction);
      await page.waitForFunction(
        () => document.querySelector("docviewkit-viewer")?.dataset.objectSelectCount === "1",
        null,
        { timeout: 30_000 },
      );
      const displayInteraction = await page.locator("docviewkit-viewer").evaluate(async (viewer, point) => {
        viewer.config = {
          ...viewer.config,
          features: { ...viewer.config.features, interactionMode: "display" },
        };
        const target = viewer.shadowRoot?.elementFromPoint(point.x, point.y);
        target?.dispatchEvent(new MouseEvent("click", {
          bubbles: true,
          composed: true,
          clientX: point.x,
          clientY: point.y,
        }));
        await new Promise((resolve) => setTimeout(resolve, 100));
        return {
          mode: viewer.shadowRoot?.querySelector(".shell")?.dataset.interactionMode,
          objectSelectCount: viewer.dataset.objectSelectCount,
          textItems: viewer.shadowRoot?.querySelectorAll(".text-layer-item").length ?? 0,
        };
      }, textInteraction);
      if (displayInteraction.mode !== "display"
        || displayInteraction.objectSelectCount !== "1"
        || displayInteraction.textItems !== 0) {
        throw new Error(`${name} did not keep display interaction mode inert: ${JSON.stringify(displayInteraction)}`);
      }
      const pinchFixture = process.env.DOCVIEWKIT_PINCH_FIXTURE;
      if (pinchFixture !== undefined) {
        await page.locator("#viewer-file-input").setInputFiles(pinchFixture);
        await page.waitForFunction(
          () => {
            const state = document.querySelector("docviewkit-viewer")?.state;
            return state?.info?.format === "docx" && state.status === "ready";
          },
          null,
          { timeout: 30_000 },
        );
      }
      const supportsContinuousPageMode = await page.locator("docviewkit-viewer").evaluate((viewer) => {
        const units = viewer.state.info?.units ?? [];
        return units.length > 1 && units.every((unit) => unit.type !== "sheet");
      });
      const pinchPageMode = process.env.DOCVIEWKIT_PINCH_PAGE_MODE
        ?? (supportsContinuousPageMode ? "continuous" : "single");
      if (pinchPageMode !== "continuous" && pinchPageMode !== "single") {
        throw new Error("DOCVIEWKIT_PINCH_PAGE_MODE must be continuous or single");
      }
      if (pinchPageMode === "continuous" && !supportsContinuousPageMode) {
        throw new Error("DOCVIEWKIT_PINCH_PAGE_MODE=continuous requires a multi-unit non-sheet fixture");
      }
      if (await page.locator("docviewkit-viewer").evaluate((viewer) => viewer.state.pageMode) !== pinchPageMode) {
        await page.locator("docviewkit-viewer").evaluate((viewer) => {
          viewer.shadowRoot?.querySelector('[data-action="page-mode"]')?.click();
        });
        await page.waitForFunction(
          (mode) => {
            const viewer = document.querySelector("docviewkit-viewer");
            if (viewer?.state.pageMode !== mode) return false;
            if (mode === "continuous") return true;
            const surface = viewer.shadowRoot?.querySelector(".stage .surface");
            return surface instanceof HTMLCanvasElement && surface.width > 1 && surface.height > 1;
          },
          pinchPageMode,
          { timeout: 30_000 },
        );
        await page.waitForTimeout(100);
      }
      const pinchUnitIndex = Number(process.env.DOCVIEWKIT_PINCH_UNIT_INDEX ?? 0);
      if (!Number.isInteger(pinchUnitIndex) || pinchUnitIndex < 0) {
        throw new Error("DOCVIEWKIT_PINCH_UNIT_INDEX must be a non-negative integer");
      }
      if (pinchUnitIndex !== 0) {
        await page.locator("docviewkit-viewer").evaluate(
          (viewer, unitIndex) => viewer.reveal({ kind: "unit", unitIndex }),
          pinchUnitIndex,
        );
        await page.waitForFunction(
          (unitIndex) => document.querySelector("docviewkit-viewer")?.state.currentUnitIndex === unitIndex,
          pinchUnitIndex,
          { timeout: 30_000 },
        );
        await page.waitForTimeout(100);
      }
      const pinchZoom = await page.locator("docviewkit-viewer").evaluate(async (
        viewer,
        { pinchDeltaY, focalX, focalY },
      ) => {
        const workspace = viewer.shadowRoot?.querySelector(".workspace");
        if (!(workspace instanceof HTMLElement)) return undefined;
        const currentCanvas = viewer.state.pageMode === "continuous"
          ? viewer.shadowRoot?.querySelector(
              `.continuous-page[data-unit-index="${viewer.state.currentUnitIndex}"] canvas`,
            )
          : viewer.shadowRoot?.querySelector(".stage .surface");
        if (!(currentCanvas instanceof HTMLCanvasElement)) return undefined;
        let canvasWidthChanges = 0;
        let canvasResets = 0;
        const observer = new MutationObserver((records) => {
          for (const record of records) {
            if (record.attributeName !== "width") continue;
            canvasWidthChanges += 1;
            if (currentCanvas.width <= 1) canvasResets += 1;
          }
        });
        observer.observe(currentCanvas, { attributes: true, attributeFilter: ["width"] });
        const inkSample = (canvas, bounds) => {
          const width = 320;
          const height = Math.max(1, Math.round(width * bounds.height / Math.max(1, bounds.width)));
          const sample = document.createElement("canvas");
          sample.width = width;
          sample.height = height;
          const context = sample.getContext("2d", { willReadFrequently: true });
          if (context === null) return undefined;
          context.drawImage(canvas, 0, 0, width, height);
          const pixels = context.getImageData(0, 0, width, height).data;
          const ink = new Uint8Array(width * height);
          for (let index = 0; index < ink.length; index += 1) {
            const offset = index * 4;
            const luminance = (pixels[offset] + pixels[offset + 1] + pixels[offset + 2]) / 3;
            ink[index] = luminance < 224 ? 1 : 0;
          }
          return { width, height, ink };
        };
        const inkTranslation = (preview, final) => {
          if (preview.width !== final.width || preview.height !== final.height) return undefined;
          let best = { dx: 0, dy: 0, mismatches: Number.POSITIVE_INFINITY };
          for (let dy = -4; dy <= 4; dy += 1) {
            for (let dx = -4; dx <= 4; dx += 1) {
              let mismatches = 0;
              for (let y = 4; y < preview.height - 4; y += 1) {
                for (let x = 4; x < preview.width - 4; x += 1) {
                  const shiftedX = x + dx;
                  const shiftedY = y + dy;
                  mismatches += preview.ink[y * preview.width + x]
                    ^ final.ink[shiftedY * final.width + shiftedX];
                }
              }
              if (mismatches < best.mismatches) best = { dx, dy, mismatches };
            }
          }
          return best;
        };
        const bounds = workspace.getBoundingClientRect();
        const initialZoom = viewer.state.zoom;
        const initialScroll = { left: workspace.scrollLeft, top: workspace.scrollTop };
        const anchoredUnitIndex = viewer.state.currentUnitIndex;
        const initialSurface = viewer.state.pageMode === "continuous"
          ? viewer.shadowRoot?.querySelector(
              `.continuous-page[data-unit-index="${anchoredUnitIndex}"]`,
            )
          : viewer.shadowRoot?.querySelector(".stage");
        const initialRect = initialSurface?.getBoundingClientRect();
        const focalClientY = bounds.top + bounds.height * focalY;
        const anchoredClientY = initialRect === undefined
          ? focalClientY
          : Math.min(initialRect.bottom, Math.max(initialRect.top, focalClientY));
        const anchoredDocumentY = initialRect === undefined
          ? 0
          : (anchoredClientY - initialRect.top) / initialZoom;
        for (let index = 0; index < 40; index += 1) {
          workspace.dispatchEvent(new WheelEvent("wheel", {
            bubbles: true,
            cancelable: true,
            ctrlKey: true,
            clientX: bounds.left + bounds.width * focalX,
            clientY: bounds.top + bounds.height * focalY,
            deltaY: pinchDeltaY,
          }));
          await new Promise((resolve) => setTimeout(resolve, 16));
        }
        await new Promise((resolve) => setTimeout(resolve, 50));
        const previewSurface = viewer.state.pageMode === "continuous"
          ? viewer.shadowRoot?.querySelector(
              `.continuous-page[data-unit-index="${anchoredUnitIndex}"]`,
            )
          : viewer.shadowRoot?.querySelector(".stage");
        const previewRect = previewSurface?.getBoundingClientRect();
        const previewScroll = { left: workspace.scrollLeft, top: workspace.scrollTop };
        const previewInk = previewRect === undefined ? undefined : inkSample(currentCanvas, previewRect);
        const highResolutionDeadline = performance.now() + 3_000;
        while (performance.now() < highResolutionDeadline) {
          const pageElement = viewer.state.pageMode === "continuous"
            ? viewer.shadowRoot?.querySelector(
                `.continuous-page[data-unit-index="${anchoredUnitIndex}"]`,
              )
            : viewer.shadowRoot?.querySelector(".stage");
          const canvas = pageElement?.querySelector("canvas");
          const pageRect = pageElement?.getBoundingClientRect();
          if (
            canvas instanceof HTMLCanvasElement
            && pageRect !== undefined
            && pageRect.width > 0
            && canvas.width / pageRect.width >= window.devicePixelRatio * .9
          ) {
            break;
          }
          await new Promise((resolve) => setTimeout(resolve, 50));
        }
        observer.disconnect();
        const currentPage = viewer.state.pageMode === "continuous"
          ? viewer.shadowRoot?.querySelector(
              `.continuous-page[data-unit-index="${anchoredUnitIndex}"]`,
            )
          : viewer.shadowRoot?.querySelector(".stage");
        const settledCanvas = currentPage?.querySelector("canvas");
        const finalRect = currentPage?.getBoundingClientRect();
        const finalScroll = { left: workspace.scrollLeft, top: workspace.scrollTop };
        const finalInk = settledCanvas instanceof HTMLCanvasElement && finalRect !== undefined
          ? inkSample(settledCanvas, finalRect)
          : undefined;
        const currentPageWidth = finalRect?.width ?? 0;
        const workspaceRect = workspace.getBoundingClientRect();
        const settleShift = previewRect === undefined || finalRect === undefined
          ? Number.POSITIVE_INFINITY
          : Math.max(
              Math.abs(previewRect.left - finalRect.left),
              Math.abs(previewRect.top - finalRect.top),
              Math.abs(previewRect.right - finalRect.right),
              Math.abs(previewRect.bottom - finalRect.bottom),
            );
        const settleScrollShift = Math.hypot(
          previewScroll.left - finalScroll.left,
          previewScroll.top - finalScroll.top,
        );
        const contentTranslation = previewInk === undefined || finalInk === undefined
          ? undefined
          : inkTranslation(previewInk, finalInk);
        const contentShift = contentTranslation === undefined || finalRect === undefined
          ? Number.POSITIVE_INFINITY
          : Math.hypot(
              contentTranslation.dx * finalRect.width / previewInk.width,
              contentTranslation.dy * finalRect.height / previewInk.height,
            );
        return {
          initialZoom,
          finalZoom: viewer.state.zoom,
          canvasWidthChanges,
          canvasResets,
          settleShift,
          settleScrollShift,
          verticalScrollChange: Math.abs(finalScroll.top - initialScroll.top),
          verticalPageShift: initialRect === undefined || finalRect === undefined
            ? Number.POSITIVE_INFINITY
            : Math.abs(finalRect.top - initialRect.top),
          pointerVerticalDrift: finalRect === undefined
            ? Number.POSITIVE_INFINITY
            : Math.abs(finalRect.top + anchoredDocumentY * viewer.state.zoom - anchoredClientY),
          horizontalCenterError: finalRect !== undefined && finalRect.width <= workspace.clientWidth
            ? Math.abs((finalRect.left + finalRect.right) / 2
              - (workspaceRect.left + workspaceRect.right) / 2)
            : 0,
          contentShift,
          renderPixelRatio: settledCanvas instanceof HTMLCanvasElement && currentPageWidth > 0
            ? settledCanvas.width / currentPageWidth
            : 0,
          previewTransformCount: Array.from(
            viewer.shadowRoot?.querySelectorAll(".stage, .continuous-page .text-layer") ?? [],
          ).filter((element) => element instanceof HTMLElement && element.style.transform !== "").length,
        };
      }, {
        pinchDeltaY: Number(process.env.DOCVIEWKIT_PINCH_DELTA_Y ?? -1.5),
        focalX: Number(process.env.DOCVIEWKIT_PINCH_FOCAL_X ?? .73),
        focalY: Number(process.env.DOCVIEWKIT_PINCH_FOCAL_Y ?? .31),
      });
      if (pinchZoom === undefined
        || Math.abs(pinchZoom.finalZoom - pinchZoom.initialZoom) < .01
        || pinchZoom.canvasWidthChanges > 4
        || pinchZoom.canvasResets !== 0
        || pinchZoom.settleShift > 1
        || pinchZoom.settleScrollShift > 1
        || (pinchPageMode === "single"
          && pinchZoom.finalZoom < pinchZoom.initialZoom
          && pinchZoom.verticalScrollChange > 1)
        || ((pinchPageMode === "continuous" || pinchZoom.finalZoom > pinchZoom.initialZoom)
          && pinchZoom.pointerVerticalDrift > 2)
        || pinchZoom.horizontalCenterError > 1
        || pinchZoom.contentShift > 1
        || pinchZoom.renderPixelRatio < deviceScaleFactor * .9
        || pinchZoom.previewTransformCount !== 0) {
        throw new Error(`${name} rerendered too often during one pinch gesture: ${JSON.stringify(pinchZoom)}`);
      }
      const selectionClick = pinchFixture === undefined
        ? undefined
        : await page.locator("docviewkit-viewer").evaluate(async (viewer) => {
            viewer.config = {
              ...viewer.config,
              features: { ...viewer.config.features, interactionMode: "object" },
            };
            const workspace = viewer.shadowRoot?.querySelector(".workspace");
            const canvas = viewer.state.pageMode === "continuous"
              ? viewer.shadowRoot?.querySelector(
                  `.continuous-page[data-unit-index="${viewer.state.currentUnitIndex}"] canvas`,
                )
              : viewer.shadowRoot?.querySelector(".stage .surface");
            if (!(workspace instanceof HTMLElement) || !(canvas instanceof HTMLCanvasElement)) return undefined;
            const rect = canvas.getBoundingClientRect();
            const before = { left: workspace.scrollLeft, top: workspace.scrollTop };
            let selected = false;
            for (const [xFactor, yFactor] of [[.25, .2], [.75, .2], [.25, .75], [.75, .75]]) {
              selected = await new Promise((resolve) => {
                const handler = () => {
                  clearTimeout(timer);
                  resolve(true);
                };
                const timer = setTimeout(() => {
                  viewer.removeEventListener("docviewkit-objectselect", handler);
                  resolve(false);
                }, 200);
                viewer.addEventListener("docviewkit-objectselect", handler, { once: true });
                canvas.dispatchEvent(new MouseEvent("click", {
                  bubbles: true,
                  composed: true,
                  clientX: rect.left + rect.width * xFactor,
                  clientY: rect.top + rect.height * yFactor,
                }));
              });
              if (selected) break;
            }
            await new Promise((resolve) => setTimeout(resolve, 300));
            return {
              selected,
              scrollShift: Math.hypot(
                workspace.scrollLeft - before.left,
                workspace.scrollTop - before.top,
              ),
            };
          });
      if (selectionClick !== undefined && (!selectionClick.selected || selectionClick.scrollShift > 1)) {
        throw new Error(`${name} moved the view after a direct object click: ${JSON.stringify(selectionClick)}`);
      }
      await page.setViewportSize({ width: 360, height: 640 });
      const mobile = await page.locator("docviewkit-viewer").evaluate((viewer) => {
        const statusbar = viewer.shadowRoot?.querySelector(".statusbar");
        const statusRect = statusbar?.getBoundingClientRect();
        return {
          statusWidth: statusbar?.clientWidth ?? 0,
          statusScrollWidth: statusbar?.scrollWidth ?? 0,

        };
      });
      if (mobile.statusWidth === 0 || mobile.statusScrollWidth > mobile.statusWidth) {
        throw new Error(`${name} overflowed the narrow status bar: ${JSON.stringify(mobile)}`);
      }
      await page.locator("docviewkit-viewer").evaluate((viewer) => {
        const transfer = new DataTransfer();
        transfer.items.add(new File(["Name,Value\nDropped,1"], "dropped.csv", { type: "text/csv" }));
        viewer.dispatchEvent(new DragEvent("drop", { bubbles: true, cancelable: true, dataTransfer: transfer }));
      });
      await page.waitForFunction(
        () => document.querySelector("docviewkit-viewer")?.state.info?.format === "csv",
        null,
        { timeout: 30_000 },
      );
      await page.goto(`${url}examples/viewer.html?fixture=visual-baseline.pptx&fileDrop=false&filePicker=true`, { waitUntil: "domcontentloaded" });
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true", null, { timeout: 30_000 });
      if (!await page.locator(".open-button").isVisible()) throw new Error(`${name} did not show the explicitly enabled file picker`);
      const disabledDropPrevented = await page.locator("docviewkit-viewer").evaluate((viewer) => {
        const transfer = new DataTransfer();
        transfer.items.add(new File(["Name,Value\nDropped,1"], "dropped.csv", { type: "text/csv" }));
        const event = new DragEvent("drop", { bubbles: true, cancelable: true, dataTransfer: transfer });
        viewer.dispatchEvent(event);
        return event.defaultPrevented;
      });
      await page.waitForTimeout(100);
      if (!disabledDropPrevented
        || await page.locator("docviewkit-viewer").evaluate((viewer) => viewer.state.info?.format) !== "pptx") {
        throw new Error(`${name} did not safely reject a dropped file after explicit disablement`);
      }
      await page.setViewportSize({ width: 1440, height: 900 });
      await page.goto(`${url}examples/viewer.html?fixture=word-first-page-text.docx`, { waitUntil: "domcontentloaded" });
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true", null, { timeout: 30_000 });
      await page.locator("docviewkit-viewer").evaluate((viewer) => {
        viewer.config = {
          ...viewer.config,
          features: { ...viewer.config.features, interactionMode: "text" },
        };
      });
      await page.waitForFunction(
        () => document.querySelector("docviewkit-viewer")?.shadowRoot
          ?.querySelector('[data-object-id="docx:paragraph:12:fragment:0"]') !== null,
        null,
        { timeout: 30_000 },
      );
      const textSelectionLayout = await page.locator("docviewkit-viewer").evaluate((viewer) => {
        const item = viewer.shadowRoot?.querySelector('[data-object-id="docx:paragraph:12:fragment:0"]');
        if (!(item instanceof HTMLElement)) return undefined;
        const range = document.createRange();
        range.selectNodeContents(item);
        const fragments = [...item.querySelectorAll(".text-layer-fragment")];
        return {
          text: range.toString(),
          layoutSource: item.dataset.layoutSource,
          lineCount: new Set(fragments.map((fragment) => fragment.dataset.textLine)).size,
          fragmentLineCount: Math.max(0, ...fragments.map((fragment) => {
            const fragmentRange = document.createRange();
            fragmentRange.selectNodeContents(fragment);
            return [...fragmentRange.getClientRects()].filter((bounds) => bounds.width > 0 && bounds.height > 0).length;
          })),
        };
      });
      if (textSelectionLayout === undefined
        || !textSelectionLayout.text.includes("2021年7月26日")
        || textSelectionLayout.layoutSource !== "render"
        // Fallback fonts may wrap the rendered paragraph; its individual fragments must not wrap again.
        || textSelectionLayout.lineCount < 1 || textSelectionLayout.fragmentLineCount !== 1) {
        throw new Error(`${name} text selection escaped the rendered line: ${JSON.stringify(textSelectionLayout)}`);
      }
      await page.locator("#viewer-file-input").setInputFiles(
        new URL("../tests/fixtures/word-table-auto-height.doc", import.meta.url).pathname,
      );
      await page.waitForFunction(
        () => document.querySelector("docviewkit-viewer")?.state.info?.format === "doc",
        null,
        { timeout: 60_000 },
      );
      await page.locator("docviewkit-viewer").evaluate((viewer) => {
        viewer.config = {
          ...viewer.config,
          features: { ...viewer.config.features, interactionMode: "text" },
        };
      });
      await page.waitForFunction(
        () => [...(document.querySelector("docviewkit-viewer")?.shadowRoot?.querySelectorAll(".text-layer-item") ?? [])]
          .some((item) => item.textContent === "三晋先锋隐私协议"),
        null,
        { timeout: 60_000 },
      );
      const legacyDocTitleSelection = await page.locator("docviewkit-viewer").evaluate((viewer) => {
        const item = [...(viewer.shadowRoot?.querySelectorAll(".text-layer-item") ?? [])]
          .find((candidate) => candidate.textContent === "三晋先锋隐私协议");
        const body = [...(viewer.shadowRoot?.querySelectorAll(".text-layer-item") ?? [])]
          .find((candidate) => candidate.textContent?.startsWith("您的信任对我们非常重要"));
        const surface = item?.closest(".continuous-page")?.querySelector("canvas")
          ?? viewer.shadowRoot?.querySelector(".stage .surface");
        if (!(item instanceof HTMLElement) || !(body instanceof HTMLElement)
          || !(surface instanceof HTMLCanvasElement)) return undefined;
        const surfaceBounds = surface.getBoundingClientRect();
        const range = document.createRange();
        range.selectNodeContents(item);
        const selectionBounds = range.getBoundingClientRect();
        const bodyFragments = [...body.querySelectorAll(".text-layer-fragment")];
        const context = surface.getContext("2d");
        const pixels = context?.getImageData(0, 0, surface.width, surface.height);
        const scaleX = surface.width / surfaceBounds.width;
        const scaleY = surface.height / surfaceBounds.height;
        const inkFragments = bodyFragments.filter((fragment) => (fragment.textContent ?? "").trim().length > 0);
        const selectionLines = new Map();
        for (const fragment of bodyFragments) {
          const range = document.createRange();
          range.selectNodeContents(fragment);
          const bounds = range.getBoundingClientRect();
          if (bounds.width <= 0 || bounds.height <= 0) continue;
          const line = Number(fragment.dataset.textLine);
          const entries = selectionLines.get(line) ?? [];
          // Range tops vary with fallback font ascent; the renderer owns line placement.
          entries.push({ left: bounds.left, right: bounds.right, top: fragment.getBoundingClientRect().top });
          selectionLines.set(line, entries);
        }
        const selectionContinuity = [...selectionLines.values()].map((entries) => {
          const ordered = entries.toSorted((left, right) => left.left - right.left);
          let right = ordered[0]?.right ?? 0;
          let maxGap = 0;
          for (const bounds of ordered.slice(1)) {
            maxGap = Math.max(maxGap, bounds.left - right);
            right = Math.max(right, bounds.right);
          }
          const tops = ordered.map(({ top }) => top);
          return {
            maxGap,
            topSpread: Math.max(...tops) - Math.min(...tops),
          };
        });
        const bodyInkCoverage = pixels === undefined || inkFragments.length === 0 ? 0
          : inkFragments.filter((fragment) => {
              const bounds = fragment.getBoundingClientRect();
              const left = Math.max(0, Math.floor((bounds.left - surfaceBounds.left) * scaleX));
              const right = Math.min(surface.width, Math.ceil((bounds.right - surfaceBounds.left) * scaleX));
              const top = Math.max(0, Math.floor((bounds.top - surfaceBounds.top) * scaleY));
              const bottom = Math.min(surface.height, Math.ceil((bounds.bottom - surfaceBounds.top) * scaleY));
              for (let y = top; y < bottom; y += 1) {
                for (let x = left; x < right; x += 1) {
                  const offset = (y * surface.width + x) * 4;
                  if ((pixels.data[offset] ?? 255) + (pixels.data[offset + 1] ?? 255)
                    + (pixels.data[offset + 2] ?? 255) < 660) return true;
                }
              }
              return false;
            }).length / inkFragments.length;
        return {
          centerOffset: Math.abs(
            (selectionBounds.left + selectionBounds.right - surfaceBounds.left - surfaceBounds.right) / 2,
          ),
          selectionWidth: selectionBounds.width,
          surfaceWidth: surfaceBounds.width,
          bodyLayoutSource: body.dataset.layoutSource,
          bodyFragmentCount: body.querySelectorAll(".text-layer-fragment").length,
          bodyLineCount: new Set(bodyFragments.map((fragment) => fragment.dataset.textLine)).size,
          bodySelectionMaxGap: Math.max(...selectionContinuity.map(({ maxGap }) => maxGap)),
          bodySelectionTopSpread: Math.max(...selectionContinuity.map(({ topSpread }) => topSpread)),
          bodyInkCoverage,
        };
      });
      if (legacyDocTitleSelection === undefined
        || legacyDocTitleSelection.centerOffset > 1
        || legacyDocTitleSelection.selectionWidth >= legacyDocTitleSelection.surfaceWidth * 0.75
        || legacyDocTitleSelection.bodyLayoutSource !== "render"
        || legacyDocTitleSelection.bodyFragmentCount < 3
        || legacyDocTitleSelection.bodyLineCount < 3
        || legacyDocTitleSelection.bodySelectionMaxGap > 1
        || legacyDocTitleSelection.bodySelectionTopSpread > 1
        || legacyDocTitleSelection.bodyInkCoverage < 0.8) {
        throw new Error(`${name} legacy DOC selection ignored paragraph layout: ${JSON.stringify(legacyDocTitleSelection)}`);
      }
      const titleSelected = await page.locator("docviewkit-viewer").evaluate((viewer) => {
        const item = [...(viewer.shadowRoot?.querySelectorAll(".text-layer-item") ?? [])]
          .find((candidate) => candidate.textContent === "三晋先锋隐私协议");
        if (!(item instanceof HTMLElement)) return false;
        const range = document.createRange();
        range.selectNodeContents(item);
        const selection = viewer.shadowRoot?.getSelection?.() ?? document.getSelection();
        selection?.removeAllRanges();
        selection?.addRange(range);
        viewer.shadowRoot?.dispatchEvent(new Event("selectionchange"));
        return true;
      });
      if (!titleSelected) throw new Error(`${name} could not select the real DOC title`);
      await page.evaluate(async () => {
        await document.fonts.ready;
        await Promise.allSettled(document.getAnimations().map((animation) => animation.finished));
        await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      });
      const titleSelectionHeight = await page.locator("docviewkit-viewer").evaluate((viewer) => {
        const item = [...(viewer.shadowRoot?.querySelectorAll(".text-layer-item") ?? [])]
          .find((candidate) => candidate.textContent === "三晋先锋隐私协议");
        const fragment = item?.querySelector(".text-layer-fragment");
        const surface = item?.closest(".continuous-page")?.querySelector("canvas");
        if (!(fragment instanceof HTMLElement) || !(surface instanceof HTMLCanvasElement)) return undefined;
        const fragmentRange = document.createRange();
        fragmentRange.selectNodeContents(fragment);
        // Exclude paragraph leading so the pixel probe cannot capture the next line.
        const fragmentBounds = fragmentRange.getBoundingClientRect();
        const highlightBounds = [...(viewer.shadowRoot?.querySelectorAll(".text-selection-highlight") ?? [])]
          .map((highlight) => highlight.getBoundingClientRect())
          .find((bounds) => bounds.right > fragmentBounds.left && bounds.left < fragmentBounds.right);
        const selectionBounds = highlightBounds ?? fragmentRange.getBoundingClientRect();
        const surfaceBounds = surface.getBoundingClientRect();
        const context = surface.getContext("2d");
        const pixels = context?.getImageData(0, 0, surface.width, surface.height);
        if (selectionBounds.width <= 0 || selectionBounds.height <= 0 || pixels === undefined) return undefined;
        const scaleX = surface.width / surfaceBounds.width;
        const scaleY = surface.height / surfaceBounds.height;
        const left = Math.max(0, Math.floor((fragmentBounds.left - surfaceBounds.left) * scaleX));
        const right = Math.min(surface.width, Math.ceil((fragmentBounds.right - surfaceBounds.left) * scaleX));
        const top = Math.max(0,
          Math.floor((fragmentBounds.top - fragmentBounds.height * 0.25 - surfaceBounds.top) * scaleY));
        const bottom = Math.min(surface.height,
          Math.ceil((fragmentBounds.bottom + fragmentBounds.height * 0.25 - surfaceBounds.top) * scaleY));
        let inkTop = bottom;
        let inkBottom = top;
        for (let y = top; y < bottom; y += 1) {
          for (let x = left; x < right; x += 1) {
            const offset = (y * surface.width + x) * 4;
            if ((pixels.data[offset] ?? 255) + (pixels.data[offset + 1] ?? 255)
              + (pixels.data[offset + 2] ?? 255) >= 660) continue;
            inkTop = Math.min(inkTop, y);
            inkBottom = Math.max(inkBottom, y + 1);
          }
        }
        if (inkBottom <= inkTop) return undefined;
        const inkBounds = {
          top: surfaceBounds.top + inkTop / scaleY,
          bottom: surfaceBounds.top + inkBottom / scaleY,
        };
        return {
          fragmentTop: fragmentBounds.top,
          fragmentBottom: fragmentBounds.bottom,
          highlightTop: selectionBounds.top,
          highlightBottom: selectionBounds.bottom,
          inkTop: inkBounds.top,
          inkBottom: inkBounds.bottom,
          heightRatio: selectionBounds.height / (inkBounds.bottom - inkBounds.top),
          centerOffset: Math.abs(
            (selectionBounds.top + selectionBounds.bottom - inkBounds.top - inkBounds.bottom) / 2,
          ),
        };
      });
      if (titleSelectionHeight === undefined
        || titleSelectionHeight.highlightTop > titleSelectionHeight.inkTop + 1
        || titleSelectionHeight.highlightBottom < titleSelectionHeight.inkBottom - 1
        // Font ascent/descent need not be symmetric; bound padding and contain all ink.
        || titleSelectionHeight.heightRatio > 1.5) {
        throw new Error(`${name} legacy DOC title selection height missed the rendered glyphs: ${JSON.stringify(titleSelectionHeight)}`);
      }
      await page.locator("docviewkit-viewer").evaluate(async (viewer) => {
        const body = [...(viewer.shadowRoot?.querySelectorAll(".text-layer-item") ?? [])]
          .find((candidate) => candidate.textContent?.startsWith("您的信任对我们非常重要"));
        body?.querySelector(".text-layer-fragment")?.scrollIntoView({ block: "center", inline: "center" });
        await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      });
      const dragSelection = await page.locator("docviewkit-viewer").evaluate((viewer) => {
        const body = [...(viewer.shadowRoot?.querySelectorAll(".text-layer-item") ?? [])]
          .find((candidate) => candidate.textContent?.startsWith("您的信任对我们非常重要"));
        const fragments = body === undefined ? [] : [...body.querySelectorAll(".text-layer-fragment")]
          .filter((fragment) => (fragment.textContent ?? "").trim().length > 0);
        const firstText = fragments[0]?.firstChild;
        const startOffset = firstText?.textContent?.indexOf("我们非常重要") ?? -1;
        if (firstText?.nodeType !== Node.TEXT_NODE || startOffset < 0) return undefined;
        const startGlyph = document.createRange();
        startGlyph.setStart(firstText, startOffset);
        startGlyph.setEnd(firstText, startOffset + 1);
        const first = startGlyph.getBoundingClientRect();
        const thirdLine = fragments.find((fragment) => Number(fragment.dataset.textLine) >= 2)?.getBoundingClientRect();
        return first === undefined || thirdLine === undefined ? undefined : {
          // Stay before the glyph midpoint, where subpixel rounding can select the next character.
          start: { x: first.left + first.width / 4, y: first.top + first.height / 2 },
          end: { x: thirdLine.right - 1, y: thirdLine.top + thirdLine.height / 2 },
        };
      });
      if (dragSelection === undefined) throw new Error(`${name} could not locate real DOC drag-selection endpoints`);
      await page.mouse.move(dragSelection.start.x, dragSelection.start.y);
      await page.mouse.down();
      await page.mouse.move(dragSelection.end.x, dragSelection.end.y, { steps: 8 });
      await page.mouse.up();
      const draggedSelection = await page.locator("docviewkit-viewer").evaluate((viewer, points) => ({
        text: viewer.shadowRoot?.getSelection?.()?.toString() ?? document.getSelection()?.toString() ?? "",
        startTarget: viewer.shadowRoot?.elementFromPoint(points.start.x, points.start.y)?.tagName
          + "." + viewer.shadowRoot?.elementFromPoint(points.start.x, points.start.y)?.className,
        endTarget: viewer.shadowRoot?.elementFromPoint(points.end.x, points.end.y)?.tagName
          + "." + viewer.shadowRoot?.elementFromPoint(points.end.x, points.end.y)?.className,
      }), dragSelection);
      if (!draggedSelection.text.startsWith("我们非常重要")
        || !draggedSelection.text.includes("保护您的个人信息安全可控")) {
        throw new Error(`${name} real DOC pointer drag did not produce deterministic text: ${JSON.stringify(draggedSelection)}`);
      }
      const expectedCopiedSelection = "我们非常重要，我们深知个人信息对您的重要性，我们将按法律法规要求，采取相应安全保护措施，尽力保护您的个人信息安全可控";
      const syntheticCopiedText = await page.locator("docviewkit-viewer").evaluate((viewer) => {
        const body = [...(viewer.shadowRoot?.querySelectorAll(".text-layer-item") ?? [])]
          .find((candidate) => candidate.textContent?.startsWith("您的信任对我们非常重要"));
        const target = body?.querySelector(".text-layer-fragment");
        if (target === null || target === undefined) return undefined;
        const event = new ClipboardEvent("copy", {
          bubbles: true,
          composed: true,
          cancelable: true,
          clipboardData: new DataTransfer(),
        });
        target.dispatchEvent(event);
        return event.clipboardData.getData("text/plain");
      });
      if (syntheticCopiedText !== expectedCopiedSelection) {
        throw new Error(`${name} real DOC copy did not join wrapped text fragments: ${JSON.stringify(syntheticCopiedText)}`);
      }
      if (name === "chromium") {
        await page.context().grantPermissions(["clipboard-read", "clipboard-write"], { origin: new URL(url).origin });
        await page.keyboard.press("ControlOrMeta+C");
        const copiedText = await page.evaluate(() => navigator.clipboard.readText());
        if (copiedText !== expectedCopiedSelection) {
          throw new Error(`${name} real DOC clipboard did not join wrapped text fragments: ${JSON.stringify(copiedText)}`);
        }
      }
      await page.waitForTimeout(50);
      const selectionHighlightCount = await page.locator("docviewkit-viewer").evaluate((viewer) =>
        viewer.shadowRoot?.querySelectorAll(".text-selection-highlight").length ?? 0);
      const selectedLineBands = await page.locator("docviewkit-viewer").evaluate((viewer) => {
        const selection = viewer.shadowRoot?.getSelection?.() ?? document.getSelection();
        if (selection === null || selection.rangeCount === 0) return [];
        const lines = [];
        for (const bounds of selection.getRangeAt(0).getClientRects()) {
          if (bounds.width <= 0 || bounds.height <= 0) continue;
          const middle = (bounds.top + bounds.bottom) / 2;
          const line = lines.find((candidate) => middle >= candidate.top && middle <= candidate.bottom);
          if (line === undefined) {
            lines.push({ left: bounds.left, right: bounds.right, top: bounds.top, bottom: bounds.bottom });
          } else {
            line.left = Math.min(line.left, bounds.left);
            line.right = Math.max(line.right, bounds.right);
            line.top = Math.min(line.top, bounds.top);
            line.bottom = Math.max(line.bottom, bounds.bottom);
          }
        }
        return lines.length > 0 ? lines
          : [...(viewer.shadowRoot?.querySelectorAll(".text-selection-highlight") ?? [])].map((highlight) => {
              const bounds = highlight.getBoundingClientRect();
              return { left: bounds.left, right: bounds.right, top: bounds.top, bottom: bounds.bottom };
            });
      });
      const selectionScreenshot = await page.screenshot({ type: "png" });
      const selectionPaintCoverage = await page.evaluate(async ({ png, bands }) => {
        const image = new Image();
        image.src = `data:image/png;base64,${png}`;
        await image.decode();
        const canvas = document.createElement("canvas");
        canvas.width = image.naturalWidth;
        canvas.height = image.naturalHeight;
        const context = canvas.getContext("2d");
        context?.drawImage(image, 0, 0);
        const pixels = context?.getImageData(0, 0, canvas.width, canvas.height);
        if (pixels === undefined || bands.length === 0) return 0;
        const scaleX = canvas.width / innerWidth;
        const scaleY = canvas.height / innerHeight;
        return Math.min(...bands.map((band) => {
          const left = Math.max(0, Math.floor(band.left * scaleX));
          const right = Math.min(canvas.width, Math.ceil(band.right * scaleX));
          const top = Math.max(0, Math.floor(band.top * scaleY));
          const bottom = Math.min(canvas.height, Math.ceil(band.bottom * scaleY));
          let paintedPixels = 0;
          for (let x = left; x < right; x += 1) {
            for (let y = top; y < bottom; y += 1) {
              const offset = (y * canvas.width + x) * 4;
              const red = pixels.data[offset] ?? 255;
              const green = pixels.data[offset + 1] ?? 255;
              const blue = pixels.data[offset + 2] ?? 255;
              if (blue - red > 12 && blue - green > 4 && red > 140) paintedPixels += 1;
            }
          }
          return paintedPixels / Math.max(1, (right - left) * (bottom - top));
        }));
      }, { png: selectionScreenshot.toString("base64"), bands: selectedLineBands });
      if (selectionPaintCoverage < 0.4 || selectionHighlightCount < 1) {
        throw new Error(`${name} legacy DOC selection paint was fragmented: ${JSON.stringify({
          selectionPaintCoverage,
          selectionHighlightCount,
        })}`);
      }
      const zoomSelectionDrag = await page.locator("docviewkit-viewer").evaluate((viewer) => {
        const item = [...(viewer.shadowRoot?.querySelectorAll(".text-layer-item") ?? [])]
          .find((candidate) => candidate.textContent?.startsWith("生效时间"));
        const fragment = item?.querySelector(".text-layer-fragment");
        const range = document.createRange();
        if (fragment !== null && fragment !== undefined) range.selectNodeContents(fragment);
        const bounds = fragment === null || fragment === undefined ? undefined : range.getBoundingClientRect();
        return bounds === undefined ? undefined : {
          start: { x: bounds.left + bounds.width * 0.4, y: bounds.top + bounds.height / 2 },
          end: { x: bounds.left + bounds.width * 0.9, y: bounds.top + bounds.height / 2 },
        };
      });
      if (zoomSelectionDrag === undefined) throw new Error(`${name} could not locate DOC zoom-selection text`);
      await page.mouse.move(zoomSelectionDrag.start.x, zoomSelectionDrag.start.y);
      await page.mouse.down();
      await page.mouse.move(zoomSelectionDrag.end.x, zoomSelectionDrag.end.y, { steps: 8 });
      await page.mouse.up();
      await page.waitForTimeout(50);
      const preZoomTextSelection = await page.locator("docviewkit-viewer").evaluate((viewer) => ({
        text: (viewer.shadowRoot?.getSelection?.() ?? document.getSelection())?.toString() ?? "",
        highlightCount: viewer.shadowRoot?.querySelectorAll(".text-selection-highlight").length ?? 0,
      }));
      if (preZoomTextSelection.text.length === 0 || preZoomTextSelection.highlightCount === 0) {
        throw new Error(`${name} could not establish the DOC selection before zoom: ${JSON.stringify(preZoomTextSelection)}`);
      }
      const selectedZoom = await page.locator("docviewkit-viewer").evaluate((viewer) => {
        const item = [...(viewer.shadowRoot?.querySelectorAll(".text-layer-item") ?? [])]
          .find((candidate) => candidate.textContent?.startsWith("生效时间"));
        const width = item?.querySelector(".text-layer-fragment")?.getBoundingClientRect().width ?? 0;
        const zoom = viewer.state.zoom;
        viewer.shadowRoot?.querySelector('[data-action="zoom-in"]')?.click();
        return {
          zoom,
          width,
          selectionText: (viewer.shadowRoot?.getSelection?.() ?? document.getSelection())?.toString() ?? "",
        };
      });
      if (selectedZoom.selectionText.length !== 0) {
        throw new Error(`${name} zoom kept a selection attached to stale text-layer nodes: ${JSON.stringify(selectedZoom)}`);
      }
      await page.waitForFunction(
        ({ zoom, width }) => {
          const viewer = document.querySelector("docviewkit-viewer");
          const item = [...(viewer?.shadowRoot?.querySelectorAll(".text-layer-item") ?? [])]
            .find((candidate) => candidate.textContent?.startsWith("生效时间"));
          const fragment = item?.querySelector(".text-layer-fragment");
          return (viewer?.state.zoom ?? 0) > zoom
            && (fragment?.getBoundingClientRect().width ?? 0) > width;
        },
        selectedZoom,
        { timeout: 30_000 },
      );
      await page.waitForTimeout(100);
      const zoomedTextSelection = await page.locator("docviewkit-viewer").evaluate((viewer) => {
        const item = [...(viewer.shadowRoot?.querySelectorAll(".text-layer-item") ?? [])]
          .find((candidate) => candidate.textContent?.startsWith("生效时间"));
        const fragments = [...(item?.querySelectorAll(".text-layer-fragment") ?? [])]
          .map((fragment) => fragment.getBoundingClientRect());
        const highlights = [...(viewer.shadowRoot?.querySelectorAll(".text-selection-highlight") ?? [])]
          .map((highlight) => highlight.getBoundingClientRect());
        return {
          highlightCount: highlights.length,
          orphanedHighlightCount: highlights.filter((highlight) => !fragments.some((fragment) =>
            highlight.right > fragment.left && highlight.left < fragment.right
            && highlight.bottom > fragment.top && highlight.top < fragment.bottom)).length,
        };
      });
      if (zoomedTextSelection.orphanedHighlightCount !== 0) {
        throw new Error(`${name} zoom retained stale text-selection rectangles: ${JSON.stringify(zoomedTextSelection)}`);
      }
      await page.locator("#viewer-file-input").setInputFiles(
        new URL("../tests/fixtures/frozen-panes.xlsx", import.meta.url).pathname,
      );
      await page.waitForFunction(
        () => document.querySelector("docviewkit-viewer")?.state.info?.format === "xlsx"
          && document.querySelector("docviewkit-viewer")?.state.status === "ready",
        null,
        { timeout: 60_000 },
      );
      await page.locator("docviewkit-viewer").evaluate((viewer) => {
        viewer.dataset.objectSelectCount = "0";
        viewer.addEventListener("docviewkit-objectselect", () => {
          viewer.dataset.objectSelectCount = String(Number(viewer.dataset.objectSelectCount ?? "0") + 1);
        }, { once: true });
        viewer.config = {
          ...viewer.config,
          features: { ...viewer.config.features, interactionMode: "text" },
        };
      });
      await page.waitForFunction(
        () => [...(document.querySelector("docviewkit-viewer")?.shadowRoot?.querySelectorAll(".text-layer-item") ?? [])]
          .some((item) => item.textContent === "区域 001"),
        null,
        { timeout: 30_000 },
      );
      const sheetCellPoint = await page.locator("docviewkit-viewer").evaluate((viewer) => {
        const item = [...(viewer.shadowRoot?.querySelectorAll(".text-layer-item") ?? [])]
          .find((candidate) => candidate.textContent === "区域 001");
        const bounds = item?.querySelector(".text-layer-fragment")?.getBoundingClientRect();
        return bounds === undefined ? undefined : {
          x: bounds.left + bounds.width / 2,
          y: bounds.top + bounds.height / 2,
        };
      });
      if (sheetCellPoint === undefined) throw new Error(`${name} could not locate XLSX cell A3 in text mode`);
      await page.mouse.click(sheetCellPoint.x, sheetCellPoint.y);
      await page.waitForFunction(() => document.querySelector("docviewkit-viewer")
        ?.shadowRoot?.querySelector(".cell-address")?.value === "A3");
      const textModeFormulaBar = await page.locator("docviewkit-viewer").evaluate((viewer) => ({
        address: viewer.shadowRoot?.querySelector(".cell-address")?.value,
        value: viewer.shadowRoot?.querySelector(".formula-value")?.value,
        objectSelectCount: Number(viewer.dataset.objectSelectCount ?? "0"),
      }));
      if (textModeFormulaBar.address !== "A3"
        || textModeFormulaBar.value !== "区域 001"
        || textModeFormulaBar.objectSelectCount !== 0) {
        throw new Error(`${name} text-mode spreadsheet click did not update the formula bar: ${JSON.stringify(textModeFormulaBar)}`);
      }
      const textSelectionFixture = process.env.DOCVIEWKIT_TEXT_SELECTION_FIXTURE;
      if (textSelectionFixture !== undefined) {
        await page.locator("#viewer-file-input").setInputFiles(textSelectionFixture);
        await page.waitForFunction(
          () => document.querySelector("docviewkit-viewer")?.state.status === "ready",
          null,
          { timeout: 60_000 },
        );
        await page.locator("docviewkit-viewer").evaluate((viewer) => {
          viewer.config = {
            ...viewer.config,
            features: { ...viewer.config.features, interactionMode: "text" },
          };
        });
        await page.waitForFunction(
          () => (document.querySelector("docviewkit-viewer")?.shadowRoot
            ?.querySelectorAll(".text-layer-item").length ?? 0) > 0,
          null,
          { timeout: 60_000 },
        );
        const suppliedSelectionLayout = await page.locator("docviewkit-viewer").evaluate((viewer) => {
          const items = Array.from(viewer.shadowRoot?.querySelectorAll(".text-layer-item") ?? [])
            .filter((item) => item instanceof HTMLElement);
          const item = items.find((candidate) => candidate.textContent?.includes("提示条款"));
          const surface = item?.closest(".continuous-page")?.querySelector("canvas")
            ?? viewer.shadowRoot?.querySelector(".stage .surface");
          const fragments = item === undefined ? [] : [...item.querySelectorAll(".text-layer-fragment")]
            .filter((fragment) => (fragment.textContent ?? "").trim().length > 0);
          const surfaceBounds = surface?.getBoundingClientRect();
          const context = surface instanceof HTMLCanvasElement ? surface.getContext("2d") : null;
          const pixels = context?.getImageData(0, 0, surface.width, surface.height);
          const inkCoverage = pixels === undefined || surfaceBounds === undefined || fragments.length === 0 ? 0
            : fragments.filter((fragment) => {
                const bounds = fragment.getBoundingClientRect();
                const scaleX = surface.width / surfaceBounds.width;
                const scaleY = surface.height / surfaceBounds.height;
                const left = Math.max(0, Math.floor((bounds.left - surfaceBounds.left) * scaleX));
                const right = Math.min(surface.width, Math.ceil((bounds.right - surfaceBounds.left) * scaleX));
                const top = Math.max(0, Math.floor((bounds.top - surfaceBounds.top) * scaleY));
                const bottom = Math.min(surface.height, Math.ceil((bounds.bottom - surfaceBounds.top) * scaleY));
                for (let y = top; y < bottom; y += 1) {
                  for (let x = left; x < right; x += 1) {
                    const offset = (y * surface.width + x) * 4;
                    if ((pixels.data[offset] ?? 255) + (pixels.data[offset + 1] ?? 255)
                      + (pixels.data[offset + 2] ?? 255) < 660) return true;
                  }
                }
                return false;
              }).length / fragments.length;
          return {
            format: viewer.state.info?.format,
            text: items.map((item) => item.textContent ?? "").join("\n"),
            layoutSource: item?.dataset.layoutSource,
            fragmentCount: fragments.length,
            inkCoverage,
          };
        });
        if (suppliedSelectionLayout.format !== "pages"
          || !suppliedSelectionLayout.text.includes("提示条款")
          || suppliedSelectionLayout.layoutSource !== "render"
          || suppliedSelectionLayout.fragmentCount < 1
          || suppliedSelectionLayout.inkCoverage < 0.8) {
          throw new Error(`${name} supplied text selection escaped its rendered page: ${JSON.stringify(suppliedSelectionLayout)}`);
        }
      }
      // Real OFD text selection: glyph positions, reverse/cross-line drags and zoom.
      await page.goto(`${url}examples/viewer.html?fixture=ofdrw-nalaizhuyi.ofd`, { waitUntil: "domcontentloaded" });
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
      await page.locator("docviewkit-viewer").evaluate((viewer) => {
        viewer.config = { ...viewer.config, features: { ...viewer.config.features, interactionMode: "text" } };
      });
      for (const zoomed of [false, true]) {
        if (zoomed) {
          await page.locator("docviewkit-viewer").evaluate((viewer) => {
            viewer.shadowRoot.querySelectorAll(".text-layer-fragment").forEach((node) => { node.dataset.beforeZoom = "true"; });
            viewer.shadowRoot.querySelector('[data-action="zoom-in"]').click();
          });
          await page.waitForFunction(() => !document.querySelector("docviewkit-viewer")
            .shadowRoot.querySelector(".text-layer-fragment[data-before-zoom]"));
        }
        await page.waitForFunction(() => {
          const viewer = document.querySelector("docviewkit-viewer");
          return viewer.shadowRoot.querySelectorAll(".text-layer-fragment").length > 100
            && !viewer.shadowRoot.querySelector(".text-layer").style.transform;
        });
        for (const [start, end, reverse] of [[0, 5, false], [0, 5, true], [0, 35, false]]) {
          const drag = await page.locator("docviewkit-viewer").evaluate((viewer, { start, end }) => {
            // Each case starts a new selection, not a native drag of the previous selection.
            (viewer.shadowRoot.getSelection?.() ?? document.getSelection())?.removeAllRanges();
            const item = [...viewer.shadowRoot.querySelectorAll(".text-layer-item")]
              .find((node) => node.textContent.startsWith("中国"));
            const all = [...viewer.shadowRoot.querySelectorAll(".text-layer-fragment")];
            const fragments = all.slice(all.indexOf(item.querySelector(".text-layer-fragment")));
            const point = (index, ratio) => {
              const bounds = fragments[index].getBoundingClientRect();
              return { x: bounds.left + bounds.width * ratio, y: bounds.top + bounds.height / 2 };
            };
            return {
              anchor: point(start, 0.1), focus: point(end, 0.9),
              expected: fragments.slice(start, end + 1).map((node) => node.textContent).join(""),
            };
          }, { start, end });
          const [anchor, focus] = reverse ? [drag.focus, drag.anchor] : [drag.anchor, drag.focus];
          await page.mouse.move(anchor.x, anchor.y);
          await page.mouse.down();
          await page.mouse.move(focus.x, focus.y, { steps: 20 });
          await page.mouse.up();
          await page.waitForFunction(() => document.querySelector("docviewkit-viewer")
            .shadowRoot.querySelectorAll(".text-selection-highlight").length > 0);
          const copied = await page.locator("docviewkit-viewer").evaluate((viewer) => {
            const event = new ClipboardEvent("copy", {
              bubbles: true, cancelable: true, composed: true, clipboardData: new DataTransfer(),
            });
            viewer.shadowRoot.dispatchEvent(event);
            return event.clipboardData.getData("text/plain");
          });
          if (copied.replace(/\n/gu, "") !== drag.expected) {
            throw new Error(`${name} OFD drag/copy mismatch (${zoomed}, ${reverse}): ${JSON.stringify({ copied, expected: drag.expected })}`);
          }
        }
      }
      const odtSpacing = await page.evaluate(async () => {
        const { createOfficeEngine } = await import("/dist/engine.js");
        const engine = await createOfficeEngine({ execution: "inline", wasm: "/dist/office-viewer-odf.wasm" });
        const document = await engine.open(await (await fetch("/tests/fixtures/oasis-contextual-spacing-section.odt")).arrayBuffer());
        try {
          const objects = (await document.listObjects()).filter((object) => object.type === "paragraph");
          const frame = await document.render({ unitIndex: 0, includeTextFragments: true });
          try {
            const paragraphs = objects.map((object) => ({
              bounds: object.bounds,
              lines: [...new Set(frame.textFragments.filter((fragment) => fragment.objectId === object.id).map((fragment) => fragment.transform.f))],
            }));
            const lineHeight = paragraphs[0].lines[1] - paragraphs[0].lines[0];
            const close = (a, b) => Math.abs(a - b) < 0.1;
            if (document.info.units.length !== 1 || paragraphs.length !== 6) throw new Error("ODT pagination or paragraphs changed");
            if (!document.diagnostics().some(({ code }) => code === "FONT_METRICS_APPLIED")) throw new Error("ODT font relayout was skipped");
            if (!close(paragraphs[0].bounds.y, 20 * 96 / 25.4)) throw new Error("ODT page-top margin was added twice");
            for (const p of paragraphs.slice(0, 5)) {
              if (!close(p.bounds.height, p.lines.length * lineHeight)) throw new Error("ODT reserved height differs from painted lines");
              if (!close(p.lines[0], p.bounds.y)) throw new Error("ODT acquired text-box padding");
            }
            for (let index = 1; index < 5; index += 1) {
              const gap = index === 2 || index === 4 ? 10 * 96 / 25.4 : 0;
              if (!close(paragraphs[index].lines[0] - paragraphs[index - 1].lines.at(-1), lineHeight + gap)) {
                throw new Error("ODT visible paragraph or section spacing is incorrect");
              }
            }
            return { pages: document.info.units.length, lineHeight };
          } finally { frame.bitmap.close(); }
        } finally { document.close(); engine.close(); }
      });
      console.log(`${name}: ODT visible spacing ${JSON.stringify(odtSpacing)}`);
      // ano.ofd embeds valid outlines in SFNTs with a lowercase OS/2 tag and invalid cmap terminator.
      await page.goto(`${url}examples/viewer.html?fixture=ofdrw-ano.ofd`, { waitUntil: "domcontentloaded" });
      await page.waitForFunction(() => document.documentElement.dataset.ready === "true");
      await page.locator("docviewkit-viewer").evaluate((viewer) => {
        viewer.config = { ...viewer.config, features: { ...viewer.config.features, interactionMode: "text" } };
        viewer.shadowRoot.querySelectorAll(".continuous-page")[1].scrollIntoView();
      });
      await page.waitForFunction(() => {
        const fragments = [...document.querySelector("docviewkit-viewer").shadowRoot.querySelectorAll(".text-layer-fragment")];
        const subsetGlyphs = fragments.filter((node) => /[\ue000-\uf8ff]/u.test(node.textContent));
        return subsetGlyphs.length > 50 && subsetGlyphs.every((node) => node.style.font.includes("embedded"));
      }, null, { timeout: 30_000 });
      const ofdFonts = await page.evaluate(async () => {
        const { createOfficeEngine } = await import("/dist/engine.js");
        const { extendedFormatPack } = await import("/dist/extended-formats.js");
        const engine = await createOfficeEngine({ execution: "inline", formatPack: async () => extendedFormatPack });
        const document = await engine.open(await (await fetch("/tests/fixtures/ofdrw-ano.ofd")).arrayBuffer());
        try {
          const frame = await document.render({ unitIndex: 1, includeTextFragments: true });
          frame.bitmap.close();
          return document.diagnostics().filter(({ code }) => code === "FONT_LOAD_FAILED");
        } finally { document.close(); engine.close(); }
      });
      if (ofdFonts.length > 0) throw new Error(`${name} rejected OFD subset fonts: ${JSON.stringify(ofdFonts)}`);
      if (errors.length !== 0) throw new Error(`${name} browser errors:\n${errors.join("\n")}`);
      console.log(`${name}: viewer ready, ${state.canvasCount} canvas element(s), pinch ${JSON.stringify(pinchZoom)}`);
    } finally {
      await browser.close();
    }
  }
} finally {
  server.kill("SIGTERM");
  lines.close();
}
