import { spawn } from "node:child_process";
import { createHash } from "node:crypto";
import {
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  renameSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { createRequire } from "node:module";
import { basename, dirname, relative, resolve, sep } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const runnerPath = fileURLToPath(import.meta.url);
const root = resolve(dirname(runnerPath), "..");
const defaultPdfRoot = resolve(root, "tests/external/pdfjs/test/pdfs");
const pdfjsRoot = resolve(root, "tests/external/pdfjs");
const pdfjsWebRoot = resolve(pdfjsRoot, "build/generic/web");
const defaultOutput = resolve(root, "output/accuracy/pdfjs-reference-corpus");
const resultPrefix = "OFFICEVIEWER_PDFJS_RESULT=";
const schemaVersion = 1;
const standardFixturePasswords = ["password"];

function usage() {
  return [
    "Usage: node scripts/run-pdfjs-reference-corpus.mjs [options]",
    "",
    "Options:",
    "  --pdf-root <directory>       PDF directory (default: tests/external/pdfjs/test/pdfs)",
    "  --output <directory>         PNGs and ledger (default: output/accuracy/pdfjs-reference-corpus)",
    "  --concurrency <count>        Isolated PDF workers (default: 4)",
    "  --host <chromium|node>       PDF.js canvas host (default: chromium)",
    "  --timeout-ms <milliseconds>  Per-PDF timeout (default: 90000)",
    "  --limit <count>              Deterministic prefix for runner smoke tests",
    "  --match <pattern>            Only basenames containing this text",
    "  --resume                     Reuse entries whose SHA-256 already completed",
    "  --worker <pdf>               Internal single-PDF worker mode",
    "  --help                       Show this help",
  ].join("\n");
}

function integer(value, option, minimum = 1) {
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed < minimum) {
    throw new Error(`${option} must be an integer >= ${minimum}`);
  }
  return parsed;
}

function parseArguments(arguments_) {
  const options = {
    pdfRoot: defaultPdfRoot,
    output: defaultOutput,
    concurrency: 4,
    host: "chromium",
    timeoutMs: 90_000,
    resume: false,
    match: undefined,
    limit: undefined,
    worker: undefined,
    workerRelative: undefined,
    workerSha256: undefined,
    workerStem: undefined,
  };
  for (let index = 0; index < arguments_.length; index += 1) {
    const argument = arguments_[index];
    if (argument === "--help") return { ...options, help: true };
    if (argument === "--resume") {
      options.resume = true;
      continue;
    }
    const value = arguments_[++index];
    if (value === undefined) throw new Error(`${usage()}\nMissing ${argument} value`);
    if (argument === "--pdf-root") options.pdfRoot = resolve(value);
    else if (argument === "--output") options.output = resolve(value);
    else if (argument === "--concurrency") options.concurrency = integer(value, argument);
    else if (argument === "--timeout-ms") options.timeoutMs = integer(value, argument, 1000);
    else if (argument === "--host") {
      if (!new Set(["chromium", "node"]).has(value)) throw new Error(`${argument} must be chromium or node`);
      options.host = value;
    }
    else if (argument === "--limit") options.limit = integer(value, argument);
    else if (argument === "--match") options.match = value;
    else if (argument === "--worker") options.worker = resolve(value);
    else if (argument === "--worker-relative") options.workerRelative = value;
    else if (argument === "--worker-sha256") options.workerSha256 = value;
    else if (argument === "--worker-stem") options.workerStem = value;
    else throw new Error(`${usage()}\nUnknown option ${argument}`);
  }
  return { ...options, help: false };
}

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

function safeRelative(rootPath, path) {
  const value = relative(rootPath, path);
  if (value === "" || value.startsWith(`..${sep}`) || value === ".." || value.startsWith(sep)) {
    throw new Error(`Path escapes the PDF root: ${path}`);
  }
  return value.split(sep).join("/");
}

function pdfFiles(directory) {
  const files = [];
  const visit = (current) => {
    for (const entry of readdirSync(current, { withFileTypes: true })) {
      const path = resolve(current, entry.name);
      if (entry.isDirectory()) visit(path);
      else if (entry.isFile() && entry.name.toLowerCase().endsWith(".pdf")) files.push(path);
    }
  };
  visit(directory);
  return files.sort((left, right) => left.localeCompare(right, "en"));
}

function safeStem(relativePath, digest) {
  const normalized = relativePath.normalize("NFKC").replace(/\.pdf$/iu, "")
    .replace(/[^A-Za-z0-9._-]+/gu, "-").replace(/^-+|-+$/gu, "");
  return `${normalized || "document"}-${digest.slice(0, 12)}`;
}

function writeJsonAtomic(path, value) {
  mkdirSync(dirname(path), { recursive: true });
  const temporary = `${path}.tmp-${process.pid}`;
  writeFileSync(temporary, `${JSON.stringify(value, null, 2)}\n`, { mode: 0o600 });
  renameSync(temporary, path);
}

function pdfjsCommit() {
  const gitPath = resolve(pdfjsRoot, ".git");
  if (!existsSync(gitPath)) return "unknown";
  const head = readFileSync(resolve(gitPath, "HEAD"), "utf8").trim();
  if (!head.startsWith("ref: ")) return head;
  const ref = head.slice("ref: ".length);
  const loose = resolve(gitPath, ref);
  if (existsSync(loose)) return readFileSync(loose, "utf8").trim();
  const packed = readFileSync(resolve(gitPath, "packed-refs"), "utf8")
    .split(/\r?\n/u).find((line) => line.endsWith(` ${ref}`));
  return packed?.split(" ", 1)[0] ?? "unknown";
}

function filePasswords() {
  const manifestPath = resolve(pdfjsRoot, "test/test_manifest.json");
  if (!existsSync(manifestPath)) return new Map();
  const manifest = JSON.parse(readFileSync(manifestPath, "utf8"));
  const passwords = new Map();
  for (const entry of manifest) {
    if (typeof entry.file !== "string" || typeof entry.password !== "string") continue;
    const path = entry.file.replace(/^pdfs\//u, "");
    const values = passwords.get(path) ?? new Set();
    values.add(entry.password);
    passwords.set(path, values);
  }
  return new Map([...passwords].map(([path, values]) => [path, [...values]]));
}

function browserLikeCanvasContext(context) {
  return new Proxy(context, {
    get(target, property) {
      const value = Reflect.get(target, property, target);
      return typeof value === "function" ? value.bind(target) : value;
    },
    set(target, property, value) {
      try {
        return Reflect.set(target, property, value, target);
      } catch (cause) {
        // Browsers ignore invalid CSS font declarations and retain the current
        // font. Native canvas bindings throw instead, which is host-only drift.
        if (property === "font") return true;
        throw cause;
      }
    },
  });
}

function routedAsset(pathname, pdfPath) {
  const roots = [
    ["/build/", resolve(pdfjsRoot, "build/generic/build")],
    ["/cmaps/", resolve(pdfjsWebRoot, "cmaps")],
    ["/standard_fonts/", resolve(pdfjsWebRoot, "standard_fonts")],
    ["/wasm/", resolve(pdfjsWebRoot, "wasm")],
  ];
  if (pathname === "/document.pdf") return pdfPath;
  if (pathname === "/viewer.css") return resolve(pdfjsWebRoot, "viewer.css");
  for (const [prefix, directory] of roots) {
    if (!pathname.startsWith(prefix)) continue;
    const path = resolve(directory, pathname.slice(prefix.length));
    if (path.startsWith(`${directory}${sep}`) && existsSync(path) && statSync(path).isFile()) return path;
  }
  return undefined;
}

function contentType(path) {
  if (path.endsWith(".mjs") || path.endsWith(".js")) return "text/javascript; charset=utf-8";
  if (path.endsWith(".pdf")) return "application/pdf";
  if (path.endsWith(".wasm")) return "application/wasm";
  return "application/octet-stream";
}

async function chromiumPdfjsRender(options, outputDirectory, started, passwords) {
  const { chromium } = await import("playwright-core");
  const browser = await chromium.launch({ channel: "chrome", headless: true });
  try {
    const context = await browser.newContext({
      locale: "en-US",
      timezoneId: "Asia/Shanghai",
      deviceScaleFactor: 1,
    });
    await context.addInitScript(() => {
      Math.sumPrecise ??= (values) => {
        let sum = 0;
        let compensation = 0;
        for (const value of values) {
          const next = sum + value;
          compensation += Math.abs(sum) >= Math.abs(value)
            ? sum - next + value
            : value - next + sum;
          sum = next;
        }
        return sum + compensation;
      };
    });
    await context.route("http://pdfjs.local/**", async (route) => {
      const pathname = new URL(route.request().url()).pathname;
      if (pathname === "/") {
        await route.fulfill({
          contentType: "text/html; charset=utf-8",
          body: "<!doctype html><html><head><link rel=\"stylesheet\" href=\"/viewer.css\"></head><body></body></html>",
        });
        return;
      }
      const asset = routedAsset(pathname, options.worker);
      if (asset === undefined) {
        await route.fulfill({ status: 404, body: "Not Found" });
        return;
      }
      await route.fulfill({ contentType: contentType(asset), body: readFileSync(asset) });
    });
    const page = await context.newPage();
    await page.goto("http://pdfjs.local/", { waitUntil: "domcontentloaded" });
    const metadata = await page.evaluate(async ({ passwords: candidatePasswords }) => {
      const pdfjs = await import("/build/pdf.mjs");
      pdfjs.GlobalWorkerOptions.workerSrc = "/build/pdf.worker.mjs";
      const loadingTask = pdfjs.getDocument({
        url: "/document.pdf",
        cMapUrl: "/cmaps/",
        cMapPacked: true,
        standardFontDataUrl: "/standard_fonts/",
        wasmUrl: "/wasm/",
        useWasm: true,
        useSystemFonts: true,
        enableXfa: true,
        isEvalSupported: false,
      });
      let passwordIndex = 0;
      loadingTask.onPassword = (updatePassword) => {
        const password = candidatePasswords[passwordIndex++];
        if (password === undefined) throw new Error("PDF_PASSWORD_REQUIRED");
        updatePassword(password);
      };
      globalThis.pdfjsCorpusLoadingTask = loadingTask;
      globalThis.pdfjsCorpusApi = pdfjs;
      globalThis.pdfjsCorpusDocument = await loadingTask.promise;
      return { pageCount: globalThis.pdfjsCorpusDocument.numPages };
    }, { passwords });
    const pages = [];
    for (let pageNumber = 1; pageNumber <= metadata.pageCount; pageNumber += 1) {
      const pageStarted = performance.now();
      const capture = await page.evaluate(async (number) => {
        const pdfPage = await globalThis.pdfjsCorpusDocument.getPage(number);
        const viewport = pdfPage.getViewport({ scale: 1 });
        const xfaHtml = await pdfPage.getXfa();
        if (xfaHtml) {
          document.body.replaceChildren();
          document.body.style.margin = "0";
          const container = document.createElement("div");
          container.id = "pdfjs-xfa-capture";
          container.style.position = "relative";
          container.style.overflow = "hidden";
          container.style.width = `${Math.ceil(viewport.width)}px`;
          container.style.height = `${Math.ceil(viewport.height)}px`;
          container.style.background = "#ffffff";
          document.body.append(container);
          await globalThis.pdfjsCorpusApi.XfaLayer.render({
            viewport: viewport.clone({ dontFlip: true }),
            div: container,
            xfaHtml,
            annotationStorage: globalThis.pdfjsCorpusDocument.annotationStorage,
            linkService: { addLinkAttributes() {}, getDestinationHash() { return ""; } },
            intent: "display",
          });
          return { width: Math.ceil(viewport.width), height: Math.ceil(viewport.height), xfa: true };
        }
        const canvas = document.createElement("canvas");
        canvas.width = Math.ceil(viewport.width);
        canvas.height = Math.ceil(viewport.height);
        await pdfPage.render({
          canvasContext: canvas.getContext("2d"),
          viewport,
          background: "#ffffff",
          intent: "display",
        }).promise;
        pdfPage.cleanup();
        return { width: canvas.width, height: canvas.height, dataUrl: canvas.toDataURL("image/png") };
      }, pageNumber);
      const png = capture.xfa
        ? await page.locator("#pdfjs-xfa-capture").screenshot({ type: "png" })
        : Buffer.from(capture.dataUrl.slice("data:image/png;base64,".length), "base64");
      const filename = `page-${String(pageNumber).padStart(4, "0")}.png`;
      writeFileSync(resolve(outputDirectory, filename), png, { mode: 0o600 });
      pages.push({
        page: pageNumber,
        width: capture.width,
        height: capture.height,
        png: `reference/${options.workerStem}/${filename}`,
        pngSha256: sha256(png),
        durationMs: Math.round((performance.now() - pageStarted) * 10) / 10,
      });
    }
    await page.evaluate(async () => globalThis.pdfjsCorpusLoadingTask.destroy());
    return {
      path: options.workerRelative,
      sha256: options.workerSha256,
      bytes: statSync(options.worker).size,
      status: "rendered",
      oracle: "pdf.js",
      oracleHost: "chromium",
      pages,
      pageCount: pages.length,
      durationMs: Math.round((performance.now() - started) * 10) / 10,
    };
  } finally {
    await browser.close();
  }
}

async function worker(options) {
  if (options.workerRelative === undefined || options.workerSha256 === undefined
      || options.workerStem === undefined) {
    throw new Error("Internal worker metadata is incomplete");
  }
  const started = performance.now();
  const outputDirectory = resolve(options.output, "reference", options.workerStem);
  mkdirSync(outputDirectory, { recursive: true });
  const passwords = JSON.parse(process.env.OFFICEVIEWER_PDF_PASSWORDS ?? "[]");
  if (options.host === "chromium") {
    try {
      return await chromiumPdfjsRender(options, outputDirectory, started, passwords);
    } catch (cause) {
      return {
        path: options.workerRelative,
        sha256: options.workerSha256,
        bytes: statSync(options.worker).size,
        status: "pdfjs-failed",
        oracle: "acrobat-required",
        code: cause?.name ?? cause?.code ?? "PDFJS_CHROMIUM_ERROR",
        message: cause instanceof Error ? cause.message : String(cause),
        durationMs: Math.round((performance.now() - started) * 10) / 10,
      };
    }
  }
  const require = createRequire(resolve(pdfjsRoot, "package.json"));
  const canvasApi = require("@napi-rs/canvas");
  for (const key of ["DOMMatrix", "ImageData", "Path2D"]) globalThis[key] = canvasApi[key];
  // The pinned PDF.js commit uses the ES2026 Math.sumPrecise proposal, while
  // supported Node runtimes do not all expose it yet. This compensated sum is
  // exact for the byte and integer arrays used by PDF.js and stable for widths.
  Math.sumPrecise ??= (values) => {
    let sum = 0;
    let compensation = 0;
    for (const value of values) {
      if (typeof value !== "number") throw new TypeError("Math.sumPrecise values must be numbers");
      const next = sum + value;
      compensation += Math.abs(sum) >= Math.abs(value)
        ? sum - next + value
        : value - next + sum;
      sum = next;
    }
    return sum + compensation;
  };
  const pdfjs = await import(pathToFileURL(resolve(pdfjsRoot, "build/generic/build/pdf.mjs")));
  const fileUrl = (path) => `${pathToFileURL(resolve(path)).href}/`;
  const bytes = readFileSync(options.worker);
  let passwordIndex = 0;
  const loadingTask = pdfjs.getDocument({
    data: new Uint8Array(bytes),
    cMapUrl: fileUrl(resolve(pdfjsWebRoot, "cmaps")),
    cMapPacked: true,
    standardFontDataUrl: fileUrl(resolve(pdfjsWebRoot, "standard_fonts")),
    wasmUrl: fileUrl(resolve(pdfjsWebRoot, "wasm")),
    useWasm: false,
    // PDF.js' Node test environment uses bundled standard-font data. Native
    // canvas libraries reject a few malformed PDF font names that browsers
    // sanitize, so system font lookup would create host-only false failures.
    useSystemFonts: false,
    enableXfa: true,
    isEvalSupported: false,
  });
  loadingTask.onPassword = (updatePassword) => {
    const password = passwords[passwordIndex++];
    if (password === undefined) throw new Error("PDF_PASSWORD_REQUIRED");
    updatePassword(password);
  };
  let document;
  try {
    document = await loadingTask.promise;
    const pages = [];
    for (let pageNumber = 1; pageNumber <= document.numPages; pageNumber += 1) {
      const pageStarted = performance.now();
      const page = await document.getPage(pageNumber);
      const viewport = page.getViewport({ scale: 1 });
      const width = Math.ceil(viewport.width);
      const height = Math.ceil(viewport.height);
      if (!Number.isSafeInteger(width) || !Number.isSafeInteger(height) || width < 1 || height < 1) {
        throw new Error(`PDFJS_INVALID_PAGE_SIZE: page ${pageNumber} is ${width}x${height}`);
      }
      const canvas = canvasApi.createCanvas(width, height);
      await page.render({
        canvasContext: browserLikeCanvasContext(canvas.getContext("2d")),
        viewport,
        background: "#ffffff",
        intent: "display",
      }).promise;
      const png = canvas.toBuffer("image/png");
      const filename = `page-${String(pageNumber).padStart(4, "0")}.png`;
      writeFileSync(resolve(outputDirectory, filename), png, { mode: 0o600 });
      pages.push({
        page: pageNumber,
        width,
        height,
        png: `reference/${options.workerStem}/${filename}`,
        pngSha256: sha256(png),
        durationMs: Math.round((performance.now() - pageStarted) * 10) / 10,
      });
      page.cleanup();
    }
    return {
      path: options.workerRelative,
      sha256: options.workerSha256,
      bytes: bytes.length,
      status: "rendered",
      oracle: "pdf.js",
      oracleHost: "node",
      pages,
      pageCount: pages.length,
      durationMs: Math.round((performance.now() - started) * 10) / 10,
    };
  } catch (cause) {
    if (cause?.name === "Error") {
      try {
        await loadingTask.destroy().catch(() => undefined);
        return await chromiumPdfjsRender(options, outputDirectory, started, passwords);
      } catch (browserCause) {
        return {
          path: options.workerRelative,
          sha256: options.workerSha256,
          bytes: bytes.length,
          status: "pdfjs-failed",
          oracle: "acrobat-required",
          code: browserCause?.name ?? browserCause?.code ?? "PDFJS_CHROMIUM_ERROR",
          message: browserCause instanceof Error ? browserCause.message : String(browserCause),
          nodeMessage: cause instanceof Error ? cause.message : String(cause),
          durationMs: Math.round((performance.now() - started) * 10) / 10,
        };
      }
    }
    return {
      path: options.workerRelative,
      sha256: options.workerSha256,
      bytes: bytes.length,
      status: "pdfjs-failed",
      oracle: "acrobat-required",
      code: cause?.name ?? cause?.code ?? "PDFJS_ERROR",
      message: cause instanceof Error ? cause.message : String(cause),
      durationMs: Math.round((performance.now() - started) * 10) / 10,
    };
  } finally {
    await loadingTask.destroy().catch(() => undefined);
  }
}

function workerProcess(options, entry, passwords) {
  return new Promise((resolvePromise) => {
    const arguments_ = [
      "--max-old-space-size=1024",
      runnerPath,
      "--worker", entry.path,
      "--worker-relative", entry.relative,
      "--worker-sha256", entry.sha256,
      "--worker-stem", entry.stem,
      "--host", options.host,
      "--output", options.output,
    ];
    const child = spawn(process.execPath, arguments_, {
      cwd: root,
      env: { ...process.env, OFFICEVIEWER_PDF_PASSWORDS: JSON.stringify(passwords) },
      stdio: ["ignore", "pipe", "pipe"],
    });
    let stdout = "";
    let stderr = "";
    child.stdout.on("data", (chunk) => { stdout += chunk; });
    child.stderr.on("data", (chunk) => { stderr += chunk; });
    const timeout = setTimeout(() => child.kill("SIGKILL"), options.timeoutMs);
    child.once("error", (cause) => {
      clearTimeout(timeout);
      resolvePromise({
        path: entry.relative,
        sha256: entry.sha256,
        bytes: entry.bytes,
        status: "pdfjs-failed",
        oracle: "acrobat-required",
        code: "WORKER_START_FAILED",
        message: cause.message,
      });
    });
    child.once("exit", (code, signal) => {
      clearTimeout(timeout);
      const line = stdout.split(/\r?\n/u).findLast((value) => value.startsWith(resultPrefix));
      if (line !== undefined) {
        resolvePromise(JSON.parse(Buffer.from(line.slice(resultPrefix.length), "base64url").toString("utf8")));
        return;
      }
      resolvePromise({
        path: entry.relative,
        sha256: entry.sha256,
        bytes: entry.bytes,
        status: "pdfjs-failed",
        oracle: "acrobat-required",
        code: signal === "SIGKILL" ? "PDFJS_TIMEOUT" : "PDFJS_WORKER_FAILED",
        message: (stderr || stdout || `Worker exited with ${code ?? signal}`).trim().slice(-4_000),
      });
    });
  });
}

async function main(options) {
  const required = [
    resolve(pdfjsRoot, "build/generic/build/pdf.mjs"),
    resolve(pdfjsRoot, "build/generic/build/pdf.worker.mjs"),
  ];
  if (options.host === "node") required.push(resolve(pdfjsRoot, "node_modules/@napi-rs/canvas"));
  const missing = required.filter((path) => !existsSync(path));
  if (missing.length !== 0) {
    throw new Error(
      `PDF.js reference runtime is missing: ${missing.join(", ")}. `
      + "Run npm ci and npx gulp generic in tests/external/pdfjs first.",
    );
  }
  const allFiles = pdfFiles(options.pdfRoot);
  let files = options.match === undefined
    ? allFiles
    : allFiles.filter((path) => basename(path).toLowerCase().includes(options.match.toLowerCase()));
  if (options.limit !== undefined) files = files.slice(0, options.limit);
  if (files.length === 0) throw new Error("No PDF files matched the requested corpus selection");
  const entries = files.map((path) => {
    const bytes = readFileSync(path);
    const digest = sha256(bytes);
    const relativePath = safeRelative(options.pdfRoot, path);
    return {
      path,
      relative: relativePath,
      sha256: digest,
      bytes: statSync(path).size,
      stem: safeStem(relativePath, digest),
    };
  });
  const ledgerPath = resolve(options.output, "ledger.json");
  const previous = options.resume && existsSync(ledgerPath)
    ? JSON.parse(readFileSync(ledgerPath, "utf8"))
    : undefined;
  const completed = new Map((previous?.files ?? [])
    .filter((entry) => entry.oracleHost === options.host)
    .map((entry) => [`${entry.path}\0${entry.sha256}`, entry]));
  const results = [];
  const pending = [];
  for (const entry of entries) {
    const prior = completed.get(`${entry.relative}\0${entry.sha256}`);
    if (prior !== undefined && ["rendered", "acrobat-rendered"].includes(prior.status)) results.push(prior);
    else pending.push(entry);
  }
  const passwords = filePasswords();
  const started = new Date().toISOString();
  const writeLedger = () => {
    const files_ = [...results].sort((left, right) => left.path.localeCompare(right.path, "en"));
    const rendered = files_.filter(({ status }) => status === "rendered");
    const acrobatRequired = files_.filter(({ status }) => status === "pdfjs-failed");
    writeJsonAtomic(ledgerPath, {
      schemaVersion,
      generatedAt: new Date().toISOString(),
      startedAt: started,
      corpus: {
        root: options.pdfRoot,
        discoveredPdfs: allFiles.length,
        selectedPdfs: entries.length,
        selectedBytes: entries.reduce((sum, entry) => sum + entry.bytes, 0),
      },
      environment: {
        pdfjsCommit: pdfjsCommit(),
        renderer: options.host === "chromium"
          ? "pdf.js generic build in Chromium"
          : "pdf.js generic build with @napi-rs/canvas",
        oracleHost: options.host,
        scale: 1,
        devicePixelRatio: 1,
        background: "#ffffff",
        useWasm: options.host === "chromium",
        useSystemFonts: options.host === "chromium",
        activeContent: false,
      },
      summary: {
        completedPdfs: files_.length,
        pendingPdfs: entries.length - files_.length,
        renderedPdfs: rendered.length,
        renderedPages: rendered.reduce((sum, entry) => sum + entry.pageCount, 0),
        chromiumRenderedPdfs: rendered.filter(({ oracleHost }) => oracleHost === "chromium").length,
        acrobatRequiredPdfs: acrobatRequired.length,
      },
      files: files_,
    });
  };
  writeLedger();
  let next = 0;
  const runNext = async () => {
    while (next < pending.length) {
      const index = next++;
      const entry = pending[index];
      const result = await workerProcess(options, entry, [
        ...(passwords.get(entry.relative) ?? []),
        ...standardFixturePasswords,
      ]);
      results.push(result);
      writeLedger();
      const pageDetail = result.status === "rendered" ? `, ${result.pageCount} pages` : `, ${result.code}`;
      console.log(
        `${String(results.length).padStart(String(entries.length).length)}/${entries.length} `
        + `${result.status.padEnd(13)} ${entry.relative}${pageDetail}`,
      );
    }
  };
  await Promise.all(Array.from({ length: Math.min(options.concurrency, pending.length) }, runNext));
  writeLedger();
  const failed = results.filter(({ status }) => status === "pdfjs-failed").length;
  console.log(
    `PDF.js reference corpus complete: ${results.length}/${entries.length} PDFs, `
    + `${failed} require Acrobat, ledger=${ledgerPath}`,
  );
}

const options = parseArguments(process.argv.slice(2));
if (options.help) {
  console.log(usage());
} else if (options.worker !== undefined) {
  const result = await worker(options);
  console.log(`${resultPrefix}${Buffer.from(JSON.stringify(result)).toString("base64url")}`);
} else {
  await main(options);
}
