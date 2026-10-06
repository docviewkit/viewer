import { locale, t } from "./i18n.js?v=20261002-3";

const MAX_FILE_BYTES = 128 * 1024 * 1024;
const MAX_RENDER_PIXELS = 24_000_000;
const FORMAT_PACK_CANDIDATES = Object.freeze(["odf", "iwork", "legacy-office", "wps", "pdf", "xps"]);
const ENGLISH_PPTX_SAMPLE_BASE64 = "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";
const sdkManifest = await fetch("/sdk/version.json", { cache: "no-store" }).then((response) => {
  if (!response.ok) throw new Error(`SDK version request failed with ${response.status}`);
  return response.json();
});
if (typeof sdkManifest.version !== "string" || !/^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/u.test(sdkManifest.version)) {
  throw new Error("SDK version manifest is invalid");
}
const sdkBase = sdkManifest.immutableAssets === false
  ? "/sdk/"
  : `/sdk/v${sdkManifest.version}/`;
const viewerPromise = import(`${sdkBase}viewer.js`);
const runtimePromise = import(`${sdkBase}extended-formats.js`).then(async ({ extendedFormatPack }) => {
  const sources = await Promise.all(FORMAT_PACK_CANDIDATES.map((candidate) => extendedFormatPack.load(candidate)));
  const coreUrl = `${sdkBase}office-viewer-core.wasm`;
  const modules = new Map(await Promise.all([...new Set([coreUrl, ...sources.map(String)])].map(async (url) => {
    try {
      const response = await fetch(url);
      if (!response.ok) throw new Error(`SDK format preload failed with ${response.status}: ${url}`);
      return [url, await WebAssembly.compile(await response.arrayBuffer())];
    } catch {
      return [url, url];
    }
  })));
  return {
    wasm: modules.get(coreUrl),
    formatPack: Object.freeze({
      async load(candidate) {
        const source = await extendedFormatPack.load(candidate);
        return modules.get(String(source)) ?? source;
      },
    }),
  };
});
const [, runtime] = await Promise.all([
  viewerPromise,
  runtimePromise,
]);

function element(id) {
  const node = document.getElementById(id);
  if (!node) throw new Error(`Demo element #${id} is missing.`);
  return node;
}

const ui = {
  fileInput: element("demo-file-input"),
  sampleButton: element("demo-sample-button"),
  format: element("demo-format"),
  fileName: element("demo-file-name"),
  fileMeta: element("demo-file-meta"),
  status: element("demo-status"),
  host: element("demo-viewer-host"),
  viewer: element("demo-viewer"),
};

let openRevision = 0;
let dragDepth = 0;
let disposed = false;

function localize(chinese, english) {
  return locale === "en" ? english : chinese;
}

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function errorMessage(error) {
  if (error instanceof Error) return error.message;
  return localize("未知错误", "Unknown error");
}

function setStatus(message, tone = "ready") {
  ui.status.textContent = message;
  ui.status.dataset.tone = tone;
}

function sampleFile() {
  const bytes = Uint8Array.from(atob(ENGLISH_PPTX_SAMPLE_BASE64), (character) => character.charCodeAt(0));
  return new File([bytes], "DocViewKit-visual-sample.pptx", { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
}

async function openFile(file) {
  if (!(file instanceof File)) return;
  if (file.size > MAX_FILE_BYTES) {
    setStatus(localize("文件超过 128 MB 限制。", "The file exceeds the 128 MB limit."), "error");
    return;
  }

  const revision = ++openRevision;
  ui.format.textContent = localize("识别中", "Identifying");
  ui.fileName.textContent = file.name || localize("未命名文件", "Untitled file");
  ui.fileMeta.textContent = `${formatBytes(file.size)} · ${t("浏览器本地")}`;
  setStatus(localize("正在打开文档…", "Opening document…"), "loading");

  try {
    let password;
    let info;
    for (let attempt = 0; attempt < 3; attempt += 1) {
      try {
        info = await ui.viewer.open(file, password === undefined ? {} : { password });
        break;
      } catch (error) {
        if (error?.code !== "PDF_PASSWORD_REQUIRED" && error?.code !== "PDF_PASSWORD_INCORRECT") throw error;
        const promptMessage = error.code === "PDF_PASSWORD_INCORRECT"
          ? localize("密码错误，请重试：", "Incorrect password. Try again:")
          : localize("请输入 PDF 密码：", "Enter the PDF password:");
        const supplied = window.prompt(promptMessage);
        if (supplied === null) throw error;
        password = supplied;
      }
    }
    if (info === undefined) throw new Error(localize("PDF 密码重试次数过多", "Too many PDF password attempts"));
    if (revision !== openRevision || disposed) return;
    ui.format.textContent = info.format.toUpperCase();
    ui.fileMeta.textContent = `${formatBytes(file.size)} · ${info.kind} · ${t("浏览器本地")}`;
    setStatus(localize("文档已就绪。", "Document ready."));
  } catch (error) {
    if (revision !== openRevision || disposed || error?.name === "AbortError") return;
    ui.format.textContent = localize("不支持", "Unsupported");
    setStatus(localize(`打开失败：${errorMessage(error)}`, `Open failed: ${errorMessage(error)}`), "error");
  }
}

ui.viewer.config = {
  locale,
  theme: document.documentElement.dataset.theme ?? "auto",
  navigation: "auto",
  features: { interactionMode: "object", interactionModeSwitcher: true },
  engine: {
    wasm: runtime.wasm,
    formatPack: () => Promise.resolve(runtime.formatPack),
    limits: { inputBytes: MAX_FILE_BYTES, renderPixels: MAX_RENDER_PIXELS },
  },
};

ui.fileInput.addEventListener("change", () => {
  const [file] = ui.fileInput.files ?? [];
  void openFile(file);
  ui.fileInput.value = "";
});
ui.sampleButton.addEventListener("click", () => void openFile(sampleFile()));

ui.viewer.addEventListener("docviewkit-diagnostic", (event) => {
  const diagnostics = event.detail?.diagnostics ?? [];
  const count = diagnostics.filter(({ severity }) => severity === "warning" || severity === "error" || severity === "fatal").length;
  if (count > 0) setStatus(localize(`${count} 项兼容诊断可查看。`, `${count} compatibility diagnostics available.`));
});

ui.host.addEventListener("dragenter", (event) => {
  event.preventDefault();
  dragDepth += 1;
  ui.host.dataset.dragging = "true";
});
ui.host.addEventListener("dragover", (event) => event.preventDefault());
ui.host.addEventListener("dragleave", (event) => {
  event.preventDefault();
  dragDepth = Math.max(0, dragDepth - 1);
  if (dragDepth === 0) ui.host.dataset.dragging = "false";
});
ui.host.addEventListener("drop", (event) => {
  event.preventDefault();
  dragDepth = 0;
  ui.host.dataset.dragging = "false";
  const [file] = event.dataTransfer?.files ?? [];
  void openFile(file);
});

function syncTheme() {
  const theme = document.documentElement.dataset.theme;
  if (theme === "light" || theme === "dark") ui.viewer.setAttribute("theme", theme);
}

const themeObserver = new MutationObserver(syncTheme);
themeObserver.observe(document.documentElement, { attributes: true, attributeFilter: ["data-theme"] });
syncTheme();

window.addEventListener("pagehide", () => {
  disposed = true;
  themeObserver.disconnect();
  ui.viewer.destroy();
}, { once: true });

ui.fileInput.disabled = false;
ui.sampleButton.disabled = false;
setStatus(localize("本地 Viewer 已就绪。", "Local Viewer ready."));
void openFile(sampleFile());
