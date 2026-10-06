import "/dist/viewer.js";

const viewer = document.querySelector("#viewer");
if (viewer === null) throw new Error("Viewer element is missing");
const fileInput = document.querySelector("#viewer-file-input");
if (!(fileInput instanceof HTMLInputElement)) throw new Error("Viewer file input is missing");

const parameters = new URLSearchParams(location.search);
const locale = parameters.get("locale") ?? "en";
const themeValue = parameters.get("theme") ?? "auto";
const theme = themeValue === "light" || themeValue === "dark" ? themeValue : "auto";
const fileDropEnabled = parameters.get("fileDrop") !== "false";
const filePickerEnabled = parameters.get("filePicker") === "true";

document.documentElement.lang = locale;
viewer.lang = locale;
viewer.setAttribute("theme", theme);
viewer.config = {
  locale,
  theme,
  navigation: "auto",
  features: {
    interactionModeSwitcher: parameters.get("interactionModeSwitcher") !== "false",
    print: parameters.get("print") !== "false",
    fullscreen: parameters.get("fullscreen") !== "false",
  },
  engine: {
    formatPack: () => import("/dist/extended-formats.js")
      .then(({ extendedFormatPack }) => extendedFormatPack),
  },
};

const openButton = fileInput.closest("label");
const openLabel = locale.toLowerCase().startsWith("zh") ? "打开本地文档" : "Open local document";
fileInput.ariaLabel = openLabel;
if (openButton instanceof HTMLElement) {
  openButton.title = openLabel;
  openButton.hidden = !filePickerEnabled;
}

async function openFile(file) {
  if (!(file instanceof File)) return;
  await viewer.open(file);
  document.documentElement.dataset.ready = "true";
}

function reportOpenError(cause) {
  document.documentElement.dataset.error = cause instanceof Error ? cause.message : String(cause);
  console.error(cause);
}

async function openFixture() {
  const fixture = parameters.get("fixture") ?? "visual-baseline.pptx";
  if (fixture.startsWith(".") || fixture.length > 160 || /[\\/\0]/u.test(fixture)
    || !/\.(?:ppt|pptx|pptm|ppsx|ppsm|potx|potm|odp|otp|fodp|dps|xls|xlsx|xlsm|xltx|xltm|ods|ots|fods|et|doc|docx|docm|dotx|dotm|odt|ott|wps|csv|rtf|pdf|xps|oxps|ofd)$/iu.test(fixture)) {
    throw new Error("fixture must be a supported local test filename");
  }
  const response = await fetch(`/tests/fixtures/${encodeURIComponent(fixture)}`, { cache: "no-store" });
  if (!response.ok) throw new Error(`fixture load failed: HTTP ${response.status}`);
  const file = new File([await response.arrayBuffer()], fixture, { type: "application/octet-stream" });
  await openFile(file);
}

fileInput.addEventListener("change", () => {
  const [file] = fileInput.files ?? [];
  void openFile(file).catch(reportOpenError);
  fileInput.value = "";
});

for (const eventName of ["dragenter", "dragover", "dragleave", "drop"]) {
  viewer.addEventListener(eventName, (event) => event.preventDefault());
}
if (fileDropEnabled) {
  viewer.addEventListener("drop", (event) => {
    const [file] = event.dataTransfer?.files ?? [];
    void openFile(file).catch(reportOpenError);
  });
}

openFixture().catch((cause) => {
  document.documentElement.dataset.error = cause instanceof Error ? cause.message : String(cause);
  throw cause;
});
