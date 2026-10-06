const version = "__DOCVIEWKIT_VERSION__";
const sdkBase = new URL(`../sdk/v${version}/`, document.baseURI);
await import(new URL("viewer.js", sdkBase).href);

const viewer = document.querySelector("#viewer");
const fileInput = document.querySelector("#file");
const status = document.querySelector("#status");

viewer.config = {
  locale: "en",
  theme: "auto",
  navigation: "auto",
  engine: {
    formatPack: () => import(new URL("extended-formats.js", sdkBase).href)
      .then(({ extendedFormatPack }) => extendedFormatPack),
  },
};

fileInput.addEventListener("change", async () => {
  const [file] = fileInput.files ?? [];
  if (!file) return;
  status.textContent = `Opening ${file.name}…`;
  try {
    const info = await viewer.open(file);
    status.textContent = `${info.format.toUpperCase()} ready · processed locally`;
  } catch (error) {
    status.textContent = error instanceof Error ? error.message : "The file could not be opened.";
  } finally {
    fileInput.value = "";
  }
});

status.textContent = `Viewer v${version} ready`;
