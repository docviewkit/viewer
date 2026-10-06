import { createReadStream } from "node:fs";
import { realpath, stat } from "node:fs/promises";
import { createServer } from "node:http";
import { extname, resolve, sep } from "node:path";

const root = await realpath(resolve(process.cwd()));
const values = new Map();
for (let index = 2; index < process.argv.length; index += 1) {
  const argument = process.argv[index];
  if (!argument.startsWith("--")) continue;
  const [key, inlineValue] = argument.split("=", 2);
  const next = process.argv[index + 1];
  if (inlineValue !== undefined) {
    values.set(key, inlineValue);
  } else if (next !== undefined && !next.startsWith("--")) {
    values.set(key, next);
    index += 1;
  } else {
    values.set(key, "true");
  }
}
const fixtureRoot = resolve(values.get("--fixture-root") ?? resolve(root, "tests/fixtures"));
const fontRoot = resolve(values.get("--font-root") ?? resolve(root, "node_modules/@embedpdf/fonts-sc/fonts"));
const servedRoots = await Promise.all([
  ["/examples/", resolve(root, "examples")],
  ["/dist/", resolve(root, "dist")],
  ["/tests/fixtures/", fixtureRoot],
  ["/tests/fonts/", fontRoot],
].map(async ([prefix, logicalRoot]) => ({
  prefix,
  logicalRoot,
  realRoot: await realpath(logicalRoot),
})));
for (const required of [
  "examples/viewer.html",
  "examples/viewer.js",
  "examples/inspector.html",
  "examples/inspector.js",
  "dist/index.js",
  "dist/engine.js",
  "dist/viewer.js",
  "dist/extended-formats.js",
  "dist/office-viewer-core.wasm",
  "dist/office-viewer-calc.wasm",
  "dist/office-viewer-odf.wasm",
  "dist/office-viewer-legacy-office.wasm",
  "dist/office-viewer-pdf.wasm",
  "dist/office-viewer-xps.wasm",
  "dist/office-viewer-ofd.wasm",
  "dist/image-codec-worker.js",
  "dist/resvg.wasm",
]) {
  try {
    if (!(await stat(resolve(root, required))).isFile()) throw new Error();
  } catch {
    throw new Error(`Viewer asset is missing: ${required}. Run npm run build before npm run inspect.`);
  }
}
const host = values.get("--host") ?? "127.0.0.1";
const portText = values.get("--port") ?? process.env.PORT ?? "4173";
const port = Number(portText);
if (!Number.isSafeInteger(port) || port < 0 || port > 65_535) {
  throw new Error(`Invalid port: ${portText}`);
}

const contentTypes = new Map([
  [".css", "text/css; charset=utf-8"],
  [".csv", "text/csv; charset=utf-8"],
  [".html", "text/html; charset=utf-8"],
  [".js", "text/javascript; charset=utf-8"],
  [".json", "application/json; charset=utf-8"],
  [".map", "application/json; charset=utf-8"],
  [".mjs", "text/javascript; charset=utf-8"],
  [".otf", "font/otf"],
  [".png", "image/png"],
  [".rtf", "application/rtf"],
  [".xps", "application/vnd.ms-xpsdocument"],
  [".oxps", "application/oxps"],
  [".ofd", "application/ofd"],
  [".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"],
  [".doc", "application/msword"],
  [".docm", "application/vnd.ms-word.document.macroEnabled.12"],
  [".dotm", "application/vnd.ms-word.template.macroEnabled.12"],
  [".dotx", "application/vnd.openxmlformats-officedocument.wordprocessingml.template"],
  [".odp", "application/vnd.oasis.opendocument.presentation"],
  [".fodp", "application/vnd.oasis.opendocument.presentation-flat-xml"],
  [".otp", "application/vnd.oasis.opendocument.presentation-template"],
  [".ods", "application/vnd.oasis.opendocument.spreadsheet"],
  [".ots", "application/vnd.oasis.opendocument.spreadsheet-template"],
  [".odt", "application/vnd.oasis.opendocument.text"],
  [".fodt", "application/vnd.oasis.opendocument.text-flat-xml"],
  [".ott", "application/vnd.oasis.opendocument.text-template"],
  [".potm", "application/vnd.ms-powerpoint.template.macroEnabled.12"],
  [".potx", "application/vnd.openxmlformats-officedocument.presentationml.template"],
  [".ppsm", "application/vnd.ms-powerpoint.slideshow.macroEnabled.12"],
  [".ppsx", "application/vnd.openxmlformats-officedocument.presentationml.slideshow"],
  [".pptm", "application/vnd.ms-powerpoint.presentation.macroEnabled.12"],
  [".pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation"],
  [".ppt", "application/vnd.ms-powerpoint"],
  [".dps", "application/vnd.kingsoft.presentation"],
  [".wps", "application/vnd.kingsoft.writer"],
  [".key", "application/vnd.apple.keynote"],
  [".pages", "application/vnd.apple.pages"],
  [".numbers", "application/vnd.apple.numbers"],
  [".pdf", "application/pdf"],
  [".ttf", "font/ttf"],
  [".woff", "font/woff"],
  [".woff2", "font/woff2"],
  [".xlsm", "application/vnd.ms-excel.sheet.macroEnabled.12"],
  [".xltm", "application/vnd.ms-excel.template.macroEnabled.12"],
  [".xltx", "application/vnd.openxmlformats-officedocument.spreadsheetml.template"],
  [".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"],
  [".xls", "application/vnd.ms-excel"],
  [".et", "application/vnd.kingsoft.spreadsheets"],
  [".wasm", "application/wasm"],
]);

function headers(contentType) {
  return {
    "Cache-Control": "no-store",
    "Content-Security-Policy": "default-src 'self'; script-src 'self' 'wasm-unsafe-eval'; worker-src 'self'; connect-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; media-src 'self' blob:; font-src 'self' blob: data:; object-src 'none'; base-uri 'none'",
    "Content-Type": contentType,
    "X-Content-Type-Options": "nosniff",
  };
}

function sendText(response, statusCode, message) {
  response.writeHead(statusCode, headers("text/plain; charset=utf-8"));
  response.end(message);
}

const server = createServer(async (request, response) => {
  if (request.method !== "GET" && request.method !== "HEAD") {
    response.setHeader("Allow", "GET, HEAD");
    sendText(response, 405, "Method Not Allowed");
    return;
  }

  let pathname;
  try {
    pathname = decodeURIComponent(new URL(request.url ?? "/", "http://localhost").pathname);
  } catch {
    sendText(response, 400, "Bad Request");
    return;
  }
  if (pathname === "/") {
    response.writeHead(302, {
      ...headers("text/plain; charset=utf-8"),
      Location: "/examples/viewer.html?fixture=visual-baseline.pptx",
    });
    response.end("Redirecting to DocViewKit Viewer");
    return;
  }

  if (pathname.split("/").some((segment) => segment.startsWith("."))) {
    sendText(response, 403, "Forbidden");
    return;
  }

  const servedRoot = servedRoots.find(({ prefix }) => pathname.startsWith(prefix));
  if (servedRoot === undefined) {
    sendText(response, 404, "Not Found");
    return;
  }
  const filePath = resolve(servedRoot.logicalRoot, pathname.slice(servedRoot.prefix.length));
  if (!filePath.startsWith(`${servedRoot.logicalRoot}${sep}`)) {
    sendText(response, 403, "Forbidden");
    return;
  }

  let metadata;
  let realFilePath;
  try {
    realFilePath = await realpath(filePath);
    if (!realFilePath.startsWith(`${servedRoot.realRoot}${sep}`)) {
      sendText(response, 403, "Forbidden");
      return;
    }
    metadata = await stat(realFilePath);
  } catch {
    sendText(response, 404, "Not Found");
    return;
  }
  if (!metadata.isFile()) {
    sendText(response, 404, "Not Found");
    return;
  }

  const contentType = contentTypes.get(extname(realFilePath).toLowerCase()) ?? "application/octet-stream";
  response.writeHead(200, {
    ...headers(contentType),
    "Content-Length": metadata.size,
  });
  if (request.method === "HEAD") {
    response.end();
    return;
  }
  const stream = createReadStream(realFilePath);
  stream.on("error", () => response.destroy());
  stream.pipe(response);
});

server.on("error", (cause) => {
  console.error(cause instanceof Error ? cause.message : cause);
  process.exitCode = 1;
});

server.listen(port, host, () => {
  const displayHost = host === "0.0.0.0" || host === "::" ? "127.0.0.1" : host;
  const address = server.address();
  const listeningPort = typeof address === "object" && address !== null ? address.port : port;
  console.log(`DocViewKit Viewer: http://${displayHost}:${listeningPort}/`);
});
