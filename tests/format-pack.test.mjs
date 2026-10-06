import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";

import {
  detectFormatPackCandidate,
  FormatPackRuntime,
} from "../dist/format-pack.js";
import { OfficeEngineError } from "../dist/index.js";
import { createZip } from "./zip-fixture.mjs";

const iwork = createZip({
  "Metadata/Properties.plist": new TextEncoder().encode("plist"),
  "Index/Document.iwa": Uint8Array.of(0, 1, 2, 3),
});
const odf = createZip({
  mimetype: new TextEncoder().encode("application/vnd.oasis.opendocument.text"),
  "META-INF/manifest.xml": new TextEncoder().encode("<manifest:manifest/>"),
  "content.xml": new TextEncoder().encode("<office:document-content/>"),
}, { compress: false });
const fodp = new TextEncoder().encode(`<?xml version="1.0"?>
  <office:document office:mimetype="application/vnd.oasis.opendocument.presentation"/>`);
const fodg = new TextEncoder().encode(`<?xml version="1.0"?>
  <office:document office:mimetype="application/vnd.oasis.opendocument.graphics"/>`);
const fodt = new TextEncoder().encode(`<?xml version="1.0"?>
  <office:document office:mimetype="application/vnd.oasis.opendocument.text"/>`);
const fods = new TextEncoder().encode(`<?xml version="1.0"?>
  <office:document office:mimetype="application/vnd.oasis.opendocument.spreadsheet"/>`);
const appleIworkZip = createZip({
  "Metadata/Properties.plist": new TextEncoder().encode("plist"),
  "Index/Document.iwa": Uint8Array.of(0, 1, 2, 3),
}, { redundantLocalZip64: true });
const genericZip = createZip({
  "notes/readme.txt": new TextEncoder().encode("Index/Document.iwa Metadata/Properties.plist"),
});
const iworkDirectoryPackage = createZip({
  "Report.pages/Index.zip": Uint8Array.of(0),
  "Report.pages/Metadata/Properties.plist": new TextEncoder().encode("plist"),
});
const legacyOffice = Uint8Array.of(
  0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1,
  0, 0, 0, 0,
);
const pdf = new TextEncoder().encode("%PDF-1.7\n%%EOF");
const prefixedPdf = new TextEncoder().encode("%!PS-compatible\n%PDF-1.7\n%%EOF");
const headerlessPdf = Uint8Array.from([
  0x25, 0xe2, 0xe3, 0xcf, 0xd3, 0x0a,
  ...new TextEncoder().encode("1 0 obj << /Type /Catalog >> endobj\ntrailer << /Root 1 0 R >>"),
]);
const xps = createZip({
  "FixedDocSeq.fdseq": new TextEncoder().encode("<FixedDocumentSequence/>") ,
  "Documents/1/FixedDoc.fdoc": new TextEncoder().encode("<FixedDocument/>") ,
  "Documents/1/Pages/1.fpage": new TextEncoder().encode("<FixedPage/>") ,
});
const ofd = createZip({
  "OFD.xml": new TextEncoder().encode("<ofd:OFD xmlns:ofd=\"http://www.ofdspec.org/2016\"><ofd:DocBody><ofd:DocRoot>Doc_0/Document.xml</ofd:DocRoot></ofd:DocBody></ofd:OFD>"),
  "Doc_0/Document.xml": new TextEncoder().encode("<ofd:Document xmlns:ofd=\"http://www.ofdspec.org/2016\"><ofd:Pages><ofd:Page ID=\"1\" BaseLoc=\"Pages/Page_0/Content.xml\"/></ofd:Pages></ofd:Document>"),
  "Doc_0/Pages/Page_0/Content.xml": new TextEncoder().encode("<ofd:Page xmlns:ofd=\"http://www.ofdspec.org/2016\"/>") ,
});

function deferred() {
  let resolve;
  const promise = new Promise((done) => {
    resolve = done;
  });
  return { promise, resolve };
}

function unsupportedNative(overrides = {}) {
  return {
    async open() {
      throw new OfficeEngineError("UNSUPPORTED_FORMAT", "unsupported");
    },
    close() {},
    ...overrides,
  };
}

function inertPackEngine(open) {
  return { open, close() {} };
}

test("opens recognized iWork bytes with a lazily loaded native format pack", async () => {
  const calls = [];
  const document = { marker: "native-iwork-document" };
  const nativeEngine = {
    async open(bytes) {
      calls.push(["native-open", bytes]);
      throw new OfficeEngineError("UNSUPPORTED_FORMAT", "unsupported");
    },
    close() {},
  };
  const runtime = new FormatPackRuntime(
    nativeEngine,
    async () => ({
      async load(candidate) {
        calls.push(["load", candidate]);
        return "iwork-wasm";
      },
    }),
    async (candidate, wasm) => {
      calls.push(["create", candidate, wasm]);
      return {
        async open(bytes) {
          calls.push(["pack-open", bytes]);
          return document;
        },
        close() {},
      };
    },
  );

  assert.equal(await runtime.open(iwork), document);
  assert.deepEqual(calls, [
    ["native-open", iwork],
    ["load", "iwork"],
    ["create", "iwork", "iwork-wasm"],
    ["pack-open", iwork],
  ]);
  runtime.close();
});

test("opens recognized ODF bytes only after the base engine rejects them", async () => {
  const calls = [];
  const document = { marker: "odf-document" };
  const runtime = new FormatPackRuntime(
    unsupportedNative({ async open() { calls.push("base"); throw new OfficeEngineError("UNSUPPORTED_FORMAT", "unsupported"); } }),
    async () => ({
      async load(candidate) {
        calls.push(`load:${candidate}`);
        return "odf-wasm";
      },
    }),
    async (candidate, wasm) => {
      calls.push(`create:${candidate}:${wasm}`);
      return inertPackEngine(async () => document);
    },
  );

  assert.equal(await runtime.open(odf), document);
  assert.deepEqual(calls, ["base", "load:odf", "create:odf:odf-wasm"]);
  runtime.close();
});

test("routes flat ODF XML before native CSV sniffing", async () => {
  const calls = [];
  const document = { marker: "flat-odf-document" };
  const runtime = new FormatPackRuntime(
    unsupportedNative({
      async open() {
        calls.push("base");
        throw new OfficeEngineError("FORMAT_INVALID", "a CSV quote must begin at the start of a field");
      },
    }),
    async () => ({
      async load(candidate) {
        calls.push(`load:${candidate}`);
        return "odf-wasm";
      },
    }),
    async () => inertPackEngine(async () => document),
  );

  assert.equal(await runtime.open(fodp, { fileName: "fit-to-size.fodp" }), document);
  assert.equal(await runtime.open(fods, { fileName: "sheet.fods" }), document);
  assert.deepEqual(calls, ["load:odf"]);
  runtime.close();
});

test("routes flat ODT XML before native flat-document sniffing", async () => {
  const calls = [];
  const document = { marker: "flat-odt-document" };
  const runtime = new FormatPackRuntime(
    unsupportedNative({
      async open() {
        calls.push("base");
        throw new OfficeEngineError("UNSUPPORTED_FORMAT", "unsupported");
      },
    }),
    async () => ({
      async load(candidate) {
        calls.push(`load:${candidate}`);
        return "odf-wasm";
      },
    }),
    async () => inertPackEngine(async () => document),
  );

  assert.equal(await runtime.open(fodt, { fileName: "document.fodt" }), document);
  assert.deepEqual(calls, ["load:odf"]);
  runtime.close();
});

test("native and pack parsers share an immutable byte snapshot across fallback", async () => {
  const mutable = iwork.slice();
  let nativeBytes;
  let packBytes;
  const document = { marker: "stable-input" };
  const runtime = new FormatPackRuntime(
    {
      async open(bytes) {
        nativeBytes = bytes;
        await Promise.resolve();
        throw new OfficeEngineError("UNSUPPORTED_FORMAT", "unsupported");
      },
      close() {},
    },
    async () => ({ async load() { return "iwork-wasm"; } }),
    async () => inertPackEngine(async (bytes) => {
      packBytes = bytes;
      return document;
    }),
  );

  const opening = runtime.open(mutable);
  mutable.fill(0);
  assert.equal(await opening, document);
  assert.notEqual(nativeBytes, mutable);
  assert.equal(packBytes, nativeBytes);
  assert.equal(detectFormatPackCandidate(packBytes), "iwork");
  runtime.close();
});

test("extended format pack shares one iWork/PDF Wasm asset", async () => {
  const { extendedFormatPack } = await import("../dist/extended-formats.js");
  const odfWasm = await extendedFormatPack.load("odf");
  const iworkWasm = await extendedFormatPack.load("iwork");
  const legacyWasm = await extendedFormatPack.load("legacy-office");
  const wpsWasm = await extendedFormatPack.load("wps");
  const pdfWasm = await extendedFormatPack.load("pdf");
  const xpsWasm = await extendedFormatPack.load("xps");
  const ofdWasm = await extendedFormatPack.load("ofd");

  assert.equal(odfWasm instanceof URL, true);
  assert.equal(iworkWasm instanceof URL, true);
  assert.equal(legacyWasm instanceof URL, true);
  assert.equal(wpsWasm instanceof URL, true);
  assert.equal(odfWasm.pathname.endsWith("/office-viewer-odf.wasm"), true);
  assert.equal(iworkWasm.pathname.endsWith("/office-viewer-pdf.wasm"), true);
  assert.equal(legacyWasm.pathname.endsWith("/office-viewer-legacy-office.wasm"), true);
  assert.equal(wpsWasm.pathname.endsWith("/office-viewer-legacy-office.wasm"), true);
  assert.equal(pdfWasm.pathname.endsWith("/office-viewer-pdf.wasm"), true);
  assert.equal(iworkWasm.href, pdfWasm.href);
  assert.equal(xpsWasm.pathname.endsWith("/office-viewer-xps.wasm"), true);
  assert.equal(ofdWasm.pathname.endsWith("/office-viewer-ofd.wasm"), true);
});

test("standalone XPS pack exposes only the XPS Wasm", async () => {
  const { xpsFormatPack } = await import("../dist/xps-formats.js");
  const xpsWasm = await xpsFormatPack.load("xps");
  assert.equal(xpsWasm instanceof URL, true);
  assert.equal(xpsWasm.pathname.endsWith("/office-viewer-xps.wasm"), true);
  await assert.rejects(
    xpsFormatPack.load("pdf"),
    (cause) => cause?.code === "UNSUPPORTED_FORMAT",
  );
});

test("standalone OFD pack exposes only the OFD Wasm", async () => {
  const { ofdFormatPack } = await import("../dist/ofd-formats.js");
  const ofdWasm = await ofdFormatPack.load("ofd");
  assert.equal(ofdWasm instanceof URL, true);
  assert.equal(ofdWasm.pathname.endsWith("/office-viewer-ofd.wasm"), true);
  await assert.rejects(
    ofdFormatPack.load("pdf"),
    (cause) => cause?.code === "UNSUPPORTED_FORMAT",
  );
});

test("standalone WPS pack exposes only the shared audited CFB Wasm", async () => {
  const { wpsFormatPack } = await import("../dist/wps-formats.js");
  const wpsWasm = await wpsFormatPack.load("wps");
  assert.equal(wpsWasm instanceof URL, true);
  assert.equal(wpsWasm.pathname.endsWith("/office-viewer-legacy-office.wasm"), true);
  await assert.rejects(
    wpsFormatPack.load("legacy-office"),
    (cause) => cause?.code === "UNSUPPORTED_FORMAT",
  );
});

test("recognizes the supplied Foxit OFD with a directory-first ZIP layout", async () => {
  const bytes = await readFile(new URL("./fixtures/foxit-directory-first.ofd", import.meta.url));
  assert.equal(bytes.readUInt32LE(0), 0x06054b50);
  assert.equal(detectFormatPackCandidate(bytes), "ofd");
  assert.equal(detectFormatPackCandidate(bytes.subarray(0, 22)), undefined);
  const invalid = Buffer.from(bytes);
  invalid.writeUInt32LE(0xffffffff, invalid.length - 6);
  assert.equal(detectFormatPackCandidate(invalid), undefined);
});

test("dispatches only structurally recognized optional format bytes", () => {
  assert.equal(detectFormatPackCandidate(odf), "odf");
  assert.equal(detectFormatPackCandidate(fodp), "odf");
  assert.equal(detectFormatPackCandidate(fodg), "odf");
  assert.equal(detectFormatPackCandidate(iwork), "iwork");
  assert.equal(detectFormatPackCandidate(legacyOffice), "legacy-office");
  assert.equal(detectFormatPackCandidate(legacyOffice, "report.wps"), "wps");
  assert.equal(detectFormatPackCandidate(legacyOffice, "budget.ET"), "wps");
  assert.equal(detectFormatPackCandidate(legacyOffice, "slides.dps"), "wps");
  assert.equal(detectFormatPackCandidate(legacyOffice, "slides.dpt"), "legacy-office");
  assert.equal(detectFormatPackCandidate(genericZip, "forged.wps"), undefined);
  assert.equal(detectFormatPackCandidate(pdf), "pdf");
  assert.equal(detectFormatPackCandidate(prefixedPdf), "pdf");
  assert.equal(detectFormatPackCandidate(headerlessPdf), "pdf");
  assert.equal(detectFormatPackCandidate(xps), "xps");
  assert.equal(detectFormatPackCandidate(ofd), "ofd");
  assert.equal(
    detectFormatPackCandidate(Uint8Array.of(0x25, 0xe2, 0xe3, 0xcf, 0xd3)),
    undefined,
  );
  assert.equal(detectFormatPackCandidate(appleIworkZip), "iwork");
  assert.equal(detectFormatPackCandidate(iworkDirectoryPackage), "iwork");
  assert.equal(detectFormatPackCandidate(genericZip), undefined);
  assert.equal(
    detectFormatPackCandidate(new TextEncoder().encode("Index/Document.iwa Metadata/Properties.plist")),
    undefined,
  );
});

test("routes structurally valid CFB bytes to WPS only with an explicit WPS filename", async () => {
  const calls = [];
  const document = { marker: "wps-document" };
  const runtime = new FormatPackRuntime(
    unsupportedNative(),
    async () => ({
      async load(candidate) {
        calls.push(["load", candidate]);
        return "wps-wasm";
      },
    }),
    async (candidate, wasm) => {
      calls.push(["create", candidate, wasm]);
      return inertPackEngine(async (_bytes, options) => {
        calls.push(["open", options.fileName]);
        return document;
      });
    },
  );

  assert.equal(await runtime.open(legacyOffice, { fileName: "report.wps" }), document);
  assert.deepEqual(calls, [
    ["load", "wps"],
    ["create", "wps", "wps-wasm"],
    ["open", "report.wps"],
  ]);
  runtime.close();
});

test("routes Apple's redundant local ZIP64 pair to the strict iWork pack", async () => {
  const document = { marker: "apple-iwork" };
  let loads = 0;
  const runtime = new FormatPackRuntime(
    {
      async open() {
        throw new OfficeEngineError("PACKAGE_ZIP64_UNSUPPORTED", "base ZIP policy rejected it");
      },
      close() {},
    },
    async () => ({
      async load(candidate) {
        assert.equal(candidate, "iwork");
        loads += 1;
        return "iwork-wasm";
      },
    }),
    async () => inertPackEngine(async (bytes) => {
      assert.equal(detectFormatPackCandidate(bytes), "iwork");
      return document;
    }),
  );

  assert.equal(await runtime.open(appleIworkZip), document);
  assert.equal(loads, 1);
  runtime.close();
});

test("routes Apple's valid unflagged UTF-8 paths only to the strict iWork pack", async () => {
  const document = { marker: "apple-iwork-unflagged-utf8" };
  let loads = 0;
  const runtime = new FormatPackRuntime(
    {
      async open() {
        throw new OfficeEngineError(
          "PACKAGE_ZIP_INVALID",
          "non-ASCII ZIP path requires the UTF-8 flag",
        );
      },
      close() {},
    },
    async () => ({
      async load(candidate) {
        assert.equal(candidate, "iwork");
        loads += 1;
        return "iwork-wasm";
      },
    }),
    async () => inertPackEngine(async () => document),
  );

  assert.equal(await runtime.open(iwork), document);
  assert.equal(loads, 1);
  runtime.close();
});

test("keeps native open first and does not import a pack on native success", async () => {
  const document = { marker: "base-document" };
  let imports = 0;
  let creates = 0;
  const runtime = new FormatPackRuntime(
    {
      async open(bytes) {
        assert.notEqual(bytes, iwork);
        assert.deepEqual(bytes, iwork);
        return document;
      },
      close() {},
    },
    async () => {
      imports += 1;
      throw new Error("must stay lazy");
    },
    async () => {
      creates += 1;
      throw new Error("must stay lazy");
    },
  );

  assert.equal(await runtime.open(iwork), document);
  assert.equal(imports, 0);
  assert.equal(creates, 0);
  runtime.close();
});

test("preserves the base error when packs are disabled or content is not a candidate", async () => {
  for (const [formatPack, bytes] of [[false, odf], [false, iwork], [undefined, genericZip]]) {
    const failure = new OfficeEngineError("UNSUPPORTED_FORMAT", "base failure");
    let creates = 0;
    const runtime = new FormatPackRuntime(
      unsupportedNative({ async open() { throw failure; } }),
      formatPack,
      async () => {
        creates += 1;
        throw new Error("must not create");
      },
    );
    await assert.rejects(runtime.open(bytes), (cause) => cause === failure);
    assert.equal(creates, 0);
    runtime.close();
  }
});

test("does not bypass base security and validation failures for recognized content", async () => {
  for (const [bytes, failure] of [
    [legacyOffice, new OfficeEngineError("INPUT_SIZE_LIMIT", "too large")],
    [iwork, new OfficeEngineError("PACKAGE_ZIP_INVALID", "bad CRC")],
  ]) {
    let imports = 0;
    const runtime = new FormatPackRuntime(
      unsupportedNative({ async open() { throw failure; } }),
      async () => {
        imports += 1;
        throw new Error("must not import");
      },
      async () => { throw new Error("must not create"); },
    );

    await assert.rejects(runtime.open(bytes), (cause) => cause === failure);
    assert.equal(imports, 0);
    runtime.close();
  }
});

test("single-flights one pack engine per candidate and one pack import per Engine", async () => {
  let imports = 0;
  const loads = [];
  const creates = [];
  const packOpens = [];
  const runtime = new FormatPackRuntime(
    unsupportedNative(),
    async () => {
      imports += 1;
      return {
        async load(candidate) {
          loads.push(candidate);
          await Promise.resolve();
          return `${candidate}-wasm`;
        },
      };
    },
    async (candidate, wasm) => {
      creates.push([candidate, wasm]);
      await Promise.resolve();
      return inertPackEngine(async (bytes) => {
        packOpens.push([candidate, bytes]);
        return { candidate };
      });
    },
  );

  const [first, second, legacy, third] = await Promise.all([
    runtime.open(iwork),
    runtime.open(iwork),
    runtime.open(legacyOffice),
    runtime.open(iwork),
  ]);
  assert.deepEqual([first.candidate, second.candidate, legacy.candidate, third.candidate], [
    "iwork",
    "iwork",
    "legacy-office",
    "iwork",
  ]);
  assert.equal(imports, 1);
  assert.deepEqual(loads.toSorted(), ["iwork", "legacy-office"]);
  assert.deepEqual(creates.toSorted(([left], [right]) => left.localeCompare(right)), [
    ["iwork", "iwork-wasm"],
    ["legacy-office", "legacy-office-wasm"],
  ]);
  assert.equal(packOpens.length, 4);
  assert.deepEqual(
    packOpens.map(([, bytes]) => detectFormatPackCandidate(bytes)).toSorted(),
    ["iwork", "iwork", "iwork", "legacy-office"],
  );
  runtime.close();
});

test("retries a candidate after a transient pack load failure", async () => {
  let loads = 0;
  const document = { marker: "recovered" };
  const runtime = new FormatPackRuntime(
    unsupportedNative(),
    async () => ({
      async load() {
        loads += 1;
        if (loads === 1) throw new Error("temporary load failure");
        return "iwork-wasm";
      },
    }),
    async () => inertPackEngine(async () => document),
  );

  await assert.rejects(runtime.open(iwork), /temporary load failure/u);
  assert.equal(await runtime.open(iwork), document);
  assert.equal(loads, 2);
  runtime.close();
});

test("retries the lazy pack factory after a transient import failure", async () => {
  let imports = 0;
  const document = { marker: "recovered-import" };
  const runtime = new FormatPackRuntime(
    unsupportedNative(),
    async () => {
      imports += 1;
      if (imports === 1) throw new Error("temporary import failure");
      return { async load() { return "iwork-wasm"; } };
    },
    async () => inertPackEngine(async () => document),
  );

  await assert.rejects(runtime.open(iwork), /temporary import failure/u);
  assert.equal(await runtime.open(iwork), document);
  assert.equal(imports, 2);
  runtime.close();
});

test("includes lazy pack loading in the open timeout without restarting its budget", async () => {
  const timeoutMs = 30;
  const never = new Promise(() => {});
  const runtime = new FormatPackRuntime(
    unsupportedNative(),
    async () => ({ async load() { return never; } }),
    async () => { throw new Error("must not create"); },
  );

  await assert.rejects(
    runtime.open(iwork, { timeoutMs }),
    (cause) => cause?.code === "OPERATION_TIMEOUT" && cause.message.includes(`${timeoutMs}ms`),
  );
  runtime.close();
});

test("passes the remaining default budget to a pack Worker but not inline execution", async () => {
  const seen = [];
  for (const execution of ["worker", "inline"]) {
    const runtime = new FormatPackRuntime(
      unsupportedNative(),
      async () => ({ async load() { return "iwork-wasm"; } }),
      async () => inertPackEngine(async (_bytes, options) => {
        seen.push([execution, options.timeoutMs]);
        return { execution };
      }),
      undefined,
      execution,
    );
    await runtime.open(iwork);
    runtime.close();
  }

  assert.equal(seen[0][0], "worker");
  assert.equal(Number.isFinite(seen[0][1]) && seen[0][1] > 0 && seen[0][1] <= 60_000, true);
  assert.deepEqual(seen[1], ["inline", undefined]);
});

test("does not return a late document when its promise wins the timer-task race", async () => {
  const opened = deferred();
  const started = deferred();
  let documentCloses = 0;
  const runtime = new FormatPackRuntime(
    unsupportedNative(),
    async () => ({ async load() { return "iwork-wasm"; } }),
    async () => inertPackEngine(async () => {
      started.resolve();
      return opened.promise;
    }),
  );

  const opening = runtime.open(iwork, { timeoutMs: 5 });
  await started.promise;
  const blockedUntil = performance.now() + 10;
  while (performance.now() < blockedUntil) {
    // Keep the timer task queued so the resolved promise exercises the absolute-deadline check.
  }
  opened.resolve({ close() { documentCloses += 1; } });
  await assert.rejects(opening, (cause) => cause?.code === "OPERATION_TIMEOUT");
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(documentCloses, 1);
  runtime.close();
});

test("cancels while the optional pack is loading", async () => {
  const loading = deferred();
  const started = deferred();
  const controller = new AbortController();
  const runtime = new FormatPackRuntime(
    unsupportedNative(),
    async () => ({
      async load() {
        started.resolve();
        return loading.promise;
      },
    }),
    async () => { throw new Error("must not create"); },
  );

  const opening = runtime.open(iwork, { signal: controller.signal });
  await started.promise;
  controller.abort("test cancellation");
  await assert.rejects(
    opening,
    (cause) => cause?.code === "OPERATION_ABORTED" && cause.cause === "test cancellation",
  );
  runtime.close();
  loading.resolve("unused-wasm");
});

test("close rejects pending opens and closes a pack engine that finishes creating late", async () => {
  const creation = deferred();
  const started = deferred();
  let nativeCloses = 0;
  let packCloses = 0;
  const runtime = new FormatPackRuntime(
    unsupportedNative({ close() { nativeCloses += 1; } }),
    async () => ({ async load() { return "iwork-wasm"; } }),
    async () => {
      started.resolve();
      return creation.promise;
    },
  );

  const opening = runtime.open(iwork);
  await started.promise;
  runtime.close();
  await assert.rejects(opening, (cause) => cause?.code === "ENGINE_CLOSED");
  creation.resolve({
    async open() { return { close() {} }; },
    close() { packCloses += 1; },
  });
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(nativeCloses, 1);
  assert.equal(packCloses, 1);
  runtime.close();
  assert.equal(nativeCloses, 1);
});

test("close during the lazy factory prevents a later pack load side effect", async () => {
  const factory = deferred();
  const started = deferred();
  let loads = 0;
  const runtime = new FormatPackRuntime(
    unsupportedNative(),
    async () => {
      started.resolve();
      return factory.promise;
    },
    async () => { throw new Error("must not create"); },
  );

  const opening = runtime.open(iwork);
  await started.promise;
  runtime.close();
  factory.resolve({ async load() { loads += 1; return "unused"; } });
  await assert.rejects(opening, (cause) => cause?.code === "ENGINE_CLOSED");
  assert.equal(loads, 0);
});

// The ZIP detector only routes a candidate; the XPS parser validates the OPC relationships.
test("XPS candidates include interleaved physical pieces and legacy XAML parts", () => {
  for (const entries of [
    { "Sequence.fdseq/[0].last.piece": new TextEncoder().encode("sequence"),
      "Document.fdoc/[0].last.piece": new TextEncoder().encode("document"),
      "Pages/1.fpage/[0].last.piece": new TextEncoder().encode("page") },
    { "FixedDocSeq.xaml": new TextEncoder().encode("sequence"),
      "FixedDoc.xaml": new TextEncoder().encode("document"), "Pages/1.xaml": new TextEncoder().encode("page") },
  ]) assert.equal(detectFormatPackCandidate(createZip(entries)), "xps");
});

test("routes an OPC XPS document with a missing sequence to the validating parser", () => {
  const entries = {
    "[Content_Types].xml": new Uint8Array(), "_rels/.rels": new Uint8Array(),
    "Documents/1/FixedDoc.fdoc": new Uint8Array(), "Documents/1/Pages/1.fpage": new Uint8Array(),
  };
  assert.equal(detectFormatPackCandidate(createZip(entries)), "xps");
  delete entries["_rels/.rels"];
  assert.equal(detectFormatPackCandidate(createZip(entries)), undefined);
});

test("routes an OPC XPS sequence with a missing document index to the validating parser", () => {
  const entries = {
    "[Content_Types].xml": new Uint8Array(), "_rels/.rels": new Uint8Array(),
    "FixedDocSeq.fdseq": new Uint8Array(), "Documents/1/Pages/1.fpage": new Uint8Array(),
  };
  assert.equal(detectFormatPackCandidate(createZip(entries)), "xps");
  delete entries["_rels/.rels"];
  assert.equal(detectFormatPackCandidate(createZip(entries)), undefined);
});
