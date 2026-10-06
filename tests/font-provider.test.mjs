import assert from "node:assert/strict";
import test from "node:test";

import {
  collectFontFamilies,
  collectFontRequests,
  documentFontRuns,
  FontProviderCache,
} from "../dist/font.js";

test("font provider batches normalized face demands and reuses a full-face cache entry", async () => {
  const calls = [];
  const cache = new FontProviderCache(async (requests, context) => {
    calls.push({ requests, context });
    return [{ family: "Demand Sans", bytes: Uint8Array.of(1, 2, 3) }];
  }, {
    maxBytes: 64,
    policy: "deterministic",
    timeoutMs: 50,
  });

  const first = await cache.resolve([
    {
      family: "  Demand Sans  ",
      style: "normal",
      weight: 400,
      stretch: "normal",
      codePoints: [0x4e2d, 0x41, 0x41],
      language: "zh-cn",
      script: "hans",
    },
    {
      family: "Demand Sans",
      style: "normal",
      weight: 400,
      stretch: "normal",
      codePoints: [0x42],
      language: "zh-CN",
      script: "Hans",
    },
  ]);

  assert.equal(calls.length, 1);
  assert.deepEqual(calls[0].requests, [{
    family: "Demand Sans",
    style: "normal",
    weight: 400,
    stretch: "normal",
    codePoints: [0x41, 0x42, 0x4e2d],
    language: "zh-CN",
    script: "Hans",
  }]);
  assert.equal(Object.isFrozen(calls[0].requests), true);
  assert.equal(Object.isFrozen(calls[0].requests[0]), true);
  assert.equal(Object.isFrozen(calls[0].requests[0].codePoints), true);
  assert.equal(calls[0].context.policy, "deterministic");
  assert.equal(calls[0].context.signal instanceof AbortSignal, true);
  assert.deepEqual(first.diagnostics, []);
  assert.deepEqual([...new Uint8Array(first.assets[0].bytes)], [1, 2, 3]);

  new Uint8Array(first.assets[0].bytes)[0] = 99;
  const cached = await cache.resolve([{
    family: "Demand Sans",
    style: "normal",
    weight: 400,
    stretch: "normal",
    codePoints: [0x43],
  }]);

  assert.equal(calls.length, 1);
  assert.deepEqual(cached.diagnostics, []);
  assert.deepEqual([...new Uint8Array(cached.assets[0].bytes)], [1, 2, 3]);
});

test("font provider verifies expected SHA-256 and budgets duplicate content once", async () => {
  const bytes = Uint8Array.of(1, 2, 3);
  const cache = new FontProviderCache(async () => [
    { family: "Content Sans", bytes },
    { family: "Content Serif", bytes },
  ], {
    maxBytes: 6,
    timeoutMs: 50,
  });

  const result = await cache.resolve([
    {
      family: "Content Sans",
      style: "normal",
      weight: 400,
      stretch: "normal",
      codePoints: [0x41],
      expectedSha256: "039058C6F2C0CB492C533B0A4D14EF77CC0F78ABCCCED5287D84A1A2011CFB81",
    },
    {
      family: "Content Serif",
      style: "normal",
      weight: 400,
      stretch: "normal",
      codePoints: [0x42],
    },
  ]);

  assert.deepEqual(result.diagnostics, []);
  assert.deepEqual(result.assets.map(({ family }) => family), ["Content Sans", "Content Serif"]);
});

test("font provider rejects an integrity mismatch with a request-scoped diagnostic", async () => {
  const cache = new FontProviderCache(async () => [
    { family: "Pinned Sans", bytes: Uint8Array.of(1, 2, 3) },
  ], {
    maxBytes: 64,
    timeoutMs: 50,
  });
  const expectedSha256 = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";

  const result = await cache.resolve([{
    family: "Pinned Sans",
    style: "normal",
    weight: 400,
    stretch: "normal",
    codePoints: [0x41],
    expectedSha256,
  }]);

  assert.deepEqual(result.assets, []);
  assert.equal(result.diagnostics.length, 1);
  assert.deepEqual(result.diagnostics[0], {
    code: "FONT_INTEGRITY_MISMATCH",
    severity: "error",
    fidelity: "not-rendered",
    phase: "security",
    message: "Font provider returned an unexpected binary for Pinned Sans",
    details: {
      requestKey: `pinned sans;normal;400;normal;sha256=${expectedSha256}`,
      family: "Pinned Sans",
      expectedSha256,
      actualSha256: "039058c6f2c0cb492c533b0a4d14ef77cc0f78abccced5287d84a1a2011cfb81",
    },
  });
  assert.equal(Object.isFrozen(result.diagnostics), true);
  assert.equal(Object.isFrozen(result.diagnostics[0]), true);
  assert.equal(Object.isFrozen(result.diagnostics[0].details), true);
});

test("font provider verifies the manifest digest returned with a font asset", async () => {
  const expectedSha256 = "bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb";
  const cache = new FontProviderCache(async () => [{
    family: "Manifest Sans",
    bytes: Uint8Array.of(1, 2, 3),
    sha256: expectedSha256,
  }], { maxBytes: 64, timeoutMs: 50 });

  const result = await cache.resolve([{
    family: "Manifest Sans",
    style: "normal",
    weight: 400,
    stretch: "normal",
    codePoints: [0x41],
  }]);

  assert.deepEqual(result.assets, []);
  assert.equal(result.diagnostics[0].code, "FONT_INTEGRITY_MISMATCH");
  assert.equal(result.diagnostics[0].details.expectedSha256, expectedSha256);
});

test("font provider rejects an over-budget batch before copying and hashing it", async () => {
  const cache = new FontProviderCache(async () => [
    { family: "Batch One", bytes: Uint8Array.of(1, 2) },
    { family: "Batch Two", bytes: Uint8Array.of(3, 4) },
  ], { maxBytes: 3, timeoutMs: 50 });
  const requests = ["Batch One", "Batch Two"].map((family) => ({
    family,
    style: "normal",
    weight: 400,
    stretch: "normal",
    codePoints: [0x41],
  }));

  const result = await cache.resolve(requests);

  assert.deepEqual(result.assets, []);
  assert.equal(result.diagnostics.length, 2);
  assert.ok(result.diagnostics.every(({ code }) => code === "FONT_PROVIDER_INVALID_RESPONSE"));
  assert.ok(result.diagnostics.every(({ details }) => /cache budget/u.test(details.error)));
});

test("font provider reports every requested face it did not return", async () => {
  const cache = new FontProviderCache(async () => [], {
    maxBytes: 64,
    timeoutMs: 50,
    policy: "local-first",
  });

  const result = await cache.resolve([{
    family: "Missing Serif",
    style: "italic",
    weight: 600,
    stretch: "condensed",
    codePoints: [0x41, 0x42],
    language: "en-US",
    script: "Latn",
  }]);

  assert.deepEqual(result.assets, []);
  assert.deepEqual(result.diagnostics, [{
    code: "FONT_PROVIDER_MISSING",
    severity: "warning",
    fidelity: "approximate",
    phase: "layout",
    message: "Font provider did not return Missing Serif italic 600 condensed",
    details: {
      requestKey: "missing serif;italic;600;condensed;sha256=any",
      family: "Missing Serif",
      style: "italic",
      weight: 600,
      stretch: "condensed",
      codePointCount: 2,
      language: "en-US",
      script: "Latn",
    },
  }]);
});

test("font provider failures become retryable request-scoped diagnostics", async () => {
  let calls = 0;
  const cache = new FontProviderCache(async () => {
    calls += 1;
    throw new Error("service unavailable");
  }, {
    maxBytes: 64,
    timeoutMs: 50,
  });
  const request = {
    family: "Retry Sans",
    style: "normal",
    weight: 400,
    stretch: "normal",
    codePoints: [0x41],
  };

  const first = await cache.resolve([request]);
  const second = await cache.resolve([request]);

  assert.equal(calls, 2);
  for (const result of [first, second]) {
    assert.deepEqual(result.assets, []);
    assert.deepEqual(result.diagnostics, [{
      code: "FONT_PROVIDER_FAILED",
      severity: "warning",
      fidelity: "approximate",
      phase: "layout",
      message: "Font provider failed while resolving Retry Sans",
      details: {
        requestKey: "retry sans;normal;400;normal;sha256=any",
        family: "Retry Sans",
        error: "service unavailable",
      },
    }]);
  }
});

test("font requests collect actual Unicode scalars and keep font faces distinct", () => {
  const base = {
    numericId: 1,
    id: "object:1",
    type: "paragraph",
    unitIndex: 0,
    bounds: { x: 0, y: 0, width: 100, height: 20 },
    source: {
      format: "odt",
      part: "content.xml",
      kind: "text-range",
      path: "/p[1]",
      mapping: "exact",
    },
    z: 0,
  };
  const requests = collectFontRequests([
    {
      ...base,
      text: "A😀A",
      visual: {
        kind: "text",
        geometry: "rectangle",
        fill: 0,
        stroke: 0,
        strokeWidth: 0,
        fontFamily: "Demand Sans",
        fontSize: 12,
        color: 0,
        bold: false,
        italic: false,
        align: "start",
      },
    },
    {
      ...base,
      numericId: 2,
      id: "object:2",
      text: "BA",
      visual: {
        kind: "text-layout",
        layout: {},
        visual: {
          kind: "rich-text",
          geometry: "rectangle",
          fill: { kind: "none" },
          stroke: { kind: "none" },
          strokeWidth: 0,
          align: "start",
          lineHeight: 14,
          runs: [{
            text: "BA",
            fontFamily: "Demand Sans",
            fontSize: 12,
            color: 0,
            bold: true,
            italic: true,
            underline: false,
            strikethrough: false,
            highlight: 0,
            baselineShift: 0,
            letterSpacing: 0,
          }],
        },
      },
    },
  ]);

  assert.deepEqual(requests, [
    {
      family: "Demand Sans",
      style: "normal",
      weight: 400,
      stretch: "normal",
      codePoints: [0x41, 0x1f600],
    },
    {
      family: "Demand Sans",
      style: "italic",
      weight: 700,
      stretch: "normal",
      codePoints: [0x41, 0x42],
    },
  ]);
});

test("font discovery traverses groups and DrawingML wrappers", () => {
  const object = {
    numericId: 1,
    id: "object:wrapped-font",
    type: "paragraph",
    unitIndex: 0,
    bounds: { x: 0, y: 0, width: 100, height: 20 },
    source: {
      format: "pptx",
      part: "ppt/slides/slide1.xml",
      kind: "text-range",
      path: "/p:sp[1]/p:txBody[1]",
      mapping: "exact",
    },
    z: 0,
    text: "Wrapped",
    visual: {
      kind: "stroke-style",
      style: {
        cap: "flat",
        join: "miter",
        compound: "single",
        alignment: "center",
        miterLimit: 4,
        dash: [],
      },
      visual: {
        kind: "advanced-effect",
        visual: {
          kind: "image-color-change",
          from: 0xffffffff,
          to: 0x00000000,
          useAlpha: true,
          visual: {
            kind: "group",
            children: [{
              bounds: { x: 0, y: 0, width: 100, height: 20 },
              visual: {
                kind: "rich-text",
                geometry: "rectangle",
                fill: { kind: "none" },
                stroke: { kind: "none" },
                strokeWidth: 0,
                align: "start",
                lineHeight: 14,
                runs: [{
                  text: "Wrapped",
                  fontFamily: "Wrapped Sans",
                  fontSize: 12,
                  color: 0,
                  bold: true,
                  italic: true,
                  underline: false,
                  strikethrough: false,
                  highlight: 0,
                  baselineShift: 0,
                  letterSpacing: 0,
                }],
              },
            }],
          },
        },
      },
    },
  };

  assert.deepEqual(collectFontFamilies([object]), ["Wrapped Sans"]);
  assert.deepEqual(collectFontRequests([object]), [{
    family: "Wrapped Sans",
    style: "italic",
    weight: 700,
    stretch: "normal",
    codePoints: [0x57, 0x61, 0x64, 0x65, 0x70, 0x72],
  }]);
  assert.deepEqual(documentFontRuns(object, {
    resolve: (family) => ({ family: `${family} Resolved`, source: "host" }),
  }), [{
    start: 0,
    end: 7,
    authoredFamily: "Wrapped Sans",
    renderedFamily: "Wrapped Sans Resolved",
    source: "host",
  }]);
});

test("font request limits count unique scalars rather than repeated characters", () => {
  const object = {
    numericId: 1,
    id: "long-run",
    type: "paragraph",
    unitIndex: 0,
    bounds: { x: 0, y: 0, width: 100, height: 20 },
    source: { format: "docx", part: "word/document.xml", kind: "paragraph", mapping: "derived" },
    z: 0,
    text: "A".repeat(65_537),
    visual: {
      kind: "text",
      geometry: "rectangle",
      fill: 0,
      stroke: 0,
      strokeWidth: 0,
      fontFamily: "Repeated Sans",
      fontSize: 12,
      color: 0,
      bold: false,
      italic: false,
      align: "start",
    },
  };

  assert.deepEqual(collectFontRequests([object])[0].codePoints, [0x41]);
});

test("font provider deduplicates concurrent requests for the same face", async () => {
  let calls = 0;
  let release;
  const cache = new FontProviderCache(async () => {
    calls += 1;
    await new Promise((resolve) => { release = resolve; });
    return [{ family: "Concurrent Sans", bytes: Uint8Array.of(7) }];
  }, { maxBytes: 64, timeoutMs: 100 });
  const face = {
    family: "Concurrent Sans",
    style: "normal",
    weight: 400,
    stretch: "normal",
  };

  const first = cache.resolve([{ ...face, codePoints: [0x41] }]);
  const second = cache.resolve([{ ...face, codePoints: [0x42] }]);
  await Promise.resolve();
  assert.equal(calls, 1);
  release();

  const results = await Promise.all([first, second]);
  assert.deepEqual(results.map(({ assets }) => assets.length), [1, 1]);
});

test("font provider timeout aborts the provider context and returns a diagnostic", async () => {
  let providerSignal;
  const cache = new FontProviderCache(async (_requests, context) => {
    providerSignal = context.signal;
    await new Promise((_, reject) => {
      context.signal.addEventListener("abort", () => reject(new Error("provider aborted")), { once: true });
    });
    return [];
  }, { maxBytes: 64, timeoutMs: 5 });

  const result = await cache.resolve([{
    family: "Slow Sans",
    style: "normal",
    weight: 400,
    stretch: "normal",
    codePoints: [0x41],
  }]);

  assert.equal(providerSignal.aborted, true);
  assert.equal(result.diagnostics[0].code, "FONT_PROVIDER_TIMEOUT");
  assert.equal(result.diagnostics[0].details.requestKey, "slow sans;normal;400;normal;sha256=any");
});

test("a zero-byte provider cache reports capacity without invoking the provider", async () => {
  let calls = 0;
  const cache = new FontProviderCache(async () => {
    calls += 1;
    return [];
  }, { maxBytes: 0 });

  const result = await cache.resolve([{
    family: "No Budget Sans",
    style: "normal",
    weight: 400,
    stretch: "normal",
    codePoints: [0x41],
  }]);

  assert.equal(calls, 0);
  assert.deepEqual(result.assets, []);
  assert.equal(result.diagnostics[0].code, "FONT_PROVIDER_CACHE_LIMIT");
  assert.equal(result.diagnostics[0].details.cacheLimit, 0);
  assert.equal(result.diagnostics[0].details.requestKey, "no budget sans;normal;400;normal;sha256=any");
});

test("closing the provider cache is terminal even when a provider ignores abort", async () => {
  let release;
  let started;
  const providerStarted = new Promise((resolve) => { started = resolve; });
  const cache = new FontProviderCache(async () => {
    started();
    await new Promise((resolve) => { release = resolve; });
    return [{ family: "Late Sans", bytes: Uint8Array.of(1) }];
  }, { maxBytes: 64, timeoutMs: 1_000 });
  const request = {
    family: "Late Sans",
    style: "normal",
    weight: 400,
    stretch: "normal",
    codePoints: [0x41],
  };

  const pending = cache.resolve([request]);
  await providerStarted;
  cache.close();

  await assert.rejects(pending, (error) => error?.code === "ENGINE_CLOSED");
  await assert.rejects(cache.resolve([request]), (error) => error?.code === "ENGINE_CLOSED");
  release();
});

test("closing before provider dispatch cancels the scheduled call", async () => {
  let calls = 0;
  const cache = new FontProviderCache(async () => {
    calls += 1;
    return [];
  }, { maxBytes: 64, timeoutMs: 1_000 });
  const pending = cache.resolve([{
    family: "Cancelled Sans",
    style: "normal",
    weight: 400,
    stretch: "normal",
    codePoints: [0x41],
  }]);

  cache.close();

  await assert.rejects(pending, (error) => error?.code === "ENGINE_CLOSED");
  await Promise.resolve();
  assert.equal(calls, 0);
});

test("one resolve never silently evicts an earlier face from its own provider batches", async () => {
  const cache = new FontProviderCache(async (requests) => requests.map(({ family }) => ({
    family,
    bytes: Uint8Array.of(Number(family.slice("Batch Face ".length))),
  })), { maxBytes: 128, timeoutMs: 1_000 });
  const requests = Array.from({ length: 129 }, (_, index) => ({
    family: `Batch Face ${index}`,
    style: "normal",
    weight: 400,
    stretch: "normal",
    codePoints: [0x41],
  }));

  const result = await cache.resolve(requests);

  assert.equal(result.assets.length, 128);
  assert.equal(result.diagnostics.length, 1);
  assert.equal(result.diagnostics[0].code, "FONT_PROVIDER_CACHE_LIMIT");
  assert.equal(result.assets.length + result.diagnostics.length, requests.length);
});

test("merging requests caps total Unicode scalar demand per exact face", async () => {
  let received;
  const cache = new FontProviderCache(async (requests) => {
    received = requests;
    return [];
  }, { maxBytes: 64, timeoutMs: 1_000 });
  const base = {
    family: "Scalar Sans",
    style: "normal",
    weight: 400,
    stretch: "normal",
  };

  await cache.resolve([
    { ...base, codePoints: Array.from({ length: 40_000 }, (_, index) => 0x10000 + index) },
    { ...base, codePoints: Array.from({ length: 40_000 }, (_, index) => 0x20000 + index) },
  ]);

  assert.equal(received.length, 1);
  assert.equal(received[0].codePoints.length, 65_536);
});
