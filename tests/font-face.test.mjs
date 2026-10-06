import assert from "node:assert/strict";
import test from "node:test";

import { registerFonts } from "../dist/font.js";

class FaceSet {
  constructor() {
    this.faces = new Set();
  }

  add(face) {
    this.faces.add(face);
  }

  delete(face) {
    return this.faces.delete(face);
  }

  check() {
    return true;
  }
}

class AvailableFace {
  constructor(family, source, descriptors) {
    this.family = family;
    this.source = source;
    this.descriptors = descriptors;
  }

  async load() {
    return this;
  }
}

function asset(family, byte, overrides = {}) {
  return {
    family,
    bytes: Uint8Array.of(byte).buffer,
    style: "normal",
    weight: 400,
    stretch: "normal",
    ...overrides,
  };
}

function request(family, overrides = {}) {
  return {
    family,
    style: "normal",
    weight: 400,
    stretch: "normal",
    codePoints: [0x41],
    ...overrides,
  };
}

test("exact faces keep their source while one authored family shares a CSS alias", async () => {
  const fontSet = new FaceSet();
  const registration = await registerFonts(
    [asset("Mixed Sans", 2, { style: "italic" })],
    50,
    { FontFace: AvailableFace, fontSet },
    undefined,
    [asset("Mixed Sans", 1)],
    ["Mixed Sans"],
    {
      providerAssets: [asset("Mixed Sans", 3, { weight: 700 })],
      policy: "deterministic",
      requestedFaces: [
        request("Mixed Sans"),
        request("Mixed Sans", { style: "italic" }),
        request("Mixed Sans", { weight: 700 }),
      ],
    },
  );

  const regular = registration.resolveFace(request("Mixed Sans"));
  const italic = registration.resolveFace(request("Mixed Sans", { style: "italic" }));
  const bold = registration.resolveFace(request("Mixed Sans", { weight: 700 }));

  assert.equal(regular.source, "embedded");
  assert.equal(italic.source, "host");
  assert.equal(bold.source, "provider");
  assert.equal(regular.family, italic.family);
  assert.equal(italic.family, bold.family);
  assert.match(regular.family, /^OfficeViewer \d+ embedded 0$/u);
  assert.deepEqual(
    [regular, italic, bold].map(({ style, weight, stretch }) => [style, weight, stretch]),
    [
      ["normal", 400, "normal"],
      ["italic", 400, "normal"],
      ["normal", 700, "normal"],
    ],
  );
  assert.equal([...fontSet.faces].every((face) => face.family === regular.family), true);
  assert.deepEqual(registration.resolve("Mixed Sans"), {
    family: regular.family,
    source: "embedded",
  });
  registration.close();
});

test("font policy chooses browser or provider for the same unresolved exact face", async () => {
  for (const [policy, expectedSource, expectedKind] of [
    ["deterministic", "provider", "object"],
    ["local-first", "browser", "string"],
  ]) {
    const fontSet = new FaceSet();
    const registration = await registerFonts(
      [],
      50,
      { FontFace: AvailableFace, fontSet },
      undefined,
      [],
      ["Policy Sans"],
      {
        providerAssets: [asset("Policy Sans", 4, { weight: 700 })],
        policy,
        requestedFaces: [request("Policy Sans", { weight: 700 })],
      },
    );

    const resolved = registration.resolveFace(request("Policy Sans", { weight: 700 }));
    assert.equal(resolved.source, expectedSource);
    assert.equal(
      resolved.family,
      policy === "local-first" ? "Policy Sans" : registration.resolve("Policy Sans").family,
    );
    const exactFace = [...fontSet.faces].find((face) => face.descriptors.weight === "700");
    assert.equal(typeof exactFace.source, expectedKind);
    registration.close();
  }
});

test("deterministic policy never probes an unmanaged local font", async () => {
  const fontSet = new FaceSet();
  const requested = request("Unmanaged Sans");
  const registration = await registerFonts(
    [],
    50,
    { FontFace: AvailableFace, fontSet },
    undefined,
    [],
    [],
    { policy: "deterministic", requestedFaces: [requested] },
  );

  assert.equal(fontSet.faces.size, 0);
  assert.equal(registration.hasExactFace(requested), false);
  assert.equal(registration.resolveFace(requested).source, "fallback");
  registration.close();
});

test("local-first falls back when the runtime cannot verify local fonts", async () => {
  const requested = request("等线", { codePoints: [0x4e2d] });
  const registration = await registerFonts(
    [],
    50,
    undefined,
    undefined,
    [],
    [],
    { policy: "local-first", requestedFaces: [requested] },
  );

  assert.equal(registration.hasExactFace(requested), false);
  assert.equal(registration.resolveFace(requested).source, "fallback");
  assert.equal(registration.resolveFace(requested).family, "Hiragino Sans");
  registration.close();
});

test("missing-face fallback follows each run's Unicode coverage demand", async () => {
  const fontSet = new FaceSet();
  const base = request("Missing Script Font", {
    codePoints: [0x41, 0x4e2d, 0x627, 0xf123],
  });
  const registration = await registerFonts(
    [],
    50,
    { FontFace: AvailableFace, fontSet },
    undefined,
    [],
    [],
    { policy: "deterministic", requestedFaces: [base] },
  );

  assert.equal(registration.resolveFace({ ...base, codePoints: [0x41, 0x416] }).family, "Calibri");
  assert.equal(registration.resolveFace({ ...base, codePoints: [0x4e2d] }).family, "Hiragino Sans");
  assert.equal(registration.resolveFace({ ...base, codePoints: [0x627, 0x5d0, 0xe01] }).family, "Arial");
  assert.equal(registration.resolveFace({ ...base, codePoints: [0xf123] }).family, "Segoe UI Symbol");
  assert.equal(registration.resolveFace({ ...base, codePoints: [0x41, 0x4e2d] }).family, "Calibri");
  registration.close();
});

test("local-first coverage checks use the requested subset text", async () => {
  class RecordingFaceSet extends FaceSet {
    checks = [];

    check(font, text) {
      this.checks.push({ font, text });
      return true;
    }
  }
  const fontSet = new RecordingFaceSet();
  const requested = request("Private Icon Subset", { codePoints: [0xf123] });
  const registration = await registerFonts(
    [],
    50,
    { FontFace: AvailableFace, fontSet },
    undefined,
    [],
    [],
    { policy: "local-first", requestedFaces: [requested] },
  );

  assert.equal(fontSet.checks.length, 2);
  assert.equal(fontSet.checks.every(({ text }) => text === "\uf123"), true);
  assert.equal(registration.resolveFace(requested).source, "browser");
  registration.close();
});

test("family fallback does not claim an unavailable exact face", async () => {
  const fontSet = new FaceSet();
  const registration = await registerFonts(
    [asset("Regular Only", 8)],
    50,
    { FontFace: AvailableFace, fontSet },
  );
  const regular = request("Regular Only");
  const bold = request("Regular Only", { weight: 700 });

  assert.equal(registration.hasExactFace(regular), true);
  assert.equal(registration.hasExactFace(bold), false);
  assert.equal(registration.resolveFace(bold).family, registration.resolveFace(regular).family);
  registration.close();
});

test("an unavailable bold local face does not masquerade as an exact match", async () => {
  class ExactLocalFace extends AvailableFace {
    async load() {
      if (this.source === 'local("Existing Sans")' && this.descriptors.weight === "700") {
        throw new Error("bold face is not installed");
      }
      return this;
    }
  }
  const fontSet = new FaceSet();
  const bold = request("Existing Sans", { weight: 700 });

  const registration = await registerFonts(
    [],
    50,
    { FontFace: ExactLocalFace, fontSet },
    undefined,
    [],
    [],
    { policy: "local-first", requestedFaces: [bold] },
  );

  assert.equal(registration.hasExactFace(bold), false);
  assert.equal(registration.resolveFace(bold).source, "fallback");
  registration.close();
});

test("an unavailable local face does not hide an available provider face", async () => {
  class NoLocalFace extends AvailableFace {
    async load() {
      if (typeof this.source === "string") throw new Error("local face is unavailable");
      return this;
    }
  }
  const fontSet = new FaceSet();
  const bold = request("Exact Sans", { weight: 700 });

  const registration = await registerFonts(
    [],
    50,
    { FontFace: NoLocalFace, fontSet },
    undefined,
    [],
    [],
    {
      policy: "local-first",
      providerAssets: [asset("Exact Sans", 9, { weight: 700 })],
      requestedFaces: [bold],
    },
  );

  assert.equal(registration.resolveFace(bold).source, "provider");
  registration.close();
});

test("local-first resolves SimSun through the platform-compatible STSong face", async () => {
  class CompatibleSongFace extends AvailableFace {
    static sources = [];

    async load() {
      CompatibleSongFace.sources.push(this.source);
      if (this.source !== 'local("STSong")') throw new Error("local face is unavailable");
      return this;
    }
  }
  const fontSet = new FaceSet();
  const simSun = request("宋体", { codePoints: [0x4e2d] });

  const registration = await registerFonts(
    [],
    50,
    { FontFace: CompatibleSongFace, fontSet },
    undefined,
    [],
    [],
    { policy: "local-first", requestedFaces: [simSun] },
  );

  const resolved = registration.resolveFace(simSun);
  assert.equal(resolved.source, "browser");
  assert.match(resolved.family, /^OfficeViewer \d+ browser \d+$/u);
  assert.deepEqual(CompatibleSongFace.sources, [
    'local("宋体")',
    'local("SimSun")',
    'local("NSimSun")',
    'local("STSong")',
    'local("Songti SC")',
  ]);
  registration.close();
});

test("local-first probes common cross-platform Office font replacements on demand", async () => {
  const replacements = new Map([
    ["楷体_GB2312", ['local("STKaiti")', 0x4e2d]],
    ["微软雅黑", ['local("PingFang SC")', 0x4e2d]],
    ["MS Mincho", ['local("Hiragino Mincho ProN")', 0x3042]],
    ["Malgun Gothic", ['local("Apple SD Gothic Neo")', 0xac00]],
    ["Traditional Arabic", ['local("Geeza Pro")', 0x627]],
    ["David", ['local("New Peninim MT")', 0x5d0]],
    ["Mangal", ['local("Kohinoor Devanagari")', 0x915]],
    ["Leelawadee UI", ['local("Thonburi")', 0xe01]],
    ["Calibri", ['local("Carlito")', 0x41]],
    ["Arial", ['local("Helvetica")', 0x41]],
    ["Times New Roman", ['local("Times")', 0x41]],
    ["Book Antiqua", ['local("Palatino")', 0x41]],
    ["Courier New", ['local("Courier")', 0x41]],
  ]);
  const availableSources = new Set([...replacements.values()].map(([source]) => source));
  availableSources.add('local("Aptos")');
  class CompatibleFace extends AvailableFace {
    async load() {
      if (!availableSources.has(this.source)) {
        throw new Error("local face is unavailable");
      }
      return this;
    }
  }
  const requestedFaces = [...replacements].map(([family, [, codePoint]]) => (
    request(family, { codePoints: [codePoint] })
  ));
  const fontSet = new FaceSet();

  const registration = await registerFonts(
    [],
    50,
    { FontFace: CompatibleFace, fontSet },
    undefined,
    [],
    [],
    { policy: "local-first", requestedFaces },
  );

  assert.equal(requestedFaces.every((face) => registration.resolveFace(face).source === "browser"), true);
  assert.deepEqual(
    new Set([...fontSet.faces].map(({ source }) => source)),
    new Set([...replacements.values()].map(([source]) => source)),
  );
  registration.close();
});

test("compatible local fonts keep their real weight so bold can be synthesized", async () => {
  class CompatibleSongFace extends AvailableFace {
    static loaded = [];

    async load() {
      if (this.source !== 'local("STSong")') throw new Error("local face is unavailable");
      CompatibleSongFace.loaded.push(this);
      return this;
    }
  }
  const fontSet = new FaceSet();
  const boldSimSun = request("宋体", { weight: 700, codePoints: [0x4e2d] });

  const registration = await registerFonts(
    [],
    50,
    { FontFace: CompatibleSongFace, fontSet },
    undefined,
    [],
    [],
    { policy: "local-first", requestedFaces: [boldSimSun] },
  );

  assert.equal(registration.resolveFace(boldSimSun).source, "browser");
  assert.equal(CompatibleSongFace.loaded.length, 1);
  assert.equal(CompatibleSongFace.loaded[0].descriptors.weight, "400");
  registration.close();
});

test("inline registrations use realm-unique aliases across documents", async () => {
  const fontSet = new FaceSet();
  const runtime = { FontFace: AvailableFace, fontSet };
  const first = await registerFonts([], 50, runtime, undefined, [asset("Alpha Sans", 1)]);
  const second = await registerFonts([], 50, runtime, undefined, [asset("Beta Sans", 2)]);

  assert.notEqual(first.resolve("Alpha Sans").family, second.resolve("Beta Sans").family);

  first.close();
  second.close();
});
