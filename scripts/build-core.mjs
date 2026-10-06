import { mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { delimiter, dirname, resolve } from "node:path";
import binaryen from "binaryen";

function prepareCalculator(input, output) {
  const required = new Set([
    "memory", "ov_alloc", "ov_free", "ov_calc_abi_version", "ov_calculate",
    "ov_result_pointer", "ov_result_clear",
  ]);
  const module = binaryen.readBinary(readFileSync(input));
  let bytes;
  try {
    for (let index = module.getNumExports() - 1; index >= 0; index -= 1) {
      const { name } = binaryen.getExportInfo(module.getExportByIndex(index));
      if (!required.delete(name)) module.removeExport(name);
    }
    if (required.size !== 0) throw new Error(`Calculator ABI is missing: ${[...required].join(", ")}`);
    bytes = Buffer.from(module.emitBinary());
  } finally {
    module.dispose();
  }
  // wasm-bindgen's CLI metadata is not consumed by our direct Wasm ABI.
  let offset = 8;
  const sections = [bytes.subarray(0, offset)];
  function readU32() {
    let value = 0;
    for (let shift = 0; shift < 35 && offset < bytes.length; shift += 7) {
      const byte = bytes[offset++];
      value += (byte & 127) * 2 ** shift;
      if (byte < 128 && value <= 0xffff_ffff) return value;
    }
    throw new Error("Invalid Wasm section length");
  }
  while (offset < bytes.length) {
    const start = offset;
    const id = bytes[offset++];
    const length = readU32();
    const end = offset + length;
    if (end > bytes.length) throw new Error("Truncated Wasm section");
    let name;
    if (id === 0) {
      const nameLength = readU32();
      if (offset + nameLength > end) throw new Error("Truncated Wasm custom section name");
      name = bytes.toString("utf8", offset, offset + nameLength);
    }
    if (name !== "__wasm_bindgen_unstable") sections.push(bytes.subarray(start, end));
    offset = end;
  }
  const prepared = Buffer.concat(sections);
  if (!WebAssembly.validate(prepared)) throw new Error("Invalid calculator Wasm");
  writeFileSync(output, prepared);
}

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const rustupCargo = spawnSync(
  "rustup",
  ["which", "cargo"],
  { cwd: root, encoding: "utf8" },
);
const cargoCommand = process.env.CARGO
  ?? (rustupCargo.status === 0
    ? rustupCargo.stdout.trim()
    : (process.platform === "win32" ? "cargo.exe" : "cargo"));
const cargoBin = dirname(cargoCommand);
const wasmOptCommand = resolve(
  root,
  "node_modules/.bin",
  process.platform === "win32" ? "wasm-opt.cmd" : "wasm-opt",
);
mkdirSync(resolve(root, "dist"), { recursive: true });
rmSync(resolve(root, "dist/office-viewer-iwork.wasm"), { force: true });

const builds = [
  { features: undefined, output: "office-viewer-core.wasm", targetDir: "target/wasm/base" },
  {
    features: "calculation-service",
    output: "office-viewer-calc.wasm",
    targetDir: "target/wasm/calculation",
  },
  { features: "odf-formats", output: "office-viewer-odf.wasm", targetDir: "target/wasm/odf" },
  {
    features: "legacy-office-formats",
    output: "office-viewer-legacy-office.wasm",
    targetDir: "target/wasm/legacy-office",
  },
  {
    features: "iwork-formats,pdf-formats",
    output: "office-viewer-pdf.wasm",
    targetDir: "target/wasm/iwork-pdf",
  },
  { features: "xps-formats", output: "office-viewer-xps.wasm", targetDir: "target/wasm/xps" },
  { features: "ofd-formats", output: "office-viewer-ofd.wasm", targetDir: "target/wasm/ofd" },
];

for (const build of builds) {
  const args = [
    "build",
    "--release",
    "--target",
    "wasm32-unknown-unknown",
    "--target-dir",
    build.targetDir,
    "-p",
    "office-viewer-core",
  ];
  if (build.features !== undefined) {
    args.push("--no-default-features", "--features", build.features);
  }
  const cargo = spawnSync(cargoCommand, args, {
    cwd: root,
    stdio: "inherit",
    env: {
      ...process.env,
      PATH: `${cargoBin}${delimiter}${process.env.PATH ?? ""}`,
    },
  });
  if (cargo.status !== 0) process.exit(cargo.status ?? 1);
  let input = resolve(
    root,
    build.targetDir,
    "wasm32-unknown-unknown/release/office_viewer_core.wasm",
  );
  if (build.features === "calculation-service") {
    const prepared = resolve(root, build.targetDir, "calculator-runtime.wasm");
    prepareCalculator(input, prepared);
    input = prepared;
  }
  const output = resolve(root, `dist/${build.output}`);
  const wasmOpt = spawnSync(
    wasmOptCommand,
    [
      "-O3",
      "--enable-bulk-memory",
      "--enable-nontrapping-float-to-int",
      input,
      "-o",
      output,
    ],
    { cwd: root, stdio: "inherit" },
  );
  if (wasmOpt.status !== 0) process.exit(wasmOpt.status ?? 1);
}
