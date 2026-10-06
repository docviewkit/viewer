import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";

import { createOfficeEngine } from "../dist/engine.js";

const wasm = await readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url));

const emptySnapshot = Object.freeze({
  info: {
    format: "pptx",
    kind: "presentation",
    title: null,
    units: [],
  },
  diagnostics: [],
  objects: [],
});

class FakeWorker {
  constructor() {
    this.onmessage = null;
    this.onerror = null;
    this.onmessageerror = null;
    this.messages = [];
    this.transfers = [];
    this.terminated = false;
  }

  postMessage(message, transfer) {
    if (this.terminated) {
      throw new Error("worker terminated");
    }
    this.messages.push(message);
    this.transfers.push(transfer);
  }

  terminate() {
    this.terminated = true;
  }

  respondToOpen(snapshot = emptySnapshot) {
    const request = this.messages.findLast((message) => message.type === "open");
    assert.ok(request, "expected an open request");
    queueMicrotask(() => {
      this.onmessage?.({
        data: { id: request.id, ok: true, type: "open", snapshot },
      });
    });
  }

  requestFonts(requests) {
    const request = this.messages.findLast((message) => message.type === "open");
    assert.ok(request, "expected an open request");
    queueMicrotask(() => {
      this.onmessage?.({
        data: { id: request.id, ok: true, type: "font-request", requests },
      });
    });
  }

  fail(message = "worker failed") {
    this.onerror?.({ message });
  }
}

class CountingSignal {
  constructor() {
    this.aborted = false;
    this.reason = undefined;
    this.listeners = new Set();
  }

  addEventListener(type, listener) {
    if (type === "abort") this.listeners.add(listener);
  }

  removeEventListener(type, listener) {
    if (type === "abort") this.listeners.delete(listener);
  }

  abort(reason) {
    if (this.aborted) return;
    this.aborted = true;
    this.reason = reason;
    for (const listener of [...this.listeners]) listener.call(this, { type: "abort" });
  }
}

class SliceTrap extends Uint8Array {
  slice() {
    throw new Error("input was copied before its size was checked");
  }
}

async function withWorkerGlobal(run, WorkerClass = FakeWorker) {
  const descriptor = Object.getOwnPropertyDescriptor(globalThis, "Worker");
  Object.defineProperty(globalThis, "Worker", {
    configurable: true,
    value: WorkerClass,
    writable: true,
  });
  try {
    return await run();
  } finally {
    if (descriptor) Object.defineProperty(globalThis, "Worker", descriptor);
    else delete globalThis.Worker;
  }
}

async function waitForWorkerMessage(worker, type) {
  const deadline = performance.now() + 5_000;
  while (performance.now() < deadline) {
    const message = worker.messages.findLast((candidate) => candidate.type === type);
    if (message !== undefined) return message;
    await new Promise((resolve) => setTimeout(resolve, 10));
  }
  assert.fail(`timed out waiting for Worker ${type} message`);
}

test("worker open remains pending until a font request is resolved and the final snapshot arrives", async () => {
  await withWorkerGlobal(async () => {
    const worker = new FakeWorker();
    const providerCalls = [];
    const engine = await createOfficeEngine({
      execution: "worker",
      wasm,
      workerFactory: () => worker,
      fontProvider: async (requests, context) => {
        providerCalls.push({ requests, context });
        await new Promise((resolve) => setTimeout(resolve, 50));
        return [{
          family: "Worker Sans",
          bytes: Uint8Array.of(0, 1, 0, 0, 7, 8, 9),
          style: "italic",
          weight: 600,
          stretch: "condensed",
        }];
      },
    });
    const requests = [{
      family: "Worker Sans",
      style: "italic",
      weight: 600,
      stretch: "condensed",
      codePoints: [0x41, 0x4e2d],
    }];
    const opening = engine.open(new Uint8Array([1]));
    let state = "pending";
    void opening.then(
      () => { state = "resolved"; },
      () => { state = "rejected"; },
    );

    try {
      worker.requestFonts(requests);
      const response = await waitForWorkerMessage(worker, "font-response");

      assert.equal(state, "pending", "font-request must not complete open");
      assert.equal(providerCalls.length, 1);
      assert.deepEqual(providerCalls[0].requests, requests);
      assert.equal(providerCalls[0].context.policy, "local-first");
      assert.equal(response.id, worker.messages.find((message) => message.type === "open").id);
      assert.deepEqual(response.diagnostics, []);
      assert.equal(response.fonts.length, 1);
      assert.deepEqual({
        family: response.fonts[0].family,
        style: response.fonts[0].style,
        weight: response.fonts[0].weight,
        stretch: response.fonts[0].stretch,
      }, {
        family: "Worker Sans",
        style: "italic",
        weight: 600,
        stretch: "condensed",
      });
      assert.deepEqual([...new Uint8Array(response.fonts[0].bytes)], [0, 1, 0, 0, 7, 8, 9]);
      const responseIndex = worker.messages.indexOf(response);
      assert.equal(worker.transfers[responseIndex].includes(response.fonts[0].bytes), true);

      worker.respondToOpen();
      const document = await opening;
      assert.equal(state, "resolved");
      document.close();
    } finally {
      engine.close();
      await opening.catch(() => undefined);
    }
  });
});

test("engine.close terminates a pending open", async () => {
  await withWorkerGlobal(async () => {
    const worker = new FakeWorker();
    const engine = await createOfficeEngine({
      execution: "worker",
      wasm,
      workerFactory: () => worker,
    });

    const pending = engine.open(new Uint8Array([1]));
    engine.close();

    if (!worker.terminated) worker.respondToOpen();
    const outcome = await pending.then(
      (document) => {
        document.close();
        return "resolved";
      },
      (error) => error?.code,
    );

    assert.equal(worker.terminated, true);
    assert.equal(outcome, "ENGINE_CLOSED");
  });
});

test("a worker channel failure permanently closes the document", async () => {
  await withWorkerGlobal(async () => {
    const worker = new FakeWorker();
    const engine = await createOfficeEngine({
      execution: "worker",
      wasm,
      workerFactory: () => worker,
    });
    const opening = engine.open(new Uint8Array([1]));
    worker.respondToOpen();
    const document = await opening;

    worker.fail("channel gone");
    const outcome = await Promise.race([
      document.hitTest({ unitIndex: 0, x: 0, y: 0 }).then(
        () => "resolved",
        (error) => error?.code,
      ),
      new Promise((resolve) => setTimeout(() => resolve("pending"), 20)),
    ]);

    document.close();
    engine.close();
    assert.equal(worker.terminated, true);
    assert.equal(outcome, "WORKER_FAILED");
  });
});

test("malformed worker errors close the channel instead of orphaning a request", async () => {
  await withWorkerGlobal(async () => {
    const worker = new FakeWorker();
    const engine = await createOfficeEngine({
      execution: "worker",
      wasm,
      workerFactory: () => worker,
    });
    const opening = engine.open(new Uint8Array([1]));
    worker.respondToOpen();
    const document = await opening;

    const rendering = document.render({ unitIndex: 0 });
    const request = worker.messages.findLast((message) => message.type === "render");
    let handlerFailure;
    try {
      worker.onmessage?.({
        data: {
          id: request.id,
          ok: false,
          error: { code: "BAD", message: "malformed", diagnostics: [null] },
        },
      });
    } catch (error) {
      handlerFailure = error;
    }
    const outcome = await Promise.race([
      rendering.then(() => "resolved", (error) => error?.code),
      new Promise((resolve) => setTimeout(() => resolve("pending"), 20)),
    ]);

    document.close();
    engine.close();
    assert.equal(handlerFailure, undefined);
    assert.equal(outcome, "WORKER_PROTOCOL_ERROR");
    assert.equal(worker.terminated, true);
  });
});

test("hit testing rejects fractional units and invalid limits before Worker dispatch", async () => {
  await withWorkerGlobal(async () => {
    const worker = new FakeWorker();
    const engine = await createOfficeEngine({
      execution: "worker",
      wasm,
      workerFactory: () => worker,
    });
    const opening = engine.open(new Uint8Array([1]));
    worker.respondToOpen({
      ...emptySnapshot,
      info: {
        format: "pptx",
        kind: "presentation",
        units: [{ type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 960, height: 720 }],
      },
    });
    const document = await opening;
    const messages = worker.messages.length;

    await assert.rejects(
      document.hitTest({ unitIndex: 0.5, x: 0, y: 0 }),
      (error) => error?.code === "INVALID_UNIT",
    );
    await assert.rejects(
      document.hitTest({ unitIndex: 0, x: 0, y: 0, limit: Number.NaN }),
      (error) => error?.code === "INVALID_HIT_TEST",
    );
    await assert.rejects(
      document.hitTest({ unitIndex: 0, x: 0, y: 0, limit: 257 }),
      (error) => error?.code === "INVALID_HIT_TEST",
    );
    assert.equal(worker.messages.length, messages);

    document.close();
    engine.close();
  });
});

test("worker metadata and on-demand object DTOs are deeply copied and immutable", async () => {
  await withWorkerGlobal(async () => {
    const worker = new FakeWorker();
    const engine = await createOfficeEngine({
      execution: "worker",
      wasm,
      workerFactory: () => worker,
    });
    const objectResponse = {
      numericId: 1,
      id: "shape:7",
      type: "text-box",
      unitIndex: 0,
      bounds: { x: 10, y: 20, width: 30, height: 40 },
      text: "Hello",
      fontRuns: [{
        start: 0,
        end: 5,
        authoredFamily: "Aptos",
        renderedFamily: "OfficeViewer browser 0",
        source: "browser",
      }],
      source: {
        format: "pptx",
        part: "ppt/slides/slide1.xml",
        kind: "shape",
        shapeId: 7,
        row: 2,
        column: 3,
        textRange: [4, 9],
        mapping: "exact",
      },
    };
    const snapshot = {
      info: {
        format: "pptx",
        kind: "presentation",
        units: [
          { type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 960, height: 720 },
        ],
      },
      diagnostics: [{
        code: "APPROXIMATE",
        severity: "warning",
        fidelity: "approximate",
        phase: "layout",
        message: "Approximate layout",
        details: { feature: "test" },
      }],
    };
    const documentDiagnostics = [
      ...snapshot.diagnostics,
      {
        code: "FONT_LOAD_FAILED",
        severity: "warning",
        fidelity: "approximate",
        phase: "render",
        message: "Embedded font could not be loaded",
        details: { family: "Broken PDF Sans" },
      },
    ];
    const opening = engine.open(new Uint8Array([1]));
    worker.respondToOpen(snapshot);
    const document = await opening;
    const gettingObject = document.getObject("shape:7");
    const getRequest = await waitForWorkerMessage(worker, "get-object");
    worker.onmessage?.({
      data: {
        id: getRequest.id,
        ok: true,
        type: "get-object",
        object: objectResponse,
        documentDiagnostics,
      },
    });
    const object = await gettingObject;

    snapshot.info.units[0].name = "mutated";
    objectResponse.bounds.x = 999;
    objectResponse.fontRuns[0].authoredFamily = "mutated";
    objectResponse.source.textRange[0] = 999;
    snapshot.diagnostics[0].details.feature = "mutated";
    documentDiagnostics[1].details.family = "mutated";

    assert.equal(document.info.units[0].name, "Slide 1");
    assert.equal(object.bounds.x, 10);
    assert.equal(object.fontRuns[0].authoredFamily, "Aptos");
    assert.deepEqual(object.source.textRange, [4, 9]);
    assert.equal(document.diagnostics()[0].details.feature, "test");
    assert.equal(document.diagnostics()[1].details.family, "Broken PDF Sans");
    assert.equal(Object.isFrozen(document.info.units[0]), true);
    assert.equal(Object.isFrozen(object), true);
    assert.equal(Object.isFrozen(object.bounds), true);
    assert.equal(Object.isFrozen(object.fontRuns), true);
    assert.equal(Object.isFrozen(object.fontRuns[0]), true);
    assert.equal(Object.isFrozen(object.source), true);
    assert.equal(Object.isFrozen(object.source.textRange), true);
    assert.equal(Object.isFrozen(document.diagnostics()[0].details), true);
    assert.equal(Object.isFrozen(document.diagnostics()[1].details), true);

    document.close();
    engine.close();
  });
});

test("aborting an open reports OPERATION_ABORTED", async () => {
  await withWorkerGlobal(async () => {
    const worker = new FakeWorker();
    const engine = await createOfficeEngine({
      execution: "worker",
      wasm,
      workerFactory: () => worker,
    });
    const signal = new CountingSignal();
    const opening = engine.open(new Uint8Array([1]), { signal });

    signal.abort("cancelled");

    await assert.rejects(opening, (error) => error?.code === "OPERATION_ABORTED");
    assert.equal(signal.listeners.size, 0);
    assert.equal(worker.terminated, true);
    engine.close();
  });
});

test("an already-aborted open rejects and leaves no listener", async () => {
  await withWorkerGlobal(async () => {
    const worker = new FakeWorker();
    let workerStarts = 0;
    const engine = await createOfficeEngine({
      execution: "worker",
      wasm,
      workerFactory: () => {
        workerStarts += 1;
        return worker;
      },
    });
    const signal = new CountingSignal();
    signal.abort("cancelled before open");

    await assert.rejects(
      engine.open(new Uint8Array([1]), { signal }),
      (error) => error?.code === "OPERATION_ABORTED",
    );
    assert.equal(signal.listeners.size, 0);
    assert.equal(workerStarts, 0);
    assert.equal(worker.terminated, false);
    engine.close();
  });
});

test("timing out an open removes its abort listener", async () => {
  await withWorkerGlobal(async () => {
    const worker = new FakeWorker();
    const engine = await createOfficeEngine({
      execution: "worker",
      wasm,
      workerFactory: () => worker,
    });
    const signal = new CountingSignal();

    await assert.rejects(
      engine.open(new Uint8Array([1]), { signal, timeoutMs: 5 }),
      (error) => error?.code === "OPERATION_TIMEOUT",
    );

    assert.equal(signal.listeners.size, 0);
    assert.equal(worker.terminated, true);
    engine.close();
  });
});

test("input size is rejected before bytes are copied", async () => {
  await withWorkerGlobal(async () => {
    const worker = new FakeWorker();
    const engine = await createOfficeEngine({
      execution: "worker",
      wasm,
      workerFactory: () => worker,
      limits: { inputBytes: 4 },
    });

    await assert.rejects(
      engine.open(new SliceTrap(8)),
      (error) => error?.code === "INPUT_SIZE_LIMIT",
    );

    engine.close();
  });

  const engine = await createOfficeEngine({
    execution: "inline",
    wasm,
    limits: { inputBytes: 4 },
  });
  try {
    await assert.rejects(
      engine.open(new SliceTrap(8)),
      (error) => error?.code === "INPUT_SIZE_LIMIT",
    );
  } finally {
    engine.close();
  }
});

test("worker input from another JavaScript realm is normalized to a transferable ArrayBuffer", async () => {
  await withWorkerGlobal(async () => {
    const worker = new FakeWorker();
    const engine = await createOfficeEngine({
      execution: "worker",
      wasm,
      workerFactory: () => worker,
    });
    const foreign = vm.runInNewContext("Uint8Array.from([1, 2, 3])");
    assert.equal(foreign instanceof Uint8Array, false);

    const opening = engine.open(foreign);
    try {
      const request = worker.messages.at(-1);
      assert.equal(request.bytes instanceof ArrayBuffer, true);
      assert.deepEqual([...new Uint8Array(request.bytes)], [1, 2, 3]);
      assert.equal(worker.transfers.at(-1).includes(request.bytes), true);
    } finally {
      worker.respondToOpen();
      const document = await opening;
      document.close();
      engine.close();
    }
  });
});

test("transferInput explicitly transfers the caller-owned ArrayBuffer without a host copy", async () => {
  await withWorkerGlobal(async () => {
    const worker = new FakeWorker();
    const engine = await createOfficeEngine({ execution: "worker", wasm, workerFactory: () => worker });
    const input = Uint8Array.of(1, 2, 3).buffer;
    const opening = engine.open(input, { transferInput: true });
    try {
      const request = await waitForWorkerMessage(worker, "open");
      assert.equal(request.bytes, input);
      assert.equal(worker.transfers.at(-1).includes(input), true);
    } finally {
      worker.respondToOpen();
      const document = await opening;
      document.close();
      engine.close();
    }
  });
});

test("PDF passwords cross the Worker boundary as transferred bytes and are cleared after open", async () => {
  await withWorkerGlobal(async () => {
    const worker = new FakeWorker();
    const engine = await createOfficeEngine({ execution: "worker", wasm, workerFactory: () => worker });
    const opening = engine.open(new Uint8Array([1]), { password: "secret" });
    const request = await waitForWorkerMessage(worker, "open");
    assert.deepEqual([...new Uint8Array(request.password)], [...new TextEncoder().encode("secret")]);
    assert.equal(worker.transfers.at(-1).includes(request.password), true);
    worker.respondToOpen();
    const document = await opening;
    assert.deepEqual([...new Uint8Array(request.password)], [0, 0, 0, 0, 0, 0]);
    document.close();
    engine.close();
  });
});

test("render options reach the Worker and cancellation does not terminate the document", async () => {
  await withWorkerGlobal(async () => {
    const worker = new FakeWorker();
    const engine = await createOfficeEngine({ execution: "worker", wasm, workerFactory: () => worker });
    const opening = engine.open(new Uint8Array([1]));
    worker.respondToOpen({
      ...emptySnapshot,
      info: {
        format: "pptx",
        kind: "presentation",
        units: [{ type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 960, height: 720 }],
      },
    });
    const document = await opening;
    const signal = new CountingSignal();
    const rendering = document.render({ unitIndex: 0 }, {
      priority: "interactive",
      supersedeKey: "surface",
      signal,
    });
    const request = await waitForWorkerMessage(worker, "render");
    assert.equal(request.priority, "interactive");
    assert.equal(request.supersedeKey, "surface");

    signal.abort("new viewport");
    const cancellation = await waitForWorkerMessage(worker, "cancel-render");
    assert.equal(cancellation.targetId, request.id);
    worker.onmessage?.({
      data: {
        id: request.id,
        ok: false,
        error: { code: "OPERATION_ABORTED", message: "Rendering was cancelled", diagnostics: [] },
      },
    });
    await assert.rejects(rendering, (error) => error?.code === "OPERATION_ABORTED");
    assert.equal(worker.terminated, false);
    assert.equal(signal.listeners.size, 0);

    document.close();
    engine.close();
  });
});

test("document query deadlines terminate the dedicated Worker", async () => {
  await withWorkerGlobal(async () => {
    const worker = new FakeWorker();
    const engine = await createOfficeEngine({ execution: "worker", wasm, workerFactory: () => worker });
    const opening = engine.open(new Uint8Array([1]));
    worker.respondToOpen();
    const document = await opening;

    await assert.rejects(
      document.searchText({ query: "missing" }, { timeoutMs: 5 }),
      (error) => error?.code === "OPERATION_TIMEOUT",
    );
    assert.equal(worker.terminated, true);
    await assert.rejects(
      document.listObjects(),
      (error) => error?.code === "OPERATION_TIMEOUT",
    );
    engine.close();
  });
});

test("render deadlines terminate a Worker that cannot reach a cancellation checkpoint", async () => {
  await withWorkerGlobal(async () => {
    const worker = new FakeWorker();
    const engine = await createOfficeEngine({ execution: "worker", wasm, workerFactory: () => worker });
    const opening = engine.open(new Uint8Array([1]));
    worker.respondToOpen({
      ...emptySnapshot,
      info: {
        format: "pptx",
        kind: "presentation",
        units: [{ type: "slide", index: 0, id: "unit:0", name: "Slide 1", width: 960, height: 720 }],
      },
    });
    const document = await opening;

    await assert.rejects(
      document.render({ unitIndex: 0 }, { timeoutMs: 5 }),
      (error) => error?.code === "OPERATION_TIMEOUT",
    );
    assert.equal(worker.terminated, true);
    engine.close();
  });
});

test("default execution rejects when Worker is unavailable", async () => {
  const descriptor = Object.getOwnPropertyDescriptor(globalThis, "Worker");
  delete globalThis.Worker;
  try {
    await assert.rejects(
      createOfficeEngine({ wasm }),
      (error) => error?.code === "UNSUPPORTED_ENVIRONMENT",
    );
  } finally {
    if (descriptor) Object.defineProperty(globalThis, "Worker", descriptor);
  }
});

test("default worker execution keeps one document-free Worker warm", async () => {
  const workers = [];
  class RecordingWorker extends FakeWorker {
    constructor() {
      super();
      workers.push(this);
    }
  }
  await withWorkerGlobal(async () => {
    const engine = await createOfficeEngine({ execution: "worker", wasm });
    assert.equal(workers.length, 1, "engine creation should prewarm one Worker");
    const opening = engine.open(new Uint8Array([1]));
    assert.equal(workers.length, 2, "opening should replenish the idle Worker");
    workers[0].respondToOpen();
    const document = await opening;
    document.close();
    assert.equal(workers[0].terminated, true);
    assert.equal(workers[1].terminated, false);
    engine.close();
    assert.equal(workers[1].terminated, true);
  }, RecordingWorker);
});

test("inline execution explicitly rejects timeoutMs", async () => {
  const engine = await createOfficeEngine({ execution: "inline", wasm });
  try {
    await assert.rejects(
      engine.open(new Uint8Array([1]), { timeoutMs: 1 }),
      (error) => error?.code === "INLINE_TIMEOUT_UNSUPPORTED",
    );
  } finally {
    engine.close();
  }
});
