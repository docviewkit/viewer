# DocViewKit Viewer

DocViewKit is a lightweight document viewer for everyday OA attachment previews, approval workflows, admin portals, CRM/ERP systems, cloud drives, and customer-facing apps, as well as enterprise SaaS, AI knowledge bases, legal review, and financial audit. Quick integration starts with the ready-made Viewer: import the component, mount `<docviewkit-viewer>`, and call `open(file)`. A compact core and on-demand format packs keep unused parsers out of the initial load, with no document-conversion server to deploy. Use the Engine API when you need custom rendering or source-object access.

Apache-2.0 software with optional paid support, customization and enterprise delivery.
[Source repository](https://github.com/docviewkit/viewer) · [Documentation](https://docviewkit.com/docs/quickstart/) · [Contact services](mailto:novalag778@gmail.com)

Open source, high-performance, lightweight document viewing—right in your frontend application.

DocViewKit Viewer is a frontend-local, read-only Web Component for Office, PDF, OFD, OpenDocument, iWork, WPS Office, XPS, CSV, and RTF files. It uses a compact core with on-demand format packs for fast, responsive viewing across Chrome, Edge, Firefox, and Safari. Document bytes stay on the device.

## Frontend runtime compatibility

The Viewer runs in frontend environments that provide DOM/Custom Elements, module Workers, WebAssembly, Canvas/OffscreenCanvas, ImageBitmap, and FontFace. These include browsers and compatible WebView hosts in Electron, Tauri, Ionic and Capacitor. React Native and Flutter applications can embed it through a WebView that provides the same APIs; their native rendering layers do not run the Web Component directly. Validate the actual host engine, asset loading, CSP, fonts, and required interactions on each target platform.

## Install

```sh
npm install @docviewkit/viewer
```

Mount the component in your application:

```html
<docviewkit-viewer></docviewkit-viewer>
```

Then import it in client code and pass a document supplied by your application:

```js
import "@docviewkit/viewer";

const viewer = document.querySelector("docviewkit-viewer");
await viewer.open(file);
```

Keep the package's Worker, Wasm, font and codec assets together when deploying. The host application owns document access and authentication; no DocViewKit website account or application registration is required. See the [complete integration guide](https://docviewkit.com/docs/quickstart/) for asset URLs, CSP and private attachments.

## Engine API

```js
import { createOfficeEngine } from "@docviewkit/viewer/engine";

const engine = await createOfficeEngine();
const document = await engine.open(await file.arrayBuffer());
// Render, search, hit-test and inspect source objects through the Engine API.
document.close();
engine.close();
```

## First-open performance

Configure optional formats before the first `open()` and start their small JavaScript dispatcher import early:

```js
import "@docviewkit/viewer";

const formatPack = import("@docviewkit/viewer/extended-formats")
  .then(({ extendedFormatPack }) => extendedFormatPack);

const viewer = document.querySelector("docviewkit-viewer");
viewer.config = { engine: { formatPack: () => formatPack } };
await viewer.open(file);
```

Only the Wasm module matching the detected document family is then loaded. Use `wps-formats`, `xps-formats`, or `ofd-formats` when the product accepts only that optional family. OFD 1.0 and 1.1 fixed-layout documents are parsed and rendered locally; see the format-boundaries page for the implemented subset and signature-validation limits. Serve versioned runtime assets with HTTP caching, reuse the Viewer with `close()` between documents, and avoid preloading every Wasm module or starting all-page rendering and document-wide search before first paint.

## Why DocViewKit

- **Free to use:** every supported format and the Engine API are available under Apache-2.0.
- **High performance:** documents are parsed and rendered locally without a conversion service.
- **Lightweight by design:** the compact core and optional format packs keep unused capabilities out of the startup path.
- **Broad compatibility:** one Viewer covers modern and legacy document formats in browsers and compatible cross-platform WebView hosts.
- **Fast and responsive:** local processing, lazy loading, and incremental rendering keep viewing interactions immediate.
- **Frontend-only:** the component runs entirely in the host frontend and requires no backend service.
- **Private by default:** normal viewing never uploads document bytes.

## Licensing

DocViewKit Viewer and Engine are licensed under Apache-2.0. All supported formats, Engine APIs, UI customization, minimal mode and text watermarks require no runtime license or mandatory branding. Optional paid services cover support, customization and enterprise delivery; they do not restrict open source usage rights. Third-party components and fonts retain their own licenses.

These terms apply to releases built from this source; older published versions retain their included licenses. Retain [LICENSE](LICENSE), [NOTICE](NOTICE) and [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md) when redistributing the corresponding materials.

- [Online demo](https://docviewkit.com/en/demo/)
- [Support, customization and enterprise delivery](mailto:novalag778@gmail.com)
- [Releases](https://github.com/docviewkit/viewer/releases)
- [Bug reports](https://github.com/docviewkit/viewer/issues)
- [Format boundaries](https://docviewkit.com/docs/supported-formats/)

The source repository is [docviewkit/viewer](https://github.com/docviewkit/viewer); its README covers building and contributing. Do not attach confidential documents to public issues; follow [the security policy](https://github.com/docviewkit/viewer/blob/main/docs/SECURITY.md) for security reports.
