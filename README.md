# DocViewKit Viewer

Free, high-performance, lightweight document viewing—right in your browser.

DocViewKit Viewer is a browser-local, read-only Web Component for Office, PDF, OpenDocument, iWork, WPS Office, XPS, CSV, RTF, and Markdown files. It uses a compact core with on-demand format packs for fast, responsive viewing across Chrome, Edge, Firefox, and Safari. Document bytes stay in the browser.

## Install

```sh
npm install @docviewkit/viewer
```

```js
import "@docviewkit/viewer";

const viewer = document.querySelector("docviewkit-viewer");
await viewer.open(file);
```

## Why DocViewKit

- **Free to use:** every supported format is available in the free Viewer.
- **High performance:** documents are parsed and rendered locally without a conversion service.
- **Lightweight by design:** the compact core and optional format packs keep unused capabilities out of the startup path.
- **Broad compatibility:** one Viewer covers modern and legacy document formats across major browser engines.
- **Fast and responsive:** local processing, lazy loading, and incremental rendering keep viewing interactions immediate.
- **Frontend-only:** the component runs entirely in the browser and requires no backend service.
- **Private by default:** normal viewing never uploads document bytes.

## Licensing

Without a valid license, the Viewer keeps rendering and displays an English unlicensed watermark with a link to obtain a free license. A valid free license removes the unlicensed watermark and retains the fixed DocViewKit brand. Commercial licenses add white-label presentation, deeper integration, and support options; they do not unlock additional document formats.

- Online demo: https://docviewkit.com/demo/
- Free license: https://docviewkit.com/portal/
- Releases: https://github.com/docviewkit/viewer/releases
- Bug reports: https://github.com/docviewkit/viewer/issues
- Format boundaries: https://docviewkit.com/docs/?doc=support

This repository contains compiled artifacts, not product source code. Do not attach confidential documents to public issues; follow [SECURITY.md](SECURITY.md) for security reports.
