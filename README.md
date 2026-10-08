# DocViewKit Viewer

**Lightweight document previews. Quick integration. Fits your business.**

Give your users a place to read Word, Excel, PowerPoint, PDF, OFD and more inside your application. DocViewKit brings document viewing, search and source location to everyday OA attachments, approval workflows, CRM/ERP, cloud drives and enterprise apps. Documents are parsed and rendered locally in the frontend, with no conversion server to deploy.

[Website](https://docviewkit.com/en/) · [Try your files](https://docviewkit.com/en/demo/) · [Documentation](https://docviewkit.com/docs/quickstart/) · [npm](https://www.npmjs.com/package/@docviewkit/viewer)

## Why DocViewKit

- **Quick to integrate.** Start with a ready-made Viewer instead of assembling a preview interface. Navigation, search, zoom and printing are already available.
- **Files stay local.** Parsing, rendering and search run in your application's frontend. DocViewKit does not upload documents to a conversion service; your application keeps control of attachment access and permissions.
- **Lightweight by design.** A compact core and on-demand format packs keep unused parsers out of the initial load. Pages, slides and spreadsheet viewports are rendered as needed.
- **One viewing experience.** Use the same component for modern and legacy Office files, WPS, PDF, OFD, OpenDocument, iWork and other supported formats.
- **Find the source behind a result.** Search and jump to relevant content. Connect your application's search results or AI citations to document pages, objects and regions so users can check the original material.
- **Make it your own.** Adapt the interface, commands, theme and text watermarks to your product. Use the Engine API for custom rendering and source-object access.
- **Open source, free to use.** Viewer, Engine and every implemented format are available under Apache-2.0. Paid services cover support, customization and enterprise delivery.

## Where it fits

- **OA, approvals and internal tools:** read attachments alongside the task being reviewed.
- **CRM, ERP and customer portals:** view contracts, records and business documents within the existing workflow.
- **Cloud drives, enterprise SaaS and AI knowledge bases:** preview files and bring search results or citations back to their source.
- **Legal review, audit and public-sector systems:** inspect documents and locate supporting passages while using your own access controls.
- **Education and training:** read course materials, presentations and assignments in the application.

## What your users can do

Browse pages, slides and sheets with navigation, thumbnails, zoom and fullscreen. Search document text, select and copy it, and jump to the relevant content. Add text watermarks to pages, thumbnails and printed output. Choose a full Viewer interface or a minimal document surface that fits your product.

The Viewer is read-only. Your application owns document storage, authorization and business workflows; DocViewKit provides viewing and source location.

## Supported formats

| Document family | Examples |
| --- | --- |
| Modern Microsoft Office | Word `.docx`, Excel `.xlsx`, PowerPoint `.pptx`, including supported macro-enabled, template and slideshow variants |
| Legacy Microsoft Office | Word `.doc`, Excel `.xls`, PowerPoint `.ppt` |
| WPS Office | `.wps`, `.et`, `.dps` |
| Fixed-layout documents | PDF, OFD 1.0/1.1, XPS and OpenXPS |
| OpenDocument | `.odt`, `.ods`, `.odp`, `.odg`, plus supported templates and flat XML variants |
| Apple iWork | `.pages`, `.numbers`, `.key` single-file packages |
| Text and tabular data | CSV and RTF |

Compatibility and rendering depth vary by format and document. Try your actual files in the [online Demo](https://docviewkit.com/en/demo/) and check the [supported formats and limits](https://docviewkit.com/docs/supported-formats/). Some legacy iWork paths still return embedded previews; these remain implementation gaps and do not establish native-content rendering or fidelity.

## Get started

Install the Viewer:

```sh
npm install @docviewkit/viewer
```

Place it in your page:

```html
<docviewkit-viewer></docviewkit-viewer>
```

Import it in client code, enable optional formats on demand, and open a `File` supplied by your application:

```js
import "@docviewkit/viewer";

const viewer = document.querySelector("docviewkit-viewer");
viewer.config = {
  engine: {
    formatPack: () => import("@docviewkit/viewer/extended-formats")
      .then(({ extendedFormatPack }) => extendedFormatPack),
  },
};
await viewer.open(file);
```

Only the optional format module needed by the document is loaded. Applications limited to OFD can choose `ofd-formats` instead. Keep the matching Worker, Wasm, font and codec assets with your deployment; the [integration guide](https://docviewkit.com/docs/quickstart/) covers asset hosting and private attachments.

## Works with your frontend

Use the same Viewer in JavaScript, React, Vue and Angular applications, across Chrome, Edge, Firefox and Safari. Compatible WebViews also support Electron, Tauri, Ionic and Capacitor. React Native and Flutter applications can embed it through a WebView with the required browser capabilities; see [runtime compatibility](https://docviewkit.com/docs/browser-compatibility/).

[React and Next.js](https://docviewkit.com/react-office-viewer/) · [Vue](https://docviewkit.com/vue-office-viewer/) · [Angular](https://docviewkit.com/angular-office-viewer/)

## Open source and professional services

DocViewKit Viewer and Engine use [Apache-2.0](LICENSE). All implemented formats, Engine APIs, UI customization, minimal mode and text watermarks are available without runtime licenses, domain registration or mandatory branding. Third-party components and fonts retain their own licenses; preserve [LICENSE](LICENSE), [NOTICE](NOTICE) and [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md) when redistributing the corresponding materials. These terms apply to releases built from this source; older published versions retain their included licenses. Trademark rights are separate.

For integration assistance, document compatibility work, performance tuning, custom interfaces or enterprise deployment, contact [novalag778@gmail.com](mailto:novalag778@gmail.com). Scope, acceptance cases, delivery dates and SLA are agreed for the service.

## Explore further

- [Viewer API](https://docviewkit.com/docs/viewer-api/) and [source location](https://docviewkit.com/docs/source-location/)
- [Performance guide](https://docviewkit.com/docs/performance/) and [accuracy and fidelity](https://docviewkit.com/docs/accuracy-fidelity/)
- [Integration and architecture](https://github.com/docviewkit/viewer/blob/main/docs/ARCHITECTURE.md#integration-reference)
- [Build and contribute](https://github.com/docviewkit/viewer/blob/main/docs/RELEASING.md#local-development)
- [Source and releases](https://github.com/docviewkit/viewer) · [Report a problem](https://github.com/docviewkit/viewer/issues) · [Security policy](https://github.com/docviewkit/viewer/blob/main/docs/SECURITY.md)

Validate your real documents before adopting the Viewer. Share confidential support files only through an agreed private channel.
