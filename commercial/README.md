# DocViewKit website

DocViewKit is a lightweight document viewer for everyday OA attachment previews, approval workflows, admin portals, CRM/ERP systems, cloud drives, and customer-facing apps, as well as enterprise SaaS, AI knowledge bases, legal review, and financial audit. Quick integration starts with the ready-made Viewer: import the component, mount `<docviewkit-viewer>`, and call `open(file)`. A compact core and on-demand format packs keep unused parsers out of the initial load, with no document-conversion server to deploy. Use the Engine API when you need custom rendering or source-object access.

This website source is part of [docviewkit/viewer](https://github.com/docviewkit/viewer) and uses [Apache-2.0](../LICENSE). It presents the open source Viewer and Engine plus optional support, customization and enterprise delivery. These terms apply to releases built from this source; older published versions retain their included licenses. DocViewKit Omni retains its separately documented product license.

This directory is intentionally isolated from the OfficeViewer SDK, Viewer examples, and format engine. It contains only the public web surface:

- `/en/` and `/zh-cn/` — localized landing pages
- `/en/for-ide/` and `/zh-cn/for-ide/` — DocViewKit Omni for IDE
- `/en/for-browser/` and `/zh-cn/for-browser/` — DocViewKit Omni for Browser
- `/en/privacy/docviewkit-omni/` and `/zh-cn/privacy/docviewkit-omni/` — browser-extension privacy policy
- `/en/license/docviewkit-omni/` and `/zh-cn/license/docviewkit-omni/` — Omni product license
- `/en/demo/` and `/zh-cn/demo/` — browser-local Viewer demo
- `/docs/<slug>/` — server-rendered, canonical documentation pages
- `/robots.txt`, `/sitemap.xml`, and `/docs.json` — crawler and machine-readable discovery
- `/llms.txt` — AI-oriented documentation index
- `/llms-full.txt` — complete plain-text documentation

The website is intended for `https://docviewkit.com`, with the same paths as the local application. DNS, TLS, production deployment, and service agreements are separate release gates.

The product supports Office, WPS, PDF, OFD 1.0/1.1, OpenDocument, iWork, XPS, CSV, and RTF with local frontend processing. OFD uses the on-demand `ofd-formats` pack; the published supported-formats page describes its implemented subset and signature-integrity limits.

The Viewer runs in frontend environments that provide DOM/Custom Elements, module Workers, WebAssembly, Canvas/OffscreenCanvas, ImageBitmap, and FontFace. These include browsers and compatible WebView hosts in Electron, Tauri, Ionic and Capacitor. React Native and Flutter applications can embed it through a WebView that provides the same APIs; their native rendering layers do not run the Web Component directly. Validate the actual host engine, asset loading, CSP, fonts, and required interactions on each target platform.

## Run

From the repository root, build the SDK assets before starting the online demo:

```sh
npm ci
npm run build
npm --prefix commercial run dev
```

The default address is `http://127.0.0.1:4310`. Override with `PORT` or `HOST`. The website needs no database, user accounts or authentication service.

Production may set `DOCVIEWKIT_GOOGLE_SITE_VERIFICATION`, `DOCVIEWKIT_BING_SITE_VERIFICATION`, and `DOCVIEWKIT_INDEXNOW_KEY`. The first two add the official webmaster verification meta tags to both localized home pages. The IndexNow key is served at `/<key>.txt`; after deployment, submit every canonical URL with:

```sh
DOCVIEWKIT_INDEXNOW_KEY=replace-with-production-key npm run indexnow
```

## Search verification after deployment

After publishing, inspect `/en/`, `/zh-cn/`, `/docs/overview/`, and `/docs/quickstart/` in Google Search Console. Check the live URL, selected canonical, indexability, and last crawl; request indexing for the changed pages and confirm `/sitemap.xml` is submitted. IndexNow is separate from Google's URL Inspection workflow.

Save a pre-release baseline, then review daily for the first 48 hours: indexing status and the Search results report filtered by page and target queries such as `JavaScript document viewer`, `OA attachment preview`, and `纯前端文档预览`. Compare impressions, clicks, CTR, and average position using the same country/device filters. Also record page impressions in the Generative AI performance report when available; it is not a separate AI-click attribution report. Continue over a longer comparable window if data is sparse.

The localized home derives FAQPage from its six visible plain-text answers; documentation already emits TechArticle. These describe content, not ranking guarantees. Google discontinued FAQ rich results in May 2026, and recrawling can take days to weeks. After 48 hours without a change, check crawl/index status and data availability before drawing conclusions about content or Schema.

Sources: [Google AI optimization guidance](https://developers.google.com/search/docs/fundamentals/ai-optimization-guide), [FAQ rich result retirement](https://developers.google.com/search/updates#may-2026), [Generative AI report](https://support.google.com/webmasters/answer/16984139), [recrawl guidance](https://developers.google.com/search/docs/crawling-indexing/ask-google-to-recrawl).

## Verify

```sh
cd commercial
npm run verify
```

Support, customization and enterprise delivery are handled directly at [novalag778@gmail.com](mailto:novalag778@gmail.com). Public issues go to [GitHub](https://github.com/docviewkit/viewer/issues). The website does not accept document uploads. Retired customer data under `commercial/data/` remains ignored by Git; the website no longer opens or modifies it.

The Viewer and Engine are Apache-2.0 open source and require no runtime license. Paid services cover support, customization and enterprise delivery. No registration, application record or runtime license is required. The official Demo uses the same Viewer. Every lazy format pack used by the Demo, including `office-viewer-xps.wasm` and `office-viewer-ofd.wasm`, must be present in the public SDK asset allowlist.
