import test from "node:test";
import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFile, readdir } from "node:fs/promises";
import { getDocsCatalog, renderDocsPage, renderLlmsFull, renderLlmsIndex } from "../src/docs.mjs";

const root = new URL("..", import.meta.url);

async function text(path) {
  return readFile(new URL(path, root), "utf8");
}

test("project READMEs describe the open source package and direct service contact", async () => {
  for (const path of ["../README.md", "../release/FREE_VIEWER_README.md", "README.md"]) {
    const content = await text(path);
    assert.match(content, /Apache-2\.0/, path);
    assert.match(content, /https:\/\/github\.com\/docviewkit\/viewer/, path);
    assert.match(content, /mailto:novalag778@gmail\.com/, path);
    assert.match(content, /older published (?:packages|versions) retain their included licenses/i, path);
    assert.doesNotMatch(content, /\/portal\/|DOCVIEWKIT_ADMIN_EMAILS|DOCVIEWKIT_ENABLE_PREVIEW_PORTAL|public package is an engine/, path);
  }
  const source = await text("../README.md");
  assert.match(source, /git clone https:\/\/github\.com\/docviewkit\/viewer\.git/);
  assert.match(source, /npm ci/);
  const publicPackage = await text("../release/FREE_VIEWER_README.md");
  assert.match(publicPackage, /<docviewkit-viewer><\/docviewkit-viewer>/);
});

test("documentation identifiers and anchors are unique", async () => {
  const catalog = JSON.parse(await text("content/docs.json"));
  const slugs = catalog.docs.map((doc) => doc.slug);
  assert.equal(new Set(slugs).size, slugs.length);
  for (const doc of catalog.docs) {
    const sections = doc.sections.map((section) => section.id);
    assert.equal(new Set(sections).size, sections.length, `duplicate section id in ${doc.slug}`);
    assert.ok(doc.summary.length >= 30);
  }
});

test("every page has mobile viewport, skip navigation and external scripts only", async () => {
  const pages = new Map([
    ["public/index.html", "https://docviewkit.com/en/"],
    ["public/for-ide/index.html", "https://docviewkit.com/en/for-ide/"],
    ["public/for-browser/index.html", "https://docviewkit.com/en/for-browser/"],
    ["public/demo/index.html", "https://docviewkit.com/en/demo/"],
    ["public/privacy/docviewkit-omni/index.html", "https://docviewkit.com/en/privacy/docviewkit-omni/"],
    ["public/license/docviewkit-omni/index.html", "https://docviewkit.com/en/license/docviewkit-omni/"],
  ]);
  for (const [path, canonical] of pages) {
    const html = await text(path);
    assert.match(html, /name="viewport" content="width=device-width, initial-scale=1"/);
    assert.match(html, /class="skip-link"/);
    assert.match(html, new RegExp(`<link rel="canonical" href="${canonical}">`));
    assert.doesNotMatch(html, /<script(?![^>]*src=)/);
    assert.doesNotMatch(html, /<img[^>]*(?!alt=)[^>]*>/);
  }
});

test("Omni product pages describe real capabilities and current installation paths", async () => {
  const ide = await text("public/for-ide/index.html");
  assert.match(ide, /VS Code 与 JetBrains IDE/u);
  assert.match(ide, /marketplace\.visualstudio\.com\/items\?itemName=docviewkit\.docviewkit-omni/u);
  assert.match(ide, /code --install-extension docviewkit\.docviewkit-omni/u);
  assert.match(ide, /class="availability available">现已提供<\/span><h3>JetBrains IDE/u);
  assert.match(ide, /插件已通过 JetBrains Marketplace 审核，可直接在 IDE 的插件市场中搜索并安装。/u);
  assert.doesNotMatch(ide, /商店审核中|通过审核后|未经商店签名/u);
  assert.doesNotMatch(ide, /plugins\.jetbrains\.com\/plugin\//u);

  const browser = await text("public/for-browser/index.html");
  assert.match(browser, /chromewebstore\.google\.com\/detail\/docviewkit-omni\/mfjdlohbjkajelkffaamkcfjpdbbmfil/u);
  assert.match(browser, /docviewkit-omni-browser\/releases\/tag\/v0\.1\.11/u);
  assert.match(browser, /按需来源权限/u);
  for (const name of ["Microsoft Edge", "Firefox"]) {
    assert.ok(browser.includes(`class="availability available">现已提供</span><h3>${name}</h3>`), name);
  }
  assert.doesNotMatch(browser, /商店发布中|尚未生效|下载测试包|使用 Chrome 商店安装/u);
  assert.match(browser, /macOS 版准备中/u);
});

test("public surfaces use the registered DocViewKit brand consistently", async () => {
  const retiredBrand = ["Folio", "Lens"].join("");
  const retiredFreeTier = ["Free", "Viewer"].join(" ");
  for (const path of [
    "public/index.html",
    "public/demo/index.html",
    "public/privacy/docviewkit-omni/index.html",
    "public/license/docviewkit-omni/index.html",
    "public/assets/i18n.js",
    "content/docs.json",
  ]) {
    const content = await text(path);
    assert.match(content, /DocViewKit/);
    assert.ok(!content.includes(retiredBrand));
    assert.ok(!content.toLowerCase().includes(retiredBrand.toLowerCase()));
    assert.ok(!content.includes(retiredFreeTier));
  }
});

test("landing states the free, high-performance, lightweight and compatible product contract", async () => {
  const html = await text("public/index.html");
  assert.match(html, /Office、WPS、PDF、OFD、OpenDocument 和 iWork/);
  assert.match(html, /全部格式免费/);
  assert.match(html, /高性能本地渲染/);
  assert.match(html, /轻量按需加载/);
  assert.match(html, /多种前端环境兼容/);
  assert.match(html, /浏览器与跨平台应用能共用一套 Viewer 吗/);
  assert.match(html, /Electron、Tauri、Ionic 和 Capacitor/);
  assert.match(html, /WebView/);
  assert.match(html, /<h3>如何添加文字水印？<\/h3>/);
  assert.doesNotMatch(html, /viewer\.config = \{ watermark:/);
  const i18n = await text("public/assets/i18n.js");
  assert.match(i18n, /Office, WPS, PDF, OFD, OpenDocument and iWork/);
  assert.match(i18n, /Consistent UI across browsers/);
  assert.match(i18n, /Can browsers and cross-platform apps share one Viewer/);
  assert.match(i18n, /Electron, Tauri, Ionic and Capacitor/);
  assert.match(i18n, /required Web APIs/);
  assert.match(i18n, /Validate a real file/);
  assert.match(i18n, /Which formats does the free version support/);
  assert.match(i18n, /How do I add a text watermark/);

  const sdkPackage = JSON.parse(await text("../package.json"));
  assert.deepEqual(sdkPackage.peerDependencies ?? {}, {});
});

test("landing leads with lightweight integration and covers everyday and professional scenarios", async () => {
  const html = await text("public/index.html");
  const i18n = await text("public/assets/i18n.js");
  assert.match(html, /轻量预览，<br><span>快速集成，<br>融入每一种业务。<\/span>/u);
  for (const scenario of [
    "合同审阅",
    "AI 知识库",
    "普通 OA",
    "网盘",
    "审批流、后台管理与内部工具",
    "CRM、ERP 与人力资源系统",
    "电子政务与合规系统",
    "教育与在线学习平台",
    "工业、建筑与供应链管理",
    "电子商务与跨境贸易",
    "法律、审计与金融风控",
  ]) assert.match(html, new RegExp(scenario, "u"));
  assert.match(i18n, /"轻量预览，": "Lightweight viewing, "/u);
  assert.match(i18n, /"快速集成，": "quick integration,"/u);
  assert.match(i18n, /"融入每一种业务。": "for every business\."/u);
  assert.match(i18n, /OA and everyday attachments/u);
  assert.match(i18n, /legal work, audit, and financial risk management/u);
});

test("design system includes accessibility and responsive contracts", async () => {
  const css = await text("public/assets/styles.css");
  assert.match(css, /:focus-visible/);
  assert.match(css, /prefers-reduced-motion/);
  assert.match(css, /min-height: 44px/);
  assert.match(css, /@media \(max-width: 620px\)/);
  assert.match(css, /\[data-theme="dark"\]/);
  assert.match(css, /\.docs-nav a\[hidden\]/);
});

test("localized marketing surfaces expose one persistent language switch", async () => {
  for (const path of ["public/index.html", "public/demo/index.html", "public/privacy/docviewkit-omni/index.html", "public/license/docviewkit-omni/index.html"]) {
    const html = await text(path);
    assert.match(html, /data-language-toggle/);
  }
  const i18n = await text("public/assets/i18n.js");
  assert.match(i18n, /Office, WPS, PDF & OFD Document Viewer SDK \| DocViewKit/);
  assert.match(i18n, /Developer Documentation/);
  assert.match(i18n, /docviewkit-language/);
  assert.match(i18n, /pathLocale = location\.pathname\.match/);
  assert.match(i18n, /document\.cookie/);
  assert.match(i18n, /navigator\.languages/);
  assert.doesNotMatch(i18n, /localStorage/);
});

test("DocViewKit Omni privacy policy matches the extension data boundary", async () => {
  const html = await text("public/privacy/docviewkit-omni/index.html");
  assert.match(html, /不会向 DocViewKit 或我们的服务器上传、发送或出售文档内容/u);
  assert.match(html, /精确来源/u);
  assert.match(html, /不会在安装时申请全站访问权限/u);
  assert.match(html, /不会自动发送给我们/u);
  assert.match(html, /不包含广告、遥测或远程运行时代码/u);
  assert.match(html, /novalag778@gmail\.com/u);
});

test("DocViewKit Omni EULA grants free proprietary use with marketplace boundaries", async () => {
  const html = await text("public/license/docviewkit-omni/index.html");
  assert.match(html, /软件保持专有和闭源；免费不表示开源/u);
  assert.match(html, /用于个人或内部业务目的/u);
  assert.match(html, /不授予源代码许可/u);
  assert.match(html, /Microsoft、GitHub、JetBrains 及其关联方不是本协议的一方/u);
  assert.match(html, /按“现状”和“可用”状态提供/u);
  assert.match(html, /privacy\/docviewkit-omni/u);
  assert.match(html, /novalag778@gmail\.com/u);
});

test("documentation renders complete canonical HTML from the shared versioned catalog", async () => {
  const sdkPackage = JSON.parse(await text("../package.json"));
  const catalog = getDocsCatalog();
  assert.equal(catalog.version, sdkPackage.version);
  assert.ok(catalog.docs.some((doc) => doc.slug === "supported-formats"));
  assert.ok(catalog.docs.some((doc) => doc.slug === "accuracy-fidelity"));
  const performance = catalog.docs.find((doc) => doc.slug === "performance");
  assert.deepEqual(performance.sections.map((section) => section.id), [
    "first-open", "loading-boundaries", "measurement",
  ]);
  assert.match(JSON.stringify(performance), /viewer\.config\.engine\.formatPack/);
  assert.match(JSON.stringify(performance), /does not fetch every Wasm binary/);
  const extensionTable = catalog.docs
    .find((doc) => doc.slug === "supported-formats")
    .sections.find((section) => section.id === "accepted-inputs")
    .blocks.find((block) => block.type === "table");
  assert.deepEqual(extensionTable.headers, ["Family", "Accepted file extensions", "Runtime path"]);
  assert.deepEqual(extensionTable.rows.map((row) => row[1]), [
    ".pptx, .pptm, .ppsx, .ppsm, .potx, .potm", ".xlsx, .xlsm, .xltx, .xltm",
    ".docx, .docm, .dotx, .dotm", ".csv, .rtf",
    ".odp, .otp, .fodp, .odg, .otg, .fodg, .ods, .ots, .fods, .odt, .ott, .fodt", ".pages, .numbers, .key",
    ".doc, .xls, .ppt", ".wps, .et, .dps", ".pdf, .xps, .oxps", ".ofd",
  ]);
  assert.match(JSON.stringify(catalog.docs.find((doc) => doc.slug === "supported-formats")), /OFD 1\.0 and 1\.1/);

  const demo = await text("public/demo/index.html");
  assert.match(demo, /accept="[^"]*\.ofd/);
  assert.match(demo, /Office、PDF、OFD、ODF/);

  for (const doc of catalog.docs) {
    const html = renderDocsPage("https://docviewkit.com", doc.slug);
    const path = doc.path || `/docs/${doc.slug}/`;
    assert.match(html, new RegExp(`<link rel="canonical" href="https://docviewkit\\.com${path}">`));
    assert.match(html, new RegExp(`<h1>${doc.title.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}</h1>`));
    assert.match(html, /application\/ld\+json/);
    assert.match(html, /TechArticle/);
    assert.match(html, /data-docs-menu/);
    assert.match(html, /id="docs-search-input"/);
    assert.doesNotMatch(html, /class="skeleton"|fetch\("\/api\/docs"\)/);
  }

  const quickstart = renderDocsPage("https://docviewkit.com", "quickstart");
  assert.match(quickstart, /import '@docviewkit\/viewer'/);
  assert.match(quickstart, /data-docs-copy/);
  assert.doesNotMatch(quickstart, /@docviewkit\/sdk\/viewer/);
  assert.match(renderLlmsIndex("https://docviewkit.com"), /https:\/\/docviewkit\.com\/docs\/quickstart\//);
  assert.match(renderLlmsIndex("https://docviewkit.com"), /https:\/\/docviewkit\.com\/react-office-viewer\//);
  assert.match(renderLlmsIndex("https://docviewkit.com"), /https:\/\/docviewkit\.com\/docs\.json/);
  assert.match(renderLlmsFull("https://docviewkit.com"), /viewer\.config\.engine\.formatPack/);
});

test("framework guides are distinct, source-backed integration tasks", () => {
  const catalog = getDocsCatalog();
  const guides = ["react-office-viewer", "vue-office-viewer", "angular-office-viewer"]
    .map((slug) => catalog.docs.find((doc) => doc.slug === slug));
  assert.ok(guides.every(Boolean));
  assert.deepEqual(guides.map((doc) => doc.path), [
    "/react-office-viewer/", "/vue-office-viewer/", "/angular-office-viewer/",
  ]);

  const react = renderDocsPage("https://docviewkit.com", "react-office-viewer");
  assert.match(react, /<title>React Office Document Viewer \| DocViewKit<\/title>/u);
  assert.match(react, /useEffect/u);
  assert.match(react, /'use client'/u);
  assert.match(react, /import\('@docviewkit\/viewer'\)/u);

  const vue = renderDocsPage("https://docviewkit.com", "vue-office-viewer");
  assert.match(vue, /isCustomElement/u);
  assert.match(vue, /watchEffect/u);

  const angular = renderDocsPage("https://docviewkit.com", "angular-office-viewer");
  assert.match(angular, /CUSTOM_ELEMENTS_SCHEMA/u);
  assert.match(angular, /NO_ERRORS_SCHEMA is not required/u);

  for (const html of [react, vue, angular]) {
    assert.match(html, /@docviewkit\/viewer/u);
    assert.match(html, /\/docs\/quickstart\//u);
    assert.match(html, /application\/ld\+json/u);
  }
});

test("website sends issue feedback to the public GitHub tracker", async () => {
  const landing = await text("public/index.html");
  assert.match(landing, /href="https:\/\/github\.com\/docviewkit\/viewer\/issues\/new\/choose" target="_blank" rel="noopener noreferrer"/u);
});

test("online demo embeds the public Viewer instead of duplicating its controls", async () => {
  const html = await text("public/demo/index.html");
  const script = await text("public/assets/demo.js");
  assert.match(html, /<docviewkit-viewer id="demo-viewer"/);
  assert.match(script, /fetch\("\/sdk\/version\.json", \{ cache: "no-store" \}\)/);
  assert.match(script, /sdkManifest\.immutableAssets === false/);
  assert.match(script, /const viewerPromise = import\(`\$\{sdkBase\}viewer\.js`\)/);
  assert.doesNotMatch(script, /site-license|websiteLicense|license:/);
  assert.match(script, /import\(`\$\{sdkBase\}extended-formats\.js`\)/);
  assert.match(script, /FORMAT_PACK_CANDIDATES = Object\.freeze\(\["odf", "iwork", "legacy-office", "wps", "pdf", "xps"\]\)/);
  assert.match(script, /WebAssembly\.compile\(await response\.arrayBuffer\(\)\)/);
  assert.match(script, /return modules\.get\(String\(source\)\) \?\? source/);
  assert.match(script, /wasm: runtime\.wasm/);
  assert.match(script, /formatPack: \(\) => Promise\.resolve\(runtime\.formatPack\)/);
  assert.match(script, /ui\.viewer\.open\(file,/);
  assert.match(script, /ui\.viewer\.config =/);
  assert.doesNotMatch(html, /demo-viewer-mode|demo-mode-switcher/);
  assert.doesNotMatch(script, /modeSwitcher|demo-viewer-mode/);
  assert.match(script, /features: \{ interactionMode: "object", interactionModeSwitcher: true \}/);
  assert.match(script, /DocViewKit-visual-sample\.pptx/);
  assert.doesNotMatch(`${html}\n${script}`, /DocViewKit-视觉示例\.pptx/);
  const encodedSample = /ENGLISH_PPTX_SAMPLE_BASE64 = "([^"]+)"/u.exec(script)?.[1];
  assert.ok(encodedSample);
  assert.equal(createHash("sha256").update(Buffer.from(encodedSample, "base64")).digest("hex"), "efd1e2569508910946812831fa1f671eaf2cb1343171a7e05f1f9ebc8032ef81");
  assert.match(script, /application\/vnd\.openxmlformats-officedocument\.presentationml\.presentation/);
  assert.doesNotMatch(script, /text\/csv|\.csv"/);
  assert.doesNotMatch(html, /id="demo-canvas"|id="demo-search-form"|id="demo-unit-list"/);
  assert.doesNotMatch(script, /createOfficeEngine|document\.render|searchText/);
});

test("landing has one primary demo action", async () => {
  const html = await text("public/index.html");
  assert.match(html, /href="\/demo\/"/u);
  const hero = html.match(/<div class="hero-actions">([\s\S]*?)<\/div>/u)?.[1] || "";
  assert.equal((hero.match(/href="\/demo\/"/gu) || []).length, 1);
  assert.match(hero, /验证真实文件/);
  assert.doesNotMatch(hero, /在线体验/);
});

test("Enterprise contact uses mailto and displays the address", async () => {
  const html = await text("public/index.html");
  assert.match(html, /href="mailto:novalag778@gmail\.com\?subject=DocViewKit%20Enterprise"/u);
  assert.match(html, /<span class="small">novalag778@gmail\.com<\/span>/u);
});

test("landing states fast location, framework support and stable LLM resource names", async () => {
  const html = await text("public/index.html");
  const i18n = await text("public/assets/i18n.js");
  assert.match(html, /<h3>如何快速定位原文？<\/h3>/u);
  assert.match(html, /快速接入所有主流前端框架/u);
  assert.match(html, /原生 JavaScript、React、Vue、Angular/u);
  assert.match(html, /纯前端文档预览 SDK/u);
  assert.match(html, /href="\/react-office-viewer\/"/u);
  assert.match(html, /href="\/vue-office-viewer\/"/u);
  assert.match(html, /href="\/angular-office-viewer\/"/u);
  assert.match(html, /href="\/llm-full\.txt">llm-full\.txt<\/a>/u);
  assert.match(i18n, /every major frontend framework/u);
  assert.doesNotMatch(`${html}\n${i18n}`, /答案直接回到原文|完整纯文本/u);
});

test("Enterprise offers format-scoped performance customization", async () => {
  const html = await text("public/index.html");
  const i18n = await text("public/assets/i18n.js");
  const catalog = JSON.stringify(getDocsCatalog());
  assert.match(html, /企业性能优化定制：按实际格式裁剪未使用模块，缩小交付体积并优化首开性能/u);
  assert.match(i18n, /Custom enterprise performance optimization: trim unused modules/u);
  assert.match(catalog, /format-specific build can omit unused format modules/u);
  assert.match(renderLlmsFull("https://docviewkit.com"), /contract-scoped performance customization/u);
});

test("open source capability lists include minimal rendering-only mode", async () => {
  const html = await text("public/index.html");
  const catalog = JSON.stringify(getDocsCatalog());
  assert.match(html, /界面定制、极简模式和文字水印/u);
  assert.match(catalog, /minimal mode with only the document rendering workspace/u);
});

test("public copy states Apache-2.0 rights and optional service scope", async () => {
  const publicRoot = new URL("public/", root);
  const publicFiles = (await readdir(publicRoot, { recursive: true }))
    .filter((path) => /\.(?:html|js)$/u.test(path));
  const content = await Promise.all([
    ...publicFiles.map((path) => readFile(new URL(path, publicRoot), "utf8")),
    text("content/docs.json"),
    text("README.md"),
  ]).then((parts) => parts.join("\n"));
  assert.doesNotMatch(content, /完全免费用于生产|Completely free (?:for|production)|免费生产/iu);
  assert.doesNotMatch(content, /自动删除|automatic deletion|deleted after 7.?30 days/iu);
  assert.doesNotMatch(content, /源文件始终是可选|source file is always optional|only when you choose to attach/iu);
  assert.doesNotMatch(content, /\bpage:\s*6/);
  assert.match(content, /Apache-2\.0/iu);
  assert.match(content, /支持、定制和企业交付|support, customization and enterprise delivery/iu);
  assert.match(content, /@docviewkit\/viewer/);
  assert.match(content, /无[需须]许可证|no runtime license/iu);
  assert.doesNotMatch(content, /Apache-2\.0 开源版本正在准备|Source publication is being prepared|upcoming release/iu);
  assert.match(content, /Apache-2\.0 from v0\.2\.75/);
  const { english } = await import("../public/assets/i18n.js");
  assert.equal(english.get("Viewer 和 Engine 源码采用 Apache-2.0。"), "Viewer and Engine source code is licensed under Apache-2.0.");
  assert.doesNotMatch(content, /免费许可证|create (?:and renew a )?free licen[cs]e|site-license/iu);
});

test("public introductions cover OFD and capable frontend hosts", async () => {
  for (const path of ["../README.md", "../release/FREE_VIEWER_README.md", "README.md", "../release/pages/index.html", "content/docs.json"]) {
    const content = await text(path);
    for (const term of ["OFD", "Electron", "Tauri", "WebView"]) assert.ok(content.includes(term), `${path}: ${term}`);
  }
  const runtime = renderDocsPage("https://docviewkit.com", "browser-compatibility");
  for (const term of ["OffscreenCanvas", "FontFace", "React Native", "Flutter", "WebView"]) assert.ok(runtime.includes(term), term);
  const llms = renderLlmsFull("https://docviewkit.com");
  assert.match(llms, /frontend-local/);
});


test("every product entry positions lightweight viewing for everyday OA attachments", async () => {
  for (const path of ["../README.md", "../release/FREE_VIEWER_README.md", "README.md", "../release/pages/index.html", "public/index.html"]) {
    const content = await text(path);
    assert.match(content, /OA/, path);
    assert.match(content, /轻量|lightweight/iu, path);
    assert.match(content, /快速|quick/iu, path);
  }
  const { english } = await import("../public/assets/i18n.js");
  const html = await text("public/index.html");
  for (const copy of html.matchAll(/>([^<>]*OA[^<>]*)</gu)) {
    assert.match(english.get(copy[1]) || "", /OA/, `English translation: ${copy[1]}`);
  }
  assert.match(html.match(/<meta name="description" content="([^"]+)"/u)[1], /OA/u);
  assert.match(html.match(/<div class="hero-copy">([\s\S]*?)<div class="hero-actions">/u)[1], /OA/u);
  for (const output of [renderDocsPage("https://docviewkit.com", "overview"), renderLlmsIndex("https://docviewkit.com"), renderLlmsFull("https://docviewkit.com")]) {
    assert.match(output, /OA/);
    assert.match(output, /lightweight/iu);
    assert.match(output, /quick/iu);
  }
});

test("OA integration evidence is available in both human and AI documentation", () => {
  const overview = renderDocsPage("https://docviewkit.com", "overview");
  const quickstart = renderDocsPage("https://docviewkit.com", "quickstart");
  const full = renderLlmsFull("https://docviewkit.com");
  for (const phrase of ["1 Viewer component", "4 input types", "0 conversion services", "cold versus warm cache", "Private URLs do not need to become public"]) {
    for (const output of [overview, full]) assert.ok(output.includes(phrase), phrase);
  }
  for (const phrase of ["How do I preview an authenticated OA attachment?", "/api/attachments/42/content", "credentials: 'same-origin'", "if (!response.ok)", "await viewer.open(attachment)"]) {
    for (const output of [quickstart, full]) assert.ok(output.includes(phrase), phrase);
  }
  assert.ok(overview.includes('"@type":"TechArticle"'));
});

test("marketing copy leads with capabilities and helpful next steps", async () => {
  for (const path of ["public/index.html", "public/for-ide/index.html", "public/for-browser/index.html", "public/demo/index.html"]) {
    const html = await text(path);
    assert.ok(!/不会|失败|不等于|不代表|不提供|不需要|不支持|不能/u.test(html), path);
  }
  const demo = await text("public/demo/index.html");
  const { english } = await import("../public/assets/i18n.js");
  assert.equal(english.get("浏览器本地"), "Browser-local");
  assert.ok(demo.includes("文件在本地处理"));
  assert.ok(demo.includes("体验文档的字体、分页、图表与搜索效果"));
});
