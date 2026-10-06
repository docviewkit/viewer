import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { once } from "node:events";
import { connect } from "node:net";
import { createCommercialApp } from "../src/app.mjs";

const publicDir = new URL("../public", import.meta.url).pathname;
const packageVersion = JSON.parse(await readFile(new URL("../../package.json", import.meta.url), "utf8")).version;

async function startApp(t, options = {}) {
  const app = createCommercialApp({
    publicDir,
    ...options,
  });
  app.server.listen(0, "127.0.0.1");
  await once(app.server, "listening");
  t.after(() => app.close());
  const address = app.server.address();
  return { app, baseUrl: `http://127.0.0.1:${address.port}` };
}

test("website removes account surfaces and offers services without registration", async (t) => {
  const { baseUrl } = await startApp(t);
  for (const path of ["/portal/", "/portal/index.html", "/en/portal/", "/zh-cn/portal/", "/en/portal", "/zh-cn/portal", "/assets/portal.js"]) {
    assert.equal((await fetch(`${baseUrl}${path}`)).status, 404, path);
  }
  for (const path of ["/api/auth/register", "/api/auth/login", "/api/auth/logout", "/api/auth/password", "/api/me", "/api/portal/summary", "/api/admin/summary", "/api/apps", "/api/entitlement-requests", "/api/service-requests", "/api/site-license", "/api/licenses", "/api/licenses/lic_00000000-0000-0000-0000-000000000000/renew"]) {
    for (const method of ["GET", "POST"]) {
      const response = await fetch(`${baseUrl}${path}`, { method });
      assert.equal(response.status, 404, `${method} ${path}`);
      assert.equal(response.headers.has("set-cookie"), false, path);
    }
  }
  for (const path of ["/en/", "/zh-cn/", "/en/demo/", "/zh-cn/demo/", "/docs/licensing/", "/llms.txt", "/llms-full.txt"]) {
    const response = await fetch(`${baseUrl}${path}`);
    assert.equal(response.status, 200, path);
    const content = await response.text();
    assert.doesNotMatch(content, /\/portal\/|\/api\/auth\/|DOCVIEWKIT_ADMIN_EMAILS|Customer Portal|客户 Portal|登录|注册/u, path);
    assert.match(content, /mailto:novalag778@gmail\.com/u, path);
  }
});

test("website runs without database storage and preserves historical files", async (t) => {
  const dataDir = await mkdtemp(join(tmpdir(), "docviewkit-retired-storage-"));
  t.after(() => rm(dataDir, { recursive: true, force: true }));
  const filename = join(dataDir, "portal.sqlite");
  const history = Buffer.from("Historical customer records must stay untouched.");
  await writeFile(filename, history);
  const { app, baseUrl } = await startApp(t, { dataDir });
  assert.equal("database" in app, false);
  assert.equal((await fetch(`${baseUrl}/api/health`)).status, 200);
  assert.deepEqual(await readFile(filename), history);
  assert.deepEqual(await readdir(dataDir), ["portal.sqlite"]);
});

test("Apache-2.0 SDK assets stay public without accounts or signing keys", async (t) => {
  const { baseUrl } = await startApp(t);
  assert.equal((await fetch(`${baseUrl}/api/license-public-key`)).status, 404);
  for (const asset of ["index.js", "accuracy.js", "LICENSE", "NOTICE", "THIRD_PARTY_NOTICES.md", "third-party-licenses/pako-MIT.txt"]) {
    const response = await fetch(`${baseUrl}/sdk/${asset}`);
    assert.equal(response.status, 200, asset);
    assert.deepEqual(Buffer.from(await response.arrayBuffer()), await readFile(new URL(`../../dist/${asset}`, import.meta.url)));
  }
});

test("SDK serves only approved bundled fonts and their licenses with versioned caching", async (t) => {
  const { baseUrl } = await startApp(t);
  for (const [family, version] of [["Carlito", "1.103"], ["Caladea", "1.002"]]) {
    for (const variant of ["Regular", "Bold", "Italic", "BoldItalic"]) {
      const name = `${family}-${variant}.ttf`;
      const expected = await readFile(new URL(`../../third_party/${family.toLowerCase()}/${name}`, import.meta.url));
      for (const prefix of ["/sdk/", `/sdk/v${packageVersion}/`]) {
        const response = await fetch(`${baseUrl}${prefix}${name}`);
        assert.equal(response.status, 200, `${prefix}${name}`);
        assert.equal(response.headers.get("content-type"), "font/ttf");
        assert.equal(response.headers.get("cache-control"), prefix === "/sdk/" ? "no-cache" : "public, max-age=31536000, immutable");
        assert.deepEqual(Buffer.from(await response.arrayBuffer()), expected);
      }
    }
    const license = await fetch(`${baseUrl}/sdk/v${packageVersion}/third-party-licenses/${family}-${version}-LICENSE.txt`);
    assert.equal(license.status, 200);
    assert.equal(await license.text(), await readFile(new URL(`../../third_party/${family.toLowerCase()}/LICENSE.txt`, import.meta.url), "utf8"));
  }
  for (const path of ["Unknown-Regular.ttf", "Carlito-Light.ttf", "nested/Carlito-Regular.ttf", "third-party-licenses/Unknown-LICENSE.txt", "v0.0.0/Carlito-Regular.ttf"]) {
    assert.equal((await fetch(`${baseUrl}/sdk/${path}`)).status, 404, path);
  }
});

function jsonLd(html) {
  const source = /<script type="application\/ld\+json">([^<]+)<\/script>/u.exec(html)?.[1];
  assert.ok(source, "missing JSON-LD");
  return JSON.parse(source);
}

test("home exposes every supported extension in visible localized format content", async (t) => {
  const { baseUrl } = await startApp(t);
  const catalog = JSON.parse(await readFile(new URL("../content/docs.json", import.meta.url), "utf8"));
  const formats = catalog.docs.find((doc) => doc.slug === "supported-formats");
  const inputs = formats.sections.find((section) => section.id === "accepted-inputs");
  const extensions = inputs.blocks.find((block) => block.type === "table").rows.flatMap((row) => row[1].split(", "));
  for (const locale of ["en", "zh-cn"]) {
    const html = await fetch(`${baseUrl}/${locale}/`).then((response) => response.text());
    const section = html.match(/<section[^>]*id="formats"[^>]*>([\s\S]*?)<\/section>/)?.[1];
    assert.ok(section, `${locale}: visible format section`);
    const visibleExtensions = new Set(section.match(/\.[a-z][a-z0-9]+/g));
    for (const extension of extensions) assert.ok(visibleExtensions.has(extension), `${locale}: ${extension}`);
    for (const name of ["Word", "Excel", "PowerPoint", "WPS", "PDF", "OFD", "OpenDocument", "Pages", "Numbers", "Keynote"]) {
      assert.ok(section.includes(name), `${locale}: ${name}`);
    }
    assert.match(section, /href="\/docs\/supported-formats\/"/);
    if (locale === "en") assert.doesNotMatch(section, /[\u3400-\u9fff]/u);
  }
});

test("website describes Apache-2.0 source and links its public repository", async (t) => {
  const { baseUrl } = await startApp(t);
  for (const locale of ["en", "zh-cn"]) {
    const html = await fetch(`${baseUrl}/${locale}/`).then((response) => response.text());
    assert.match(html, /Apache-2\.0/);
    assert.match(html, /href="https:\/\/github\.com\/docviewkit\/viewer"/);
    assert.match(html, locale === "en" ? /Viewer and Engine source code is licensed under Apache-2\.0\./ : /Viewer 和 Engine 源码采用 Apache-2\.0。/);
    assert.doesNotMatch(html, /US\$999|商业授权|Request a commercial license/);
  }
  for (const path of ["/docs/licensing/", "/llms-full.txt"]) {
    const text = await fetch(`${baseUrl}${path}`).then((response) => response.text());
    assert.match(text, /Apache-2\.0 from v0\.2\.75/);
    assert.match(text, /previously published packages retain their included licenses/);
  }
});

test("serves landing, human docs, structured docs and AI text", async (t) => {
  const { baseUrl } = await startApp(t, {
    indexNowKey: "docviewkit-test-key",
    siteVerification: { google: "google-test-token", bing: "bing-test-token" },
  });

  assert.equal((await fetch(`${baseUrl}/`, { redirect: "manual" })).status, 308);
  assert.equal((await fetch(`${baseUrl}/`, { redirect: "manual" })).headers.get("location"), "/en/");
  assert.equal((await fetch(`${baseUrl}/docs/?doc=support`, { redirect: "manual" })).headers.get("location"), "/docs/supported-formats/");

  for (const path of ["/en/", "/zh-cn/", "/en/for-ide/", "/zh-cn/for-ide/", "/en/for-browser/", "/zh-cn/for-browser/", "/docs/quickstart/", "/react-office-viewer/", "/vue-office-viewer/", "/angular-office-viewer/", "/en/demo/", "/en/privacy/docviewkit-omni/", "/zh-cn/privacy/docviewkit-omni/", "/en/license/docviewkit-omni/", "/zh-cn/license/docviewkit-omni/"]) {
    const response = await fetch(`${baseUrl}${path}`);
    assert.equal(response.status, 200);
    assert.match(response.headers.get("content-security-policy"), /object-src 'none'/);
    if (path === "/en/demo/") assert.match(response.headers.get("content-security-policy"), /style-src 'self' 'unsafe-inline'/);
    else assert.doesNotMatch(response.headers.get("content-security-policy"), /style-src[^;]*unsafe-inline/);
    assert.match(await response.text(), /DocViewKit/);
  }

  const englishHome = await fetch(`${baseUrl}/en/`).then((response) => response.text());
  assert.match(englishHome, /<html lang="en">/);
  assert.match(englishHome, /<title>Office, WPS, PDF & OFD Document Viewer SDK \| DocViewKit<\/title>/);
  assert.match(englishHome, /lightweight frontend-local document viewer SDK for quick integration into everyday OA attachments/);
  assert.match(englishHome, /href="https:\/\/docviewkit\.com\/en\/"/);
  assert.match(englishHome, /hreflang="zh-CN" href="https:\/\/docviewkit\.com\/zh-cn\/"/);
  assert.match(englishHome, /SoftwareApplication/);
  assert.match(englishHome, new RegExp(`softwareVersion":"${packageVersion.replaceAll(".", "\\.")}"`));
  assert.match(englishHome, /<h1>Lightweight viewing, <br><span>quick integration,<br>for every business\.<\/span><\/h1>/);
  assert.match(englishHome, /href="\/react-office-viewer\/"/);
  assert.match(englishHome, /name="google-site-verification" content="google-test-token"/);
  assert.match(englishHome, /name="msvalidate\.01" content="bing-test-token"/);
  assert.equal(await fetch(`${baseUrl}/docviewkit-test-key.txt`).then((response) => response.text()), "docviewkit-test-key\n");
  assert.ok(jsonLd(englishHome)["@graph"].some((entry) => entry["@type"] === "SoftwareApplication"));

  assert.match(englishHome, /Browsers and compatible cross-platform WebView hosts/);
  assert.match(englishHome, /Electron, Tauri, Ionic and Capacitor/);

  const chineseHome = await fetch(`${baseUrl}/zh-cn/`).then((response) => response.text());
  assert.match(chineseHome, /<title>Office、WPS、PDF、OFD 文档预览 SDK｜DocViewKit<\/title>/);
  assert.match(chineseHome, /轻量级纯前端文档预览 SDK，可快速集成到普通 OA 附件预览/);

  const englishIde = await fetch(`${baseUrl}/en/for-ide/`).then((response) => response.text());
  assert.match(englishIde, /<title>DocViewKit Omni for IDE \| Document preview in your IDE<\/title>/u);
  assert.match(englishIde, /href="\/en\/for-browser\/"/u);
  assert.doesNotMatch(englishIde, /[\u3400-\u9fff]/u);

  const englishBrowser = await fetch(`${baseUrl}/en/for-browser/`).then((response) => response.text());
  assert.match(englishBrowser, /<title>DocViewKit Omni for Browser \| Local document preview<\/title>/u);
  assert.match(englishBrowser, /Chrome Web Store/u);
  assert.doesNotMatch(englishBrowser, /[\u3400-\u9fff]/u);

  const privacyHtml = await fetch(`${baseUrl}/en/privacy/docviewkit-omni/`).then((response) => response.text());
  assert.match(privacyHtml, /<title>DocViewKit Omni Privacy Policy<\/title>/);
  assert.match(privacyHtml, /does not request site-wide access at installation/u);
  assert.match(privacyHtml, /novalag778@gmail\.com/u);

  const licenseHtml = await fetch(`${baseUrl}/en/license/docviewkit-omni/`).then((response) => response.text());
  assert.match(licenseHtml, /<title>DocViewKit Omni End User License Agreement<\/title>/);
  assert.match(licenseHtml, /free does not mean open source/u);
  assert.match(licenseHtml, /JetBrains, and their affiliates are not parties to this Agreement/u);

  const docsHtml = await fetch(`${baseUrl}/docs/quickstart/`).then((response) => response.text());
  assert.match(docsHtml, /<h1>Viewer quickstart<\/h1>/);
  assert.match(docsHtml, /TechArticle/);
  assert.doesNotMatch(docsHtml, /class="skeleton"/);
  assert.ok(jsonLd(docsHtml)["@graph"].some((entry) => entry["@type"] === "TechArticle"));

  const reactHtml = await fetch(`${baseUrl}/react-office-viewer/`).then((response) => response.text());
  assert.match(reactHtml, /<title>React Office Document Viewer \| DocViewKit<\/title>/);
  assert.match(reactHtml, /'use client'/);
  assert.equal((await fetch(`${baseUrl}/react-office-viewer`, { redirect: "manual" })).headers.get("location"), "/react-office-viewer/");

  const robots = await fetch(`${baseUrl}/robots.txt`).then((response) => response.text());
  assert.match(robots, /User-agent: OAI-SearchBot/);
  assert.match(robots, /User-agent: Claude-SearchBot/);
  assert.match(robots, /Sitemap: https:\/\/docviewkit\.com\/sitemap\.xml/);

  const sitemapResponse = await fetch(`${baseUrl}/sitemap.xml`);
  assert.match(sitemapResponse.headers.get("content-type"), /application\/xml/);
  const sitemap = await sitemapResponse.text();
  assert.match(sitemap, /https:\/\/docviewkit\.com\/en\//);
  assert.match(sitemap, /https:\/\/docviewkit\.com\/zh-cn\//);
  assert.match(sitemap, /https:\/\/docviewkit\.com\/en\/for-ide\//);
  assert.match(sitemap, /https:\/\/docviewkit\.com\/zh-cn\/for-browser\//);
  assert.match(sitemap, /https:\/\/docviewkit\.com\/en\/privacy\/docviewkit-omni\//);
  assert.match(sitemap, /https:\/\/docviewkit\.com\/zh-cn\/privacy\/docviewkit-omni\//);
  assert.match(sitemap, /https:\/\/docviewkit\.com\/en\/license\/docviewkit-omni\//);
  assert.match(sitemap, /https:\/\/docviewkit\.com\/zh-cn\/license\/docviewkit-omni\//);
  assert.match(sitemap, /https:\/\/docviewkit\.com\/docs\/supported-formats\//);
  assert.match(sitemap, /https:\/\/docviewkit\.com\/react-office-viewer\//);
  assert.match(sitemap, /https:\/\/docviewkit\.com\/vue-office-viewer\//);
  assert.match(sitemap, /https:\/\/docviewkit\.com\/angular-office-viewer\//);
  assert.doesNotMatch(sitemap, /portal|\?doc=/);

  const www = await fetch(`${baseUrl}/en/`, { headers: { "x-forwarded-host": "www.docviewkit.com" }, redirect: "manual" });
  assert.equal(www.status, 308);
  assert.equal(www.headers.get("location"), "https://docviewkit.com/en/");
  assert.equal((await fetch(`${baseUrl}/index.html`, { redirect: "manual" })).headers.get("location"), "/en/");
  assert.equal((await fetch(`${baseUrl}/for-ide/`, { redirect: "manual" })).headers.get("location"), "/en/for-ide/");
  assert.equal((await fetch(`${baseUrl}/en/for-browser`, { redirect: "manual" })).headers.get("location"), "/en/for-browser/");
  assert.equal((await fetch(`${baseUrl}/privacy/docviewkit-omni/`, { redirect: "manual" })).headers.get("location"), "/en/privacy/docviewkit-omni/");
  assert.equal((await fetch(`${baseUrl}/en/privacy/docviewkit-omni`, { redirect: "manual" })).headers.get("location"), "/en/privacy/docviewkit-omni/");
  assert.equal((await fetch(`${baseUrl}/license/docviewkit-omni/`, { redirect: "manual" })).headers.get("location"), "/en/license/docviewkit-omni/");
  assert.equal((await fetch(`${baseUrl}/en/license/docviewkit-omni`, { redirect: "manual" })).headers.get("location"), "/en/license/docviewkit-omni/");

  const sdkEntry = await fetch(`${baseUrl}/sdk/index.js`);
  assert.equal(sdkEntry.status, 200);

  const viewerEntry = await fetch(`${baseUrl}/sdk/viewer.js`);
  assert.equal(viewerEntry.status, 200);
  assert.equal(viewerEntry.headers.get("cache-control"), "no-cache");
  assert.match(viewerEntry.headers.get("content-type"), /javascript/);
  const viewerSource = await viewerEntry.text();
  assert.match(viewerSource, /defineDocViewKitViewer/);
  const runtimeAsset = viewerSource.match(/\.\/(runtime-[A-Z0-9]+\.js)/u)?.[1];
  assert.ok(runtimeAsset);
  assert.equal((await fetch(`${baseUrl}/sdk/${runtimeAsset}`)).status, 200);
  assert.equal((await fetch(`${baseUrl}/sdk/lcms.wasm`)).status, 200);
  assert.equal((await fetch(`${baseUrl}/sdk/office-viewer-xps.wasm`)).status, 200);
  assert.equal((await fetch(`${baseUrl}/sdk/office-viewer-ofd.wasm`)).status, 200);

  const version = await fetch(`${baseUrl}/sdk/version.json`);
  assert.equal(version.status, 200);
  assert.equal(version.headers.get("cache-control"), "no-store");
  const sdkManifest = await version.json();
  const sdkVersion = sdkManifest.version;
  assert.match(sdkVersion, /^\d+\.\d+\.\d+/u);
  assert.equal(sdkManifest.immutableAssets, true);

  const versionedViewer = await fetch(`${baseUrl}/sdk/v${sdkVersion}/viewer.js`);
  assert.equal(versionedViewer.status, 200);
  assert.equal(versionedViewer.headers.get("cache-control"), "public, max-age=31536000, immutable");
  assert.equal((await fetch(`${baseUrl}/sdk/v${sdkVersion}/office-viewer-xps.wasm`)).status, 200);
  assert.equal((await fetch(`${baseUrl}/sdk/v${sdkVersion}/office-viewer-ofd.wasm`)).status, 200);

  assert.equal((await fetch(`${baseUrl}/api/site-license`)).status, 404);

  const staleViewer = await fetch(`${baseUrl}/sdk/v0.0.0/viewer.js`);
  assert.equal(staleViewer.status, 404);

  const docsResponse = await fetch(`${baseUrl}/api/docs`);
  const docs = { response: docsResponse, payload: await docsResponse.json() };
  assert.equal(docs.response.status, 200);
  assert.ok(docs.payload.docs.length >= 7);
  assert.ok(docs.payload.docs.some((doc) => doc.slug === "issue-reporting"));
  assert.equal(docs.payload.version, packageVersion);
  assert.deepEqual(await fetch(`${baseUrl}/docs.json`).then((response) => response.json()), docs.payload);

  const llms = await fetch(`${baseUrl}/llms.txt`).then((response) => response.text());
  const llmFull = await fetch(`${baseUrl}/llm-full.txt`).then((response) => response.text());
  const llmsFull = await fetch(`${baseUrl}/llms-full.txt`).then((response) => response.text());
  assert.match(llms, /Machine-readable resources/);
  assert.match(llms, new RegExp(`${baseUrl.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}/docs\\.json`));
  assert.match(llmsFull, /# Complete documentation/);
  assert.match(llmsFull, /github\.com\/docviewkit\/viewer\/issues\/new\/choose/);
  assert.doesNotMatch(llmsFull, /site-license|\/api\/licenses\/:licenseId\/renew/);
  assert.match(llmsFull, /No website account or application registration is required/);
  assert.match(llmsFull, /mailto:novalag778@gmail\.com/);
  assert.equal(llmFull, llmsFull);
});

test("selects the browser language once and then prefers the language cookie", async (t) => {
  const { baseUrl } = await startApp(t);
  const redirectOptions = { redirect: "manual" };

  const chineseFirstVisit = await fetch(`${baseUrl}/`, {
    ...redirectOptions,
    headers: { "accept-language": "en-US;q=0.4, zh-CN;q=0.9" },
  });
  assert.equal(chineseFirstVisit.headers.get("location"), "/zh-cn/");

  const chinesePage = await fetch(`${baseUrl}/zh-cn/`);
  const languageCookie = chinesePage.headers.get("set-cookie")?.split(";", 1)[0];
  assert.equal(languageCookie, "docviewkit-language=zh-CN");
  const rememberedVisit = await fetch(`${baseUrl}/`, {
    ...redirectOptions,
    headers: { cookie: languageCookie, "accept-language": "en-US" },
  });
  assert.equal(rememberedVisit.headers.get("location"), "/zh-cn/");

  const englishFirstVisit = await fetch(`${baseUrl}/demo/`, {
    ...redirectOptions,
    headers: { "accept-language": "fr-FR, en-US;q=0.8" },
  });
  assert.equal(englishFirstVisit.headers.get("location"), "/en/demo/");
});

test("rejects a malformed Host header without escaping the request error boundary", async (t) => {
  const { app } = await startApp(t);
  const address = app.server.address();
  const response = await new Promise((resolve, reject) => {
    const socket = connect(address.port, "127.0.0.1");
    let data = "";
    socket.setEncoding("utf8");
    socket.setTimeout(1_000, () => socket.destroy(new Error("server did not respond")));
    socket.on("data", (chunk) => { data += chunk; });
    socket.on("end", () => resolve(data));
    socket.on("error", reject);
    socket.end("GET / HTTP/1.1\r\nHost: [\r\nConnection: close\r\n\r\n");
  });
  assert.match(response, /^HTTP\/1\.1 400 /u);
});

test("development SDK routes bypass immutable version caches", async (t) => {
  const { baseUrl } = await startApp(t, { immutableSdkAssets: false });
  const manifest = await fetch(`${baseUrl}/sdk/version.json`).then((response) => response.json());
  assert.equal(manifest.immutableAssets, false);

  const versionedViewer = await fetch(`${baseUrl}/sdk/v${manifest.version}/viewer.js`);
  assert.equal(versionedViewer.status, 200);
  assert.equal(versionedViewer.headers.get("cache-control"), "no-cache");
});

test("localized landing answers OA integration questions with matching FAQ structured data", async (t) => {
  const { baseUrl } = await startApp(t);
  for (const locale of ["en", "zh-cn"]) {
    const html = await fetch(`${baseUrl}/${locale}/`).then((response) => response.text());
    const faq = jsonLd(html)["@graph"].find((entry) => entry["@type"] === "FAQPage");
    assert.ok(faq, `${locale}: visible FAQs must have matching structured data`);
    const visible = [...html.matchAll(/<details><summary>([^<]+)<\/summary><p>([^<]+)<\/p><\/details>/gu)];
    assert.equal(visible.length, 6);
    assert.deepEqual(faq.mainEntity, visible.map(([, name, answer]) => ({
      "@type": "Question", name, acceptedAnswer: { "@type": "Answer", text: answer },
    })));
    const lead = /<div class="hero-copy">[\s\S]*?<p class="lede">([^<]+)<\/p>/u.exec(html)?.[1];
    assert.ok(lead.includes("OA"));
    assert.ok((locale === "en" ? lead.split(/\s+/u).length : [...lead].length) <= 60);
    for (const [, , answer] of visible) {
      assert.ok((locale === "en" ? answer.split(/\s+/u).length : [...answer].length) <= 150);
    }
    for (const path of ["/docs/quickstart/", "/docs/performance/", "/docs/supported-formats/", "/docs/licensing/"]) {
      assert.ok(html.includes(`href="${path}"`), `${locale}: link to evidence ${path}`);
    }
    if (locale === "en") assert.doesNotMatch(html, /[\u3400-\u9fff]/u);
  }
});
