import { createServer } from "node:http";
import { readFile } from "node:fs/promises";
import { existsSync, readFileSync } from "node:fs";
import { extname, normalize, resolve } from "node:path";
import { getDocsCatalog, getIndexablePaths, renderDocsPage, renderLlmsFull, renderLlmsIndex } from "./docs.mjs";
import { english } from "../public/assets/i18n.js";

const PUBLIC_ORIGIN = "https://docviewkit.com";
const LANGUAGE_COOKIE = "docviewkit-language";
const STATIC_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".svg": "image/svg+xml",
  ".txt": "text/plain; charset=utf-8",
  ".ico": "image/x-icon",
  ".ttf": "font/ttf",
  ".wasm": "application/wasm"
};
const PUBLIC_SDK_ASSETS = new Set([
  "index.js",
  "accuracy.js",
  "LICENSE",
  "NOTICE",
  "THIRD_PARTY_NOTICES.md",
  "viewer.js",
  "worker.js",
  "image-codec-worker.js",
  "extended-formats.js",
  "types.js",
  "office-viewer-core.wasm",
  "office-viewer-odf.wasm",
  "office-viewer-legacy-office.wasm",
  "office-viewer-pdf.wasm",
  "office-viewer-xps.wasm",
  "office-viewer-ofd.wasm",
  "lcms.wasm",
  "resvg.wasm",
  "version.json",
  "third-party-licenses/Carlito-1.103-LICENSE.txt",
  "third-party-licenses/Caladea-1.002-LICENSE.txt",
]);

function securityHeaders(contentType, { allowWasm = false, allowInlineStyle = false } = {}) {
  return {
    "Content-Type": contentType,
    "Content-Security-Policy": `default-src 'self'; script-src 'self'${allowWasm ? " 'wasm-unsafe-eval'" : ""}; style-src 'self'${allowInlineStyle ? " 'unsafe-inline'" : ""}; img-src 'self' data: blob:; media-src 'self' blob:; connect-src 'self'; worker-src 'self'; object-src 'none'; base-uri 'self'; form-action 'self'; frame-ancestors 'none'`,
    "Referrer-Policy": "strict-origin-when-cross-origin",
    "X-Content-Type-Options": "nosniff",
    "X-Frame-Options": "DENY",
    "Permissions-Policy": "camera=(), microphone=(), geolocation=(), payment=()"
  };
}

function sendJson(response, status, payload, extraHeaders = {}) {
  response.writeHead(status, {
    ...securityHeaders("application/json; charset=utf-8"),
    "Cache-Control": "no-store",
    ...extraHeaders
  });
  response.end(`${JSON.stringify(payload)}\n`);
}

function sendText(response, status, text, contentType = "text/plain; charset=utf-8") {
  response.writeHead(status, {
    ...securityHeaders(contentType),
    "Cache-Control": "no-store"
  });
  response.end(text);
}

function sendHtml(response, html, { allowWasm = false, allowInlineStyle = false, headers = {} } = {}) {
  response.writeHead(200, {
    ...securityHeaders("text/html; charset=utf-8", { allowWasm, allowInlineStyle }),
    "Cache-Control": "no-cache",
    ...headers,
  });
  response.end(html);
}

function redirect(response, location, status = 308) {
  response.writeHead(status, {
    ...securityHeaders("text/plain; charset=utf-8"),
    "Cache-Control": "no-store",
    Location: location,
  });
  response.end(`Redirecting to ${location}\n`);
}

function localizedPath(locale, path = "/") {
  return `/${locale}${path}`;
}

function languageCookie(locale, { secure = false } = {}) {
  const value = locale === "zh-cn" ? "zh-CN" : "en";
  return `${LANGUAGE_COOKIE}=${value}; Path=/; SameSite=Lax; Max-Age=31536000${secure ? "; Secure" : ""}`;
}

function preferredLocale(request) {
  const saved = String(request.headers.cookie || "")
    .split(";")
    .map((part) => part.trim())
    .find((part) => part.startsWith(`${LANGUAGE_COOKIE}=`))
    ?.slice(LANGUAGE_COOKIE.length + 1);
  if (saved === "en") return "en";
  if (saved === "zh-CN") return "zh-cn";
  const languages = String(request.headers["accept-language"] || "")
    .split(",")
    .map((entry, index) => {
      const [tag, ...parameters] = entry.trim().toLowerCase().split(";");
      const qualityValue = parameters.find((parameter) => parameter.trim().startsWith("q="))?.trim().slice(2);
      const quality = qualityValue === undefined ? 1 : Number(qualityValue);
      return { tag, quality: Number.isFinite(quality) ? quality : 0, index };
    })
    .filter(({ quality }) => quality > 0)
    .sort((left, right) => right.quality - left.quality || left.index - right.index);
  const supported = languages.find(({ tag }) => /^(?:en|zh)(?:-|$)/u.test(tag))?.tag;
  return supported?.startsWith("zh") ? "zh-cn" : "en";
}

function requestedLocale(request, url) {
  const parameter = url.searchParams.get("lang");
  if (parameter === "en") return "en";
  if (parameter === "zh-CN") return "zh-cn";
  return preferredLocale(request);
}

function escapeAttribute(value) {
  return String(value).replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;");
}

function localizeStaticHtml(source, locale, path, version, siteVerification) {
  let html = source;
  if (locale === "en") {
    for (const [original, translation] of [...english].sort(([left], [right]) => right.length - left.length)) {
      html = html.replaceAll(original, translation);
    }
  }
  const language = locale === "en" ? "en" : "zh-CN";
  const canonical = `${PUBLIC_ORIGIN}${localizedPath(locale, path)}`;
  const alternateEn = `${PUBLIC_ORIGIN}${localizedPath("en", path)}`;
  const alternateZh = `${PUBLIC_ORIGIN}${localizedPath("zh-cn", path)}`;
  const title = /<title>([^<]+)<\/title>/u.exec(html)?.[1] || "DocViewKit";
  const description = /<meta name="description" content="([^"]+)">/u.exec(html)?.[1] || "";
  const escapedTitle = escapeAttribute(title);
  const escapedDescription = escapeAttribute(description);
  const isHome = path === "/";
  const structuredData = JSON.stringify({
    "@context": "https://schema.org",
    "@graph": isHome ? [
      {
        "@type": "Organization",
        "@id": `${PUBLIC_ORIGIN}/#organization`,
        name: "DocViewKit",
        url: `${PUBLIC_ORIGIN}/en/`,
        logo: `${PUBLIC_ORIGIN}/favicon.svg`,
        sameAs: ["https://github.com/docviewkit/viewer", "https://www.npmjs.com/package/@docviewkit/viewer"],
      },
      {
        "@type": "WebSite",
        "@id": `${PUBLIC_ORIGIN}/#website`,
        name: "DocViewKit",
        url: `${PUBLIC_ORIGIN}/en/`,
        inLanguage: ["en", "zh-CN"],
        publisher: { "@id": `${PUBLIC_ORIGIN}/#organization` },
      },
      {
        "@type": "SoftwareApplication",
        "@id": `${PUBLIC_ORIGIN}/#viewer`,
        name: "DocViewKit Viewer",
        url: canonical,
        applicationCategory: "DeveloperApplication",
        operatingSystem: "Browsers and compatible cross-platform WebView hosts",
        softwareVersion: version,
        description,
        offers: { "@type": "Offer", price: "0", priceCurrency: "USD" },
        publisher: { "@id": `${PUBLIC_ORIGIN}/#organization` },
      },
      {
        "@type": "FAQPage",
        "@id": `${canonical}#faq`,
        url: `${canonical}#faq`,
        inLanguage: language,
        // ponytail: plain-text FAQs only; use an HTML parser if answers gain rich markup.
        mainEntity: [...html.matchAll(/<details><summary>([^<]+)<\/summary><p>([^<]+)<\/p><\/details>/gu)]
          .map(([, name, text]) => ({
            "@type": "Question", name, acceptedAnswer: { "@type": "Answer", text },
          })),
      },
    ] : [{
      "@type": "WebPage",
      "@id": `${canonical}#page`,
      url: canonical,
      name: title,
      description,
      inLanguage: language,
      isPartOf: { "@id": `${PUBLIC_ORIGIN}/#website` },
    }],
  }).replaceAll("<", "\\u003c");
  html = html
    .replace(/<html lang="[^"]+">/u, `<html lang="${language}">`)
    .replace(/\s*<link rel="canonical"[^>]+>/u, "")
    .replace("</title>", `</title>\n  <link rel="canonical" href="${canonical}">\n  <link rel="alternate" hreflang="en" href="${alternateEn}">\n  <link rel="alternate" hreflang="zh-CN" href="${alternateZh}">\n  <link rel="alternate" hreflang="x-default" href="${alternateEn}">\n  <meta property="og:type" content="website">\n  <meta property="og:site_name" content="DocViewKit">\n  <meta property="og:title" content="${escapedTitle}">\n  <meta property="og:description" content="${escapedDescription}">\n  <meta property="og:url" content="${canonical}">\n  <meta name="twitter:card" content="summary">\n  <script type="application/ld+json">${structuredData}</script>`)
    .replaceAll('href="/"', `href="/${locale}/"`)
    .replaceAll('href="/for-ide/"', `href="/${locale}/for-ide/"`)
    .replaceAll('href="/for-browser/"', `href="/${locale}/for-browser/"`)
    .replaceAll('href="/demo/"', `href="/${locale}/demo/"`);
  if (isHome) {
    const verification = [
      siteVerification.google ? `<meta name="google-site-verification" content="${escapeAttribute(siteVerification.google)}">` : "",
      siteVerification.bing ? `<meta name="msvalidate.01" content="${escapeAttribute(siteVerification.bing)}">` : "",
    ].filter(Boolean).join("\n  ");
    if (verification) html = html.replace("</head>", `  ${verification}\n</head>`);
  }
  return html;
}

function renderRobots() {
  const publicRules = ["Allow: /", "Disallow: /api/"];
  return `${[
    "User-agent: *", ...publicRules,
    "", "User-agent: OAI-SearchBot", ...publicRules,
    "", "User-agent: ChatGPT-User", ...publicRules,
    "", "User-agent: Claude-SearchBot", ...publicRules,
    "", "User-agent: Claude-User", ...publicRules,
    "", `Sitemap: ${PUBLIC_ORIGIN}/sitemap.xml`, "",
  ].join("\n")}`;
}

function renderSitemap() {
  const catalog = getDocsCatalog();
  return `<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n${getIndexablePaths().map((path) => `  <url><loc>${PUBLIC_ORIGIN}${path}</loc><lastmod>${catalog.updatedAt}</lastmod></url>`).join("\n")}\n</urlset>\n`;
}

function isSecureRequest(request) {
  return request.socket.encrypted || request.headers["x-forwarded-proto"] === "https";
}

function requestBaseUrl(request) {
  const protocol = isSecureRequest(request) ? "https" : "http";
  try {
    return new URL(`${protocol}://${request.headers.host || "127.0.0.1"}`).origin;
  } catch {
    const error = new Error("Request URL or Host is invalid.");
    error.status = 400;
    throw error;
  }
}

export function createCommercialApp({
  publicDir,
  sdkDir = resolve(publicDir, "../../dist"),
  immutableSdkAssets = true,
  indexNowKey = "",
  siteVerification = {},
}) {
  const staticRoot = resolve(publicDir);
  const sdkRoot = resolve(sdkDir);
  const sdkManifestPath = resolve(sdkRoot, "version.json");
  const sdkManifest = existsSync(sdkManifestPath)
    ? JSON.parse(readFileSync(sdkManifestPath, "utf8"))
    : null;
  const sdkVersion = typeof sdkManifest?.version === "string" ? sdkManifest.version : null;
  const publicIndexNowKey = /^[A-Za-z0-9_-]{8,128}$/u.test(indexNowKey) ? indexNowKey : "";

  function handleApi(request, response, url) {
    if (request.method === "GET" && url.pathname === "/api/health") {
      return sendJson(response, 200, { ok: true, service: "docviewkit-commercial" });
    }
    if (request.method === "GET" && url.pathname === "/api/docs") {
      return sendJson(response, 200, getDocsCatalog());
    }
    return sendJson(response, 404, { error: "API route not found." });
  }

  async function serveStatic(response, pathname) {
    const sdkAsset = pathname.startsWith("/sdk/");
    const root = sdkAsset ? sdkRoot : staticRoot;
    let relativePath = decodeURIComponent(sdkAsset ? pathname.slice(4) : pathname);
    if (relativePath.endsWith("/")) relativePath += "index.html";
    if (sdkAsset) {
      relativePath = relativePath.replace(/^\/+/u, "");
      const versioned = /^v([^/]+)\/(.+)$/u.exec(relativePath);
      if (versioned !== null) {
        if (sdkVersion === null || versioned[1] !== sdkVersion) return false;
        relativePath = versioned[2];
      }
      if (!PUBLIC_SDK_ASSETS.has(relativePath)
        && !/^runtime-[A-Z0-9]+\.js$/u.test(relativePath)
        && !/^(?:Carlito|Caladea)-(?:Regular|Bold|Italic|BoldItalic)\.ttf$/u.test(relativePath)
        && !/^third-party-licenses\/[^/]+(?:-LICENSE|\.(?:txt|md))$/u.test(relativePath)) return false;
    }
    const candidate = resolve(root, `./${normalize(relativePath).replace(/^[/\\]+/u, "")}`);
    if (!candidate.startsWith(`${root}/`) || !existsSync(candidate)) {
      return false;
    }
    const extension = extname(candidate).toLowerCase();
    const contentType = STATIC_TYPES[extension] || "application/octet-stream";
    const body = pathname === "/sdk/version.json" && sdkManifest !== null
      ? Buffer.from(`${JSON.stringify({
          ...sdkManifest,
          immutableAssets: immutableSdkAssets,
        })}\n`)
      : await readFile(candidate);
    response.writeHead(200, {
      ...securityHeaders(contentType, {
        allowWasm: sdkAsset || pathname.startsWith("/demo/"),
        allowInlineStyle: pathname.startsWith("/demo/"),
      }),
      "Cache-Control": pathname === "/sdk/version.json"
        ? "no-store"
        : immutableSdkAssets && /^\/sdk\/v[^/]+\//u.test(pathname)
          ? "public, max-age=31536000, immutable"
          : extension === ".html" || sdkAsset
            ? "no-cache"
            : "public, max-age=3600"
    });
    response.end(body);
    return true;
  }

  async function serveLocalizedStatic(request, response, locale, pathname) {
    const files = new Map([
      ["/", "index.html"],
      ["/for-ide/", "for-ide/index.html"],
      ["/for-browser/", "for-browser/index.html"],
      ["/demo/", "demo/index.html"],
      ["/privacy/docviewkit-omni/", "privacy/docviewkit-omni/index.html"],
      ["/license/docviewkit-omni/", "license/docviewkit-omni/index.html"],
    ]);
    const relativePath = files.get(pathname);
    if (!relativePath) return false;
    const html = await readFile(resolve(staticRoot, relativePath), "utf8");
    sendHtml(response, localizeStaticHtml(html, locale, pathname, getDocsCatalog().version, siteVerification), {
      allowWasm: pathname === "/demo/",
      allowInlineStyle: pathname === "/demo/",
      headers: {
        "Content-Language": locale === "en" ? "en" : "zh-CN",
        "Set-Cookie": languageCookie(locale, { secure: isSecureRequest(request) }),
      },
    });
    return true;
  }

  const server = createServer(async (request, response) => {
    try {
      const url = new URL(request.url || "/", requestBaseUrl(request));
      const forwardedHost = String(request.headers["x-forwarded-host"] || request.headers.host || "").split(",", 1)[0].trim().split(":", 1)[0];
      if (forwardedHost === "www.docviewkit.com") {
        return redirect(response, `${PUBLIC_ORIGIN}${url.pathname}${url.search}`);
      }
      if (url.pathname === "/robots.txt") {
        return sendText(response, 200, renderRobots());
      }
      if (url.pathname === "/sitemap.xml") {
        return sendText(response, 200, renderSitemap(), "application/xml; charset=utf-8");
      }
      if (url.pathname === "/docs.json") {
        return sendJson(response, 200, getDocsCatalog());
      }
      if (publicIndexNowKey && url.pathname === `/${publicIndexNowKey}.txt`) {
        return sendText(response, 200, `${publicIndexNowKey}\n`);
      }
      if (url.pathname === "/llms.txt") {
        return sendText(response, 200, renderLlmsIndex(requestBaseUrl(request)));
      }
      if (url.pathname === "/llm-full.txt" || url.pathname === "/llms-full.txt") {
        return sendText(response, 200, renderLlmsFull(requestBaseUrl(request)));
      }
      if (url.pathname.startsWith("/api/")) {
        return handleApi(request, response, url);
      }
      if (request.method !== "GET" && request.method !== "HEAD") {
        return sendJson(response, 405, { error: "Method not allowed." }, { Allow: "GET, HEAD" });
      }
      if (url.pathname === "/") {
        return redirect(response, `/${requestedLocale(request, url)}/`);
      }
      if (["/for-ide/", "/for-browser/", "/demo/", "/privacy/docviewkit-omni/", "/license/docviewkit-omni/"].includes(url.pathname)) {
        return redirect(response, `/${requestedLocale(request, url)}${url.pathname}`);
      }
      const staticIndexRedirects = new Map([
        ["/index.html", "/"],
        ["/for-ide/index.html", "/for-ide/"],
        ["/for-browser/index.html", "/for-browser/"],
        ["/demo/index.html", "/demo/"],
        ["/privacy/docviewkit-omni/index.html", "/privacy/docviewkit-omni/"],
        ["/license/docviewkit-omni/index.html", "/license/docviewkit-omni/"],
      ]);
      if (staticIndexRedirects.has(url.pathname)) {
        return redirect(response, `/${requestedLocale(request, url)}${staticIndexRedirects.get(url.pathname)}`);
      }
      if (url.pathname === "/docs/") {
        const requested = url.searchParams.get("doc") || "quickstart";
        const slug = requested === "support" ? "supported-formats" : requested;
        return redirect(response, `/docs/${encodeURIComponent(slug)}/`);
      }
      const missingSlashDocs = /^\/docs\/([a-z0-9-]+)$/u.exec(url.pathname);
      if (missingSlashDocs) return redirect(response, `${url.pathname}/`);
      if (url.pathname === "/docs/support/") return redirect(response, "/docs/supported-formats/");
      const docsMatch = /^\/docs\/([a-z0-9-]+)\/$/u.exec(url.pathname);
      if (docsMatch) {
        const html = renderDocsPage(PUBLIC_ORIGIN, docsMatch[1]);
        if (html) return sendHtml(response, html, { headers: { "Content-Language": "en" } });
      }
      const directDoc = getDocsCatalog().docs.find((doc) => doc.path === url.pathname);
      if (directDoc) {
        return sendHtml(response, renderDocsPage(PUBLIC_ORIGIN, directDoc.slug), { headers: { "Content-Language": "en" } });
      }
      const missingSlashDirectDoc = getDocsCatalog().docs.find((doc) => doc.path === `${url.pathname}/`);
      if (missingSlashDirectDoc) return redirect(response, missingSlashDirectDoc.path);
      const localizedMatch = /^\/(en|zh-cn)(\/|\/for-ide\/|\/for-browser\/|\/demo\/|\/privacy\/docviewkit-omni\/|\/license\/docviewkit-omni\/)$/u.exec(url.pathname);
      if (localizedMatch && await serveLocalizedStatic(request, response, localizedMatch[1], localizedMatch[2])) return;
      const missingSlashLocalized = /^\/(en|zh-cn)(|\/for-ide|\/for-browser|\/demo|\/privacy\/docviewkit-omni|\/license\/docviewkit-omni)$/u.exec(url.pathname);
      if (missingSlashLocalized) return redirect(response, `${url.pathname}/`);
      if (await serveStatic(response, url.pathname)) return;
      return sendText(response, 404, "Page not found.\n");
    } catch (error) {
      const status = Number(error.status) || 500;
      const message = status >= 500 ? "The service could not complete this request." : error.message;
      if (status >= 500) console.error(error);
      return sendJson(response, status, { error: message, ...(error.field ? { field: error.field } : {}), ...(error.code ? { code: error.code } : {}) });
    }
  });

  return {
    server,
    close() {
      return new Promise((resolveClose, rejectClose) => {
        server.close((error) => {
          if (error) rejectClose(error);
          else resolveClose();
        });
      });
    }
  };
}
