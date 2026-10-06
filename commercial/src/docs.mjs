import { readFileSync } from "node:fs";

const docsUrl = new URL("../content/docs.json", import.meta.url);
const packageUrl = new URL("../../package.json", import.meta.url);
const catalog = {
  ...JSON.parse(readFileSync(docsUrl, "utf8")),
  version: JSON.parse(readFileSync(packageUrl, "utf8")).version,
};

export function getDocsCatalog() {
  return catalog;
}

function docPath(doc) {
  return doc.path || `/docs/${encodeURIComponent(doc.slug)}/`;
}

export function getIndexablePaths() {
  return [
    "/en/",
    "/zh-cn/",
    "/en/for-ide/",
    "/zh-cn/for-ide/",
    "/en/for-browser/",
    "/zh-cn/for-browser/",
    "/en/demo/",
    "/zh-cn/demo/",
    "/en/privacy/docviewkit-omni/",
    "/zh-cn/privacy/docviewkit-omni/",
    "/en/license/docviewkit-omni/",
    "/zh-cn/license/docviewkit-omni/",
    ...catalog.docs.map(docPath),
  ];
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function blockToHtml(block) {
  switch (block.type) {
    case "p":
      return `<p>${escapeHtml(block.text)}</p>`;
    case "callout":
      return `<div class="docs-callout${block.tone === "warning" ? " warning" : ""}"><strong>${escapeHtml(block.title)}</strong><p>${escapeHtml(block.text)}</p></div>`;
    case "list":
      return `<ul>${block.items.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`;
    case "code":
      return `<div class="docs-code"><div class="docs-code-header"><span>${escapeHtml(block.language || "text")}</span><button class="copy-button" type="button" data-docs-copy>Copy</button></div><pre><code>${escapeHtml(block.content)}</code></pre></div>`;
    case "table":
      return `<div class="docs-table-wrap"><table class="docs-table"><thead><tr>${block.headers.map((heading) => `<th>${escapeHtml(heading)}</th>`).join("")}</tr></thead><tbody>${block.rows.map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`).join("")}</tbody></table></div>`;
    case "links":
      return `<ul class="docs-link-list">${block.items.map((item) => `<li><a href="${escapeHtml(item.href)}">${escapeHtml(item.label)}</a></li>`).join("")}</ul>`;
    default:
      return "";
  }
}

function schemaJson(value) {
  return JSON.stringify(value).replaceAll("<", "\\u003c");
}

export function renderDocsPage(baseUrl, slug) {
  const root = baseUrl.replace(/\/$/u, "");
  const doc = catalog.docs.find((item) => item.slug === slug);
  if (!doc) return null;
  const canonical = `${root}${docPath(doc)}`;
  const pageTitle = doc.pageTitle || `${doc.title} — DocViewKit Docs`;
  const navigation = catalog.groups.map((group) => {
    const links = catalog.docs
      .filter((item) => item.group === group.id)
      .map((item) => `<a href="${docPath(item)}"${item.slug === doc.slug ? ' aria-current="page"' : ""}>${escapeHtml(item.title)}</a>`)
      .join("");
    return `<div class="docs-group"><div class="docs-group-title">${escapeHtml(group.label)}</div><nav class="docs-nav" aria-label="${escapeHtml(group.label)}">${links}</nav></div>`;
  }).join("");
  const sections = doc.sections.map((section) => `<section id="${escapeHtml(section.id)}"><h2>${escapeHtml(section.title)}</h2>${section.blocks.map(blockToHtml).join("")}</section>`).join("");
  const outline = doc.sections.map((section) => `<a href="#${escapeHtml(section.id)}">${escapeHtml(section.title)}</a>`).join("");
  const structuredData = schemaJson({
    "@context": "https://schema.org",
    "@graph": [
      {
        "@type": "TechArticle",
        "@id": `${canonical}#article`,
        headline: doc.title,
        description: doc.summary,
        dateModified: catalog.updatedAt,
        inLanguage: "en",
        mainEntityOfPage: canonical,
        author: { "@type": "Organization", "@id": `${root}/#organization`, name: "DocViewKit", url: `${root}/` },
        publisher: { "@id": `${root}/#organization` },
      },
      {
        "@type": "BreadcrumbList",
        itemListElement: [
          { "@type": "ListItem", position: 1, name: "DocViewKit", item: `${root}/en/` },
          { "@type": "ListItem", position: 2, name: "Documentation", item: `${root}/docs/quickstart/` },
          { "@type": "ListItem", position: 3, name: doc.title, item: canonical },
        ],
      },
    ],
  });
  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="description" content="${escapeHtml(doc.summary)}">
  <meta property="og:type" content="article">
  <meta property="og:site_name" content="DocViewKit">
  <meta property="og:title" content="${escapeHtml(pageTitle)}">
  <meta property="og:description" content="${escapeHtml(doc.summary)}">
  <meta property="og:url" content="${canonical}">
  <meta name="twitter:card" content="summary">
  <title>${escapeHtml(pageTitle)}</title>
  <link rel="canonical" href="${canonical}">
  <link rel="icon" href="/favicon.svg" type="image/svg+xml">
  <link rel="stylesheet" href="/assets/styles.css?v=20261002-3">
  <script type="application/ld+json">${structuredData}</script>
  <script type="module" src="/assets/site.js?v=20261002-3"></script>
  <script type="module" src="/assets/docs.js?v=20260817-1"></script>
</head>
<body class="docs-body">
  <a class="skip-link" href="#docs-article">Skip to documentation</a>
  <header class="site-header"><nav class="site-nav container" aria-label="Documentation navigation"><div class="nav-actions"><button class="icon-button docs-mobile-toggle" type="button" data-docs-menu aria-controls="docs-sidebar" aria-expanded="false" aria-label="Open documentation menu"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16M4 12h16M4 18h16"/></svg></button><a class="brand" href="/en/">DocViewKit <span class="small">/ Docs</span></a></div><div class="nav-actions"><a class="nav-link" href="/en/demo/">Demo</a><a class="nav-link" href="/llms.txt">llms.txt</a><a class="button button-primary" href="mailto:novalag778@gmail.com?subject=DocViewKit%20Services">Contact services</a></div></nav></header>
  <div class="docs-shell">
    <aside class="docs-sidebar" id="docs-sidebar" aria-label="Documentation table of contents"><div class="docs-search"><svg class="inline-icon" viewBox="0 0 24 24" aria-hidden="true"><circle cx="11" cy="11" r="6"/><path d="m16 16 4 4"/></svg><label class="sr-only" for="docs-search-input">Search documentation</label><input id="docs-search-input" type="search" placeholder="Search documentation" autocomplete="off"></div>${navigation}<div class="ai-doc-card"><strong>For AI coding tools</strong><p class="small">HTML and plain-text documentation share one source.</p><a href="/llm-full.txt">Read llm-full.txt</a></div></aside>
    <main class="docs-main"><article class="docs-article" id="docs-article"><header class="docs-article-header"><h1>${escapeHtml(doc.title)}</h1><p class="lede">${escapeHtml(doc.summary)}</p></header>${sections}</article></main>
    <aside class="docs-outline" aria-label="On this page"><div class="outline-title">On this page</div><nav class="outline-list">${outline}</nav></aside>
  </div>
</body>
</html>`;
}

function blockToText(block) {
  switch (block.type) {
    case "p":
      return block.text;
    case "callout":
      return `${block.title}: ${block.text}`;
    case "list":
      return block.items.map((item) => `- ${item}`).join("\n");
    case "code":
      return `\`\`\`${block.language || "text"}\n${block.content}\n\`\`\``;
    case "table": {
      const header = `| ${block.headers.join(" | ")} |`;
      const separator = `| ${block.headers.map(() => "---").join(" | ")} |`;
      const rows = block.rows.map((row) => `| ${row.join(" | ")} |`).join("\n");
      return `${header}\n${separator}\n${rows}`;
    }
    case "links":
      return block.items.map((item) => `- [${item.label}](${item.href})`).join("\n");
    default:
      return "";
  }
}

export function renderLlmsIndex(baseUrl) {
  const root = baseUrl.replace(/\/$/, "");
  const lines = [
    `# ${catalog.product}`,
    "",
    `> ${catalog.description}`,
    "",
    `Version: ${catalog.version}`,
    `Updated: ${catalog.updatedAt}`,
    "",
    "## Product boundary",
    "",
    "DocViewKit is a frontend-local, read-only document viewer SDK for Office, PDF, OFD and other supported formats. It does not provide editing, collaboration, OCR, RAG, embeddings, model calls, or AI answers.",
    "",
    "## Documentation",
    ""
  ];

  for (const doc of catalog.docs) {
    lines.push(`- [${doc.title}](${root}${docPath(doc)}): ${doc.summary}`);
  }

  lines.push(
    "",
    "## Machine-readable resources",
    "",
    `- [Complete documentation](${root}/llms-full.txt)`,
    `- [Structured documentation JSON](${root}/docs.json)`,
    "- [Support, customization and enterprise delivery](mailto:novalag778@gmail.com)",
    "",
    "Pin the package version, runtime assets, and documented support boundary together."
  );

  return `${lines.join("\n")}\n`;
}

export function renderLlmsFull(baseUrl) {
  const lines = [renderLlmsIndex(baseUrl).trim(), "", "# Complete documentation", ""];

  for (const doc of catalog.docs) {
    lines.push(`## ${doc.title}`, "", doc.summary, "");
    for (const section of doc.sections) {
      lines.push(`### ${section.title}`, "");
      for (const block of section.blocks) {
        lines.push(blockToText(block), "");
      }
    }
  }

  return `${lines.join("\n").trim()}\n`;
}
