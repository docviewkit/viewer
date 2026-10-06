# DocViewKit GEO / SEO 官方一手资料基线

> 调研日期：2026-08-17。本文只采用 Google Search Central、Bing Webmaster、Schema.org、OpenAI、Anthropic 的官方文档；不采用第三方“GEO 排名因子”、营销案例或搜索摘要。本文回答“官方资料能证明什么”，不把建议表述成收录、排名或引用保证。

## 1. 结论先行

1. **对 Google，GEO 不是一套独立技术栈。** Google 明确说明 AI Overviews / AI Mode 建立在核心搜索索引、排名与质量系统之上；页面要成为 AI 回答的支持链接，首先必须已被索引并有资格展示 snippet。Google 也明确表示不需要 AI 专用标记或 `llms.txt`，没有专门的 schema.org 类型。[Google 生成式 AI 优化指南](https://developers.google.com/search/docs/fundamentals/ai-optimization-guide)
2. **对 Bing / Copilot，“可引用性”已有可测量的官方口径。** Bing Webmaster Tools 的 AI Performance 报告提供 Total Citations、Average Cited Pages、Grounding Queries 和逐 URL 引用活动；这比第三方 GEO 分数更适合作为验收基线。[Bing AI Performance](https://blogs.bing.com/webmaster/February-2026/Introducing-AI-Performance-in-Bing-Webmaster-Tools-Public-Preview)
3. **最短有效路径是：抓取与索引打底，稳定 URL 与实体信息消歧，再发布能独立回答真实问题的一手内容。** 不应先批量制造关键词页、FAQ 页或“AI 优化文本”。Google 明确反对为覆盖查询变体而规模化造页；Bing 官方建议用清晰标题、表格、FAQ、实例、数据和来源提高 AI 引用的准确性。[Google 生成式 AI 优化指南](https://developers.google.com/search/docs/fundamentals/ai-optimization-guide)、[Bing AI Performance](https://blogs.bing.com/webmaster/February-2026/Introducing-AI-Performance-in-Bing-Webmaster-Tools-Public-Preview)
4. **结构化数据用于准确描述可见实体，不是 GEO 捷径。** Google 推荐 JSON-LD，但要求标记与页面可见主内容一致，并明确不保证富结果或 AI 引用。[Google 结构化数据指南](https://developers.google.com/search/docs/appearance/structured-data/sd-policies)

## 2. 官方事实与 DocViewKit 的直接含义

| 主题 | 官方可确认事实 | 对 DocViewKit 的含义 |
|---|---|---|
| Google AI 可见性 | AI 功能从 Search index 检索页面；候选页面必须已索引且可展示 snippet；不存在额外准入标记 | 先修普通 SEO 的抓取、索引、正文、内部链接和页面体验，再谈 GEO |
| Google query fan-out | AI 功能会并发检索原问题的相关子问题 | 用少量深入页面覆盖一个完整购买/集成任务；不要为同义词机械拆出薄页 |
| Bing / Copilot 引用 | Bing 可报告引用量、被引 URL 和 grounding query | 以真实被引 URL 和 query 建立内容迭代闭环，不使用不可核验的“AI visibility score” |
| 文本可摘录性 | Google 要求重要内容以文本形式可用；Bing建议清晰标题、表格、FAQ、证据、数据和来源 | 核心能力、格式矩阵、隐私边界、性能口径、限制不得只存在于视频、Canvas 或截图中 |
| robots 控制 | Google Search AI 由 Googlebot 访问控制；OpenAI / Anthropic 为搜索、训练、用户请求设置不同 user-agent | 不要把“拒绝训练”和“拒绝搜索引用”混为一谈；逐类设置并测试 |
| sitemap 与更新 | Google、Bing 均使用 sitemap 发现 URL；Bing强调准确 `lastmod`，并建议 IndexNow 通知新增、更新、删除 | sitemap 只列 canonical、可索引 URL；内容发布流水线写真实修改时间并触发 IndexNow |
| canonical | Google把重定向和 `rel="canonical"` 视为强信号，sitemap 为较弱信号；方法可叠加 | 每个公开页面自引用 canonical；HTTP/HTTPS、尾斜杠、参数和语言版本不能互相冲突 |
| 结构化数据 | Google 推荐 JSON-LD；标记必须代表可见主内容；无专用 AI Schema | 只标真实存在的 WebSite、Organization、SoftwareApplication、文章与面包屑实体 |

主要依据：[Google AI features](https://developers.google.com/search/docs/appearance/ai-features)、[Google 生成式 AI 优化指南](https://developers.google.com/search/docs/fundamentals/ai-optimization-guide)、[Google sitemap](https://developers.google.com/search/docs/crawling-indexing/sitemaps/build-sitemap)、[Google canonical](https://developers.google.com/search/docs/crawling-indexing/consolidate-duplicate-urls)、[Bing sitemap 与 AI 搜索](https://blogs.bing.com/webmaster/July-2025/Keeping-Content-Discoverable-with-Sitemaps-in-AI-Powered-Search)、[Bing AI Performance](https://blogs.bing.com/webmaster/February-2026/Introducing-AI-Performance-in-Bing-Webmaster-Tools-Public-Preview)。

## 2A. DocViewKit 线上实测审计

2026-08-17 直接请求 `https://docviewkit.com/` 及其公开入口，并对照当前公开 npm 元数据，得到以下结果。搜索结果抽样不是 Search Console 的替代品；是否已收录应最终以 Google Search Console 与 Bing Webmaster Tools 为准。

| 优先级 | 实测结果 | 影响与最小修复 |
|---|---|---|
| P0 | `/robots.txt` 与 `/sitemap.xml` 均返回 404 | 缺少统一抓取策略与 URL 发现入口。先增加两个根路径，sitemap 只列 200、canonical、可索引页面，并提交 Google / Bing |
| P0 | 文档正文由浏览器请求 `/api/docs` 后注入；原始 `/docs/` HTML 只有 skeleton。所有 `?doc=...` 变体又 canonical 到 `/docs/` | 搜索引擎需二次渲染才能读正文，且各主题不能成为独立引用目标。把文档输出为服务端或构建时生成的静态 HTML，并使用 `/docs/quickstart/`、`/docs/viewer-api/` 等独立 self-canonical URL |
| P0 | 线上 `llms.txt`、`llm-full.txt` 与 `/api/docs` 自报 v0.2.43，npm `@docviewkit/viewer` 最新为 v0.2.45 | Agent 可能引用过期产品事实。发布流水线应从同一版本源生成站点、docs、llms、npm 元数据，并在发布门禁中做版本一致性断言 |
| P0 | Quickstart 先安装 `@docviewkit/viewer`，随后示例却导入 `@docviewkit/sdk/viewer`；公开包 README 使用的是 `import "@docviewkit/viewer"` | 当前最关键的可引用代码示例不自洽。修为公开包入口，并用可执行 quickstart 测试防回归 |
| P1 | `https://www.docviewkit.com/` 返回 200，而不是 301 到 apex；页面 canonical 虽指向 apex，但主机仍重复可访问 | 将 `www` 永久重定向到 `https://docviewkit.com/`，并让内部链接、sitemap、JSON-LD 和 `og:url` 全部使用 apex |
| P1 | 首页原始 HTML 是 `lang="zh-CN"` 与中文 title/H1；英文由同一 URL 的客户端语言切换提供，且没有 `hreflang` | 英文查询和中文查询共用一个 canonical，搜索引擎无法稳定选择语言版本。使用 `/en/` 与 `/zh-cn/` 独立 HTML、自引用 canonical、互相 `hreflang` 和 `x-default` |
| P1 | 首页已有 title、description、self-canonical、H1/H2、正文和普通 `<a>` 链接，这是良好基础；但首页、Docs、Demo、Portal 均无 JSON-LD、Open Graph、Twitter Card | 保留现有语义 HTML；先加最小 `WebSite`、`Organization`、`SoftwareApplication` 和文档页 `TechArticle` / `BreadcrumbList`。社交元数据用于稳定分享摘要，不把它当排名捷径 |
| P1 | 当前公开内容主要是首页、Demo、单壳 Docs、Portal；缺少可独立索引的支持格式、浏览器兼容、架构/隐私、精度证据、性能方法和版本化 release notes 页面 | 竞争查询结果已经由拥有专题落地页、格式清单、框架示例和一手证据的产品占据。优先发布 6 个事实页，不批量生成同义词薄页 |
| P2 | `llms.txt` 与 `llm-full.txt` 已存在且首页 footer 可发现 | 可保留为非 Google 的附加机器入口，但必须由同一内容源生成、链接全为 canonical 200 URL；它不能替代 HTML、sitemap 或索引 |

当前最大问题不是文案长度，而是**公开事实没有稳定、独立、版本一致的 HTML URL**。因此实施顺序必须是：技术可发现性与版本一致性 → 独立事实页 → 一手证据 → 外部实体与引用扩散；Schema 和 `llms.txt` 只能辅助，不能倒序。

## 3. 建议的信息架构与内容优先级

### P0：让产品事实拥有稳定、可索引、可引用的 URL

至少让以下用户任务各有一个明确 canonical 页面；已有页面能承载就复用，不为关键词另造近重复页面：

| 页面 / 主题 | 必须可直接摘录的事实 |
|---|---|
| 产品首页 | DocViewKit 是什么、目标用户、浏览器端还是服务端、支持的核心任务、主 CTA |
| Supported formats | 格式、扩展名、查看/解析能力、已知限制、版本或更新时间；不要只有 logo |
| Architecture / privacy | 文件是否上传、运行位置、Worker / Wasm 边界、网络请求、数据保留边界 |
| Security | 不执行的活动内容、资源预算、隔离边界、明确“不保证”的能力 |
| Accuracy / fidelity | 测试语料、验收口径、结果日期、失败与降级如何披露 |
| SDK docs / quickstart | 最小安装和渲染代码、浏览器支持、包入口、错误处理、版本号 |
| Pricing / licensing | 真实可公开的许可范围、评估方式、销售入口；若不公开价格就明确“联系销售”，不要伪造 Offer |
| Comparisons / alternatives | 按部署、格式、隐私、对象 API、编辑/只读等可验证维度比较，并链接双方一手资料 |

每页建议使用：一个准确 `<title>`、一个与标题一致的可见 `h1`、开头 1–2 句直接定义、清晰 `h2/h3`、HTML 表格、真实示例、限制说明、作者/维护方、首次发布日期与**真实**更新时间、稳定锚点。Google 说明 title link 会参考 `<title>`、主视觉标题、heading、`og:title`、锚文本和 `WebSite` 数据；snippet 主要来自页面正文，也可能采用 meta description。[Google title links](https://developers.google.com/search/docs/appearance/title-link)、[Google snippets](https://developers.google.com/search/docs/appearance/snippet)

### P1：发布无法由通用 AI 轻易重写的一手证据

优先级应是：

1. 可复现的格式兼容性与像素/对象级精度报告；
2. 真实文件的解析失败、最佳努力降级和修复复盘；
3. 浏览器、文件大小、页数、内存与加载阶段都写清楚的性能测试；
4. 具体集成场景的端到端代码与边界，如 AI citation 回到页、单元格或对象；
5. 有证据链接的产品比较与选择指南。

Google 明确把独特观点、第一手经验、非同质化、对用户有用的内容列为长期更可能影响生成式搜索表现的因素，并反对重述现有网页或批量覆盖查询变体。[Google 生成式 AI 优化指南](https://developers.google.com/search/docs/fundamentals/ai-optimization-guide)、[Google people-first content](https://developers.google.com/search/docs/fundamentals/creating-helpful-content)

### P2：上线后的提交、加速与观测顺序

#### 上线当天：先确认可抓取，再提交

1. **验证站点所有权。** 在 Google Search Console 与 Bing Webmaster Tools 中添加并验证 `https://docviewkit.com/`；提交前先确认首页和 P0 页面返回 200、无 `noindex`、canonical 一致、正文无需登录即可读取。
2. **向 Google 提交 sitemap。** 在 Search Console 的 Sitemaps 报告提交根 sitemap URL。Google说明“提交”只是告知 sitemap 的位置，不是把文件上传给 Google；报告可查看处理历史与解析错误。[Google Sitemaps report](https://support.google.com/webmasters/answer/7451001)
3. **用 Google URL Inspection 检查少量关键页。** 对首页、Supported formats、Architecture / privacy、Security、Accuracy、Quickstart 等少量 canonical URL 运行 live test；核对 fetch、rendered HTML、indexability、结构化数据和 indexed view 中的 Google-selected canonical，通过后点击 **Request indexing**。URL Inspection 适合少量 URL；大量新页或更新页应使用 sitemap，并写准确 `<lastmod>`。[Google URL Inspection](https://support.google.com/webmasters/answer/9012289)、[Ask Google to recrawl](https://developers.google.com/search/docs/crawling-indexing/ask-google-to-recrawl)
4. **向 Bing 提交 sitemap。** 在 Bing Webmaster Tools 的 Sitemaps 工具提交同一 canonical sitemap，并检查处理状态、发现 URL 数及错误。Bing也会发现 robots.txt 中声明的 sitemap。[Bing Sitemaps](https://www.bing.com/webmasters/help/sitemaps-3b5cf6ed)
5. **对 Bing 启用 IndexNow。** 发布流水线在 URL 新增、实质更新、重定向或删除时提交 IndexNow；它通知 Bing及其他参与搜索引擎。Bing将 IndexNow列为强烈推荐的自动化方式，提交状态可在 Webmaster Tools 的 IndexNow 页面观察。[Bing URL Submission](https://www.bing.com/webmasters/help/URL-Submission-62f2860b)、[Bing IndexNow](https://www.bing.com/webmasters/help/indexnow-0z209wby)
6. **只对少量首发重点页使用 Bing URL Submission。** 在 Submit URLs 中提交首页与 P0 页面即可；不要把人工提交当作发布流水线。Bing明确说明重复人工提交不会加速收录，大批量或持续更新应使用 IndexNow。[Bing URL Submission](https://www.bing.com/webmasters/help/URL-Submission-62f2860b)
7. **验证 AI 搜索爬虫能真实到达。** robots.txt 放行 `OAI-SearchBot`、`Claude-SearchBot` 及需要的用户触发爬虫，同时在 CDN / WAF 放行官方 IP；从访问日志核对 user-agent 与来源 IP，而不能只看到 robots 规则就判定成功。具体爬虫与 IP 链接见 4.1 节。

#### 上线后 1–14 天：按漏斗排障，不重复催收录

| 平台 | 先看什么 | 发现问题时怎么做 |
|---|---|---|
| Google Search Console | sitemap 成功/错误、Page indexing、URL Inspection 的 last crawl / indexed HTML / Google-selected canonical、目标 URL 是否已索引 | 先修 fetch、robots、`noindex`、canonical、空正文或渲染问题；仅对修复后的少量重点页重新 Request indexing |
| Bing Webmaster Tools | sitemap 处理状态、URL Inspection 的 crawl/index 状态、IndexNow 提交与首次索引状态 | 修复抓取或质量问题；新/改/删 URL 重新触发一次 IndexNow，不循环人工提交 |
| OpenAI | 服务端日志中的 `OAI-SearchBot` + 官方 IP；分析工具中的 `utm_source=chatgpt.com` referral | 检查 robots、CDN/WAF 与正文 snippet 控制；不要把 `GPTBot` 训练授权误当成搜索授权 |
| Anthropic | 服务端日志中的 `Claude-SearchBot` / `Claude-User` + 官方 IP；人工高价值查询的来源链接 | 检查每个子域的 robots 和 CDN/WAF；不要把 `ClaudeBot` 训练授权误当成搜索授权 |

Google说明抓取可能需要数天到数周；Request indexing 有配额，重复请求同一 URL 不会更快，也不保证进入索引。[Ask Google to recrawl](https://developers.google.com/search/docs/crawling-indexing/ask-google-to-recrawl)、[Google crawling/indexing FAQ](https://developers.google.com/search/help/crawling-index-faq)

#### 有数据后：建立搜索与 AI 引用闭环

- **Google：** 观察目标 query、page、impressions、clicks 和索引覆盖；账户出现 Generative AI performance report 时，再看 AI features 的 impressions、pages、countries、devices 与时间趋势。该报告处于分阶段提供状态，未出现时不能据此断言“Google AI 没有引用”。[Google Generative AI performance report announcement](https://developers.google.com/search/blog/2026/06/gen-ai-performance-reports)
- **Bing / Copilot：** 在 AI Performance 观察 Total Citations、Average Cited Pages、Grounding Queries、逐 URL citation activity 和趋势。这些数据代表引用活动，不代表页面排名、权威或在单次答案中的位置。[Bing AI Performance](https://blogs.bing.com/webmaster/February-2026/Introducing-AI-Performance-in-Bing-Webmaster-Tools-Public-Preview)
- **站内转化：** 分开记录 Google、Bing/Copilot、ChatGPT、Claude 的 referral，以及进入 docs、demo、试用和销售页面的后续行为；搜索展现、AI 引用和业务转化是三个不同指标。
- **人工复核：** 每月用一组固定的高价值问题检查 Google、Bing/Copilot、ChatGPT、Claude：是否引用正确 canonical、是否摘录到最新事实、是否把限制一起引用；不要只记录“提到了品牌”。

Bing说明 sitemap负责完整 URL 覆盖，IndexNow负责实时 URL 变化，两者互补；AI Performance 才是其官方 AI 引用观测入口。[Bing sitemap 与 AI 搜索](https://blogs.bing.com/webmaster/July-2025/Keeping-Content-Discoverable-with-Sitemaps-in-AI-Powered-Search)、[Bing AI Performance](https://blogs.bing.com/webmaster/February-2026/Introducing-AI-Performance-in-Bing-Webmaster-Tools-Public-Preview)

## 4. 抓取、robots、sitemap 与 canonical

### 4.1 robots 原则

- 公开营销页、文档、证据报告与格式页应允许 `Googlebot`、`Bingbot` 及目标 AI 搜索爬虫访问；CDN、WAF 和登录墙也必须允许匿名读取正文。
- `robots.txt` 用于控制抓取，不用于 canonical，也不应被当作可靠的 noindex 机制；要禁止索引，使用 `noindex` 或鉴权。Google明确指出，被 robots 阻止的 URL 仍可能在没有正文的情况下被索引。[Google technical SEO](https://developers.google.com/search/docs/fundamentals/get-started)、[Google canonical](https://developers.google.com/search/docs/crawling-indexing/consolidate-duplicate-urls)
- 不要对希望被引用的正文设置 `nosnippet`、过小的 `max-snippet` 或 `data-nosnippet`。Google要求 AI 支持链接具备 snippet 资格；Bing也说明 `data-nosnippet` 标记的内容不会出现在搜索 snippet 或 AI 摘要中。[Google robots meta](https://developers.google.com/search/docs/crawling-indexing/robots-meta-tag)、[Bing data-nosnippet](https://blogs.bing.com/webmaster/October-2025/Bing-Introduces-Support-for-the-data-nosnippet-HTML-Attribute)

一个足够简单的基础文件可以是：

```txt
User-agent: *
Disallow: /private/
Disallow: /internal-api/

Sitemap: https://docviewkit.com/sitemap.xml
```

要保留 ChatGPT / Claude 搜索发现和用户请求访问，同时把“是否用于训练”作为独立选择，可明确写成：

```txt
User-agent: OAI-SearchBot
Allow: /
Disallow: /private/
Disallow: /internal-api/

User-agent: ChatGPT-User
Allow: /
Disallow: /private/
Disallow: /internal-api/

User-agent: Claude-SearchBot
Allow: /
Disallow: /private/
Disallow: /internal-api/

User-agent: Claude-User
Allow: /
Disallow: /private/
Disallow: /internal-api/
```

- OpenAI 将 `OAI-SearchBot` 定义为 ChatGPT Search 的搜索爬虫，将 `GPTBot` 定义为可能用于训练生成式基础模型的爬虫，将 `ChatGPT-User` 定义为用户触发、非自动全网抓取的访问。三者可独立控制；若退出训练，可另加 `User-agent: GPTBot` + `Disallow: /`，而不阻止 `OAI-SearchBot`。OpenAI同时提醒，因 `ChatGPT-User` 访问由用户发起，robots.txt 规则可能不适用，敏感内容仍必须依靠鉴权。[OpenAI crawlers](https://developers.openai.com/api/docs/bots)
- Anthropic 将 `Claude-SearchBot`、`Claude-User`、`ClaudeBot` 分别用于搜索索引、用户触发检索和模型开发/潜在训练。若退出训练，可另加 `User-agent: ClaudeBot` + `Disallow: /`，而保留前两者；每个需要控制的子域都要有自己的 robots.txt。[Anthropic crawler controls](https://support.claude.com/en/articles/8896518-does-anthropic-crawl-data-from-the-web-and-how-can-site-owners-block-the-crawler)
- CDN / WAF 还需允许官方公布的来源 IP；仅写 robots 但在网络层返回 403、验证码或 JS challenge，仍无法抓取。OpenAI提供 [OAI-SearchBot IP 列表](https://openai.com/searchbot.json) 和 [ChatGPT-User IP 列表](https://openai.com/chatgpt-user.json)，Anthropic提供 [bot IP 列表](https://claude.com/crawling/bots.json)。生产日志应用 **user-agent token + 当前官方 CIDR** 联合验证；不要只信可伪造的 UA，也不要把官方示例中的版本号写死。
- OpenAI说明 robots.txt 更新后，其 Search 系统约需 24 小时调整；Anthropic没有公布固定生效时间，不能承诺同样时限。[OpenAI crawlers](https://developers.openai.com/api/docs/bots)
- OpenAI说明公开网站可能出现在 ChatGPT Search，但无展示或排名保证；要让正文进入 summary/snippet 并获得清晰引用，不应阻止 `OAI-SearchBot`。ChatGPT 引荐流量会携带 `utm_source=chatgpt.com`。[OpenAI publisher FAQ](https://help.openai.com/en/articles/12627856-publishers-and-developers-faq)、[ChatGPT Search](https://help.openai.com/en/articles/9237897-chatgpt-search)
- Claude Web Search 会提供直接引用及原始来源链接；禁止 `Claude-SearchBot` 会降低搜索可见性，禁止 `Claude-User` 会阻止用户问题触发的页面检索。[Claude Web Search](https://support.claude.com/en/articles/10684626-enable-and-use-web-search)

指定分组必须完整表达该 bot 的规则，避免误以为它会自动继承 `User-agent: *`。Bing特别提醒其指定 `Bingbot` 规则会覆盖通用分组。[Bing robots.txt](https://www.bing.com/webmasters/help/how-to-create-a-robots-txt-file-cb7c31ec)

### 4.2 sitemap

- 使用根目录 UTF-8 XML sitemap，只列完整绝对 URL、200、canonical、允许索引的公开页面。
- `lastmod` 只在正文或重要结构确实变化时更新，不要每天伪造“新鲜度”。
- 在 `robots.txt` 声明 sitemap，并分别提交 Google Search Console 与 Bing Webmaster Tools。
- 发布、实质更新、重定向或删除 URL 时调用 IndexNow；它加速发现，不保证收录或排名。

Google说明 sitemap 会影响 canonical 选择但只是较弱信号，且不保证抓取或收录；Bing说明准确 `lastmod` 有助于决定复抓优先级。[Google sitemap](https://developers.google.com/search/docs/crawling-indexing/sitemaps/build-sitemap)、[Bing sitemap](https://www.bing.com/webmasters/help/sitemaps-3b5cf6ed)、[Bing URL submission](https://www.bing.com/webmasters/help/URL-Submission-62f2860b)

### 4.3 canonical 与多语言

- 每个 canonical HTML 页面放自引用 `<link rel="canonical" href="…">`。
- 统一 `https`、主机名、大小写、尾斜杠和跟踪参数；旧地址使用 301 到唯一目标。
- sitemap、内部链接、`og:url`、结构化数据 `url/@id` 必须指向同一个 canonical。
- 若中英文页面实质翻译，各自 self-canonical，并用互相完整的 `hreflang` 连接；不要把中文 canonical 到英文或反之。

Google把 301 和 `rel="canonical"` 视为强信号、sitemap 视为弱信号，并建议组合一致使用。[Google canonical](https://developers.google.com/search/docs/crawling-indexing/consolidate-duplicate-urls)

## 5. 结构化数据：够用即可

建议 JSON-LD 采用真实实体和稳定 `@id`：

- 首页：[`WebSite`](https://schema.org/WebSite) + [`Organization`](https://schema.org/Organization)，表述官方站名、组织名、logo、官网 URL 与真实 `sameAs`；Google说明首页 `WebSite` 数据可表达站点名称偏好。[Google site names](https://developers.google.com/search/docs/appearance/site-names)
- 产品页：[`SoftwareApplication`](https://schema.org/SoftwareApplication)，只填页面可见且可证实的 `name`、`description`、`applicationCategory`、`operatingSystem`、`softwareVersion`、`featureList`、`offers` 等；没有公开价格就不构造价格。
- 技术文章 / 研究：[`TechArticle`](https://schema.org/TechArticle)，如实填 `headline`、`author`、`datePublished`、`dateModified`、`about`、`citation`。
- 有可见导航层级时才加 [`BreadcrumbList`](https://schema.org/BreadcrumbList)。

Google推荐 JSON-LD，但要求结构化数据描述可见主内容，并明确“正确标记”不保证富结果。对 AI Search 不存在特殊 schema；因此不要虚构 `AggregateRating`、Review、价格、作者、发布日期或 FAQ。[Google structured data intro](https://developers.google.com/search/docs/appearance/structured-data/intro-structured-data)、[Google structured data policies](https://developers.google.com/search/docs/appearance/structured-data/sd-policies)、[Google 生成式 AI 优化指南](https://developers.google.com/search/docs/fundamentals/ai-optimization-guide)

## 6. 明确不做或延后

- **不把 `llms.txt` 当作 Google GEO 项目。** Google明确表示其不使用该文件，加入它既不提升也不损害 Google 可见性；只有确认其他目标服务读取且维护成本可接受时才作为附加渠道。
- **不批量生成“格式 × 行业 × 浏览器 × 问题”薄页。** 只在页面有独立事实、示例或决策价值时拆分。
- **不为 GEO 单独复制正文或做隐藏 AI 文本。** 结构化数据、HTML 正文、截图、视频必须对同一实体保持一致。
- **不承诺排名、索引或 AI 引用。** Google和Bing都明确提交、合规或结构化数据不构成展示保证。
- **不先上复杂 Schema 图谱。** 首页实体、产品实体、文章和面包屑覆盖当前需求；真实内容出现新实体后再扩展。

## 7. 90 天验收口径

| 时间 | 最小交付 | 可验证结果 |
|---|---|---|
| 0–14 天 | robots、sitemap、canonical、唯一 title/description、公开正文、Search Console / Bing 验证 | P0 URL 可被 URL Inspection 抓取；canonical 一致；sitemap 无关键错误 |
| 15–45 天 | 补齐产品、格式、架构/隐私、安全、精度、快速开始等核心事实页；加入最小 JSON-LD | 结构化数据验证通过且与正文一致；Google/Bing 已发现核心 URL |
| 46–90 天 | 发布 3–6 篇有实测数据/样本/代码的一手内容；启用 IndexNow 与引用监测 | Search Console 有目标 query/page 数据；Bing AI Performance 可用时建立 citations 与 grounding query 基线 |

最终 KPI 应按漏斗分层：**已发现 URL → 已索引 canonical URL → 有目标查询展现 → 被 AI 引用的 URL / query → 文档、demo、试用或销售转化**。前一层没有数据时，不应直接用内容数量或 Schema 数量替代成果。
