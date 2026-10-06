const languageKey = "docviewkit-language";

export const english = new Map(Object.entries({
  "Viewer 和 Engine 源码采用 Apache-2.0。": "Viewer and Engine source code is licensed under Apache-2.0.",
  "联系支持与定制": "Contact support and customization",
  "联系服务": "Contact services",
  "联系与服务": "Contact and services",
  "GitHub 源码": "Source on GitHub",
  "Omni 隐私政策": "Omni Privacy Policy",
  "Omni 使用许可": "Omni License",
  "浏览器本地": "Browser-local",
  "免费安装，即刻开启设备本地的文档预览体验。": "Install for free and start viewing documents locally on your device.",
  "Safari 版本正在准备中，将通过 App Store 提供安装。": "The Safari version is in preparation, with installation planned through the App Store.",
  "结合页面导航、缩放、搜索与文字选择，让文档查阅自然融入浏览流程。": "Page navigation, zoom, search, and text selection make document reading a natural part of browsing.",
  "专注阅读体验": "Focus on reading",
  "文档在设备本地处理，来源权限按需授权，隐私政策清晰说明数据处理方式。": "Documents are processed locally on your device, with source permissions granted on demand and data handling explained in the privacy policy.",
  "由你选择文件、发起预览，并按需授权文档来源，轻松掌控每一次查看。": "Choose your files, start previews, and grant access to document sources on demand, keeping you in control of each viewing session.",
  "安装插件后，直接打开受支持文件，即可在 IDE 中预览。": "Install the plugin, then open a supported file to preview it in your IDE.",
  "Word、Excel、PowerPoint、PDF、OpenDocument、Apple Pages / Numbers / Keynote、WPS Office、XPS、RTF 与 CSV；相关模板和宏启用文件以安全只读方式呈现。": "Word, Excel, PowerPoint, PDF, OpenDocument, Apple Pages / Numbers / Keynote, WPS Office, XPS, RTF, and CSV, including read-only viewing of related templates and macro-enabled files.",
  "以只读方式呈现文档，隔离宏等主动内容，提供页面导航、内容查看与兼容诊断。": "View documents in read-only mode with active content isolated, plus page navigation and compatibility diagnostics.",
  "在 VS Code 与 JetBrains IDE 的编辑区内预览 Office、PDF、OpenDocument、iWork 等文件。本地解析、即开即看，让文档与代码同处一个工作空间。": "Preview Office, PDF, OpenDocument, iWork, and more in the editor area of VS Code and JetBrains IDEs. Local parsing keeps documents and code together in one workspace.",
  "DocViewKit Omni for IDE 在 VS Code 与 JetBrains IDE 中本地预览 Office、PDF、OpenDocument、iWork 等文档，让文档阅读融入开发流程。": "Preview Office, PDF, OpenDocument, iWork, and more locally in VS Code and JetBrains IDEs with DocViewKit Omni for IDE, bringing document reading into your development workflow.",
  "让软件产品在前端本地展示原文、搜索内容并定位引用。": "Display original documents, search content, and locate citations locally in your frontend app.",
  "Viewer、Engine API、界面定制、极简模式和水印均可按 Apache-2.0 使用。需要接入协助、专项兼容修复、性能优化或企业 SLA 时，可选择支持、定制和企业交付服务。": "Viewer, Engine APIs, UI customization, minimal mode and watermarks are available under Apache-2.0. Request support, customization or enterprise delivery for integration assistance, document compatibility work, performance optimization or an enterprise SLA.",
  "开源 Viewer 与 Engine 包含当前提供的全部格式，如 Office、WPS、PDF、OFD、OpenDocument 和 iWork，支持界面定制、极简模式和文字水印。在线 Demo 可直接体验你的文档，支持格式页提供详细功能说明。": "The open source Viewer and Engine include every available format, including Office, WPS, PDF, OFD, OpenDocument and iWork, plus UI customization, minimal mode and watermarks. Try your documents in the online demo and explore supported-format details.",
  "沿用现有权限体系。你的应用完成鉴权并获取附件，再交给 Viewer 在本地解析。接入指南提供私有附件示例，并说明跨域资源的 CORS 配置，让预览融入已有业务流程。": "Use your existing permissions. Your application authorizes access, fetches the attachment, and passes it to the Viewer for local parsing. The integration guide includes a private-attachment example and CORS guidance for cross-origin resources.",
  "如何预览业务系统中的私有附件？": "How do I preview private business attachments?",
  "复用 1 个 Viewer 组件，按需加载格式，并利用浏览器缓存。性能指南提供首开优化和冷、热启动测量方法，帮助你结合实际文件与运行环境优化体验。": "Reuse one Viewer component, load formats on demand, and use browser caching. The performance guide covers first-open optimization and cold- and warm-start measurement so you can tune the experience for your files and runtime.",
  "如何获得轻量、流畅的预览体验？": "How do I get lightweight, responsive previews?",
  "挂载界面只需 3 步：导入组件、放置 Viewer、调用 open(file)。open() 支持 File、Blob、ArrayBuffer、Uint8Array 这 4 类输入，并配有现成预览工具栏。完整接入文档涵盖静态资源部署和附件鉴权。": "Mount the UI in three steps: import the component, mount the Viewer, and call open(file). Four input types are supported: File, Blob, ArrayBuffer, and Uint8Array, with a ready-made preview toolbar. The integration guide covers asset deployment and attachment authorization.",
  "适合。普通 OA、审批、后台管理和 CRM/ERP 附件都是主要场景，企业 SaaS、AI 知识库和专业审阅也可使用同一组件。现成 Viewer 集成了预览界面，直接嵌入即可使用。": "Yes. Everyday OA, approvals, admin portals, and CRM/ERP attachments are primary use cases. Enterprise SaaS, AI knowledge bases, and professional review can use the same component. The ready-made Viewer includes the preview UI for direct embedding.",
  "Viewer 与 Engine 全部能力按 Apache-2.0 提供；付费服务覆盖支持、定制和企业交付。": "Every Viewer and Engine capability is available under Apache-2.0; paid services cover support, customization and enterprise delivery.",
  "普通 OA、审批流、后台管理与内部工具，嵌入组件即可拥有本地附件预览。": "Embed local attachment previews in everyday OA systems, approval workflows, admin portals, and internal tools.",
  "通过只读预览、主动内容隔离和资源管理，让用户专注查看文档；兼容信息可在诊断面板中查看。": "Read-only viewing, active-content isolation, and resource management keep the focus on the document. Compatibility details are available in the diagnostics panel.",
  "静态资源随应用托管，文档在前端本地解析，附件继续使用现有业务系统的存储与权限。": "Host static assets with your app and parse documents locally in the frontend, using your existing attachment storage and permissions.",
  "如何随现有应用部署？": "How does it fit my existing deployment?",
  "1 个 Viewer 组件复用多种格式；扩展格式按需加载，让应用启动更轻盈。": "One Viewer component serves multiple formats, with on-demand format loading for a lighter startup.",
  "查看支持格式与功能说明，选择你的文档开始体验。": "Explore supported formats and features, then try your own documents.",
  "本地解析 Office、WPS、PDF、OFD、OpenDocument 和 iWork，覆盖新版与旧版 Office，让多格式文档使用同一套预览体验。": "Parse Office, WPS, PDF, OFD, OpenDocument and iWork locally, including modern and legacy Office, through one consistent viewing experience.",
  "随应用部署静态资源，即可在前端完成文档解析与呈现": "Host static assets with your app for frontend-local document parsing and rendering.",
  "轻松部署": "Simple deployment",
  "解析、渲染与搜索均在前端本地完成，文件就地处理，交互快速响应": "Parse, render, and search locally in the frontend for responsive document interactions.",
  "纯前端，轻松部署": "Frontend-local, simple deployment",
  "DocViewKit 是适用于普通 OA 到企业应用的轻量文档预览组件，快速集成，在前端本地完成解析与呈现。": "DocViewKit is a lightweight document viewer for everyday OA attachments and enterprise apps, with quick integration and local frontend parsing and rendering.",
  "DocViewKit 是轻量级纯前端文档预览 SDK，可快速集成到普通 OA 附件预览、审批、CRM、ERP、网盘、企业 SaaS 和 AI 知识库。支持 Office、WPS、PDF、OFD 等格式，按需加载、本地解析，轻松嵌入现有应用。": "DocViewKit is a lightweight frontend-local document viewer SDK for quick integration into everyday OA attachments, approvals, CRM, ERP, cloud drives, enterprise SaaS, and AI knowledge bases. View Office, WPS, PDF, OFD, and more with on-demand loading and local processing, embedded in your existing app.",
  "支持 Office、PDF、OFD、ODF、iWork、CSV 和 RTF 等格式。带上你的实际文件，体验文档的字体、分页、图表与搜索效果。": "View Office, PDF, OFD, ODF, iWork, CSV, RTF, and more. Try your own files to explore fonts, pagination, charts, and search.",
  "在当前浏览器中完成解析与预览。": "Parse and view documents in your current browser.",
  "文件在本地处理": "Files stay local",
  "在浏览器中直接体验 DocViewKit 文档预览、搜索和来源定位，文件在本地处理。": "Try DocViewKit document viewing, search, and source location directly in your browser, with local file processing.",
  "查看开源许可与服务范围": "See the open source license and service scope",
  "查看完整接入与私有附件示例": "See complete integration and private-attachment examples",
  "查看按需加载与测量方法": "See on-demand loading and measurement",
  "什么时候需要专业服务？": "When should I request paid services?",
  "快速集成需要哪些步骤？": "What does quick integration involve?",
  "如何确认适合你的实际文件？": "How do I check compatibility with my own files?",
  "开源软件与专业服务如何选择？": "How do open source software and paid services fit together?",
  "普通 OA 和专业业务都能用吗？": "Does it fit everyday OA and professional workflows?",
  "如何处理文档中的主动内容？": "How is active document content handled?",
  "如何把搜索结果带回原文？": "How do search results link back to the source?",
  "浏览器与跨平台应用能共用一套 Viewer 吗？": "Can browsers and cross-platform apps share one Viewer?",
  "预览之外，还有哪些现成功能？": "What else does the Viewer include?",
  "只需导入组件、放置 Viewer、调用 open(file) 这 3 步即可挂载预览界面。原生 JavaScript、React、Vue、Angular 等项目均可复用；上线时还需配置静态资源托管和附件鉴权。": "Mount a preview UI in three steps: import the component, mount the Viewer, and call open(file). Reuse it in vanilla JavaScript, React, Vue, Angular, and other projects. Production setup also requires asset hosting and attachment authorization.",
  "如何快速接入所有主流前端框架？": "How do I integrate with every major frontend framework?",
  "如何快速定位原文？": "How do users locate source content?",
  "轻量体现在哪里？": "What makes it lightweight?",
  "为什么日常附件预览也适合用 DocViewKit？": "Why use DocViewKit for everyday attachment previews?",
  "支持哪些文档格式？": "Which document formats are supported?",
  "普通 OA 系统的附件预览适合用 DocViewKit 吗？": "Is DocViewKit suitable for everyday OA attachment previews?",
  "适合。普通 OA 附件、审批材料、后台管理和内部工具都是主要使用场景。导入现成 Viewer 组件并调用 open(file)，即可接入预览；无需使用底层 Engine API，也无需部署转换服务。轻量核心与按需格式加载减少初始加载内容，专业场景可进一步使用搜索与原文定位。": "Yes. Everyday OA attachments, approval documents, admin portals, and internal tools are primary use cases. Import the ready-made Viewer and call open(file) to add previews without using the Engine API or deploying a conversion service. A compact core and on-demand format loading reduce the initial load; professional workflows can also use search and source location.",
  "支持格式与边界": "Supported formats and boundaries",
  "新版 Microsoft Office": "Modern Microsoft Office",
  "Word（.docx、.docm、.dotx、.dotm）；Excel（.xlsx、.xlsm、.xltx、.xltm）；PowerPoint（.pptx、.pptm、.ppsx、.ppsm、.potx、.potm）。": "Word (.docx, .docm, .dotx, .dotm); Excel (.xlsx, .xlsm, .xltx, .xltm); PowerPoint (.pptx, .pptm, .ppsx, .ppsm, .potx, .potm).",
  "旧版 Microsoft Office": "Legacy Microsoft Office",
  "Word（.doc）、Excel（.xls）、PowerPoint（.ppt）文档预览。": "Preview Word (.doc), Excel (.xls) and PowerPoint (.ppt) documents.",
  "WPS 文字（.wps）、WPS 表格（.et）、WPS 演示（.dps）。": "WPS Writer (.wps), WPS Spreadsheets (.et) and WPS Presentation (.dps).",
  "PDF、OFD 与 XPS": "PDF, OFD and XPS",
  "固定版式文档预览：.pdf、.ofd、.xps、.oxps。": "Fixed-layout document viewing: .pdf, .ofd, .xps, .oxps.",
  "OpenDocument（ODF）": "OpenDocument (ODF)",
  "文本文档（.odt、.ott、.fodt）、电子表格（.ods、.ots、.fods）、演示文稿（.odp、.otp、.fodp）、绘图（.odg、.otg、.fodg）。": "Text documents (.odt, .ott, .fodt), spreadsheets (.ods, .ots, .fods), presentations (.odp, .otp, .fodp) and drawings (.odg, .otg, .fodg).",
  "Pages（.pages）、Numbers（.numbers）、Keynote（.key）单文件文档包。": "Pages (.pages), Numbers (.numbers) and Keynote (.key) single-file document packages.",
  "CSV 与 RTF": "CSV and RTF",
  "CSV 表格（.csv）与 RTF 富文本文档（.rtf）。": "CSV spreadsheets (.csv) and RTF rich-text documents (.rtf).",
  "Office、WPS、PDF、OFD 文档预览 SDK｜DocViewKit": "Office, WPS, PDF & OFD Document Viewer SDK | DocViewKit",
  "开发者文档 — DocViewKit": "Developer Documentation — DocViewKit",
  "在线 Demo — DocViewKit": "Online Demo — DocViewKit",
  "DocViewKit Omni 隐私政策": "DocViewKit Omni Privacy Policy",
  "DocViewKit Omni 最终用户许可协议": "DocViewKit Omni End User License Agreement",
  "DocViewKit 是纯前端文档预览 SDK，支持 Word、Excel、PowerPoint 新旧格式，以及 WPS、PDF、OFD、OpenDocument、iWork、XPS、CSV 和 RTF。文档在前端本地解析，可接入浏览器及跨平台应用的 WebView，无需上传或部署转换服务。": "DocViewKit is a frontend-local JavaScript document viewer SDK for modern and legacy Word, Excel and PowerPoint files, plus WPS, PDF, OFD, OpenDocument, iWork, XPS, CSV and RTF. Runs in browsers and compatible cross-platform WebViews. No uploads or conversion server.",
  "DocViewKit Omni 浏览器扩展隐私政策：文档在设备本地处理，不上传文档，不采集遥测。": "Privacy policy for the DocViewKit Omni browser extension: documents are processed on the device, are not uploaded, and no telemetry is collected.",
  "DocViewKit Omni 免费专有软件最终用户许可协议。": "End User License Agreement for the free proprietary DocViewKit Omni software.",
  "产品": "Product",
  "兼容保障": "Compatibility",
  "版本": "Plans",
  "文档": "Docs",
  "专业服务": "Paid services",
  "打开导航菜单": "Open navigation menu",
  "主导航": "Primary navigation",
  "DocViewKit 首页": "DocViewKit home",
  "产品特性摘要": "Product highlights",
  "产品优势": "Product advantages",
  "切换深浅主题": "Toggle color theme",
  "切换到浅色主题": "Switch to light theme",
  "切换到深色主题": "Switch to dark theme",
  "隐私政策导航": "Privacy policy navigation",
  "DocViewKit Omni 浏览器扩展": "DocViewKit Omni browser extension",
  "隐私政策": "Privacy Policy",
  "DocViewKit Omni 是本地、只读的文档预览扩展。我们不会接收你的文档，也不采集扩展遥测。": "DocViewKit Omni is a local, read-only document preview extension. We do not receive your documents or collect extension telemetry.",
  "生效日期": "Effective date",
  "2026 年 8 月 27 日": "August 27, 2026",
  "适用产品": "Applies to",
  "DocViewKit Omni for Chrome、Edge 和 Firefox": "DocViewKit Omni for Chrome, Edge, and Firefox",
  "我们不收集什么": "What we do not collect",
  "DocViewKit Omni 不会向 DocViewKit 或我们的服务器上传、发送或出售文档内容、文件名、文件路径、浏览历史、账户信息、使用统计或诊断遥测。": "DocViewKit Omni does not upload, transmit, or sell document content, filenames, file paths, browsing history, account information, usage statistics, or diagnostic telemetry to DocViewKit or our servers.",
  "本地文件处理": "Local file processing",
  "用户选择或拖入的本地文件只在浏览器扩展页面中解析和渲染。扩展不会自动把文件或其内容传输给 DocViewKit 或第三方转换服务。": "Local files selected or dropped by the user are parsed and rendered only within the browser extension page. The extension does not automatically transmit files or their contents to DocViewKit or a third-party conversion service.",
  "远程文档与来源权限": "Remote documents and origin permissions",
  "只有当用户主动选择预览某个 HTTP 或 HTTPS 文档链接时，扩展才会请求访问该链接所属的精确来源。扩展不会在安装时申请全站访问权限。": "The extension requests access to the exact origin of an HTTP or HTTPS document link only after the user explicitly chooses to preview it. The extension does not request site-wide access at installation.",
  "远程文档由用户的浏览器直接向原始网站请求，并在设备本地渲染。与普通网页请求一样，原始网站可能接收到网络请求所必需的信息，例如 IP 地址和标准请求头；这些信息不会发送给 DocViewKit。": "The user's browser requests a remote document directly from its source website and renders it on the device. As with a normal web request, the source website may receive information required for the network request, such as the IP address and standard request headers; this information is not sent to DocViewKit.",
  "扩展存储": "Extension storage",
  "扩展只会在浏览器会话存储中短暂保存用户主动选择的预览请求，以便把链接交给预览标签页。请求被读取后即删除。扩展不会持久保存文档内容、文件名、路径或浏览历史。": "The extension briefly stores a user-selected preview request in browser session storage so it can hand the link to the preview tab. The request is deleted after it is read. The extension does not persist document content, filenames, paths, or browsing history.",
  "诊断与剪贴板": "Diagnostics and clipboard",
  "预览失败时，用户可以主动复制经过脱敏的诊断摘要。该操作只写入用户设备的剪贴板，不会自动发送给我们。诊断摘要不包含文档内容、文件路径、文件名或原始网址。": "When a preview fails, the user may explicitly copy a redacted diagnostic summary. This action writes only to the clipboard on the user's device and does not automatically send anything to us. The summary does not include document content, file paths, filenames, or original URLs.",
  "数据共享与保留": "Data sharing and retention",
  "由于扩展不会向我们收集或传输用户数据，我们不会共享、出售或保留扩展用户数据。扩展不包含广告、遥测或远程运行时代码。": "Because the extension does not collect or transmit user data to us, we do not share, sell, or retain extension user data. The extension contains no advertising, telemetry, or remotely hosted runtime code.",
  "儿童隐私": "Children's privacy",
  "DocViewKit Omni 不以儿童为目标，也不会有意收集儿童的个人信息。": "DocViewKit Omni is not directed to children and does not knowingly collect children's personal information.",
  "政策更新": "Policy updates",
  "如果扩展的数据处理方式发生变化，我们会在本页面更新政策和生效日期，并按适用的浏览器商店规则向用户披露。": "If the extension's data handling changes, we will update this page and its effective date and disclose the change as required by the applicable browser store rules.",
  "联系我们": "Contact us",
  "如对本政策有疑问，请发送邮件至": "If you have questions about this policy, email",
  "返回 DocViewKit 首页": "Back to DocViewKit home",
  "许可协议导航": "License agreement navigation",
  "最终用户许可协议": "End User License Agreement",
  "本协议授予你免费使用 DocViewKit Omni 的有限许可。软件保持专有和闭源；免费不表示开源。": "This agreement grants you a limited license to use DocViewKit Omni at no charge. The software remains proprietary and closed source; free does not mean open source.",
  "2026 年 8 月 28 日": "August 28, 2026",
  "协议版本": "Agreement version",
  "DocViewKit Omni for VS Code 与 JetBrains IDE": "DocViewKit Omni for VS Code and JetBrains IDEs",
  "1. 协议双方与接受": "1. Parties and acceptance",
  "本最终用户许可协议（“协议”）由你与运营 DocViewKit 的个人或实体（“许可方”）订立。DocViewKit 是许可方用于本产品的品牌名称。Microsoft、GitHub、JetBrains 及其关联方不是本协议的一方，也不负责本产品。": "This End User License Agreement (the \"Agreement\") is between you and the individual or entity operating DocViewKit (the \"Licensor\"). DocViewKit is the brand name used by the Licensor for the Product. Microsoft, GitHub, JetBrains, and their affiliates are not parties to this Agreement and are not responsible for the Product.",
  "下载、安装或使用 DocViewKit Omni（“软件”）即表示你同意本协议。如果你不同意，请不要下载、安装或使用软件。": "By downloading, installing, or using DocViewKit Omni (the \"Software\"), you agree to this Agreement. If you do not agree, do not download, install, or use the Software.",
  "2. 免费许可": "2. Free license",
  "在你遵守本协议的前提下，许可方授予你一项免费、有限、非独占、不可转让、不可再许可且可撤销的许可，允许你在自己控制的兼容设备和 IDE 中安装并使用软件，用于个人或内部业务目的。": "Subject to your compliance with this Agreement, the Licensor grants you a royalty-free, limited, non-exclusive, non-transferable, non-sublicensable, and revocable license to install and use the Software on compatible devices and IDEs under your control for personal or internal business purposes.",
  "本协议不授予源代码许可、所有权或除上述使用权之外的任何权利。": "This Agreement grants no source-code license, ownership interest, or right other than the right to use the Software stated above.",
  "3. 使用限制": "3. Restrictions",
  "除适用法律明确允许外，你不得复制（合理备份除外）、修改、制作衍生作品、反向工程、反编译或反汇编软件；不得出售、出租、再许可、重新分发或以其他方式向第三方提供软件；不得移除权利声明、规避技术限制，或将软件用于违法、侵权或破坏性活动。": "Except where expressly permitted by applicable law, you may not copy the Software other than for reasonable backup, modify it, create derivative works, reverse engineer, decompile, or disassemble it; sell, rent, sublicense, redistribute, or otherwise make it available to a third party; remove proprietary notices, circumvent technical restrictions, or use it for unlawful, infringing, or destructive activity.",
  "4. 所有权与第三方组件": "4. Ownership and third-party components",
  "软件及其相关知识产权归许可方或其许可人所有。软件包含的第三方组件继续适用其各自的许可和声明；如第三方许可与本协议冲突，就该第三方组件以第三方许可为准。": "The Software and its related intellectual property are owned by the Licensor or its licensors. Third-party components included in the Software remain subject to their respective licenses and notices; if a third-party license conflicts with this Agreement, that license controls for the applicable component.",
  "5. 文档与数据": "5. Documents and data",
  "软件提供只读文档预览，不保证与原始创作软件完全一致，也不应作为文档编辑、恢复、合规归档或唯一备份工具。你应保留原始文件和独立备份，并负责确认预览结果是否适合你的用途。": "The Software provides read-only document preview and does not guarantee exact equivalence with the original authoring application. It must not be used as a document editor, recovery tool, compliance archive, or sole backup. You should retain original files and independent backups and are responsible for determining whether preview results are suitable for your purposes.",
  "软件的数据处理方式见": "The Software's data handling is described in the",
  "6. 更新、支持与可用性": "6. Updates, support, and availability",
  "许可方可以提供、变更或停止更新和支持，也可以修改或停止软件的全部或部分功能。除适用法律另有要求外，免费许可不包含服务级别、维护期限或响应时间承诺。": "The Licensor may provide, change, or discontinue updates and support and may modify or discontinue any or all Software functionality. Unless applicable law requires otherwise, the free license includes no service level, maintenance period, or response-time commitment.",
  "7. 终止": "7. Termination",
  "如果你违反本协议，许可自动终止。终止后，你必须停止使用并删除你控制的所有软件副本。关于所有权、免责声明、责任限制和其他按其性质应继续有效的条款在终止后继续有效。": "The license terminates automatically if you breach this Agreement. Upon termination, you must stop using and delete all copies of the Software under your control. Terms concerning ownership, disclaimers, limitations of liability, and any other terms that by their nature should survive will remain effective after termination.",
  "8. 免责声明": "8. Disclaimer of warranties",
  "在适用法律允许的最大范围内，软件按“现状”和“可用”状态提供，不附带任何明示、默示或法定保证，包括对适销性、特定用途适用性、不侵权、准确性、可靠性或不中断运行的保证。适用法律不允许排除的保证不受本条影响。": "TO THE MAXIMUM EXTENT PERMITTED BY APPLICABLE LAW, THE SOFTWARE IS PROVIDED \"AS IS\" AND \"AS AVAILABLE\" WITHOUT EXPRESS, IMPLIED, OR STATUTORY WARRANTIES, INCLUDING WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, NON-INFRINGEMENT, ACCURACY, RELIABILITY, OR UNINTERRUPTED OPERATION. WARRANTIES THAT APPLICABLE LAW DOES NOT ALLOW TO BE EXCLUDED ARE NOT AFFECTED.",
  "9. 责任限制": "9. Limitation of liability",
  "在适用法律允许的最大范围内，许可方及其许可人不对任何间接、附带、特殊、惩罚性或后果性损失，或数据、利润、收入、业务机会或商誉损失承担责任。对于与免费软件有关的全部索赔，许可方的累计责任以你在索赔发生前十二个月内为软件实际支付的金额为上限。适用法律不允许限制或排除的责任不受本条影响。": "TO THE MAXIMUM EXTENT PERMITTED BY APPLICABLE LAW, THE LICENSOR AND ITS LICENSORS WILL NOT BE LIABLE FOR INDIRECT, INCIDENTAL, SPECIAL, PUNITIVE, OR CONSEQUENTIAL DAMAGES, OR FOR LOSS OF DATA, PROFITS, REVENUE, BUSINESS OPPORTUNITY, OR GOODWILL. THE LICENSOR'S AGGREGATE LIABILITY FOR ALL CLAIMS RELATING TO THE FREE SOFTWARE IS LIMITED TO THE AMOUNT YOU ACTUALLY PAID FOR THE SOFTWARE DURING THE TWELVE MONTHS BEFORE THE CLAIM AROSE. LIABILITY THAT APPLICABLE LAW DOES NOT ALLOW TO BE LIMITED OR EXCLUDED IS NOT AFFECTED.",
  "10. 一般条款": "10. General terms",
  "本协议构成双方关于软件的完整协议。任何条款被认定不可执行时，应在必要的最小范围内调整或分离，其余条款继续有效。许可方未立即执行某项权利不构成放弃。本协议受适用于双方及本交易的强制性法律约束。": "This Agreement is the entire agreement between the parties concerning the Software. If any term is held unenforceable, it will be modified or severed only to the minimum extent necessary, and the remaining terms will continue in effect. A delay in enforcing a right is not a waiver. This Agreement is subject to mandatory laws applicable to the parties and the transaction.",
  "如对本协议有疑问，请发送邮件至": "If you have questions about this Agreement, email",
  "DocViewKit Omni for IDE｜IDE 文档预览": "DocViewKit Omni for IDE | Document preview in your IDE",
  "DocViewKit Omni for Browser｜浏览器文档预览": "DocViewKit Omni for Browser | Local document preview",
  "DocViewKit Omni for Browser 在 Chrome、Edge、Firefox 与 Safari 中本地预览 Office、PDF、OpenDocument、iWork 等文档。": "Preview Office, PDF, OpenDocument, iWork, and more locally in Chrome, Edge, Firefox, and Safari with DocViewKit Omni for Browser.",
  "不离开 IDE，": "Stay in your IDE—",
  "直接打开真实文档。": "open real documents directly.",
  "安装 VS Code 扩展": "Install the VS Code extension",
  "查看全部安装方式": "View all installation options",
  "全部格式免费": "Every format is free",
  "本地只读": "Local and read-only",
  "VS Code 与 JetBrains": "VS Code and JetBrains",
  "DocViewKit Omni 在 IDE 编辑区预览文档的界面示意": "DocViewKit Omni previewing a document inside an IDE editor",
  "本地预览 · 第 1 / 12 页": "Local preview · Page 1 of 12",
  "留在工作流": "Stay in your workflow",
  "双击文件，在普通编辑标签页中直接查看": "Double-click a file to view it in a normal editor tab",
  "文档不上传": "No document uploads",
  "解析、渲染和搜索都在 IDE 内完成": "Parsing, rendering, and search run inside the IDE",
  "按需加载": "On-demand loading",
  "只加载当前格式需要的组件，减少启动负担": "Load only the components required by the current format",
  "双端一致": "One Viewer, two IDE families",
  "VS Code 与 JetBrains 复用同一 Viewer 产物": "VS Code and JetBrains use the same Viewer build",
  "为开发者的文档工作流而做。": "Built for the developer document workflow.",
  "不新增独立工作台，预览跟随 IDE 的标签页、主题、语言和文件生命周期。": "No separate workspace. Preview follows the IDE's tabs, theme, language, and file lifecycle.",
  "原位查看": "View in place",
  "从文件树直接打开 DOCX、XLSX、PPTX、PDF 等资料，代码与文档保持在同一上下文。": "Open DOCX, XLSX, PPTX, PDF, and more from the file tree, keeping code and documents in the same context.",
  "内置阅读工具": "Built-in reading tools",
  "按格式提供页面或工作表导航、缩放、搜索、文字选择与安全链接。": "Navigate pages or sheets, zoom, search, select text, and follow safe links where supported.",
  "文件变化可见": "Stay current with file changes",
  "源文件被替换或修改后安全刷新，避免继续查看已经过时的内容。": "Refresh safely when the source file is replaced or changed so stale content does not remain visible.",
  "安全只读": "Safe, read-only viewing",
  "支持格式": "Supported formats",
  "一套扩展，覆盖常用与专业文档。": "One extension for everyday and specialist documents.",
  "下载安装。": "Download and install.",
  "现已提供": "Available now",
  "打开 VS Code Marketplace 页面并点击 Install。": "Open the VS Code Marketplace page and select Install.",
  "按提示在 VS Code 中完成安装。": "Follow the prompt to complete installation in VS Code.",
  "打开文档，或选择“打开方式 → DocViewKit Omni”。": "Open a document, or choose Open With → DocViewKit Omni.",
  "前往 VS Code Marketplace": "Open VS Code Marketplace",
  "复制命令": "Copy command",
  "打开 Settings / Preferences → Plugins。": "Open Settings / Preferences → Plugins.",
  "在 Marketplace 搜索“DocViewKit Omni”并安装。": "Search Marketplace for “DocViewKit Omni” and install it.",
  "重启 IDE 后，直接打开受支持文档。": "Restart the IDE, then open a supported document directly.",
  "插件已通过 JetBrains Marketplace 审核，可直接在 IDE 的插件市场中搜索并安装。": "The plugin has passed JetBrains Marketplace review. Search for it in your IDE’s plugin marketplace and install it directly.",
  "让文档回到开发上下文。": "Keep documents in the development context.",
  "安装扩展后，用项目中的真实文件验证格式与版面效果。": "After installing, validate format and layout with a real file from your project.",
  "免费安装 For IDE": "Install For IDE for free",
  "本地、只读的文档预览产品与 SDK。": "Local, read-only document preview products and SDK.",
  "资源": "Resources",
  "条款": "Legal",
  "许可协议": "License agreement",
  "不用上传，": "No upload required—",
  "在浏览器里直接看文档。": "view documents directly in your browser.",
  "选择或拖入本地文件，或者右键网页中的文档链接，在独立标签页中安全预览。文档在设备本地处理，不经过 DocViewKit 服务器。": "Choose or drop a local file, or right-click a document link on the web, to preview it safely in a separate tab. Documents are processed on your device and never pass through DocViewKit servers.",
  "安装 Chrome 扩展": "Install the Chrome extension",
  "免费使用": "Free to use",
  "不采集遥测": "No telemetry",
  "按需来源权限": "On-demand origin access",
  "DocViewKit Omni 在浏览器标签页中预览文档的界面示意": "DocViewKit Omni previewing a document in a browser tab",
  "本地扩展页": "local extension page",
  "docviewkit omni · 本地扩展页": "docviewkit omni · local extension page",
  "把文档放在这里": "Drop a document here",
  "或选择本地文件": "or choose a local file",
  "正在本地打开": "Opening locally",
  "本地文件": "Local files",
  "选择或拖入文件，字节只留在扩展页": "Choose or drop a file; its bytes stay in the extension page",
  "远程链接": "Remote links",
  "右键文档链接，主动发起一次预览": "Right-click a document link to start a one-time preview",
  "最小权限": "Minimal permissions",
  "只在需要时请求目标链接的精确来源": "Request only the exact origin required for the selected link",
  "无远程代码": "No remote code",
  "Viewer 与格式组件全部随扩展打包": "Viewer and format components are packaged with the extension",
  "更直接，也更可控。": "More direct, more controlled.",
  "打开本地文件": "Open local files",
  "点击扩展按钮进入预览页，选择文件或直接拖入；支持 Office、PDF、OpenDocument、iWork 等格式。": "Open the preview page from the toolbar, then choose or drop a file. Office, PDF, OpenDocument, iWork, and more are supported.",
  "预览网页链接": "Preview web links",
  "对文档链接使用右键菜单，确认来源权限后从原始网站读取并在本地渲染。": "Use the context menu on a document link, approve its origin, then fetch from the source website and render locally.",
  "隐私边界清楚": "Clear privacy boundaries",
  "使用方式": "How it works",
  "三步开始预览。": "Start previewing in three steps.",
  "安装扩展": "Install the extension",
  "从浏览器商店添加 DocViewKit Omni。": "Add DocViewKit Omni from your browser's extension store.",
  "选择文档": "Choose a document",
  "点击工具栏按钮，选择或拖入本地文件。": "Select the toolbar button, then choose or drop a local file.",
  "本地查看": "View locally",
  "在独立标签页中导航、缩放、搜索和选择文字。": "Navigate, zoom, search, and select text in a separate tab.",
  "商店安装适合日常使用；GitHub 安装包只用于测试或受控部署。": "Store installation is for everyday use. GitHub packages are only for testing or controlled deployment.",
  "打开 Chrome Web Store，点击“添加至 Chrome”，确认扩展权限即可。": "Open Chrome Web Store, select Add to Chrome, and confirm the extension permissions.",
  "前往 Chrome Web Store": "Open Chrome Web Store",
  "已在 Microsoft Edge Add-ons 发布。搜索“DocViewKit Omni”，点击“获取”并确认扩展权限即可安装。": "Now available on Microsoft Edge Add-ons. Search for “DocViewKit Omni”, select Get, and confirm the extension permissions to install.",
  "在 Edge Add-ons 搜索": "Search Edge Add-ons",
  "已在 Firefox Add-ons 发布。搜索“DocViewKit Omni”，点击“添加到 Firefox”并确认扩展权限即可安装。": "Now available on Firefox Add-ons. Search for “DocViewKit Omni”, select Add to Firefox, and confirm the extension permissions to install.",
  "在 Firefox Add-ons 搜索": "Search Firefox Add-ons",
  "macOS 版准备中": "macOS release in preparation",
  "开发者手动安装 Chrome / Edge 测试包": "Manually install the Chrome / Edge test package",
  "从 GitHub Release 下载 Chrome ZIP 并解压。": "Download and extract the Chrome ZIP from GitHub Releases.",
  "打开扩展管理页并启用开发者模式。": "Open the extensions management page and enable developer mode.",
  "选择“加载已解压的扩展”，指向解压目录。": "Choose Load unpacked and select the extracted directory.",
  "打开 v0.1.11 Release": "Open the v0.1.11 release",
  "先用一个真实文件试试看。": "Try it with a real file.",
  "免费安装 For Browser": "Install For Browser for free",
  "GitHub Release": "GitHub release",
  "轻量预览，": "Lightweight viewing, ",
  "快速集成，": "quick integration,",
  "融入每一种业务。": "for every business.",
  "开始技术评估": "Start technical evaluation",
  "阅读快速开始": "Read the quickstart",
  "在线体验": "Try the demo",
  "在线 Demo": "Online Demo",
  "打开在线 Demo": "Open the online demo",
  "验证真实文件": "Validate a real file",
  "轻量按需加载": "Lightweight, on-demand loading",
  "高性能本地渲染": "High-performance local rendering",
  "多种前端环境兼容": "Compatible frontend environments",
  "纯前端本地运行": "Runs locally in the frontend",
  "零外部运行时依赖": "Zero external runtime dependencies",
  "跨浏览器界面一致": "Consistent UI across browsers",
  "DocViewKit Viewer 展示原文并定位引用区域的界面示意": "DocViewKit Viewer displaying an original document and locating a cited region",
  "引用已定位": "Citation located",
  "前端本地": "Frontend-local",
  "纯前端运行": "Frontend-only",
  "无服务端依赖": "No server-side dependency",
  "核心与格式能力独立加载，未使用的能力不进入启动路径": "Core and format capabilities load independently; unused capabilities stay out of the startup path.",
  "解析、渲染与搜索均在浏览器完成，文档无需上传": "Parsing, rendering, and search happen in the browser, with no document upload required.",
  "不依赖 React、Vue 等框架或第三方前端运行库": "No dependency on frameworks such as React or Vue, or on third-party frontend runtime libraries.",
  "不部署文档转换服务，也不要求安装完整 Office": "No document-conversion service to deploy and no full Office installation required.",
  "免费支持全部格式": "Every format is free",
  "开源 Viewer 与 Engine 支持全部可用格式，按 Apache-2.0 使用与定制": "The open source Viewer and Engine support every available format under Apache-2.0.",
  "只加载当前文档需要的功能，打开更快、资源占用更低": "Load only what the current document needs for faster startup and lower resource use.",
  "广泛兼容": "Broad compatibility",
  "覆盖多种现代与传统格式，可用于浏览器和跨平台应用": "Covers modern and legacy formats in browsers and cross-platform apps.",
  "把文档预览和原文定位直接嵌入你的产品；用户权限、答案和业务流程仍由你的系统管理。": "Embed document viewing and source navigation directly in your product while your system continues to manage permissions, answers, and business workflows.",
  "从搜索结果或业务系统跳转到对应页面和区域，快速核对内容与引用。": "Jump from search results or business workflows to the relevant page and region to verify content and citations quickly.",
  "约定交付范围与验收标准": "Agreed delivery scope and acceptance criteria",
  "框架集成指南": "Framework integration guides",
  "React 与 Next.js 集成": "React and Next.js integration",
  "Vue 集成": "Vue integration",
  "Angular 集成": "Angular integration",
  "复制代码": "Copy code",
  "复制": "Copy",
  "已复制": "Copied",
  "复制失败": "Copy failed",
  "除了只读预览，还提供搜索、文字选择、对象定位和原文跳转，方便用户核对内容与引用。": "Beyond read-only viewing, DocViewKit provides search, text selection, object location, and source navigation so users can verify content and citations.",
  "查看运行环境要求": "View runtime requirements",
  "可用于浏览器及 Electron、Tauri、Ionic 和 Capacitor 的兼容 WebView。React Native、Flutter 等应用也可通过 WebView 接入，宿主需提供所需 Web 能力。": "Use browsers or compatible WebViews in Electron, Tauri, Ionic and Capacitor. React Native and Flutter apps can also embed the Viewer through a WebView with the required Web APIs.",
  "把搜索结果或 AI 回答直接带回原文，定位到相关页面、对象和区域。": "Take search results or AI answers directly back to the relevant page, object, or region in the source document.",
  "从日常附件预览到专业文档审阅，把多格式预览嵌入各类业务系统；需要时再接入搜索与原文定位。": "Embed multi-format viewing in business apps, from everyday attachments to professional review; add search and source location when needed.",
  "OA 与日常附件预览": "OA and everyday attachments",
  "内容与知识管理": "Content and knowledge management",
  "网盘、企业 SaaS 与 AI 知识库中的文档预览，也适用于合同审阅、VDR 预览、法律、审计与金融风控。": "Document viewing for cloud drives, enterprise SaaS, and AI knowledge bases, as well as contract review, VDR viewing, legal work, audit, and financial risk management.",
  "企业业务系统": "Enterprise business systems",
  "CRM、ERP 与人力资源系统中的档案、附件和业务文档预览。": "View records, attachments, and business documents in CRM, ERP, and HR systems.",
  "政务与合规": "Government and compliance",
  "电子政务与合规系统中的材料查阅、依据核验与原文定位。": "Review submissions, verify supporting evidence, and locate source content in e-government and compliance systems.",
  "产业与贸易协作": "Industry and trade collaboration",
  "工业、建筑与供应链管理，以及电子商务与跨境贸易。": "Industrial, construction, and supply chain management, plus e-commerce and cross-border trade.",
  "教育与培训": "Education and training",
  "教育与在线学习平台中的课件、资料和作业预览。": "View courseware, learning materials, and assignments in education and online learning platforms.",
  "开源软件": "Open source software",
  "适合直接嵌入文档预览或使用 Engine API，按 Apache-2.0 免费商用、修改和再分发。": "Embed document previews or use the Engine API, with commercial use, modification and redistribution under Apache-2.0.",
  "免费": "Free",
  "支持全部可用文件格式": "Every available file format",
  "高性能、轻量级前端 Viewer": "High-performance, lightweight frontend Viewer",
  "开始免费使用": "Start using it free",
  "用于验证集成路径、格式边界和真实文件表现，不授予公开生产权利。": "Validate integration, format boundaries, and real-file behavior without general production rights.",
  "按你的应用与真实文件提供接入支持、兼容性诊断和定制开发。": "Integration support, compatibility diagnosis and customization for your application and real documents.",
  "预览": "Preview",
  "按服务范围报价": "Quoted by service scope",
  "技术评估": "Technical evaluation",
  "企业交付": "Enterprise delivery",
  "限定范围的集成与兼容验证": "Scoped integration and compatibility validation",
  "文档始终在前端本地处理": "Documents always process locally in the frontend",
  "集成指导与兼容性诊断": "Integration guidance and compatibility diagnosis",
  "业务界面与渲染流程定制": "Custom business interfaces and rendering workflows",
  "专项保真修复与真实 Case 回归": "Document fidelity fixes with real case regressions",
  "约定响应时间与升级协助": "Agreed response times and upgrade assistance",
  "适合需要部署验证、性能定制、验收材料和 SLA 的企业产品。": "For enterprise products requiring deployment validation, performance customization, acceptance evidence and an SLA.",
  "部署验证、培训与 SLA": "Deployment validation, training and SLA",
  "按业务需求定制现有文件格式": "Customize existing file formats for your requirements",
  "企业性能优化定制：按实际格式裁剪未使用模块，缩小交付体积并优化首开性能": "Custom enterprise performance optimization: trim unused modules for the required format set to reduce delivery size and improve first-open performance",
  "单独开发并交付新的文件格式": "Develop and deliver new file formats",
  "定制服务端渲染": "Custom server-side rendering",
  "联系企业交付": "Contact enterprise delivery",
  "常见问题": "Frequently asked questions",
  "DocViewKit 会把文档上传到服务器吗？": "Does DocViewKit upload documents to a server?",
  "不会。解析和呈现在浏览器或 WebView 内完成；只有你主动提交的诊断信息才会发送给我们，源文件不会自动上传。": "No. Parsing and rendering happen inside the browser or WebView. Only diagnostics you explicitly submit are sent to us; source files are never uploaded automatically.",
  "它与在线 Office 有什么区别？": "How is this different from an online Office suite?",
  "DocViewKit 是只读、可嵌入的本地 Viewer SDK，不提供编辑、保存或协作，也不要求部署完整 Office 服务。": "DocViewKit is a read-only, embeddable, local Viewer SDK. It does not provide editing, saving, or collaboration, and it does not require a full Office server deployment.",
  "免费版本支持哪些格式？": "Which formats does the free version support?",
  "AI 开发工具可以直接完成集成吗？": "Can an AI coding tool integrate it directly?",
  "可以。AI Coding 工具能够直接读取稳定的纯文本文档和结构化接口说明，快速生成集成代码并根据项目环境调整。": "Yes. AI coding tools can read stable plain-text documentation and structured API descriptions, generate integration code, and adapt it to the project environment.",
  "创建预览应用": "Create a preview application",
  "能力": "Capabilities",
  "开发者": "Developers",
  "报告问题": "Report an issue",
  "在浏览器中打开一个真实文档。": "Open a real document in the browser.",
  "选择本地文件，直接体验预览、页面导航和全文搜索。文件只在当前浏览器标签页中处理。": "Choose a local file and try preview, page navigation, and full-text search. The file is processed only in this browser tab.",
  "DocViewKit 在线 Demo": "DocViewKit online demo",
  "选择本地文件": "Choose a local file",
  "打开示例": "Open sample",
  "准备中": "Preparing",
  "正在初始化本地 Viewer…": "Initializing the local Viewer…",
  "将文件拖到 Viewer 中打开": "Drop a file to open it in the Viewer",
  "可以使用自己的文件": "Use your own file",
  "查看格式边界": "View format boundaries",
  "跳到文档正文": "Skip to documentation",
  "文档主导航": "Documentation navigation",
  "打开文档目录": "Open documentation menu",
  "文档目录": "Documentation menu",
  "搜索文档": "Search documentation",
  "为 AI Coding 工具准备": "Built for AI coding tools",
  "网页与纯文本来自同一份内容源。": "Web and plain-text documentation share one source.",
  "读取 llms-full.txt": "Read llms-full.txt",
  "正在加载文档": "Loading documentation",
  "本文目录": "On this page",
  "本页内容": "On this page",
  "没有匹配文档": "No matching documents",
  "文档暂时无法加载": "Documentation is temporarily unavailable",
  "请刷新页面，或直接读取 /llms-full.txt。": "Refresh the page or read /llms-full.txt directly.",
  "跳到主要内容": "Skip to main content",
  "应用": "Applications",
  "开发文档": "Developer docs",
  "查看全部": "View all",
  "用户": "Users",
  "申请": "requests",
  "环境": "Environment",
  "测试": "Staging",
  "开发": "Development",
  "完成": "Done",
  "有效": "Active",
  "支持与定制": "Support and customization",
  "支持": "Support",
  "定制": "Customization",
  "如何添加文字水印？": "How do I add a text watermark?",
  "Viewer 与 Engine 可通过一个配置为页面、缩略图和打印结果添加文字水印，保留源文档原貌。": "Viewer and Engine add text watermarks to pages, thumbnails and printing through a single setting while preserving the source document.",
  "开源软件支持哪些格式？": "Which formats does the open source software support?",
  "先用你的真实文件确认格式、版面、搜索和原文定位效果，再选择合适的专业服务。": "Verify formats, layout, search and source navigation with your own documents, then choose the services you need."
}));

function resolveLocale() {
  if (typeof location === "undefined") return "en";
  const pathLocale = location.pathname.match(/^\/(en|zh-cn)(?:\/|$)/u)?.[1];
  if (pathLocale) return pathLocale === "en" ? "en" : "zh-CN";
  const parameter = new URLSearchParams(location.search).get("lang");
  if (parameter === "en" || parameter === "zh-CN") return parameter;
  const saved = document.cookie
    .split(";")
    .map((part) => part.trim())
    .find((part) => part.startsWith(`${languageKey}=`))
    ?.slice(languageKey.length + 1);
  if (saved === "en" || saved === "zh-CN") return saved;
  const browserLanguage = (navigator.languages?.length ? navigator.languages : [navigator.language])
    .find((language) => /^(?:en|zh)(?:-|$)/iu.test(language));
  return browserLanguage?.toLowerCase().startsWith("zh") ? "zh-CN" : "en";
}

export const locale = resolveLocale();

export function t(value) {
  return locale === "en" ? english.get(value) || value : value;
}

function translateTextNode(node) {
  const parent = node.parentElement;
  if (!parent || parent.closest("code, pre")) return;
  const original = node.nodeValue;
  const trimmed = original.trim();
  if (!trimmed) return;
  let translated = english.get(trimmed);
  if (!translated) return;
  const start = original.indexOf(trimmed);
  node.nodeValue = `${original.slice(0, start)}${translated}${original.slice(start + trimmed.length)}`;
}

function translateElement(element) {
  for (const attribute of ["aria-label", "placeholder", "title"]) {
    const value = element.getAttribute?.(attribute);
    if (value && english.has(value)) element.setAttribute(attribute, english.get(value));
  }
  if (element.matches?.('meta[name="description"]')) {
    const content = element.getAttribute("content");
    if (content && english.has(content)) element.setAttribute("content", english.get(content));
  }
}

function translateSubtree(root) {
  if (root.nodeType === Node.TEXT_NODE) {
    translateTextNode(root);
    return;
  }
  if (root.nodeType !== Node.ELEMENT_NODE && root.nodeType !== Node.DOCUMENT_NODE) return;
  if (root.nodeType === Node.ELEMENT_NODE) translateElement(root);
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT);
  let node;
  while ((node = walker.nextNode())) {
    if (node.nodeType === Node.TEXT_NODE) translateTextNode(node);
    else translateElement(node);
  }
}

function languageUrl(targetLocale) {
  const url = new URL(location.href);
  const prefix = targetLocale === "en" ? "/en" : "/zh-cn";
  const suffix = url.pathname.replace(/^\/(?:en|zh-cn)(?=\/|$)/u, "") || "/";
  url.pathname = `${prefix}${suffix}`;
  url.searchParams.delete("lang");
  return url;
}

export function initializeI18n() {
  document.documentElement.lang = locale;
  if (locale === "en") {
    translateSubtree(document);
    const observer = new MutationObserver((mutations) => {
      for (const mutation of mutations) {
        for (const node of mutation.addedNodes) translateSubtree(node);
      }
    });
    observer.observe(document.documentElement, { childList: true, subtree: true });
  }

  document.querySelectorAll("[data-language-toggle]").forEach((button) => {
    const targetLocale = locale === "en" ? "zh-CN" : "en";
    button.textContent = locale === "en" ? "中文" : "EN";
    button.setAttribute("aria-label", locale === "en" ? "切换到中文" : "Switch to English");
    button.addEventListener("click", () => {
      location.href = languageUrl(targetLocale).href;
    });
  });

  return locale;
}
