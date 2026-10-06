# 文档 SDK 与 Viewer 应用场景全网验证

> 调研日期与网页访问日期：2026-08-10。本文刻意不沿用项目原定位，只从买方任务重新判断。外部证据仅采用厂商官方产品页、官方文档/API、标准、监管/政府资料和厂商官方案例；不采用媒体、聚合站、咨询报告或搜索摘要。动态价格不含税、折扣和合同条款，采购前仍须复核。

## 1. 结论先行

上一轮列出的 12 类任务，大多确有真实需求，但“需求存在”不等于“适合成为 DocViewKit 的独立产品方向”。纠偏后的核心判断是：

1. **唯一达到 P0 验证优先级的主方向是“可核验的 AI 文档证据引用”**：宿主把答案、风险或结论绑定到文档中的原句、单元格、表格、图形或区域，用户点击即可回源复核。Box、Adobe、Microsoft、Apryse 已把这一交互产品化，证明需求，也证明它不是蓝海。
2. **抽取人工复核降为 P1 集成能力，不应做通用 HITL 平台。** Azure、AWS、Box 都输出置信度和几何坐标，复核任务成立；但 Google Document AI HITL 于 2024-01-16 deprecated、AWS A2I 自 2026-07-30 起不再接受新客户。DocViewKit 又没有 OCR，因此仅在 born-digital 文档，或外部 OCR/AI 已提供 anchor/bbox 时匹配较强。
3. **“文档即业务界面”、财务血缘、Preflight、隔离查看都成立为 P1 能力或准入项，而不是四个独立产品。** 它们共用“稳定目标定位 + overlay + 结构化诊断”原语。
4. **Diff、归档迁移、无障碍和培训/离线手册均有成熟需求，但当前能力缺口或上下游系统过大，降为 P2。** 泛用 Diff 已有 Adobe、Apryse、Nutrient；归档核心是批量识别、fixity、元数据、验证和迁移；LMS 核心是 SCORM/xAPI、进度与同步。
5. **数字销售室和通用附件预览作为主方向不成立。** DocSend 已把权限链接、访客身份、逐页分析、动态水印和数据室打包；Box、Microsoft、Google、Adobe 及 PDF.js 已将普通预览商品化。两者最多是 P3 分发场景。
6. **本地处理不是独特护城河。** Apryse 与 Nutrient 都已商业化浏览器客户端处理、离线/air-gap 或自托管。DocViewKit 若竞争，差异必须落在宿主中立、Office 原生 source mapping、长尾格式、可解释降级、包体/性能和可验收语料，而不能只说“文件不上传”。

最重要的竞争性裁决是：**市场命题强成立，DocViewKit 的差异化尚未成立。** Apryse 2026 当前 WebViewer 首页已经直接使用“document is the work”，并明确写出 LLM/抽取结果仍需人工核验、核验就在 Viewer 中完成。它既是本方向最强的一手验证，也是最强反证：这不是 DocViewKit 独创的新品类，而是已有成熟领跑者的赛道。[Apryse WebViewer](https://apryse.com/products/webviewer)。

这里的“可核验”只表示用户可以看到结论所指向的原始内容，**不表示 AI 引用一定正确、渲染达到法律认证、文件一定无恶意内容，或系统已具备取证资质**。

## 2. 裁决口径

- **成立**：至少两个独立一手来源证明当前任务存在，且当前 SDK 有可复用基础。
- **部分成立**：任务存在，但此前范围、买方强度或当前匹配被高估。
- **不成立（作为主方向）**：市场可能成立，但 Viewer 只是可替代底座，无法据现有证据形成独立购买理由。
- **未验证**：一手证据不足；本文不以故事补齐。
- **证据 A / B / C**：A 为两个以上独立来源且包含当前产品、监管或付费信号；B 为两个独立官方来源但对购买或本项目匹配仅有间接证明；C 为单一来源或主要依赖推断。
- **置信度**衡量本次裁决，不衡量未来市场一定成功。

文中使用 **[事实]**、**[推断]**、**[待验证假设]** 区分证据层级。厂商客户数、ROI 和案例成效均按“厂商披露”处理，不视为独立审计事实。

## 3. 十二类方向总矩阵

| # | 场景 | 对上一轮结论的裁决 | 证据 / 置信度 | 修订优先级 | DocViewKit 应扮演的角色 |
|---:|---|---|---|---|---|
| 1 | AI 证据 Viewer / 引用回原文 | **成立** | A / 高 | **P0** | 宿主中立的 evidence target、回源、临时高亮；不做模型与知识库 |
| 2 | AI/OCR/规则抽取人工复核 | **部分成立** | A / 高 | **P1 集成** | 接收外部字段、置信度和 anchor/bbox；不做 OCR、队列、人员管理和 HITL 后端 |
| 3 | Preflight / 生成质量 / 可观测性 | **部分成立** | A（PDF/结构）+ C（Office 视觉付费）/ 中高 | **P1 实验** | “本引擎可渲染性与降级报告”，不宣称通用 Office 合规或印前认证 |
| 4 | 文档即业务界面 / 审批叠加 | **部分成立** | A（文件级审批）+ B（对象级叠加）/ 中高 | **P1** | overlay/event 原语；权限、任务和状态归宿主 |
| 5 | 财务数字 / 表格血缘 | **部分成立** | A / 高 | **P1 垂直试点** | 单元格/公式/图表数据源回溯；不做财务报表平台或 XBRL 申报 |
| 6 | 语义 + 视觉 Diff | **成立于市场，部分成立于本项目** | A / 高 | **P2** | anchor 稳定后先做同格式、source-aware Office diff；不先做通用 PDF diff |
| 7 | 不可信文档安全查看 / 取证 | **安全查看部分成立；取证未验证** | A（需求）+ C（本项目资质）/ 中高 | **P1 准入；取证 P2/排除** | 只读、阻断活动内容、诊断；不得宣传沙箱、恶意检测、CDR 或取证认证 |
| 8 | 遗留档案迁移验收 / 长尾格式 | **部分成立** | A（档案任务）+ B（Viewer 角色）/ 高 | **P2 合作** | 人工验收与可视化辅助；批量 ID、fixity、迁移链路交给档案工具 |
| 9 | 无障碍替代视图 / 检查 | **部分成立** | A / 高 | **P2 / table stakes** | 先保证 Viewer UI 可访问；文档语义检查与重排需独立语料和标准验收 |
| 10 | 培训课件 / 设备手册离线指导 | **部分成立** | A（培训）+ B（现场手册）/ 中高 | **P2 集成** | 静态文档底座与热点；SCORM/xAPI、动画、进度、离线同步归宿主 |
| 11 | 数字销售室 / 阅读分析 | **作为主方向不成立** | C / 中高 | **P3** | 可提供页级事件/水印接口；身份、链接、权限、分析后台才是产品 |
| 12 | 通用附件预览 / LMS 普通预览 | **作为主方向不成立** | A / 高 | **P3 / 分发** | 免费入口、集成样例和获客面；不作为主要付费理由 |

**优先级纠偏结果：** P0 从多个“看似相邻的产品”收敛为一个底层：**基于标准 selector、可持久解析、能暴露映射质量的文档证据目标**。P1 才是把它用于抽取复核、审批、财务血缘和诊断；P2/P3 不应抢占主线。

## 4. 当前能力匹配与真实缺口

### 4.1 已经具备的底座

**[事实，本仓库]** 当前 `OfficeDocument` 已提供 `render`、`hitTest`、`searchText`、`listObjects`、`getObject` 和结构化 `diagnostics`；`DocumentObject.source` 能暴露 PPTX shape ID、XLSX A1/公式、DOCX 段落/表格范围、PDF 对象/流偏移等格式原生引用，并区分 `exact`、`derived`、`approximate`。详见 [`src/types.ts`](../../src/types.ts) 与 [`docs/ARCHITECTURE.md`](../ARCHITECTURE.md)。

**[事实，本仓库]** 当前安全模型默认每文档一个 Worker，限制 ZIP/XML/OLE/PDF 资源预算，不执行宏、ActiveX、OLE、PDF JavaScript 或外部关系，并输出 `ACTIVE_CONTENT_BLOCKED` 等诊断。详见 [`docs/SECURITY.md`](../SECURITY.md)。这比普通图片预览更适合构成“结果—原始对象”的可解释链路。

### 4.2 不能假装已经具备的能力

| 缺口 | 当前证据 | 对产品结论的影响 |
|---|---|---|
| 持久 anchor | 架构明确说明 object ID 只在一次打开的文档生命周期内稳定 | 不能把当前 `objectId` 存进 AI 答案、审批或审计记录后长期回放 |
| SourceRef 解析 | `viewer.reveal({kind:"source"})` 实际只接 `objectId` 或显式 `region`，没有接收完整 `SourceRef` 的解析器 | “点击引用回原对象”仍是增量能力，不是现成功能 |
| 通用 overlay | 有选中态、文本层、toolbar slot 和事件，但没有宿主注入的持久 mark/label/state API | 抽取复核、审批和风险标注尚不能用一个公共原语实现 |
| OCR / AI / 工作流 | README 明确 AI 与业务工作流归宿主；仓库没有 OCR、模型、人员/任务状态 | 通用 HITL、合同平台、LMS 或数据室不是当前产品 |
| Diff | 没有双文档对象对齐、变化模型、同步视图与审查状态 | 成熟市场需求不能直接转化为近期能力 |
| 无障碍文档语义 | Viewer chrome 有 ARIA 与文本选择，但没有跨格式完整阅读顺序、标题/表格语义、alt-text 检查和屏幕阅读器验收 | 不能把“Viewer UI 可操作”表述为“文档无障碍” |
| 安全/取证资质 | Security 文档明确 Worker 不是独立源安全沙箱，且安全敏感部署在披露渠道、签名和支持策略完成前视为 pre-release | 只能描述具体阻断和预算，不能作“安全打开”“恶意检测”或取证承诺 |

## 5. 分方向验证

### 5.1 AI 证据 Viewer：成立，P0

- **[事实]** Box AI 在 Preview 内回答问题，引用悬停显示对应源文本；Adobe Acrobat AI 的引用编号可跳到相关段落并高亮；Microsoft Copilot Studio 的生成式答案也把回答链接到知识源。三家独立产品共同证明“回答旁直接回源”已成为企业 AI 的标准信任交互。[Box AI for Documents](https://support.box.com/hc/en-us/articles/22158484213267-Box-AI-for-Documents)、[Adobe AI-generated answers](https://helpx.adobe.com/acrobat/using/get-ai-generated-answers.html)、[Microsoft RAG guidance](https://learn.microsoft.com/en-us/microsoft-copilot-studio/guidance/retrieval-augmented-generation)。
- **[事实/反证]** Box 明示数字、表格、图表等内容可能不准或不被处理；Adobe 明示 attribution 可能错误或跳到无关位置，要求用户核对。这说明证据 Viewer 的价值是真实的，也说明“有 citation”不等于“citation 正确”。[Box AI limitations](https://support.box.com/hc/en-us/articles/22158484213267-Box-AI-for-Documents)、[Adobe AI user disclosures](https://helpx.adobe.com/au/acrobat/desktop/use-acrobat-ai/understand-usage-policies/user-disclosures.html)。
- **[事实/竞争]** Apryse 已几乎逐字把“文档就是工作、模型结果仍需人工核验、Viewer 是核验发生处”写入 WebViewer 产品定位，并覆盖合同、理赔、申报、合规等用途；Nutrient 也已提供带引用、可自带模型的 AI Assistant。因此该方向需求很强，但不是空白市场。[Apryse WebViewer](https://apryse.com/products/webviewer)、[Nutrient AI Assistant](https://www.nutrient.io/sdk/ai-assistant/getting-started/)。
- **[待验证假设]** 可能的候选差异不是“也能显示 citation”，而是让 Office/ODF/iWork/老 Office/XPS 的原生对象或单元格成为 citation target，并在失败时报告 mapping quality。公开资料不足以断言 Apryse/Nutrient 没有 source mapping，必须用 API 与客户语料实测，不能先宣称独有。
- **[待验证假设]** AI/RAG 厂商会为这一点单独购买 SDK。现有证据只证明同类能力被套件和 SDK 商业化，未证明买方会选择 DocViewKit。

### 5.2 抽取人工复核：部分成立，P1 集成

- **[事实]** Adobe Acrobat Analyzer 把每个抽取属性链接到原文位置，支持高亮、赞/踩反馈和跨文件属性比较；Box Extract API 于 2026-03 提供字段级 citation 与 bounding box。二者直接验证“结构化结果 + 原件定位 + 人工确认”的交互。[Adobe validate attributes](https://helpx.adobe.com/acrobat-analyzer/using/attributes/validate-extracted-attributes.html)、[Adobe compare attributes](https://helpx.adobe.com/acrobat-analyzer/using/collections/compare-files.html)、[Box Extract citations and bounding boxes](https://support.box.com/hc/en-us/articles/50042285165331-Support-for-citations-and-bounding-boxes-in-Box-Extract-Agent-APIs-Mar-2026)。
- **[事实]** Azure Document Intelligence 与 AWS Textract 都返回文本/表格的 polygon 或 bounding box、置信度，并建议低置信度结果进入人工复核。[Azure Layout model](https://learn.microsoft.com/en-us/azure/ai-services/document-intelligence/prebuilt/layout?view=doc-intel-3.1.0)、[AWS Textract best practices](https://docs.aws.amazon.com/textract/latest/dg/textract-best-practices.html)。
- **[反证]** Google Document AI managed HITL 于 2024-01-16 deprecated，2025-01-16 后停止；AWS A2I 自 2026-07-30 起不再接受新客户且不再规划新功能。它们说明复核需求存在，但不能据此推断独立通用 HITL 产品普遍增长。[Google deprecations](https://docs.cloud.google.com/document-ai/docs/deprecation)、[Google HITL request review](https://docs.cloud.google.com/document-ai/docs/hitl/request-review)、[AWS A2I core components](https://docs.aws.amazon.com/sagemaker/latest/dg/a2i-getting-started-core-components.html)。
- **[反证/当前匹配]** DocViewKit 没有 OCR。Azure 对 Office 的 layout 支持也有边界，例如嵌入图片和 XLSX 表格分析并非等同于 PDF 页面 OCR。故当前最合理范围是 born-digital 对象，或外部服务已经给出页码、bbox、文本 quote/source address 的结果。
- **[推断]** 云厂商退出 managed HITL 可能给“宿主自有复核前端”留下集成缺口，但这不是增长率或愿付费的直接证明。

### 5.3 Preflight / 生成质量：部分成立，P1 实验

- **[事实]** Adobe Acrobat Pro Preflight 已形成成熟的 PDF/PDF-A/PDF-X/印前检查、fixup、问题对象定位和 XML/文本/PDF 报告；这证明 PDF 生产前检查有付费市场。[Adobe Preflight profiles](https://helpx.adobe.com/acrobat/using/preflight-profiles-acrobat-pro.html)、[Adobe Preflight reports](https://helpx.adobe.com/acrobat/using/preflight-reports-acrobat-pro.html)。
- **[事实/边界]** Microsoft `OpenXmlValidator` 返回描述、错误类型、节点、路径和 part，但它验证的是 Open XML schema/包结构；Open XML Markup Compatibility 文档又明确互操作性还取决于应用支持。Schema 通过不能证明 Word/Excel/PowerPoint 的视觉结果正确。[Microsoft Open XML validation](https://learn.microsoft.com/en-us/office/open-xml/word/how-to-validate-a-word-processing-document)、[Markup Compatibility](https://learn.microsoft.com/en-us/office/open-xml/general/introduction-to-markup-compatibility)。
- **[推断]** 当前 diagnostics 很适合固化成“DocViewKit 对此文件的可渲染性、缺失资源、阻断内容和降级报告”，但一手证据不足以证明跨 Office 视觉 QA 已有与 Adobe prepress 同强度的独立预算。
- **裁决边界**：不能把解析成功、schema 合法或某次截图相似，表述为 Office 兼容认证；先用文档生成 SaaS 的真实语料做 P1 设计伙伴实验。

### 5.4 文档即业务界面 / 审批叠加：部分成立，P1

- **[事实]** Box 可在文件 Preview 的 Activity 侧栏分配 Approval Task，审批者直接批准/拒绝；Box annotations 可把区域批注定位回预览原处。[Box comments and tasks](https://support.box.com/hc/en-us/articles/360043695954-Adding-Comments-and-Tasks)、[Box annotations](https://support.box.com/hc/en-us/articles/360048016153-Annotating-Documents)。
- **[事实]** Google Drive Approvals 与 SharePoint document-library approvals 都让用户在文件预览/库内审批，并处理编辑导致审批重置或取消等状态边界。[Google Drive approvals](https://support.google.com/drive/answer/9387535?hl=en)、[SharePoint approvals](https://support.microsoft.com/en-us/sharepoint/data-and-lists/approvals-in-lists-document-libraries)。
- **[反证]** 这些一手证据主要证明“文件级审批 + 批注”，没有充分证明“任意 Office 对象都是业务控件”本身是独立采购类别。
- **[推断]** DocViewKit 只应提供 mark/label/state、点击事件和 reveal；任务、权限、审批状态、审计日志仍归宿主。把它卖成完整审批平台会越过现有边界并正面撞上 Box/M365。

### 5.5 财务数字与表格血缘：部分成立，P1 垂直试点

- **[事实]** Microsoft Excel 自带 Trace Precedents/Dependents 和企业版 Inquire Cell Relationship；Workiva 则把 Spreadsheet、Document、Presentation 的 source/destination links、变化人/时间/内容和 linked-files report 产品化。这证明数字来源与跨文档一致性是明确任务。[Excel formula relationships](https://support.microsoft.com/en-us/excel/display-the-relationships-between-formulas-and-cells)、[Workiva linking](https://support.workiva.com/hc/en-us/articles/360036000591-What-is-Linking)、[Workiva linked files report](https://support.workiva.com/hc/en-us/articles/360036000671-Use-the-Linked-Files-Report)。
- **[事实/监管]** SEC Inline XBRL 把人可读财报与机器可读事实放在同一文档，点击数据点可查看标签、定义、期间等上下文；SEC 还免费发布 Renderer/Previewer 与验证错误。这是“数字—披露位置—机器事实”血缘的官方实例。[SEC Inline XBRL](https://www.sec.gov/data-research/structured-data/inline-xbrl)、[SEC XBRL Validation and Rendering](https://www.sec.gov/data-research/xbrl-validation-rendering)。
- **[当前匹配]** XLSX `SourceRef` 已有 sheet、A1 address 与 formula，是强基础；但还没有依赖图、命名区域、外部工作簿链、图表系列数据源、隐藏表风险规则或 XBRL 语义。
- **[推断]** 可先验证“AI/KPI 结论点回具体单元格与公式”的窄场景；不能据此宣称审计、合并报表或申报平台。

### 5.6 语义 + 视觉 Diff：市场成立，本项目 P2

- **[事实]** Word Legal Blackline、Adobe Compare Files、Apryse compare/multi-viewer 和 Nutrient comparison 都已覆盖文本、格式、图形或视觉比较；Nutrient 将比较作为需授权组件。市场与付费信号明确。[Word legal blackline](https://support.microsoft.com/en-us/word/compare-document-differences-using-the-legal-blackline-option)、[Adobe Compare Files](https://helpx.adobe.com/acrobat/using/compare-documents.html)、[Apryse diffing](https://docs.apryse.com/web/guides/diffing)、[Nutrient compare documents](https://www.nutrient.io/guides/web/comparison/compare-documents/)。
- **[反证]** Word 网页版不提供完整 Compare；Nutrient 明示自动视觉比较依赖页面对齐，版式变化时需三点手工对齐。这说明真实 Diff 不是对两张截图做减法。[Word browser differences](https://support.microsoft.com/en-us/word/differences-between-using-a-document-in-the-browser-and-in-word)、[Nutrient comparison](https://www.nutrient.io/guides/web/comparison/compare-documents/)。
- **[推断]** DocViewKit 的机会只在同格式、source-aware Office 差异，例如段落、单元格、公式、图表数据源和版面同时对齐；在持久 anchor 未完成前不应启动通用 Diff。

### 5.7 不可信文档安全查看 / 取证：部分成立，P1 准入

- **[事实]** Microsoft 对互联网、邮件附件、非可信位置和文件验证失败的 Office 文件使用 Protected View；Microsoft 365 E5/Suite 的 Safe Documents 还会扫描在 Protected View 中打开的文件。这是受限查看与检测进入企业付费套件的采购信号，但不等于客户为 Viewer 单项付费。[Microsoft Protected View](https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653)、[Defender Safe Documents](https://learn.microsoft.com/en-us/office365/servicedescriptions/microsoft-defender-for-office-365-features)。
- **[事实/竞争]** Apryse 与 Nutrient 都提供客户端、本地、自托管或 air-gapped 选择；Azure Document Intelligence 的 disconnected containers 甚至采用审批、承诺用量和企业协议。这证明离线/隔离有预算，也证明“本地处理”不是 DocViewKit 独有。[Apryse deployment](https://docs.apryse.com/web/guides/deployment-options)、[Nutrient deployment](https://www.nutrient.io/sdk/deployment-options/)、[Azure disconnected containers](https://learn.microsoft.com/en-us/azure/ai-services/document-intelligence/containers/disconnected?view=doc-intel-4.0.0)。
- **[当前反证]** Worker 不是独立源安全沙箱；本项目没有恶意软件检测、虚拟机 detonation、CDR 重建输出、取证写保护/证据链与安全认证，且安全敏感部署目前仍标为 pre-release。
- **裁决边界**：可说“活动内容不执行、外部关系不获取、资源有界并有诊断”；不能说“文件安全”“隔离沙箱”“取证工具”。安全工程是企业采购准入项，不是现成护城河。

### 5.8 遗留档案迁移与长尾格式：部分成立，P2 合作

- **[事实]** 英国国家档案馆 DROID/PRONOM 把精确识别文件格式和版本视为数字保存的基础；NARA 的风险框架与移交指南要求评估格式风险、重要属性、可接受格式和迁移行动。[DROID](https://www.nationalarchives.gov.uk/information-management/manage-information/preserving-digital-records/droid/)、[PRONOM overview](https://www.nationalarchives.gov.uk/help/pronom/overview.htm)、[NARA Digital Preservation Framework](https://www.archives.gov/preservation/digital-preservation/risk)、[NARA transfer guidance](https://www.archives.gov/records-mgmt/policy/transfer-guidance.html)。
- **[事实/反证]** DROID/PRONOM 免费提供批量格式识别；SEC Renderer、veraPDF 等领域工具也可免费验证特定标准。归档项目核心还包括 fixity、元数据、格式 ID、重要属性、转换、审计轨迹和批量处理，Viewer 只覆盖人眼验收的一部分。[veraPDF](https://verapdf.org/software/)。
- **[推断]** 老 Office、ODF、iWork、WPS、XPS 等支持可成为迁移验收插件，但必须用客户档案语料逐格式证明质量；不能从“格式在列表中”推断档案机构愿意采购。

### 5.9 无障碍替代视图与检查：部分成立，P2 / table stakes

- **[事实/监管]** 美国 Section 508 标准覆盖联邦机构开发、采购、维护或使用的软件与电子文档，电子文档测试基线明确包括 spreadsheet、presentation 等非 Web 文档；W3C WCAG 2.2 要求文本替代、信息关系、意义顺序、键盘操作和 reflow 等。[U.S. Access Board ICT standards](https://www.access-board.gov/ict/about/)、[ICT Testing Baseline for Electronic Documents](https://ictbaseline.access-board.gov/document-baselines/)、[WCAG 2.2](https://www.w3.org/TR/WCAG22/)。
- **[事实/竞争]** Microsoft Accessibility Checker 在 Word、Excel、PowerPoint、Web 和 Mac 中免费定位对象并给出错误/警告；Adobe Acrobat Pro 有 PDF 标签、阅读顺序、替代文本和无障碍报告工具。[Microsoft Accessibility Checker rules](https://support.microsoft.com/en-US/accessibility/office-accessibility/rules-for-the-accessibility-checker)、[Adobe PDF accessibility](https://helpx.adobe.com/au/acrobat/using/create-verify-pdf-accessibility.html)。
- **[反证]** 自动检查也不能捕获所有问题，Microsoft 明确要求视觉复核；当前 DocViewKit 解析语义又不足以跨格式证明阅读顺序、表格头、alt text 与重排正确。
- **[推断]** 先把 Viewer 本身做成键盘/屏幕阅读器可用是准入项；“替代视图/检查器”必须另建语义模型和标准验收，暂不作为主卖点。

### 5.10 培训课件 / 设备手册离线指导：部分成立，P2 集成

- **[事实]** iSpring 将 PowerPoint 转成 HTML5、SCORM 1.2/2004、AICC、xAPI、cmi5，保留动画并加入测验；Presenter 当前公开价为 570 美元/作者/年，LMS 明示移动应用支持离线学习。这证明“复用 PPT 做培训”和离线学习有付费市场。[iSpring Presenter](https://www.ispringsolutions.com/ispring-presenter)、[iSpring LMS pricing](https://www.ispring.com/pricing)。
- **[事实]** IBM Maximo Mobile 支持为离线使用自动下载附件；Microsoft Dynamics 365 Field Service 明确面向无网络或网络不稳的偏远/地下现场。这验证了离线手册是现场工作流的一部分。[IBM Maximo Mobile properties](https://www.ibm.com/docs/en/masv-and-l/maximo-manage/cd?topic=mobile-maximo-properties)、[Dynamics 365 Field Service offline](https://learn.microsoft.com/en-us/dynamics365/field-service/mobile/set-up-offline-profile)。
- **[反证]** DocViewKit 没有 PPT 完整动画/触发器、SCORM/xAPI、测验、学习进度、移动下载同步或现场工单；这些才是 iSpring/LMS/Field Service 的主要购买内容。
- **[推断]** 可做宿主的静态课件/手册原版面底座与热点 overlay，不应做 LMS 或现场服务平台。

### 5.11 数字销售室 / 阅读分析：作为主方向不成立，P3

- **[事实]** Dropbox DocSend 的 Performance/Activity 提供逐页停留、Top Pages、版本对比、dropoff、访客身份/设备/位置和导出；Advanced Data Rooms 又提供数据室分析，动态水印仅在 Advanced/Advanced Data Rooms 等高阶计划开放。[DocSend Performance](https://help.dropbox.com/share/dropbox-docsend-documents-performance-tab)、[DocSend Activity](https://help.dropbox.com/share/dropbox-docsend-documents-activity-tab)、[DocSend watermarks](https://help.dropbox.com/share/dropbox-docsend-watermarks)、[DocSend data-room analytics](https://help.dropbox.com/organize/dropbox-docsend-data-room-analytics)。
- **[反证]** DocSend 还明示停留时间存在前台标签、VPN、重开标签等测量边界；“页面停留”不是内容理解的可靠证明。[DocSend visitor time](https://help.dropbox.com/share/dropbox-docsend-visitor-time-on-document)。
- **[推断]** 市场和付费成立，但产品核心是身份、链接、权限、NDA、通知、分析存储和数据室；Viewer 只贡献渲染与事件。最多提供宿主可选的页可见事件和动态水印接口，不能据此建立新定位。
- **证据限制**：本项由 DocSend 同一官方产品家族的功能和付费层交叉验证，未取得第二家独立数字销售室厂商的一手证据，因此证据降为 C；“已有一个成熟付费替代”足以否定其近期主方向，却不足以估算整个市场。

### 5.12 通用附件预览 / LMS 普通预览：作为主方向不成立，P3

- **[事实]** Box Content Preview 可嵌入并预览 120+ 文件类型；Microsoft Graph/SharePoint Embedded 生成短时 iframe 预览 URL；Google Drive 可预览 Office、PDF、XPS 等，但官方也提示预览为缩放版本、显示可能不同。[Box Content Preview](https://developer.box.com/guides/embed/ui-elements/preview/)、[SharePoint Embedded preview](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/build/preview-files)、[Google Drive supported files](https://support.google.com/drive/answer/37603?hl=en)。
- **[事实/免费替代]** Microsoft Edge 内置本地、在线及网页嵌入 PDF 的查看、表单、批注和朗读；Adobe PDF Embed API 免费且不限嵌入次数；Mozilla PDF.js 以 Apache-2.0 提供浏览器 PDF Viewer。PDF-only 与普通附件查看已高度商品化。[Microsoft Edge PDF reader](https://learn.microsoft.com/en-us/deployedge/microsoft-edge-pdf)、[Adobe PDF Embed API](https://developer.adobe.com/document-services/apis/pdf-embed/)、[PDF.js](https://mozilla.github.io/pdf.js/getting_started/)。
- **[事实]** Canvas 的官方学生/基础指南把 DocViewer 作为作业查看、教师批注和反馈入口，说明 LMS 预览很常见，但价值来自作业、评分与反馈系统。[Canvas Student Guide](https://community.canvaslms.com/html/assets/Canvas_Student_Guide.pdf)。
- **裁决**：通用预览仍应保持免费/低摩擦，作为演示、分发和集成入口；不能作为主要商业差异。

## 6. 竞争替代分析

| 替代类别 | 已验证能力 | 对 DocViewKit 的直接压力 | 仍可能存在的窄缝 |
|---|---|---|---|
| Apryse WebViewer | 客户端 Office/PDF、多格式、批注、比较、编辑、AI 人工验证、离线；商业许可 | **最高**，其产品页几乎已占据“高风险文档验证”叙事 | 宿主中立 source mapping、ODF/iWork/WPS/XPS、透明降级、体积/价格；均需实测而非宣传 |
| Nutrient | 浏览器 Office/PDF、AI 引用、自带模型、Diff、本地/自托管/air-gap，年度组件许可 | **最高**，本地与隐私不再独特 | 不转 PDF 的原生 source 语义、开放 anchor/diagnostics、较轻部署 |
| Microsoft Office / SharePoint / Edge | Office 原生查看/编辑/审批、Copilot 引用与比较、Protected View、Graph preview、免费 OOXML validator/Accessibility Checker | 套件客户的默认选择；Office Online Server 可本地但较重 | 非 Microsoft 存储、无 WOPI/专用服务器、长尾格式、宿主自带 AI |
| Box | 120+ 预览、批注/审批、AI 引用、Extract/OCR、工作流；按席位与 AI units 商业化 | 内容已经在 Box 的客户没有理由再买通用 Viewer | 文件无需进入 Box、浏览器原字节、本地/私有宿主、格式原生 target |
| Adobe | 免费 PDF Embed、Acrobat AI 引用、Analyzer 抽取复核、PDF Diff/Preflight/Accessibility | PDF-only、PDF preflight、PDF diff 基本无切口 | Office/ODF/iWork/老格式的对象证据与实际 renderability |
| Google Workspace / Document AI | Drive 普通预览、Document AI text anchors/bbox；managed HITL 已退出 | 普通预览与云抽取很强，HITL 后端退出 | 做外部 Document AI 结果的宿主自有复核界面，但不是做 OCR |
| Azure / AWS | OCR/布局/字段/置信度/bbox、用量定价；Azure 有 disconnected containers | 抽取能力与企业采购已成熟 | 作为结果适配器与 evidence viewer 的上游，而不是正面竞争模型 |
| ONLYOFFICE / Collabora | 服务器端 Office 查看、编辑、协作、WOPI/自托管；Collabora 可 air-gap | 客户只要完整编辑协作时明显更合适 | 真正只读、无需服务器/WOPI、轻量嵌入 |
| PDF.js / 浏览器内置 / Google Drive | 免费或套件内普通预览 | 压低基础 Viewer 付费意愿 | 非 PDF 高保真、对象 API、结构化诊断与证据 target |

### 6.1 买方与付费信号

| 可能买方 | 已观察到的一手付费/采购信号 | 证据能证明什么 | 不能证明什么 |
|---|---|---|---|
| AI、合同、理赔、文档处理 SaaS | Apryse 生产使用需商业 key；Nutrient 为年度、多年、按组件与部署授权；Adobe/Box 将 AI、比较或 Extract 放入商业套餐 | 嵌入式验证、批注、比较和抽取组件能作为 SDK/套件收费 | 买方会为 DocViewKit 的 source mapping 另付费 |
| 文档生成、申报、出版团队 | Acrobat Pro Preflight 为付费能力；SEC 官方 Renderer/Previewer 同时用于验证与实际呈现 | “发出前发现错误”是明确任务；PDF/监管输出尤其强 | 普通 Office 视觉 QA 已有同强度独立预算 |
| 金融、报告与合规团队 | Workiva 将跨 Spreadsheet/Document/Presentation linking、历史和 XBRL 工作流商业化 | 数字一致性与血缘能进入企业报告采购 | 只读 Viewer 足以替代 Workiva 或审计系统 |
| 安全/隔离环境 | Safe Documents 需要 M365 E5/Defender Suite；Nutrient air-gap 需离线许可；Azure disconnected containers 需审批、承诺用量和企业协议 | 隔离、检测与离线部署有企业预算 | 纯客户端或无上传本身足以赢单 |
| 培训与现场作业 | iSpring Presenter 当前公开价 570 美元/作者/年，LMS 另按活跃用户；Maximo/Dynamics 把离线附件放进移动作业产品 | 课件复用、进度、离线同步有付费市场 | 只显示 PPT/PDF 就能替代 LMS/Field Service |
| 销售/融资材料 | DocSend 将水印、NDA、权限和数据室分析限制在 Advanced/Data Rooms 等付费层 | 阅读分析与受控分享可以收费 | 页级事件本身是独立产品或可靠意向指标 |

**[事实/付费代理]** Nutrient 采用年度组件许可并对 air-gap 许可询价；Apryse 生产使用需要商业 key；Box Business/Enterprise 将 AI、工作流和 API 额度放入付费套餐；Google Document AI、AWS Textract、Azure Document Intelligence 按页/用量或承诺量收费；Adobe Acrobat Pro/Studio 与 DocSend Advanced/Data Rooms 把 Diff、AI、Preflight、水印或分析放在付费层。它们证明相关工作已有预算。

**[边界]** 这些只是“类别能收费”的代理证据，不能证明 DocViewKit 的品牌、质量、格式组合或定价已获得买方认可。公开厂商客户数和 ROI 也只能作为厂商披露。

## 7. 建议的产品原语与边界

### 7.1 第一原语：基于 W3C Web Annotation 的持久目标

不要从零发明只含 `objectId + rect` 的 `DocumentAnchor`。W3C Web Annotation Data Model Recommendation 已定义 `target/source/selector/state`，以及 Fragment、TextQuote、TextPosition、DataPosition、SVG 和 Range selector；标准还明确指出 TextPosition 在文档变化后很脆弱，应配合 state。[Web Annotation Data Model](https://www.w3.org/TR/annotation-model/)。

建议在标准模型上增加 DocViewKit 扩展，而不是替代标准：

```ts
interface DocumentEvidenceTarget {
  /** Host URI/URN for the source resource, following the W3C target model. */
  source: string;
  state: {
    type: "DocViewKitDocumentState";
    documentHash: string;
    versionId?: string;
  };
  selector: readonly (
    | { type: "TextQuoteSelector"; exact: string; prefix?: string; suffix?: string }
    | { type: "TextPositionSelector"; start: number; end: number }
    | { type: "DocViewKitRegionSelector"; unitIndex: number; region: Rect }
    | { type: "DocViewKitSourceSelector"; source: SourceRef }
  )[];
}

interface ResolvedEvidenceTarget {
  target: DocumentEvidenceTarget;
  usedSelector: number;
  mapping: "exact" | "derived" | "approximate" | "unresolved";
}
```

这里的 `mapping` 是 `resolveTarget()` 的解析结果，不是持久 target 对事实正确性的自我声明。最小 API 是 `resolveTarget()`、`reveal()`、`mark()`、`unmark()` 与对象/mark 点击事件。解析时应返回实际命中的 selector、映射质量、失败原因和候选，不应静默猜测。业务状态与持久化继续由宿主负责。

这是 **W3C-shaped extension**：复用 W3C 的 target/selector/state 结构，但 `DocViewKitRegionSelector` 与 `DocViewKitSourceSelector` 是项目扩展，不能宣称其本身已标准化。跨格式使用 `TextQuoteSelector`/`TextPositionSelector` 还必须先定义确定性、版本化的 normalized text stream；否则相同文件在不同解析器版本或格式适配器中的 `start/end` 无法可靠重放。

### 7.2 第二原语：结构化 renderability 报告

把已有 `Diagnostic` 固化为带 schema/version、文档 hash、unit/object/SourceRef、严重度、fidelity、phase、稳定 code 和修复提示的报告。产品命名应是 **renderability/compatibility report**，避免与 PDF/X、PDF/A、Office schema 合规或恶意扫描混同。

### 7.3 明确不做

- 不做通用 AI、RAG、OCR、抽取模型或置信度算法。
- 不做审批、人员队列、权限、审计存储、数据室和阅读分析后台。
- 不做完整 Office 编辑器、LMS、SCORM/xAPI 运行时或现场服务系统。
- 不在 anchor 未稳定前做通用 Diff。
- 不在安全审计、签名、披露与支持策略完备前宣称安全沙箱、CDR、恶意检测或取证。

## 8. 建议验证顺序与否证门槛

以下是 **[待验证假设]**，不是市场事实：

1. **P0：AI 引用回源样例。** 用 DOCX、XLSX、PPTX、ODF 各一组真实文件，外部生成答案与 selector；关闭重开后仍能解析，文档改版时明确返回精确、近似或失效。若客户只接受 PDF 页码/bbox、对 Office 原生对象无额外价值，则核心差异被否证。
2. **P1：外部抽取复核适配。** 先适配一种公开坐标模型（Azure、Textract、Document AI 任选其一）和一种 born-digital `SourceRef` 模型。若主要买方样本都是扫描件且要求内置 OCR，则当前产品匹配被否证，应合作而非扩张。
3. **P1：renderability 报告。** 由文档生成 SaaS 提供真实失败语料，验证 diagnostics 是否能在发出前拦截其支持事故。若错误只在 Word/Excel 私有布局中可见且本引擎无法可靠检测，则不应产品化为 Preflight。
4. **P1：安全部署准入。** 完成私有漏洞报告渠道、发布签名、SBOM/依赖审计、支持策略、CSP 部署模板和恶意语料 fuzz 证据；此前不以“安全查看器”销售。
5. **P1 垂直试点：财务血缘。** 在 selector 稳定后，仅验证“结论/图表点回单元格与公式”；若买方需要完整依赖计算、合并报表和审计控制，则应与财务平台集成而非扩张。
6. **P2：同格式 Diff、档案和无障碍。** 只有当上述 selector/diagnostics 已被设计伙伴复用，且客户愿提供验收语料时再启动；不得用格式数量或单个演示推导市场成功。

## 9. 一手来源登记

除特别注明外，页面未显示稳定发布日期时记为“未显示/持续更新”；所有来源均于 **2026-08-10** 访问。

| 主题 | 一手来源 | 页面日期或官方节点 |
|---|---|---|
| 标准 anchor | [W3C Web Annotation Data Model](https://www.w3.org/TR/annotation-model/) | W3C Recommendation，2017-02-23 |
| AI 引用 | [Box AI for Documents](https://support.box.com/hc/en-us/articles/22158484213267-Box-AI-for-Documents) | 未显示 |
| AI 引用 | [Adobe AI-generated answers](https://helpx.adobe.com/acrobat/using/get-ai-generated-answers.html) | 2026-06-07 |
| AI 引用边界 | [Adobe AI user disclosures](https://helpx.adobe.com/au/acrobat/desktop/use-acrobat-ai/understand-usage-policies/user-disclosures.html) | 未显示/持续更新 |
| RAG 引用 | [Microsoft Copilot Studio RAG guidance](https://learn.microsoft.com/en-us/microsoft-copilot-studio/guidance/retrieval-augmented-generation) | 未显示/持续更新 |
| 直接竞品 | [Apryse WebViewer](https://apryse.com/products/webviewer)；[deployment](https://docs.apryse.com/web/guides/deployment-options)；[license FAQ](https://docs.apryse.com/web/faq/add-license) | 未显示/持续更新 |
| 直接竞品 | [Nutrient Office Viewer](https://www.nutrient.io/sdk/solutions/office-viewer/)；[AI Assistant](https://www.nutrient.io/sdk/ai-assistant/getting-started/)；[pricing](https://www.nutrient.io/sdk/pricing/) | 未显示/持续更新 |
| 抽取复核 | [Adobe Analyzer validate](https://helpx.adobe.com/acrobat-analyzer/using/attributes/validate-extracted-attributes.html) | 2026-02-27 |
| 抽取复核 | [Adobe Analyzer compare](https://helpx.adobe.com/acrobat-analyzer/using/collections/compare-files.html) | 2026-06-02 |
| 抽取坐标 | [Box Extract citations/bboxes](https://support.box.com/hc/en-us/articles/50042285165331-Support-for-citations-and-bounding-boxes-in-Box-Extract-Agent-APIs-Mar-2026) | 2026-03 GA |
| 抽取坐标 | [Azure Layout model](https://learn.microsoft.com/en-us/azure/ai-services/document-intelligence/prebuilt/layout?view=doc-intel-3.1.0)；[AWS Textract bbox](https://docs.aws.amazon.com/textract/latest/APIReference/API_BoundingBox.html) | 未显示/持续更新 |
| 抽取复核边界 | [AWS Textract best practices](https://docs.aws.amazon.com/textract/latest/dg/textract-best-practices.html) | 未显示/持续更新 |
| HITL 反证 | [Google Document AI deprecations](https://docs.cloud.google.com/document-ai/docs/deprecation) | HITL deprecated 2024-01-16；停止 2025-01-16 |
| HITL 反证 | [Google HITL request review](https://docs.cloud.google.com/document-ai/docs/hitl/request-review) | 官方标明 2025-01-16 后不可用 |
| HITL 反证 | [AWS A2I core components](https://docs.aws.amazon.com/sagemaker/latest/dg/a2i-getting-started-core-components.html) | 新客户截止 2026-07-30 |
| Preflight | [Adobe profiles](https://helpx.adobe.com/acrobat/using/preflight-profiles-acrobat-pro.html)；[reports](https://helpx.adobe.com/acrobat/using/preflight-reports-acrobat-pro.html) | 2026-02-26 |
| Open XML | [Microsoft validation](https://learn.microsoft.com/en-us/office/open-xml/word/how-to-validate-a-word-processing-document) | 2025-01-21 |
| Open XML 边界 | [Markup Compatibility](https://learn.microsoft.com/en-us/office/open-xml/general/introduction-to-markup-compatibility) | 未显示/持续更新 |
| 审批 | [Box tasks](https://support.box.com/hc/en-us/articles/360043695954-Adding-Comments-and-Tasks)；[Google Drive approvals](https://support.google.com/drive/answer/9387535?hl=en)；[SharePoint approvals](https://support.microsoft.com/en-us/sharepoint/data-and-lists/approvals-in-lists-document-libraries) | 未显示/持续更新 |
| 对象批注 | [Box annotations](https://support.box.com/hc/en-us/articles/360048016153-Annotating-Documents) | 未显示/持续更新 |
| 财务血缘 | [Microsoft formula relationships](https://support.microsoft.com/en-us/excel/display-the-relationships-between-formulas-and-cells) | 未显示/持续更新 |
| 财务血缘 | [Workiva linking](https://support.workiva.com/hc/en-us/articles/360036000591-What-is-Linking)；[linked files report](https://support.workiva.com/hc/en-us/articles/360036000671-Use-the-Linked-Files-Report) | 2023-06-28；2025-02-28 |
| 财务监管 | [SEC Inline XBRL](https://www.sec.gov/data-research/structured-data/inline-xbrl)；[XBRL validation/rendering](https://www.sec.gov/data-research/xbrl-validation-rendering) | 2016-06-14；2024-07-02，后者复核 2026-01-21 |
| Diff | [Word legal blackline](https://support.microsoft.com/en-us/word/compare-document-differences-using-the-legal-blackline-option)；[Adobe Compare](https://helpx.adobe.com/acrobat/using/compare-documents.html) | 未显示；2026-02-26 |
| Diff | [Apryse diffing](https://docs.apryse.com/web/guides/diffing)；[Nutrient comparison](https://www.nutrient.io/guides/web/comparison/compare-documents/) | 未显示/持续更新 |
| Diff 边界 | [Word browser differences](https://support.microsoft.com/en-us/word/differences-between-using-a-document-in-the-browser-and-in-word) | 未显示/持续更新 |
| 安全需求 | [Microsoft Protected View](https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653)；[Defender features](https://learn.microsoft.com/en-us/office365/servicedescriptions/microsoft-defender-for-office-365-features) | 未显示/持续更新 |
| air-gap | [Nutrient deployment](https://www.nutrient.io/sdk/deployment-options/)；[Azure disconnected containers](https://learn.microsoft.com/en-us/azure/ai-services/document-intelligence/containers/disconnected?view=doc-intel-4.0.0) | 未显示；2026-04-25 |
| 档案 | [DROID](https://www.nationalarchives.gov.uk/information-management/manage-information/preserving-digital-records/droid/)；[PRONOM](https://www.nationalarchives.gov.uk/help/pronom/overview.htm) | DROID 当前版 6.8.1；其余持续更新 |
| 档案 | [NARA framework](https://www.archives.gov/preservation/digital-preservation/risk)；[transfer guidance](https://www.archives.gov/records-mgmt/policy/transfer-guidance.html) | 复核 2025-06-20；复核 2024-05-24 |
| PDF/A 验证 | [veraPDF](https://verapdf.org/software/) | 未显示/持续更新 |
| 无障碍标准 | [WCAG 2.2](https://www.w3.org/TR/WCAG22/)；[U.S. Access Board](https://www.access-board.gov/ict/about/)；[document baseline](https://ictbaseline.access-board.gov/document-baselines/) | W3C REC 2024-12-12；baseline v1.0 2024-09-30 |
| 无障碍竞品 | [Microsoft Checker rules](https://support.microsoft.com/en-US/accessibility/office-accessibility/rules-for-the-accessibility-checker)；[Adobe accessibility](https://helpx.adobe.com/au/acrobat/using/create-verify-pdf-accessibility.html) | 未显示；2025-08-01 |
| 培训 | [iSpring Presenter](https://www.ispringsolutions.com/ispring-presenter)；[iSpring LMS pricing](https://www.ispring.com/pricing) | 未显示/当前页面 |
| 现场离线 | [IBM Maximo Mobile properties](https://www.ibm.com/docs/en/masv-and-l/maximo-manage/cd?topic=mobile-maximo-properties)；[Dynamics Field Service offline](https://learn.microsoft.com/en-us/dynamics365/field-service/mobile/set-up-offline-profile) | 未显示；2026-06-26 |
| 销售室 | [DocSend Performance](https://help.dropbox.com/share/dropbox-docsend-documents-performance-tab)；[Activity](https://help.dropbox.com/share/dropbox-docsend-documents-activity-tab) | 2026-02-09；2025-12-10 |
| 销售室 | [DocSend watermarks](https://help.dropbox.com/share/dropbox-docsend-watermarks)；[data-room analytics](https://help.dropbox.com/organize/dropbox-docsend-data-room-analytics) | 2026-02-16；2026-03-06 |
| 阅读时间边界 | [DocSend visitor time](https://help.dropbox.com/share/dropbox-docsend-visitor-time-on-document) | 2025-12-09 |
| 通用预览 | [Box Preview](https://developer.box.com/guides/embed/ui-elements/preview/)；[SharePoint Embedded preview](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/build/preview-files)；[Google Drive files](https://support.google.com/drive/answer/37603?hl=en) | 未显示；2026-07-27；未显示 |
| 免费 PDF | [Adobe PDF Embed](https://developer.adobe.com/document-services/apis/pdf-embed/)；[Mozilla PDF.js](https://mozilla.github.io/pdf.js/getting_started/) | 未显示/持续更新 |
| 浏览器 PDF | [Microsoft Edge PDF reader](https://learn.microsoft.com/en-us/deployedge/microsoft-edge-pdf) | 2024-07-18 |
| LMS 预览 | [Canvas Student Guide PDF](https://community.canvaslms.com/html/assets/Canvas_Student_Guide.pdf) | 2025-06-27 |

## 10. 最终判断

**准确的部分：** Viewer 的高价值确实在“真实文档内容 + 对象级定位 + 宿主业务/AI 结果”交汇处，而不是附件预览本身；AI 引用、复核、审批、数字血缘和未来 Diff 的确可以共享底层目标模型。

**需要修正的部分：** 通用抽取复核、Office Preflight、文档即业务界面和安全查看此前优先级偏高；离线并非独特；Diff 复杂度与竞争被低估；归档、无障碍、培训、销售室的主要价值都在 Viewer 之外。

因此，建议把产品表述收敛为：

> **浏览器本地、宿主中立的多格式可核验文档证据层：真实渲染原文件，用标准化 selector 把外部 AI、抽取与业务结果定位回原始对象，并显式报告映射与兼容性质量。**

这是一条有充分需求证据、但竞争激烈的方向。它是否能成为 DocViewKit 的业务，仍取决于下一步设计伙伴能否证明三件事：现有套件无法满足其格式/部署边界；原生 Office 对象定位比 PDF 页码/bbox 有可量化价值；客户愿意为这一差异而不是为完整 AI、工作流或编辑套件付费。
