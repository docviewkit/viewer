# flyfish-dev/file-viewer 源码研究与 OfficeViewer 借鉴判断

> 研究日期：2026-08-29
> 上游快照：[`815d52ee423ed503906a7fd4ccb69ad2304a2178`](https://github.com/flyfish-dev/file-viewer/commit/815d52ee423ed503906a7fd4ccb69ad2304a2178)
> 范围：只使用 `flyfish-dev/file-viewer` 仓库源码、清单、工作流、Issues、Releases 和提交历史等一手材料；不把官网营销文案或 npm 包描述单独当成实现证据。

## 结论先行

这个项目最有价值的不是把它的 84 个 npm 发布目标搬进 OfficeViewer，而是它把“格式注册、能力等级、已知限制、资产交付、依赖许可证、发布证据”做成了机器可检查的数据链。上游公开目录记录 244 个扩展名、34 条预览链路，其中 221 个标为 stable、23 个标为 experimental，并为每条链路给出 fidelity level、容器和 known limits；但这些等级是项目自己的声明，不等于我们要求的原生 Office 视觉认证。[生成格式目录](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/docs/generated/format-catalog.md#L1-L45)

对 OfficeViewer 最值得直接借鉴的四点：

1. 用一份生成式 format catalog 绑定扩展名、renderer、稳定度、证据和已知限制，而不是把支持表散落在 README。[生成格式目录](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/docs/generated/format-catalog.md#L1-L45)
2. 用 capability manifest 同时声明 renderer、格式、静态资产、许可证策略、体量等级和所属 profile，使“装了代码但漏发 Worker/WASM/font”成为可检测错误。[Archive capability manifest](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/archive/file-viewer.capability.json#L1-L22)
3. 重格式按 renderer/preset 懒加载，并为默认 profile 设置依赖闭包与静态资产预算；它解决的是交付闭包，不是文档语义复用。[架构说明](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/README.en.md#L219-L247) [profile budgets](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/ecosystem/profile-budgets.json#L1-L40)
4. 格式缺陷要求真实/脱敏文件、同视口视觉证据、受影响包与回滚说明，这与 OfficeViewer 的真实制品验收原则高度一致。[贡献证据契约](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/CONTRIBUTING.md#L18-L39) [PR 证据要求](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/CONTRIBUTING.md#L65-L94)

不能照搬的核心边界：

- 上游的统一主要发生在壳层 API、生命周期、toolbar、zoom、print/export 和扩展名路由；格式内容模型仍由多个独立引擎和第三方库各自输出 DOM、Canvas 或虚拟表格。它不是 OfficeViewer 项目规则所要求的“共享文档模型、协议与渲染语义”。[core/renderer 架构](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/README.en.md#L219-L234) [presentation 的双引擎边界](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/presentation/README.en.md#L1-L55)
- 上游的 iWork 可显式启用 `embeddedPreview: 'fallback'`，当结构化结果受限时把包内预览图放进最终内容；这直接违反 OfficeViewer 禁止用内嵌缩略图/预览图作最终渲染兜底的规则。[iWork fallback 开关](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/iwork/src/iwork.ts#L259-L285) [上游预览图查找](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/iwork/src/parser.ts#L195-L213)
- 上游把 ODT/ODP 归为 stable structured，但当前实现只从 `content.xml` 收集标题/段落文本并套固定纸张样式，不是布局渲染，不能据此降低 OfficeViewer 的 ODF 保真标准。[ODF 文本提取](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/word/src/openDocument.ts#L52-L99)
- 公开 CI 只安装并运行 Chromium；根 `pnpm test` 也只覆盖日语 i18n 与 thumbnail。仓库文档虽记录若干 Chromium/Firefox/WebKit 验收，但公开 checkout 没有一个由根脚本和 public CI 执行的完整三内核格式矩阵，因此不能把宣传的跨浏览器结论视为当前公开可复核的全矩阵证明。[根测试脚本](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/package.json#L27-L52) [Public CI](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/.github/workflows/public-ci.yml#L16-L87)

## 1. 架构

### 1.1 四层装配，而非统一排版内核

上游结构可概括为：框架组件 / Web Component → `@file-viewer/core` → preset 或 renderer → 本地 Worker、WASM、字体和 vendor assets。`core` 是框架无关 TypeScript；React、Vue、Svelte、jQuery 和 Web Component 各自维护原生生命周期，不互相嵌套。[README 架构图与职责](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/README.en.md#L219-L239)

`core` 中的 registry 维护 renderer ID 与扩展名的唯一映射，重复占用扩展名会抛错；dispatcher 再把 renderer ID 映射到 handler，并允许基于内容签名做一次 renderer redirect。这一设计能避免 suffix 路由散落，是值得复用的“装配层单一事实源”。[registry](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/core/src/registry/registry.ts#L52-L116) [dispatcher](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/core/src/rendering/dispatcher.ts#L11-L92)

但 renderer 只共享协议，不共享格式语义。例如 Office preset 只是把 PDF、Word、Spreadsheet、Presentation、OFD、iWork、WordPerfect、Hangul renderer 聚合成数组；它没有把这些格式映射到共同的页面、段落、表格、形状或绘制模型。[Office preset](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/presets/office/src/index.ts#L1-L31)

**对 OfficeViewer 的判断：** 可以借 registry/capability/preset 的装配思路；不能用“renderer 都实现同一个 handler”替代当前共享模型/协议/渲染语义要求。新增格式仍必须先证明哪些语义严格等价并进入共同层，哪些必须留在适配层。[OfficeViewer 项目规则](https://github.com/wjfree/OfficeViewer/blob/51c74efb99741362094333e9966ee33399a72976/AGENTS.md#L3-L8)

### 1.2 发布模块数很大，标准 profile 才是纠偏机制

上游提供 light/full 组件、多个框架版本、renderer、preset、capability、asset、CLI 和兼容别名，共称 84 个 npm targets。Full 兼容包固定保留旧的 221-extension/32-pipeline 合同，新项目则推荐 standard profile 和显式 renderer。[包选择与 Full 兼容边界](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/README.en.md#L117-L167)

预算文件为 standard preset + web 设置上限：最多 98 包、packed closure 30,408,704 bytes（约 30.4 MB 十进制）、unpacked closure 101,711,872 bytes（约 101.7 MB）、静态资产 24,117,248 bytes；它还明确禁止 PPT、CAD、DICOM、iWork、Typst 等重包进入 standard。[standard budget](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/ecosystem/profile-budgets.json#L3-L32)

同一文件记录的 v2.4.0 `@file-viewer/web-full` 实测闭包为 274 包、packed 310,928,143 bytes、unpacked 957,831,702 bytes、静态资产 448,964,222 bytes。也就是说模块化不是微优化，而是对近 1 GB 解包闭包的必要控制。[legacy Full reference](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/ecosystem/profile-budgets.json#L84-L103)

**对 OfficeViewer 的判断：** 借鉴“能力清单 + profile + 闭包预算 + 资产收据”，但不要复制 84 包的产品矩阵。OfficeViewer 当前单 SDK/可选 format pack 更适合保持少实体；只有消费者确实需要独立升级、独立下载或独立许可证边界时才拆包。

## 2. 格式支持的真实实现路径

| 格式族 | 上游真实路径 | 证据与边界 | 对 OfficeViewer |
| --- | --- | --- | --- |
| DOCX | `renderer-word` 动态导入 wrapper，实际依赖外部 npm `@file-viewer/docx@0.3.28`；该引擎源码不在本仓库 | [Word handler](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/word/src/index.ts#L45-L79) [package dependency](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/word/package.json#L60-L79) | 只能审计集成、Worker 选路和 DOM 后处理，不能从这个仓库独立验证 DOCX parser/layout 内核；不可把其 fidelity 声明当完整源码证据。 |
| Binary DOC | 仓库内有 CFB/FIB/CLX/FKP/OfficeArt parser，renderer 走 Worker/HTML | [DOC 源码目录与入口](https://github.com/flyfish-dev/file-viewer/tree/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/doc/src) [DOC/Binary 分流](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/word/src/index.ts#L52-L72) | 可参考内容签名纠错；不要复制 HTML-only 模型，优先映射现有共享文档模型。 |
| ODT/ODP | JSZip 解 `content.xml`，ODP 按 `draw:page` 收集 `text:p`，ODT 收集 `text:h/text:p`，再生成固定 DOM | [实现](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/word/src/openDocument.ts#L52-L99) | 属于可读文本降级，不是布局引擎；不适合 OfficeViewer 的 ODF 主路径。 |
| XLSX/XLS/ODS/CSV | `styled-exceljs` + `e-virt-table` + Worker；大于 1 MB 可自动走静态 Worker，显示层采用窗口化/虚拟表格 | [依赖与 Worker 构建](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/spreadsheet/package.json#L74-L97) [1 MB threshold](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/spreadsheet/src/spreadsheet.ts#L108-L120) | 虚拟化值得借鉴；parser 和 display model 是独立栈，不能替代 OfficeViewer 对跨格式共享 DrawingML/样式语义的要求。 |
| PPTX | 仓库内 `@file-viewer/pptx`，renderer 用 Worker 渐进处理；初始 3 页、每批 4 页，并能为单页失败插入错误页 | [PPTX 资源限制与批次](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/pptx/src/viewer.ts#L53-L102) | “页级失败隔离 + 渐进批次”符合整体可用优先，值得映射到 OfficeViewer 的诊断/增量 display-list，而不是复制 DOM renderer。 |
| Binary PPT | 外部 `@file-viewer/ppt@0.3.3`，Worker/OffscreenCanvas/WASM/CJK font，有界 IndexedDB 帧缓存；与 PPTX 严格分路 | [双引擎与回退](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/presentation/README.en.md#L48-L55) | 保持二进制与 OOXML parser 边界是正确的；共享只能发生在可证明等价的 presentation model/render semantics。 |
| PDF | renderer 自述 powered by PDF.js，构建时 stage `pdfjs-dist@5.4.624` runtime | [PDF package](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/pdf/package.json#L1-L88) | 可参考资产探测与自托管，不应把 PDF.js 的 DOM/canvas 模型引入 Office 格式共同层。 |
| iWork | 自有 IWA/Snappy 解析 + Worker，依赖 JSZip、keynote-archives、styled-exceljs；可选包内 preview fallback | [iWork dependencies](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/iwork/package.json#L64-L82) [fallback](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/iwork/src/iwork.ts#L259-L285) | 解析限额可借；最终 preview fallback 不可用。 |

上游的“244 extensions”不能等同“244 种高保真文档格式”：目录本身区分 high-fidelity / structured / basic，且许多扩展只是同容器别名。例如 DOC/DOT 是 structured，ODT/ODP/RTF 共用 structured Open Document 路径，DICOM 是 experimental/basic。[格式目录](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/docs/generated/format-catalog.md#L10-L45)

## 3. 依赖与运行时模型

### 3.1 浏览器运行，不代表没有重依赖

项目强调无需强制服务端转换，文件在允许时于浏览器内解析，Worker/WASM/font/vendor assets 可全部自托管；重链路按格式加载。[运行原则](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/README.en.md#L41-L50)

实际 renderer 依赖覆盖 PDF.js、JSZip、libarchive.js、Three.js、OpenCascade、sql.js、MapLibre、Cornerstone、DOMPurify、Mermaid、styled-exceljs、e-virt-table 等。它通过拆 renderer 限制默认闭包，而不是消除 Node/npm 模块。[renderer 目录](https://github.com/flyfish-dev/file-viewer/tree/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers) [standard 禁入重包](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/ecosystem/profile-budgets.json#L3-L32)

Full 包必须完整复制同版本 Worker/WASM/font/vendor asset tree；只复制 IIFE 入口会造成不完整部署。Vite plugin 自动复制，其他 bundler 通过同版本 copy-assets CLI。[Full 资产交付](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/README.en.md#L127-L139)

**对 OfficeViewer 的判断：** 我们刚清理 Node 依赖后不应反向追求上游的多引擎依赖广度。值得加的是“每个 format pack 的代码/WASM/font/worker 闭包预算与发布收据”，不是为每个格式再加一个 npm 生态。

### 3.2 依赖所有权并不完全公开

`@file-viewer/docx@0.3.28` 和 `@file-viewer/ppt@0.3.3` 是独立发布依赖，源码不在当前仓库对应目录；前者只能审计 wrapper，后者还保留独立许可证、可见 watermark，并声明去 watermark 需要商业授权。[DOCX dependency](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/word/package.json#L60-L79) [PPT license boundary](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/presentation/README.en.md#L48-L55)

因此“仓库 Apache-2.0”只说明仓库主体，不代表所有运行时资产都可按 Apache-2.0 处理。任何借用都必须逐包核查依赖许可证、NOTICE、字体和 vendor asset，而不能只看根 LICENSE。[根包许可证](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/package.json#L54-L61) [Apache-2.0 LICENSE](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/LICENSE#L1-L20)

## 4. 渲染、转换与降级策略

### 4.1 没有服务端统一转 PDF

上游主路径是浏览器端直接解析：不同 renderer 输出 DOM、Canvas、图片、WebGL 或虚拟表格，不经过统一中间文档转换。[浏览器原生目标](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/README.en.md#L7-L11) [renderer 架构](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/README.en.md#L219-L234)

这避免了隐私文件上传和转换服务运维，但也造成每个格式的渲染语义、故障模型和 fidelity ceiling 不同。上游自己明确提示 fidelity 会受文件结构、嵌入字体、厂商扩展和浏览器能力影响。[Honest Boundaries](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/README.en.md#L189-L197)

### 4.2 局部失败隔离做得较好的路径

PPTX 为单页解析失败构造带 slide index 的错误卡，而不是让整份演示文稿无输出；同时以首批 3 页、后续每批 4 页渐进挂载。[PPTX page error 与 batch](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/pptx/src/viewer.ts#L70-L102)

Word 的外部链接和 HTTP(S) 图片默认阻断，但嵌入图片和内部书签继续工作，体现“隔离危险资源而保留文档其余内容”。[Word 默认外部资源策略](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/word/README.en.md#L55-L59)

**对 OfficeViewer 的判断：** 借鉴它的“诊断对象 + 局部占位 + 继续渲染”产品反馈，但诊断必须进入现有协议和共同回归，不能只在每个 renderer 写一套错误 DOM。

### 4.3 与 OfficeViewer 规则冲突的降级

iWork renderer 在 `limitedPreview` 且配置 `embeddedPreview: 'fallback'` 时显示包内 `preview.jpg` / QuickLook Preview/Thumbnail；即使是显式 opt-in，这仍可能成为最终内容。[fallback 判定](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/iwork/src/iwork.ts#L259-L285) [preview 候选](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/iwork/src/parser.ts#L195-L213)

OfficeViewer 明确只允许这类图片用于缩略图导航或加载占位，禁止成为最终渲染或解析失败兜底。[OfficeViewer 项目规则](https://github.com/wjfree/OfficeViewer/blob/51c74efb99741362094333e9966ee33399a72976/AGENTS.md#L3-L5)

同时，本项目自己的 `docs/SUPPORT.md` 仍写有 Pages root-preview fallback 和 Keynote per-slide embedded preview fallback，与新项目规则存在潜在内部不一致；这是本轮只读研究发现的待审计项，不应只批评上游。[OfficeViewer SUPPORT 当前声明](https://github.com/wjfree/OfficeViewer/blob/51c74efb99741362094333e9966ee33399a72976/docs/SUPPORT.md#L32-L41)

## 5. 安全与资源限制

### 5.1 值得借鉴的具体防线

- DOC/DOCX/RTF 外部 link/resource 默认 block；只有显式 allow 才启用，未知协议和 protocol-relative URL 仍拒绝。[Word policy](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/word/README.en.md#L55-L59)
- 富文本、PPTX 等生成 DOM 的边界使用 sanitization；PPTX 在插入 HTML 前调用 sanitizer，并对单页错误文本做 HTML escape。[PPTX sanitizer 边界](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/pptx/src/viewer.ts#L70-L89)
- PPTX 默认最大源文件 160 MiB；超限在 preflight 阶段以结构化诊断失败。[PPTX limit](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/pptx/src/options.ts#L1-L5) [preflight](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/pptx/src/viewer.ts#L53-L67)
- iWork 限制累计解压 256 MiB、压缩比 200、对象数 250,000、单图 80M pixels、嵌套深度 128，并在 Snappy frame 累计输出超限时中止。[iWork limits](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/iwork/src/limits.ts#L1-L9) [Snappy enforcement](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/iwork/src/snappy.ts#L74-L93)
- TIFF 限制 32 MiB 源文件、64 页、16,384 边长、单页 32M pixels、累计 128M pixels。[TIFF limits](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/image/src/tiff.ts#L22-L26) [TIFF enforcement](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/image/src/tiff.ts#L82-L114)
- capability manifest 带 SPDX/license policy，CI 另有 renderer dependency audit 和 DICOM license ledger gate。[capability license field](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/archive/file-viewer.capability.json#L18-L22) [root audit scripts](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/package.json#L48-L52)

### 5.2 安全成熟度判断

安全意识明显高于普通 viewer：有私密漏洞报告渠道、默认外部资源阻断、DOM sanitizer、压缩/像素/对象限制、Worker 清理和许可证账本。[SECURITY](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/SECURITY.md#L1-L34)

但不能据此推断“无安全风险”。仓库维护记录明确列出过 PPTX stored XSS（GHSA-mrrf-4m5h-mppj），也说明生成 DOM 的格式引擎始终是高风险面。[上游安全公告](https://github.com/flyfish-dev/file-viewer/security/advisories/GHSA-mrrf-4m5h-mppj)

**对 OfficeViewer 的判断：** 最应借的是统一的 per-format resource budget schema 和 hostile-fixture 回归，而不是复制具体数值。数值必须依据 OfficeViewer 的 Rust/Wasm 内存、浏览器实测和格式语义校准。

## 6. 浏览器兼容性

上游 README 声称部分 WordPerfect/Hangul、iWork、表格、TIFF 等路径有 Chromium/Firefox/WebKit smoke，并在维护文档中记录若干三内核验收。[能力声明](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/README.en.md#L179-L183)

可公开复核的根 CI 边界更窄：Public CI 只安装 Chromium，并用 Chromium 跑 toolbar、sanitization、demo smoke 和 PPTX slideshow；没有安装 Firefox 或 WebKit。[Public CI browser steps](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/.github/workflows/public-ci.yml#L66-L87)

根 `verify:browser-smoke` 脚本也硬编码 `chromium.launch()`；它验证 demo UI、移动布局、搜索和 locale 等，但不是格式×三内核矩阵。[public smoke browser](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/apps/viewer-demo/scripts/public-smoke.mjs#L40-L45) [Chromium launch](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/apps/viewer-demo/scripts/public-smoke.mjs#L93-L110)

贡献文档列出的 `verify:format-support`、`verify:offline-assets` 等命令并不在公开根 `package.json` scripts 中；完整 release-channel/migration gates 也被标为 maintainer gates。因此，本研究只能确认“源码和历史记录声明有三内核证据”，不能确认当前公开 checkout 能独立重跑完整三内核格式矩阵。[贡献命令与 maintainer gates](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/CONTRIBUTING.md#L41-L63) [公开根 scripts](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/package.json#L8-L52)

**对 OfficeViewer 的判断：** 不降低现有 Chrome/Firefox/Safari 三内核目标；仍以本项目公开可运行的 browser matrix、真实原文件和 console-clean 结果为准。上游可借的是问题证据模板，不是验证覆盖结论。

## 7. 构建、发布与许可证

工作区使用 pnpm monorepo，构建顺序为 core → renderers → capabilities → presets → components → tools → thumbnail → examples → demo；公开 CI 固定 pnpm 11.0.9、Node 24、Rust wasm32 target 与 wasm-bindgen-cli 0.2.127，之后执行 build、type-check、narrow test、docs、Chromium smoke。[root build scripts](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/package.json#L6-L37) [Public CI toolchain](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/.github/workflows/public-ci.yml#L16-L64)

v3.0.0 Release 声明 84 个 npm targets、91 份冻结发布 bytes、93 个 GitHub Release assets，并发布 manifest/matrix/status/schema 等机器可读附件；这体现了很强的“发布物独立验真”意识。[v3.0.0 release](https://github.com/flyfish-dev/file-viewer/releases/tag/v3.0.0)

仓库主体 Apache-2.0，可商用、修改和分发并带专利许可；但第三方依赖、字体、WASM/vendor runtime 各有自己的许可证，尤其 binary PPT runtime 明确不受 renderer Apache-2.0 覆盖。[根 LICENSE](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/LICENSE#L1-L20) [PPT 单独许可证](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/presentation/README.en.md#L48-L55)

**对 OfficeViewer 的判断：** 借鉴 release manifest/status/schema 与逐资产 hash/size/readback；继续保持每个第三方 codec/font/WASM 的 NOTICE 和许可证边界，不因上游根许可证宽松而整体复制。

## 8. 测试成熟度

### 优点

- 贡献规则要求真实或 license-safe fixture、focused parser/browser assertion、视觉证据、离线资产检查，方向正确。[change-specific expectations](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/CONTRIBUTING.md#L86-L94)
- 仓库包含 issue-numbered 回归脚本和真实 fixture，例如 spreadsheet #178、DOC 表格、iWork 版本矩阵、signature 正负样本；格式目录公开已知限制而不是只列后缀。[spreadsheet scripts](https://github.com/flyfish-dev/file-viewer/tree/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/spreadsheet/scripts) [iWork fixtures](https://github.com/flyfish-dev/file-viewer/tree/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/iwork/test/fixtures) [format catalog](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/docs/generated/format-catalog.md#L10-L45)
- 发布记录把 parser、浏览器、冷安装、GitHub Release、生产 Demo 分成不同证据通道，避免把单一 CI 绿灯等同全部发布成功。[v3.0.0 release](https://github.com/flyfish-dev/file-viewer/releases/tag/v3.0.0)

### 不足

- 根 `pnpm test` 只执行 i18n 和 thumbnail；大量 renderer 测试被放在各包 build/verify 脚本或未公开的 maintainer gate，默认测试命令的语义过弱。[root test](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/package.json#L33-L37)
- Public CI 只覆盖 Chromium；三内核证据分散在 README、maintenance ledger 和专项脚本声明中，无法由公开根工作流一次复核。[Public CI](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/.github/workflows/public-ci.yml#L66-L87)
- `@file-viewer/docx` 与 `@file-viewer/ppt` 的关键内核不在本仓库，公开源码审计和版本回归证据存在 dependency-owner 断点。[Word manifest](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/word/package.json#L60-L79) [presentation boundary](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/presentation/README.en.md#L48-L59)
- “stable” 的含义不统一：ODP/ODT 即使只是文本提取仍标 stable/structured；因此 status 不能替代 fidelity 层级和真实原生参考比较。[catalog ODF row](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/docs/generated/format-catalog.md#L12-L20) [ODF implementation](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/packages/renderers/word/src/openDocument.ts#L52-L99)

综合判断：工程治理和发布治理成熟度较高，格式实现成熟度高度不均，公开可复核测试边界低于 release/README 声明的全量边界。

## 9. 对 OfficeViewer 的行动建议

| 优先级 | 建议 | 原因 | 是否照搬上游 |
| --- | --- | --- | --- |
| P0 | 审计并消除 Pages root-preview 与 Keynote per-slide embedded preview 最终兜底，或明确它们只可作为 loading/thumbnail | 当前本项目文档与新规则潜在冲突 | 否；这是修正自身边界 |
| P1 | 从现有格式注册/测试资料生成单一 capability catalog：格式、容器、fidelity、已知限制、parser pack、WASM/Worker/font assets、诊断、浏览器证据 | 上游最强且可低风险借鉴的治理能力 | 借数据模型，不复制 244 项营销口径 |
| P1 | 给每个 format pack 建 asset manifest 与闭包预算，并在 release gate 校验实际 dist、hash、MIME、大小和缺失资产 | 防止 runtime 代码存在但 Worker/WASM/font 漏发 | 借机制 |
| P1 | 把所有局部失败统一进现有协议诊断：format、unit/page/object、severity、fallback、source range；保留剩余内容 | 符合“局部失败隔离、整体可用优先” | 借 PPTX 页级隔离思想，不复制错误 DOM |
| P1 | 保持公开、可一键执行的 Chromium/Firefox/WebKit matrix；format catalog 的 stable 状态必须引用可重跑证据 | 避免上游“历史说三内核、公开 CI 只跑 Chromium”的证据断层 | 不照搬其 CI 边界 |
| P2 | 评估 `standard` / optional pack 的下载与解包预算，但先用当前单 SDK + format packs，不拆 84 包 | 体量可控，同时保持少实体 | 只借预算，不借包数量 |
| P2 | 对每次新增/修改格式运行跨格式重复审查，并在 catalog 记录“共享语义 / 适配层差异” | 落实 OfficeViewer 严格等价才共享的规则 | 上游 renderer 分包不能替代此项 |

### 明确不做

- 不引入完整 flyfish renderer/preset/component 依赖树。
- 不以 ODT/ODP 文本提取、embedded preview fallback 或 extension routing 作为“格式已支持”的证明。
- 不因为上游使用 DOM/Canvas/第三方引擎，就绕过 OfficeViewer 共享文档模型与 display-list。
- 不把上游 stable、release note 或 README 的三内核声明当成 OfficeViewer 的验收证据。
- 不把 Apache-2.0 根许可证推导成所有上游 runtime/font/WASM 都可直接复制。

## 10. 最终评价

`flyfish-dev/file-viewer` 是一个产品覆盖面和交付治理很强的浏览器文件查看器：模块化装配、离线资产、自托管 Worker/WASM、格式目录、资源限制、真实 fixture 和发布矩阵都值得学习。[项目定位](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/README.en.md#L41-L50) [公开工作区](https://github.com/flyfish-dev/file-viewer/blob/815d52ee423ed503906a7fd4ccb69ad2304a2178/README.en.md#L237-L260)

它不是 OfficeViewer 的可替代内核：格式间共享停在壳层协议，Office fidelity 高低不一，部分关键引擎在仓库外，ODF 路径偏文本提取，iWork 允许最终 preview fallback，公开 CI 也没有完整三内核矩阵。最优策略是吸收其机器可验证的产品治理与资源安全方法，同时坚持 OfficeViewer 更严格的真实结构解析、共享语义、局部降级诊断和三内核原文件验收边界。
