# Apple iWork 原生格式兼容性调研

> 实施状态（2026-07-18）：现代单文件 iWork 包已经落地为默认关闭、按需加载的独立 Wasm；Pages/Numbers 显示根预览，Keynote 按原生 slide tree 解析静态背景、基础段落样式文本、嵌套组和原始 JPEG/PNG，并对未覆盖的页面回退到包内逐页预览。当前产品路线不包含转换适配器，实际支持边界以 [SUPPORT.md](../SUPPORT.md) 为准。

调研日期：2026-07-15

## 结论

**可以做，但要区分“能显示预览”“能读取部分语义”和“完整兼容”。**

- `.pages`、`.numbers`、`.key` 的首屏静态预览可以较快支持；这不是对象级兼容。
- 在浏览器本地解析现代 iWork 二进制格式、生成可命中对象并保留原生源映射，在技术上可行；最合适的方向是复用 OfficeViewer 现有 Rust/Wasm 安全边界，自建受限 IWA 读取器，并把 libetonyek 仅作为协议/解析研究的参考实现和差分比较器。iWork 的 native/release 验收金标只能来自对应 Apple iWork 应用。
- 不建议把 libetonyek 直接移植进默认浏览器核心：上游没有 Wasm 构建，依赖链和公共回调接口都与本项目的依赖极简、精确源映射目标不匹配。
- Apple 没有公开现代 iWork 完整格式规范。`iWorkFileFormat` 和 libetonyek 都建立在逆向分析之上，因此“当前常见文件的只读子集”可以工程化，“所有 iWork 版本高保真兼容”不能做成稳定承诺。
- 如业务需要尽快扩大可打开范围，可以另做显式选择的 LibreOffice/macOS 转换适配器；它会改变浏览器本地、无上传、原生源映射的产品边界，不应伪装成 OfficeViewer 核心原生支持。

建议的产品承诺是：**分阶段提供现代单文件 iWork 的只读兼容；先预览、再 Pages 基础语义、再 Numbers 和 Keynote；加密文件与 macOS package 目录先明确拒绝。**

## 能做到什么

| 层级 | 用户结果 | 可行性 | 是否符合当前核心边界 |
| --- | --- | --- | --- |
| 预览回退 | 显示文件内嵌的单张 `preview.jpg`；无搜索、选择、命中和对象检查 | 高 | 基本符合，但必须标为 `preview-only` |
| 转换后查看 | Apple/LibreOffice 转成 DOCX、XLSX、PPTX 或 PDF | 高 | 外部转换不符合；用户自行导出后打开符合 |
| 原生语义子集 | 浏览器本地读取 IWA，生成 Pages 段落、Numbers 表格、Keynote 页面对象 | 中 | 符合，推荐方向 |
| 当前版本高保真 | 覆盖复杂排版、多个 Numbers 自由布局表、图表、母版、动画等 | 中低，需长期投入 | 可以逐项建设，不能一次承诺 |
| 所有历史/未来版本完整兼容 | 任意 iWork 版本等同 Apple 原应用 | 不可作为可靠承诺 | 不符合可验证的支持策略 |

这里沿用本项目自己的定义：只有生成真实渲染对象、可参与命中测试并保留原生源引用的功能，才能称为 `Supported`；只显示内嵌图片或先转成另一种格式都不能升级成原生支持，见 [现有支持政策](../SUPPORT.md) 和 [架构边界](../ARCHITECTURE.md)。

## 1. 现代 iWork 文件是什么

### 1.1 单文件与 package 目录都存在

Apple 官方说明：Pages、Numbers、Keynote 默认保存为单文件，但可切换为 package；package 是 macOS 视作一个文件的一组文件，且 Apple 明确建议需要经浏览器上传到第三方云时保留单文件形式。[Apple：Save a document as a package or a single file](https://support.apple.com/en-us/119883)

这对 OfficeViewer 有直接影响：当前入口是 `OfficeEngine.open(bytes)`，一次只接收一个字节源，并不接收目录树。第一阶段应只承诺默认的单文件形式；package 目录应提示用户在 iWork 中改存为 Single File，而不是悄悄遗漏媒体文件。

### 1.2 IWA 是逆向得到的二进制对象图，不是公开标准

原始 `iWorkFileFormat` 项目把 iWork '13 描述为 bundle：媒体在 `Data/`，轻量元数据在 `Metadata/`，序列化对象位于 `Index.zip` 内的多个 `Index/*.iwa`，顶层还可能有预览图片。每个 IWA 是 Snappy framing 包裹的连续 Protobuf 对象；对象头包含文档内唯一 identifier、消息 type、version、length 和对象/数据引用。[格式说明](https://github.com/obriensp/iWorkFileFormat/blob/8575e441beaaaa56f480fdd91721f5bb06d07d43/Docs/index.md)

Protobuf 本身不自描述；要理解 payload，解析器必须预先知道消息 schema 和 type registry。原项目明确说明，这些映射和 `.proto` 是从 iWork 应用二进制及运行时 registry 中恢复的；项目目标也仅是解释 iWork '13，仓库只有两个提交，最后一次提交在 2013 年。[项目主页与许可证](https://github.com/obriensp/iWorkFileFormat/tree/8575e441beaaaa56f480fdd91721f5bb06d07d43)

因此，容器层相对清楚，真正困难的是不断演进的对象 schema、继承关系和布局语义。Apple 当前支持文档给出打开/导出能力，但没有给出现代 `.pages/.numbers/.key` 的公开格式规范。这是本项目未来需要持续维护兼容层的根本原因。

### 1.3 本机当前样本验证

补充验证使用本机 Pages、Numbers、Keynote 14.5 各生成一个空白文件：三者都是 stored ZIP，均包含 `Index/*.iwa`、`Metadata/Properties.plist`、`preview.jpg`、`preview-web.jpg` 和 `preview-micro.jpg`，`fileFormatVersion` 均为 `14.4.1`。这说明 2013 年文档描述的总体容器/IWA 架构仍能在 14.5 样本中观察到，但这只是三个样本，不能替代版本语料。

本机 Numbers 14.5 样本可被 `numbers-parser 4.18.5` 读取到 `Sheet 1`；Keynote 14.5 样本可被 `keynote-parser 1.14.4.0` 解包，但该工具同时警告仅支持 14.4、当前版本可能不兼容。该结果不作为第三方实现背书，只用来说明 schema 漂移不是理论风险，而是需要按 iWork 版本持续验证的维护成本。

## 2. Apple 官方兼容与转换能力

Apple 当前官方矩阵是：

| 应用 | 原生格式可官方导出为 | 可打开的主要外部格式 |
| --- | --- | --- |
| Pages | DOCX、PDF、EPUB、RTF/RTFD、TXT、图片、Pages '09 | 所有 Pages 版本、DOCX、RTF/RTFD、TXT |
| Numbers | XLSX、PDF、CSV、TSV、Numbers '09 | 所有 Numbers 版本、XLSX、CSV、分隔/定宽文本 |
| Keynote | PPTX、PDF、图片、GIF、MOV/M4V、Keynote '09、HTML | 所有 Keynote 版本、PPTX |

完整列表和平台限制见 [Apple 官方转换与文件格式兼容表](https://support.apple.com/en-gb/105050)。其中没有 ODT、ODS、ODP；当前可依赖的标准 Office 交换格式是 DOCX、XLSX、PPTX。

官方支持两种转换入口：

1. 在 iPhone、iPad、Mac 的对应应用中“Export and Send”；
2. 登录 iCloud.com，在文档管理器中“Download a Copy”。

Apple 文档描述的是用户交互流程，不是公开的服务端转换 API。因此：

- 用户自行导出后由 OfficeViewer 打开，是立即可用且边界清楚的兼容路径；
- 依赖 iCloud 私有网络接口做自动转换不可作为产品架构；
- macOS 上可以围绕已安装的 iWork 应用做 host adapter，本机应用字典也确认 Pages 可导出 DOCX/PDF/RTF、Numbers 可导出 XLSX/PDF/CSV、Keynote 可导出 PPTX/PDF/HTML/图片，但它是 macOS 专属外部能力，不是浏览器本地跨平台核心。

格式转换也不是无损保证。Apple 明确提示，打开外部文件时可能报告缺失字体或外观变化；不同目标格式能力不同。[Apple 官方转换说明](https://support.apple.com/en-gb/105050)

## 3. libetonyek 能提供什么

### 3.1 支持范围必须按“声明、源码、测试”三层理解

libetonyek 是 LibreOffice 官方基础库，用于读取和转换 Keynote、Pages、Numbers。其 README 的公开一句话仍写着 Keynote 2–6、Pages 1–4、Numbers 1–2，但同一仓库的后续源码和发布记录已经包含现代二进制族：

- `KEY6Parser`：Keynote 6+；
- `PAG5Parser`：Pages 5+；
- `NUM3Parser`：Numbers 3+；
- `NEWS` 0.1.8/0.1.10 记录了 Pages 5+、Numbers 3+、Keynote 6+ 的文本、表格、公式、形状、图片和样式等改进。

直接证据见 [README](https://git.libreoffice.org/libetonyek/+/37704aa6ac808fe7f7a14b4515503c3de3bc0dbf/README.md)、[NEWS](https://git.libreoffice.org/libetonyek/+/37704aa6ac808fe7f7a14b4515503c3de3bc0dbf/NEWS)、[二进制格式分派源码](https://git.libreoffice.org/libetonyek/+/37704aa6ac808fe7f7a14b4515503c3de3bc0dbf/src/lib/EtonyekDocument.cpp)。

但上游检测测试固定的是 Keynote 6、Numbers 3、Pages 5 样本，主要断言 `isSupported()` 的类型和 confidence，并没有对当前 iWork 14/15 做完整渲染或视觉验收。[检测测试](https://git.libreoffice.org/libetonyek/+/37704aa6ac808fe7f7a14b4515503c3de3bc0dbf/src/test/EtonyekDocumentTest.cpp) 因而 `6+ / 3+ / 5+` 应理解为同一 IWA 格式族的 best-effort 解析器，不能直接翻译成“已验证所有后续版本”。

上游自己列出的已规划而未完整支持项包括 transitions、animations、presentation text auto-fit、connectors、charts 等。[FEATURES](https://git.libreoffice.org/libetonyek/+/37704aa6ac808fe7f7a14b4515503c3de3bc0dbf/FEATURES.md)

### 3.2 许可证和依赖

- libetonyek：MPL 2.0+；直接修改其 MPL 文件并分发时要遵守文件级源码义务。[许可证声明](https://git.libreoffice.org/libetonyek/+/37704aa6ac808fe7f7a14b4515503c3de3bc0dbf/README.md)
- 原始 iWorkFileFormat 文档、检查器和恢复的 schema：MIT。[LICENSE](https://github.com/obriensp/iWorkFileFormat/blob/8575e441beaaaa56f480fdd91721f5bb06d07d43/LICENSE)
- libetonyek 构建依赖 Boost、GLM、liblangtag、librevenge、libxml2、mdds、zlib；转换工具还需要 librevenge-generators 和 librevenge-stream。[README 构建说明](https://git.libreoffice.org/libetonyek/+/37704aa6ac808fe7f7a14b4515503c3de3bc0dbf/README.md)、[configure.ac](https://git.libreoffice.org/libetonyek/+/37704aa6ac808fe7f7a14b4515503c3de3bc0dbf/configure.ac)

MPL 并不阻止商业使用，但会增加发布合规和第三方声明工作；以上不是法律意见。

还需要单独看待 Apple 派生协议资料的权利边界。`iWorkFileFormat` 和 `keynote-parser` 明确说明其 Protobuf schema/type registry 来自对 iWork 应用二进制或运行时的恢复；这些项目对自身代码使用 MIT/MPL，并不自动回答 Apple 派生描述文件、应用许可条款及不同司法辖区互操作性逆向规则的问题。商用前应由法务确认可采用的资料和实现方式；在结论明确前，阶段 0 只做隔离研究，不把从 Apple 应用提取的生成文件直接并入产品仓库。可优先评估基于公开行为、授权样本和独立实现的 clean-room 方案。以上同样不是法律意见。

### 3.3 输出接口不满足 OfficeViewer 的源映射要求

libetonyek 的公共 API 只有格式识别，以及向 `RVNGPresentationInterface`、`RVNGSpreadsheetInterface`、`RVNGTextInterface` 发回调的三种 `parse()`。它没有公开 IWA component、ArchiveInfo identifier、message type 或 field path。[公共头文件](https://git.libreoffice.org/libetonyek/+/37704aa6ac808fe7f7a14b4515503c3de3bc0dbf/inc/libetonyek/EtonyekDocument.h)

这意味着直接消费公共回调，可以拿到转换后的文档语义，却拿不到本项目需要的稳定原生 source locator。若为此侵入 libetonyek 内部解析器，既依赖非公共实现，又需要长期维护 MPL fork。它适合做协议参考、解析/转换差分研究或外部转换器，不是默认核心的干净适配层，也不能作为 iWork 的 native/release 验收 oracle。

## 4. 浏览器/Wasm、服务端和 macOS 三条路径

### 4.1 直接把 libetonyek 编译到 Wasm

**理论可移植，当前没有上游现成路径，风险较高。**

仓库中没有 Emscripten/WebAssembly 构建文件或条件代码；上游是 C++11 + Autotools，并要求前述完整依赖链。库本身通过 `librevenge::RVNGInputStream` 抽象输入，不必调用 Apple 私有 API，因此不存在原则性的平台封锁；但要得到可发布 Wasm，仍需移植和裁剪整个依赖图、编写 JS ABI、控制异常/内存、接入 Worker 取消与预算，并另补源映射。

这与当前 [依赖极简 Rust core](../../Cargo.toml) 和“不引入完整第三方文档渲染器”的 [架构选择](../ARCHITECTURE.md) 相冲突。即使编译成功，也只证明能跑，不证明能达到对象协议、安全预算和准确性门槛。

结论：可以安排一次隔离的 2–4 周技术 spike 测体积、峰值内存和现代语料通过率；不建议在没有门禁结果前把它定为产品路线。

### 4.2 LibreOffice 服务端转换

**最快得到较宽格式覆盖，但会改变产品。**

LibreOffice 的 Pages import filter 直接调用 `EtonyekDocument::parse()`；Numbers 和 Keynote 也有对应 import filter。[Pages](https://git.libreoffice.org/core/+/master/writerperfect/source/writer/PagesImportFilter.cxx)、[Numbers](https://git.libreoffice.org/core/+/master/writerperfect/source/calc/NumbersImportFilter.cxx)、[Keynote](https://git.libreoffice.org/core/+/master/writerperfect/source/impress/KeynoteImportFilter.cxx)。LibreOffice 官方支持 `--headless --convert-to` 批量转换。[命令行文档](https://help.libreoffice.org/latest/en-US/text/shared/guide/start_parameters.html)

可部署独立受限进程，把 iWork 转成 PDF 或 OOXML/ODF 后交给现有 Viewer。但代价是：

- 文档离开浏览器并进入服务端/本机 native 进程；
- 新增上传、排队、存储、清理、租户隔离和 parser sandbox 责任；
- 原生 IWA source mapping 丢失，最多只能映射到生成文件；
- 转换保真仍受 libetonyek/LibreOffice 支持范围限制。

因此它适合企业显式启用的 `external conversion adapter`，不应成为默认 `OfficeEngine.open(bytes)` 的隐藏后端。

### 4.3 Apple Quick Look 与 macOS host adapter

Apple Quick Look 官方说明可预览包括 iWork 和 Microsoft Office 在内的常见文档。[Quick Look framework](https://developer.apple.com/documentation/quicklook) 这可用于未来 macOS/iOS 原生宿主中的系统预览，但它不是浏览器 Web API，也不返回 OfficeViewer 的对象树、文本范围或源 locator。

macOS host adapter 还可以调用安装的 iWork 应用导出标准格式。它的优点是 Apple 自己负责当前格式解析，缺点是平台专属、依赖应用安装/版本/用户会话，而且外部导出同样会丢失原生 IWA 映射。可作为桌面集成能力，不应污染跨平台核心协议。

### 4.4 文件内嵌 preview 回退

原始格式说明列出顶层 `preview.jpg`、`preview-web.jpg`、`preview-micro.jpg`。[iWorkFileFormat 格式说明](https://github.com/obriensp/iWorkFileFormat/blob/8575e441beaaaa56f480fdd91721f5bb06d07d43/Docs/index.md) 本机三个 14.5 空白样本也都有这些文件，但每个 `preview.jpg` 只有一张，尺寸分别为 Pages 724×1024、Numbers 720×552、Keynote 1024×576。

所以可以本地提取它们作为首屏/缩略图回退，不能把它描述为完整文档预览：它不覆盖 Pages 后续页、Numbers 其他表/画布区域、Keynote 后续幻灯片，也没有可选择文本、搜索、命中和对象源映射。预览缺失、过期或尺寸异常时还要结构化拒绝，不能回退为猜测解析。

## 5. 与 OfficeViewer 当前设计的匹配度

### 5.1 匹配的部分

- 单文件 iWork 总体仍是受限 ZIP/嵌套 `Index.zip` + 二进制 component，可复用现有 ZIP、CRC、路径校验和 Worker 生命周期。
- ArchiveInfo identifier 在文档内唯一，适合作为对象生命周期内的原生 identity；source locator 可设计为 `component + archiveId + messageType + fieldPath`，表格再增加 table/row/column，文本增加 range。
- 现有 display list、文字排版、图片解码、页面/幻灯片/表格对象和诊断模型都能复用。
- 当前每文档一个 Worker、超时终止、总对象/图片/输入预算的结构，适合隔离新的二进制解析器，见 [安全模型](../SECURITY.md)。

### 5.2 不匹配的部分

- 当前 Rust core 只引入经过边界控制的小型 Rust 编解码依赖；libetonyek 是较大的 C++ 依赖图。
- 当前解析器强调原生、格式区分的 source reference；libetonyek/libreoffice 转换输出丢失 IWA provenance。
- 当前包路径只处理一层 Office ZIP 语义；iWork 单文件内还可能有 `Index.zip`，所有嵌套展开必须共享预算，不能重置限制。
- Numbers 不是“一张 sheet 等于一个从 A1 开始的单网格”。Apple 官方允许在一张 sheet 上添加任意多个可移动、可调整尺寸的表。[Apple Numbers：Add or delete a table](https://support.apple.com/guide/numbers/add-or-delete-a-table-tandfc7ebc28/mac) 当前 `UnitDescriptor.sheet` 只有一组 rows/columns/frozen extent，若把多个表强行拼成一个网格，会损坏几何、命中和源映射。Numbers 原生支持需要先把 sheet 定义成自由画布，table 是有坐标的一级对象；现有 XLSX/ODS 单网格视图可继续作为该模型的受限特例。
- Keynote 动画/转场、Numbers 公式计算、Pages 应用特定分页都超出当前只读静态 renderer 边界，应明确诊断而不是模拟执行。

## 6. 推荐实现路线

### 阶段 0：现代语料与双 spike（3–4 周）

1. 用当前受支持 macOS/iOS iWork 版本建立有授权的 Pages/Numbers/Keynote 语料，覆盖单文件、package、密码、字体、图片、表格、图表、母版、公式、评论和损坏输入；由对应 Apple iWork 应用导出的 PDF 或固定 sheet viewport 是唯一验收金标，同时导出的 OOXML 只用于解析差分研究。
2. 固定 libetonyek 0.1.13/当前 master，在 native CLI 上跑同一语料，记录识别、解析、崩溃、超时和明显内容丢失；它只是研究比较器，不进入运行时或验收 oracle。
3. 在隔离分支写最小 Rust IWA container spike：外层 ZIP、`Index.zip`、Snappy framing、无 schema Protobuf wire reader、ArchiveInfo/object index，只输出清单与稳定 source key。
4. 可并行做一次不合入产品的 libetonyek Wasm 构建 spike，比较 Wasm 体积、峰值内存、解析时间和 provenance 改造成本。

继续门槛：当前版本单文件识别率达到 100%，container 解析无未分类崩溃；资源预算可覆盖正常样本且能拒绝构造炸弹；原生 object ID 可稳定输出；schema 未知字段可跳过并产生诊断。

### 阶段 1：诚实的 preview-only（1–2 周）

- 识别 `.pages/.numbers/.key` 单文件，不信任扩展名；验证内部结构和预览图片签名/尺寸。
- 只渲染内嵌 `preview.jpg`，DocumentInfo/diagnostics 明确 `preview-only`，不创建虚假的文本或表格对象。
- package 目录、加密文件、无预览文件给出可操作错误。

这一步能快速改善“完全打不开”，但市场文案仍应写“首屏预览”，不能写“兼容 iWork”。

### 阶段 2：共享 IWA 基础 + Pages P0（8–12 周）

- 在 Rust 内实现有硬上限的 Snappy/IWA/Protobuf reader、component index、对象引用循环检测、字符串和媒体读取。
- 新增 iWork source locator，不经 OOXML 中转。
- Pages 先支持页面尺寸、段落/文本样式、基本表格、图片、页眉页脚、显式分页；复杂浮动、脚注/评论逐项标 approximate/unsupported。
- 把 libetonyek 和 Apple 导出像素只作为独立对照，沿用 [准确性测试体系](../ACCURACY.md)。

### 阶段 3：Numbers P0（10–16 周）

- 先完成自由画布 sheet + 多 table 对象模型，再解析工作区、表坐标、行列、单元格、格式、缓存公式值、图片和基本图表。
- 不执行公式、外部数据、股票数据或脚本；只显示缓存结果和公式源信息。

### 阶段 4：Keynote P0（12–18 周）

- 支持幻灯片尺寸、母版/布局、形状、文本、图片、表格、静态图表和演讲者备注。
- 动画、转场、音视频播放继续明确不支持；只取静态 poster/preview。

### 阶段 5：硬化和发布门禁（8–12 周，可与 2–4 重叠）

- fuzz Snappy、varint、message length、nested message、对象引用图、嵌套 ZIP 和媒体；
- 增加 shared nested-expansion、component、archive-object、protobuf-depth/field/string/reference 等硬预算；
- 版本矩阵、视觉/结构/source mapping、命中、重复渲染确定性与 metamorphic tests；
- 发布文档按实际通过的 iWork 版本和功能列明范围，不使用“全兼容”。

上述估算以熟悉现有 renderer 的工程师为前提，阶段 0 前误差约 ±50%。一名工程师顺序完成三个格式的可发布 P0，现实量级约 9–14 个月；三名有文档格式经验的工程师并行，约 5–8 个日历月。高保真长尾是持续路线，不包含在 P0 中。

## 7. 安全边界

原生解析只有在以下约束全部成立时才应进入默认核心：

- 外层 ZIP 与 `Index.zip` 最多固定两层，条目数、累计展开字节、压缩比和 CPU 操作预算全程共享；
- Snappy chunk 长度、单 chunk 输出、累计输出和操作数有硬上限；
- Protobuf varint 最长 10 字节，message length 不越界，递归深度、field 数、字符串、对象、引用边和引用遍历步数受限；
- component/ArchiveInfo 重复 ID、悬空引用和循环引用产生结构化诊断；
- 媒体继续按声明 + 签名双重验证，不允许文档选择外部 URL；
- 密码文件继续拒绝。原始 iWorkFileFormat 只给出了 iWork '13 AES/PKCS7 逆向线索，并非当前 Apple 公开加密规范，[说明](https://github.com/obriensp/iWorkFileFormat/blob/8575e441beaaaa56f480fdd91721f5bb06d07d43/Docs/index.md#encryption)；
- 公式、动作、脚本、链接、动画和媒体永不执行；
- 所有解析和布局仍在一次性 document Worker，超时/取消直接终止 Worker。

## 最终建议

1. **现在可以决定做兼容，但产品目标定为“现代单文件 iWork 的只读、对象级子集”，不要定为全版本高保真。**
2. **默认核心选择 Rust 原生 IWA 路线。**它最符合现有依赖极简、浏览器本地、对象级、源映射和安全预算；libetonyek 仅用作协议/解析研究比较器和可选外部转换底座，native/release 验收只使用 Apple iWork 金标。
3. **先做 3–4 周阶段 0，再决定是否进入完整开发。**没有当前版本语料和源映射 spike，不应直接承诺交付日期。
4. **可以先发 preview-only，但 UI 和文档必须明确其限制。**它解决“看一眼”，不解决搜索、复制、检查、打印完整文档。
5. **若客户需要近期覆盖，用显式 external conversion adapter。**服务端 LibreOffice 或 macOS iWork/Quick Look 都应是单独能力，并标明文件离开浏览器核心、转换后不保留原生源映射。
6. **把商用法务审查设为阶段 0 门禁。**开源许可证审查与 Apple 派生 schema/type registry 的权利审查是两件事；后者未明确前不进入产品代码。
