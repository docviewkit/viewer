# OOXML 官方测试资产调研

调研日期：2026-07-15

## 结论

截至调研日期，在 ECMA、ISO/IEC 和 Microsoft 的公开官方入口中，**没有发现可公开下载、带完整预期结果并由 ISO/IEC 或 Ecma International 声明为 OOXML 官方一致性认证套件的文件集**。

可以直接用于本项目的官方资产分为两类：

1. [ECMA-376 官方发布包](https://ecma-international.org/publications-and-standards/standards/ecma-376/)：规范正文、Strict/Transitional XSD、RELAX NG、预定义几何和少量规范辅助资源。它们是规范资产，不是文档渲染测试套件。
2. [dotnet/Open-XML-SDK](https://github.com/dotnet/Open-XML-SDK) 的测试文件：Microsoft/.NET Open XML SDK 的实现回归语料，包含大量 DOCX、XLSX 和 PPTX，可自动化批量下载。它不是 ISO/ECMA 官方一致性套件，也没有渲染基准图。

因此，推荐将 Open XML SDK 测试文件用于 OfficeViewer 的批量解析、资源限制、诊断、对象模型和“能够完成渲染”验证；将 ECMA 模式用于结构与命名空间核对。两者都不能单独证明视觉保真度或获得 ISO/ECMA 认证。

## 资产分类

| 资产 | 发布方 | 性质 | 能否批量获取 | 对 OfficeViewer 的价值 |
| --- | --- | --- | --- | --- |
| ISO/IEC 29500 | ISO/IEC | 国际标准正文及一致性要求 | 需通过 ISO Webstore 获取；不是测试文件仓库 | 规范判定依据 |
| ECMA-376 发布包 | Ecma International | 规范、XSD/RNG、规范辅助资源 | 可以，官方 ZIP 为固定 URL | 结构验证、Strict/Transitional 词汇核对 |
| Open XML SDK 测试资产 | Microsoft/.NET Foundation 项目 | SDK 实现回归文件 | 可以，Git tag、源码归档或 Git sparse checkout | 多类 OOXML 扩展名的批量健壮性和能力覆盖测试 |
| Microsoft Standards Support | Microsoft | Office 对 ISO/IEC 29500 的实现说明和扩展规范 | 文档可下载；没有公开文件测试套件 | 解释 Office 与标准的差异 |

## 1. ISO/IEC 29500：规范，不是测试套件

[ISO/IEC 29500-1:2016 官方页面](https://www.iso.org/standard/71691.html)说明该标准定义 WordprocessingML、SpreadsheetML、PresentationML，以及 Strict 和 Transitional 一致性概念。该页面提供的是标准出版物，没有列出测试文档、执行器、预期结果或一致性证书流程。

ISO 原先的 Publicly Available Standards 目录已经关闭，官方页面要求改从 ISO/IEC Webstore 获取出版物：[ISO ITTF 关闭说明](https://standards.iso.org/ittf/PubliclyAvailableStandards/index.html)。这意味着不能把旧目录视为可持续的自动化测试文件源。

版权边界也很明确：[ISO Copyright](https://www.iso.org/copyright.html)声明 ISO 出版物受版权保护，未经许可不得复制或分发；[ISO End Customer Licence Agreement](https://www.iso.org/terms-conditions-licence-agreement.html)还对共享系统、数字集成和自动化使用设置了许可要求。即使某一出版物价格为零，也不等于可以把它复制进开源仓库或 CI 镜像。

结论：ISO/IEC 29500 是判定规范正确性的权威来源，但本项目不应把 ISO 出版物作为可再分发的测试语料。

## 2. ECMA-376 官方发布包：模式和规范辅助资源

[ECMA-376 官方页面](https://ecma-international.org/publications-and-standards/standards/ecma-376/)公开了四个部分的下载包，并将该标准关联到 ISO/IEC 29500。官方 [TC45 工作范围](https://ecma-international.org/technical-committees/tc45/)明确提到维护 OOXML 标准并提供、测试跨平台 W3C XML Schema；页面没有发布文档渲染一致性套件。

当前发布包：

- [Part 1：Fundamentals and Markup Language Reference](https://ecma-international.org/wp-content/uploads/ECMA-376-1_5th_edition_december_2016.zip)
- [Part 2：Open Packaging Conventions](https://ecma-international.org/wp-content/uploads/ECMA-376-2_5th_edition_december_2021.zip)
- [Part 3：Markup Compatibility and Extensibility](https://ecma-international.org/wp-content/uploads/ECMA-376-3_5th_edition_december_2015.zip)
- [Part 4：Transitional Migration Features](https://ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip)

对这些 ZIP 的实测内容检查结果：

- Part 1 包含 Strict XSD、Strict RELAX NG、预定义 DrawingML 几何、表格/单元格样式资源和 WordprocessingML 艺术边框资源。
- Part 1 的 `OfficeOpenXML-SpreadsheetMLStyles.zip` 中有一个 `PivotTableFormats.xlsx`；它是规范辅助资源，不构成覆盖 DOCX/XLSX/PPTX 的测试套件。
- Part 2 包含 OPC 的 XSD 和 RELAX NG。
- Part 3 只有规范正文。
- Part 4 包含 Transitional XSD 和 RELAX NG。

这些资产适合做：

- OOXML 包结构和关系类型的核对；
- Strict 与 Transitional 命名空间覆盖检查；
- XML 层面的 schema 辅助验证；
- DrawingML 预定义几何和样式字典实现依据。

这些资产不提供：

- 大规模 DOCX/XLSX/PPTX 输入集合；
- “该页面应渲染成什么样”的基准图；
- 命中测试或源对象映射的预期结果；
- 可给产品颁发一致性结论的测试执行器。

### ECMA 版权与再分发

[Ecma text copyright policy](https://ecma-international.org/policies/by-ipr/ecma-text-copyright-policy/)允许在保留版权声明和许可文本等条件下复制、发布和分发默认许可覆盖的文档，但原则上不得修改标准正文；Ecma 也建议从官网获取权威版本。不同包内的软件或模式是否另有 BSD 条款，应以对应发布物中的版权页为准。

工程上建议在测试准备阶段从官方 URL 下载并缓存，不把完整 ECMA 标准包直接打入 OfficeViewer npm 产物。若需要把模式或辅助资源长期 vendoring 到仓库，应同时保存原始版权和许可文本并做一次法务确认。

## 3. Open XML SDK 官方实现回归语料

[Open XML SDK 官方仓库](https://github.com/dotnet/Open-XML-SDK)说明该 SDK 面向 Word、Excel 和 PowerPoint 文档，并紧密跟随 Microsoft Office 对 ISO 29500 的实现。仓库的 [`test/DocumentFormat.OpenXml.Tests.Assets`](https://github.com/dotnet/Open-XML-SDK/tree/v3.5.1/test/DocumentFormat.OpenXml.Tests.Assets) 项目将测试文件作为嵌入资源使用。

为了获得可复现结果，应固定到 [v3.5.1 release](https://github.com/dotnet/Open-XML-SDK/releases/tag/v3.5.1)，对应 commit `3139fdfd27414548a41555f7848d5728f6e71a42`，不要直接跟随 `main`。

在 v3.5.1 上对 `assets` 目录实测枚举得到：

| 分类 | 数量 |
| --- | ---: |
| 全部文件 | 896 |
| Wordprocessing 文件（DOCX/DOTX/DOCM/DOTM） | 505 |
| Spreadsheet 文件（XLSX/XLTX/XLSM/XLTM） | 155 |
| Presentation 文件（PPTX/POTX/PPTM/POTM/PPSM/PPAM） | 227 |
| 其他辅助文件 | 9 |
| 解出后的资产体积 | 约 51 MiB |

主要目录：

- [`O14ISOStrict`](https://github.com/dotnet/Open-XML-SDK/tree/v3.5.1/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/O14ISOStrict)：157 个文件，涵盖 Word、Excel、PowerPoint 和 DrawingML Strict 场景。
- [`O15Conformance`](https://github.com/dotnet/Open-XML-SDK/tree/v3.5.1/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/O15Conformance)：70 个文件，但内容集中在 Word 扩展评论和 Excel Web Extension，不能按目录名推断为完整 OOXML conformance suite。
- `Robustness`：9 个 XLSX 健壮性文件。
- `v2FxTestFiles`：588 个旧版 SDK 回归文件。
- `TestFiles`：71 个常规、畸形、加密、扩展和主动内容样例。

仓库自身的 [`IsoStrictTest`](https://github.com/dotnet/Open-XML-SDK/blob/v3.5.1/test/DocumentFormat.OpenXml.Tests/IsoStrictTest/IsoStrictTest.cs)用于验证 Strict 关系识别和 Open XML SDK validator 行为；[`Robustness`](https://github.com/dotnet/Open-XML-SDK/blob/v3.5.1/test/DocumentFormat.OpenXml.Tests/OFCatTest/Robustness.cs)验证指定 XLSX 可打开且 validator 不报错。这些测试没有 Canvas、页面图像或对象边界的 golden result，说明该语料的原始目标不是视觉渲染验证。

### 自动化获取

推荐 sparse checkout，只下载测试资产，并固定 release tag：

```bash
mkdir -p .cache

git clone \
  --depth 1 \
  --branch v3.5.1 \
  --filter=blob:none \
  --sparse \
  https://github.com/dotnet/Open-XML-SDK.git \
  .cache/open-xml-sdk

git -C .cache/open-xml-sdk sparse-checkout set \
  test/DocumentFormat.OpenXml.Tests.Assets
```

也可以下载固定 tag 的源码归档：

```text
https://github.com/dotnet/Open-XML-SDK/archive/refs/tags/v3.5.1.zip
```

CI 应记录 tag、commit 和每个输入文件的 SHA-256，避免上游资产变更导致基线漂移。测试文件只进入测试缓存，不进入 SDK 发布包。

### 许可证

仓库根目录使用 [MIT License](https://github.com/dotnet/Open-XML-SDK/blob/v3.5.1/LICENSE)，允许使用、复制、修改和再分发，但要求在副本或重要部分中保留版权和许可声明。仓库 README 将项目描述为 .NET Foundation 项目，并标明 Open XML SDK by Microsoft。

若仅由 CI 从固定 tag 下载，许可证管理最简单。若要把 51 MiB 资产复制进本仓库，应保留上游 `LICENSE`、记录来源 commit，并在发布前检查文档内部是否包含需要单独处理的第三方内容、商标或隐私信息。

## 4. Microsoft Standards Support：实现说明，不是输入语料

[Word, Excel, and PowerPoint Standards Support](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-offstandlp/d5784a8b-7070-466b-befa-b7bf3724c6f0)收录 Microsoft 对 ECMA-376、ISO/IEC 29500 和 Office 扩展的技术文档，并提供整套 PDF ZIP。其 [MS-OI29500 conformance statements](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/bd9e8289-844a-42e2-9809-66c7005bd9e2)描述 Office 2010 及后续 Excel、PowerPoint、Word 对 Strict 和 Transitional 的实现行为。

这些内容应作为兼容性诊断和 Office 差异处理的依据，但官方入口没有提供对应的 DOCX/XLSX/PPTX 测试文件集和期望渲染结果。

## 5. 对 OfficeViewer 的验证方案

Open XML SDK 语料适合按以下层次接入，不能只统计“打开成功率”。

### 第一层：安全与生命周期

对全部 896 个资产执行：

- 读取文件字节，不依赖扩展名识别实际格式；
- 每个文件必须在超时和内存预算内结束；
- 不得崩溃、卡死 Worker 或泄漏文档资源；
- 宏、ActiveX、OLE、外部关系和加密包必须阻断或输出明确诊断。

### 第二层：解析和对象协议

对能够打开的 DOCX/XLSX/PPTX 执行：

- 场景快照可解码，unit/object ID 唯一且父子关系无环；
- 数值边界有限，宽高非负；
- 对象源映射与格式相符，例如 PPTX Shape ID、XLSX 单元格地址、DOCX 段落或文本范围；
- 对每个可见对象中心点执行 hit test，并验证返回对象属于同一 unit；
- 不支持内容必须产生结构化诊断，不能静默丢弃。

### 第三层：渲染执行

- 对每个 unit 在受限 viewport 下调用渲染；
- 记录成功渲染对象数、诊断数、耗时和峰值内存；
- 把“完整渲染”“带诊断的近似渲染”“预期拒绝”“引擎错误”分开统计；
- Strict 文件单独成组，避免被大量 Transitional 文件掩盖。

### 第四层：视觉保真

该上游语料没有官方 golden image，因此不能自动判断布局、字体、分页、图形和表格是否与 Office 一致。视觉层必须另建有明确授权的基准：PPTX/PPT、DOCX/DOC、XLSX/XLS 只能由固定版本的 Microsoft PowerPoint、Word、Excel 生成参考画面，并记录操作系统、字体集合、DPI 和应用版本；人工审核负责确认和签署该来源，不能替代金标应用。LibreOffice、WPS 或其他渲染器只可用于差分研究，不得作为 native/release 视觉验收基线。这样的结果属于项目自己的兼容性基线，不得称为 ISO/ECMA 官方 conformance 结果。

## 6. 可对外使用的准确表述

可以表述为：

> OfficeViewer 已使用固定版本的 Microsoft/.NET Open XML SDK 回归语料，对 OOXML 文档执行批量解析、资源限制、诊断、对象协议和渲染执行验证；Strict OOXML 使用独立分组统计。

不能表述为：

> OfficeViewer 通过了 ISO/IEC 或 Ecma International 官方 OOXML 一致性测试。

除非未来获得标准组织明确发布的套件、测试规范和合格判定规则，否则后一表述没有证据基础。
