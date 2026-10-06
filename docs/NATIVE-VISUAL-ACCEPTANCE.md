# Office / iWork / WPS 外部金标验收

“文件能打开”只表示解析没有立即失败，不代表兼容。发布验收以每页、每张幻灯片或每个工作表的最终画面为单位，对指定金标应用和 OfficeViewer Canvas 输出做可重复比较。

## 金标边界

映射是代码中的闭集，manifest 不能改写：

| 文件格式 | 唯一金标应用 | 验收单元 | 金标采集语义 |
| --- | --- | --- | --- |
| PPTX、PPT、ODP | Microsoft PowerPoint | slide | 原生 PDF 导出后按页栅格化 |
| DOCX、DOC、ODT | Microsoft Word | page | 原生 PDF 导出后按页栅格化 |
| XLSX、XLS、ODS | Microsoft Excel | sheet | 固定 A1 范围的无界面 viewport |
| `.key` | Apple Keynote | slide | 原生 PDF 导出后按页栅格化 |
| `.pages` | Apple Pages | page | 原生 PDF 导出后按页栅格化 |
| `.numbers` | Apple Numbers | sheet | 固定 A1 范围的无界面 viewport |
| `.wps` | WPS Office Writer | page | 人工复核原生 PDF 后导入并按页栅格化 |
| `.et` | WPS Office Spreadsheets | sheet | 人工复核固定 A1 范围的无界面 viewport |
| `.dps` | WPS Office Presentation | slide | 人工复核原生 PDF 后导入并按页栅格化 |

只有三套金标且不能交叉：OOXML、旧版 Office 和 ODF 只能使用对应的 Microsoft Office 应用，iWork 只能使用对应的 Apple iWork 应用，WPS/ET/DPS 只能使用 WPS Office。LibreOffice 和其他非对应应用全面禁止作为 native 或 release 验收金标；若保留它们的输出，只能用于研究，不能进入审核金标、`reference.json` 或 suite manifest。

金标不会在普通 CI 中重新生成。它们只在固定版本、字体和导出设置的隔离环境中生成，审核后连同 SHA-256 和环境信息签入；CI 只运行 OfficeViewer 候选采集和比较，避免应用升级偷偷移动基线。

## 什么才算通过

一个 native schema v2 suite 只有同时满足以下条件才通过：

1. fixture、金标 `reference.json` 和每张 golden PNG 的 SHA-256 与审核 manifest 及实际文件一致；runner 还会校验源格式/哈希、应用名称/bundle/version/build、采集语义、适用的栅格器版本、页码或 sheet 范围，以及人工 attestation 中声明的源文件/被审核 artifact 哈希绑定。页面/幻灯片 reference 的实际 PDF 也必须存在，并通过非空、字节数、`%PDF-` 签名和 SHA-256 校验；这些检查仍不能证明该 PDF 的 producer。`oracleFingerprint` 中其余 OS、语言、时区、字体等字段也是审核后锁定的声明，不是脚本对生产环境的独立证明。
2. 格式只能使用对应的 Microsoft Office、Apple iWork 或 WPS Office 金标，不能交叉，也不能逐 case 改阈值；任何声明 LibreOffice 或其他非对应应用为 oracle 的 suite 都必须直接失败。
3. 文档实际解析格式、unit 类型和 unit 总数必须与 manifest 一致；`0..unitCount-1` 缺任何一项都失败。
4. 工作表必须声明固定的 `sheetRange`，例如 `A1:H40`；OfficeViewer 会从行列轴解析出精确 viewport 并记录在候选 provenance 中。
5. 浏览器、OS、字体集、语言、时区、色彩空间、缩放和背景必须命中候选环境指纹。
6. 每个 unit 独立通过全页和固定 16×16 局部分块门槛；平均分只用于观察趋势，不能掩盖单页或小区域失败。

发布 suite 还必须覆盖本套件的全部格式，以及 `minimal`、`combination`、`enterprise`、`large` 四类可信语料。`malformed` 和 `malicious` 文件不能交给 Office/iWork 打开，只走 OfficeViewer 的结构化拒绝和安全测试。

## 锁定的阈值

当前唯一策略为 `native-visual-v1`：

| 指标 | 门槛 |
| --- | ---: |
| 单通道像素容差 | 8 |
| 邻域位移半径 | 1 px |
| 精确像素率 | 记录但不门禁 |
| 容差像素相似度 | ≥ 0.995 |
| windowed SSIM | ≥ 0.990 |
| 16×16 局部分块相似度 | 每块 ≥ 0.990 |
| AccuracyReport 总分 | ≥ 95 |

跨金标应用和浏览器的字体栅格化不可能稳定做到逐像素完全相同，所以 exact pixel 是观测指标；容差像素和 SSIM 才是发布门禁。任何放宽都必须新增策略 ID、说明理由并重新审核整个 suite，不能在 manifest 中写 `policy`。

这组门槛比“视觉差异 2% 以内”更严格：容差像素最多允许 0.5% 未匹配，同时整图 SSIM 和每个 16×16 对象区域相似度都不得低于 99%。只有 WPS release suite 中每个 unit 都通过，才能对该审核语料和固定环境声明达标；文件能打开、对象能提取或平均分达标都不能替代这一结论。

## 工作流

先构建运行时并确认工具链：

```bash
npm run build
npm run visual:native:reference -- --help
npm run visual:native:capture -- --help
```

幻灯片和页面的稳妥流程是：先在对应的 Office/iWork/WPS 中导出并人工审核 PDF，再由脚本逐字节导入、校验哈希，并固定用 `pdftoppm` 96 DPI 栅格化全部页面。脚本会记录 fixture、PDF、PNG、当前已安装应用版本/build 以及“来源由人工确认、脚本未验证”的边界：

```bash
npm run visual:native:reference -- corpus/example.pptx reviewed/example-pptx \
  --native-pdf reviewed-export.pdf --reviewed-by reviewer-id
```

在无未保存文档、无弹窗的隔离 macOS 账户或 VM 中，也可以省略 `--native-pdf`，让脚本复制 fixture 后自动调用原生应用导出；该模式需要 Automation 权限。两种方式都不得在普通 CI 中运行，且金标必须经过人工审核。

Excel（包括 ODS）、Numbers 和 WPS ET 的 PDF 是打印分页，不等于工作表画布。正式 sheet 金标必须在原生应用中以 100% 缩放、固定 A1 范围、无选择框和应用 chrome 采集，再通过 reference 脚本的 `sheet-viewport` 模式登记。`--allow-print-pages` 只适用于明确验收打印页的研究用例，不能作为 sheet 兼容性通过证据。

```bash
npm run visual:native:reference -- corpus/example.xlsx reviewed/example-xlsx \
  --sheet-viewport reviewed-capture.png --range A1:H40 --unit-index 0 \
  --reviewed-by reviewer-id
```

ODF 使用同一条 Office 金标流程：ODP 由 PowerPoint、ODT 由 Word 导出并人工审核完整 PDF；ODS 由 Excel 采集人工审核、固定 range 和 unit index 的 content-only viewport。生成器记录当前安装的 Microsoft Office 应用身份/version/build、平台/OS release/架构/语言/时区、采集设置和文件哈希；这些记录不等于对实际 producer 或完整生产环境的密码学证明：

```bash
npm run visual:native:reference -- corpus/example.odp reviewed/example-odp \
  --native-pdf reviewed-powerpoint-export.pdf --reviewed-by reviewer-id
npm run visual:native:reference -- corpus/example.ods reviewed/example-ods \
  --sheet-viewport reviewed-excel-capture.png --range A1:H40 --unit-index 0 \
  --reviewed-by reviewer-id
```

WPS 金标只允许人工复核导入：工具不会自动操作 WPS Office，也不会把内嵌缩略图、预览图或封面当作最终画面。WPS Writer/DPS 使用 `--native-pdf`，ET 使用固定范围的 `--sheet-viewport`：

```bash
npm run visual:native:reference -- corpus/example.wps reviewed/example-wps \
  --native-pdf reviewed-wps-export.pdf --reviewed-by reviewer-id
npm run visual:native:reference -- corpus/example.et reviewed/example-et \
  --sheet-viewport reviewed-et-capture.png --range A1:H40 --unit-index 0 \
  --reviewed-by reviewer-id
```

`--reviewed-by` 是每次手工 PDF/viewport 导入的必填项。它生成由人负责的 `reviewAttestation`，将 reviewer ID、闭集中的 oracle/application、fixture 哈希和被审核 PDF/PNG 哈希绑定在一起；runner 会校验字段和哈希绑定，但不会以密码学方式认证 reviewer，也无法判定手工导入文件的实际 producer，更不会独立证明完整生产环境。这部分需要由受控账户、审核记录和发布审批流程保证。

LibreOffice 可以保留为研究比较工具，但它的输出不能作为 native reference，也不能进入 regression 或 release suite。suite、reference 或 attestation 只要为格式声明错误的 oracle/application，策略就会拒绝；但手工导入路径无法识别“改名后又被人虚假声明为对应应用”的文件，伪造这种人工 attestation 属于本地证明边界之外的流程违规，必须由上述治理措施防止和追责。

候选采集与比较由同一个命令完成：

```bash
npm run visual:native:capture -- suites/native/office-release.json \
  --output output/accuracy/office-release
```

发布前一次验证全部 15 种格式时，使用聚合入口；它会分别执行 Office、iWork 与 WPS release suite，并汇总缺失或失败，任一 suite 未通过即返回非零：

```bash
npm run visual:native:release
npm run visual:native:release -- --validate-only
```

聚合入口默认保持严格：缺少任何 release manifest 或已有 suite 校验失败都会返回非零。发布工作流经明确授权使用 `--allow-missing`；该参数只允许“审核资产尚不存在”降级为警告，任何已经存在但校验失败的 suite 仍然阻断。此类发布的 `release.json` 和 GitHub Release 说明必须标记 `not-certified`，不得宣称通过原生视觉验收。

它会：

- 校验 manifest、fixture/golden 哈希和字体集；
- 启动本地只读 visual harness；
- 用锁定版本的 Playwright 驱动 Chromium，采集每个 unit 的 Canvas PNG；
- 写入候选 provenance（格式、unit 数、sheet viewport、诊断和 PNG 哈希）；
- 调用现有 AccuracyReport 比较器并输出逐 unit 报告及 `summary.json`。

只校验 suite 和资产、不启动浏览器：

```bash
npm run visual:native:capture -- suites/native/office-release.json --validate-only
```

## Manifest v2 最小示例

```json
{
  "schemaVersion": 2,
  "suiteKind": "regression",
  "oracleMode": "read-only",
  "oracleSuite": "microsoft-office",
  "thresholdPolicy": "native-visual-v1",
  "oracleFingerprint": {
    "os": "macOS",
    "osVersion": "27.0",
    "architecture": "arm64",
    "locale": "zh-CN",
    "timezone": "Asia/Shanghai",
    "colorSpace": "srgb",
    "scale": 1,
    "background": "#ffffff",
    "fontSetDigest": "sha256:<64 lowercase hex>",
    "applications": {
      "powerpoint": { "version": "16.111", "build": "16.111.26071325", "capture": "pdf-export" }
    },
    "rasterizer": { "name": "pdftoppm", "version": "26.05.0", "dpi": 96 }
  },
  "candidateFingerprint": {
    "os": "macOS",
    "osVersion": "27.0",
    "architecture": "arm64",
    "locale": "zh-CN",
    "timezone": "Asia/Shanghai",
    "colorSpace": "srgb",
    "devicePixelRatio": 1,
    "scale": 1,
    "background": "#ffffff",
    "fontSetDigest": "sha256:<64 lowercase hex>",
    "browser": "Chromium",
    "browserVersion": "150.0.0.0"
  },
  "documents": [{
    "id": "bug-314-title-wrap",
    "corpusClass": "minimal",
    "regressionBugId": "BUG-314",
    "format": "pptx",
    "fixture": "fixtures/bug-314.pptx",
    "fixtureSha256": "sha256:<64 lowercase hex>",
    "unitCount": 1,
    "units": [{
      "index": 0,
      "referenceJson": "goldens/bug-314/reference.json",
      "referenceJsonSha256": "sha256:<64 lowercase hex>",
      "goldenPng": "goldens/bug-314/page-0001.png",
      "goldenPngSha256": "sha256:<64 lowercase hex>",
      "actualPng": "candidate/bug-314-01.png",
      "actualObservationJson": "candidate/bug-314-01.json"
    }]
  }]
}
```

工作表 unit 需额外加入 `"sheetRange": "A1:H40"`。

这里的 `capture: "pdf-export"` 表示该格式要求原生应用 PDF 导出语义；若使用 `--native-pdf` 导入，`reference.json` 会标记 `reviewed-native-pdf-import` 并写入 `reviewAttestation`，明确脚本只验证声明结构、文件完整性和哈希绑定，不冒充对 reviewer 身份、PDF producer 或完整环境的自动证明。每个 unit 必须引用该 `reference.json` 及其哈希，golden PNG 必须就是其中对应页或 viewport 的原始文件。

## Bug 闭环

以后每个兼容性 Bug 都要转成资产，而不是修完即丢：最小 fixture、对应原生金标、Bug ID、自动候选截图和独立 unit 报告。这样能够明确回答“已覆盖多少格式/特性、当前有多少失败、哪些 Bug 回归”，而不是用随机文件估算未知 Bug 总量。

## 发布门禁状态

本规范和工具只建立验收能力，不虚构真实金标。仓库在签入并审核以下三套资产前，`npm test` 只能证明框架正确，不能宣称 Office/iWork/WPS/ODF 视觉兼容通过：

- `suites/native/office-release.json`：覆盖 PPTX、PPT、DOCX、DOC、XLSX、XLS、ODP、ODT、ODS；
- `suites/native/iwork-release.json`：覆盖 Keynote、Pages、Numbers；
- `suites/native/wps-release.json`：覆盖 WPS、ET、DPS。

三套 manifest 都必须是 `suiteKind: "release"`，覆盖四类可信语料，并在发布流程中分别执行 `visual:native:capture`。缺少审核 manifest 时，当前制品发布策略允许继续但必须产生工作流警告，并把发布标记为 `not-certified`；不得生成假金标或把其他测试当作替代证据。一旦 manifest 存在，任何失败、缺格式、缺 unit、缺 reference 链、声明的 oracle 应用/栅格器字段不匹配或候选环境指纹漂移都阻断发布。只有三套全部存在且通过时，发布状态才可写为 `passed`。
