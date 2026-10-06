# LibreOffice 全局文档 QA fixture 库调研

调研日期：2026-07-19

## 结论

`sc/qa/unit/data` 只是 Calc 的一块。LibreOffice 官方 [GitHub 组织](https://github.com/LibreOffice) 中，文档 QA 的主体在只读镜像 [`LibreOffice/core`](https://github.com/LibreOffice/core)；其中 Writer、Calc、Impress/Draw、OOXML/ODF、PDF、SVG、WMF/EMF、图像 filter、存储包和旧格式转换都有独立 fixture 与断言。

本文固定到官方 `core` `master` HEAD [`db5d24b4e0d0350125babbaf377d2ed7f043e161`](https://github.com/LibreOffice/core/commit/db5d24b4e0d0350125babbaf377d2ed7f043e161)（作者时间 2026-07-18 22:19:46 +02:00，提交者时间 2026-07-19 11:13:36 +02:00）。按“路径同时包含 `/qa/` 与 `/data/`”粗盘，固定树有 **10,256 个 tracked paths**。这个数字包含 `.gitignore`、证书、图片、XML/CSV oracle 和其他辅助文件，也会遗漏不放在 `data/` 下的 fixture；**不是 10,256 份可打开文档，更不是兼容性通过数**。

对 OfficeViewer 最有价值的是“fixture + 相邻测试断言”的配对：导入后模型属性、导出包 XML/XPath、round-trip、Writer 布局 dump、图形 primitive/metafile dump、像素/校验和 `pass` / `fail` 载入结果。它们不是统一的像素 golden，也不能替代 Microsoft Office / Apple iWork 原生金标。

## 固定 core 树的全局粗盘

可复核口径：

```sh
git -c core.quotepath=false ls-tree -r --name-only \
  db5d24b4e0d0350125babbaf377d2ed7f043e161 \
  | awk '/\/qa\// && /\/data\//'
```

| 模块 | `/qa/**/data/**` tracked paths | 主要内容 |
| --- | ---: | --- |
| [`sw`](https://github.com/LibreOffice/core/tree/db5d24b4e0d0350125babbaf377d2ed7f043e161/sw/qa) | 5,048 | Writer DOCX/ODT/RTF/DOC/FODT、布局、邮件合并、HTML、OOXML/RTF filter。 |
| [`sc`](https://github.com/LibreOffice/core/tree/db5d24b4e0d0350125babbaf377d2ed7f043e161/sc/qa) | 1,984 | Calc XLSX/ODS/XLS/FODS、公式、表格模型、UI/瓦片渲染。 |
| [`sd`](https://github.com/LibreOffice/core/tree/db5d24b4e0d0350125babbaf377d2ed7f043e161/sd/qa) | 889 | Impress/Draw PPTX/ODP/PPT/ODG/PDF/SVG、布局、导出与瓦片渲染。 |
| [`vcl`](https://github.com/LibreOffice/core/tree/db5d24b4e0d0350125babbaf377d2ed7f043e161/vcl/qa) | 788 | PNG/JPEG/TIFF/BMP/WebP、PDF 导出、SVM、图像 filter 与 bitmap 操作。 |
| [`chart2`](https://github.com/LibreOffice/core/tree/db5d24b4e0d0350125babbaf377d2ed7f043e161/chart2/qa) | 433 | XLSX/ODS/DOCX/PPTX/ODP 中的图表模型、导入导出和渲染。 |
| [`writerperfect`](https://github.com/LibreOffice/core/tree/db5d24b4e0d0350125babbaf377d2ed7f043e161/writerperfect/qa) | 277 | Document Liberation 旧格式/iWork 转换；其中也有占位文件。 |
| `oox` / `svgio` / `svx` / `emfio` | 453 | DrawingML/OOXML、SVG、绘图模型、WMF/EMF。 |
| `xmlsecurity` / `xmloff` / `filter` / `package` | 201 | 签名加密、ODF XML、跨格式 filter、ZIP/存储包。 |
| `lotuswordpro` / `starmath` / 其他 | 183 | LWP、MathML 及其他模块小型 corpus。 |

全局主要扩展名为 DOCX 2,244、ODT 1,277、ODS 710、FODT 689、RTF 661、PPTX 552、XLSX 547、FODS 543、DOC 357、ODP 326、XLS 317；还有 PNG 237、SVG 138、PDF 104、EMF 92、HTML 88、XML 84、PPT 81。扩展名统计只是路由线索，不代表真实 MIME、测试意图或成功期望。

## 三大办公应用 corpus

### Writer：`sw/qa`

Writer 是最大的一组。`sw/qa/**/data/**` 共 5,048 个 tracked files（106,186,613 bytes，约 101.27 MiB）；主要扩展名为 DOCX 2,107、ODT 1,117、RTF 660、FODT 580、DOC 341。最重要的子库是 [`ooxmlexport/data`](https://github.com/LibreOffice/core/tree/db5d24b4e0d0350125babbaf377d2ed7f043e161/sw/qa/extras/ooxmlexport/data) 1,642、`rtfexport/data` 447、`uiwriter/data` 419、`layout/data` 411、`core/data` 265、`ww8export/data` 228、`odfexport/data` 199、`ooxmlimport/data` 160。

相邻的 [`ooxmlexport*.cxx`](https://github.com/LibreOffice/core/tree/db5d24b4e0d0350125babbaf377d2ed7f043e161/sw/qa/extras/ooxmlexport) 会断言 Writer 文档模型、保存后 `word/*.xml` 的 XPath、round-trip，以及 `parseLayoutDump()` 产生的页/行/文字布局树。因此 Writer fixture 特别适合页面几何、段落、表格、列表、浮动对象、页眉页脚和导出包结构回归；但 layout dump 是 LibreOffice 内部布局模型，不是 Word 原生布局 oracle。

### Calc：`sc/qa`

`sc/qa/**/data/**` 粗盘有 1,984 个 tracked files，主要是 ODS 574、FODS 535、XLSX 386、XLS 310。核心库仍是 [`sc/qa/unit/data`](https://github.com/LibreOffice/core/tree/db5d24b4e0d0350125babbaf377d2ed7f043e161/sc/qa/unit/data)；它的 1,673 文件和预期结果机制在下文保留详细盘点。其他主要路径是 `sc/qa/uitest/data` 192、`sc/qa/unit/uicalc/data` 67、`sc/qa/unit/tiledrendering/data` 41。

### Impress / Draw：`sd/qa`

`sd/qa/**/data/**` 粗盘有 889 个 tracked files（39,896,674 bytes，约 38.05 MiB），主要是 PPTX 419、ODP 240、PPT 72、ODG 28、FODP 28。核心 [`sd/qa/unit/data`](https://github.com/LibreOffice/core/tree/db5d24b4e0d0350125babbaf377d2ed7f043e161/sd/qa/unit/data) 有 821 个 files，包含 PPTX 413、ODP 202、PPT 72、ODG 23、PDF 17、FODP 16，以及 CGM/SVG/PNG 等图形输入。

相邻 [`sd/qa/unit/*.cxx`](https://github.com/LibreOffice/core/tree/db5d24b4e0d0350125babbaf377d2ed7f043e161/sd/qa/unit) 检查幻灯片/画布模型、shape、主版、动画、循环与媒体元数据、导出 XML、round-trip、PDF 导入和 PNG 导出。这是 OfficeViewer 演示文档和绘图对象回归的主要上游库，但并无统一的 PowerPoint 像素级真值。

## 共享解析、渲染与包层

| 路径 | 可复核规模 | 断言机制与 OfficeViewer 用途 |
| --- | ---: | --- |
| [`filter/qa`](https://github.com/LibreOffice/core/tree/db5d24b4e0d0350125babbaf377d2ed7f043e161/filter/qa) | 37 个 `data` paths；其中 `unit/data` 31 | ODT/ODS/ODP/ODG/DOC/PPTX/PDF/混合 PDF、filter detect、SVG/PDF 导出与证书签名辅助文件。是跨应用 filter 行为，不是大型文档 corpus。 |
| [`oox/qa/unit/data`](https://github.com/LibreOffice/core/tree/db5d24b4e0d0350125babbaf377d2ed7f043e161/oox/qa/unit/data) | 153 files，3,780,550 bytes | PPTX 62、DOCX 36、OOXML 片段 `.bin` 14、ODP 13、ODT 12。检查 DrawingML、VML、shape、3D scene、MathML、VBA 压缩/加密等共享 OOXML 层。 |
| [`xmloff/qa/unit/data`](https://github.com/LibreOffice/core/tree/db5d24b4e0d0350125babbaf377d2ed7f043e161/xmloff/qa/unit/data) | 64 files，1,163,611 bytes | 主要为 FODT/ODT/DOCX/PPTX/FODG/ODP/ODG；通过 UNO 模型属性、保存再载入和 `content.xml` / `styles.xml` XPath 检查 ODF XML 语义。 |
| [`vcl/qa`](https://github.com/LibreOffice/core/tree/db5d24b4e0d0350125babbaf377d2ed7f043e161/vcl/qa) | 788 个 `data` paths；`graphicfilter` 子树 311 | PNG/JPEG/TIFF/GIF/BMP/WebP/PSD/EPS/PCX/SVM/WMF/EMF 等 decode/encode、bitmap 尺寸/像素/校验、PDF 导出。对内嵌图片与 PDF 管线直接有用。 |
| [`emfio/qa`](https://github.com/LibreOffice/core/tree/db5d24b4e0d0350125babbaf377d2ed7f043e161/emfio/qa) | 85 个 `data` paths；EMF 根 59、WMF 根 26 | 共含 EMF 66、WMF 18、PPTX 1；将 WMF/EMF 解码成 metafile XML，用 XPath 检查坐标、clip、线型、字体、填充和 bitmap，适合 Office 内嵌矢量图回归。 |
| [`svgio/qa/cppunit/data`](https://github.com/LibreOffice/core/tree/db5d24b4e0d0350125babbaf377d2ed7f043e161/svgio/qa/cppunit/data) | 126 SVG | 将 SVG 分解为 drawinglayer primitive XML 并 XPath 断言 geometry、paint、transform、mask/filter、text；是结构 oracle，不是浏览器像素 golden。 |
| [`package/qa/cppunit/data`](https://github.com/LibreOffice/core/tree/db5d24b4e0d0350125babbaf377d2ed7f043e161/package/qa/cppunit/data) | 27 files | DOCX/ODT/ODS/XLSX/ODG/ZIP/OXT 包，配合 [`test_package.cxx`](https://github.com/LibreOffice/core/blob/db5d24b4e0d0350125babbaf377d2ed7f043e161/package/qa/cppunit/test_package.cxx) 验证 ZIP/storage 导入、损坏包和 pass/fail。 |

`writerfilter` 的测试在 [`sw/qa/writerfilter`](https://github.com/LibreOffice/core/tree/db5d24b4e0d0350125babbaf377d2ed7f043e161/sw/qa/writerfilter)，不在根级 `writerfilter/qa`；当前固定树没有 `avmedia/qa`。媒体行为主要是包在 `sd` 文档内的关系和对象模型测试，不能从 core 得到一套独立音视频播放 corpus。

## LibreOffice GitHub 组织中的其他格式库

[LibreOffice 组织仓库列表](https://github.com/orgs/LibreOffice/repositories?type=all) 不只有 `core`，还有 Document Liberation 格式解析库。它们的测试数据可作为 OfficeViewer 扩展格式的上游语料，但不是 Writer/Calc/Impress 的集成渲染测试：

- [`libetonyek`](https://github.com/LibreOffice/libetonyek) 处理 Apple iWork；`src/test/data` 有 27 个顶层 fixture 条目（package 目录展开后 52 个 tracked files），覆盖 Keynote/Pages/Numbers 样本。不应把 52 个 blob 说成 52 份独立 iWork 文档。
- [`libvisio`](https://github.com/LibreOffice/libvisio) 处理 Visio；`src/test/data` 有 37 个顶层 fixture（VSD 18、VSDX 16、WMF 2、PNG 1）。
- [`libcdr`](https://github.com/LibreOffice/libcdr)、[`libmspub`](https://github.com/LibreOffice/libmspub)、[`libfreehand`](https://github.com/LibreOffice/libfreehand)、[`libabw`](https://github.com/LibreOffice/libabw) 分别面向 CorelDRAW、Microsoft Publisher、FreeHand 和 AbiWord；它们应作为各格式 parser 的独立来源和许可边界管理。

## Calc 核心库深入盘点

[`sc/qa/unit/data`](https://github.com/LibreOffice/core/tree/db5d24b4e0d0350125babbaf377d2ed7f043e161/sc/qa/unit/data) 在固定树中有 1,673 个 tracked files，151,113,893 bytes（约 144.11 MiB）。它是仍随 `core` 开发维护的 Calc 导入、导出、公式、数据模型和安全回归库。文件本身通常不构成完整测试；预期结果主要由相邻 [`sc/qa/unit/*.cxx`](https://github.com/LibreOffice/core/tree/db5d24b4e0d0350125babbaf377d2ed7f043e161/sc/qa/unit) 断言、`contentCSV/`、公式文档内嵌的 `Expected` / `Correct` 列和通用 filter 框架共同赋予。

### Calc 目录结构

目录按测试机制和输入格式混合组织，不只是按扩展名分类：

| 路径 | 固定 HEAD 文件数 | 作用 |
| --- | ---: | --- |
| [`functions/`](https://github.com/LibreOffice/core/tree/db5d24b4e0d0350125babbaf377d2ed7f043e161/sc/qa/unit/data/functions) | 508 | 公式正确性主库；其中 507 个 FODS，按 add-in、array、database、date/time、financial、information、logical、mathematical、spreadsheet、statistical、text 等分类，另有一个动态数组 XLSX。 |
| [`xlsx/`](https://github.com/LibreOffice/core/tree/db5d24b4e0d0350125babbaf377d2ed7f043e161/sc/qa/unit/data/xlsx) | 355 | XLSX 导入/导出与回归，包括 data table、pivot、shared formula、track changes、安全样本；其中实际 `.xlsx` 351 个，另含一个 ODS 依赖和 3 个占位 `.gitignore`。 |
| [`ods/`](https://github.com/LibreOffice/core/tree/db5d24b4e0d0350125babbaf377d2ed7f043e161/sc/qa/unit/data/ods) | 345 | ODS 模型、格式、图形、外链、数据透视和 OpenCL 回归；含 343 个 ODS、一个 FODS、一个关联 JPG。 |
| [`xls/`](https://github.com/LibreOffice/core/tree/db5d24b4e0d0350125babbaf377d2ed7f043e161/sc/qa/unit/data/xls) | 304 | BIFF/XLS、OpenCL、shared formula/string、track changes、数据表与安全样本；含 301 个 XLS 和 3 个占位 `.gitignore`。 |
| [`contentCSV/`](https://github.com/LibreOffice/core/tree/db5d24b4e0d0350125babbaf377d2ed7f043e161/sc/qa/unit/data/contentCSV) | 39 | 不是待测输入集合，而是部分导入测试逐单元格比对的期望结果。 |
| `123/`、`dbf/`、`dif/`、`gnumeric/`、`html/`、`qpro/`、`slk/`、`wks/` | 45 | Lotus 1-2-3、DBF、DIF、Gnumeric、HTML、Quattro Pro、SYLK、Lotus WKS 等旧格式和安全回归。 |
| `xlsb/`、`xlsm/`、`csv/`、`fods/`、`xml/` | 61 | XLSB/XLSM、文本和 XML 表格输入，以及额外 FODS 样本。 |
| `dataprovider/`、`json-mapped/`、`xml-mapped/`、`solver/` | 12 | 外部数据提供器、JSON/XML 映射和求解器 fixture。 |
| 根目录文件 | 4 | [`README`](https://github.com/LibreOffice/core/blob/db5d24b4e0d0350125babbaf377d2ed7f043e161/sc/qa/unit/data/README)、[`README.cellborders`](https://github.com/LibreOffice/core/blob/db5d24b4e0d0350125babbaf377d2ed7f043e161/sc/qa/unit/data/README.cellborders) 及两个根级 FODS fixture。 |

`functions/` 的 508 个文件进一步分为：statistical 147、mathematical 79、financial 51、add-in 49、text 44、spreadsheet 44、date/time 32、information 20、array 15、database 12、logical 9、旧的通用 `fods/` 5、dynamic array 1。

## 文件格式与数量

以下数量来自固定 commit 的 `git ls-tree -r --long HEAD sc/qa/unit/data`，扩展名按小写归一化：

| 格式 / 文件类型 | 数量 | 格式 / 文件类型 | 数量 |
| --- | ---: | --- | ---: |
| FODS | 527 | XLSX | 352 |
| ODS | 348 | XLS | 301 |
| CSV | 51 | XML | 21 |
| XLSB | 9 | SLK | 9 |
| WKS | 5 | WB2 / Quattro Pro | 5 |
| HTML | 5 | DBF | 5 |
| XLSM | 4 | JSON | 2 |
| Gnumeric | 2 | DIF | 2 |
| Lotus 1-2-3 (`.123`) | 2 | JPG | 1 |
| `.gitignore` 占位 | 20 | README 类文件 | 2 |

按 OfficeViewer 当前支持矩阵，直接可进入对应格式回归的共有 1,005 个：XLSX 352、XLSM 4、ODS 348、XLS 301。XLSB 9 个应验证明确拒绝；FODS 527 个可提供公式和安全场景线索，但不是当前支持格式，不能计入打开率。

## 相邻测试如何定义“正确”

### 1. C++ 直接断言文档模型

许多测试以 `ScModelTestBase(u"sc/qa/unit/data")` 为根目录，调用 `createScDoc("xlsx/…")` 或 `createScDoc("ods/…")`，随后直接断言单元格值、公式、命名范围、样式、筛选器、数据透视、图形锚点等。例如 [`subsequent_export_test2.cxx`](https://github.com/LibreOffice/core/blob/db5d24b4e0d0350125babbaf377d2ed7f043e161/sc/qa/unit/subsequent_export_test2.cxx#L44-L106) 对导入后的 filter、导出的 OOXML XPath、图形 anchor 以及保存再载入结果分别断言。

这意味着 fixture 文件名或目录名通常只说明场景；真正的期望值在对应测试函数中。

### 2. 导出包 XML 与 round-trip 断言

导出测试会将输入保存为 XLSX/XLSM/ODS，再通过 `parseExport()` 读取包内 XML，用 XPath 检查关系、属性、公式、样式和对象结构；部分测试还会 `saveAndReload()` 后再次检查 Calc 模型。它比“成功打开”更强，但仍主要验证被明确断言的结构，不等于页面像素或完整视觉保真。

### 3. `contentCSV/` 作为逐单元格 oracle

[`qahelper.cxx`](https://github.com/LibreOffice/core/blob/db5d24b4e0d0350125babbaf377d2ed7f043e161/sc/qa/unit/helper/qahelper.cxx#L101-L148) 的 `testFile()` 读取 `contentCSV/` 对照文件；[`csv_handler.hxx`](https://github.com/LibreOffice/core/blob/db5d24b4e0d0350125babbaf377d2ed7f043e161/sc/qa/unit/helper/csv_handler.hxx#L67-L152) 将其与载入后的 sheet 逐格比较：空值、字符串精确匹配，数值使用 `1e-10` 容差。相邻测试代码负责指定对照 CSV 和 sheet index。

### 4. 公式 FODS 自带期望值

公式 fixture 在文档中保存实际公式、`Expected`、`Correct` 和 `FunctionString` 列。[`functions_test.cxx`](https://github.com/LibreOffice/core/blob/db5d24b4e0d0350125babbaf377d2ed7f043e161/sc/qa/unit/functions_test.cxx#L21-L113) 先 hard recalc，再检查汇总单元格 `Sheet1.B3`；失败时逐行输出实际值、期望值和公式。各类别测试（例如 [`functions_financial.cxx`](https://github.com/LibreOffice/core/blob/db5d24b4e0d0350125babbaf377d2ed7f043e161/sc/qa/unit/functions_financial.cxx#L25-L30)）递归扫描对应目录并要求全部通过。

### 5. `pass` / `fail` / `indeterminate` 定义过滤器结果

安全/异常输入使用格式目录下的三个子目录。通用 [`FiltersTest::testDir()`](https://github.com/LibreOffice/core/blob/db5d24b4e0d0350125babbaf377d2ed7f043e161/unotest/source/cpp/filters-test.cxx#L58-L165) 的语义是：

- `pass/`：过滤器必须成功载入；
- `fail/`：过滤器必须拒绝；
- `indeterminate/`：执行但不约束成功或失败。

固定 HEAD 中有 60 个实际 `pass` fixture、6 个实际 `fail` fixture、0 个实际 `indeterminate` fixture；其余同名目录中的 `.gitignore` 只是保留空目录。[`filters-test.cxx`](https://github.com/LibreOffice/core/blob/db5d24b4e0d0350125babbaf377d2ed7f043e161/sc/qa/unit/filters-test.cxx#L29-L108) 将该机制用于 QPro、SYLK、XLS、XLSX、XLSM、DBF 和 WKS，并明确提醒：SYLK 导入失败后可能串联尝试 CSV/RTF，因此“pass”甚至不一定证明使用了 SYLK filter。

## 安全、CVE 与恶意样本边界

这是该目录最不能按普通演示文件处理的部分。

- 固定树中有 25 个以 `CVE` 开头、3 个以 `EDB` 开头的安全 fixture，另有 13 个 `ofz…` OSS-Fuzz 回归文件，以及若干 `crash`、`forcepoint`、损坏、外链、宏、ActiveX、超大行列和 Zip64 样本。
- [`data/README`](https://github.com/LibreOffice/core/blob/db5d24b4e0d0350125babbaf377d2ed7f043e161/sc/qa/unit/data/README#L1-L7) 明确说明：文件名含 `CVE` 的样本在仓库中以 ARCFOUR/RC4、密钥 `CVE` 加密，目的是避免源码下载触发病毒扫描器。
- 通用过滤器测试还会将 `BID`、`CVE`、`EDB`、`RC4` 前缀都视为加密输入，解密到临时文件、运行测试后删除临时明文：[实现](https://github.com/LibreOffice/core/blob/db5d24b4e0d0350125babbaf377d2ed7f043e161/unotest/source/cpp/filters-test.cxx#L24-L55)、[识别与生命周期](https://github.com/LibreOffice/core/blob/db5d24b4e0d0350125babbaf377d2ed7f043e161/unotest/source/cpp/filters-test.cxx#L87-L147)。因此不能把 Git blob 直接当作正常 XLS/XLSX 交给 OfficeViewer；需要先按清单识别，且只在隔离环境内解密。
- 上游 Calc filter helper 禁用用户交互，但设置了 `MacroExecMode::ALWAYS_EXECUTE_NO_WARN`：[代码](https://github.com/LibreOffice/core/blob/db5d24b4e0d0350125babbaf377d2ed7f043e161/sc/qa/unit/helper/scfiltertestbase.cxx#L23-L48)。复用上游测试逻辑时必须运行在无网络、低权限、资源受限、一次性工作目录或容器中；不要在开发者桌面直接批量用 LibreOffice 打开这些文件。
- 对 OfficeViewer，宏、ActiveX、OLE 和外链必须保持不执行；安全 fixture 应单列为 `security` corpus，设置超时、内存/解压预算、Worker 终止与崩溃检测。`pass` 只表示上游希望 LibreOffice filter 不崩溃并成功载入，`fail` 表示应拒绝，两者都不代表内容可信或渲染正确。

## 获取、固定与清单化

不要下载 GitHub `master` ZIP 后长期使用，也不要只记录调研日期。推荐用 partial clone + sparse checkout，并在 checkout 后验证 commit：

```sh
git clone --depth 1 --filter=blob:none --no-checkout \
  https://github.com/LibreOffice/core.git \
  .cache/libreoffice-core
git -C .cache/libreoffice-core sparse-checkout init --cone
git -C .cache/libreoffice-core sparse-checkout set \
  sc/qa/unit unotest/source/cpp
git -C .cache/libreoffice-core fetch --depth 1 origin \
  db5d24b4e0d0350125babbaf377d2ed7f043e161
git -C .cache/libreoffice-core checkout --detach FETCH_HEAD
test "$(git -C .cache/libreoffice-core rev-parse HEAD)" = \
  db5d24b4e0d0350125babbaf377d2ed7f043e161
```

如果只需要二进制 fixture，可将 sparse set 缩到 `sc/qa/unit/data`；但这样会丢失大量期望结果语义。实际接入 OfficeViewer 时，应生成独立 manifest，至少记录：

- 上游 repo URL、固定 commit、相对路径、文件尺寸、SHA-256；
- 格式和场景类别；
- 对应 LibreOffice 测试源文件/函数；
- 期望结果来源：直接断言、CSV、FODS 内嵌、pass/fail；
- 是否加密、是否安全样本、是否包含宏/外链/ActiveX/OLE；
- OfficeViewer 的预期：成功解析、明确拒绝、诊断码、资源上限，以及已知 unsupported。

升级 commit 时应生成 manifest diff 并人工审阅新增/删除的 fixture 和相邻断言，不应自动滚动 `master`。

## 许可证与再分发

LibreOffice `core` 根目录带有 [`COPYING.MPL`](https://github.com/LibreOffice/core/blob/db5d24b4e0d0350125babbaf377d2ed7f043e161/COPYING.MPL) 和 [`COPYING.LGPL`](https://github.com/LibreOffice/core/blob/db5d24b4e0d0350125babbaf377d2ed7f043e161/COPYING.LGPL)；相邻测试 C++ 源文件明确使用 MPL 2.0 头，例如 [`filters-test.cxx`](https://github.com/LibreOffice/core/blob/db5d24b4e0d0350125babbaf377d2ed7f043e161/sc/qa/unit/filters-test.cxx#L1-L10)。

但 `sc/qa/unit/data` 自身没有目录级 LICENSE 或逐个二进制文档的统一许可声明；文件名和测试注释还表明其中包含 Bugzilla、论坛、CVE、Exploit-DB、OSS-Fuzz 等不同来源的材料，文档内部也可能带有第三方文本、图片或字体。**源码仓库的项目许可文件不能替代对每个外来二进制 fixture 及其内嵌内容的来源核验。**

工程上建议：

- CI/内部测试从官方仓库固定 commit 拉取并缓存，不把整库提交到 OfficeViewer；
- 不把 fixture 默认打入 npm、Wasm、示例应用或公开测试包；
- 如确需 vendoring/公开再分发，逐文件查看 `git log --follow -- <path>`、引入提交、原 bug/附件来源和声明，保留上游路径、commit、hash、许可与 notice；来源或授权不清晰的文件排除；
- CVE/EDB/恶意样本即使许可允许，也应单独受控存储，不进入普通下载和演示分发链路。

这不是法律意见；正式商业发行前应由项目的合规流程确认选中文件。

## 对 OfficeViewer 的价值与限制

### 值得接入的部分

- **XLSX/ODS 结构回归**：共享公式、动态数组、数据验证、条件格式、数据透视、表样式、批注、图形锚点、超链接、外部引用等覆盖很深。
- **公式基准**：527 个 FODS 加上明确的 Expected/Correct 机制，可用于筛选 OfficeViewer 已实现公式和缓存值处理；不能把 LibreOffice 的计算结果直接等同于 Excel 标准行为，需按函数语义分层。
- **多格式同主题对照**：`universal-content`、formats、database、matrix 等场景跨 ODS/XLS/XLSX/旧格式存在，适合验证结构化模型一致性和诊断差异。
- **安全与韧性**：CVE、OSS-Fuzz、损坏输入、外链、宏和大表适合独立验证拒绝策略、无执行保证、资源预算、崩溃隔离和确定性。
- **导入/导出语义线索**：相邻 LibreOffice 测试提供了远比文件名更精确的预期，可转译成 OfficeViewer 自己的对象协议、诊断和渲染断言。

### 不能替代的证据

- 它主要是 LibreOffice 的回归库，预期基于 LibreOffice 模型和实现，不是 ECMA-376、ODF 或 Microsoft Excel 的独立一致性认证。
- 大量测试只断言少数模型字段或 XML 节点；未被断言的内容不能推定正确。
- 没有统一截图、页面几何、字体环境或 reviewed native golden，不能证明视觉保真，也不能替代 OfficeViewer 现有的 Microsoft Office / Apple iWork 原生 golden 政策。
- `functions/` 和 OpenCL fixture 比例很高，1,673 不是 1,673 个独立终端用户文档场景；格式数量也不能直接当功能覆盖率。
- `pass/fail` 是 LibreOffice filter 的载入预期，不应原样复制成 OfficeViewer 的成功/失败预期，尤其是 OfficeViewer 不支持的旧格式、XLSB/XLSM 或主动拒绝的能力。
- 部分文件依赖外链、字体、宏、ActiveX、特定生产者或历史 bug 上下文，脱离相邻测试代码后容易误判。

### 预渲染图像可用性

**有少量可以利用的 Microsoft Office 渲染参照，但不能把目录里的 PNG/JPG/PDF 整体视为 Office 金标。** 固定树中最明确的一组是 `sd/qa/unit/data/TextFittingComparisonWithMSO_*.pptx`：相邻 [`TextFittingTest.cxx`](https://github.com/LibreOffice/core/blob/db5d24b4e0d0350125babbaf377d2ed7f043e161/sd/qa/unit/TextFittingTest.cxx#L31-L34) 直接说明这些文档包含“来自 MSO 渲染”的 bitmap，用来肉眼比较文字缩放差异。4 个 PPTX 的 `docProps/app.xml` 均记录 `Microsoft Office PowerPoint` / `AppVersion 16.0000`，包内合计有 32 个 `ppt/media/image*.png`；slide relationship 能把图片映射回对应 slide。代表文件是 [`TextFittingComparisonWithMSO_1.pptx`](https://github.com/LibreOffice/core/blob/db5d24b4e0d0350125babbaf377d2ed7f043e161/sd/qa/unit/data/TextFittingComparisonWithMSO_1.pptx)（14 个 PNG）和 [`TextFittingComparisonWithMSO_2.pptx`](https://github.com/LibreOffice/core/blob/db5d24b4e0d0350125babbaf377d2ed7f043e161/sd/qa/unit/data/TextFittingComparisonWithMSO_2.pptx)（6 个 PNG）。它们适合做**文字 fitting 的局部视觉参照**：提取内嵌 PNG，按 slide/object 关系裁切 OfficeViewer 输出后做容差或感知差异比较；不能扩展解释为整张 slide 或其他对象都已由 PowerPoint 验证。

另有一类较弱证据：在 `sw/qa`、`sc/qa`、`sd/qa` 的 OOXML 包内共发现 410 个 `docProps/thumbnail.(jpeg|emf|wmf)` 实例（加 `chart2/qa` 后为 443）；其 `app.xml` 均声明 Microsoft Office 或 Microsoft Macintosh Office，版本跨 12/14/15/16。它们通常只有一张低分辨率文档预览，例如上述 PPTX 的 JPEG thumbnail 是 256×144、96 DPI；无法覆盖全部页面/幻灯片，而且仓库没有固定 OS、字体、缩放、DPI、渲染 API、人工审阅或“缩略图未过期”的证明。可以用作**辅助 smoke/异常检测**，不应进入发布级像素门禁。

其余图片大多不是文档渲染结果：`sw/sc/sd` 外置 PNG/JPG 共 35 个，相邻代码通常把它们作为 linked/inserted graphic 输入，例如 [`layout4.cxx`](https://github.com/LibreOffice/core/blob/db5d24b4e0d0350125babbaf377d2ed7f043e161/sw/qa/extras/layout/layout4.cxx#L1018-L1033)、[`misc-tests.cxx`](https://github.com/LibreOffice/core/blob/db5d24b4e0d0350125babbaf377d2ed7f043e161/sd/qa/unit/misc-tests.cxx#L1171-L1212) 和 [`tiledrendering.cxx`](https://github.com/LibreOffice/core/blob/db5d24b4e0d0350125babbaf377d2ed7f043e161/sc/qa/unit/tiledrendering/tiledrendering.cxx#L975-L987)。VCL 的确有 checked-in expected bitmap，例如 [`BitmapFilterTest.cxx`](https://github.com/LibreOffice/core/blob/db5d24b4e0d0350125babbaf377d2ed7f043e161/vcl/qa/cppunit/BitmapFilterTest.cxx#L225-L251) 以 4 个 PNG 的 checksum 验证膨胀/腐蚀算法；这是图像算法 oracle，不是 Word/Excel/PowerPoint 文档渲染。`sd`/`vcl` 下的 PDF 也主要由相邻测试作为 PDF 导入、解析或导出管线输入，未发现可证明其为 Microsoft Office 页面渲染金标的生产环境与源文件映射。

因此接入时应单列三种 provenance：`mso-embedded-render-reference`（目前明确的是上述 TextFitting 组）、`office-package-thumbnail`（弱参照）和 `input-or-algorithm-oracle`（不可作为 Office 视觉真值）。即使第一类有 PowerPoint 生产者、版本和 slide 映射，也仍缺 OS/字体/DPI/审阅环境，只应作为场景级补充；OfficeViewer 的发布级视觉金标仍需我们自己在封闭、固定的 Microsoft Office 环境中生成和审核。

建议按“普通解析回归 / 模型断言 / 布局与对象 / 公式 / 导出结构 / 安全拒绝 / 资源极限”分层精选，不整库直接跑打开率；第一批分别从 DOCX/ODT、XLSX/ODS、PPTX/ODP 和共享图形层选取与 OfficeViewer 已支持能力相符的场景，并把每个 fixture 映射到明确的本地断言。

## 与独立 `test-files` 仓库的区别

LibreOffice 另有旧的独立 [`test-files`](https://git.libreoffice.org/test-files) 仓库，当前 HEAD 停留在 2018 年。它是另一套历史语料，不是本次用户指出的 `core` 内置 QA 数据，也不应被当作当前全局主库。本调研的 `core` 数量均以 `LibreOffice/core@db5d24b4…` 为准，并在表内明确各自路径口径；`libetonyek`、`libvisio` 的补充数量来自各自官方仓库。
