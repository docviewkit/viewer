# 官方来源语料验证基线

验证日期：2026-07-19

## 结论

OfficeViewer 已接入四个固定版本的上游语料源，并通过正式公开 SDK 执行内容识别、受限包解析、unit 元数据、结构化诊断、首对象源映射和中心点命中测试。新增 LibreOffice 集合固定到 `core@db5d24b4…`，只扫描 Writer/Calc/Impress/Chart 中与当前六种格式对应的精选 QA 根，排除显式 `fail/`、CVE/EDB/BID/RC4 和 OSS-Fuzz 文件名：

| 来源 | 固定版本 | 测试数 | 打开 | 结构化拒绝 | 非结构化失败 |
| --- | --- | ---: | ---: | ---: | ---: |
| Microsoft Open XML SDK | v3.5.1 / `3139fdfd…` | 707 | 671 | 36 | 0 |
| OASIS ODF TC 1.3 特性文件 | `16a59d94…` | 34 | 29 | 5 | 0 |
| TDF ODF Toolkit + Validator | v0.13.0 / `b926a613…` | 350 | 269 | 81 | 0 |
| LibreOffice core QA 精选集 | `db5d24b4…` | 3,645 | 3,336 | 305 | 4 |
| **合计** |  | **4,736** | **4,305** | **427** | **4** |

4,736 个路径输入对应 4,621 个唯一 SHA-256 内容。4,305 个路径完成打开；427 个结构化拒绝既包括损坏、加密、路径异常等应拒绝输入，也包括当前解析兼容缺口。另有 4 个 LibreOffice ODP 输入成功打开但首对象中心点未命中，因此被验证器归为非结构化失败。不能将“打开”或“得到结构化结果”误写成视觉、格式一致性或标准认证通过。

按格式统计：

| 格式 | 测试数 | 打开 | 结构化拒绝 | 验证失败 |
| --- | ---: | ---: | ---: | ---: |
| DOCX | 2,015 | 1,917 | 98 | 0 |
| XLSX | 614 | 551 | 63 | 0 |
| PPTX | 633 | 623 | 10 | 0 |
| ODT | 654 | 548 | 106 | 0 |
| ODS | 577 | 473 | 104 | 0 |
| ODP | 243 | 193 | 46 | 4 |

格式统计默认按文件扩展名归类；LibreOffice 中 3 个扩展名为 DOCX、实际 `mimetype` 为 ODT 的文件使用显式内容格式覆盖，并在 suite 元数据中保留。上游目录混有专门构造的非法、加密和不完整文件，也没有统一的预期结果清单，所以这些数字是可重复兼容性基线，不是合格率。

LibreOffice 的 4 个 `TextFittingComparisonWithMSO_*.pptx` 另经过包内 provenance 校验：`docProps/app.xml` 必须是 `Microsoft Office PowerPoint 16.0000`，32 个内嵌 PNG 必须全部匹配预期数量、映射到 slide relationship，并记录 fixture/image SHA-256。这些图片只作为文字 fitting 的局部视觉参照，不是整张 slide 的发布金标。

## 已发现并修复的问题

- ZIP 读取器原先拒绝合法的 Deflate 空目录项；现在允许 `uncompressed_size = 0` 的目录，并继续实际解压和校验 CRC，非空目录仍被拒绝。
- ODS 的空 `<text:p/>` 曾被当成实质文本，可能把重复尾部空行错误扩展为大量对象；现在真正的空段落保持稀疏。
- ODT `style:default-page-layout` 中的默认属性曾被误判为无父级；现在只从具名 `style:page-layout` 读取页面尺寸，并正常忽略合法默认属性。
- DOCX 表格源定位器曾使用协议未声明的 `table` / `cell` 类型组合；现在 JS 协议和 Rust 核心统一为 `table` / `table-cell`，相关 43 个协议拒绝已消除。
- DOCX 域指令、删除域指令和删除文本曾被当成显示文本解析；现在安全忽略这些非显示内容并输出兼容性诊断，另有 22 个语料因此可正常打开。
- 上传检查页的对象详情展开曾改变 Viewer 高度、触发适应窗口重绘并清除选中对象；现在桌面工作区尺寸稳定，且相同缩放不会重复渲染。

核心行为均有 Rust 或 JS 回归测试；上传页问题另经真实浏览器回归验证。TDF `Basic.ods` 仍按预期触发 `OBJECT_LIMIT`：该文件声明约 419 万个非空重复单元格，超过默认 100 万对象预算；当前 SDK 选择受控拒绝，而不是无界展开。

## 当前主要兼容缺口

427 个结构化拒绝按代码主要分布为：`FORMAT_INVALID` 270、`PACKAGE_ZIP_INVALID` 85、`OBJECT_LIMIT` 20、`LAYOUT_BUDGET_EXCEEDED` 15、`UNSUPPORTED_FORMAT` 11、`XML_INVALID` 8、`PACKAGE_ENCRYPTED` 7。数量较多的当前兼容缺口包括：

- DOCX 运行文本边界处理：23 个 `unexpected text outside a DOCX text run`；
- 嵌套 DOCX 段落 9 个、负上页边距 9 个、嵌套表格 3 个；
- 非 `Pictures/` 路径下的 ODF 嵌入图片；
- 部分 ODT 嵌套表格单元格段落和缺少主页面的测试输入。

当前全量门禁不是绿色：`sd/qa/unit/data/fdo84043.odp` 暴露 1 个 `CORE_PROTOCOL_INVALID`；OASIS、TDF 旧集合当前打开数分别为 29 和 269，低于已有 34 和 305 基线；LibreOffice 的 `tdf111798.odp`、`tdf92076.odp`、`tdf127090.odp`、`tdf128651_CustomShapeUndo.odp` 暴露 4 个首对象中心点命中失败。格式损坏、加密和路径逃逸类拒绝属于安全预期；其余结果仍需结合每个上游测试的原始预期逐项分类，不能只优化“打开率”。

## 复现方式

快速基线会下载 15 个固定文件、逐个校验 SHA-256；其中新增 6 个 LibreOffice 文件覆盖 DOCX、ODT、XLSX、ODS、PPTX 和 ODP：

```sh
npm run test:corpus
```

全量基线使用 blobless sparse checkout 固定到上述四个 commit，扫描产品支持的六种扩展名；`--quiet` 只输出失败和最终汇总：

```sh
npm run test:corpus:full
```

全量命令同时执行“只允许改善、不允许回退”的基线门禁：总数必须保持 4,736；Microsoft、OASIS、TDF、LibreOffice 分别不得低于 651、34、305、3,336 个成功打开。任一来源回退、出现未分类失败或任何 `CORE_*` 内部协议/加载拒绝都会返回非零退出码。固定 commit 的缓存工作区如有修改、删除或未跟踪文件，将被丢弃并重新获取，避免本地污染改变统计。

缓存保存在 `.cache/official-corpus/`，本次实际占用约 809 MiB，其中 LibreOffice sparse checkout 约 499 MiB。机器可读报告写入 `output/official-corpus-report.json`，记录每个输入（包括结构化拒绝）的 SHA-256，并附带 MSO 图片级映射和哈希；两者都由 `.gitignore` 排除，不会进入 SDK 发布包。

## 验证边界

批处理脚本在 Node 中使用 inline Wasm，只验证解析和对象协议，不进行 Canvas 绘制。浏览器端 Worker、Wasm、渲染、命中和诊断展示由上传检查页验证。

四个上游来源都没有覆盖六种格式的完整官方视觉 oracle、对象边界 golden data 或统一合格规则。LibreOffice 的 32 张 MSO bitmap 只覆盖 4 个 PowerPoint 文字 fitting 场景。因此当前结果不能表述为通过 ISO、ECMA 或 OASIS 一致性认证；完整视觉保真仍需建立 OfficeViewer 自有、带授权且固定字体/系统/DPI 的图像基线。

多层准确性评估、逐文件 `AccuracyReport`、语料覆盖门禁和变形生成器现已定义在 [准确性测试体系](../ACCURACY.md)。本页的上游语料仍只是兼容性输入，不能自动升级为独立 oracle；接入准确性 suite 前，必须补充经过审阅的结构、几何、文本布局、像素和对象映射期望。
