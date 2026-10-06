# ODF 官方测试资产调研

调研日期：2026-07-15

## 结论

能够找到由 OASIS ODF TC、已关闭的 OASIS OIC TC 和 The Document Foundation ODF Toolkit 项目公开维护的测试资产，但**没有找到一个现代、完整、带统一预期结果或渲染基准，并由 OASIS 声明为 ODF 官方一致性认证套件的文件集**。

可以直接用于 OfficeViewer 的资产分为四层：

1. [OASIS ODF 1.4 官方 Schema](https://docs.oasis-open.org/office/OpenDocument/v1.4/os/schemas/)：用于 XML 词汇和包内清单结构验证，不验证视觉渲染。
2. [OASIS ODF TC 官方仓库](https://github.com/oasis-tcs/odf-tc)：包含针对新 ODF 特性制作的 ODT、ODS、ODP 测试文档，以及历史 OpenFormula 测试材料；没有覆盖全部规范条款的统一 manifest 或视觉 oracle。
3. [ODF Toolkit v0.13.0 测试语料](https://github.com/tdf/odftoolkit/tree/v0.13.0/odfdom/src/test/resources/test-input)：包含 338 个与本产品格式范围直接相关的 ODT、ODS、ODP 文件，适合批量健壮性和功能覆盖测试；它是实现回归语料，不是 OASIS 一致性套件。
4. [ODF Validator](https://odftoolkit.org/conformance/ODFValidator.html)：可递归执行 Schema 和部分包一致性检查；官方说明同时明确了其检查对象是 XML、manifest、签名等静态结构，不包含 Office 页面布局和渲染结果。

因此，OfficeViewer 应使用这些资产做“官方来源语料上的批量解析、诊断、对象协议和渲染执行验证”，不能据此宣称“通过 OASIS 官方 ODF 渲染一致性测试”。OASIS OIC 的报告也明确指出，XML 有效性是一致性的必要但不充分条件，并把视觉外观、文档结构、链接和嵌入对象列为另外的互操作维度：[State of Interoperability v2.0](https://docs.oasis-open.org/oic/StateOfInterop/v2.0/cnd01/StateOfInterop-v2.0-cnd01.html#_ODF_Conformance)。

## 资产分类

| 资产 | 发布方 | 性质 | 可否自动批量获取 | 对 OfficeViewer 的价值 |
| --- | --- | --- | --- | --- |
| ODF 1.4 标准和 Relax NG Schema | OASIS ODF TC | 规范与结构约束 | 可以，固定官方 ZIP/RNG URL | 格式版本、XML 和 manifest 基线 |
| ODF 1.3 特性测试文档 | OASIS ODF TC | 标准制定过程中的特性语料 | 可以，Git 固定 commit | ODT/ODS/ODP 新特性解析与诊断 |
| ODF2HTML 输入及参考 HTML | OASIS ODF TC | TC 自身 XSLT 回归输入 | 可以，Git 固定 commit | ODT/ODS 复杂输入；参考 HTML 不能当原生渲染基准 |
| OpenFormula 历史测试套件 | OASIS ODF TC 仓库收录 | 公式求值测试材料 | 可以，Git 或单文件下载 | 公式缓存值、公式诊断；不等于布局测试 |
| OIC Interop Advisories | 已关闭的 OASIS OIC TC | 小型互操作建议及零散测试文件 | 当前不适合批量下载 | 历史边界场景和规范歧义回归 |
| ODF Toolkit 测试输入 | The Document Foundation | Java ODF 实现回归语料 | 可以，固定 release tag | 大规模 ODT/ODS/ODP 健壮性和能力覆盖 |
| ODF Validator fixtures | The Document Foundation | 正例、反例、加密及包结构回归 | 可以，固定 release tag | 拒绝路径、错误码和结构化诊断 |

## 1. OASIS 是否发布过完整的一致性测试套件

### OIC TC 的目标与实际交付

[OASIS OIC TC charter](https://www.oasis-open.org/committees/oic/charter.php)曾把“选择并发布 ODF interoperability test corpus”列为工作范围和计划交付物，同时明确该 TC 不负责充当认证机构，也不负责编写或分发测试软件。

但 [OIC TC 最终归档页](https://www.oasis-open.org/committees/tc_home.php?wg_abbrev=oic)显示：该 TC 于 2013 年关闭，正式列出的主要成果是两份互操作报告、ODF 1.1 分析、若干 Interop Advisories 和一个 profile；技术成果入口只列出“various advisories and small test files”，没有列出覆盖 ODF 全部条款的最终 test-suite 发布包。

[Interop Advisories 归档页](https://wiki.oasis-open.org/oic/InteropAdvisories)列出了 9 个 candidate advisories 和 1 个 draft advisory，主题包括非 inline 文本、表单 fallback、图表单元格范围、删除线、SHA-256 URI、表格行列、数字签名命名空间和保护密钥。页面工作流多次把这些零散文件描述为未来 “OIC Test Suite” 的输入，而不是已经完成的正式套件。

截至调研日期，Advisories 页面中的旧 SVN 下载链接会重定向到 OASIS issue tracker 登录页，无法匿名递归列目录或批量抓取。因此这些历史文件可以作为人工补充来源，但不应成为当前 CI 的基础依赖。

### 为什么 Schema 验证不等于 SDK 一致性

OIC 的 [State of Interoperability v2.0](https://docs.oasis-open.org/oic/StateOfInterop/v2.0/cnd01/StateOfInterop-v2.0-cnd01.html#_ODF_Conformance)明确指出，现有 validator 主要检查 XML validity；这是一致性的必要条件，但不足以验证应用的一致性。该报告还说明，ODF 渲染的主要预期行为来源是已发布标准，并把视觉外观、可编辑结构、链接、嵌入对象、元数据和扩展保留分别列为互操作质量。

结论：目前没有证据支持把任何单个公开 ODF 文件目录称为“完整 OASIS 官方渲染一致性套件”。

## 2. OASIS ODF 1.4 Schema：官方结构验证资产

[ODF 1.4 OASIS Standard 发布目录](https://docs.oasis-open.org/office/OpenDocument/v1.4/os/)提供固定 ZIP，包含标准正文及相关资产：

```text
https://docs.oasis-open.org/office/OpenDocument/v1.4/os/OpenDocument-v1.4-os.zip
```

[官方 schemas 目录](https://docs.oasis-open.org/office/OpenDocument/v1.4/os/schemas/)提供：

- [`OpenDocument-v1.4-schema.rng`](https://docs.oasis-open.org/office/OpenDocument/v1.4/os/schemas/OpenDocument-v1.4-schema.rng)
- [`OpenDocument-v1.4-manifest-schema.rng`](https://docs.oasis-open.org/office/OpenDocument/v1.4/os/schemas/OpenDocument-v1.4-manifest-schema.rng)
- [`OpenDocument-v1.4-dsig-schema.rng`](https://docs.oasis-open.org/office/OpenDocument/v1.4/os/schemas/OpenDocument-v1.4-dsig-schema.rng)

这些文件可验证解包后的 `content.xml`、`styles.xml`、`META-INF/manifest.xml` 和签名 XML 是否满足相应词汇约束。它们不包含页面截图、文本测量结果、对象边界或命中测试预期。

### ODF Validator

[ODF Validator 官方说明](https://odftoolkit.org/conformance/ODFValidator.html)记录了以下可自动化能力：

- 接受文件或目录，`-r` 可递归处理；
- 按文档声明选择 ODF 版本，命令行明确列出 1.0、1.1、1.2 和 1.3；
- 验证 `content.xml`、`styles.xml`、`meta.xml`、`settings.xml`；
- 验证 manifest、数字签名及嵌入式 ODF 对象；
- 可区分普通、strict、conforming 和 extended conformance 的部分静态规则。

该版本的文档没有列出 ODF 1.4 命令行选项。因此工程上应把 ODF Validator v0.13.0 用于其明确支持的版本，并对 ODF 1.4 直接使用 OASIS 1.4 RNG 做单独结构检查，不能静默把 1.3 Schema 当作 1.4 Schema。

Schema/validator 只作为测试工具运行，不应加入 OfficeViewer 正式 Wasm/JavaScript 运行时，因此不违反产品“无第三方运行时代码”的约束。

## 3. OASIS ODF TC 官方特性测试文档

[OASIS ODF TC 官方仓库 README](https://github.com/oasis-tcs/odf-tc#overview)明确把“根据规范新增 ODF 特性提供测试文档”列为仓库目的之一。为了可复现，建议固定到本次核验的 commit：

```text
16a59d945875bd74834e82e166bebded316478da
```

### ODF 1.3 testfiles

在固定 commit 的 [`src/test/resources/odf1.3/testfiles`](https://github.com/oasis-tcs/odf-tc/tree/16a59d945875bd74834e82e166bebded316478da/src/test/resources/odf1.3/testfiles) 中实测枚举得到：

| 格式 | 数量 | 是否属于当前产品范围 |
| --- | ---: | --- |
| ODT | 14 | 是 |
| ODS | 17 | 是 |
| ODP | 3 | 是 |
| ODG | 5 | 否 |
| OTM | 1 | 否 |
| FODG | 1 | 否 |

与 OfficeViewer 当前三种 ODF 格式直接相关的文件共 34 个。文件主题包括公式与数值格式、图表坐标区域、命名范围、页面打印、页眉页脚、上下文段落间距、背景/填充、图片 MIME 类型和 presentation 图形属性。

这些文件适合检查：

- SDK 是否按内容识别 ODT、ODS、ODP；
- 已支持对象是否能稳定形成 scene 和源映射；
- 当前不支持的图表、页眉页脚、复杂背景等是否给出明确诊断；
- 每份文档是否在时间、内存和 Worker 生命周期预算内结束。

同一目录没有统一的 expected-result manifest、参考截图或对象边界清单，因此不能单凭“成功打开”判断该特性已正确实现。

### ODF2HTML 回归输入

固定 commit 的 [`src/test/resources/html-export/input`](https://github.com/oasis-tcs/odf-tc/tree/16a59d945875bd74834e82e166bebded316478da/src/test/resources/html-export/input)包含 6 个 ODT、1 个 ODS 和 1 个 ODG 打包文档。仓库 README 说明这些文件用于 ODF TC 自身的 ODF-to-HTML XSLT 回归，并将输出与 [`references/xslt-html`](https://github.com/oasis-tcs/odf-tc/tree/16a59d945875bd74834e82e166bebded316478da/src/test/resources/odf1.4/references/xslt-html)比较。

它们可以作为 OfficeViewer 的复杂输入语料，但参考 HTML 是特定 XSLT 的输出，不是 Office 原生布局的 golden image，不能直接用于 Canvas 像素比对。

### ODF 1.5 草案文件

仓库还包含 [`src/test/resources/odf1.5/testfiles`](https://github.com/oasis-tcs/odf-tc/tree/16a59d945875bd74834e82e166bebded316478da/src/test/resources/odf1.5/testfiles) 下的 7 个 ODS。该目录明确属于尚未发布的 ODF 1.5 工作材料，应作为未来格式或前向兼容测试单独分组，不能计入 ODF 1.4 合格率。

### 自动化获取

推荐使用 blobless sparse checkout，只取需要的测试目录，并固定 commit：

```bash
mkdir -p .cache

git clone \
  --filter=blob:none \
  --no-checkout \
  https://github.com/oasis-tcs/odf-tc.git \
  .cache/odf-tc

git -C .cache/odf-tc sparse-checkout init --cone
git -C .cache/odf-tc sparse-checkout set \
  src/test/resources/odf1.3/testfiles \
  src/test/resources/html-export/input
git -C .cache/odf-tc checkout --detach \
  16a59d945875bd74834e82e166bebded316478da
```

完整固定快照也可通过以下 GitHub archive URL 获取，但体积明显大于 sparse checkout：

```text
https://github.com/oasis-tcs/odf-tc/archive/16a59d945875bd74834e82e166bebded316478da.zip
```

### 许可证和再分发边界

[仓库 LICENSE.md](https://github.com/oasis-tcs/odf-tc/blob/16a59d945875bd74834e82e166bebded316478da/LICENSE.md)说明该仓库是 OASIS TC 工作仓库，受 OASIS IPR Policy 和 RF on Limited Terms 模式治理；它不是 Apache、MIT 或普通 OASIS Open Repository。仓库 README 称内容公开且可供使用，但须遵守这些 OASIS 政策。

[OASIS IPR Policy](https://www.oasis-open.org/policies-guidelines/ipr/)允许复制和提供正式 OASIS deliverable，并允许制作帮助解释或实现标准的派生材料，但要求保留版权和许可段落，且原则上不得修改标准文档本身。

工程建议：CI 从固定 commit 下载并缓存；不要把 OASIS TC 全部测试文件直接打入 npm 发布包。若需要长期复制到本仓库或对外再分发，应保留来源、commit 和相关 OASIS notice，并先做许可证审核。

## 4. OpenFormula 历史测试套件

OASIS TC 仓库的 [`formulas-testsuite-generator`](https://github.com/oasis-tcs/odf-tc/tree/16a59d945875bd74834e82e166bebded316478da/src/test/resources/odf1.2/tools/formulas-testsuite-generator)包含：

- 一份带测试注释的 ODF 1.2 OpenFormula 规范 ODT；
- 一份 2006 年生成的 [`openformula-testsuite-20060724.ods`](https://github.com/oasis-tcs/odf-tc/blob/16a59d945875bd74834e82e166bebded316478da/src/test/resources/odf1.2/tools/formulas-testsuite-generator/examples/openformula-testsuite-20060724.ods)；
- 将规范中的公式测试提取成 ODS 的生成器。

[该目录 README](https://github.com/oasis-tcs/odf-tc/blob/16a59d945875bd74834e82e166bebded316478da/src/test/resources/odf1.2/tools/formulas-testsuite-generator/README.md)明确说明，该 generator 在 ODF 1.3 期间没有被使用，并把现存 ODS 称为 2006 年的 generated example。因此它是有价值的历史公式语料，但不是 ODF 1.3/1.4 的完整公式认证套件。

OfficeViewer 当前读取公式缓存值而不承担完整公式计算，因此该 ODS 可用于测试：

- 文件能够被安全解析；
- 公式和缓存值不会破坏单元格源映射；
- 未实现的重计算能力产生准确诊断。

不应把公式计算结果正确率计入当前渲染 SDK 的支持率，也不应因为无法重算公式而把文件归类为解析失败。

生成器子目录的 [`COPYING`](https://github.com/oasis-tcs/odf-tc/blob/16a59d945875bd74834e82e166bebded316478da/src/test/resources/odf1.2/tools/formulas-testsuite-generator/generator/COPYING)是 GPL-2.0。除非确有生成新测试的需要，不应把该 generator vendoring 到 OfficeViewer；对生成的 ODS 单独再分发前仍应核对其具体版权来源。

## 5. OASIS 发布的规范 ODT：规范样例，不是测试套件

ODF 1.4 的四个标准部分本身同时以 ODT 发布：

- [Part 1 Introduction ODT](https://docs.oasis-open.org/office/OpenDocument/v1.4/os/part1-introduction/OpenDocument-v1.4-os-part1-introduction.odt)
- [Part 2 Packages ODT](https://docs.oasis-open.org/office/OpenDocument/v1.4/os/part2-packages/OpenDocument-v1.4-os-part2-packages.odt)
- [Part 3 Schema ODT](https://docs.oasis-open.org/office/OpenDocument/v1.4/os/part3-schema/OpenDocument-v1.4-os-part3-schema.odt)
- [Part 4 Formula ODT](https://docs.oasis-open.org/office/OpenDocument/v1.4/os/part4-formula/OpenDocument-v1.4-os-part4-formula.odt)

这些是 OASIS 自己发布的复杂、真实 ODT 文档，可用于大文件、表格、图片、目录、交叉引用和分页的 smoke test。但它们的权威内容是规范文本，不是“渲染测试输入 + 预期画面”，也不覆盖 ODS 和 ODP。

版权方面应遵守文件内的 OASIS notice：保留原文和版权段落，不修改后冒充官方版本。测试阶段从 OASIS URL 下载最稳妥，不把它们复制进 SDK 发行包。

## 6. ODF Toolkit 官方实现回归语料

ODF Toolkit 是 The Document Foundation 维护的 Java ODF 项目；[项目首页](https://odftoolkit.org/)将 ODFDOM、ODF Validator 和相关工具列为项目组件。建议固定到 [v0.13.0 release](https://github.com/tdf/odftoolkit/releases/tag/v0.13.0)，解引用 commit：

```text
b926a6134a2fee782076500dfc02c47c2d651cff
```

在该 tag 的 [`odfdom/src/test/resources/test-input`](https://github.com/tdf/odftoolkit/tree/v0.13.0/odfdom/src/test/resources/test-input) 中实测枚举得到：

| 格式 | 数量 |
| --- | ---: |
| ODT | 210 |
| ODS | 112 |
| ODP | 16 |
| 与本产品直接相关的合计 | 338 |

目录还包含模板、ODG/ODC/ODF、XML、JSON 和图片等辅助资产。其 ODT/ODS/ODP 文件覆盖普通内容、列表、表格、样式、图片、图表、公式、页眉页脚、密码保护、损坏 MIME 和多种生产者生成的文档。准确覆盖范围应以每个测试类和文件本身为准，不能只根据文件名推断通过条件。

[`validator/src/test/resources`](https://github.com/tdf/odftoolkit/tree/v0.13.0/validator/src/test/resources)另有 10 个 ODT 和 2 个 ODS fixture，包含有效、无效、加密、MathML 和扩展标记场景；对应 [`ValidTest`](https://github.com/tdf/odftoolkit/blob/v0.13.0/validator/src/test/java/org/odftoolkit/odfvalidator/ValidTest.java)等测试代码给出了 ODF Validator 自身的预期。但这些预期是 validator 的结果，不是 OfficeViewer 的页面渲染预期。

### 自动化获取

```bash
mkdir -p .cache

git clone \
  --depth 1 \
  --branch v0.13.0 \
  --filter=blob:none \
  --sparse \
  https://github.com/tdf/odftoolkit.git \
  .cache/odftoolkit

git -C .cache/odftoolkit sparse-checkout set \
  odfdom/src/test/resources/test-input \
  validator/src/test/resources
```

也可下载固定 tag 的源码归档：

```text
https://github.com/tdf/odftoolkit/archive/refs/tags/v0.13.0.zip
```

### 许可证

ODF Toolkit 仓库使用 [Apache License 2.0](https://github.com/tdf/odftoolkit/blob/v0.13.0/LICENSE)，并提供独立的 [NOTICE](https://github.com/tdf/odftoolkit/blob/v0.13.0/NOTICE)。Apache-2.0 允许复制和再分发，但必须满足许可证、修改说明和 NOTICE 等条件；仓库还包含 OASIS Schema、MathML 等各自 notice。

推荐仍然由 CI 从固定 tag 下载，不把 338 个文档放入 OfficeViewer npm 产物。若复制进本仓库，应同时保存上游 LICENSE、NOTICE、tag/commit 和文件 SHA-256，并检查个别历史文档是否带有额外第三方内容。

## 7. 对 OfficeViewer 的验证方案

这些资产没有统一视觉 oracle，因此测试必须分层，不能把“调用 `open()` 未抛异常”当作通过。

### 第一层：资产准入和结构基线

- 下载源固定到 tag/commit，保存 SHA-256 和来源路径；
- OASIS TC、ODF Toolkit、OIC 历史资产分别统计，不混成一个“官方套件”；
- 使用 OASIS RNG 或 ODF Validator 记录输入是否结构有效；
- 结构无效、加密、损坏和主动内容作为预期拒绝/诊断样本，不计为普通解析失败；
- ODF 1.5 草案、ODG 和模板格式单独分组，不计入 ODT/ODS/ODP 支持率。

### 第二层：安全和生命周期

对全部 ODT、ODS、ODP 执行：

- 按文件内容识别格式，不依赖扩展名；
- 每份文档必须在固定时间、内存、解压和 XML 预算内结束；
- 不能崩溃、卡死 Worker、长期占用主线程或泄漏文档资源；
- 外部资源、脚本、宏、异常图片、加密包和损坏结构必须拒绝或给出结构化诊断。

### 第三层：解析和对象协议

对成功打开的文档检查：

- scene unit/object ID 唯一，父子关系无环，数值边界有限；
- ODT 段落/文本范围、ODS 工作表与行列、ODP 对象 ID 等源映射存在且格式正确；
- 关键内容不能静默丢失；未实现特性必须产生 `unsupported` 或 `approximation` 诊断；
- 每个可见对象中心点执行 hit test，并验证命中对象属于同一 unit；
- 所有 unit 在有界 viewport 下至少执行一次 Canvas 渲染。

### 第四层：结果分类

每个文件只进入以下一种顶层结果：

1. 完整渲染；
2. 带明确诊断的近似渲染；
3. 预期拒绝；
4. SDK 缺陷。

同时记录格式、来源、文件大小、打开耗时、unit/object 数、成功绘制对象数、诊断码、首个错误和资源预算结果。OASIS TC 特性文件与 ODF Toolkit 大语料应分别出报表，避免数量较大的 ODT 文件掩盖 ODP 或官方 TC 特性样本的问题。

### 第五层：视觉保真

上述官方资产没有适用于 OfficeViewer 的原生页面截图或对象边界 golden result。视觉层必须另建项目自己的授权基准：

- ODP、ODT、ODS 分别固定 Microsoft PowerPoint、Word、Excel 版本，以及操作系统、字体、DPI 和导出设置；
- 使用上述对应 Office 应用为代表性 ODT/ODS/ODP 生成参考图；
- 对动态字体替换和分页差异设置可解释阈值；
- 人工审定对象层级、源映射和诊断是否正确。

LibreOffice、WPS 和其他渲染器只能作为解析/渲染差分研究参考，全面禁止作为 native/release 验收金标。该结果应称为“OfficeViewer 自有视觉回归基线”，不能称为 OASIS 官方 conformance 结果。

## 8. 推荐落地顺序

1. 先接入 ODF Toolkit v0.13.0 的 338 个 ODT/ODS/ODP，建立批量执行器和报告格式。
2. 再加入 OASIS ODF TC commit `16a59d...` 的 34 个 ODF 1.3 ODT/ODS/ODP 特性文件，单独标为标准组织语料。
3. 加入 ODF Validator 的 12 个正反例 fixture，重点测试拒绝路径和诊断。
4. 使用 OASIS ODF 1.4 Schema 做结构基线；不要把 Schema 通过率等同于渲染通过率。
5. 最后挑选每种格式的代表文件建立自有视觉和对象映射 golden tests。

## 9. 可对外使用的准确表述

可以表述为：

> OfficeViewer 已使用固定版本的 OASIS ODF TC 特性测试文档和 The Document Foundation ODF Toolkit 回归语料，对 ODT、ODS、ODP 执行批量解析、安全预算、兼容性诊断、对象协议和渲染执行验证；XML 结构另使用 OASIS ODF Schema 验证。

不能表述为：

> OfficeViewer 通过了 OASIS 官方 ODF 渲染一致性认证。

除非未来 OASIS 发布带完整测试规范、预期结果和合格判定流程的正式套件，否则后一表述没有充分证据。
