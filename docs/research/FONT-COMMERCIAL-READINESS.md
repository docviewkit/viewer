# OfficeViewer 字体能力与商用就绪度研究

> 观察日期：2026-07-16（Asia/Shanghai）
> 范围：当前仓库源码/测试，以及标准组织、Microsoft、浏览器与商业产品的一手资料。许可内容仅作风险识别，不构成法律意见。

## 结论

“字体相关功能是 SDK 是否具备商用能力的关键”方向正确，但必须改成更严格的表述：

> **字体是 Office 渲染商用化的 P0 必要条件和高杠杆质量门槛，但不是充分条件；真正要建设的是字体感知排版基础设施，而不只是注册或下载字体。**

字体会决定 glyph 是否存在、字形宽度、行高、换行、分页、表格高度、后续对象位置、首屏性能、跨平台一致性，也带来安全和许可风险。[Microsoft 明确说明](https://support.microsoft.com/en-US/publisher/substitute-the-missing-fonts-in-your-publication)，缺字体替换可能改变行、列、分页、行距和连字符；[嵌入字体](https://support.microsoft.com/en-US/Office/fonts/benefits-of-embedding-custom-fonts)的核心价值正是保持字体和布局。

但字体正确仍不能单独复刻 Word：兼容模式、keep/widow-orphan、分节、表格 autofit、浮动对象、字段、断词、打印机度量等都会影响版式。参见 [Word 分页控制](https://support.microsoft.com/en-us/word/line-and-page-breaks)、[compatibility mode](https://learn.microsoft.com/en-us/office/compatibility/manage-compatibility-mode-for-office)和 [usePrinterMetrics](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.useprintermetrics)。

当前判断：**0.2.0 已补齐字体参与 DOCX 分页、精确 face、懒加载 Provider、Engine 缓存、策略、完整性与诊断这组 P0 基础；但字形覆盖/cluster shaping、真实多浏览器与 Word 参考金标、许可运营仍未达到“全面商用就绪”。**

## 当前仓库的决定性证据

### 已完成的 P0 基础

- 公共 API 仍只接收字体二进制，不接受文档控制的 URL；网络由宿主 `fontProvider` 掌握，接口保持为 `fontProvider` 与 `fontPolicy`。
- family/style/weight/stretch 精确到 face；DOCX 会按 ASCII/HAnsi/EA/CS、theme 和 Unicode script 选择 run 字体。
- `open` 已形成 preliminary parse → demand → resolve/register → Canvas advance → DOCX final layout；实测宽度会改变换行和页数，固定 1em/0.55em 仅在缺失度量时回退。
- 支持 deterministic 与 local-first；deterministic 不探测环境字体，后者只按文档实际需求确认 exact local face，命中后不会调用 Provider。embedded、static host、provider、browser 来源可在公开 `fontRuns` 中观察。
- Provider 具备 Engine 级 face/content/in-flight 去重、LRU 字节预算、超时、取消、SHA-256、缺失/失败/完整性诊断；静态资产也可声明 `sha256`。
- 支持 DOCX 嵌入字体反混淆以及 PPTX/ODF 文档字体；默认每文档独立 Worker，关闭即释放字体 realm。
- OVFM 度量表在 JS 与 Wasm 两端严格校验长度、record、Unicode scalar 和有限浮点数，并受字体/对象/XML record 预算限制。

### 仍需继续收口的商用门槛

1. **advance 不是完整 shaping。** 当前表按 face/codepoint 测量孤立字符；kerning、ligature、combining cluster、Arabic/Indic shaping、variation axis 和竖排度量还需要 cluster-aware shaping/measure 方案。
2. **缺 glyph 仍不够可审计。** Provider 会收到实际 Unicode scalars，但浏览器对单个缺字的 fallback、synthetic style 与覆盖范围尚未形成逐 glyph 的公开 trace。
3. **真实视觉证据仍不足。** 自动化覆盖策略、缓存、Worker 握手和分页变化，但仍需提交 Chrome/Edge/Firefox/Safari × Windows/macOS/Linux、CJK/Arabic/Indic/emoji/variable/TTC 的固定字体金标，并与 Word/PDF 参考比较。
4. **许可不能靠代码自动解决。** `sha256` 能固定二进制，不能证明 Web/应用分发权；字体清单仍需要版本、来源、授权范围和审计台账。

这证明项目不应把“字体注入提前到 SDK 初始化”当成根治。静态字体可在 Engine 初始化准备和缓存，但每个文档仍必须先识别需求，并在**首次正式布局之前**解析、传入 Worker、注册和读取度量。

## 平台和 Word 规则

- [CSS Font Loading](https://www.w3.org/TR/css-font-loading/)允许用二进制创建 FontFace；[Canvas 标准](https://html.spec.whatwg.org/multipage/canvas.html)会使用 CSS 字体源。因此非系统字体可以被 Canvas 使用。
- Document 与 Worker 有各自的 FontFaceSet，Worker 字体集初始为空；页面注册不等于 Worker 可用。字体仍应由 SDK 在实际文档 Worker 内注册。
- FontFaceSet.check 即使 family 不存在、浏览器会 fallback 也可能返回 true，不能证明本地字体、版本或 glyph coverage。
- [CSS Fonts 4](https://www.w3.org/TR/css-fonts-4/)允许 local()，但 UA 可为隐私隐藏本地字体；同名也不代表同二进制、版本或覆盖。因此“名称精确匹配就永远本地优先”并不正确。
- SDK 不使用 Local Font Access，也不枚举用户安装字体。`local-first` 只通过 `FontFace(..., local(...))` 尝试文档实际请求的 exact face，因此不需要本地字体列表权限；需要稳定二进制和度量时仍由宿主字体或 Provider 提供。
- Word 一个 run 可同时使用 ASCII、High ANSI、East Asian、Complex Script 字体；应按字符和脚本选 face，参见 [Microsoft 实现说明](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/dcf1caba-49a9-40e3-ba36-32b9e205434f)。

## 推荐最小架构

~~~text
parse document
→ collect family/face/script/codepoint demands
→ resolve embedded/host/cache/local/service by explicit policy
→ verify hash, limits and license
→ transfer bytes to document Worker
→ FontFace.load + Worker FontFaceSet
→ font-aware shaping, line breaking and pagination
→ stable scene and Canvas render
~~~

责任边界：宿主提供静态字体或实现 fontProvider；SDK Engine 负责策略、缓存、去重、完整性、预算和诊断；Core 输出字体需求并用真实度量排版；Worker 在自己的 realm 注册；文档内容不得直接决定网络 URL。

公共 API 应保持小，例如只增加 fontProvider 与 fontPolicy，不暴露 manifest、Worker registry 或内部缓存：

~~~ts
createOfficeEngine({ fonts, fontProvider, fontPolicy: "deterministic" });
~~~

两种明确策略：

- **deterministic**：embedded → host/pinned cache → pinned service → 固定度量兼容 fallback；不读取或探测环境字体，用于验收、转换和回归。
- **local-first**：embedded → host/cache → local → service → fallback；用于快速预览和离线，但结果必须标记为环境相关。

缓存键必须是 hash + face identity，而不是 family。清单至少记录 family/full/PostScript name、style、weight、stretch、faceIndex、Unicode coverage、version/fontRevision、SHA-256、size、format、来源和许可。manifest 用 version/ETag，内容寻址字体可长期 immutable；Engine 做内存预算和 in-flight 去重。

## 安全、许可与可观测性

- 字体是复杂的不可信二进制。[CSS Fonts 4](https://www.w3.org/TR/css-fonts-4/)提示恶意字体可利用脆弱平台；Chromium/Firefox 使用 [OpenType Sanitizer](https://chromium.googlesource.com/external/github.com/khaledhosny/ots/)。P0 应限制格式、大小、数量、并发和超时，校验 magic/结构/hash，并加入畸形字体语料与 Worker 崩溃隔离测试。
- provider 只能访问宿主配置/allowlist 的 HTTPS 服务；不跟随文档 URL，不把 token 写入 URL/日志；处理 CORS 与 CSP connect-src/font-src/worker-src，参见 [CSP 3](https://www.w3.org/TR/CSP/)。
- [OpenType fsType](https://learn.microsoft.com/en-us/typography/opentype/spec/os2)区分 Installable、Restricted、Preview & Print、Editable、No Subsetting 等；但它只表达文档嵌入限制，不等于服务器/应用分发许可。[Microsoft FAQ](https://learn.microsoft.com/en-us/typography/fonts/font-faq)明确指出文档嵌入权不授予应用嵌入或把 Windows 字体复制到 Web 服务的权利。
- 每个请求应输出 requested face/codepoints、resolved face、source、hash/version、substitution reason、missing glyph、synthetic style、cache hit、耗时、字节和 integrity/license 结果；不得记录文档原文或凭据。

## 商业产品共同模式

- [Apryse](https://docs.apryse.com/web/faq/self-serve-substitute-fonts)：fonts.json 包含 coverage、family、variants 和有序匹配，按需下载；可在本地替代与统一后端之间选择。
- [Microsoft Cloud Fonts](https://support.microsoft.com/en-us/office/fonts/cloud-fonts-in-office)：由受控服务自动下载并在 Office apps 间复用；与文档嵌入并存。
- [ONLYOFFICE](https://helpcenter.onlyoffice.com/docs/installation/docs-community-install-fonts-windows.aspx)：新增字体后扫描并生成 AllFonts.js 和 font_selection.bin 度量缓存，说明成熟系统不只注册 FontFace。
- [Collabora](https://sdk.collaboraonline.com/CO-SDK-manual.pdf)：远程 HTTPS 清单、stamp/version 更新、metric-compatible fallback；缺字体可 report/log。
- [Aspose](https://docs.aspose.com/words/net/specifying-truetype-fonts-location/)与 [Syncfusion](https://help.syncfusion.com/document-processing/word/conversions/word-to-pdf/net/fallback-fonts-word-to-pdf)：source priority、按需 stream、跨文档缓存，并明确区分 font substitution 与 glyph fallback。

共同结论：清单 + 懒加载是成熟实践，但成熟能力的单位是“字体子系统”，包括布局前解析、face/coverage、缓存、诊断、安全和许可，而不是一个下载函数。

## P0 / P1 路线

**P0 商用门槛（前三项基础已在 0.2.0 实现，后续继续做深）**

1. 将 open 改为 parse → font demand → resolve/register → font-aware layout → scene；移除 DOCX 固定字符宽度作为正式分页依据。
2. 建立 family/style/weight/stretch/script/codepoint 精确模型；正确处理 OOXML theme 与 ASCII/HAnsi/EA/CS；分离 substitution 和 per-glyph fallback。
3. 提供 deterministic/local-first 策略，以及宿主 fontProvider；按 hash/version/face 缓存和并发去重。
4. 继续补齐并发上限、格式预检和许可治理；大小/超时/hash/网络地址边界已落地。
5. 在现有 face/source/provider 诊断之上补齐 missing glyph、synthetic style、cache hit/version/license trace。
6. 建立真实 Chrome/Edge/Firefox/Safari、Windows/macOS/Linux、Worker/inline、CJK/Arabic/Indic/emoji/variable/TTC、embedded/host/local/service 的 golden；与 Word/PDF 参考比较断行、分页、对象和像素。

**P1**

1. Provider 持久缓存、离线字体包、回滚和租户配额；不引入 Local Font Access 权限依赖。
2. variable axes、TTC face、颜色字体、OpenType features、竖排度量；合法前提下子集化。
3. 更完整的度量兼容 fallback、字体清单运维和许可证台账；按支持的 WebView/威胁模型决定是否增加独立 sanitizer。

字体项目之外仍有独立 P0：Word 兼容/表格/浮动/字段/分页规则、格式覆盖、资源安全、大文档性能、API/浏览器支持契约、隐私部署和 SLA。

最终验收原则：

> **如果字体尚未参与分页，就不能称为字体保真；如果缺字与替换不可审计，就不能称为商用；如果字体服务没有版本、哈希和许可边界，就不能上线。**
