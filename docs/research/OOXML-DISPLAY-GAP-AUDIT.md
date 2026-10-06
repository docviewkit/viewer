# OOXML 只读显示差距复核

日期：2026-09-23。初始审计基线：`2042470`；本批实现基线：`1fe8f58`。下述阶段记录保留当时状态，v0.2.74 的合并范围和发布边界见第 14 节，实际发布状态以对应 GitHub Release 为准。

范围：DOCX / XLSX / PPTX 的内容显示、静态外观、分页与打印版面；排除编辑、表单输入、动画、切换与自动播放。批注、修订标记、控件的静态外观仍属于显示能力，但与交互功能分别评价。初始审计按标准词汇与源码路径核对；后续实现与真实文件验证见第 8 节。不是逐条标准符合性认证，也不估算支持率。

标准基线为 [ECMA-376](https://ecma-international.org/publications-and-standards/standards/ecma-376/)：Part 1（2016，主标记语言）、Part 3（2015，兼容与扩展）、Part 4（2016，过渡特性）；Part 2 的 2021 版本是包规范更新。具体元素用 Microsoft Open XML SDK 文档中的 ISO/IEC 29500 引文交叉核对。Microsoft 扩展独立列出，不混入 ECMA 标准缺口。

“未实现”指读取、模型或渲染路径缺失；“部分”指已有可见结果但特定语义缺失；“近似”指有明确替代算法。源码确认的缺口不等于已用真实文档证明其影响程度。

已下载官方 Part 1 PDF（5039 页），定点核对条文：§17.3.3.25 Ruby、§17.6.5 文档网格、§17.15.1.50/57 装订线位置与镜像页边距、§18.3.1.10 条件格式、§18.3.1.74 手动横向分页、§20.1.8.25 效果容器、§20.1.9.19 文字变形、§21.2.2.45 显示单位、§21.2.2.115 多级分类引用、§21.4.2.16 SmartArt 布局定义。没有逐页审查整份标准。

## 1. 跨格式共同层

| 能力 | 结论与显示影响 | 当前代码证据 |
| --- | --- | --- |
| MCE 兼容分支 | **本批补齐分支选择。** 共同 tokenizer 根据作用域内命名空间 URI 选择第一个已支持的 Choice，否则 Fallback，支持嵌套和前缀重绑定。`MustUnderstand` / `ProcessContent` / Ignorable 的完整语义仍未实现 | `core/format/docx.rs:5980`；`core/format/drawingml.rs:33`；全局未找到 `MustUnderstand` / `ProcessContent` 通用处理 |
| OOXML XML 编码 | **本批补 UTF-16 LE/BE。** 共用 XML 入口按 BOM 或首字节识别、限额转码，供正文、工作表、包内容类型与关系文件复用；UTF-32 和非法代理项仍拒绝 | `core/xml.rs:xml_utf8_input`；第 12 节 |
| SmartArt 布局定义 | **部分。** 已有生成后的 drawing part 解析和多类数据布局；没有完整执行 `layoutDef` 中的布局节点、算法、约束和规则。没有生成几何时仍按已知 layout 名称选择近似算法 | `core/format/drawingml/diagram.rs:707`；`core/format/drawingml/diagram_parse.rs:143`；`core/format/pptx.rs:9917` |
| DrawingML 效果图 | **部分。** 阴影、发光、反射、软边、3D 等已存在；但固定效果字段不能表达完整 `effectDag` 的命名节点引用、容器组合和混合图。不能把识别容器内部某个阴影算作支持整个 DAG | `core/format/drawingml/effects.rs:55` 的固定字段及 `DrawingMlPictureEffectsCapture`；`core/model.rs` 的 `Visual::AdvancedEffect` |
| WordArt 预设文字变形 | **近似，缺真实预设几何。** PPTX 保留 `prstTxWarp` 名称，但普通 warp 在 Canvas 出口按 run 相位轻微旋转，并未分别实现拱形、圆形、波浪等文字包络。单 run 相位为零，甚至可能没有可见变形。VML 路径文字有独立处理，不能据此推导所有 WordArt 已完成 | `core/format/pptx.rs:6368`；`src/render.ts:6059` |
| OMML 数学排版 | **部分。** DOCX 已保留公式结构并做近似几何排版；PPTX 的纯公式、未变形形状复用同一公式盒排版，混排或变形形状仍退回线性文本。两者均缺 OpenType MATH 度量 | `core/format/docx.rs` 的 `layout_omml_for_shape`；`core/format/pptx.rs` 的 `push_shape` |

依据：[MCE 分支选择规则](https://download.microsoft.com/download/e/1/4/e14fb96f-83b8-4a2a-84db-7fa8acbe061a/Office%20Open%20XML%20Part%203%20-%20Primer.pdf)、[SmartArt LayoutDefinition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.diagrams.layoutdefinition?view=openxml-3.0.1)、[EffectDag 子元素与结构](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.effectdag?view=openxml-3.0.1)。上述 Primer 是早期标准说明材料，MCE 版本基线仍以现行 Part 3 为准。

## 2. DOCX

| 能力 | 结论与显示影响 | 当前代码证据 |
| --- | --- | --- |
| 注音 / 拼音 / Ruby | **已补横排注音。** 保留基字和注音的字体样式，消费 hps / hpsRaise 及居中、左、右、字符/空格分布对齐，参与行宽、行高和分页。竖排 Ruby 仍未完整实现 | `core/format/docx.rs` 的 RubyCapture、EquationNode::Ruby |
| altChunk 插入内容 | **未实现。** 包内通过 `altChunk` 引用的内容不加载、不渲染。应区分安全解析本地支持子集与执行 HTML 活动内容；阻断活动内容不意味着所有静态内容只能消失 | `core/format/docx.rs:8749` |
| 镜像页边距、装订线 | **基础路径已补。** 读取 gutter、rtlGutter、gutterAtTop、mirrorMargins；装订线缩减正文空间，镜像边距按物理页奇偶交换。书籍折页/每纸两页组合、复杂浮动对象的 inside/outside 锚点仍待验证 | `PageSpec::for_page`、第 10 节真实样本 |
| 文档字符网格 | **部分。** `docGrid` 消费 `linePitch`，未消费 `charSpace`；`linesAndChars` 并未实现完整字符格约束 | `core/format/docx.rs:6385` |
| 文字框链接与溢出串流 | **部分。** 文本框已有固定边界布局，未找到 `linkedTxbx` 的跨框文字串流实现。不能把一般文本框列成完全不支持 | `core/format/docx.rs:8645`；全局 `linkedTxbx` 检索 |
| 批注与修订外观 | **部分。** 批注进入独立 story 流，没有 Word 边栏气泡版面；修订采用接受后视图，删除、move-from 隐藏。查看原稿/修订标记属于尚缺显示模式，接受/拒绝操作不在本次范围 | `core/format/docx.rs:5993`、`16164` 附近的 revision 与 story 路径 |
| 日文假名压缩 | **明确未实现。** 已保留标点压缩，但 kana 压缩另有诊断，不能把 CJK 排版整体判为完整 | `core/format/docx.rs:12331` |

[Ruby 标准结构](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.ruby?view=openxml-3.0.1)包含独立的注音与基字。完整 Word 分页、复杂浮动交互、字体塑形仍属于需真实语料量化的保真度边界，本次不据此虚构新的“完全未支持”项目。

## 3. XLSX

| 能力 | 结论与显示影响 | 当前代码证据 |
| --- | --- | --- |
| 条件格式规则 | **本批补齐常用规则匹配。** 复用现有可选计算模块处理 expression、引用型 cellIs、文字、空值、错误、重复/唯一、Top/Bottom、日期区间与基本均值；按全工作表 priority 合成属性、处理 stopIfTrue。尚不覆盖 equalAverage/stdDev、x14、全部 differential 样式，也不生成原文件不存在的空白单元格。超预算或缺计算结果时局部降级并诊断 | `core/format/xlsx.rs:3593` |
| 打印重复标题、手动分页 | **已接通重复行/列标题及手动分页。** 模型、懒加载、完整解析、协议和 Viewer 打印均消费；缩放到页数时忽略手动分页，损坏的分页标记局部忽略并诊断。部分区域列分页目前按整个打印行区间处理，复杂分区尚需原生对照 | `core/format/xlsx.rs`；`src/render.ts:sheetPrintPages`；`src/viewer.ts:#printDocument` |
| 其他打印属性 | **部分。** 未保留 pageOrder、水平/垂直居中、黑白打印等对应配置，也未发现 `legacyDrawingHF` 页眉页脚图片处理。这里应单独验收打印输出，不能只看普通 worksheet viewport | 同上 PrintSettingsParser 与 SheetPrintSettings |
| 从右至左工作表视图 | **未找到实现。** `sheetView/@rightToLeft` 无消费路径；单元格文字 RTL 和工作表列排列方向是不同能力 | `core/format/xlsx.rs`、`core/model.rs`、`src/viewer.ts` 的视图与坐标路径 |
| 单元格注音 | **未实现。** shared-string 和 inline-string 都跳过 `rPh` 内容，没有注音显示布局 | `core/format/xlsx.rs:2392`、`5153` |
| 批注 / Note 静态外观 | **未找到完整读取与显示路径。** 当前 legacyDrawing 路径抽取 shape 对应的 imagedata，不读取批注文字、锚点和 VML 文本框来生成 note 外观 | `core/format/xlsx.rs:7405`；comments 关系与内容检索 |

依据：[标准条件格式规则类型](https://learn.microsoft.com/ka-ge/dotnet/api/documentformat.openxml.spreadsheet.conditionalformatvalues?view=openxml-2.19.0)、[rowBreaks 的预览及打印语义](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.rowbreaks?view=openxml-3.0.1)、[工作表 RightToLeft](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sheetview.righttoleft?view=openxml-3.0.1)。

## 4. 图表与 PPTX

| 能力 | 结论与显示影响 | 当前代码证据 |
| --- | --- | --- |
| 传统 ChartML 多级分类轴 | **本批补齐基础层级。** 传统 multiLvlStrCache 保留叶子与稀疏父级标签，三宿主消费同一布局；支持 noMultiLvlLbl。复杂多轴布局、分组分隔线、任意旋转等保真度仍需独立验证 | `core/format/drawingml/chart_parse.rs:1526`、`2832` |
| 图表显示单位 | **本批补齐基础显示。** builtInUnit/custUnit 缩放刻度文本，保留数值几何和数据标签；显示默认或作者单位文字。General 极小非零刻度使用科学记数法。单位标签的独立手动位置/完整样式仍未覆盖 | `core/format/drawingml.rs:284`；`chart_parse.rs` 轴属性处理 |
| 完整坐标轴属性 | **部分。** 本批补数值轴 minorUnit、minorGridlines、minorTickMark，以及分类轴 orientation 对柱形、折线、面积和分组标签的消费。数值轴反向、crossAx/crossesAt 仍待补。日期轴已读取 base/major/minorTimeUnit，生成按真实日历间距排列的主次刻度；任意日期轴组合与标签样式仍待验证 | `ChartValueAxis`、`Chart::minor_axis_lines`、第 10 节 |
| PPTX 批注显示 | **未实现。** 明确诊断 comments 保留于包但不渲染；不涉及添加或回复批注 | `core/format/pptx.rs:4338` |
| PPTX 备注页与讲义版面 | **部分。** 已有按幻灯片读取备注正文及样式；没有完整 notes-page/handout 固定页面输出。仅 Inspector 中有备注文字不代表能按备注母版打印 | `core/format/pptx.rs:3800`；`core/model.rs:118`；`docs/SUPPORT.md` |
| PPTX 旧 VML drawing | **未实现一般内容。** 关系层明确诊断不渲染。DOCX VML 和 XLSX VML 图片路径的存在不代表 PPTX 有同等消费能力 | `core/format/pptx.rs:4338` |

依据：[图表显示单位](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.displayunits?view=openxml-3.0.1)、[ChartML 元素目录](https://learn.microsoft.com/it-ch/dotnet/api/documentformat.openxml.drawing.charts?view=openxml-2.9.1)。

## 5. Microsoft 扩展，单列

- XLSX x14 条件格式：当前发现非顶层 conditionalFormatting 即标记 unsupported，没有完整扩展规则处理；尤其不能用基础 dataBar 支持推导扩展负值/轴/边框等均支持。证据：`core/format/xlsx.rs:3462`。
- Slicer / Timeline 的静态显示缺失：代码明确诊断不渲染。筛选交互排除后，仍存在作者保存的外观缺口。证据：`core/format/xlsx.rs:4783`。
- 嵌入 3D model 未解码；已有 DrawingML 形状 3D 投影不是 model3d。证据：`core/format/pptx.rs:4320`、`4349`。
- ChartEx 已有瀑布、树图、旭日、直方图、箱线等实现，不能泛称现代图表未支持。其完整保真度仍需逐属性和真实文件证明。

这些属于 Office 扩展覆盖，不能计入 ECMA-376 基础元素的支持分母。Office 与标准差异的定位入口见 [Microsoft Office 标准实现说明](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/db9b9b72-b10b-4e7e-844c-09f88c972219)。

## 6. 不应继续列成“未支持”的已有能力

- DOCX：嵌套表格、不等宽分栏及平衡、页底脚注、一般文字框、部分 VML/水印、首字下沉、行号已存在实现；只能继续追具体属性和保真度。
- 图表：bar/line/area/scatter/bubble/radar/pie/doughnut/stock/surface、部分 3D、复合饼图、趋势线、误差线、数据表、标记等已有实现。`ChartKind` 与 `chart_extended_elements` 是比旧支持文档更可靠的入口。
- 特效：发光、反射、软边、阴影与 DrawingML 3D 已有模型及 Canvas 路径；缺口是完整效果图和精确光照/材料等语义，而不是这些功能全无。
- OLE：有 MSGraph、部分嵌入文本/图像/旧公式等受限静态解码路径；不能把禁止激活对象直接等同于所有 OLE 静态内容未支持。也不能用对象/文档预览图充当真实解析完成的证据。

`docs/SUPPORT.md` 中“嵌套表格不支持”“脚注不在页底”“XLSX 仅 bar/line/pie”“全部 OMML 线性化”等描述已经落后于代码。这份审计记录差异，不据源码存在就直接提升公开保真度承诺。

## 7. 建议顺序与验收边界

1. 先补可能丢内容或改变读数的路径：MCE、条件格式、Ruby、传统图表单位/层级/轴语义。
2. 再补版面契约：镜像边距/装订线、字符网格、Excel 打印标题/分页、RTL 工作表、PPTX 结构化公式。
3. 再补长尾复杂外观：SmartArt 约束布局、WordArt 预设包络、效果 DAG、批注/备注页、扩展静态对象。

每项实施前须拿真实文件证明修改前失败，修改后在对应 Office 原生参考与浏览器输出中通过，并检查既有 case。共同语义落共同层，DOCX/PPTX/XLSX 保留宿主差异回归；上表仅更新已落地的范围，未标完成的缺口仍然存在。


## 8. 首批实现与回归边界

第一批实现 MCE 分支选择、常用条件格式、图表显示单位与多级分类标签；Ruby、分页/打印的后续结果见第 9 节。未做 Word/Excel/PowerPoint 原生像素对照，浏览器可见结果与自动断言不等同于完整 Office 保真度认证。

- 真实失败：nestedAlternateContent 在共同 tokenizer 中消费了双份文字；DisplayUnits 的刻度未除以十亿；testMultilevelCategoryAxis 的父级覆盖叶子；conditional_fmt_checkpriority 的基础文字规则未进入渲染。
- 适配边界：将两个真实文件的原始 ChartML 放入现有 PPTX/XLSX 容器，验证三宿主均输出单位和父子标签。这些是派生容器，不伪称新的独立真实来源。
- 条件格式使用已有 IronCalc matcher，只新增只读匹配结果出口。计算 ABI 升为 3，Core 与 calc Wasm 必须同步发布。常量 cellIs 不触发可选计算模块加载。
- 缺计算模块、超规则单元格预算或未覆盖的规则保留其他内容，输出 UnsupportedFeature 诊断。1904 日期系统的 timePeriod 和未附公式的周区间规则保留诊断，不静默套用计算引擎的 1900/Monday 语义；不声明完整 CF 支持。
- 共同布局与匹配回归：`cargo test --lib`、`cargo test --no-default-features --features calculation-service --lib`；适配/渲染回归：`node --test tests/spreadsheet.test.mjs`；真实 Viewer：`node scripts/run-ooxml-display-regression.mjs`，可用 `BROWSER=chromium|firefox|webkit` 分别运行。

真实输入均逐字节复制自 LibreOffice core commit `db5d24b4e0d0350125babbaf377d2ed7f043e161`，来源项目许可证适用；无 Office 预览图片兜底。

| 本地文件 | 上游路径 | SHA-256 |
| --- | --- | --- |
| `ooxml-nested-alternate.docx` | `sw/qa/extras/ooxmlexport/data/nestedAlternateContent.docx` | `1b0434877bac299daf1d05202bd9e183d1aad0f594ade062db06f80eb9dc0866` |
| `ooxml-conditional-priority.xlsx` | `sc/qa/unit/data/xlsx/conditional_fmt_checkpriority.xlsx` | `21e644459edc86e2d6c260d367a2cd42fca89b2f83da7a1b3a92115ff18368d7` |
| `ooxml-display-units.docx` | `chart2/qa/extras/data/docx/DisplayUnits.docx` | `6f8557b333b62220ae9537029f746c1aadac28d6648912b4ce08a0aef4dc004c` |
| `ooxml-multilevel-axis.docx` | `chart2/qa/extras/data/docx/testMultilevelCategoryAxis.docx` | `ba37d11dcde01e023c0df21f246047eb399fe9c308e3221a067e779e8e0dc62c` |

验证结果（2026-09-23）：Core 799 通过 / 5 忽略，计算模块 77 通过，JS 工作表与场景/体积回归 112 通过。Chrome、WebKit 各 4 个真实 Viewer case 通过，控制台无错误；截图在 `output/ooxml-display/`。Firefox 启动后持续占用 CPU、未进入断言，已停止本次测试进程，不计为通过。

已逐字段复核快照差异：修正多级标签空间影响旭日图的回归后，sunburst.xlsx 恢复原始 hash；background-worksheet.xlsx 仅新增可选计算模块诊断，剔除该诊断后，初始、逐页和完整场景与提交基线完全相等。只更新后一项的诊断快照。

当前优化产物 Core 为 3,588,854 字节，calc 为 2,028,244 字节；本批体积预算分别调整为 3,600,000 / 2,040,000 字节，计算模块仍按需加载。未新增依赖或随包字体，未提交或发布。


## 9. 第二批：Ruby 与工作表打印（2026-09-23）

- Ruby 复用 DOCX 现有行内组合对象、Row 几何、字体度量及各正文/表格/页眉布局入口；不增加渲染协议对象。真实文档修改前缺失 `きもん`，修改后基字与注音可见，回归检查相对位置、字号及居中。ODF 现有 ruby 是上标降级，未保留 DOCX 的基线抬升等输入语义，不能宣称严格等价；本批不改 ODF 适配行为。
- XLSX 的打印属性进入共同 SheetPrintSettings，协议升级至 74，旧协议仍可读。标题行/列按源坐标裁切，和正文拼成实际打印画布；分页预览保留正文坐标。支持整数页边界、重复标题与手动行/列分页组合；跳过缓存的自动分页。无新依赖、字体或格式专用 Canvas 出口。
- 两个原始工作簿的 metadata 回归修改前失败，修改后通过；另有完整解析回归。原始手动分页样本的 Print_Area 恰好止于分页处，因此分页几何测试显式扩展打印区，以验证边界被消费；标题样本第 5 行无文字，浏览器测试在原结构上添加可见标记及分页，验证每页 Canvas 真正绘制标题。原始 fixtures 保持逐字节不变。
- 标准边界：[RubyProperties 的半磅字号/基线抬升](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.rubyproperties?view=openxml-3.0.1)、[RowBreaks 示例中的页边界](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.rowbreaks?view=openxml-3.0.1)、[Fit To 忽略手动分页](https://support.microsoft.com/en-us/excel/insert-move-or-delete-page-breaks-in-a-worksheet)。
- 未声明完整 Office 排版一致性：竖排 Ruby、复杂脚本的分布塑形、相互交错的局部分区分页、复杂 Fit To 页数目标、跨重复标题区域的浮动图形仍需要原生样本对照；pageOrder、居中、黑白、页眉图片仍在原审计待办中。

| 本地文件 | 来源 | SHA-256 |
| --- | --- | --- |
| `ooxml-ruby.docx` | LibreOffice 上述固定 commit，`sw/qa/extras/ooxmlexport/data/tdf49073.docx` | `6c5940209c085fd384a555aeb80e16c4a3a495415839d7366a7c76ddcda4501f` |
| `ooxml-print-titles.xlsx` | LibreOffice 同 commit，`sc/qa/unit/data/xlsx/tdf91332.xlsx` | `5d65bb68c1c82e70013245d1d133d9eb4034953f1a5a5bba99ea7b4cb1e89800` |
| `ooxml-manual-breaks.xlsx` | Microsoft Open XML SDK v3.5.1，`test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/missingcalcchainpart.xlsx` | `8157632b6baee46941f615f98507380b3081269e2b2f8f0fecfb509ad8b04067` |

验证：Core 801 通过 / 5 忽略，工作表、协议、场景与体积 JS 回归 134 通过，原有场景快照无新增变化。Chrome / WebKit 各完成 5 个 Viewer 文件及实际两页打印回归，控制台无错误。Firefox 再次在启动阶段 90 秒超时，未进入断言，不计通过。API 声明快照随新增打印属性更新。最终 Core 为 3,600,850 字节，calc 为 2,028,244 字节；Core 预算为 3,615,000 字节（本批新增约 12 KB），calc 预算维持不变。未提交、未发布。

## 10. 第三批：次刻度、反向分类与装订版面（2026-09-23）

- ChartML 数值轴的次刻度/次网格线走共同解析和几何，三个宿主仅消费结果；支持次单位、内/外/交叉刻度及网格线颜色/宽度。原始 minorTickMark.xlsx 在旧 Core 中为 0 条次刻度，新实现至少 12 条。将同一原始 ChartML 放入已有 DOCX/PPTX 容器，验证适配消费；另在该 ChartML 上添加 0.5 次单位及特定颜色/宽度，三个宿主均实际绘制 7 条次网格线，旧 Core 均为 0。派生测试不伪称独立原始文件。
- 分类轴反向同时作用于数据位置、标签、面积几何与柱形系列连接线。原始 tdf132174.docx 的首项柱形和标签从底部改为顶部，数值与标签保持配对；共同层另用该图表派生面积/多级标签边界回归。
- DOCX 装订线进入页面可用空间：左侧、顶部、RTL 右侧分别验证；镜像边距依据物理页交换。真实镜像样本正文 x 从奇数页 120 px 变为偶数页 280 px，浏览器翻页可见。异常过大的装订线夹取并保留诊断，避免整篇内容消失。
- 本批不覆盖完整日期轴、数值轴反向或交叉位置；字符网格、PPTX 结构化公式仍未补。书籍折页/每纸两页与顶部装订线的组合不在已验证范围，不能声明符合该组合的完整标准语义。标准边界参见 [GutterAtTop](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.gutterattop?view=openxml-3.0.1) 与 [MirrorMargins](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.mirrormargins?view=openxml-3.0.1)。

下列原始输入来自第 8 节同一 LibreOffice 固定 commit，逐字节复制：

| 本地文件 | 上游路径 | SHA-256 |
| --- | --- | --- |
| `ooxml-minor-ticks.xlsx` | `chart2/qa/extras/data/xlsx/minorTickMark.xlsx` | `acf1d82ce0e62de44c57d9a8ff34396ca8572f2fdeb9ac1ea54a22ae70f1fd3f` |
| `ooxml-reversed-categories.docx` | `chart2/qa/extras/data/docx/tdf132174.docx` | `3374fb34920f604b55e91b8664807d4d4506b306b48472c99dc8a4f10f3dacbe` |
| `ooxml-gutter-left.docx` | `sw/qa/extras/ooxmlexport/data/gutter-left.docx` | `a058ce6c448e0bd3d3fd03e09e45e7dfda6e37e07ffb769f2a6cf3ced4cb23b0` |
| `ooxml-gutter-top.docx` | `sw/qa/extras/ooxmlexport/data/gutter-top.docx` | `be80449fb8c183222b81c20afdb962c653cc89a388fcfadf70d010e9c8d49b32` |
| `ooxml-gutter-right.docx` | `sw/qa/extras/ooxmlexport/data/rtl-gutter.docx` | `32f641b3cc63f0cadef35c3cd69026262413a495506882af12a3f70b27c6e7a5` |
| `ooxml-mirror-margins.docx` | `sw/qa/extras/ooxmlexport/data/fdo73541.docx` | `cf845c867b8f64336f010c4514cbe588042ba087fde7448c52058b7edbf83bcb` |

验证：Core 805 通过 / 5 忽略；工作表、协议、文档保真度与体积 JS 回归 148 通过，既有场景快照无新增变化。Chrome / WebKit 各通过 11 个真实 Viewer 文件及两页打印，包含 Canvas 次刻度路径、反向标签顺序和镜像边距翻页断言，控制台无错误。截图位于 `output/ooxml-display/`。未做 Office 原生像素对照；反向分类样本的长标签仍有既有换行拥挤问题，不宣称整体图表保真度已完成。

最终 Core 3,607,964 字节，沿用 3,615,000 字节预算；calc 仍为 2,028,244 字节。未新增依赖、字体或协议字段，未提交或发布。

Firefox 本批再次在浏览器启动阶段超过 90 秒，未进入 Viewer 断言；进程已退出，不计通过。

## 11. 第四批：日期轴日历刻度（2026-09-23）

- 真实 XLSX 中 `dateAx` 明确规定每月主刻度与每 7 天次刻度。修改前只在 10 天一条的数据点上检查固定数字间隔，因此缺少 2 月 1 日等月初标签，也没有次网格线。现在共享轴几何从日期序列号生成真实月初/年初或定日间隔，位置按实际经过天数投影；类别数据点几何仍用原始日期，不挪动曲线。三宿主复用已有轴线、额外标签绘制入口。
- 原始文件中的次网格线颜色为 `DDDDDD`，线宽 `w="0"`；基础解析过去把这条零宽线直接禁用。本批只把**次网格线**的该写法按默认细线绘制，保留作者颜色，不改变既有主网格线的零宽处理。
- 固定语义参考：[DateAxis 子元素](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.dateaxis?view=openxml-3.0.1) 与 [MajorTimeUnit](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.majortimeunit?view=openxml-3.0.1)。任意小数时间步长、复杂跨年自动对齐、日期轴的全部交叉语义仍未声明完成。

| 本地文件 | 上游路径 | SHA-256 |
| --- | --- | --- |
| `ooxml-calendar-axis.xlsx` | `chart2/qa/extras/data/xlsx/tdf137917.xlsx` | `3ff0f05831efa664e1100c77d9569a66d4303101ee36a7dbf91ab542121f03e7` |

验证：原始工作簿及 DOCX/PPTX 派生容器的日期轴回归在旧 Core 失败、当前 Core 通过。Core 原生回归 806 通过 / 5 忽略；JS 文档保真度、工作表、协议与体积 151 通过。Chrome / WebKit 各 12 个原始 Viewer 样本、2 个 DOCX/PPTX 日期 ChartML 派生容器及两页打印通过，日期轴截图能看到不等距的月初标签和七天浅灰网格线，控制台无错误。Firefox 本轮未重试，沿用第 10 节失败边界；没有 Office 原生像素对照。最终优化 Core 3,611,914 字节，小于 3,615,000 字节预算，calc 仍为 2,028,244 字节。未提交或发布。

## 12. 第五批：UTF-16 XML 部件（2026-09-23）

- W3C XML 1.0 §4.3.3 要求 XML 处理器读取 UTF-8 与 UTF-16。过去共用 tokenizer 在 BOM 检查时直接拒绝 UTF-16；本批在同一入口有界转为 UTF-8，再沿用原有 XML 结构、实体和 MCE 处理。DOCX/XLSX/PPTX、ODF 及 OPC 的内容类型、关系文件共用这一实现；UTF-32 仍不支持。依据：[W3C XML 1.0 字符编码](https://www.w3.org/TR/xml/#charencoding)。
- 原始回归部件为 LibreOffice 官方语料 `sc/qa/unit/data/xlsx/tdf167689_x15_namespace.xlsx` 内的 `customXml/item18.xml`（SHA-256 `81538453afbd8a62103199c406ff4b9d7c6777e78020e1112f3f147bca93f3b9`）。修复前返回 `XML_ENCODING_UNSUPPORTED`；修复后原始 LE 与仅调换字节序的 BE 版本均得到 CDATA `False`，非法代理项仍被拒绝。
- 为验证可见显示，将现有真实 `ooxml-minor-ticks.xlsx` 的包内容类型、关系、工作簿与首张工作表仅重新编码为 UTF-16。旧 Core 在 `_rels/.rels` 拒绝打开；新 Core 的对象文本和边界与原件逐项相同。DOCX `ooxml-nested-alternate.docx` 与 PPTX `corpus-chart-wall.pptx` 的正文/幻灯片 XML 作同样重编码后也保持原有对象几何和文字。Chromium / WebKit 实际 Viewer Canvas 分别可见工作表 `col1` / `col2`、DOCX `ABC`、PPTX `Anne`，控制台无错误。原件与派生件的差异仅是 XML 编码，不能称派生件为上游原始样本。

验证：Core 807 通过 / 5 忽略，工作表、文档保真度、协议及体积 JS 154 通过，`npm run check` 与格式检查通过；Chromium / WebKit 各 12 个原始 Viewer 样本、5 个派生显示载体及两页打印通过。Firefox 本轮仍在浏览器启动阶段 90 秒超时，未进入 Viewer 断言，不计通过。现有优化 Core 3,623,011 字节，因新增共用转码调整预算至 3,630,000 字节。未新增依赖、字体或协议字段；未提交或发布。

## 13. 第六批：PPTX OMML 结构化公式（2026-09-23）

- DrawingML 的数学内容在 `a14:m` 的 MCE `Choice` 中可携带真正的 OMML，`Fallback` 常是静态图像；最终显示应选择公式结构。依据：[Microsoft DrawingML Math](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/853b19c7-68a9-4f9a-a2ae-5e6cb0d02e62)。之前 PPTX 将结构化公式一律线性化为 `TextRun`；现在纯公式、无形状或文字变换的形状复用 DOCX 的公式盒布局，生成分子、分数线、分母等独立可见对象。形状太小、对象预算不足或布局失败时保留原有线性文本与诊断。混排及变形公式仍走线性路径，DOCX/PPTX 均尚未使用 OpenType MATH 度量。
- 原始容器是 LibreOffice 官方语料 `sd/qa/unit/data/pptx/tdf129372.pptx`（本地 `tests/fixtures/ooxml-omml-slide.pptx`，SHA-256 `b9bc871d807225cb92cbc2dfbead959e86920d3e0b2dc53ad6d59fb0110e45b7`），其真实 `a14:m` 公式为 `𝜕`，原样显示保持通过。分式 XML 原样取自同一语料的 `sw/qa/extras/ooxmlexport/data/math-vertical_stacks.docx` 的 `m:f`，仅替换原 PPTX 公式子树以生成派生显示载体；分子/分母及分数线语义依据：[Microsoft OMML Fraction](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.fraction?view=openxml-3.0.1)。该派生件不是上游原始样本，也没有 Office 原生像素对照。
- 新回归在修改前失败（缺少独立的上下分子、分母对象），修改后通过。真实 Viewer 的 Chromium / WebKit Canvas 均画出上下两层文字与横线，控制台无错误；源形状本身约 39×39 CSS 像素，所以分式在全页截图中很小。Firefox 再次在浏览器启动阶段 90 秒超时，未进入 Viewer 断言，不计通过。Core 807 通过 / 5 忽略，演示/格式 JS 103 通过，体积及跨格式样本 JS 35 通过，`npm run check`、`cargo fmt --all --check` 和 `git diff --check` 通过。全格式包构建成功；最终优化 Core 3,628,132 字节，低于 3,630,000 字节预算。未提交或发布。

## 14. v0.2.74 发布边界（2026-09-25）

本批还包含共同 ChartML 的 3D 几何、SmartArt 组织结构与样式、三个宿主的图片变换、VML 笔迹、DOCX 艺术边框及表格样式、GDI 字体度量，以及 PPTX 文字效果与挤出性能修复。真实样本和对应脚本一并保留；既有章节中列出的未实现范围不会因发布而自动变为完整支持。

- 快照协议为 75，计算 ABI 为 3；Core、计算模块与 JavaScript 必须作为同一版本发布。公开声明新增打印标题/分页/分片和文字描边、发光属性，已对照 v0.2.71 包复核后更新 API 摘要。
- 最终 macOS 优化 Core 为 3,771,967 字节，calc 为 2,030,982 字节。Core 预算由早期批次的 3,630,000 调整为 3,800,000 字节，包含后续上述几何与样式实现，并留不到 1% 的 Linux 代码生成余量；calc 仍限 2,040,000。完整 SDK 压缩 8,006,869 字节、解包 21,686,842 字节，保持原发布包预算，未新增依赖或随包字体。
- `Text_withExtrusion_200chars.pptx` 原样来自 Microsoft Open XML SDK v3.5.1 官方语料。逐层整页分配预算在优化前失败，优化后通过；Chromium 的 960×720 绘制三次中位数约 537→303 ms、2× 约 2145→1173 ms，累计画布分配减少约 89%。原文件在 Chromium/WebKit 的 1×、2× 和缩放局部视口均与优化前像素一致。无 Canvas 滤镜的实现保持原路径，不宣称获得相同提速。
- 本机 Chromium/WebKit 完成真实 Viewer、两页打印、图片变换、笔迹、文字效果、SmartArt 和 20 页图表回归；本机 Firefox 启动受阻，其最终结果以 Linux 三浏览器 CI 为准。缺少已审核的 Office/iWork/WPS 原生验收清单，发布仍标记 `not-certified`，不把浏览器回归当作 Office 原生保真度认证。

发布前将 8 个变化摘要与 v0.2.71 正式 npm 包逐对象核对：`chart-original.docx` 新增 14 条表格边框并扩大线图图例键；`complex2005_12rtm.docx` 新增 126 条边框、29 处作者 caps 标题和两张图片的裁剪/3D 效果，图片字节不变；`corpus-stacked-mix.pptx` 图例采用作者 24px 字号；三个 XLSX 图表样本恢复保存的 14.25/15pt 默认行高，连带锚定图表几何，空标题样本图例采用黑色；`background-worksheet.xlsx` 恢复 13.5pt 行高及 SmartArt 阴影，保留计算模块需求诊断；`piechart_outside.xls` 调整外侧标签度量与短引导线，全部标签文字保持不变。只有这些已解释的摘要更新，其他原有摘要继续严格逐字节检查。打印与辉光的旧测试分别改为断言实际片段视口和 Canvas 阴影，真实浏览器回归同时保留。

Linux 发布门禁还覆盖了两个真实 DOC 选择问题：压缩开标点的字形左移不应移动逻辑选择起点（原间隙 4.86px），Firefox 的元素边界 Range 不应把行容器计入文字高亮（原高度约为字形 2.28 倍）。分别在共享文字渲染和 Viewer 范围端点归一化处修复，原样 `word-table-auto-height.doc` 的现有矩阵在修复前失败、修复后通过。测试同时修正平台前提：复制使用 ControlOrMeta、读取 ClipboardEvent 自身数据、独立拖选清除前次范围、异步单元格点击等待地址状态；宽度对照实际渲染片段，标题像素探针不再扫描相邻行，并以字形覆盖和至多 1.5 倍高度限制高亮。Linux WebKit 辉光边缘强度为 19，旧缺陷不足 10，阈值据此保留明确区分；不要求不同字体的上下留白对称。固定 Rust 1.96 全格式原生测试 2,495 项及 JS 967 项通过，保留既有 11 项 ignored 与 3 项 skipped。

原计划标签 v0.2.72 在公开发布前取消，标签保留不移动。x86 Linux Firefox 的替代字体把日期段落排成两行，而已有测试写死为一行；选择层实际使用渲染片段，文本完整。回归改为要求完整原文、渲染布局来源，以及每个渲染片段的 DOM Range 不再额外换行，避免把宿主字体排版差异误判为选择层错误；以 v0.2.73 重新完成发布门禁。

后续 v0.2.73 也在公开发布前被门禁阻止，标签同样保留。CI Firefox 的混合替代字体产生 1.5 px 的 Range 顶边差异，选区水平连续且全部覆盖字形；同行垂直定位改为检查渲染片段盒，保留 Range 水平连续性和像素覆盖断言。Worker 字体响应测试将 100 次事件循环空转改为 5 秒实际截止时间，并用 50 ms 异步字体提供器复现修改前失败、修改后通过。以 v0.2.74 先完成 CI，再创建发布标签。

## 15. v0.2.75 发布前场景差异复核（2026-10-06）

使用已公开的 `docviewkit-viewer-0.2.74.tgz` 中的 Core Wasm，以及对应源码的协议解码器，重放原始文档的初始、逐单元和全量场景；八个旧摘要均与现有基线一致。再与候选版本逐字段比较，确认以下差异后更新这八个摘要，其余场景仍严格检查原值。所有样本的文档级信息和诊断保持一致。

| 原始文件 | 已核对的差异与回归 |
| --- | --- |
| `chart-original.docx` | 共同层折线几何分离曲线与标记，图表标题、坐标轴和图例改用共享文字布局；正文绕排把首行缩进等价转换为片段位置。保留折线及原始标题回归。 |
| `word-data-label-borders.docx` | 30 个图表文字对象由 rich-text 改为 text-layout；原始文字、源定位和边框语义保留。 |
| `corpus-stacked-mix.pptx` | 共享 plot/轴/图例约束调整图表对象的位置及刻度，变化仅在三个 ChartML 部件内；共同层和三个宿主的图表回归继续生效。 |
| `funnel-pp1.pptx`、`sunburst.xlsx`、`color_funnel.xlsx` | 每个文件仅一个自动标题的八个字段变化：标题高改为 28 px，禁用缩字和内边距，使用固定行高及垂直居中；数据对象、标签和源定位不变。 |
| `background-linear.docx` | 仅交换两个颜色 stop，修复 VML 负角度的色带方向；由原始 aqua 背景及角度/focus 边界回归覆盖。DrawingML 的 stop 顺序不变。 |
| `complex2005_12rtm.docx` | DOCX 子对象按父级浮动层级绘制，长段落保留继续换行标记；折叠嵌套表格末尾的一段空 cell-end marker，后续表格随之重新流排。仅移除该空段落，没有丢弃正文；真实嵌套表格、单元格边界及浮动图表层级回归覆盖这些路径。 |

八个场景在未更新基线时失败，更新后通过。此复核证明上述摘要变化的来源，不代表完整 Office 原生像素认证；缺少审核后的 native release manifests 时仍标记为 `not-certified`。

共享字体名称协议元数据及格式化后的源码位置使 PDF pack 从公开 v0.2.74 的 3,139,335 字节变为 3,140,235 字节，增加 900 字节，未新增或修改字体二进制。原 3,140,000 字节门禁超出 235 字节；预算仅增加 1,000 字节到 3,141,000，其余字体、Core 与 calc 预算维持原约束。
