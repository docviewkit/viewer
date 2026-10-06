# DOCX 与 Microsoft Open XML 渲染语义差距审计

调研日期：2026-07-17

## 结论

OfficeViewer 已覆盖 DOCX 的正文、常见段落与 run 样式、分页、页眉页脚、列表、表格、图片、批注/注释内容和部分东亚排版，但仍不是完整的 WordprocessingML 布局实现。当前最影响“与 Word 看起来一致”的缺口集中在：**段落断行与分页、表格布局、分节/文档网格、浮动对象、完整样式级联**。

以下结论仅针对用户可见的渲染和布局，不枚举宏、邮件合并、权限、元数据等非渲染功能。规范依据使用 Microsoft Learn 中的 Open XML SDK / ISO/IEC 29500 元素说明；实现依据为当前仓库源码。

## `test.docx` 直接命中的规范缺口

对 `/Users/wjfree/Downloads/test.docx` 解包检查后，已确认它不是只碰到通用的字体度量误差，而是实际使用了当前实现尚未完整建模的 WordprocessingML 语义：

- `settings.xml` 指定 `w:characterSpacingControl="compressPunctuation"` 和 `w:useFELayout`；当前设置解析器只读取奇偶页页眉和默认制表位，所以 Word 的东亚标点压缩/兼容布局没有生效。当前新增的 `autoSpaceDE/autoSpaceDN` 只覆盖中西文、中数字间距，不等价于这两项规则。
- `styles.xml` 中存在 `w:jc w:val="both"`；当前解析器把中、右以外的值统一降为 `Start`，因此 Word 的两端对齐没有实现。
- `numbering.xml` 使用 `w:numFmt w:val="japaneseCounting"`；当前编号格式化只专门实现十进制回退、英文字母、罗马数字和项目符号，中文/日文计数会错误退化成阿拉伯数字。
- 编号定义含 `w:tab w:val="num"`，样式也包含中心/右对齐制表位；当前编号模型不保存编号制表位，渲染协议又不保留 leader/小数点等制表属性，容易造成序号和正文起点不一致。
- 文档中的 `w:snapToGrid`、`w:adjustRightInd` 均为关闭值 `0`；它们本身虽未建模，但不是这个样例当前差异的主要来源。

## 已确认的主要缺口

| 优先级 | 能力 | Open XML / Word 要求 | 当前实现证据 | 可见影响 | 置信度 |
| --- | --- | --- | --- | --- | --- |
| P0 | 段落断行、对齐与分页 | [`w:pPr`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.paragraphpropertiesextended?view=openxml-3.0.1)包含 `bidi`、`contextualSpacing`、`kinsoku`、`overflowPunct`、`snapToGrid`、`textDirection`、`wordWrap` 等规则；对齐不只有左/中/右 | 段落模型只保存少量属性；`w:jc` 除中/右外一律落为 Start；预分页按字符宽度逐字断行，并明确声明是近似模型：[docx.rs:L671-L704](../../core/format/docx.rs#L671-L704)、[docx.rs:L2493-L2502](../../core/format/docx.rs#L2493-L2502)、[docx.rs:L3317-L3325](../../core/format/docx.rs#L3317-L3325)、[docx.rs:L5953-L6125](../../core/format/docx.rs#L5953-L6125) | 标点、英文单词、两端对齐、RTL/复杂文字、孤行控制和跨页位置会与 Word 不同；一个断行差异会级联改变后续所有页 | 高 |
| P0 | 表格宽度、AutoFit、边框和跨页 | [`w:tcW`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.tablecellwidth?view=openxml-3.0.1)是参与表格布局算法的“首选宽度”；[`w:tblLayout`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.tablelayout.type?view=openxml-3.0.1)区分 AutoFit/固定布局；[`w:tblHeader`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.tableheader?view=openxml-3.0.1)要求跨页重复表头；单元格还支持[`垂直对齐`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.tablecellproperties.tablecellverticalalignment?view=openxml-3.0.1) | 内部表格模型只有网格宽度、最小行高、合并、底色和段落；整张表强制铺满栏宽，使用固定 4px 内边距和固定灰色边框，超高行直接裁剪：[docx.rs:L1383-L1431](../../core/format/docx.rs#L1383-L1431)、[docx.rs:L6849-L6918](../../core/format/docx.rs#L6849-L6918)、[docx.rs:L6946-L6971](../../core/format/docx.rs#L6946-L6971)、[docx.rs:L7142-L7171](../../core/format/docx.rs#L7142-L7171)、[docx.rs:L7424-L7449](../../core/format/docx.rs#L7424-L7449) | 合同、报价单、报告中的列宽、行高、边框、表头重复、单元格垂直位置和分页容易明显错位 | 高 |
| P0 | 分节符、非等宽分栏与文档网格 | [`w:type`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.sectiontype?view=openxml-3.0.1)定义 next/odd/even/continuous/nextColumn 分节；[`w:cols`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.columns?view=openxml-3.0.1)支持显式非等宽 `w:col`；[`w:docGrid`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.docgrid?view=openxml-3.0.1)控制东亚文字每行字符数和每页行数 | `PageSpec`只保存等宽栏数与统一间距；解析器只读 `pgSz`、`pgMar`、`cols/@num`、`cols/@space`；每次切节都无条件新建页面：[docx.rs:L34-L80](../../core/format/docx.rs#L34-L80)、[docx.rs:L1998-L2073](../../core/format/docx.rs#L1998-L2073)、[docx.rs:L5497-L5510](../../core/format/docx.rs#L5497-L5510) | 连续分节会错误分页；奇偶页分节、下一栏分节、非等宽栏、中文稿纸/网格文档的页面结构不一致 | 高 |
| P0 | 浮动图片/文本框的定位与环绕 | DrawingML 的[`relativeFrom`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.wordprocessing.horizontalposition.relativefrom?view=openxml-3.0.1)决定坐标相对页、栏、边距或字符；[`wrapTight`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.wordprocessing.wraptight?view=openxml-3.0.1)依赖 `wrapPolygon`；[`wrapNone`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.wordprocessing.wrapnone?view=openxml-3.0.1)还受 `behindDoc` 和层级影响 | 模型未保存 `relativeFrom`、对齐、`behindDoc`、`relativeHeight` 或环绕多边形；所有 Tight/Through/Square 都退化成左右矩形排除区，坐标按栏左/页上边距解释：[docx.rs:L1110-L1135](../../core/format/docx.rs#L1110-L1135)、[docx.rs:L1284-L1350](../../core/format/docx.rs#L1284-L1350)、[docx.rs:L2724-L2772](../../core/format/docx.rs#L2724-L2772)、[docx.rs:L2918-L2953](../../core/format/docx.rs#L2918-L2953)、[docx.rs:L6228-L6251](../../core/format/docx.rs#L6228-L6251) | 带公章、签名、侧栏、浮动图片和文本框的文档会出现位置、环绕形状和前后层级错误 | 高 |
| P0 | 完整样式级联 | WordprocessingML 支持段落、字符、表格、编号、链接样式和默认属性；字符样式通过 `w:rStyle` 应用，见 Microsoft 的[样式类型说明](https://learn.microsoft.com/en-us/office/open-xml/word/how-to-create-and-add-a-character-style-to-a-word-processing-document) | 样式解析器只接收 `w:type="paragraph"`，仅保留有限的段落/run 属性；字符、表格、编号样式及条件表格样式未进入模型：[docx.rs:L3580-L3624](../../core/format/docx.rs#L3580-L3624)、[docx.rs:L3668-L3855](../../core/format/docx.rs#L3668-L3855) | 同一文档中大量“没有直接写在 run 上”的字体、字号、颜色、边框和表格样式会静默退回默认值；影响面广于单个元素 | 高 |
| P1 | 高级编号语义 | [`w:num`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.numberinginstance?view=openxml-3.0.1)允许 `lvlOverride/startOverride`；[`lvlRestart`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.levelrestart?view=openxml-3.0.1)规定多级列表何时重启；编号格式远多于十进制/字母/罗马数字 | 当前只解析抽象级别的 start/numFmt/lvlText/lvlJc/缩进/字体和 num→abstractNum 映射；未知格式退回十进制，嵌套级别统一在上级变化时重置：[docx.rs:L4122-L4259](../../core/format/docx.rs#L4122-L4259)、[docx.rs:L5322-L5395](../../core/format/docx.rs#L5322-L5395)、[docx.rs:L7551-L7559](../../core/format/docx.rs#L7551-L7559) | 复杂法律条款、中文编号、跨节列表、手动重启和实例级覆盖会显示错误号码 | 高 |
| P1 | 脚注/尾注、批注与域的页面语义 | [`footnotePr`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.footnoteproperties?view=openxml-3.0.1)定义位置、格式、起始值、重启策略；[`commentRangeStart`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.commentrangestart?view=openxml-3.0.1)锚定完整文本范围；[`fldChar`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.fieldchar?view=openxml-3.0.1)区分域指令和当前结果 | 脚注、尾注、批注统一追加到文档末尾的独立 flow；引用横坐标按字符数近似；域主要保留缓存结果，仅 PAGE 做动态替换；修订固定采用“接受全部”的可见视图：[docx.rs:L2286-L2305](../../core/format/docx.rs#L2286-L2305)、[docx.rs:L5768-L5785](../../core/format/docx.rs#L5768-L5785)、[docx.rs:L6154-L6217](../../core/format/docx.rs#L6154-L6217) | 页脚脚注不在引用页底部，批注不能准确标出范围，页码之外的 TOC/REF/日期等域可能陈旧，修订视图与 Word 当前显示状态不同 | 高 |
| P1 | 高级字符排版与制表位 | Microsoft 的 [run 属性说明](https://learn.microsoft.com/en-us/office/open-xml/word/working-with-runs)包括字符样式、边框、阴影、小型大写、字距/字偶距、文字方向等；[`ruby`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.ruby?view=openxml-3.0.1)要求拼音/注音显示在基文字上方；[`tabs`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.tabs?view=openxml-3.0.1)按样式层级叠加并支持更丰富的对齐/leader | `TextStyle`只包含常用字体、字号、颜色、粗斜体、下划线、删除线、高亮、基线和字距；run 解析范围与之相同。制表位只保留位置和左/中/右三种对齐，进入渲染协议时又只传位置：[docx.rs:L213-L244](../../core/format/docx.rs#L213-L244)、[docx.rs:L2307-L2334](../../core/format/docx.rs#L2307-L2334)、[docx.rs:L2519-L2675](../../core/format/docx.rs#L2519-L2675)、[docx.rs:L7300-L7320](../../core/format/docx.rs#L7300-L7320) | 拼音/注音、小型大写、字符缩放、复杂下划线、制表符前导点和小数点对齐等排版会丢失或退化 | 高 |

## 还存在但优先级较低的缺口

- `settings.xml` 目前只读取 `evenAndOddHeaders` 与 `defaultTabStop`；Word 的大量兼容性开关尚未建模：[docx.rs:L4050-L4114](../../core/format/docx.rs#L4050-L4114)。这类开关通常只在旧文档或特定生产器中触发，但一旦触发会改变断行、制表位和间距。
- [`w:tblpPr`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.tablepositionproperties?view=openxml-3.0.1)定义浮动表格相对页面、边距或段落的位置与绕排距离；当前表格始终进入正文流并占满栏宽。
- 嵌套表格被扁平化，VML 图形使用占位符，`altChunk` 不加载；源码已有明确诊断：[docx.rs:L2075-L2090](../../core/format/docx.rs#L2075-L2090)、[docx.rs:L2955-L2973](../../core/format/docx.rs#L2955-L2973)。
- 图片当前以默认裁剪渲染，未见 DOCX 图片裁剪、旋转、效果和形状遮罩进入 `Drawing` 模型：[docx.rs:L1284-L1305](../../core/format/docx.rs#L1284-L1305)、[docx.rs:L6699-L6723](../../core/format/docx.rs#L6699-L6723)。
- Word 2010+ 的 [`w14:rPr` 扩展](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/cdfcbd94-746b-42c4-8a0a-85efd7d24985)还定义文本轮廓、填充、阴影、发光、反射、连字、数字形式和样式集；当前字符模型没有这些视觉属性。它们影响局部像素，优先级低于会导致整页错位的布局规则。

## 建议实施顺序

1. 先统一“预分页”和 Canvas 文本布局的断行算法，并补 `kinsoku`、单词边界、两端对齐、RTL/复杂文字；这是所有后续页坐标的基础。
2. 实现 `tblW/tcW/tblLayout`、单元格边距/边框/垂直对齐、`tblHeader/cantSplit`，用 Word 基准图覆盖跨页表格。
3. 扩展 Section 模型，支持分节类型、非等宽栏、`docGrid/snapToGrid`、奇偶页起始。
4. 扩展浮动对象锚点、相对坐标、层级与 wrap polygon；随后补文本框内部布局。
5. 把样式系统升级为完整的 paragraph/character/table/numbering 级联，再补高级编号、注释和字符排版。

每一项都应使用最小 DOCX + Word 固定版本 PNG/页面几何作为独立 oracle；“文件能打开”不能证明这些语义已正确实现。
