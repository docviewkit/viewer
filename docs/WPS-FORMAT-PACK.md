# WPS FormatPack 设计与验收

## 目标与边界

`@docviewkit/sdk/wps-formats` 是 `.wps`、`.et`、`.dps` 的独立可选包。它必须：

- 在浏览器本地解析真实文档结构，复用 OfficeViewer 的对象模型、源映射、诊断、安全预算和 Canvas 渲染器；
- 不启动 WPS Office，不调用转换服务，也不把内嵌缩略图、预览图或封面作为最终结果；
- 遇到单个未知记录、资源或对象时隔离失败并继续输出其余可用内容，只有结构、安全或资源风险不可隔离时才整体失败；
- 只有通过 WPS Office 原生金标的逐页、逐 slide、逐 sheet 门禁，才允许对审核语料声明“视觉差异 2% 以内”。

这不是开放式插件注册表。`wps` 是现有闭集 `FormatPackCandidate` 中的一个正式成员，与 ODF、iWork、旧版 Office、PDF 和 XPS 共享同一受限运行时合同。

## 文件识别

2026-07-30 对 macOS WPS Office 12.1.26046 自带空白模板和复杂模板的结构检查表明，当前目标文件都是 OLE CFB 容器，并复用已有二进制 Office 记录族：

| 扩展名 | 必须可达的主流 | 解析族 | 输出种类 |
| --- | --- | --- | --- |
| `.wps` | `WordDocument`，以及适用的 `0Table`/`1Table` | Word FIB/CLX/FKP | text document |
| `.et` | `Workbook` 或 `Book` | BIFF workbook/worksheet | spreadsheet |
| `.dps` | `PowerPoint Document` 和 `Current User` | PowerPoint edit/persist/slide records | presentation |

调度分两层：

1. 输入必须先命中 CFB 魔数；单凭文件名不会加载 WPS 包。
2. 原始 `fileName` 的 `.wps`、`.et` 或 `.dps` 只负责在 WPS 与旧版 Office 两个同类 CFB 包之间选择。Wasm 仍须验证 CFB 图、可达主流和记录签名。

内部 `DocumentInfo.format` 保留底层记录族的 `doc`、`xls` 或 `ppt`。扩展名属于包选择信息，不伪造一种新的文档语义或协议格式。

## 运行时结构

```text
host bytes + original fileName
  → CFB content probe
  → candidate "wps"
  → lazy @docviewkit/sdk/wps-formats
  → shared office-viewer-legacy-office.wasm in a per-document Worker
  → validated CFB sector graph
  → Word / BIFF / PowerPoint record parser
  → source-mapped render objects + diagnostics
  → shared layout compiler and Canvas2D replay
```

WPS 是独立 npm 子路径和运行时候选，但复用旧版 Office 的 CFB Wasm 产物。两者共享同一套受审计 Rust 实现，避免复制二进制与解析器后产生体积浪费、安全分叉和兼容性漂移；WPS-only pack 不会为非 WPS 候选提供加载能力。

部分 WPS 生产文件不会写出最后一个 stream sector 的未使用 padding，也可能追加 FAT 图完全不可达的私有尾部。解析器仅在被引用短 sector 已包含目录声明的全部 stream 字节时接受缺失 padding，或忽略完全不可达的尾部，并发出容器级近似诊断；少任何一个声明字节都仍然失败，不能用该兼容路径绕过边界检查。

## 当前覆盖与差距

首版正式包已经具备：

- `.wps` 常见文字、字体、字号、颜色、强调、段落对齐、缩进和间距的结构解析；
- `.et` 常见字符串、数值、布尔、错误和缓存公式值、sheet 名称与源记录位置；
- `.dps` live edit/persist 链、slide 文本、常见旧版绘图几何、源记录位置和 shape-bound WAV；
- CFB/FAT/DIFAT/miniFAT、主流可达性、记录边界、加密/宏/ActiveX/OLE/外链阻断与资源预算；
- 独立 npm 子路径与候选、共享 CFB Wasm、懒加载调度、浏览器 MIME、发布打包和端到端 fixture。

当前不能宣称覆盖“绝大多数排版和图形”：

- WPS Writer 的精确分页、页眉页脚、表格、字段和浮动绘图仍不完整；
- ET 的 XF/字体/填充/边框/数值格式、行列几何、合并、Drawing/Chart 仍主要缺失；
- DPS 的母版/版式、主题、复杂图形、表格/图表、动画和精确文字布局仍是旧版 PPT 的有限子集。

这些缺口会产生结构化 approximate/omitted 诊断，不能以“成功打开”或格式化输出掩盖。后续兼容开发应优先扩展共享 Word/BIFF/PowerPoint 记录层；只有 WPS 私有记录无法用共享语义表达时，才增加最小、受候选边界约束的专用适配。

## 2% 视觉门禁

WPS release suite 必须同时包含 `wps`、`et`、`dps`，并覆盖 `minimal`、`combination`、`enterprise`、`large` 四类经授权语料。每个文件的每一页、slide 或固定 A1 sheet viewport 都必须与相同 WPS Office 版本、字体集、语言、色彩空间和缩放下的人工审核金标比较。

锁定的 `native-visual-v1` 门槛为：

| 指标 | 每个 unit 的要求 |
| --- | ---: |
| 容差像素相似度 | ≥ 99.5% |
| windowed SSIM | ≥ 99.0% |
| 每个 16×16 对象区域相似度 | ≥ 99.0% |
| AccuracyReport | ≥ 95 |

因此门禁比单一“差异 ≤2%”更严格，而且不能用平均值掩盖失败区域。WPS Writer/DPS 金标只接受人工复核的原生 PDF 导入，ET 只接受人工复核的固定范围 content-only PNG；工具不会自动操作 WPS Office。fixture、reference、PDF/PNG、候选输出、应用版本和环境指纹均进入哈希链。

没有经授权的完整 WPS release corpus 和人工金标时，原生视觉聚合命令默认因缺少 `wps-release.json` 而硬失败。经明确授权的制品发布流程可以将“manifest 不存在”降为警告，但必须把 Release 标记为 `not-certified`，不得生成假金标、把现有解析测试当作视觉通过证据，或宣称满足 2% 目标；一旦 manifest 存在，任何校验失败仍阻断发布。

## 测试层次

1. Rust 单元测试：CFB 图、主流识别、Word/BIFF/PowerPoint 记录边界、WPS 尾部隔离、损坏/加密/活跃内容拒绝。
2. TypeScript/Node 合同测试：闭集候选、扩展名与内容联合调度、独立子路径和惰性单例加载。
3. Wasm 端到端测试：真实构建产物分别打开 deterministic WPS/ET/DPS fixtures，验证格式、种类、对象和源映射。
4. 本机兼容冒烟：只读打开已安装 WPS Office 自带的真实 `.wps/.et/.dps` 模板，确认复杂文件局部失败不拖垮整份文档。
5. 浏览器视觉发布门：WPS Office 审核金标、完整 unit 拓扑、固定环境和逐区域阈值；这是 2% 目标的唯一发布证明。

真实供应商模板或客户文件不签入仓库；进入长期 corpus 前必须明确授权、许可、脱敏、manifest 哈希和分层用途。
