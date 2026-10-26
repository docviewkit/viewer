# iWork 静态图表覆盖矩阵

版本基线：Apple Numbers / Pages / Keynote **15.4**，本机 macOS；盘点日期：2026-10-26。

**当前结论：静态图表支持尚未完整。** 已补上雷达、气泡、混合/双轴图、手动轴范围及 Pages 行内图表路径；8 类 3D 图表仍二维化显示，不能算作 3D 支持通过。已有图表也不能因为能够生成对象或有一张截图，就升级为视觉验收通过。

可追溯矩阵：[iwork-chart-coverage.json](../../tests/fixtures/iwork-chart-coverage.json)。每个类型分 Numbers、Keynote、Pages 行内、Pages 浮动四个适配边界，记录真实文件、页/工作表索引、原生几何验证与整图视觉验收状态。索引从 0 开始。

## 类型清单

清单直接来自 Numbers 15.4 的“Chart Type”选择器：14 类 2D、8 类 3D、4 类交互式。协议中的 26（3D Donut）不在该版本选择器中，单独保留为 reserved，不纳入 22 类静态图表分母；不能据此宣称支持或不支持 Apple 当前未开放的功能。

| 类型 | 当前实现 | 本轮真实证据 |
| --- | --- | --- |
| 2D 柱、条、线、面积、饼、堆积柱、堆积条、堆积面积、散点 | 已有原生静态子集，完整样式未验收 | Keynote 原生 `add chart` 创建的 9 张 2D 图；既有 Numbers 文件场景保持不变 |
| 2D 圆环 | 已有静态子集，完整样式未验收 | Numbers Charting Basics 的 Pie and Donut Charts 工作表 |
| 2D 雷达 | 新增原生多边形、填充、标签、图例 | Numbers 原生 PDF 顶点对照；Keynote、Pages 行内与浮动真实文件回归 |
| 2D 气泡 | 新增三元组缓存、坐标、面积缩放 | Numbers 原生 PDF 的 9 个中心/半径；Keynote、Pages 行内真实文件回归 |
| 2D 混合 | 新增按系列类型绘制 | Keynote、Pages 行内真实文件；Numbers 专属混合图样本仍待补充 |
| 2D 双轴 | 新增系列轴归属、独立范围及右侧标签/标题 | Numbers 原生 PDF 的 12 个降雨折线点位；Keynote、Pages 行内真实文件回归 |
| 3D 柱、条、线、面积、饼、堆积柱、堆积条、堆积面积 | **缺口：保留数据，二维降级** | Keynote 原生创建 8 张真实 3D 图，并导出 Apple 原生结果；运行时逐图发出 `IWORK_CHART_3D_APPROXIMATE`，不计入 3D 验收通过 |
| 交互式柱、条、散点、气泡 | 不在本次静态范围 | 不把静态快照计为交互式支持 |

`partial` 表示已有实现和部分证据；`unverified` 表示没有足够证据；`missing-3d-geometry` 是明确的实现缺口。只有所有适配边界及必需特性都通过，才能将一行标记为完整支持。

## 本轮证据与重现

```sh
npm run build:core
npm run build:js
npm run test:iwork-charts
# 完整性门禁：当前应失败，不能用于声称已完整支持。
npm run test:iwork-charts -- --require-complete
# 用本轮保留的修改前 Wasm 运行完全相同的真实文件回归。
IWORK_CHART_BASELINE=output/iwork-chart-coverage/baseline.wasm \
IWORK_CHART_OUTPUT=output/iwork-chart-coverage/red \
node scripts/test-iwork-charts.mjs
```

修改前源码基线为 `3ab6d90`，`core/format/iwork.rs` SHA-256 为 `1823b5efe3f7d6cfaf261bc2e1c8e11dd20fb265b51a5a9ea93038d4b7a2820f`。本任务开始时该文件含有 GPA 未提交修改，随后由其他工作提交为上述版本；本轮保留基线文件与该提交逐字节一致，没有回滚或混入 GPA 修复。

回归报告写入 `output/iwork-chart-coverage/{red,regression}/report.json`，保留具体失败项、Chromium/WebKit 图表区域截图和真实 Viewer 截图。测试包括原生文件哈希、原生 PDF 哈希、图表归属、原生几何、可见标签、彩色绘制及浏览器控制台。局部像素或对象数量通过不能替代整图视觉验收。

原生参考位于 `tests/fixtures/iwork-chart-oracles/{numbers,keynote,pages}/`，包含 Apple PDF 及原生导出工具生成的来源、应用版本/build、环境和文件哈希。Numbers 使用打印页；本模板每个工作表恰有一页，并固定 `(72,72)` PDF 点到工作表原点的变换，不能推广成任意 Numbers 文档的页/工作表映射。

Numbers 几何金标直接从原生 PDF 提取，不从本项目解析结果反推：

```sh
# 可选测试维护工具，记录的 PyMuPDF VersionBind 为 2.0；SDK 不依赖它。
python3 scripts/extract-iwork-chart-geometry.py \
  tests/fixtures/iwork-chart-oracles/numbers > /tmp/iwork-native-geometry.json
```

雷达顶点容差为 1 CSS 像素；气泡中心/半径、双轴点位及手动轴柱形各边容差为 1.5 CSS 像素，用于容纳原生 PDF 中已取整的绘制边界。这不是发布级整图相似度阈值，不能替代 `NATIVE-VISUAL-ACCEPTANCE.md` 的发布门禁。本轮没有创建或伪造 release 视觉认证清单。

## 实现收敛与差异边界

- 雷达多边形顶点算法已从 DrawingML 收敛到共同层 `radar_polygon`。原 ChartML 包装保留 0.42 的半径系数和原有方向，iWork 使用真实 plot frame 的半径与自己的旋转。共同几何不改变任一格式的布局策略。
- 同时复核了最新已提交的 ODS/ODT/legacy 路径：legacy 复用 DrawingML；ODS 雷达按全局最大分类数补零、使用 0.45 半径，并以显式 LineTo 而非 ClosePath 收尾，且先计算全局坐标再减去原点；ODT 多边形也保留开放子路径收尾。其边界、浮点结果和描边连接与本次共享函数不能严格证明等价，因此保持独立，未为了统一改变既有语义。
- iWork 三宿主复用同一图表模型、解析与图元绘制。Pages 行内图表复用已有行内对象的锚点定位，浮动图表继续走原生 drawable 路径。
- 气泡的默认可见最小尺寸和缩放边界与 ChartML 不同；当前默认面积缩放参数用本轮原生文件验证，不能把两者的整体布局或自定义缩放宣称为严格等价。
- 混合/双轴图保留 TSCH 每系列的类型与轴 ordinal；不能用单一图表类型替代。不同轴分别求范围，折线按分类槽中心与柱形对齐。
- 损坏或未知图表以局部诊断跳过；资源限制与无法隔离的结构安全错误仍向上传递。没有引入预览图最终兜底。
- 第三方 schema 仅用于隔离的字段含义核对；未复制其代码、生成文件或引入运行时依赖。验收金标始终来自 Apple 应用。

## 尚需补齐的验收项

手动轴 Case 为 `iwork-manual-axis.key`：Apple 原生刻度为 5、16、28、39、50。修改前显示 0–30；修复后保留 5–50 的范围、刻度精度，柱形在 plot frame 内裁切，并逐边对照原生 PDF。更多手动轴变体和自动格式继承仍未验收。

1. 8 类 3D 的相机/透视、系列深度、挤出、材质与光照，及三宿主真实样本。
2. 手动轴更多边界、对数/日期轴、数字格式、负值与稀疏缓存、百分比堆积、自定义标签、颜色/线型/标记、雷达旋转和曲线网格、气泡自定义尺寸与负尺寸等边界组合。
3. 所有类型在 Pages 两条路径和 Numbers/Keynote 上的完整适配证据；裁剪、变换和继承样式边界。
4. 与 Apple 原生结果的整图及对象区域视觉验收，固定字体/浏览器/环境；Firefox 发布前实机验收。自动化仍只运行 Chromium 和 WebKit。

这些项保留未验收状态，不能因本轮局部回归通过而宣布“iWork 图表完整支持”。

本阶段回归记录：修改后真实文件/原生几何/双浏览器检查 109 项通过；iWork feature 下 517 项 Rust 测试通过；Native reference/policy 21 项通过；9 份既有文件的完整 scene 与修改前逐字段一致。完整性门禁仍保持失败。

## 静态边界补齐（继续进行中）

`iwork-chart-boundaries.key` 保存了 Keynote 15.4 原生导入的八个真实 Case：稀疏折线、正负堆积面积、百分比堆积柱、对数轴、百分比轴格式、稀疏面积、百分比堆积面积，以及含零/负值的对数轴。原生 PDF、来源哈希及直接提取的线段/矩形金标保存在 `iwork-chart-oracles/boundaries/`。路径回归同时检查原生顶点和候选线段中点，避免额外的零点折返或错误的面积闭合仍被顶点子集检查放行。

修复保留了 TSCH 的特有边界：百分比堆积按绝对值归一化，零合计为零；折线在缺失缓存处断开，面积图连接可用数据；堆积面积沿上一系列闭合并按原生层次绘制；对数轴跳过非正值，使用十进制幂刻度。分类轴读取原生“plot to edges”设置，自动轴保留显式小数位/百分比格式。原生面积透明度来自系列颜色，取消强制 0.55 的覆盖。

对数归一化坐标收敛到共同层，DrawingML 保留其既有 0..1 裁切。TSCH 的范围、刻度和百分比堆积策略与 ChartML 不同，保持适配层独立。复核最新已提交和本次未提交代码：ODS 的对数计算用于回归拟合，不能与坐标归一化合并；DrawingML 次刻度也复用共同归一化函数，仍保留其刻度枚举策略。

另外，两份既有 Numbers 文件的原生 PDF 证明圆环标题应与圆环中心对齐，不能受未显示的值轴文字宽度影响。`existing-multiple`、`existing-dice` 保留标题中心金标；`existing-team` 验证原生整数刻度 0、2、5、7、9；`existing-retirement` 验证面积不透明度。Numbers 打印页证据只用于这些固定对象特征，不能升级为全工作表视觉验收。

修改前基线 `3e27032`，修改后 192 项真实文件/原生几何/双浏览器检查通过。具体修改前失败项及哈希见 `iwork-chart-oracles/boundary-regression-evidence.json`。这些修复没有将任何完整类型或完整特性组改成通过，3D 及完整视觉验收继续进行。

原生整页门禁增加了 1280×960 Apple 幻灯片与自身比较的真实回归：修复前在第 53 个区域误报 `OBJECT_REGION_LIMIT`，修复后全部 256 个区域通过。250,000 像素限制作用于单区域工作集，区域数量及 128,000,000 次双向邻域运算预算保持不变；视觉阈值保持不变。共享准确性、协议及原生门禁 61 项测试通过，证据见 `iwork-chart-oracles/static-keynote/region-budget-regression.json`。这个门禁修复不代表任何图表类型已经通过原生视觉验收。
