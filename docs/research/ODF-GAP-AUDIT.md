# ODF 标准差距清单

对照 OASIS OpenDocument 1.2/1.3/1.4（Part 2 Packages / Part 3 Schema / Part 4 Formula）中与只读渲染相关的词汇，审计 `core/format/odt.rs`、`ods.rs`、`odp.rs`、`odf_chart.rs`、`odf_math.rs` 的实现面。标准文本以 [ODF 1.4 OS](https://docs.oasis-open.org/office/OpenDocument/v1.4/os/) 为准；测试资产见 `docs/research/ODF-TEST-ASSETS.md`。

约定：

- **已实现**：解析进共享文档模型，并产出可命中、可源映射的渲染对象。
- **近似**：保留可见内容与源映射，或给出结构化诊断后降级。
- **未实现**：标准有明确视觉/结构语义，当前适配层未映射或直接忽略。
- 范围外：纯编辑语义、宏/脚本执行、外部资源加载、动画时间轴播放（预览器只读，不承诺与编辑器一致）。

## 包与元数据

| 能力 | 状态 | 说明 |
| --- | --- | --- |
| ODF ZIP + `mimetype` + manifest | 已实现 | 含 flat ODT/ODS/ODP/FODx |
| 加密文档拒绝 / Scripts 阻断 | 已实现 | 不执行活动内容 |
| `meta.xml` 属性暴露 | 近似 | 打开信息为主，不完整映射 ODF 元数据词汇 |
| 数字签名验证 | 未实现 | 包内签名仅做安全隔离，不验签 |
| `settings.xml` 视图状态 | 部分 | ODS 冻结窗格；其余视图/打印/变更跟踪状态基本未映射 |

## 文本（ODT / office:text）

| 能力 | 状态 | 说明 |
| --- | --- | --- |
| 段落/标题/样式继承/字符样式 | 已实现 | |
| 列表（编号/项目/层级/续编） | 已实现 | `text:list-style` 含 bullet/number |
| 表格重复行列、跨行跨列、单元格样式、列宽 | 部分 | 列宽已支持 `style:column-width` / `style:rel-column-width`（回归：`authored_odt_table_column_widths_preserve_authored_ratios` 等）；无作者宽度时仍等分并诊断 |
| 分页、页边距、master 页眉页脚 | 已实现 | |
| 文本分栏（`style:columns`） | 部分 | 等宽多栏可用；不等宽、`dont-balance`、跨页断栏回退单栏 |
| 脚注/尾注 | 部分 | 可见且独立流；不贴页底、不重编号 |
| 批注 `office:annotation` | 部分 | 可见源映射；无气泡边注布局 |
| 修订 `text:tracked-changes` | 部分 | 删除内容隐藏；无接受/拒绝 UI，无完整插入/删除标记样式 |
| 字段（`text:page-number`、`page-count`、`date`、`time`、`author` 等） | 部分 | 空字段用 `office:*-value` 回填；`page-number`/`page-count` 按版面动态替换（含 master stories）。其余字段有缓存显示则用缓存；`sequence`/`simple-field`/TOC 未实现 |
| 超链接 `text:a` | 部分 | 可选中 URL；无完整注解/书签跳转 |
| 书签 `text:bookmark*`、交叉引用 `text:reference-ref` | 未实现 | |
| 目录/索引（TOC、图表目录、字母索引、文献） | 未实现 | 无 `text:table-of-content` 等索引体布局 |
| Ruby `text:ruby*` | 近似 | 基字 + 上标注音 run |
| 首字下沉 `style:drop-cap` | 已实现 | 映射共享 `TextDropCap` |
| 文本框/浮动框架/环绕 | 已实现 | 常见 wrap/anchor |
| 绘图对象 | 部分 | rect/ellipse/line/connector/caption/measure/regular-polygon/custom-shape；表内图片已支持 |
| 表单元格内图片/对象 | 部分 | 表内 `draw:image` 已渲染并源映射；嵌套对象仍未映射 |
| 表单控件 | 部分 | 表单文本框有映射；其余控件未映射 |
| 行号、条件段落样式、占位符、隐藏文字 | 未实现 | |

## 电子表格（ODS / office:spreadsheet）

| 能力 | 状态 | 说明 |
| --- | --- | --- |
| 工作表/重复行列/合并/样式/数值格式 | 已实现 | 常用 number/currency/percent/scientific/date/boolean |
| 分数格式 `number:fraction`、文本格式 `number:text-style` | 未实现 | |
| 条件格式 | 部分 | 数值比较、色阶、数据条、图标集；规则不执行 |
| 冻结窗格 | 已实现 | |
| 图表 | 部分 | bar/line/pie/scatter/area/radar；缺 stock/bubble/net/doughnut 独立语义 |
| 绘图 | 部分 | 图片帧与基础图形；复杂锚定/形状库不完整 |
| 打印设置 | 范围外 | 明确不实现；XLSX 侧 `SheetPrintSettings` 保持现状 |
| 数据验证 `table:content-validations` | 已实现 | 只读注解：条件与 help/error 元数据进诊断；宏/事件监听阻断；规则不执行 |
| 数据透视 `table:data-pilot-tables` | 未实现 | 无缓存结果单元格还原 |
| 数据库区域/筛选 `table:database-ranges` | 未实现 | |
| 分级显示/折叠 `table:table-source` 外表源 | 未实现 / 阻断 | 外部表源按外部资源策略拒绝 |
| 方案 `table:scenario`、名称表达式 | 未实现 | |
| 公式求值 | 近似 | 使用缓存值；无 ODF 公式重算（OpenFormula） |
| 备注/批注 | 未实现 | |

## 演示与绘图（ODP / ODG / office:presentation）

| 能力 | 状态 | 说明 |
| --- | --- | --- |
| 页面、母版、背景、讲者备注 | 已实现/部分 | 备注有格式化正文；无讲义页布局 |
| 形状几何 | 部分 | rect/roundrect/ellipse/line/path/polygon/polyline/connector/caption/custom-shape+enhanced-geometry；预设形状库小于 DrawingML 187 |
| 渐变/纹理/透明度/阴影/线帽/虚线/箭头 | 已实现/部分 | 双色渐变；多色标（多为 loext）不完整 |
| 表格 | 部分 | 列宽已支持；复杂表样式带不完整 |
| 图表 | 部分 | bar/line/scatter/pie |
| 公式 MathML | 近似 | 线性化文本 |
| `draw:plugin` 媒体 | 已实现 | 本地海报+控件，不自动播放 |
| 动画/切换/时间轴 `anim:*`/`smil:*` | 未实现 | 只读预览不播放 |
| 3D `dr3d:*` | 未实现 | |
| 自定义放映 `presentation:show` | 未实现 | |
| 图像映射 `draw:image-map` | 未实现 | |
| 粘合点/精确连接 | 部分 | 连接线有基础几何 |
| 形状内字段/演示者视图计时 | 未实现 | |

## 公式与图表共享层

| 能力 | 状态 | 说明 |
| --- | --- | --- |
| MathML 线性化 | 已实现 | `odf_math.rs` |
| 图表数据标签/图例/堆叠/坐标轴位置 | 部分 | 共享 `odf_chart.rs` |
| OpenFormula 求值 | 未实现 | 与 XLSX 计算服务语义不等价，不强行收敛 |

## 跨格式必须保持的实现约束

1. 渲染必须来自真实结构解析；禁止用包内缩略图/预览图兜底。
2. 局部失败隔离：未知元素、坏资源降级并诊断，不拖垮全文档。
3. 与 DOCX/PPTX/XLSX 严格等价的算法收敛到共同层；不等价差异留在适配层并写清边界（例：ODF `parse_length` 三份副本在负值与上限上不等价，暂不收敛）。
4. 每条功能合入必须附带“修改前失败、修改后通过”的真实 Case，且不引起既有通过 Case 的无解释变化。

## 实现优先级（按预览可感知收益）

1. ~~**P0** ODT 表格列宽（`style:column-width` / `style:rel-column-width`）~~ 已实现
2. ~~**P0** ODT 字段显示值（日期/页码/作者等，含空内容属性回填）~~ 已实现（sequence/TOC 仍待做）
3. ~~**P0** ODS 打印设置~~ **范围外，明确不实现**
4. ~~**P0** ODS 数据验证注解~~ 已实现
5. ~~**P1** ODS scatter/area/radar 图表；ODS 图标集条件格式~~ 已实现
6. ~~**P1** ODT `style:drop-cap`、`text:ruby`~~ 已实现（ruby 近似）
7. ~~**P1** ODT 表内图片；`draw:measure` / `draw:regular-polygon`~~ 已实现
8. **P1** 不等宽分栏；字段/TOC；修订标记样式
9. **P2** 数据透视缓存结果、筛选注解、动画只读占位诊断

## 本清单的更新方式

实现一条后，同步改状态表、补回归 Case 编号，并更新 `docs/SUPPORT.md` 对应限制句。禁止只改文档不改实现，或只改实现不留回归。
