# Excel 公式支持矩阵

这份矩阵描述的是 Coco Office 当前公式计算引擎的产品边界，不代表 Excel 公式全集覆盖。目标是优先覆盖 Agent 高频生成、业务表格计算和常见数据清洗场景；未列出的函数应视为暂不承诺支持。

## 当前稳定支持

| 类别 | 函数 |
| --- | --- |
| 条件与错误 | `IF`, `IFERROR`, `IFNA`, `IFS`, `SWITCH`, `TRUE`, `FALSE`, `AND`, `OR`, `NOT`, `NA`, `N`, `ISERROR`, `ISERR`, `ISNA`, `ISBLANK`, `ISNUMBER`, `ISTEXT` |
| 汇总与统计 | `SUM`, `AVERAGE`, `COUNT`, `COUNTA`, `COUNTBLANK`, `MIN`, `MAX`, `MEDIAN`, `PRODUCT`, `SUMSQ`, `LARGE`, `SMALL`, `RANK`, `RANK.EQ`, `RANK.AVG`, `VAR`, `VAR.P`, `VAR.S`, `STDEV`, `STDEV.P`, `STDEV.S`, `PERCENTILE`, `PERCENTILE.INC`, `PERCENTILE.EXC`, `QUARTILE`, `QUARTILE.INC`, `QUARTILE.EXC` |
| 条件汇总 | `SUMIF`, `COUNTIF`, `AVERAGEIF`, `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, `MAXIFS`, `MINIFS` |
| 查找引用 | `VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`, `XLOOKUP`, `XMATCH` |
| 数学计算 | `ABS`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `INT`, `MOD`, `CEILING`, `FLOOR`, `SIGN`, `POWER`, `SQRT`, `EXP`, `LN`, `LOG` |
| 文本处理 | `LEN`, `TRIM`, `LOWER`, `UPPER`, `LEFT`, `RIGHT`, `MID`, `FIND`, `SEARCH`, `SUBSTITUTE`, `REPLACE`, `TEXT`, `CONCAT`, `CONCATENATE`, `TEXTJOIN`, `TEXTBEFORE`, `TEXTAFTER`, `VALUE` |
| 日期时间 | `TODAY`, `DATEDIF`, `DATE`, `DATEVALUE`, `TIME`, `YEAR`, `MONTH`, `DAY`, `DAYS`, `WEEKDAY`, `HOUR`, `MINUTE`, `SECOND`, `EDATE`, `EOMONTH`, `WORKDAY`, `NETWORKDAYS` |
| 动态数组 | `FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`, `SORTBY`, `TAKE`, `DROP`, `CHOOSECOLS`, `CHOOSEROWS`, `TOCOL`, `TOROW` |
| 财务 | `PMT`, `FV`, `PV`, `NPV`, `IRR`, `RATE` |
| 其他 | `SUMPRODUCT`, `SUBTOTAL`, `AGGREGATE`, `LET` |

## 有边界的支持

- 公式支持单元格引用、区域引用、绝对引用、同表引用和常见跨表引用；不等同于完整 Excel 名称解析。
- 动态数组会计算并投影到虚拟溢出单元格；溢出区域被已有内容占用时返回 `#SPILL!`。
- `SUBTOTAL` 和 `AGGREGATE` 当前依据工作表行的 `hidden` 状态处理隐藏行；筛选、分组和更复杂的 Excel 可见性语义仍由宿主投影层决定。
- `VLOOKUP`、`HLOOKUP`、`MATCH`、`XLOOKUP`、`XMATCH` 已覆盖常见精确匹配和基础排序参数；复杂近似匹配、错误处理和数组广播不作为当前完整兼容承诺。
- 日期计算使用 Excel serial date；`TODAY()` 可由宿主注入日期，测试场景应注入固定日期以保证可重复。
- 公式引擎负责计算真值，渲染器负责显示结果；单元格格式、条件格式、表格结构化引用和图表渲染不属于本矩阵。

## 暂不承诺

- 未实现的 Excel 函数会返回 `#NAME?`，而不是静默伪造结果。
- 当前不承诺完整支持 Excel 表格结构化引用、名称管理器的全部语义、外部工作簿引用、宏/UDF、数据透视表公式、立方体函数、工程函数、数组公式的全部兼容细节，以及所有本地化函数名。
- 真实文件的公式覆盖率必须通过 `report:formula-coverage` 审计，不用“支持函数数量 / Excel 函数总数”作为产品指标。

## 验证方式

核心公式单元测试覆盖解析、错误传播、跨表依赖、动态数组、财务/统计/日期函数和完整销售订单场景。对真实 `.xlsx` 文件可运行：

```bash
corepack pnpm --filter @coco-vexel/coco-office report:formula-coverage -- <Excel目录>
corepack pnpm --filter @coco-vexel/coco-office report:formula-coverage:check
```

`--check` 会在扫描不到文件、没有公式、公式产生错误、出现未解析公式或缓存值不一致时返回失败，适合接入 CI。

## 当前业务样例基线

2026-09-09 对本地业务样例目录做审计：扫描 4 个 `.xlsx` 文件，共发现 32 个公式单元格，函数分布为 `IF` 32 次、`SUMPRODUCT` 16 次，计算错误为 0。样例文件没有保存公式缓存值，因此本次没有可比较的缓存结果；这不影响公式重新计算结果，但不能据此证明与 Excel 缓存值逐格一致。
