# 经销商库存对账 Mismatch 智能诊断系统

这是一个 Windows 本地 GUI 工具，用于把 Dealer Stock Reconciliation 的 Excel 数据转成可读的智能诊断报告。

## 启动方式

双击：

`启动_经销商库存Mismatch智能诊断系统.bat`

或用 Python 运行：

```powershell
python dealer_mismatch_diagnostic_app.py
```

## 支持的 Excel Sheet

系统会自动识别以下工作表：

- `Summary`
- `Only_in_List`
- `Only_in_SAP`
- `Mismatch_Detail`
- `Stock_Base`
- `Actual feedback`

至少需要包含 `Mismatch_Detail`、`Only_in_List` 或 `Only_in_SAP` 之一，才能运行诊断。

## 已实现功能

- Excel 导入与结构检查
- Sheet 行数、字段数、空值率预览
- Chassis 标准化：去 `VIN:`、去前导零、统一大小写、统一 USED 横杠空格、去多余空格
- List/SAP 智能配对：标准化完全一致、编辑距离 1 的疑似假 mismatch
- 五层数据链诊断：命名、真假 mismatch、库存类型、生命周期证据、Actual feedback
- 业务可读输出：Assumption、Actual situation、Data-chain judgement、Evidence reason、Process issue、Optimisation、Owner、Priority、Confidence
- Actual feedback 回写
- 统计看板：标签、优先级、Owner、Stock type、Confidence 分布
- Excel 报告导出：淡红 High、淡黄 Medium、淡绿 Low、淡蓝灰表头，关键列自动 wrap text
- 诊断结果页支持底部横向滚动条和 `< / >` 左右移动按钮
- 诊断结果页支持鼠标拖拽蓝色框选择任意单元格区域，`Ctrl+C` 或点击“复制选区”后可直接粘贴到 Excel
- 诊断结果页每列都有筛选行，双击对应列的 `Filter...` 可输入包含式过滤条件；点击“清除筛选”恢复
- 导入 Excel 后会自动运行智能诊断
- `Process issues identified` 和 `How to optimise` 会自动前置到 Chassis 后面，方便业务人员先看问题和动作
- `Assumption` 列已从诊断结果和导出报告中移除
- 导出的 Excel 会对长文本列自动 wrap text，并按文字长度估算行高，减少手动拉伸
- 新增流程诊断包字段：`Process issue category`、`Root cause hypothesis`、`Control gap`、`Recommended control`、`Required evidence`、`Next action`、`Preventive rule`、`Linked example cases`
- 新增重点流程规则：退车/退款无凭证、PGI 闭环违规、预定/二次销售风险
- 借用旧版规则体系并扩展为 `M2`、`R0-R13`、`C2`：包含 SAP 无数据、存在性违反、PGI 后缺发票、占位料号、PO/GR 缺失、已完成/离场、3110 冲销退回等规则
- 新增 `Rule_ID` 和 `Diagnostic category`，保留旧版 Category 1 / Category 2 的诊断分层
- 流程诊断包现在会按最终判断自动补全，避免新增列大面积空白
- `Actual situation after check` 从诊断主表隐藏，改由“反馈录入”页写入并在导出报告中保留
- “反馈录入”页支持同时写入实际核实情况和手动流程化改进建议
- “规则说明”页支持手动新增规则、选中规则后调整说明、优先级、Owner 和优化建议

## 主要诊断标签

- `False Positive - Naming`
- `SAP No Evidence`
- `Likely Data Timing Delay`
- `Completed / Left Stock`
- `Cancelled Sale / Refund Not Closed`
- `Demo / Used Special Handling`
- `Process Break`
- `Needs Manual Review`

## 后续可打包成 exe

如果需要分发给没有 Python 环境的同事，可以用 PyInstaller 打包：

```powershell
pyinstaller --onefile --windowed dealer_mismatch_diagnostic_app.py
```

正式版本建议把规则拆到独立 `config/rules.json`，并按真实字段继续扩展规则库。
