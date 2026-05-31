# INTEREST-CALCULATOR-CN 利息计算器

专为中国大陆律师处理民间借贷案件设计的利息计算器。它支持多笔借款、多笔还款、法定利率上限、LPR 浮动利率、先息后本冲抵规则，并能导出包含逐条计算过程的 Excel 表格，方便复核和作为证据材料整理使用。

## 主要功能

- 多笔借款、还款流水录入
- 按日期顺序自动重算每期本金、利息和欠息
- 支持年利率、月利率输入，并自动取较高折算值
- 支持 2015-09-01、2020-08-20 等法定利率规则边界分段
- 支持 LPR 数据分段计息和 LPR 倍数上限
- 还款按“先息后本”处理
- 每条流水可查看完整计算过程，包含起止日期、天数、利率、公式和冲抵结果
- 导出 Excel 明细表，包含“计算过程”列，便于人工复验
- 联网更新 LPR 失败时自动使用缓存或内置数据

## 环境要求

- Python 3.8+
- tkinter
- tkcalendar
- openpyxl

安装依赖：

```bash
pip install tkcalendar openpyxl
```

## 运行

```bash
python interest_calculator.py
```

启动后录入本金、起始日期、利率，再添加借款或还款流水。点击“汇总结果”可刷新全部计算结果；选中任意流水后点击“计算过程”可查看该条的详细公式。

## LPR 数据

程序内置了一份 LPR 历史数据，并会优先读取本地缓存。缓存位置在用户本地应用数据目录下：

```text
%LOCALAPPDATA%\利息计算器\lpr_data.json
```

如需通过 Tushare 联网更新 LPR，请先设置环境变量：

```bash
set TUSHARE_TOKEN=你的Token
```

然后在软件中点击“更新LPR”。如果未设置 token 或联网失败，程序会继续使用本地缓存；缓存不存在时使用内置数据。

## Excel 导出

点击“导出Excel”可以生成结果表，内容包括：

- 借款本金、起始日期、利率设置
- 每条流水的金额、应还利息、实还利息、当期欠息、剩余本金
- 每条流水的详细计算过程
- 汇总后的剩余本金、欠付利息、尚欠本息

## 打包

使用 PyInstaller 打包：

```bash
pyinstaller interest_calculator.spec
```

或直接执行：

```bash
pyinstaller --onefile --windowed --name "利息计算器" --icon=None interest_calculator.py
```

## 验证

语法检查：

```bash
python -m py_compile interest_calculator.py
```

本项目已重点验证：

- LPR 会取不晚于目标日期的最近一期
- 跨 2020-08-20 的区间会自动拆段
- 跨 LPR 变动日的区间会自动拆段
- “计算过程”文本能生成公式和冲抵明细
- GUI 能完成初始化

## 注意

本工具用于辅助计算和复核，不替代律师、法院或金融机构对个案事实和法律适用的判断。正式提交材料前，请结合案件事实、合同约定和现行裁判规则进行人工复核。
