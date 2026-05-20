# 利息计算器 - 构建说明

## 技术栈
- Python 3.8+
- tkinter（标准库）
- PyInstaller（打包）

## 依赖
- tkinter（Python内置）
- tkcalendar（pip install tkcalendar）
- urllib（Python内置）
- json（Python内置）

## 构建命令

```bash
pyinstaller --onefile --windowed --name "利息计算器" --icon=None interest_calculator.py
```

或使用 spec 文件方式：

```bash
pyinstaller interest_calculator.spec
```

## 运行
```bash
python interest_calculator.py
```

## LPR数据
- 程序启动时自动从 Tushare 获取最新LPR数据
- 缓存到本地 lpr_data.json
- 联网失败时读取缓存，缓存也不存在则使用内置默认值(3.45%)

## 注意事项
- tkinter DateEntry 组件需要 tkcalendar 库（pip install tkcalendar）
- 如果打包后运行异常，检查Python版本和tkinter安装
