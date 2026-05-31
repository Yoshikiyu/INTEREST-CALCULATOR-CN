"""
利息计算器 - 多轮还款明细表
分段计息，支持LPR浮动利率
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
from datetime import datetime, date
import calendar
import json
import os
import urllib.request
import re
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# 常量定义
DATE_START_LEGAL = date(2015, 9, 1)
DATE_END_LEGAL = date(2020, 8, 19)
DATE_START_LPR = date(2020, 8, 20)

# 法定利率上限
RATE_CAP_PRE_2015 = 0.24
RATE_CAP_2015_TO_2020_FULL = 0.36
RATE_CAP_2015_TO_2020_PARTIAL = 0.24

# 内置默认LPR值（当网络获取失败时使用）
DEFAULT_LPR = 3.45  # 3.45% (percentage format)

# 内置LPR历史数据（2019年8月至今主要记录）
BUILTIN_LPR_DATA = [
    {'date': '2019-08-20', 'lpr': 4.31},
    {'date': '2019-09-20', 'lpr': 4.20},
    {'date': '2019-10-21', 'lpr': 4.20},
    {'date': '2019-11-20', 'lpr': 4.15},
    {'date': '2019-12-20', 'lpr': 4.15},
    {'date': '2020-01-20', 'lpr': 4.15},
    {'date': '2020-02-20', 'lpr': 4.05},
    {'date': '2020-03-20', 'lpr': 4.05},
    {'date': '2020-04-20', 'lpr': 3.85},
    {'date': '2020-05-20', 'lpr': 3.85},
    {'date': '2020-06-22', 'lpr': 3.85},
    {'date': '2020-07-20', 'lpr': 3.85},
    {'date': '2020-08-20', 'lpr': 3.85},
    {'date': '2020-09-21', 'lpr': 3.85},
    {'date': '2020-10-20', 'lpr': 3.85},
    {'date': '2020-11-20', 'lpr': 3.85},
    {'date': '2020-12-21', 'lpr': 3.85},
    {'date': '2021-01-20', 'lpr': 3.85},
    {'date': '2021-02-20', 'lpr': 3.85},
    {'date': '2021-03-22', 'lpr': 3.85},
    {'date': '2021-04-20', 'lpr': 3.85},
    {'date': '2021-05-20', 'lpr': 3.85},
    {'date': '2021-06-21', 'lpr': 3.85},
    {'date': '2021-07-20', 'lpr': 3.85},
    {'date': '2021-08-20', 'lpr': 3.85},
    {'date': '2021-09-22', 'lpr': 3.85},
    {'date': '2021-10-20', 'lpr': 3.85},
    {'date': '2021-11-22', 'lpr': 3.85},
    {'date': '2021-12-20', 'lpr': 3.80},
    {'date': '2022-01-20', 'lpr': 3.70},
    {'date': '2022-02-21', 'lpr': 3.70},
    {'date': '2022-03-21', 'lpr': 3.70},
    {'date': '2022-04-20', 'lpr': 3.70},
    {'date': '2022-05-20', 'lpr': 3.70},
    {'date': '2022-06-20', 'lpr': 3.70},
    {'date': '2022-07-20', 'lpr': 3.70},
    {'date': '2022-08-22', 'lpr': 3.65},
    {'date': '2022-09-20', 'lpr': 3.65},
    {'date': '2022-10-20', 'lpr': 3.65},
    {'date': '2022-11-21', 'lpr': 3.65},
    {'date': '2022-12-20', 'lpr': 3.65},
    {'date': '2023-01-20', 'lpr': 3.65},
    {'date': '2023-02-20', 'lpr': 3.65},
    {'date': '2023-03-20', 'lpr': 3.65},
    {'date': '2023-04-20', 'lpr': 3.65},
    {'date': '2023-05-22', 'lpr': 3.65},
    {'date': '2023-06-20', 'lpr': 3.55},
    {'date': '2023-07-20', 'lpr': 3.55},
    {'date': '2023-08-21', 'lpr': 3.45},
    {'date': '2023-09-20', 'lpr': 3.45},
    {'date': '2023-10-20', 'lpr': 3.45},
    {'date': '2023-11-20', 'lpr': 3.45},
    {'date': '2023-12-20', 'lpr': 3.45},
    {'date': '2024-01-22', 'lpr': 3.45},
    {'date': '2024-02-20', 'lpr': 3.45},
    {'date': '2024-03-20', 'lpr': 3.45},
    {'date': '2024-04-22', 'lpr': 3.45},
    {'date': '2024-05-20', 'lpr': 3.45},
    {'date': '2024-06-20', 'lpr': 3.45},
    {'date': '2024-07-22', 'lpr': 3.35},
    {'date': '2024-08-20', 'lpr': 3.35},
    {'date': '2024-09-20', 'lpr': 3.35},
    {'date': '2024-10-21', 'lpr': 3.35},
    {'date': '2024-11-20', 'lpr': 3.35},
    {'date': '2024-12-20', 'lpr': 3.35},
    {'date': '2025-01-20', 'lpr': 3.35},
    {'date': '2025-02-20', 'lpr': 3.35},
    {'date': '2025-03-20', 'lpr': 3.35},
    {'date': '2025-04-21', 'lpr': 3.35},
    {'date': '2025-05-20', 'lpr': 3.35},
]

# LPR数据文件
LPR_DATA_FILE = "lpr_data.json"


def get_date_from_widget(widget):
    """从DateEntry widget获取date对象"""
    if hasattr(widget, 'get_date'):
        return widget.get_date()
    return None


def days_between(d1, d2):
    """计算两个日期之间的天数"""
    return (d2 - d1).days


def get_rate_for_period(start_date, end_date, user_rate, use_lpr, lpr_multiplier, lpr_data):
    """
    根据时间段获取适用的实际利率
    规则：实际利率 = min(用户输入利率, 法定上限)
    """
    # 判断时间段适用哪个法定上限
    cap = get_legal_cap(start_date, end_date)

    # 如果使用LPR，需要计算LPR×倍数
    if use_lpr:
        effective_lpr_rate = get_lpr_rate_for_date(end_date, lpr_data) * lpr_multiplier
        # LPR模式的法定上限是LPR×4
        cap = min(cap, effective_lpr_rate)

    # 实际利率 = min(用户输入, 法定上限)
    return min(user_rate, cap)


def get_legal_cap(start_date, end_date):
    """
    根据时间段判断法定利率上限
    判断逻辑基于该段的起点和终点
    """
    # 起点在2015-09-01之前，终点任意 -> 24%
    if start_date < DATE_START_LEGAL:
        return RATE_CAP_PRE_2015

    # 起点在2015-09-01~2020-08-19区间
    if DATE_START_LEGAL <= start_date <= DATE_END_LEGAL:
        # 终点也在区间内 -> 36%
        if end_date <= DATE_END_LEGAL:
            return RATE_CAP_2015_TO_2020_FULL
        # 终点超出区间 -> 24%
        else:
            return RATE_CAP_2015_TO_2020_PARTIAL

    # 起点在2020-08-20之后 -> LPR×4（在调用处动态计算）
    if start_date >= DATE_START_LPR:
        return 999  # 临时占位，实际在get_rate_for_period中重新计算

    return RATE_CAP_PRE_2015  # 默认


def get_lpr_rate_for_date(target_date, lpr_data):
    """获取指定日期适用的LPR值（找最近的但不晚于该日期的LPR）"""
    if not lpr_data:
        return DEFAULT_LPR / 100  # 转换为小数

    # 按日期升序排列，找到最近的不晚于目标日期的LPR
    applicable_lpr = DEFAULT_LPR / 100
    for record in lpr_data:
        record_date = datetime.strptime(record['date'], '%Y-%m-%d').date()
        if record_date <= target_date:
            applicable_lpr = record['lpr'] / 100  # 转换为小数
            break

    return applicable_lpr


def fetch_lpr_data():
    """
    获取LPR历史数据
    使用Tushare Pro API: shibor_lpr
    """
    try:
        import urllib.request
        import json

        url = 'https://api.tushare.pro/'
        params = {
            'api_name': 'shibor_lpr',
            'token': 'b9c020b005ffd77055b895c04bd9494763d04b9bed0a73b3f8b3d83d',
            'params': {'start_date': '20190801', 'end_date': '20261231'}
        }

        req = urllib.request.Request(
            url,
            data=json.dumps(params).encode('utf-8'),
            headers={'Content-Type': 'application/json', 'User-Agent': 'Mozilla/5.0'},
            method='POST'
        )

        with urllib.request.urlopen(req, timeout=30) as response:
            result = json.loads(response.read().decode('utf-8'))
            items = result.get('data', {}).get('items', [])

            lpr_list = []
            for item in items:
                # Without fields param: ['20190820', 4.25, 4.85]
                date_val = item[0]
                lpr_val = item[1]
                if (isinstance(date_val, str) and len(date_val) == 8 and
                    lpr_val is not None and isinstance(lpr_val, (int, float))):
                    lpr_list.append({
                        'date': f'{date_val[:4]}-{date_val[4:6]}-{date_val[6:8]}',
                        'lpr': float(lpr_val)
                    })

            lpr_list.sort(key=lambda x: x['date'])
            if lpr_list:
                print(f"Tushare获取LPR成功: {len(lpr_list)}条记录")
                return lpr_list

    except Exception as e:
        print(f"Tushare获取LPR失败: {e}")

    return None


def load_lpr_from_cache():
    """从本地文件加载LPR数据"""
    if os.path.exists(LPR_DATA_FILE):
        try:
            with open(LPR_DATA_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return None


def save_lpr_to_cache(lpr_data):
    """保存LPR数据到本地文件"""
    try:
        with open(LPR_DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(lpr_data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"保存LPR数据失败: {e}")


class InterestCalculatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("利息计算器")
        self.root.geometry("900x700")
        self.root.resizable(True, True)

        # LPR数据
        self.lpr_data = None
        self.load_or_fetch_lpr_data()

        # 样式
        self.setup_styles()

        # 变量
        self.principal_var = tk.StringVar()
        self.start_date_var = tk.StringVar()
        self.annual_rate_var = tk.StringVar()
        self.monthly_rate_var = tk.StringVar()
        self.use_lpr_var = tk.BooleanVar()
        self.lpr_multiplier_var = tk.StringVar(value="4")

        # 创建界面
        self.create_widgets()

    def setup_styles(self):
        self.style = ttk.Style()
        self.style.configure('Title.TLabel', font=('Arial', 12, 'bold'))
        self.style.configure('Result.TLabel', font=('Arial', 11))
        self.style.configure('Input.TEntry', font=('Arial', 11))

    def load_or_fetch_lpr_data(self):
        """加载或获取LPR数据"""
        # 先尝试从本地加载缓存数据
        self.lpr_data = load_lpr_from_cache()

        # 尝试联网获取最新数据
        online_data = fetch_lpr_data()
        if online_data:
            self.lpr_data = online_data
            save_lpr_to_cache(online_data)
        # 如果获取失败，保留本地缓存数据（已在self.lpr_data中）
        # 如果本地也没有，使用内置历史数据
        if not self.lpr_data:
            self.lpr_data = BUILTIN_LPR_DATA.copy()

    def update_lpr_data(self):
        """手动更新LPR数据"""
        online_data = fetch_lpr_data()
        if online_data:
            self.lpr_data = online_data
            save_lpr_to_cache(online_data)
            # 更新显示
            latest_lpr = self.lpr_data[-1]
            lpr_info = f"当前LPR(1年期): {latest_lpr.get('lpr', 3.45)}% ({latest_lpr.get('date', '')})"
            self.lpr_info_label.config(text=lpr_info)
            messagebox.showinfo("成功", "LPR数据已更新")
        else:
            # 获取失败，保留原有数据，显示警告
            if self.lpr_data:
                latest_lpr = self.lpr_data[-1]
                lpr_info = f"当前LPR(1年期): {latest_lpr.get('lpr', 3.45)}% ({latest_lpr.get('date', '')}) [缓存]"
                self.lpr_info_label.config(text=lpr_info)
            messagebox.showwarning("警告", "无法获取最新LPR数据，已使用本地缓存数据")

    def create_widgets(self):
        """创建所有界面组件"""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 借款信息区域
        self.create_loan_info_frame(main_frame)

        # 利率设置区域
        self.create_rate_frame(main_frame)

        # 还款明细表区域
        self.create_repayment_frame(main_frame)

        # 按钮区域
        self.create_button_frame(main_frame)

        # 汇总输出区域
        self.create_summary_frame(main_frame)

    def create_loan_info_frame(self, parent):
        """借款信息区域"""
        frame = ttk.LabelFrame(parent, text="借款信息", padding="10")
        frame.pack(fill=tk.X, pady=(0, 10))

        row = ttk.Frame(frame)
        row.pack(fill=tk.X)

        ttk.Label(row, text="本金:").pack(side=tk.LEFT, padx=(0, 5))
        self.principal_entry = ttk.Entry(row, textvariable=self.principal_var, width=15, style='Input.TEntry')
        self.principal_entry.pack(side=tk.LEFT, padx=(0, 20))

        ttk.Label(row, text="起始日期:").pack(side=tk.LEFT, padx=(0, 5))
        self.start_date_entry = DateEntry(row, width=12, dateformat='%Y-%m-%d')
        self.start_date_entry.pack(side=tk.LEFT)

    def create_rate_frame(self, parent):
        """利率设置区域"""
        frame = ttk.LabelFrame(parent, text="利率设置", padding="10")
        frame.pack(fill=tk.X, pady=(0, 10))

        row1 = ttk.Frame(frame)
        row1.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(row1, text="年利率(%):").pack(side=tk.LEFT, padx=(0, 5))
        self.annual_rate_entry = ttk.Entry(row1, textvariable=self.annual_rate_var, width=10, style='Input.TEntry')
        self.annual_rate_entry.pack(side=tk.LEFT, padx=(0, 20))

        ttk.Label(row1, text="月利率(%):").pack(side=tk.LEFT, padx=(0, 5))
        self.monthly_rate_entry = ttk.Entry(row1, textvariable=self.monthly_rate_var, width=10, style='Input.TEntry')
        self.monthly_rate_entry.pack(side=tk.LEFT)

        row2 = ttk.Frame(frame)
        row2.pack(fill=tk.X)

        self.use_lpr_check = ttk.Checkbutton(row2, text="使用LPR利率", variable=self.use_lpr_var)
        self.use_lpr_check.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(row2, text="倍数:").pack(side=tk.LEFT, padx=(0, 5))
        self.lpr_multiplier_entry = ttk.Entry(row2, textvariable=self.lpr_multiplier_var, width=5, style='Input.TEntry')
        self.lpr_multiplier_entry.pack(side=tk.LEFT)

        # LPR信息显示
        if self.lpr_data:
            latest_lpr = self.lpr_data[-1] if self.lpr_data else {'lpr': 3.45, 'date': '2024-12-20'}
            lpr_info = f"当前LPR(1年期): {latest_lpr.get('lpr', 3.45)}% ({latest_lpr.get('date', '')})"
            self.lpr_info_label = ttk.Label(row2, text=lpr_info, foreground='blue')
            self.lpr_info_label.pack(side=tk.LEFT, padx=(20, 10))

        ttk.Button(row2, text="更新LPR", command=self.update_lpr_data).pack(side=tk.LEFT)

    def create_repayment_frame(self, parent):
        """还款明细表区域"""
        frame = ttk.LabelFrame(parent, text="资金流水明细", padding="10")
        frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 创建表格
        columns = ('序号', '日期', '类型', '金额', '应还利息', '实还利息', '当期欠息', '当期剩余本金')
        self.repayment_tree = ttk.Treeview(frame, columns=columns, show='headings', height=8)

        # 设置列宽
        self.repayment_tree.heading('序号', text='序号')
        self.repayment_tree.heading('日期', text='日期')
        self.repayment_tree.heading('类型', text='类型')
        self.repayment_tree.heading('金额', text='金额')
        self.repayment_tree.heading('应还利息', text='应还利息')
        self.repayment_tree.heading('实还利息', text='实还利息')
        self.repayment_tree.heading('当期欠息', text='当期欠息')
        self.repayment_tree.heading('当期剩余本金', text='当期剩余本金')

        self.repayment_tree.column('序号', width=45, anchor='center')
        self.repayment_tree.column('日期', width=110, anchor='center')
        self.repayment_tree.column('类型', width=50, anchor='center')
        self.repayment_tree.column('金额', width=100, anchor='e')
        self.repayment_tree.column('应还利息', width=100, anchor='e')
        self.repayment_tree.column('实还利息', width=100, anchor='e')
        self.repayment_tree.column('当期欠息', width=100, anchor='e')
        self.repayment_tree.column('当期剩余本金', width=120, anchor='e')

        # 滚动条
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.repayment_tree.yview)
        self.repayment_tree.configure(yscrollcommand=scrollbar.set)

        self.repayment_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 存储还款行数据
        self.repayment_rows = []

    def create_button_frame(self, parent):
        """按钮区域"""
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(frame, text="添加还款", command=self.add_repayment_row).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(frame, text="添加借款", command=self.add_borrowing_row).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(frame, text="修改", command=self.edit_repayment_row).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(frame, text="删除", command=self.delete_repayment_row).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(frame, text="清空", command=self.clear_repayments).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(frame, text="导出Excel", command=self.export_to_excel).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(frame, text="汇总结果", command=self.calculate).pack(side=tk.RIGHT)

    def create_summary_frame(self, parent):
        """汇总输出区域"""
        frame = ttk.LabelFrame(parent, text="汇总输出", padding="10")
        frame.pack(fill=tk.X)

        row = ttk.Frame(frame)
        row.pack(fill=tk.X)

        ttk.Label(row, text="剩余本金:", font=('Arial', 11, 'bold')).pack(side=tk.LEFT, padx=(0, 5))
        self.remaining_principal_label = ttk.Label(row, text="0.00", font=('Arial', 11), width=15)
        self.remaining_principal_label.pack(side=tk.LEFT, padx=(0, 30))

        ttk.Label(row, text="欠付利息:", font=('Arial', 11, 'bold')).pack(side=tk.LEFT, padx=(0, 5))
        self.arrears_interest_label = ttk.Label(row, text="0.00", font=('Arial', 11), width=15)
        self.arrears_interest_label.pack(side=tk.LEFT, padx=(0, 30))

        ttk.Label(row, text="尚欠本息:", font=('Arial', 11, 'bold')).pack(side=tk.LEFT, padx=(0, 5))
        self.total_arrears_label = ttk.Label(row, text="0.00", font=('Arial', 11), width=15)
        self.total_arrears_label.pack(side=tk.LEFT)

    def add_repayment_row(self):
        """添加一行还款记录 - 弹出输入对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("添加还款记录")
        dialog.geometry("300x150")
        dialog.transient(self.root)
        dialog.grab_set()

        # 居中
        dialog.geometry("+{}+{}".format(
            self.root.winfo_x() + self.root.winfo_width() // 2 - 150,
            self.root.winfo_y() + self.root.winfo_height() // 2 - 75
        ))

        # 日期
        ttk.Label(dialog, text="还款日期:").grid(row=0, column=0, padx=10, pady=10, sticky='e')
        date_entry = DateEntry(dialog, width=12, dateformat='%Y-%m-%d')
        date_entry.grid(row=0, column=1, padx=10, pady=10)

        # 金额
        ttk.Label(dialog, text="还款金额:").grid(row=1, column=0, padx=10, pady=10, sticky='e')
        amount_var = tk.StringVar()
        amount_entry = ttk.Entry(dialog, textvariable=amount_var, width=15)
        amount_entry.grid(row=1, column=1, padx=10, pady=10)

        def on_confirm():
            try:
                repayment_date = date_entry.get_date()
                amount = float(amount_var.get()) if amount_var.get() else 0

                if amount < 0:
                    messagebox.showerror("输入错误", "还款金额不能为负数", parent=dialog)
                    return

                # 存储行数据
                row_data = {
                    'num': len(self.repayment_rows) + 1,
                    'date': repayment_date,
                    'type': 'repay',
                    'amount': amount,
                    'interest_due': 0,
                    'interest_paid': 0,
                    'current_arrears': 0,
                    'remaining_principal': 0
                }
                self.repayment_rows.append(row_data)

                # 重新排序并更新序号
                self.repayment_rows.sort(key=lambda x: x['date'])
                for i, row in enumerate(self.repayment_rows):
                    row['num'] = i + 1

                # 自动计算并显示
                self.auto_calculate_row(row_data)
                self.refresh_treeview()
                dialog.destroy()
            except ValueError:
                messagebox.showerror("输入错误", "请输入有效的金额", parent=dialog)

        def on_cancel():
            dialog.destroy()

        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=10)
        ttk.Button(btn_frame, text="确定", command=on_confirm).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=on_cancel).pack(side=tk.LEFT, padx=5)

        amount_entry.focus()
        dialog.bind('<Return>', lambda e: on_confirm())
        dialog.bind('<Escape>', lambda e: on_cancel())

    def add_borrowing_row(self):
        """添加一笔借款记录 - 弹出输入对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("添加借款记录")
        dialog.geometry("300x150")
        dialog.transient(self.root)
        dialog.grab_set()

        # 居中
        dialog.geometry("+{}+{}".format(
            self.root.winfo_x() + self.root.winfo_width() // 2 - 150,
            self.root.winfo_y() + self.root.winfo_height() // 2 - 75
        ))

        # 日期
        ttk.Label(dialog, text="借款日期:").grid(row=0, column=0, padx=10, pady=10, sticky='e')
        date_entry = DateEntry(dialog, width=12, dateformat='%Y-%m-%d')
        date_entry.grid(row=0, column=1, padx=10, pady=10)

        # 金额
        ttk.Label(dialog, text="借款金额:").grid(row=1, column=0, padx=10, pady=10, sticky='e')
        amount_var = tk.StringVar()
        amount_entry = ttk.Entry(dialog, textvariable=amount_var, width=15)
        amount_entry.grid(row=1, column=1, padx=10, pady=10)

        def on_confirm():
            try:
                borrow_date = date_entry.get_date()
                amount = float(amount_var.get()) if amount_var.get() else 0

                if amount <= 0:
                    messagebox.showerror("输入错误", "借款金额必须大于0", parent=dialog)
                    return

                row_data = {
                    'num': len(self.repayment_rows) + 1,
                    'date': borrow_date,
                    'type': 'borrow',
                    'amount': amount,
                    'interest_due': 0,
                    'interest_paid': 0,
                    'current_arrears': 0,
                    'remaining_principal': 0
                }
                self.repayment_rows.append(row_data)

                # 重新排序并更新序号
                self.repayment_rows.sort(key=lambda x: x['date'])
                for i, row in enumerate(self.repayment_rows):
                    row['num'] = i + 1

                # 重新计算所有行
                self.recalculate_all()
                dialog.destroy()
            except ValueError:
                messagebox.showerror("输入错误", "请输入有效的金额", parent=dialog)

        def on_cancel():
            dialog.destroy()

        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=10)
        ttk.Button(btn_frame, text="确定", command=on_confirm).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=on_cancel).pack(side=tk.LEFT, padx=5)

        amount_entry.focus()
        dialog.bind('<Return>', lambda e: on_confirm())
        dialog.bind('<Escape>', lambda e: on_cancel())

    def edit_repayment_row(self):
        """修改选中的记录（借款或还款）"""
        selection = self.repayment_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要修改的行")
            return

        # 获取选中的行
        item = selection[0]
        values = self.repayment_tree.item(item, 'values')
        row_num = int(values[0])

        # 找到对应的行数据
        row_data = None
        for row in self.repayment_rows:
            if row['num'] == row_num:
                row_data = row
                break

        if not row_data:
            messagebox.showerror("错误", "未找到对应的记录")
            return

        is_borrow = row_data.get('type') == 'borrow'
        type_label = "借款" if is_borrow else "还款"

        # 弹出修改对话框
        dialog = tk.Toplevel(self.root)
        dialog.title(f"修改{type_label}记录")
        dialog.geometry("300x150")
        dialog.transient(self.root)
        dialog.grab_set()

        # 居中
        dialog.geometry("+{}+{}".format(
            self.root.winfo_x() + self.root.winfo_width() // 2 - 150,
            self.root.winfo_y() + self.root.winfo_height() // 2 - 75
        ))

        # 日期
        ttk.Label(dialog, text=f"{type_label}日期:").grid(row=0, column=0, padx=10, pady=10, sticky='e')
        date_entry = DateEntry(dialog, width=12, dateformat='%Y-%m-%d')
        date_entry.set_date(row_data['date'])
        date_entry.grid(row=0, column=1, padx=10, pady=10)

        # 金额
        ttk.Label(dialog, text=f"{type_label}金额:").grid(row=1, column=0, padx=10, pady=10, sticky='e')
        amount_var = tk.StringVar(value=str(row_data['amount']))
        amount_entry = ttk.Entry(dialog, textvariable=amount_var, width=15)
        amount_entry.grid(row=1, column=1, padx=10, pady=10)

        def on_confirm():
            try:
                new_date = date_entry.get_date()
                amount = float(amount_var.get()) if amount_var.get() else 0

                if is_borrow and amount <= 0:
                    messagebox.showerror("输入错误", "借款金额必须大于0", parent=dialog)
                    return
                if not is_borrow and amount < 0:
                    messagebox.showerror("输入错误", "还款金额不能为负数", parent=dialog)
                    return

                # 更新数据
                row_data['date'] = new_date
                row_data['amount'] = amount

                # 重新排序并更新序号
                self.repayment_rows.sort(key=lambda x: x['date'])
                for i, row in enumerate(self.repayment_rows):
                    row['num'] = i + 1

                # 重新计算所有行
                self.recalculate_all()
                dialog.destroy()
            except ValueError:
                messagebox.showerror("输入错误", "请输入有效的金额", parent=dialog)

        def on_cancel():
            dialog.destroy()

        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=10)
        ttk.Button(btn_frame, text="确定", command=on_confirm).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=on_cancel).pack(side=tk.LEFT, padx=5)

        amount_entry.focus()
        amount_entry.select_range(0, tk.END)
        dialog.bind('<Return>', lambda e: on_confirm())
        dialog.bind('<Escape>', lambda e: on_cancel())

    def delete_repayment_row(self):
        """删除选中的记录"""
        selection = self.repayment_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要删除的行")
            return

        item = selection[0]
        values = self.repayment_tree.item(item, 'values')
        row_num = int(values[0])

        if messagebox.askyesno("确认", f"确定要删除第 {row_num} 条记录吗?"):
            # 找到并删除
            for i, row in enumerate(self.repayment_rows):
                if row['num'] == row_num:
                    self.repayment_rows.pop(i)
                    break

            # 重新排序并更新序号
            self.repayment_rows.sort(key=lambda x: x['date'])
            for i, row in enumerate(self.repayment_rows):
                row['num'] = i + 1

            # 重新计算所有行
            self.recalculate_all()

    def recalculate_all(self):
        """重新计算所有行的数据"""
        if not self.repayment_rows:
            return

        try:
            principal = float(self.principal_var.get())
        except ValueError:
            principal = 0

        start_date = self.start_date_entry.get_date()
        annual_rate = float(self.annual_rate_var.get()) / 100 if self.annual_rate_var.get() else 0
        monthly_rate = float(self.monthly_rate_var.get()) / 100 if self.monthly_rate_var.get() else 0
        user_rate = max(annual_rate, monthly_rate * 12)
        use_lpr = self.use_lpr_var.get()
        multiplier = float(self.lpr_multiplier_var.get()) if self.lpr_multiplier_var.get() else 4

        prev_date = start_date
        prev_remaining = principal
        brought_forward_arrears = 0

        for row in self.repayment_rows:
            # 计算本期应还利息（按当前剩余本金、上期日期到本期日期）
            interest_due = self.calculate_segment_interest(
                prev_date, row['date'], prev_remaining, user_rate, use_lpr, multiplier
            )
            total_interest_due = interest_due + brought_forward_arrears

            if row.get('type') == 'borrow':
                # 借款：利息全部挂账，本金增加
                row['interest_due'] = interest_due
                row['interest_paid'] = 0
                row['current_arrears'] = total_interest_due
                row['remaining_principal'] = prev_remaining + row['amount']
            else:
                # 还款：先息后本
                interest_paid = min(row['amount'], total_interest_due)
                principal_paid = row['amount'] - interest_paid
                current_arrears = total_interest_due - interest_paid
                remaining_principal = prev_remaining - principal_paid

                row['interest_due'] = interest_due
                row['interest_paid'] = interest_paid
                row['current_arrears'] = current_arrears
                row['remaining_principal'] = remaining_principal

            prev_date = row['date']
            prev_remaining = row['remaining_principal']
            brought_forward_arrears = row['current_arrears']

        self.refresh_treeview()
        self.update_summary()

    def refresh_treeview(self):
        """刷新treeview显示"""
        # 清空所有项
        for item in self.repayment_tree.get_children():
            self.repayment_tree.delete(item)

        # 重新插入所有行
        for row in self.repayment_rows:
            type_text = "借款" if row.get('type') == 'borrow' else "还款"
            values = (
                row['num'],
                row['date'].strftime('%Y-%m-%d') if row['date'] else '',
                type_text,
                f"{row['amount']:.2f}" if row['amount'] else '',
                f"{row['interest_due']:.2f}" if row['interest_due'] else '0.00',
                f"{row['interest_paid']:.2f}" if row['interest_paid'] else '0.00',
                f"{row['current_arrears']:.2f}" if row['current_arrears'] else '0.00',
                f"{row['remaining_principal']:.2f}" if row['remaining_principal'] else '0.00'
            )
            item = self.repayment_tree.insert('', 'end', values=values)
            # 借款行用蓝色标记
            if row.get('type') == 'borrow':
                self.repayment_tree.tag_configure('borrow', foreground='blue')
                self.repayment_tree.item(item, tags=('borrow',))

        # 更新汇总显示
        self.update_summary()

    def update_summary(self):
        """更新汇总显示"""
        if not self.repayment_rows:
            self.remaining_principal_label.config(text="0.00")
            self.arrears_interest_label.config(text="0.00")
            self.total_arrears_label.config(text="0.00")
            return

        last_row = self.repayment_rows[-1]
        remaining_principal = last_row['remaining_principal']
        arrears_interest = last_row['current_arrears']
        total_arrears = remaining_principal + arrears_interest

        self.remaining_principal_label.config(text=f"{remaining_principal:.2f}")
        self.arrears_interest_label.config(text=f"{arrears_interest:.2f}")
        self.total_arrears_label.config(text=f"{total_arrears:.2f}")

    def export_to_excel(self):
        """导出计算结果到Excel"""
        if not self.repayment_rows:
            messagebox.showwarning("警告", "没有可导出的记录")
            return

        # 获取保存路径
        file_path = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[('Excel文件', '*.xlsx'), ('所有文件', '*.*')],
            title='保存Excel文件'
        )
        if not file_path:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "利息计算结果"

            # 样式定义
            header_font = Font(bold=True, size=11)
            header_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            center_align = Alignment(horizontal='center', vertical='center')
            right_align = Alignment(horizontal='right', vertical='center')

            # 借款信息区域
            ws['A1'] = '借款本金'
            ws['B1'] = float(self.principal_var.get()) if self.principal_var.get() else 0
            ws['C1'] = '起始日期'
            ws['D1'] = self.start_date_entry.get_date().strftime('%Y-%m-%d') if self.start_date_entry.get_date() else ''

            ws['A2'] = '年利率(%)'
            ws['B2'] = float(self.annual_rate_var.get()) if self.annual_rate_var.get() else 0
            ws['C2'] = '月利率(%)'
            ws['D2'] = float(self.monthly_rate_var.get()) if self.monthly_rate_var.get() else 0
            ws['E2'] = '使用LPR'
            ws['F2'] = '是' if self.use_lpr_var.get() else '否'
            ws['G2'] = '倍数'
            ws['H2'] = float(self.lpr_multiplier_var.get()) if self.lpr_multiplier_var.get() else 4

            # 明细表标题行
            headers = ['序号', '日期', '类型', '金额', '应还利息', '实还利息', '当期欠息', '当期剩余本金']
            header_row = 4
            for col, header in enumerate(headers, start=1):
                cell = ws.cell(row=header_row, column=col, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.border = thin_border
                cell.alignment = center_align

            # 数据行
            for row_idx, row_data in enumerate(self.repayment_rows, start=header_row + 1):
                type_text = "借款" if row_data.get('type') == 'borrow' else "还款"

                ws.cell(row=row_idx, column=1, value=row_data['num']).border = thin_border
                ws.cell(row=row_idx, column=1).alignment = center_align
                ws.cell(row=row_idx, column=2, value=row_data['date'].strftime('%Y-%m-%d')).border = thin_border
                ws.cell(row=row_idx, column=2).alignment = center_align
                ws.cell(row=row_idx, column=3, value=type_text).border = thin_border
                ws.cell(row=row_idx, column=3).alignment = center_align
                ws.cell(row=row_idx, column=4, value=row_data['amount']).border = thin_border
                ws.cell(row=row_idx, column=4).alignment = right_align
                ws.cell(row=row_idx, column=4).number_format = '0.00'
                ws.cell(row=row_idx, column=5, value=row_data['interest_due']).border = thin_border
                ws.cell(row=row_idx, column=5).alignment = right_align
                ws.cell(row=row_idx, column=5).number_format = '0.00'
                ws.cell(row=row_idx, column=6, value=row_data['interest_paid']).border = thin_border
                ws.cell(row=row_idx, column=6).alignment = right_align
                ws.cell(row=row_idx, column=6).number_format = '0.00'
                ws.cell(row=row_idx, column=7, value=row_data['current_arrears']).border = thin_border
                ws.cell(row=row_idx, column=7).alignment = right_align
                ws.cell(row=row_idx, column=7).number_format = '0.00'
                ws.cell(row=row_idx, column=8, value=row_data['remaining_principal']).border = thin_border
                ws.cell(row=row_idx, column=8).alignment = right_align
                ws.cell(row=row_idx, column=8).number_format = '0.00'

            # 汇总行
            summary_row = header_row + len(self.repayment_rows) + 1
            ws.cell(row=summary_row, column=1, value='汇总').font = header_font
            ws.cell(row=summary_row, column=4, value=float(self.remaining_principal_label.cget("text"))).font = header_font
            ws.cell(row=summary_row, column=4).number_format = '0.00'
            ws.cell(row=summary_row, column=7, value=float(self.arrears_interest_label.cget("text"))).font = header_font
            ws.cell(row=summary_row, column=7).number_format = '0.00'
            ws.cell(row=summary_row, column=8, value=float(self.total_arrears_label.cget("text"))).font = header_font
            ws.cell(row=summary_row, column=8).number_format = '0.00'

            # 设置列宽
            ws.column_dimensions['A'].width = 8
            ws.column_dimensions['B'].width = 12
            ws.column_dimensions['C'].width = 8
            ws.column_dimensions['D'].width = 12
            ws.column_dimensions['E'].width = 12
            ws.column_dimensions['F'].width = 12
            ws.column_dimensions['G'].width = 12
            ws.column_dimensions['H'].width = 15

            wb.save(file_path)
            messagebox.showinfo("成功", f"已导出到:\n{file_path}")
        except Exception as e:
            messagebox.showerror("错误", f"导出失败:\n{str(e)}")

    def clear_repayments(self):
        """清空所有记录"""
        for item in self.repayment_tree.get_children():
            self.repayment_tree.delete(item)
        self.repayment_rows = []

    def validate_inputs(self):
        """验证输入"""
        try:
            principal = float(self.principal_var.get())
            if principal <= 0:
                messagebox.showerror("输入错误", "本金必须大于0")
                return False
        except ValueError:
            messagebox.showerror("输入错误", "请输入有效的本金金额")
            return False

        start_date = self.start_date_entry.get_date()
        if not start_date:
            messagebox.showerror("输入错误", "请选择起始日期")
            return False

        try:
            annual_rate = float(self.annual_rate_var.get()) / 100 if self.annual_rate_var.get() else 0
            monthly_rate = float(self.monthly_rate_var.get()) / 100 if self.monthly_rate_var.get() else 0
            user_rate = max(annual_rate, monthly_rate * 12)
            if user_rate <= 0:
                messagebox.showerror("输入错误", "请输入有效的利率")
                return False
        except ValueError:
            messagebox.showerror("输入错误", "请输入有效的利率值")
            return False

        if self.use_lpr_var.get():
            try:
                multiplier = float(self.lpr_multiplier_var.get())
                if multiplier <= 0:
                    messagebox.showerror("输入错误", "倍数必须大于0")
                    return False
            except ValueError:
                messagebox.showerror("输入错误", "请输入有效的倍数")
                return False

        # 检查事件记录
        valid_events = []
        for row in self.repayment_rows:
            if row['date'] and row['amount'] > 0:
                valid_events.append(row)

        if not valid_events:
            messagebox.showerror("输入错误", "请至少添加一笔借款或还款记录")
            return False

        # 按日期排序
        valid_events.sort(key=lambda x: x['date'])

        return True

    def auto_calculate_row(self, current_row):
        """
        自动计算单行数据（添加记录后立即计算）
        支持借款（本金增加）和还款（先息后本）
        """
        # 获取借款基本信息
        try:
            principal = float(self.principal_var.get())
        except ValueError:
            principal = 0

        start_date = self.start_date_entry.get_date()
        annual_rate = float(self.annual_rate_var.get()) / 100 if self.annual_rate_var.get() else 0
        monthly_rate = float(self.monthly_rate_var.get()) / 100 if self.monthly_rate_var.get() else 0
        user_rate = max(annual_rate, monthly_rate * 12)
        use_lpr = self.use_lpr_var.get()
        multiplier = float(self.lpr_multiplier_var.get()) if self.lpr_multiplier_var.get() else 4

        # 找到上一期记录
        current_index = self.repayment_rows.index(current_row)
        previous_row = self.repayment_rows[current_index - 1] if current_index > 0 else None

        # 确定起始日期和起始本金
        if previous_row:
            prev_date = previous_row['date']
            prev_remaining = previous_row['remaining_principal']
            brought_forward_arrears = previous_row['current_arrears']
        else:
            prev_date = start_date
            prev_remaining = principal
            brought_forward_arrears = 0

        # 计算本期应还利息（基于剩余本金）
        interest_due = self.calculate_segment_interest(
            prev_date, current_row['date'], prev_remaining, user_rate, use_lpr, multiplier
        )
        total_interest_due = interest_due + brought_forward_arrears

        if current_row.get('type') == 'borrow':
            # 借款：利息全部挂账，本金增加
            current_row['interest_due'] = interest_due
            current_row['interest_paid'] = 0
            current_row['current_arrears'] = total_interest_due
            current_row['remaining_principal'] = prev_remaining + current_row['amount']
        else:
            # 还款：先息后本
            interest_paid = min(current_row['amount'], total_interest_due)
            principal_paid = current_row['amount'] - interest_paid
            current_arrears = total_interest_due - interest_paid
            remaining_principal = prev_remaining - principal_paid

            current_row['interest_due'] = interest_due
            current_row['interest_paid'] = interest_paid
            current_row['current_arrears'] = current_arrears
            current_row['remaining_principal'] = remaining_principal

    def calculate_segment_interest(self, start_date, end_date, principal, user_rate, use_lpr, multiplier):
        """
        计算一段时间的利息（考虑分段）
        """
        # 收集所有分割点
        split_points = set([start_date, end_date])

        # 添加LPR变动点
        if use_lpr and self.lpr_data:
            for record in self.lpr_data:
                record_date = datetime.strptime(record['date'], '%Y-%m-%d').date()
                if start_date < record_date < end_date:
                    split_points.add(record_date)

        # 排序
        split_points = sorted(list(split_points))

        total_interest = 0
        remaining = principal

        for i in range(len(split_points) - 1):
            segment_start = split_points[i]
            segment_end = split_points[i + 1]
            days = days_between(segment_start, segment_end)

            if days <= 0:
                continue

            # 获取该段适用利率
            rate = get_rate_for_period(segment_start, segment_end, user_rate, use_lpr, multiplier, self.lpr_data)

            # 计算利息
            segment_interest = remaining * (rate / 360) * days
            total_interest += segment_interest

        return total_interest

    def calculate(self):
        """执行计算"""
        if not self.validate_inputs():
            return

        principal = float(self.principal_var.get())
        start_date = self.start_date_entry.get_date()
        annual_rate = float(self.annual_rate_var.get()) / 100 if self.annual_rate_var.get() else 0
        monthly_rate = float(self.monthly_rate_var.get()) / 100 if self.monthly_rate_var.get() else 0
        user_rate = max(annual_rate, monthly_rate * 12)
        use_lpr = self.use_lpr_var.get()
        multiplier = float(self.lpr_multiplier_var.get()) if self.lpr_multiplier_var.get() else 4

        # 获取有效事件记录并排序
        valid_events = []
        for row in self.repayment_rows:
            if row['date'] and row['amount'] > 0:
                valid_events.append(row)
        valid_events.sort(key=lambda x: x['date'])

        # 计算每笔事件
        current_principal = principal
        current_date = start_date
        brought_forward_arrears = 0

        for row in valid_events:
            event_date = row['date']
            event_amount = row['amount']

            # 计算应还利息（从上次日期到本次日期，基于当前剩余本金）
            interest_due = self.calculate_segment_interest(
                current_date, event_date, current_principal, user_rate, use_lpr, multiplier
            )
            total_interest_due = interest_due + brought_forward_arrears

            if row.get('type') == 'borrow':
                # 借款：利息全部挂账，本金增加
                row['interest_due'] = interest_due
                row['interest_paid'] = 0
                row['current_arrears'] = total_interest_due
                row['remaining_principal'] = current_principal + event_amount
            else:
                # 还款：先息后本
                interest_paid = min(event_amount, total_interest_due)
                principal_paid = event_amount - interest_paid
                current_arrears = total_interest_due - interest_paid
                remaining_principal = current_principal - principal_paid

                row['interest_due'] = interest_due
                row['interest_paid'] = interest_paid
                row['current_arrears'] = current_arrears
                row['remaining_principal'] = remaining_principal

            current_date = event_date
            current_principal = row['remaining_principal']
            brought_forward_arrears = row['current_arrears']

        # 更新treeview显示（包括汇总）
        self.refresh_treeview()


def main():
    root = tk.Tk()
    app = InterestCalculatorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()