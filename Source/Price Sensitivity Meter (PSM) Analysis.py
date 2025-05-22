import tkinter as tk
from tkinter import filedialog
import os
import pandas as pd
import matplotlib.pyplot as plt
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

# 定义语言字典
languages = {
    "zh": {
        "title": "价格敏感度测试模型分析",
        "select_button_text": "选择文件",
        "file_entry_placeholder": "请输入待分析 Excel 文件的完整路径",
        "analyze_button_text": "分析文件",
        "no_file_selected": "请选择有效的文件路径。",
        "file_not_exists": "文件不存在，请重新选择。",
        "analysis_error": "分析文件时出错: {}",
        "analysis_complete": "分析完成，结果已保存到 {}，PSM 图已保存。",
        "no_save_path_selected": "未选择保存路径，结果未保存。",
        "columns_stats": ["价格点", "太便宜比例", "便宜比例", "贵比例", "太贵比例"],
        "switch_language_button_text": "切换语言",
        "column_name_hint": "列名应为 TooCheap, Cheap, Expensive, TooExpensive"
    },
    "en": {
        "title": "Price Sensitivity Meter (PSM) Analysis",
        "select_button_text": "Select File",
        "file_entry_placeholder": "Please enter the full path of the Excel file to be analyzed",
        "analyze_button_text": "Analyze File",
        "no_file_selected": "Please select a valid file path.",
        "file_not_exists": "The file does not exist. Please select again.",
        "analysis_error": "An error occurred while analyzing the file: {}",
        "analysis_complete": "Analysis completed. The results have been saved to {}, and the PSM plot has been saved.",
        "no_save_path_selected": "No save path selected. The results were not saved.",
        "columns_stats": ["Price Point", "Too Cheap Ratio", "Cheap Ratio", "Expensive Ratio", "Too Expensive Ratio"],
        "switch_language_button_text": "Switch Language",
        "column_name_hint": "Column names should be TooCheap, Cheap, Expensive, TooExpensive"
    }
}

# 当前语言，默认为英文
current_language = "en"


def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, file_path)
        file_entry.config(foreground='black')


def on_entry_click(event):
    if file_entry.get() == languages[current_language]["file_entry_placeholder"]:
        file_entry.delete(0, tk.END)
        file_entry.config(foreground='black')


def on_focusout(event):
    if file_entry.get() == "":
        file_entry.insert(0, languages[current_language]["file_entry_placeholder"])
        file_entry.config(foreground='gray')


def analyze_file():
    file_path = file_entry.get()
    if file_path == languages[current_language]["file_entry_placeholder"]:
        result_label.config(text=languages[current_language]["no_file_selected"])
        return
    if not os.path.exists(file_path):
        result_label.config(text=languages[current_language]["file_not_exists"])
        return
    try:
        # 读取 Excel 文件
        df = pd.read_excel(file_path)

        # 假设数据集中包含 'TooCheap', 'Cheap', 'Expensive', 'TooExpensive' 列
        price_points = df.index
        too_cheap_ratio = df['TooCheap'] / df['TooCheap'].sum()
        cheap_ratio = df['Cheap'] / df['Cheap'].sum()
        expensive_ratio = df['Expensive'] / df['Expensive'].sum()
        too_expensive_ratio = df['TooExpensive'] / df['TooExpensive'].sum()

        # 计算交叉点
        indifference_point = None
        optimal_price_point = None
        lower_bound = None
        upper_bound = None

        for i in range(len(price_points) - 1):
            if cheap_ratio[i] < expensive_ratio[i] and cheap_ratio[i + 1] > expensive_ratio[i + 1]:
                indifference_point = price_points[i]
            if too_cheap_ratio[i] < too_expensive_ratio[i] and too_cheap_ratio[i + 1] > too_expensive_ratio[i + 1]:
                optimal_price_point = price_points[i]
            if too_cheap_ratio[i] < cheap_ratio[i] and too_cheap_ratio[i + 1] > cheap_ratio[i + 1]:
                lower_bound = price_points[i]
            if too_expensive_ratio[i] < expensive_ratio[i] and too_expensive_ratio[i + 1] > expensive_ratio[i + 1]:
                upper_bound = price_points[i]

        # 绘制 PSM 图
        plt.figure(figsize=(10, 6))
        plt.plot(price_points, too_cheap_ratio, label='Too Cheap')
        plt.plot(price_points, cheap_ratio, label='Cheap')
        plt.plot(price_points, expensive_ratio, label='Expensive')
        plt.plot(price_points, too_expensive_ratio, label='Too Expensive')

        if indifference_point:
            plt.axvline(x=indifference_point, color='r', linestyle='--',
                        label=f'Indifference Point: {indifference_point}')
        if optimal_price_point:
            plt.axvline(x=optimal_price_point, color='g', linestyle='--',
                        label=f'Optimal Price Point: {optimal_price_point}')
        if lower_bound:
            plt.axvline(x=lower_bound, color='b', linestyle='--', label=f'Lower Bound: {lower_bound}')
        if upper_bound:
            plt.axvline(x=upper_bound, color='m', linestyle='--', label=f'Upper Bound: {upper_bound}')

        plt.title('Price Sensitivity Meter (PSM)')
        plt.xlabel('Price')
        plt.ylabel('Ratio')
        plt.legend()

        psm_plot_path = os.path.splitext(file_path)[0] + '_psm_plot.png'
        plt.savefig(psm_plot_path)
        plt.close()

        # 保存结果到 DataFrame
        columns_stats = languages[current_language]["columns_stats"]
        data = {
            columns_stats[0]: price_points,
            columns_stats[1]: too_cheap_ratio,
            columns_stats[2]: cheap_ratio,
            columns_stats[3]: expensive_ratio,
            columns_stats[4]: too_expensive_ratio
        }
        result_df = pd.DataFrame(data)

        # 让用户选择保存路径
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            # 保存 DataFrame 到 Excel 文件
            result_df.to_excel(save_path, index=False)

            # 设置 wraplength 属性让文本自动换行
            result_label.config(text=languages[current_language]["analysis_complete"].format(save_path), wraplength=400)
        else:
            result_label.config(text=languages[current_language]["no_save_path_selected"])

    except Exception as e:
        result_label.config(text=languages[current_language]["analysis_error"].format(str(e)))


def switch_language(event):
    global current_language
    if current_language == "zh":
        current_language = "en"
    else:
        current_language = "zh"

    # 更新界面文字
    root.title(languages[current_language]["title"])
    select_button.config(text=languages[current_language]["select_button_text"])
    file_entry.delete(0, tk.END)
    file_entry.insert(0, languages[current_language]["file_entry_placeholder"])
    file_entry.config(foreground='gray')
    analyze_button.config(text=languages[current_language]["analyze_button_text"])
    switch_language_label.config(text=languages[current_language]["switch_language_button_text"])
    column_name_hint_label.config(text=languages[current_language]["column_name_hint"])


# 创建主窗口
root = ttk.Window(themename="flatly")
root.title(languages[current_language]["title"])

# 获取屏幕的宽度和高度
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# 设置窗口的宽度和高度
window_width = 500
window_height = 300

# 计算窗口应该放置的位置
x = (screen_width - window_width) // 2
y = (screen_height - window_height) // 2

# 设置窗口的位置和大小
root.geometry(f"{window_width}x{window_height}+{x}+{y}")

# 创建一个框架来包含按钮和输入框
frame = ttk.Frame(root)
frame.pack(expand=True)

# 创建文件选择按钮
select_button = ttk.Button(frame, text=languages[current_language]["select_button_text"], command=select_file,
                           bootstyle=PRIMARY)
select_button.pack(pady=10)

# 创建文件路径输入框
file_entry = ttk.Entry(frame, width=50)
file_entry.insert(0, languages[current_language]["file_entry_placeholder"])
file_entry.config(foreground='gray')
file_entry.bind('<FocusIn>', on_entry_click)
file_entry.bind('<FocusOut>', on_focusout)
file_entry.pack(pady=5)

# 创建分析按钮
analyze_button = ttk.Button(frame, text=languages[current_language]["analyze_button_text"], command=analyze_file,
                            bootstyle=SUCCESS)
analyze_button.pack(pady=10)

# 创建列名提示标签
column_name_hint_label = ttk.Label(frame, text=languages[current_language]["column_name_hint"], foreground="gray")
column_name_hint_label.pack(pady=5)

# 创建切换语言标签
switch_language_label = ttk.Label(frame, text=languages[current_language]["switch_language_button_text"],
                                  foreground="gray", cursor="hand2")
switch_language_label.bind("<Button-1>", switch_language)
switch_language_label.pack(pady=10)

# 创建结果显示标签
result_label = ttk.Label(root, text="", justify=tk.LEFT)
result_label.pack(pady=10)

# 运行主循环
root.mainloop()
