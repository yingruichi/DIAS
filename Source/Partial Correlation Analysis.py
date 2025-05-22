import tkinter as tk
from tkinter import filedialog
import openpyxl
import os
import pandas as pd
import numpy as np
from scipy import stats
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import matplotlib.pyplot as plt
import matplotlib

# 设置支持中文的字体
matplotlib.rcParams['font.family'] = 'SimHei'
# 解决负号显示问题
matplotlib.rcParams['axes.unicode_minus'] = False

# 定义语言字典
languages = {
    "zh": {
        "title": "偏相关分析",
        "select_button_text": "选择文件",
        "file_entry_placeholder": "请输入待分析 Excel 文件的完整路径",
        "analyze_button_text": "分析文件",
        "no_file_selected": "请选择有效的文件路径。",
        "file_not_exists": "文件不存在，请重新选择。",
        "analysis_error": "分析文件时出错: {}",
        "analysis_complete": "分析完成，结果已保存到 {}，相关图片已保存。",
        "no_save_path_selected": "未选择保存路径，结果未保存。",
        "columns_stats": ["变量对", "偏相关系数", "p 值", "结果解读"],
        "interpretation_low_p": "p 值小于 0.05，表明该变量对之间的偏相关性显著。",
        "interpretation_high_p": "p 值大于等于 0.05，表明该变量对之间的偏相关性不显著。",
        "switch_language_button_text": "切换语言"
    },
    "en": {
        "title": "Partial Correlation Analysis",
        "select_button_text": "Select File",
        "file_entry_placeholder": "Please enter the full path of the Excel file to be analyzed",
        "analyze_button_text": "Analyze File",
        "no_file_selected": "Please select a valid file path.",
        "file_not_exists": "The file does not exist. Please select again.",
        "analysis_error": "An error occurred while analyzing the file: {}",
        "analysis_complete": "Analysis completed. The results have been saved to {}, and the relevant images have been saved.",
        "no_save_path_selected": "No save path selected. The results were not saved.",
        "columns_stats": ["Variable Pair", "Partial Correlation Coefficient", "p-value", "Result Interpretation"],
        "interpretation_low_p": "The p-value is less than 0.05, indicating that the partial correlation between this variable pair is significant.",
        "interpretation_high_p": "The p-value is greater than or equal to 0.05, indicating that the partial correlation between this variable pair is not significant.",
        "switch_language_button_text": "Switch Language"
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


# 计算偏相关系数的函数
def partial_corr(data, x, y, z):
    sub_data = data[[x, y] + z]
    sub_data = sub_data.dropna()
    X = sub_data[[x] + z]
    Y = sub_data[[y] + z]
    beta_x = np.linalg.lstsq(X, sub_data[x], rcond=None)[0]
    beta_y = np.linalg.lstsq(Y, sub_data[y], rcond=None)[0]
    resid_x = sub_data[x] - np.dot(X, beta_x)
    resid_y = sub_data[y] - np.dot(Y, beta_y)
    corr, p = stats.pearsonr(resid_x, resid_y)
    return corr, p


def analyze_file():
    file_path = file_entry.get()
    if file_path == languages[current_language]["file_entry_placeholder"]:
        result_label.config(text=languages[current_language]["no_file_selected"])
        return
    if not os.path.exists(file_path):
        result_label.config(text=languages[current_language]["file_not_exists"])
        return
    try:
        # 打开 Excel 文件
        df = pd.read_excel(file_path)

        # 获取所有变量名
        variables = df.columns.tolist()

        # 存储结果的列表
        results = []

        # 进行偏相关分析
        for i in range(len(variables)):
            for j in range(i + 1, len(variables)):
                x = variables[i]
                y = variables[j]
                other_vars = [var for var in variables if var not in [x, y]]
                corr, p = partial_corr(df, x, y, other_vars)
                pair = f"{x} - {y}"
                if p < 0.05:
                    interpretation = languages[current_language]["interpretation_low_p"]
                else:
                    interpretation = languages[current_language]["interpretation_high_p"]
                results.append([pair, corr, p, interpretation])

        # 创建结果 DataFrame
        result_df = pd.DataFrame(results, columns=languages[current_language]["columns_stats"])

        # 绘制偏相关系数的柱状图
        plt.figure()
        plt.bar(result_df["变量对" if current_language == "zh" else "Variable Pair"],
                result_df["偏相关系数" if current_language == "zh" else "Partial Correlation Coefficient"])
        plt.xlabel('变量对' if current_language == "zh" else 'Variable Pair')
        plt.ylabel('偏相关系数' if current_language == "zh" else 'Partial Correlation Coefficient')
        plt.title('偏相关系数分析结果' if current_language == "zh" else 'Partial Correlation Analysis Results')
        plt.xticks(rotation=45)

        # 保存图片
        image_path = os.path.splitext(file_path)[0] + '_partial_corr_plot.png'
        plt.savefig(image_path)
        plt.close()

        # 让用户选择保存路径
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            # 保存 DataFrame 到 Excel 文件
            result_df.to_excel(save_path, index=False)

            # 打开保存的 Excel 文件并调整列宽
            wb = openpyxl.load_workbook(save_path)
            ws = wb.active

            for column in ws.columns:
                max_length = 0
                column_letter = openpyxl.utils.get_column_letter(column[0].column)
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2)
                ws.column_dimensions[column_letter].width = adjusted_width

            # 保存调整列宽后的 Excel 文件
            wb.save(save_path)

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