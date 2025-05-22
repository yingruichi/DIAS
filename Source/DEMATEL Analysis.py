import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox
import openpyxl
import os
import numpy as np
import pandas as pd
from tkinter import filedialog
import tkinter as tk
import matplotlib.pyplot as plt
import pathlib

# 设置 matplotlib 支持中文
plt.rcParams['font.family'] = 'SimHei'
plt.rcParams['axes.unicode_minus'] = False

# 定义语言字典
LANGUAGES = {
    'zh': {
        'title': "DEMATEL 分析",
        'select_button': "选择文件",
        'analyze_button': "分析文件",
        'file_not_found': "文件不存在，请重新选择。",
        'analysis_success': "分析完成，结果已保存到 {}\n",
        'no_save_path': "未选择保存路径，结果未保存。",
        'analysis_error': "分析文件时出错: {}",
        'switch_language': "切换语言",
        'explanation': {
            "综合影响矩阵": "反映因素之间综合影响关系的矩阵",
            "原因度": "衡量因素对其他因素影响程度的指标",
            "中心度": "衡量因素在系统中重要程度的指标"
        },
        'interpretation': {
            "综合影响矩阵": "矩阵元素值越大，对应因素之间的影响越强",
            "原因度": "原因度为正，该因素为原因因素；原因度为负，该因素为结果因素",
            "中心度": "中心度越大，该因素在系统中越重要"
        }
    },
    'en': {
        'title': "DEMATEL Analysis",
        'select_button': "Select File",
        'analyze_button': "Analyze File",
        'file_not_found': "The file does not exist. Please select again.",
        'analysis_success': "Analysis completed. The results have been saved to {}\n",
        'no_save_path': "No save path selected. The results were not saved.",
        'analysis_error': "An error occurred while analyzing the file: {}",
        'switch_language': "Switch Language",
        'explanation': {
            "综合影响矩阵": "A matrix reflecting the comprehensive influence relationship between factors",
            "原因度": "An indicator to measure the influence degree of a factor on other factors",
            "中心度": "An indicator to measure the importance of a factor in the system"
        },
        'interpretation': {
            "综合影响矩阵": "The larger the matrix element value, the stronger the influence between corresponding factors",
            "原因度": "If the causal degree is positive, the factor is a causal factor; if negative, it is a result factor",
            "中心度": "The larger the centrality, the more important the factor in the system"
        }
    }
}

# 当前语言
current_language = 'en'


def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, file_path)
        file_entry.configure(style="TEntry")


def dematel_analysis(data):
    """
    进行 DEMATEL 分析
    :param data: 直接影响矩阵数据
    :return: 综合影响矩阵、原因度、中心度
    """
    # 归一化直接影响矩阵
    n = data.shape[0]
    max_sum_row = np.max(np.sum(data, axis=1))
    max_sum_col = np.max(np.sum(data, axis=0))
    max_value = max(max_sum_row, max_sum_col)
    D = data / max_value

    # 计算综合影响矩阵
    I = np.eye(n)
    T = np.dot(D, np.linalg.inv(I - D))

    # 计算原因度和中心度
    sum_row = np.sum(T, axis=1)
    sum_col = np.sum(T, axis=0)
    causal_degree = sum_row - sum_col
    centrality = sum_row + sum_col

    return T, causal_degree, centrality


def analyze_file():
    global current_language
    file_path = file_entry.get()
    if file_path == "请输入待分析 Excel 文件的完整路径" or file_path == "Please enter the full path of the Excel file to be analyzed":
        file_path = ""
    if not os.path.exists(file_path):
        result_label.config(text=LANGUAGES[current_language]['file_not_found'])
        return
    try:
        # 打开 Excel 文件
        df = pd.read_excel(file_path, header=None)
        data = df.values

        # 进行 DEMATEL 分析
        T, causal_degree, centrality = dematel_analysis(data)

        # 整理数据
        factors = [f"因素{i + 1}" for i in range(data.shape[0])]
        T_df = pd.DataFrame(T, index=factors, columns=factors)
        causal_degree_df = pd.DataFrame(causal_degree, index=factors, columns=["原因度"])
        centrality_df = pd.DataFrame(centrality, index=factors, columns=["中心度"])

        # 添加解释说明
        explanations = LANGUAGES[current_language]['explanation']
        interpretations = LANGUAGES[current_language]['interpretation']
        explanation_df = pd.DataFrame([explanations])
        explanation_df = explanation_df.reindex(columns=["综合影响矩阵", "原因度", "中心度"])
        explanation_df.insert(0, "统计量_解释说明", "解释说明" if current_language == 'zh' else "Explanation")

        # 添加分析结果解读
        interpretation_df = pd.DataFrame([interpretations])
        interpretation_df = interpretation_df.reindex(columns=["综合影响矩阵", "原因度", "中心度"])
        interpretation_df.insert(0, "统计量_结果解读", "结果解读" if current_language == 'zh' else "Interpretation")

        # 让用户选择保存路径
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            # 保存到 Excel 文件
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                T_df.to_excel(writer, sheet_name="综合影响矩阵", index=True)
                causal_degree_df.to_excel(writer, sheet_name="原因度", index=True)
                centrality_df.to_excel(writer, sheet_name="中心度", index=True)
                explanation_df.to_excel(writer, sheet_name="解释说明", index=False)
                interpretation_df.to_excel(writer, sheet_name="结果解读", index=False)

                # 自动调整列宽
                for sheet_name in writer.sheets:
                    worksheet = writer.sheets[sheet_name]
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = openpyxl.utils.get_column_letter(column[0].column)
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = (max_length + 2)
                        worksheet.column_dimensions[column_letter].width = adjusted_width

            result_msg = LANGUAGES[current_language]['analysis_success'].format(save_path)
            result_label.config(text=result_msg, wraplength=400)

            # 生成原因度和中心度柱状图
            fig, axes = plt.subplots(2, 1, figsize=(8, 10))
            axes[0].bar(factors, causal_degree)
            axes[0].set_title('原因度柱状图' if current_language == 'zh' else 'Bar Chart of Causal Degree')
            axes[0].set_xlabel('因素' if current_language == 'zh' else 'Factors')
            axes[0].set_ylabel('原因度' if current_language == 'zh' else 'Causal Degree')

            axes[1].bar(factors, centrality)
            axes[1].set_title('中心度柱状图' if current_language == 'zh' else 'Bar Chart of Centrality')
            axes[1].set_xlabel('因素' if current_language == 'zh' else 'Factors')
            axes[1].set_ylabel('中心度' if current_language == 'zh' else 'Centrality')

            # 保存图片
            img_path = os.path.splitext(save_path)[0] + '_charts.png'
            plt.tight_layout()
            plt.savefig(img_path)
            plt.close()

        else:
            result_label.config(text=LANGUAGES[current_language]['no_save_path'])

    except Exception as e:
        result_label.config(text=LANGUAGES[current_language]['analysis_error'].format(str(e)))


def switch_language():
    global current_language
    current_language = 'en' if current_language == 'zh' else 'zh'
    root.title(LANGUAGES[current_language]['title'])
    select_button.config(text=LANGUAGES[current_language]['select_button'])
    analyze_button.config(text=LANGUAGES[current_language]['analyze_button'])
    language_label.config(text=LANGUAGES[current_language]['switch_language'])
    # 切换语言时更新提示信息
    file_entry.delete(0, tk.END)
    if current_language == 'zh':
        file_entry.insert(0, "请输入待分析 Excel 文件的完整路径")
        file_entry.configure(style="Gray.TEntry")
    else:
        file_entry.insert(0, "Please enter the full path of the Excel file to be analyzed")
        file_entry.configure(style="Gray.TEntry")


def on_entry_click(event):
    """当用户点击输入框时，清除提示信息"""
    if file_entry.get() == "请输入待分析 Excel 文件的完整路径" or file_entry.get() == "Please enter the full path of the Excel file to be analyzed":
        file_entry.delete(0, tk.END)
        file_entry.configure(style="TEntry")


def on_focusout(event):
    """当用户离开输入框时，如果没有输入内容，恢复提示信息"""
    if file_entry.get() == "":
        if current_language == 'zh':
            file_entry.insert(0, "请输入待分析 Excel 文件的完整路径")
            file_entry.configure(style="Gray.TEntry")
        else:
            file_entry.insert(0, "Please enter the full path of the Excel file to be analyzed")
            file_entry.configure(style="Gray.TEntry")


# 创建主窗口
root = ttk.Window(themename="flatly")
root.title(LANGUAGES[current_language]['title'])

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

# 创建自定义样式
style = ttk.Style()
style.configure("Gray.TEntry", foreground="gray")

# 创建文件选择按钮
select_button = ttk.Button(root, text=LANGUAGES[current_language]['select_button'], command=select_file,
                           bootstyle=PRIMARY)
select_button.pack(pady=10)

# 创建文件路径输入框
file_entry = ttk.Entry(root, width=50, style="Gray.TEntry")
if current_language == 'zh':
    file_entry.insert(0, "请输入待分析 Excel 文件的完整路径")
else:
    file_entry.insert(0, "Please enter the full path of the Excel file to be analyzed")
file_entry.pack(pady=5)
file_entry.bind("<FocusIn>", on_entry_click)
file_entry.bind("<FocusOut>", on_focusout)

# 创建分析按钮
analyze_button = ttk.Button(root, text=LANGUAGES[current_language]['analyze_button'], command=analyze_file,
                            bootstyle=SUCCESS)
analyze_button.pack(pady=10)

# 创建语言切换标签
language_label = ttk.Label(root, text=LANGUAGES[current_language]['switch_language'], cursor="hand2")
language_label.pack(pady=10)
language_label.bind("<Button-1>", lambda event: switch_language())

# 创建结果显示标签
result_label = ttk.Label(root, text="", justify=tk.LEFT)
result_label.pack(pady=10)

# 运行主循环
root.mainloop()