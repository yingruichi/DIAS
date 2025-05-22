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
        'title': "模糊层次分析法 FAHP 分析",
        'select_button': "选择文件",
        'analyze_button': "分析文件",
        'file_not_found': "文件不存在，请重新选择。",
        'analysis_success': "分析完成，结果已保存到 {}\n",
        'no_save_path': "未选择保存路径，结果未保存。",
        'analysis_error': "分析文件时出错: {}",
        'switch_language': "切换语言",
        'explanation': {
            "模糊特征向量": "反映各因素相对重要性的模糊向量",
            "一致性指标 CI": "衡量模糊判断矩阵一致性的指标",
            "随机一致性指标 RI": "根据矩阵阶数确定的随机一致性指标",
            "一致性比率 CR": "CI 与 RI 的比值，判断矩阵是否具有满意一致性"
        },
        'interpretation': {
            "模糊特征向量": "模糊特征向量值越大，对应因素越重要",
            "一致性指标 CI": "CI 值越小，矩阵一致性越好",
            "随机一致性指标 RI": "不同阶数矩阵有对应标准值",
            "一致性比率 CR": "CR < 0.1 时，矩阵具有满意一致性，结果可信"
        }
    },
    'en': {
        'title': "Fuzzy Analytic Hierarchy Process (FAHP) Analysis",
        'select_button': "Select File",
        'analyze_button': "Analyze File",
        'file_not_found': "The file does not exist. Please select again.",
        'analysis_success': "Analysis completed. The results have been saved to {}\n",
        'no_save_path': "No save path selected. The results were not saved.",
        'analysis_error': "An error occurred while analyzing the file: {}",
        'switch_language': "Switch Language",
        'explanation': {
            "模糊特征向量": "A fuzzy vector reflecting the relative importance of each factor",
            "一致性指标 CI": "An indicator to measure the consistency of the fuzzy judgment matrix",
            "随机一致性指标 RI": "A random consistency indicator determined by the order of the matrix",
            "一致性比率 CR": "The ratio of CI to RI to determine if the matrix has satisfactory consistency"
        },
        'interpretation': {
            "模糊特征向量": "The larger the value in the fuzzy eigenvector, the more important the corresponding factor",
            "一致性指标 CI": "The smaller the CI value, the better the consistency of the matrix",
            "随机一致性指标 RI": "There are corresponding standard values for matrices of different orders",
            "一致性比率 CR": "When CR < 0.1, the matrix has satisfactory consistency and the results are reliable"
        }
    }
}

# 当前语言
current_language = 'en'

# 随机一致性指标 RI 表
RI_TABLE = {
    1: 0, 2: 0, 3: 0.58, 4: 0.90, 5: 1.12, 6: 1.24, 7: 1.32, 8: 1.41, 9: 1.45
}


def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, file_path)
        file_entry.configure(style="TEntry")


def fahp_analysis(data):
    """
    进行模糊层次分析法 FAHP 分析
    :param data: 模糊判断矩阵数据
    :return: 模糊特征向量、一致性指标 CI、一致性比率 CR
    """
    # 计算模糊特征向量
    row_sums = np.sum(data, axis=1)
    fuzzy_eigenvector = row_sums / np.sum(row_sums)

    # 计算模糊判断矩阵的最大特征值
    weighted_sum = np.dot(data, fuzzy_eigenvector)
    max_eigenvalue = np.sum(weighted_sum / (fuzzy_eigenvector * len(data)))

    # 计算一致性指标 CI
    n = data.shape[0]
    CI = (max_eigenvalue - n) / (n - 1)

    # 计算随机一致性指标 RI
    RI = RI_TABLE.get(n, None)
    if RI is None:
        raise ValueError("判断矩阵阶数超出支持范围")

    # 计算一致性比率 CR
    CR = CI / RI

    return fuzzy_eigenvector, CI, RI, CR


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

        # 进行 FAHP 分析
        fuzzy_eigenvector, CI, RI, CR = fahp_analysis(data)

        # 整理数据
        data = [
            ["模糊特征向量", fuzzy_eigenvector.tolist(), ""],
            ["一致性指标 CI", CI, ""],
            ["随机一致性指标 RI", RI, ""],
            ["一致性比率 CR", CR, ""]
        ]
        headers = ["统计量", "统计量值", "p值"]
        df = pd.DataFrame(data, columns=headers)

        # 添加解释说明
        explanations = LANGUAGES[current_language]['explanation']
        interpretations = LANGUAGES[current_language]['interpretation']
        explanation_df = pd.DataFrame([explanations])
        explanation_df = explanation_df.reindex(columns=["模糊特征向量", "一致性指标 CI", "随机一致性指标 RI", "一致性比率 CR"])
        explanation_df.insert(0, "统计量_解释说明", "解释说明" if current_language == 'zh' else "Explanation")

        # 添加分析结果解读
        interpretation_df = pd.DataFrame([interpretations])
        interpretation_df = interpretation_df.reindex(columns=["模糊特征向量", "一致性指标 CI", "随机一致性指标 RI", "一致性比率 CR"])
        interpretation_df.insert(0, "统计量_结果解读", "结果解读" if current_language == 'zh' else "Interpretation")

        # 合并数据、解释说明和结果解读
        combined_df = pd.concat([df, explanation_df, interpretation_df], ignore_index=True)

        # 让用户选择保存路径
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            # 保存到 Excel 文件
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                combined_df.to_excel(writer, index=False)
                worksheet = writer.sheets['Sheet1']
                # 自动调整列宽
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

            # 生成模糊特征向量柱状图
            fig, ax = plt.subplots()
            ax.bar(range(len(fuzzy_eigenvector)), fuzzy_eigenvector)
            ax.set_title('模糊特征向量柱状图' if current_language == 'zh' else 'Bar Chart of Fuzzy Eigenvector')
            ax.set_xlabel('因素' if current_language == 'zh' else 'Factors')
            ax.set_ylabel('权重' if current_language == 'zh' else 'Weights')
            # 保存图片
            img_path = os.path.splitext(save_path)[0] + '_fuzzy_eigenvector.png'
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