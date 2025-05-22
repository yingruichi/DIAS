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
        'title': "秩和比(RSR)分析",
        'select_button': "选择文件",
        'analyze_button': "分析文件",
        'file_not_found': "文件不存在，请重新选择。",
        'analysis_success': "分析完成，结果已保存到 {}\n",
        'no_save_path': "未选择保存路径，结果未保存。",
        'analysis_error': "分析文件时出错: {}",
        'switch_language': "切换语言",
        'explanation': {
            "秩矩阵": "将原始数据转换为秩次后得到的矩阵",
            "秩和比(RSR)": "反映各评价对象综合水平的统计量",
            "RSR 分布直方图": "展示 RSR 值分布情况的直方图",
            "回归方程": "用于拟合 RSR 值与概率单位之间关系的方程",
            "RSR 排序结果": "根据 RSR 值对各评价对象进行排序的结果"
        },
        'interpretation': {
            "秩矩阵": "便于后续计算秩和比",
            "秩和比(RSR)": "值越大，综合水平越高",
            "RSR 分布直方图": "直观观察 RSR 值的分布特征",
            "回归方程": "用于进一步分析 RSR 值的变化规律",
            "RSR 排序结果": "排名越靠前，综合水平越高"
        }
    },
    'en': {
        'title': "Rank - Sum Ratio (RSR) Analysis",
        'select_button': "Select File",
        'analyze_button': "Analyze File",
        'file_not_found': "The file does not exist. Please select again.",
        'analysis_success': "Analysis completed. The results have been saved to {}\n",
        'no_save_path': "No save path selected. The results were not saved.",
        'analysis_error': "An error occurred while analyzing the file: {}",
        'switch_language': "Switch Language",
        'explanation': {
            "秩矩阵": "The matrix obtained after converting the original data into ranks",
            "秩和比(RSR)": "A statistic reflecting the comprehensive level of each evaluation object",
            "RSR 分布直方图": "A histogram showing the distribution of RSR values",
            "回归方程": "An equation used to fit the relationship between RSR values and probability units",
            "RSR 排序结果": "The result of ranking each evaluation object according to RSR values"
        },
        'interpretation': {
            "秩矩阵": "Facilitate the subsequent calculation of the rank - sum ratio",
            "秩和比(RSR)": "The larger the value, the higher the comprehensive level",
            "RSR 分布直方图": "Visually observe the distribution characteristics of RSR values",
            "回归方程": "Used to further analyze the change rule of RSR values",
            "RSR 排序结果": "The higher the ranking, the higher the comprehensive level"
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


def rsr_method(data):
    """
    实现秩和比(RSR)法
    :param data: 原始数据矩阵
    :return: 秩矩阵, 秩和比(RSR), RSR 排序结果
    """
    # 计算秩矩阵
    rank_matrix = np.apply_along_axis(lambda x: pd.Series(x).rank().values, 0, data)

    # 计算秩和比(RSR)
    RSR = rank_matrix.sum(axis=1) / (rank_matrix.shape[0] * rank_matrix.shape[1])

    # 对 RSR 进行排序
    ranking = np.argsort(-RSR) + 1

    return rank_matrix, RSR, ranking


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

        # 将数据转换为浮点类型
        data = data.astype(float)

        # 进行 RSR 分析
        rank_matrix, RSR, ranking = rsr_method(data)

        # 整理数据
        data = [
            ["秩矩阵", rank_matrix.tolist(), ""],
            ["秩和比(RSR)", RSR.tolist(), ""],
            ["RSR 排序结果", ranking.tolist(), ""]
        ]
        headers = ["统计量", "统计量值", "p值"]
        df = pd.DataFrame(data, columns=headers)

        # 添加解释说明
        explanations = LANGUAGES[current_language]['explanation']
        interpretations = LANGUAGES[current_language]['interpretation']
        explanation_df = pd.DataFrame([explanations])
        explanation_df = explanation_df.reindex(
            columns=["秩矩阵", "秩和比(RSR)", "RSR 分布直方图", "回归方程", "RSR 排序结果"])
        explanation_df.insert(0, "统计量_解释说明", "解释说明" if current_language == 'zh' else "Explanation")

        # 添加分析结果解读
        interpretation_df = pd.DataFrame([interpretations])
        interpretation_df = interpretation_df.reindex(
            columns=["秩矩阵", "秩和比(RSR)", "RSR 分布直方图", "回归方程", "RSR 排序结果"])
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

            # 生成 RSR 分布直方图
            fig, ax = plt.subplots()
            ax.hist(RSR, bins=10)
            ax.set_title(
                'RSR 分布直方图' if current_language == 'zh' else 'Histogram of RSR Distribution')
            ax.set_xlabel('秩和比(RSR)' if current_language == 'zh' else 'Rank - Sum Ratio (RSR)')
            ax.set_ylabel('频数' if current_language == 'zh' else 'Frequency')
            # 保存图片
            img_path = os.path.splitext(save_path)[0] + '_rsr_histogram.png'
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