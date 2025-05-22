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
        'title': "指数平滑法分析",
        'select_button': "选择文件",
        'analyze_button': "分析文件",
        'file_not_found': "文件不存在，请重新选择。",
        'analysis_success': "分析完成，结果已保存到 {}\n",
        'no_save_path': "未选择保存路径，结果未保存。",
        'analysis_error': "分析文件时出错: {}",
        'switch_language': "切换语言",
        'explanation': {
            "原始数据": "输入的待分析数据",
            "一次指数平滑值": "通过指数平滑法计算得到的一次平滑值序列",
            "预测值": "基于一次指数平滑值得到的预测值序列",
            "预测结果折线图": "展示原始数据和预测值的折线图"
        },
        'interpretation': {
            "原始数据": "作为分析的基础数据",
            "一次指数平滑值": "反映数据的平滑趋势",
            "预测值": "反映未来趋势的预测结果",
            "预测结果折线图": "直观展示原始数据和预测值的变化趋势"
        }
    },
    'en': {
        'title': "Exponential Smoothing Method Analysis",
        'select_button': "Select File",
        'analyze_button': "Analyze File",
        'file_not_found': "The file does not exist. Please select again.",
        'analysis_success': "Analysis completed. The results have been saved to {}\n",
        'no_save_path': "No save path selected. The results were not saved.",
        'analysis_error': "An error occurred while analyzing the file: {}",
        'switch_language': "Switch Language",
        'explanation': {
            "原始数据": "The input data to be analyzed",
            "一次指数平滑值": "The first-order exponentially smoothed value sequence calculated by the exponential smoothing method",
            "预测值": "The predicted value sequence based on the first-order exponentially smoothed values",
            "预测结果折线图": "A line chart showing the original data and predicted values"
        },
        'interpretation': {
            "原始数据": "As the basic data for analysis",
            "一次指数平滑值": "Reflects the smoothing trend of the data",
            "预测值": "The predicted results reflecting future trends",
            "预测结果折线图": "Visually display the changing trends of the original data and predicted values"
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


def exponential_smoothing(x, alpha):
    """
    一次指数平滑法
    :param x: 原始数据序列
    :param alpha: 平滑系数
    :return: 一次指数平滑值序列
    """
    s = [x[0]]
    for i in range(1, len(x)):
        s.append(alpha * x[i] + (1 - alpha) * s[i - 1])
    return s


def analyze_file():
    global current_language
    file_path = file_entry.get()
    if file_path == "请输入待分析 Excel 文件的完整路径" or file_path == "Please enter the full path of the Excel file to be analyzed":
        file_path = ""
    if not os.path.exists(file_path):
        result_text.delete(1.0, tk.END)
        result_text.insert(tk.END, LANGUAGES[current_language]['file_not_found'])
        return
    try:
        # 打开 Excel 文件
        df = pd.read_excel(file_path, header=None)
        data = df.values.flatten()

        # 将数据转换为浮点类型
        data = data.astype(float)

        # 进行指数平滑分析，平滑系数取 0.3
        alpha = 0.3
        smoothed_values = exponential_smoothing(data, alpha)
        # 预测值为最后一个平滑值
        pred_values = smoothed_values + [smoothed_values[-1]]

        # 整理数据
        data = [
            ["原始数据", data.tolist(), ""],
            ["一次指数平滑值", smoothed_values, ""],
            ["预测值", pred_values, ""]
        ]
        headers = ["统计量", "统计量值", "p值"]
        df = pd.DataFrame(data, columns=headers)

        # 添加解释说明
        explanations = LANGUAGES[current_language]['explanation']
        interpretations = LANGUAGES[current_language]['interpretation']
        explanation_df = pd.DataFrame([explanations])
        explanation_df = explanation_df.reindex(
            columns=["原始数据", "一次指数平滑值", "预测值", "预测结果折线图"])
        explanation_df.insert(0, "统计量_解释说明", "解释说明" if current_language == 'zh' else "Explanation")

        # 添加分析结果解读
        interpretation_df = pd.DataFrame([interpretations])
        interpretation_df = interpretation_df.reindex(
            columns=["原始数据", "一次指数平滑值", "预测值", "预测结果折线图"])
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
            result_text.delete(1.0, tk.END)
            result_text.insert(tk.END, result_msg)

            # 生成预测结果折线图
            plt.figure()
            plt.plot(range(len(data)), data, label='原始数据')
            plt.plot(range(len(pred_values)), pred_values, label='预测值', linestyle='--')
            plt.title('预测结果折线图' if current_language == 'zh' else 'Line Chart of Prediction Results')
            plt.xlabel('时间步' if current_language == 'zh' else 'Time Step')
            plt.ylabel('值' if current_language == 'zh' else 'Value')
            plt.legend()

            # 保存图片
            img_path = os.path.splitext(save_path)[0] + '_prediction_chart.png'
            plt.savefig(img_path)
            plt.close()

        else:
            result_text.delete(1.0, tk.END)
            result_text.insert(tk.END, LANGUAGES[current_language]['no_save_path'])

    except Exception as e:
        result_text.delete(1.0, tk.END)
        result_text.insert(tk.END, LANGUAGES[current_language]['analysis_error'].format(str(e)))


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

# 创建结果显示文本框
result_text = tk.Text(root, height=4, width=60, wrap=tk.WORD)
result_text.pack(pady=10)

# 运行主循环
root.mainloop()
