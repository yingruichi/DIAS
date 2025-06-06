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
        'title': "功效系数分析",
        'select_button': "选择文件",
        'analyze_button': "分析文件",
        'file_not_found': "文件不存在，请重新选择。",
        'analysis_success': "分析完成，结果已保存到 {}\n",
        'no_save_path': "未选择保存路径，结果未保存。",
        'analysis_error': "分析文件时出错: {}",
        'switch_language': "切换语言",
        'explanation': {
            "各指标实际值": "每个评价指标的实际测量值",
            "各指标不允许值": "每个评价指标的最低可接受值",
            "各指标满意值": "每个评价指标的理想值",
            "功效系数向量": "根据各指标实际值、不允许值和满意值计算得到的功效系数",
            "综合功效系数": "所有指标功效系数的加权平均值"
        },
        'interpretation': {
            "各指标实际值": "反映各指标的实际表现",
            "各指标不允许值": "作为指标表现的下限参考",
            "各指标满意值": "作为指标表现的上限参考",
            "功效系数向量": "值越高，说明该指标表现越好",
            "综合功效系数": "综合反映所有指标的整体表现，值越高越好"
        }
    },
    'en': {
        'title': "Efficacy Coefficient Analysis",
        'select_button': "Select File",
        'analyze_button': "Analyze File",
        'file_not_found': "The file does not exist. Please select again.",
        'analysis_success': "Analysis completed. The results have been saved to {}\n",
        'no_save_path': "No save path selected. The results were not saved.",
        'analysis_error': "An error occurred while analyzing the file: {}",
        'switch_language': "Switch Language",
        'explanation': {
            "各指标实际值": "The actual measured values of each evaluation indicator",
            "各指标不允许值": "The minimum acceptable values of each evaluation indicator",
            "各指标满意值": "The ideal values of each evaluation indicator",
            "功效系数向量": "The efficacy coefficients calculated based on the actual values, unacceptable values, and satisfactory values of each indicator",
            "综合功效系数": "The weighted average of the efficacy coefficients of all indicators"
        },
        'interpretation': {
            "各指标实际值": "Reflects the actual performance of each indicator",
            "各指标不允许值": "Serves as the lower limit reference for indicator performance",
            "各指标满意值": "Serves as the upper limit reference for indicator performance",
            "功效系数向量": "The higher the value, the better the performance of the indicator",
            "综合功效系数": "Comprehensively reflects the overall performance of all indicators. The higher the value, the better"
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


def efficacy_coefficient_analysis(actual_values, unacceptable_values, satisfactory_values, weights):
    """
    进行功效系数分析
    :param actual_values: 各指标实际值
    :param unacceptable_values: 各指标不允许值
    :param satisfactory_values: 各指标满意值
    :param weights: 各指标权重
    :return: 功效系数向量和综合功效系数
    """
    efficacy_coefficients = (actual_values - unacceptable_values) / (satisfactory_values - unacceptable_values) * 40 + 60
    comprehensive_efficacy_coefficient = np.dot(efficacy_coefficients, weights)
    return efficacy_coefficients, comprehensive_efficacy_coefficient


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

        # 假设第一行为各指标实际值，第二行为各指标不允许值，第三行为各指标满意值，第四行为各指标权重
        actual_values = data[0]
        unacceptable_values = data[1]
        satisfactory_values = data[2]
        weights = data[3]

        # 进行功效系数分析
        efficacy_coefficients, comprehensive_efficacy_coefficient = efficacy_coefficient_analysis(actual_values,
                                                                                                  unacceptable_values,
                                                                                                  satisfactory_values,
                                                                                                  weights)

        # 整理数据
        data = [
            ["各指标实际值", actual_values.tolist(), ""],
            ["各指标不允许值", unacceptable_values.tolist(), ""],
            ["各指标满意值", satisfactory_values.tolist(), ""],
            ["功效系数向量", efficacy_coefficients.tolist(), ""],
            ["综合功效系数", [comprehensive_efficacy_coefficient], ""]
        ]
        headers = ["统计量", "统计量值", "p值"]
        df = pd.DataFrame(data, columns=headers)

        # 添加解释说明
        explanations = LANGUAGES[current_language]['explanation']
        interpretations = LANGUAGES[current_language]['interpretation']
        explanation_df = pd.DataFrame([explanations])
        explanation_df = explanation_df.reindex(
            columns=["各指标实际值", "各指标不允许值", "各指标满意值", "功效系数向量", "综合功效系数"])
        explanation_df.insert(0, "统计量_解释说明", "解释说明" if current_language == 'zh' else "Explanation")

        # 添加分析结果解读
        interpretation_df = pd.DataFrame([interpretations])
        interpretation_df = interpretation_df.reindex(
            columns=["各指标实际值", "各指标不允许值", "各指标满意值", "功效系数向量", "综合功效系数"])
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

            # 生成功效系数向量柱状图
            fig, ax = plt.subplots()
            ax.bar(range(len(efficacy_coefficients)), efficacy_coefficients)
            ax.set_title(
                '功效系数向量柱状图' if current_language == 'zh' else 'Bar Chart of Efficacy Coefficient Vector')
            ax.set_xlabel('指标' if current_language == 'zh' else 'Indicators')
            ax.set_ylabel('功效系数' if current_language == 'zh' else 'Efficacy Coefficient')
            # 保存图片
            img_path = os.path.splitext(save_path)[0] + '_efficacy_coefficient.png'
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