import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox
import openpyxl
import os
import numpy as np
import pandas as pd
from tkinter import filedialog
import tkinter as tk
from scipy import stats
import matplotlib.pyplot as plt
import pathlib
import matplotlib

# 设置 matplotlib 支持中文
matplotlib.rcParams['font.family'] = 'SimHei'  # 使用黑体字体，你可以根据需要替换为其他支持中文的字体
matplotlib.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题

# 定义语言字典
LANGUAGES = {
    'zh': {
        'title': "单样本方差分析",
        'select_button': "选择文件",
        'analyze_button': "分析文件",
        'file_not_found': "文件不存在，请重新选择。",
        'analysis_success': "分析完成，结果已保存到 {}\n",
        'no_save_path': "未选择保存路径，结果未保存。",
        'analysis_error': "分析文件时出错: {}",
        'switch_language': "切换语言",
        'explanation': {
            "单样本方差分析": "用于检验一个样本的均值是否与某个已知的总体均值存在显著差异。",
            "样本量": "样本中的观测值数量。",
            "均值": "样本数据的平均值。",
            "t统计量": "衡量样本均值与总体均值之间差异的统计量。",
            "自由度": "用于计算t分布的参数。",
            "p值": "p值小于显著性水平（通常为0.05）时，拒绝原假设，认为样本均值与总体均值存在显著差异；否则，接受原假设，认为样本均值与总体均值无显著差异。",
            "效应量": "反映样本均值与总体均值之间差异的程度。"
        },
        'interpretation': {
            "t统计量": "t统计量的绝对值越大，说明样本均值与总体均值之间的差异越显著。",
            "p值": "用于判断样本均值与总体均值之间是否存在显著差异。",
            "自由度": "影响t分布的形状，进而影响p值的计算。",
            "样本量": "样本量的大小会影响统计检验的功效，较大的样本量通常能提供更准确的结果。",
            "均值": "反映样本数据的平均水平。",
            "效应量": "效应量越大，说明样本均值与总体均值之间的差异越大。"
        }
    },
    'en': {
        'title': "One-sample ANOVA",
        'select_button': "Select File",
        'analyze_button': "Analyze File",
        'file_not_found': "The file does not exist. Please select again.",
        'analysis_success': "Analysis completed. The results have been saved to {}\n",
        'no_save_path': "No save path selected. The results were not saved.",
        'analysis_error': "An error occurred while analyzing the file: {}",
        'switch_language': "Switch Language",
        'explanation': {
            "单样本方差分析": "Used to test whether the mean of a sample is significantly different from a known population mean.",
            "样本量": "The number of observations in the sample.",
            "均值": "The average value of the sample data.",
            "t统计量": "A statistic that measures the difference between the sample mean and the population mean.",
            "自由度": "Parameters used to calculate the t-distribution.",
            "p值": "When the p-value is less than the significance level (usually 0.05), the null hypothesis is rejected, indicating a significant difference between the sample mean and the population mean; otherwise, the null hypothesis is accepted, indicating no significant difference.",
            "效应量": "Reflects the degree of difference between the sample mean and the population mean."
        },
        'interpretation': {
            "t统计量": "The larger the absolute value of the t-statistic, the more significant the difference between the sample mean and the population mean.",
            "p值": "Used to determine whether there is a significant difference between the sample mean and the population mean.",
            "自由度": "Affects the shape of the t-distribution, which in turn affects the calculation of the p-value.",
            "样本量": "The sample size affects the power of the statistical test. A larger sample size usually provides more accurate results.",
            "均值": "Reflects the average level of the sample data.",
            "效应量": "The larger the effect size, the greater the difference between the sample mean and the population mean."
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
        file_entry.configure(style="TEntry")  # 恢复默认样式


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
        df = pd.read_excel(file_path)

        # 检查数据是否为数值类型
        numerical_df = df.select_dtypes(include=[np.number])
        if numerical_df.empty:
            raise ValueError("数据中没有数值列，无法进行方差分析。")

        # 假设总体均值为 0
        population_mean = 0

        # 取第一列数据进行单样本 t 检验
        sample = numerical_df.iloc[:, 0]
        t_stat, p_value = stats.ttest_1samp(sample, population_mean)

        # 计算自由度
        df = len(sample) - 1

        # 计算效应量（Cohen's d）
        cohen_d = (sample.mean() - population_mean) / sample.std()

        # 计算样本量和均值
        sample_size = len(sample)
        mean = sample.mean()

        # 整理数据
        data = [
            ["方差分析", t_stat, df, p_value, cohen_d],
            ["样本量", sample_size, "", "", ""],
            ["均值", mean, "", "", ""]
        ]
        headers = ["统计量", "t统计量", "自由度", "p值", "效应量（Cohen's d）"]
        df = pd.DataFrame(data, columns=headers)

        # 添加解释说明
        explanations = LANGUAGES[current_language]['explanation']
        interpretations = LANGUAGES[current_language]['interpretation']
        explanation_df = pd.DataFrame([explanations])
        explanation_df = explanation_df.reindex(
            columns=["单样本方差分析", "样本量", "均值", "t统计量", "自由度", "p值", "效应量"])
        explanation_df.insert(0, "统计量_解释说明", "解释说明" if current_language == 'zh' else "Explanation")

        # 添加分析结果解读
        interpretation_df = pd.DataFrame([interpretations])
        interpretation_df = interpretation_df.reindex(columns=["t统计量", "p值", "自由度", "样本量", "均值", "效应量"])
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

            result_msg = LANGUAGES[current_language]['analysis_success'].format(
                save_path)
            result_label.config(text=result_msg, wraplength=400)

            # 绘制箱线图
            plt.figure(figsize=(10, 6))
            numerical_df.iloc[:, 0].plot.box()
            plt.title('箱线图' if current_language == 'zh' else 'Box Plot')
            plt.ylabel('数值' if current_language == 'zh' else 'Values')
            plt.savefig(save_path.replace('.xlsx', '_boxplot.png'))
            plt.close()

            # 绘制柱状图
            plt.figure(figsize=(10, 6))
            bars = plt.bar(['样本'], [mean])
            for bar in bars:
                height = bar.get_height()
                plt.annotate(f'{height:.2f}',
                             xy=(bar.get_x() + bar.get_width() / 2, height),
                             xytext=(0, 3),  # 3 points vertical offset
                             textcoords="offset points",
                             ha='center', va='bottom')
            plt.title('柱状图' if current_language == 'zh' else 'Bar Chart')
            plt.ylabel('均值' if current_language == 'zh' else 'Mean')
            plt.savefig(save_path.replace('.xlsx', '_barplot.png'))
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
        file_entry.configure(style="TEntry")  # 恢复默认样式


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