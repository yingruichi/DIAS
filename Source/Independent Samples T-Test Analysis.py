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

# 定义语言字典
LANGUAGES = {
    'zh': {
        'title': "独立样本 t 检验分析",
        'select_button': "选择文件",
        'analyze_button': "分析文件",
        'file_not_found': "文件不存在，请重新选择。",
        'analysis_success': "分析完成，结果已保存到 {}\n",
        'no_save_path': "未选择保存路径，结果未保存。",
        'analysis_error': "分析文件时出错: {}",
        'switch_language': "切换语言",
        'explanation': {
            "独立样本 t 检验": "用于比较两个独立样本的均值是否有显著差异。",
            "样本量": "每个样本中的观测值数量。",
            "均值": "样本数据的平均值。",
            "标准差": "样本数据的离散程度。",
            "t 统计量": "用于衡量两个样本均值差异的程度。",
            "自由度": "在统计分析中能够自由取值的变量个数。",
            "p 值": "用于判断两个样本均值是否有显著差异的指标。",
            "置信区间": "均值差异的可能范围。"
        },
        'interpretation': {
            "t 统计量": "t 统计量的绝对值越大，说明两个样本均值的差异越显著。",
            "自由度": "自由度越大，t 分布越接近正态分布。",
            "p 值": "p 值小于显著性水平（通常为 0.05）时，拒绝原假设，认为两个样本均值存在显著差异；否则，接受原假设，认为两个样本均值无显著差异。",
            "置信区间": "如果置信区间不包含 0，说明两个样本均值存在显著差异。"
        }
    },
    'en': {
        'title': "Independent Samples T-Test Analysis",
        'select_button': "Select File",
        'analyze_button': "Analyze File",
        'file_not_found': "The file does not exist. Please select again.",
        'analysis_success': "Analysis completed. The results have been saved to {}\n",
        'no_save_path': "No save path selected. The results were not saved.",
        'analysis_error': "An error occurred while analyzing the file: {}",
        'switch_language': "Switch Language",
        'explanation': {
            "独立样本 t 检验": "Used to compare whether the means of two independent samples are significantly different.",
            "样本量": "The number of observations in each sample.",
            "均值": "The average value of the sample data.",
            "标准差": "The degree of dispersion of the sample data.",
            "t 统计量": "Used to measure the degree of difference between the means of two samples.",
            "自由度": "The number of variables that can take on independent values in a statistical analysis.",
            "p 值": "An indicator used to determine whether the means of two samples are significantly different.",
            "置信区间": "The possible range of the difference in means."
        },
        'interpretation': {
            "t 统计量": "The larger the absolute value of the t statistic, the more significant the difference between the means of the two samples.",
            "自由度": "The larger the degrees of freedom, the closer the t distribution is to the normal distribution.",
            "p 值": "When the p-value is less than the significance level (usually 0.05), the null hypothesis is rejected, indicating a significant difference between the means of the two samples; otherwise, the null hypothesis is accepted, indicating no significant difference.",
            "置信区间": "If the confidence interval does not contain 0, it indicates a significant difference between the means of the two samples."
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
            raise ValueError("数据中没有数值列，无法进行独立样本 t 检验。")
        if len(numerical_df.columns) != 2:
            raise ValueError("数据必须包含两列数值数据，用于独立样本 t 检验。")

        # 进行独立样本 t 检验
        sample1 = numerical_df.iloc[:, 0]
        sample2 = numerical_df.iloc[:, 1]
        t_stat, p_value = stats.ttest_ind(sample1, sample2)
        df = len(sample1) + len(sample2) - 2
        mean_diff = sample1.mean() - sample2.mean()
        std_err = stats.sem(sample1 - sample2)
        conf_int = stats.t.interval(0.95, df, loc=mean_diff, scale=std_err)

        # 计算样本量、均值和标准差
        sample_sizes = numerical_df.count()
        means = numerical_df.mean()
        stds = numerical_df.std()

        # 整理数据
        data = [
            ["独立样本 t 检验", t_stat, df, p_value, conf_int],
            ["样本量", sample_sizes.to_dict(), "", "", ""],
            ["均值", means.to_dict(), "", "", ""],
            ["标准差", stds.to_dict(), "", "", ""]
        ]
        headers = ["统计量", "t 统计量", "自由度", "p 值", "置信区间"]
        df = pd.DataFrame(data, columns=headers)

        # 添加解释说明
        explanations = LANGUAGES[current_language]['explanation']
        interpretations = LANGUAGES[current_language]['interpretation']
        explanation_df = pd.DataFrame([explanations])
        explanation_df = explanation_df.reindex(columns=["独立样本 t 检验", "样本量", "均值", "标准差", "t 统计量", "自由度", "p 值", "置信区间"])
        explanation_df.insert(0, "统计量_解释说明", "解释说明" if current_language == 'zh' else "Explanation")

        # 添加分析结果解读
        interpretation_df = pd.DataFrame([interpretations])
        interpretation_df = interpretation_df.reindex(columns=["t 统计量", "自由度", "p 值", "置信区间"])
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

            # 绘制图表
            plot_results(sample1, sample2)
        else:
            result_label.config(text=LANGUAGES[current_language]['no_save_path'])

    except Exception as e:
        result_label.config(text=LANGUAGES[current_language]['analysis_error'].format(str(e)))


def plot_results(sample1, sample2):
    # 柱状图
    plt.figure(figsize=(12, 8))
    plt.subplot(2, 2, 1)
    means = [sample1.mean(), sample2.mean()]
    labels = ['Sample 1', 'Sample 2']
    plt.bar(labels, means)
    plt.title('Bar Chart')
    plt.xlabel('Samples')
    plt.ylabel('Mean')

    # 误差线图
    plt.subplot(2, 2, 2)
    stds = [sample1.std(), sample2.std()]
    plt.errorbar(labels, means, yerr=stds, fmt='o')
    plt.title('Error Bar Chart')
    plt.xlabel('Samples')
    plt.ylabel('Mean')

    # 箱线图
    plt.subplot(2, 2, 3)
    plt.boxplot([sample1, sample2])
    plt.title('Box Plot')
    plt.xlabel('Samples')
    plt.ylabel('Value')

    # 折线图
    plt.subplot(2, 2, 4)
    plt.plot(sample1, label='Sample 1')
    plt.plot(sample2, label='Sample 2')
    plt.title('Line Chart')
    plt.xlabel('Index')
    plt.ylabel('Value')
    plt.legend()

    plt.tight_layout()
    plt.show()


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