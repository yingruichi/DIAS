import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import openpyxl
import os
import numpy as np
import pandas as pd
from tkinter import filedialog
import tkinter as tk
from scipy import stats
import matplotlib.pyplot as plt

# 设置支持中文的字体
plt.rcParams['font.family'] = 'SimHei'  # 可以根据系统情况选择 'Microsoft YaHei' 等
# 解决负号显示问题
plt.rcParams['axes.unicode_minus'] = False

# 定义语言字典
LANGUAGES = {
    'zh': {
        'title': "配对样本Wilcoxon检验分析",
        'select_button': "选择文件",
        'analyze_button': "分析文件",
        'file_not_found': "文件不存在，请重新选择。",
        'analysis_success': "分析完成，结果已保存到 {}\n",
        'no_save_path': "未选择保存路径，结果未保存。",
        'analysis_error': "分析文件时出错: {}",
        'switch_language': "切换语言",
        'explanation': {
            "配对样本Wilcoxon检验": "用于检验配对样本之间是否存在显著差异。",
            "样本量": "配对样本中的观测值对数。",
            "中位数": "配对样本差值数据的中间值，将数据分为上下两部分。",
            "t统计量": "配对样本Wilcoxon检验的统计量值，用于衡量配对样本之间的差异程度。",
            "自由度": "在统计计算中能够自由变化的变量个数，对于配对样本Wilcoxon检验，自由度等于样本量减1。",
            "p值": "p值小于显著性水平（通常为0.05）时，拒绝原假设，认为配对样本之间存在显著差异；否则，接受原假设，认为无显著差异。",
            "均值差异的置信区间": "包含真实均值差异的一个区间，反映了估计的不确定性。"
        },
        'interpretation': {
            "t统计量": "t统计量的绝对值越大，说明配对样本之间的差异越显著。",
            "自由度": "自由度影响t分布的形状，进而影响p值的计算。",
            "p值": "用于判断配对样本之间是否存在显著差异的依据。",
            "样本量": "样本量的大小会影响统计检验的功效，较大的样本量通常能提供更准确的结果。",
            "中位数": "中位数反映了配对样本差值数据的中心位置，可用于比较配对样本之间的差异。",
            "均值差异的置信区间": "如果置信区间不包含0，说明配对样本之间存在显著差异。"
        }
    },
    'en': {
        'title': "Paired-Sample Wilcoxon Test Analysis",
        'select_button': "Select File",
        'analyze_button': "Analyze File",
        'file_not_found': "The file does not exist. Please select again.",
        'analysis_success': "Analysis completed. The results have been saved to {}\n",
        'no_save_path': "No save path selected. The results were not saved.",
        'analysis_error': "An error occurred while analyzing the file: {}",
        'switch_language': "Switch Language",
        'explanation': {
            "配对样本Wilcoxon检验": "Used to test whether there is a significant difference between paired samples.",
            "样本量": "The number of pairs of observations in the paired samples.",
            "中位数": "The middle value of the difference data of the paired samples, dividing the data into two parts.",
            "t统计量": "The test statistic value of the paired-sample Wilcoxon test, used to measure the degree of difference between paired samples.",
            "自由度": "The number of independent variables in a statistical calculation. For the paired-sample Wilcoxon test, the degrees of freedom equal the sample size minus 1.",
            "p值": "When the p-value is less than the significance level (usually 0.05), the null hypothesis is rejected, indicating a significant difference between paired samples; otherwise, the null hypothesis is accepted, indicating no significant difference.",
            "均值差异的置信区间": "An interval that contains the true mean difference, reflecting the uncertainty of the estimate."
        },
        'interpretation': {
            "t统计量": "The larger the absolute value of the t-statistic, the more significant the difference between paired samples.",
            "自由度": "The degrees of freedom affect the shape of the t-distribution, which in turn affects the calculation of the p-value.",
            "p值": "The basis for determining whether there is a significant difference between paired samples.",
            "样本量": "The sample size affects the power of the statistical test. A larger sample size usually provides more accurate results.",
            "中位数": "The median reflects the central position of the difference data of the paired samples and can be used to compare the difference between paired samples.",
            "均值差异的置信区间": "If the confidence interval does not contain 0, it indicates a significant difference between paired samples."
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
            raise ValueError("数据中没有数值列，无法进行配对样本Wilcoxon检验。")
        if numerical_df.shape[1] != 2:
            raise ValueError("数据必须包含两列数值数据，用于配对样本Wilcoxon检验。")

        # 进行配对样本Wilcoxon检验
        t_stat, p_value = stats.wilcoxon(numerical_df.iloc[:, 0], numerical_df.iloc[:, 1])

        # 计算样本量、中位数
        sample_size = numerical_df.count().values[0]
        differences = numerical_df.iloc[:, 0] - numerical_df.iloc[:, 1]
        median = differences.median()

        # 计算自由度
        degrees_of_freedom = sample_size - 1

        # 计算均值差异的置信区间
        mean_diff = differences.mean()
        std_err = differences.std() / np.sqrt(sample_size)
        confidence_interval = stats.t.interval(0.95, degrees_of_freedom, loc=mean_diff, scale=std_err)

        # 整理数据
        data = [
            ["配对样本Wilcoxon检验", t_stat, degrees_of_freedom, p_value, confidence_interval],
            ["样本量", sample_size, "", "", ""],
            ["中位数", median, "", "", ""]
        ]
        headers = ["统计量", "t统计量", "自由度", "p值", "均值差异的置信区间"]
        df = pd.DataFrame(data, columns=headers)

        # 添加解释说明
        explanations = LANGUAGES[current_language]['explanation']
        interpretations = LANGUAGES[current_language]['interpretation']
        explanation_df = pd.DataFrame([explanations])
        explanation_df = explanation_df.reindex(columns=["配对样本Wilcoxon检验", "样本量", "中位数", "t统计量", "自由度", "p值", "均值差异的置信区间"])
        explanation_df.insert(0, "统计量_解释说明", "解释说明" if current_language == 'zh' else "Explanation")

        # 添加分析结果解读
        interpretation_df = pd.DataFrame([interpretations])
        interpretation_df = interpretation_df.reindex(columns=["t统计量", "自由度", "p值", "样本量", "中位数", "均值差异的置信区间"])
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

            # 绘制柱状图
            plt.figure(figsize=(8, 6))
            plt.bar(['样本1', '样本2'], [numerical_df.iloc[:, 0].mean(), numerical_df.iloc[:, 1].mean()])
            plt.title('柱状图' if current_language == 'zh' else 'Bar Chart')
            plt.ylabel('均值' if current_language == 'zh' else 'Mean')
            plt.savefig(save_path.replace('.xlsx', '_bar_chart.png'))
            plt.close()

            # 绘制误差线图
            plt.figure(figsize=(8, 6))
            plt.errorbar(['样本1', '样本2'], [numerical_df.iloc[:, 0].mean(), numerical_df.iloc[:, 1].mean()],
                         yerr=[numerical_df.iloc[:, 0].std() / np.sqrt(sample_size),
                               numerical_df.iloc[:, 1].std() / np.sqrt(sample_size)], fmt='o')
            plt.title('误差线图' if current_language == 'zh' else 'Error Bar Chart')
            plt.ylabel('均值' if current_language == 'zh' else 'Mean')
            plt.savefig(save_path.replace('.xlsx', '_error_bar_chart.png'))
            plt.close()

            # 绘制箱线图
            plt.figure(figsize=(8, 6))
            numerical_df.plot(kind='box')
            plt.title('箱线图' if current_language == 'zh' else 'Box Plot')
            plt.ylabel('数值' if current_language == 'zh' else 'Value')
            plt.savefig(save_path.replace('.xlsx', '_box_plot.png'))
            plt.close()

            # 绘制折线图
            plt.figure(figsize=(8, 6))
            numerical_df.plot(kind='line')
            plt.title('折线图' if current_language == 'zh' else 'Line Chart')
            plt.ylabel('数值' if current_language == 'zh' else 'Value')
            plt.savefig(save_path.replace('.xlsx', '_line_chart.png'))
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