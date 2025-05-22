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
import pandas.plotting as pd_plotting

# 定义语言字典
LANGUAGES = {
    'zh': {
        'title': "Kendall相关性分析",
        'select_button': "选择文件",
        'analyze_button': "分析文件",
        'file_not_found': "文件不存在，请重新选择。",
        'analysis_success': "分析完成，结果已保存到 {}\n",
        'no_save_path': "未选择保存路径，结果未保存。",
        'analysis_error': "分析文件时出错: {}",
        'switch_language': "切换语言",
        'explanation': {
            "Kendall相关系数": "用于衡量两个有序变量之间的相关性，考虑了变量的顺序信息。"
        },
        'interpretation': {
            "相关系数": "相关系数的绝对值越接近1，说明两个变量之间的相关性越强；接近0则表示相关性较弱。",
            "p值": "p值小于显著性水平（通常为0.05）时，拒绝原假设，认为两个变量之间存在显著相关性；否则，接受原假设，认为两个变量之间无显著相关性。"
        }
    },
    'en': {
        'title': "Kendall Correlation Analysis",
        'select_button': "Select File",
        'analyze_button': "Analyze File",
        'file_not_found': "The file does not exist. Please select again.",
        'analysis_success': "Analysis completed. The results have been saved to {}\n",
        'no_save_path': "No save path selected. The results were not saved.",
        'analysis_error': "An error occurred while analyzing the file: {}",
        'switch_language': "Switch Language",
        'explanation': {
            "Kendall相关系数": "Used to measure the correlation between two ordinal variables, taking into account the order information of the variables."
        },
        'interpretation': {
            "相关系数": "The closer the absolute value of the correlation coefficient is to 1, the stronger the correlation between the two variables; close to 0 indicates a weak correlation.",
            "p值": "When the p-value is less than the significance level (usually 0.05), the null hypothesis is rejected, indicating a significant correlation between the two variables; otherwise, the null hypothesis is accepted, indicating no significant correlation."
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
            raise ValueError("数据中没有数值列，无法进行相关性分析。")

        # 计算Kendall相关性
        kendall_corr = numerical_df.corr(method='kendall')

        # 计算p值
        def calculate_pvalues(df):
            df = df.dropna()._get_numeric_data()
            dfcols = pd.DataFrame(columns=df.columns)
            pvalues = dfcols.transpose().join(dfcols, how='outer')
            for r in df.columns:
                for c in df.columns:
                    _, p = stats.kendalltau(df[r], df[c])
                    pvalues.loc[r, c] = p
            return pvalues

        kendall_pvalues = calculate_pvalues(numerical_df)

        # 整理数据
        data = []
        correlation_types = ["Kendall相关系数"]
        explanations = LANGUAGES[current_language]['explanation']
        interpretations = LANGUAGES[current_language]['interpretation']

        for i, (corr, pvalues) in enumerate(zip([kendall_corr], [kendall_pvalues])):
            for col1 in corr.columns:
                for col2 in corr.columns:
                    if col1 != col2:
                        data.append([f"{correlation_types[i]} ({col1} vs {col2})", corr.loc[col1, col2], pvalues.loc[col1, col2]])

        headers = ["统计量", "相关系数", "p值"]
        df = pd.DataFrame(data, columns=headers)

        # 添加解释说明
        explanation_df = pd.DataFrame([explanations])
        explanation_df = explanation_df.reindex(columns=correlation_types)
        explanation_df.insert(0, "统计量", "解释说明" if current_language == 'zh' else "Explanation")

        # 添加分析结果解读
        interpretation_df = pd.DataFrame([interpretations])
        interpretation_df = interpretation_df.reindex(columns=["相关系数", "p值"])
        interpretation_df.insert(0, "统计量", "结果解读" if current_language == 'zh' else "Interpretation")

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

            # 生成相关性热力图
            desktop_path = pathlib.Path.home() / 'Desktop'
            plot_path = desktop_path / 'correlation_heatmap.png'
            plt.figure(figsize=(10, 8))
            plt.imshow(kendall_corr, cmap='coolwarm', interpolation='nearest')
            plt.colorbar()
            plt.xticks(range(len(kendall_corr.columns)), kendall_corr.columns, rotation=45)
            plt.yticks(range(len(kendall_corr.columns)), kendall_corr.columns)
            for i in range(len(kendall_corr.columns)):
                for j in range(len(kendall_corr.columns)):
                    plt.text(j, i, f'{kendall_corr.iloc[i, j]:.2f}', ha='center', va='center', color='black')
            plt.title('Kendall Correlation Heatmap')
            plt.savefig(plot_path)
            plt.close()

            # 生成散点图矩阵
            scatter_matrix_path = desktop_path / 'scatter_matrix.png'
            pd_plotting.scatter_matrix(numerical_df, alpha=0.8, figsize=(10, 10), diagonal='hist')
            plt.suptitle('Scatter Matrix')
            plt.savefig(scatter_matrix_path)
            plt.close()

            # 生成相关性柱状图
            selected_variable = numerical_df.columns[0]
            correlation_column = kendall_corr[selected_variable]
            bar_plot_path = desktop_path / 'correlation_bar_plot.png'
            plt.figure(figsize=(10, 6))
            correlation_column.plot(kind='bar')
            plt.title(f'Correlation with {selected_variable}')
            plt.xlabel('Variables')
            plt.ylabel('Correlation Coefficient')
            plt.xticks(rotation=45)
            plt.tight_layout()
            plt.savefig(bar_plot_path)
            plt.close()

            result_msg = LANGUAGES[current_language]['analysis_success'].format(
                save_path) + f"相关性热力图已保存到 {plot_path}"
            result_msg += f"\n散点图矩阵已保存到 {scatter_matrix_path}"
            result_msg += f"\n相关性柱状图已保存到 {bar_plot_path}"
            result_label.config(text=result_msg, wraplength=400)
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