import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox
import openpyxl
import os
import numpy as np
import pandas as pd
from tkinter import filedialog
import tkinter as tk
import tkinter.simpledialog  # 新增导入
import matplotlib.pyplot as plt
import pathlib
import statsmodels.api as sm

# 设置 matplotlib 支持中文
plt.rcParams['font.family'] = 'SimHei'  # 设置字体为黑体，可根据系统情况修改为其他支持中文的字体
plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题

# 定义语言字典
LANGUAGES = {
    'zh': {
        'title': "中介作用分析",
        'select_button': "选择文件",
        'analyze_button': "分析文件",
        'file_not_found': "文件不存在，请重新选择。",
        'analysis_success': "分析完成，结果已保存到 {}\n",
        'no_save_path': "未选择保存路径，结果未保存。",
        'analysis_error': "分析文件时出错: {}",
        'switch_language': "切换语言",
        'explanation': {
            "自变量对因变量的总效应": "自变量直接对因变量产生的影响。",
            "自变量对中介变量的效应": "自变量对中介变量产生的影响。",
            "中介变量对因变量的效应（控制自变量）": "在控制自变量的情况下，中介变量对因变量产生的影响。",
            "中介效应": "自变量通过中介变量对因变量产生的间接影响。",
            "样本量": "参与分析的样本数量。"
        },
        'interpretation': {
            "自变量对因变量的总效应": "总效应显著表示自变量对因变量有直接影响。",
            "自变量对中介变量的效应": "该效应显著表示自变量能够影响中介变量。",
            "中介变量对因变量的效应（控制自变量）": "此效应显著表示中介变量在控制自变量后仍对因变量有影响。",
            "中介效应": "中介效应显著表示自变量通过中介变量对因变量产生了间接影响。",
            "样本量": "样本量的大小会影响统计结果的可靠性，较大的样本量通常能提供更可靠的结果。"
        }
    },
    'en': {
        'title': "Mediation Analysis",
        'select_button': "Select File",
        'analyze_button': "Analyze File",
        'file_not_found': "The file does not exist. Please select again.",
        'analysis_success': "Analysis completed. The results have been saved to {}\n",
        'no_save_path': "No save path selected. The results were not saved.",
        'analysis_error': "An error occurred while analyzing the file: {}",
        'switch_language': "Switch Language",
        'explanation': {
            "自变量对因变量的总效应": "The total effect of the independent variable on the dependent variable.",
            "自变量对中介变量的效应": "The effect of the independent variable on the mediator variable.",
            "中介变量对因变量的效应（控制自变量）": "The effect of the mediator variable on the dependent variable while controlling for the independent variable.",
            "中介效应": "The indirect effect of the independent variable on the dependent variable through the mediator variable.",
            "样本量": "The number of samples involved in the analysis."
        },
        'interpretation': {
            "自变量对因变量的总效应": "A significant total effect indicates that the independent variable has a direct impact on the dependent variable.",
            "自变量对中介变量的效应": "A significant effect indicates that the independent variable can influence the mediator variable.",
            "中介变量对因变量的效应（控制自变量）": "A significant effect indicates that the mediator variable still has an impact on the dependent variable after controlling for the independent variable.",
            "中介效应": "A significant mediation effect indicates that the independent variable has an indirect impact on the dependent variable through the mediator variable.",
            "样本量": "The sample size affects the reliability of the statistical results. A larger sample size usually provides more reliable results."
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


def mediation_analysis(data, ind_var, med_var, dep_var):
    # 第一步：自变量对因变量的总效应
    X1 = data[ind_var]
    X1 = sm.add_constant(X1)
    model1 = sm.OLS(data[dep_var], X1).fit()
    total_effect = model1.params[ind_var]
    p_value_total = model1.pvalues[ind_var]

    # 第二步：自变量对中介变量的效应
    X2 = data[ind_var]
    X2 = sm.add_constant(X2)
    model2 = sm.OLS(data[med_var], X2).fit()
    effect_ind_med = model2.params[ind_var]
    p_value_ind_med = model2.pvalues[ind_var]

    # 第三步：中介变量对因变量的效应（控制自变量）
    X3 = data[[ind_var, med_var]]
    X3 = sm.add_constant(X3)
    model3 = sm.OLS(data[dep_var], X3).fit()
    effect_med_dep = model3.params[med_var]
    p_value_med_dep = model3.pvalues[med_var]

    # 第四步：中介效应
    mediation_effect = effect_ind_med * effect_med_dep

    sample_size = len(data)

    return total_effect, p_value_total, effect_ind_med, p_value_ind_med, effect_med_dep, p_value_med_dep, mediation_effect, sample_size


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

        # 让用户输入自变量、中介变量和因变量的列名
        ind_var = tkinter.simpledialog.askstring("输入信息", "请输入自变量的列名")
        med_var = tkinter.simpledialog.askstring("输入信息", "请输入中介变量的列名")
        dep_var = tkinter.simpledialog.askstring("输入信息", "请输入因变量的列名")

        if not ind_var or not med_var or not dep_var:
            result_label.config(text="未输入完整的变量名，分析取消。")
            return

        # 进行中介作用分析
        total_effect, p_value_total, effect_ind_med, p_value_ind_med, effect_med_dep, p_value_med_dep, mediation_effect, sample_size = mediation_analysis(
            df, ind_var, med_var, dep_var)

        # 整理数据
        data = [
            ["自变量对因变量的总效应", total_effect, p_value_total],
            ["自变量对中介变量的效应", effect_ind_med, p_value_ind_med],
            ["中介变量对因变量的效应（控制自变量）", effect_med_dep, p_value_med_dep],
            ["中介效应", mediation_effect, ""],
            ["样本量", sample_size, ""]
        ]
        headers = ["统计量", "统计量值", "p值"]
        df_result = pd.DataFrame(data, columns=headers)

        # 添加解释说明
        explanations = LANGUAGES[current_language]['explanation']
        interpretations = LANGUAGES[current_language]['interpretation']
        explanation_df = pd.DataFrame([explanations])
        explanation_df = explanation_df.reindex(
            columns=["自变量对因变量的总效应", "自变量对中介变量的效应", "中介变量对因变量的效应（控制自变量）", "中介效应", "样本量"])
        explanation_df.insert(0, "统计量_解释说明",
                              "解释说明" if current_language == 'zh' else "Explanation")

        # 添加分析结果解读
        interpretation_df = pd.DataFrame([interpretations])
        interpretation_df = interpretation_df.reindex(
            columns=["自变量对因变量的总效应", "自变量对中介变量的效应", "中介变量对因变量的效应（控制自变量）", "中介效应", "样本量"])
        interpretation_df.insert(0, "统计量_结果解读",
                                 "结果解读" if current_language == 'zh' else "Interpretation")

        # 合并数据、解释说明和结果解读
        combined_df = pd.concat(
            [df_result, explanation_df, interpretation_df], ignore_index=True)

        # 让用户选择保存路径
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            # 保存到 Excel 文件
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                combined_df.to_excel(writer, index=False)
                worksheet = writer.sheets['Sheet1']
                # 自动调整列宽
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = openpyxl.utils.get_column_letter(
                        column[0].column)
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

            # 生成图片（这里简单示例为中介效应柱状图）
            fig, ax = plt.subplots()
            effects = [total_effect, effect_ind_med, effect_med_dep, mediation_effect]
            labels = ["自变量对因变量总效应", "自变量对中介变量效应", "中介变量对因变量效应", "中介效应"] if current_language == 'zh' else [
                "Total Effect of Independent on Dependent", "Effect of Independent on Mediator",
                "Effect of Mediator on Dependent (Controlling Independent)", "Mediation Effect"]
            ax.bar(labels, effects)
            ax.set_title('中介作用分析结果' if current_language == 'zh' else 'Mediation Analysis Results')
            ax.set_ylabel('效应值' if current_language == 'zh' else 'Effect Value')

            # 保存图片
            img_path = os.path.splitext(save_path)[0] + '.png'
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