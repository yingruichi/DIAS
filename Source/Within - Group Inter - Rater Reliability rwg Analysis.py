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
import pingouin as pg  # 用于计算 ICC

# 设置 matplotlib 支持中文
plt.rcParams['font.family'] = 'SimHei'  # 设置字体为黑体，可根据系统情况修改为其他支持中文的字体
plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题

# 定义语言字典
LANGUAGES = {
    'zh': {
        'title': "组内评分者信度rwg分析",
        'select_button': "选择文件",
        'analyze_button': "分析文件",
        'file_not_found': "文件不存在，请重新选择。",
        'analysis_success': "分析完成，结果已保存到 {}\n",
        'no_save_path': "未选择保存路径，结果未保存。",
        'analysis_error': "分析文件时出错: {}",
        'switch_language': "切换语言",
        'explanation': {
            "rwg值": "组内评分者信度rwg用于评估组内成员评分的一致性，值越接近1表示一致性越高。",
            "Rwg值标准差SD": "Rwg值的标准差，反映了Rwg值的离散程度。",
            "P25": "Rwg值的第25百分位数。",
            "中位数": "Rwg值的中位数。",
            "P75": "Rwg值的第75百分位数。",
            "ICC1": "组内相关系数1，用于衡量组内评分者之间的一致性。",
            "ICC2": "组内相关系数2，考虑了评分者和项目的交互作用。",
            "MSB": "组间均方，反映了组间差异。",
            "MSW": "组内均方，反映了组内差异。",
            "F值": "F检验统计量，用于检验组间差异是否显著。",
            "p值": "F检验的p值，用于判断组间差异是否显著。"
        },
        'interpretation': {
            "rwg值": "rwg值越接近1，说明组内成员的评分越一致；值越低，说明组内成员的评分差异越大。",
            "Rwg值标准差SD": "标准差越大，说明Rwg值的离散程度越大。",
            "P25": "第25百分位数较低表示有25%的Rwg值低于该值。",
            "中位数": "中位数反映了Rwg值的中间水平。",
            "P75": "第75百分位数较高表示有75%的Rwg值低于该值。",
            "ICC1": "ICC1值越接近1，组内评分者之间的一致性越高。",
            "ICC2": "ICC2值越接近1，考虑交互作用后组内评分者之间的一致性越高。",
            "MSB": "MSB值越大，组间差异越明显。",
            "MSW": "MSW值越大，组内差异越明显。",
            "F值": "F值越大，说明组间差异越可能显著。",
            "p值": "p值小于0.05时，说明组间差异显著。"
        }
    },
    'en': {
        'title': "Within-Group Inter-Rater Reliability rwg Analysis",
        'select_button': "Select Files",
        'analyze_button': "Analyze Files",
        'file_not_found': "The file does not exist. Please select again.",
        'analysis_success': "Analysis completed. The results have been saved to {}\n",
        'no_save_path': "No save path selected. The results were not saved.",
        'analysis_error': "An error occurred while analyzing the file: {}",
        'switch_language': "Switch Language",
        'explanation': {
            "rwg值": "The within-group inter-rater reliability rwg is used to evaluate the consistency of ratings within a group. A value closer to 1 indicates higher consistency.",
            "Rwg值标准差SD": "The standard deviation of the rwg values, reflecting the dispersion of the rwg values.",
            "P25": "The 25th percentile of the rwg values.",
            "中位数": "The median of the rwg values.",
            "P75": "The 75th percentile of the rwg values.",
            "ICC1": "Intraclass correlation coefficient 1, used to measure the consistency between raters within a group.",
            "ICC2": "Intraclass correlation coefficient 2, considering the interaction between raters and items.",
            "MSB": "Mean square between groups, reflecting the differences between groups.",
            "MSW": "Mean square within groups, reflecting the differences within groups.",
            "F值": "F-test statistic, used to test whether the differences between groups are significant.",
            "p值": "The p-value of the F-test, used to determine whether the differences between groups are significant."
        },
        'interpretation': {
            "rwg值": "The closer the rwg value is to 1, the more consistent the ratings within the group; the lower the value, the greater the difference in ratings within the group.",
            "Rwg值标准差SD": "A larger standard deviation indicates a greater dispersion of the rwg values.",
            "P25": "A lower 25th percentile means that 25% of the rwg values are below this value.",
            "中位数": "The median reflects the middle level of the rwg values.",
            "P75": "A higher 75th percentile means that 75% of the rwg values are below this value.",
            "ICC1": "The closer the ICC1 value is to 1, the higher the consistency between raters within the group.",
            "ICC2": "The closer the ICC2 value is to 1, the higher the consistency between raters within the group considering the interaction.",
            "MSB": "A larger MSB value indicates more obvious differences between groups.",
            "MSW": "A larger MSW value indicates more obvious differences within groups.",
            "F值": "A larger F value indicates that the differences between groups are more likely to be significant.",
            "p值": "When the p-value is less than 0.05, the differences between groups are significant."
        }
    }
}

# 当前语言
current_language = 'en'


def select_file():
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_paths:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, ", ".join(file_paths))
        file_entry.configure(style="TEntry")  # 恢复默认样式


def rwg_analysis(data, group_column, rating_columns):
    all_results = []
    groups = data[group_column].unique()
    for group in groups:
        group_data = data[data[group_column] == group][rating_columns]
        # 检查数据是否为空或者只有一个样本
        if group_data.shape[0] < 2:
            print(f"Group {group} has less than 2 samples. Skipping...")
            continue
        k = group_data.shape[1]  # 评分者数量
        n = group_data.shape[0]  # 项目数量
        var_within = group_data.var(axis=1).mean()
        expected_var = (k ** 2 - 1) / 12
        rwg = 1 - (var_within / expected_var)
        result = {
            f"{group}_rwg值": rwg
        }
        all_results.append(result)
    return all_results

    # 计算 Rwg 值的统计量
    rwg_sd = np.std(rwg_values)
    rwg_p25 = np.percentile(rwg_values, 25)
    rwg_median = np.median(rwg_values)
    rwg_p75 = np.percentile(rwg_values, 75)

    # 计算 ICC1 和 ICC2
    icc_data = pd.melt(data, id_vars=[group_column], value_vars=rating_columns)
    icc_data.columns = ['Group', 'Rater', 'Score']
    icc = pg.intraclass_corr(data=icc_data, targets='Group', raters='Rater', ratings='Score')
    icc1 = icc[icc['Type'] == 'ICC1']['ICC'].values[0]
    icc2 = icc[icc['Type'] == 'ICC2']['ICC'].values[0]

    # 计算 MSB, MSW, F 值, p 值
    anova = pg.anova(data=icc_data, dv='Score', between='Group')
    msb = anova['MS'][0]
    msw = anova['MS'][1]
    f_value = anova['F'][0]
    p_value = anova['p-unc'][0]

    additional_stats = {
        "Rwg值标准差SD": rwg_sd,
        "P25": rwg_p25,
        "中位数": rwg_median,
        "P75": rwg_p75,
        "ICC1": icc1,
        "ICC2": icc2,
        "MSB": msb,
        "MSW": msw,
        "F值": f_value,
        "p值": p_value
    }
    all_results.append(additional_stats)
    return all_results


def analyze_file():
    global current_language
    file_paths = file_entry.get().split(", ")
    if not file_paths or file_paths[0] == "请输入待分析 Excel 文件的完整路径" or file_paths[0] == "Please enter the full path of the Excel file to be analyzed":
        file_paths = []
    for file_path in file_paths:
        if not os.path.exists(file_path):
            result_label.config(text=LANGUAGES[current_language]['file_not_found'])
            return
    try:
        group_column = tkinter.simpledialog.askstring("输入信息", "请输入分组列的列名")
        if not group_column:
            result_label.config(text="未输入有效的分组列名，分析取消。")
            return
        rating_columns = []
        while True:
            rating_column = tkinter.simpledialog.askstring("输入信息", "请输入评分列的列名（点击取消结束输入）")
            if rating_column is None:
                break
            if rating_column.strip():
                rating_columns.append(rating_column.strip())
            else:
                print("输入的列名不能为空，请重新输入。")

        if not rating_columns:
            result_label.config(text="未输入有效的评分列名，分析取消。")
            return

        all_results = []
        file_names = []
        for file_path in file_paths:
            # 打开 Excel 文件
            df = pd.read_excel(file_path)

            # 进行rwg分析
            rwg_results = rwg_analysis(df, group_column, rating_columns)
            all_results.extend(rwg_results)
            file_names.extend([os.path.basename(file_path)] * len(rwg_results))

        # 整理数据
        all_data = []
        for i, results in enumerate(all_results):
            if results:
                data = []
                for key, value in results.items():
                    data.append([f"{file_names[i]}_{key}", value])
                all_data.extend(data)
        headers = ["指标", "数值"]
        df_result = pd.DataFrame(all_data, columns=headers)

        # 添加解释说明
        explanations = LANGUAGES[current_language]['explanation']
        interpretations = LANGUAGES[current_language]['interpretation']
        explanation_df = pd.DataFrame([explanations])
        explanation_df = explanation_df.reindex(
            columns=["rwg值", "Rwg值标准差SD", "P25", "中位数", "P75", "ICC1", "ICC2", "MSB", "MSW", "F值", "p值"])
        explanation_df.insert(0, "指标_解释说明",
                              "解释说明" if current_language == 'zh' else "Explanation")

        # 添加分析结果解读
        interpretation_df = pd.DataFrame([interpretations])
        interpretation_df = interpretation_df.reindex(
            columns=["rwg值", "Rwg值标准差SD", "P25", "中位数", "P75", "ICC1", "ICC2", "MSB", "MSW", "F值", "p值"])
        interpretation_df.insert(0, "指标_结果解读",
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

            # 生成rwg值柱状图
            rwg_values = [result[f"{list(result.keys())[0].split('_')[0]}_rwg值"] for result in all_results if 'rwg值' in list(result.keys())[0]]
            group_names = [list(result.keys())[0].split('_')[0] for result in all_results if 'rwg值' in list(result.keys())[0]]
            fig, ax = plt.subplots()
            ax.bar(group_names, rwg_values)
            ax.set_title('组内评分者信度rwg值' if current_language == 'zh' else 'Within-Group Inter-Rater Reliability rwg Values')
            ax.set_ylabel('rwg值' if current_language == 'zh' else 'rwg Value')
            ax.set_xlabel('分组' if current_language == 'zh' else 'Group')
            plt.xticks(rotation=45)

            # 保存图片
            img_path = os.path.splitext(save_path)[0] + '_rwg.png'
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