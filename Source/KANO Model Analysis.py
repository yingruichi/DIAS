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

# 设置 matplotlib 支持中文
plt.rcParams['font.family'] = 'SimHei'  # 设置字体为黑体，可根据系统情况修改为其他支持中文的字体
plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题

# 定义语言字典
LANGUAGES = {
    'zh': {
        'title': "KANO模型分析",
        'select_button': "选择文件",
        'analyze_button': "分析文件",
        'file_not_found': "文件不存在，请重新选择。",
        'analysis_success': "分析完成，结果已保存到 {}\n",
        'no_save_path': "未选择保存路径，结果未保存。",
        'analysis_error': "分析文件时出错: {}",
        'switch_language': "切换语言",
        'explanation': {
            "基本型需求（M）": "用户认为产品必须具备的功能，缺乏这些功能会导致用户不满。",
            "期望型需求（O）": "用户的满意度随该需求的满足程度而线性增加。",
            "兴奋型需求（A）": "用户没有预期到的需求，满足这些需求会极大提高用户满意度。",
            "无差异型需求（I）": "用户对该需求的满足与否不太关心。",
            "反向型需求（R）": "满足该需求会导致用户不满。",
            "可疑结果（Q）": "回答存在矛盾，结果不可靠。"
        },
        'interpretation': {
            "基本型需求（M）": "应确保产品满足基本型需求，以避免用户不满。",
            "期望型需求（O）": "可根据资源情况逐步提升期望型需求的满足程度，以提高用户满意度。",
            "兴奋型需求（A）": "挖掘和满足兴奋型需求可以使产品脱颖而出，吸引更多用户。",
            "无差异型需求（I）": "可以适当减少在无差异型需求上的投入。",
            "反向型需求（R）": "应避免满足反向型需求，以免引起用户不满。",
            "可疑结果（Q）": "需要重新确认用户回答，确保结果可靠性。"
        },
        'better_worse': {
            'better': 'Better系数',
            'worse': 'Worse系数'
        }
    },
    'en': {
        'title': "KANO Model Analysis",
        'select_button': "Select File",
        'analyze_button': "Analyze File",
        'file_not_found': "The file does not exist. Please select again.",
        'analysis_success': "Analysis completed. The results have been saved to {}\n",
        'no_save_path': "No save path selected. The results were not saved.",
        'analysis_error': "An error occurred while analyzing the file: {}",
        'switch_language': "Switch Language",
        'explanation': {
            "基本型需求（M）": "Basic requirements that users expect the product to have. Lack of these features will lead to user dissatisfaction.",
            "期望型需求（O）": "Expected requirements where user satisfaction increases linearly with the degree of fulfillment.",
            "兴奋型需求（A）": "Exciting requirements that users do not expect. Meeting these requirements can greatly improve user satisfaction.",
            "无差异型需求（I）": "Indifferent requirements that users do not care much about whether they are met or not.",
            "反向型需求（R）": "Reverse requirements where meeting them will lead to user dissatisfaction.",
            "可疑结果（Q）": "The responses are contradictory, and the results are unreliable."
        },
        'interpretation': {
            "基本型需求（M）": "Ensure that the product meets basic requirements to avoid user dissatisfaction.",
            "期望型需求（O）": "Gradually improve the fulfillment of expected requirements according to available resources to enhance user satisfaction.",
            "兴奋型需求（A）": "Discover and meet exciting requirements to make the product stand out and attract more users.",
            "无差异型需求（I）": "Reduce investment in indifferent requirements appropriately.",
            "反向型需求（R）": "Avoid meeting reverse requirements to prevent user dissatisfaction.",
            "可疑结果（Q）": "Reconfirm user responses to ensure result reliability."
        },
        'better_worse': {
            'better': 'Better Coefficient',
            'worse': 'Worse Coefficient'
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


def kano_analysis(data, positive_question_columns, negative_question_columns):
    kano_results = {}
    better_worse_results = {}
    for i in range(len(positive_question_columns)):
        positive_responses = data[positive_question_columns[i]]
        negative_responses = data[negative_question_columns[i]]
        category = classify_kano(positive_responses, negative_responses)
        kano_results[positive_question_columns[i]] = category
        better, worse = calculate_better_worse(positive_responses, negative_responses)
        better_worse_results[positive_question_columns[i]] = (better, worse)
    return kano_results, better_worse_results


def classify_kano(positive_responses, negative_responses):
    counts = {
        'A': 0, 'O': 0, 'M': 0, 'I': 0, 'R': 0, 'Q': 0
    }
    for pos, neg in zip(positive_responses, negative_responses):
        if pos == 5 and neg == 1:
            counts['A'] += 1
        elif pos == 5 and neg == 2:
            counts['A'] += 1
        elif pos == 5 and neg == 3:
            counts['O'] += 1
        elif pos == 4 and neg == 1:
            counts['A'] += 1
        elif pos == 4 and neg == 2:
            counts['O'] += 1
        elif pos == 4 and neg == 3:
            counts['O'] += 1
        elif pos == 3 and neg == 1:
            counts['O'] += 1
        elif pos == 3 and neg == 2:
            counts['O'] += 1
        elif pos == 3 and neg == 3:
            counts['I'] += 1
        elif pos == 2 and neg == 1:
            counts['I'] += 1
        elif pos == 2 and neg == 2:
            counts['I'] += 1
        elif pos == 2 and neg == 3:
            counts['M'] += 1
        elif pos == 1 and neg == 1:
            counts['R'] += 1
        elif pos == 1 and neg == 2:
            counts['M'] += 1
        elif pos == 1 and neg == 3:
            counts['M'] += 1
        else:
            counts['Q'] += 1
    max_count_category = max(counts, key=counts.get)
    category_mapping = {
        'A': "兴奋型需求（A）",
        'O': "期望型需求（O）",
        'M': "基本型需求（M）",
        'I': "无差异型需求（I）",
        'R': "反向型需求（R）",
        'Q': "可疑结果（Q）"
    }
    return category_mapping[max_count_category]


def calculate_better_worse(positive_responses, negative_responses):
    a_count = 0
    o_count = 0
    m_count = 0
    i_count = 0
    r_count = 0
    total_count = len(positive_responses)
    for pos, neg in zip(positive_responses, negative_responses):
        if pos == 5 and neg == 1:
            a_count += 1
        elif pos == 5 and neg == 2:
            a_count += 1
        elif pos == 5 and neg == 3:
            o_count += 1
        elif pos == 4 and neg == 1:
            a_count += 1
        elif pos == 4 and neg == 2:
            o_count += 1
        elif pos == 4 and neg == 3:
            o_count += 1
        elif pos == 3 and neg == 1:
            o_count += 1
        elif pos == 3 and neg == 2:
            o_count += 1
        elif pos == 3 and neg == 3:
            i_count += 1
        elif pos == 2 and neg == 1:
            i_count += 1
        elif pos == 2 and neg == 2:
            i_count += 1
        elif pos == 2 and neg == 3:
            m_count += 1
        elif pos == 1 and neg == 1:
            r_count += 1
        elif pos == 1 and neg == 2:
            m_count += 1
        elif pos == 1 and neg == 3:
            m_count += 1
    better = (a_count + o_count) / total_count
    worse = -(m_count + r_count) / total_count
    return better, worse


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

        # 让用户输入正向问题和负向问题的列名
        positive_question_columns = tkinter.simpledialog.askstring("输入信息", "请输入正向问题的列名，用逗号分隔").split(',')
        negative_question_columns = tkinter.simpledialog.askstring("输入信息", "请输入负向问题的列名，用逗号分隔").split(',')

        if not positive_question_columns or not negative_question_columns:
            result_label.config(text="未输入完整的问题列名，分析取消。")
            return

        # 进行KANO模型分析
        kano_results, better_worse_results = kano_analysis(df, positive_question_columns, negative_question_columns)

        # 整理数据
        data = []
        for question, category in kano_results.items():
            better, worse = better_worse_results[question]
            data.append([question, category, better, worse])
        headers = ["问题", "KANO分类", LANGUAGES[current_language]['better_worse']['better'], LANGUAGES[current_language]['better_worse']['worse']]
        df_result = pd.DataFrame(data, columns=headers)

        # 添加解释说明
        explanations = LANGUAGES[current_language]['explanation']
        interpretations = LANGUAGES[current_language]['interpretation']
        explanation_df = pd.DataFrame([explanations])
        explanation_df = explanation_df.reindex(
            columns=["基本型需求（M）", "期望型需求（O）", "兴奋型需求（A）", "无差异型需求（I）", "反向型需求（R）", "可疑结果（Q）"])
        explanation_df.insert(0, "KANO分类_解释说明",
                              "解释说明" if current_language == 'zh' else "Explanation")

        # 添加分析结果解读
        interpretation_df = pd.DataFrame([interpretations])
        interpretation_df = interpretation_df.reindex(
            columns=["基本型需求（M）", "期望型需求（O）", "兴奋型需求（A）", "无差异型需求（I）", "反向型需求（R）", "可疑结果（Q）"])
        interpretation_df.insert(0, "KANO分类_结果解读",
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

            # 生成 Better 和 Worse 象限图
            generate_better_worse_plot(df_result, save_path, current_language)

            # 生成 KANO 图
            generate_kano_plot(df_result, save_path, current_language)

        else:
            result_label.config(text=LANGUAGES[current_language]['no_save_path'])

    except Exception as e:
        result_label.config(text=LANGUAGES[current_language]['analysis_error'].format(str(e)))


def generate_better_worse_plot(df_result, save_path, language):
    better = df_result[LANGUAGES[language]['better_worse']['better']]
    worse = df_result[LANGUAGES[language]['better_worse']['worse']]
    labels = df_result['问题']

    plt.figure(figsize=(10, 8))
    plt.scatter(better, worse)
    for i, label in enumerate(labels):
        plt.annotate(label, (better[i], worse[i]), textcoords="offset points", xytext=(0, 10), ha='center')
    plt.axhline(y=0, color='k')
    plt.axvline(x=0, color='k')
    plt.xlabel(LANGUAGES[language]['better_worse']['better'])
    plt.ylabel(LANGUAGES[language]['better_worse']['worse'])
    plt.title('Better - Worse 象限图' if language == 'zh' else 'Better - Worse Quadrant Plot')
    img_path = os.path.splitext(save_path)[0] + '_better_worse.png'
    plt.savefig(img_path)
    plt.close()


def generate_kano_plot(df_result, save_path, language):
    category_counts = df_result['KANO分类'].value_counts()
    plt.figure(figsize=(10, 8))
    plt.bar(category_counts.index, category_counts.values)
    plt.xlabel('KANO分类' if language == 'zh' else 'KANO Category')
    plt.ylabel('数量' if language == 'zh' else 'Count')
    plt.title('KANO模型分析结果' if language == 'zh' else 'KANO Model Analysis Results')
    img_path = os.path.splitext(save_path)[0] + '_kano.png'
    plt.savefig(img_path)
    plt.close()


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