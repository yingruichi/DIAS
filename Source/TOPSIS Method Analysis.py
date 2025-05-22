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
        'title': "TOPSIS 法分析",
        'select_button': "选择文件",
        'analyze_button': "分析文件",
        'file_not_found': "文件不存在，请重新选择。",
        'analysis_success': "分析完成，结果已保存到 {}\n",
        'no_save_path': "未选择保存路径，结果未保存。",
        'analysis_error': "分析文件时出错: {}",
        'switch_language': "切换语言",
        'explanation': {
            "标准化决策矩阵": "对原始决策矩阵进行标准化处理后的矩阵",
            "加权标准化决策矩阵": "考虑各属性权重后的标准化决策矩阵",
            "正理想解": "各属性的最优值构成的向量",
            "负理想解": "各属性的最劣值构成的向量",
            "各方案到正理想解的距离": "各方案与正理想解的欧几里得距离",
            "各方案到负理想解的距离": "各方案与负理想解的欧几里得距离",
            "各方案的相对贴近度": "反映各方案与正理想解的相对接近程度",
            "方案排序结果": "根据相对贴近度对各方案进行排序的结果"
        },
        'interpretation': {
            "标准化决策矩阵": "消除不同属性量纲的影响",
            "加权标准化决策矩阵": "体现各属性在决策中的重要性",
            "正理想解": "作为衡量各方案优劣的最优参考点",
            "负理想解": "作为衡量各方案优劣的最劣参考点",
            "各方案到正理想解的距离": "距离越小，方案越优",
            "各方案到负理想解的距离": "距离越大，方案越优",
            "各方案的相对贴近度": "值越接近 1，方案越优",
            "方案排序结果": "排名越靠前，方案越优"
        }
    },
    'en': {
        'title': "TOPSIS Method Analysis",
        'select_button': "Select File",
        'analyze_button': "Analyze File",
        'file_not_found': "The file does not exist. Please select again.",
        'analysis_success': "Analysis completed. The results have been saved to {}\n",
        'no_save_path': "No save path selected. The results were not saved.",
        'analysis_error': "An error occurred while analyzing the file: {}",
        'switch_language': "Switch Language",
        'explanation': {
            "标准化决策矩阵": "The matrix after standardizing the original decision matrix",
            "加权标准化决策矩阵": "The standardized decision matrix considering the weights of each attribute",
            "正理想解": "The vector composed of the optimal values of each attribute",
            "负理想解": "The vector composed of the worst values of each attribute",
            "各方案到正理想解的距离": "The Euclidean distance between each alternative and the positive ideal solution",
            "各方案到负理想解的距离": "The Euclidean distance between each alternative and the negative ideal solution",
            "各方案的相对贴近度": "Reflects the relative closeness of each alternative to the positive ideal solution",
            "方案排序结果": "The result of ranking each alternative according to the relative closeness"
        },
        'interpretation': {
            "标准化决策矩阵": "Eliminate the influence of different attribute dimensions",
            "加权标准化决策矩阵": "Reflect the importance of each attribute in the decision-making",
            "正理想解": "As the optimal reference point for measuring the advantages and disadvantages of each alternative",
            "负理想解": "As the worst reference point for measuring the advantages and disadvantages of each alternative",
            "各方案到正理想解的距离": "The smaller the distance, the better the alternative",
            "各方案到负理想解的距离": "The larger the distance, the better the alternative",
            "各方案的相对贴近度": "The closer the value is to 1, the better the alternative",
            "方案排序结果": "The higher the ranking, the better the alternative"
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


def topsis_method(decision_matrix, weight_vector):
    """
    实现 TOPSIS 法
    :param decision_matrix: 决策矩阵
    :param weight_vector: 属性权重向量
    :return: 各方案的相对贴近度和方案排序结果
    """
    # 标准化决策矩阵
    standardized_matrix = decision_matrix / np.sqrt(np.sum(decision_matrix ** 2, axis=0))

    # 加权标准化决策矩阵
    weighted_matrix = standardized_matrix * weight_vector

    # 正理想解和负理想解
    positive_ideal_solution = np.max(weighted_matrix, axis=0)
    negative_ideal_solution = np.min(weighted_matrix, axis=0)

    # 各方案到正理想解和负理想解的距离
    distances_to_positive = np.sqrt(np.sum((weighted_matrix - positive_ideal_solution) ** 2, axis=1))
    distances_to_negative = np.sqrt(np.sum((weighted_matrix - negative_ideal_solution) ** 2, axis=1))

    # 各方案的相对贴近度
    relative_closeness = distances_to_negative / (distances_to_positive + distances_to_negative)

    # 方案排序结果
    ranking = np.argsort(-relative_closeness) + 1

    return standardized_matrix, weighted_matrix, positive_ideal_solution, negative_ideal_solution, \
           distances_to_positive, distances_to_negative, relative_closeness, ranking


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

        # 假设第一行为属性权重向量，其余行为决策矩阵
        weight_vector = data[0]
        decision_matrix = data[1:]

        # 进行 TOPSIS 分析
        standardized_matrix, weighted_matrix, positive_ideal_solution, negative_ideal_solution, \
        distances_to_positive, distances_to_negative, relative_closeness, ranking = topsis_method(decision_matrix,
                                                                                                 weight_vector)

        # 整理数据
        data = [
            ["标准化决策矩阵", standardized_matrix.tolist(), ""],
            ["加权标准化决策矩阵", weighted_matrix.tolist(), ""],
            ["正理想解", positive_ideal_solution.tolist(), ""],
            ["负理想解", negative_ideal_solution.tolist(), ""],
            ["各方案到正理想解的距离", distances_to_positive.tolist(), ""],
            ["各方案到负理想解的距离", distances_to_negative.tolist(), ""],
            ["各方案的相对贴近度", relative_closeness.tolist(), ""],
            ["方案排序结果", ranking.tolist(), ""]
        ]
        headers = ["统计量", "统计量值", "p值"]
        df = pd.DataFrame(data, columns=headers)

        # 添加解释说明
        explanations = LANGUAGES[current_language]['explanation']
        interpretations = LANGUAGES[current_language]['interpretation']
        explanation_df = pd.DataFrame([explanations])
        explanation_df = explanation_df.reindex(
            columns=["标准化决策矩阵", "加权标准化决策矩阵", "正理想解", "负理想解", "各方案到正理想解的距离",
                     "各方案到负理想解的距离", "各方案的相对贴近度", "方案排序结果"])
        explanation_df.insert(0, "统计量_解释说明", "解释说明" if current_language == 'zh' else "Explanation")

        # 添加分析结果解读
        interpretation_df = pd.DataFrame([interpretations])
        interpretation_df = interpretation_df.reindex(
            columns=["标准化决策矩阵", "加权标准化决策矩阵", "正理想解", "负理想解", "各方案到正理想解的距离",
                     "各方案到负理想解的距离", "各方案的相对贴近度", "方案排序结果"])
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

            # 生成各方案相对贴近度柱状图
            fig, ax = plt.subplots()
            ax.bar(range(len(relative_closeness)), relative_closeness)
            ax.set_title(
                '各方案相对贴近度柱状图' if current_language == 'zh' else 'Bar Chart of Relative Closeness of Each Alternative')
            ax.set_xlabel('方案编号' if current_language == 'zh' else 'Alternative Number')
            ax.set_ylabel('相对贴近度' if current_language == 'zh' else 'Relative Closeness')
            # 保存图片
            img_path = os.path.splitext(save_path)[0] + '_relative_closeness.png'
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