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
        'title': "NPS净推荐值分析",
        'select_button': "选择文件",
        'analyze_button': "分析文件",
        'file_not_found': "文件不存在，请重新选择。",
        'analysis_success': "分析完成，结果已保存到 {}\n",
        'no_save_path': "未选择保存路径，结果未保存。",
        'analysis_error': "分析文件时出错: {}",
        'switch_language': "切换语言",
        'explanation': {
            "推荐者": "给出9 - 10分的客户，是产品或服务的忠实拥护者，会积极推荐给他人。",
            "被动者": "给出7 - 8分的客户，对产品或服务基本满意，但不会主动推荐。",
            "贬损者": "给出0 - 6分的客户，对产品或服务不满意，可能会向他人抱怨。",
            "NPS净推荐值": "NPS = 推荐者比例 - 贬损者比例，反映了客户对产品或服务的整体态度。"
        },
        'interpretation': {
            "推荐者": "应关注推荐者的需求，提供更好的服务，鼓励他们继续推荐。",
            "被动者": "可以通过改进产品或服务，将被动者转化为推荐者。",
            "贬损者": "及时了解贬损者的不满原因，采取措施改进，避免负面影响扩大。",
            "NPS净推荐值": "NPS值越高，说明客户对产品或服务越满意，忠诚度越高。"
        }
    },
    'en': {
        'title': "NPS Net Promoter Score Analysis",
        'select_button': "Select Files",
        'analyze_button': "Analyze Files",
        'file_not_found': "The file does not exist. Please select again.",
        'analysis_success': "Analysis completed. The results have been saved to {}\n",
        'no_save_path': "No save path selected. The results were not saved.",
        'analysis_error': "An error occurred while analyzing the file: {}",
        'switch_language': "Switch Language",
        'explanation': {
            "推荐者": "Customers who give a score of 9 - 10 are loyal advocates of the product or service and will actively recommend it to others.",
            "被动者": "Customers who give a score of 7 - 8 are generally satisfied with the product or service but will not actively recommend it.",
            "贬损者": "Customers who give a score of 0 - 6 are dissatisfied with the product or service and may complain to others.",
            "NPS净推荐值": "NPS = Percentage of Promoters - Percentage of Detractors, which reflects the overall attitude of customers towards the product or service."
        },
        'interpretation': {
            "推荐者": "Pay attention to the needs of promoters, provide better services, and encourage them to continue recommending.",
            "被动者": "Improve the product or service to convert passives into promoters.",
            "贬损者": "Understand the reasons for detractors' dissatisfaction in a timely manner, take measures to improve, and avoid the expansion of negative impacts.",
            "NPS净推荐值": "The higher the NPS value, the more satisfied and loyal the customers are with the product or service."
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


def nps_analysis(data, question_columns):
    all_results = []
    for question_column in question_columns:
        if question_column not in data.columns:
            print(f"列名 {question_column} 不在文件中，跳过该列分析。")
            continue
        responses = data[question_column]
        promoters = (responses >= 9).sum()
        passives = ((responses >= 7) & (responses <= 8)).sum()
        detractors = (responses <= 6).sum()
        total = len(responses)
        promoter_percentage = promoters / total * 100
        passive_percentage = passives / total * 100
        detractor_percentage = detractors / total * 100
        nps = promoter_percentage - detractor_percentage

        # 各分数段占比情况
        score_bins = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
        score_counts = pd.cut(responses, bins=score_bins, right=False).value_counts().sort_index()
        score_percentages = score_counts / total * 100

        results = {
            f"{question_column}_推荐者数量": promoters,
            f"{question_column}_推荐者比例": promoter_percentage,
            f"{question_column}_被动者数量": passives,
            f"{question_column}_被动者比例": passive_percentage,
            f"{question_column}_贬损者数量": detractors,
            f"{question_column}_贬损者比例": detractor_percentage,
            f"{question_column}_NPS净推荐值": nps,
            f"{question_column}_各分数段占比情况": score_percentages
        }
        all_results.append(results)
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
        question_columns = []
        while True:
            question_column = tkinter.simpledialog.askstring("输入信息", "请输入NPS问题的列名（点击取消结束输入）")
            if question_column is None:
                break
            if question_column.strip():
                question_columns.append(question_column.strip())
            else:
                print("输入的列名不能为空，请重新输入。")

        if not question_columns:
            result_label.config(text="未输入有效的问题列名，分析取消。")
            return

        all_results = []
        all_score_percentages = []
        file_names = []
        for file_path in file_paths:
            # 打开 Excel 文件
            df = pd.read_excel(file_path)

            # 进行NPS分析
            nps_results = nps_analysis(df, question_columns)
            all_results.extend(nps_results)
            for result in nps_results:
                for key, value in result.items():
                    if "_各分数段占比情况" in key:
                        all_score_percentages.append(value)
            file_names.extend([os.path.basename(file_path)] * len([res for res in nps_results if res]))

        # 整理数据
        all_data = []
        for i, results in enumerate(all_results):
            if results:
                data = []
                for key, value in results.items():
                    if "_各分数段占比情况" not in key:
                        data.append([f"{file_names[i]}_{key}", value])
                all_data.extend(data)
        headers = ["指标", "数值"]
        df_result = pd.DataFrame(all_data, columns=headers)

        # 添加解释说明
        explanations = LANGUAGES[current_language]['explanation']
        interpretations = LANGUAGES[current_language]['interpretation']
        explanation_df = pd.DataFrame([explanations])
        explanation_df = explanation_df.reindex(
            columns=["推荐者", "被动者", "贬损者", "NPS净推荐值"])
        explanation_df.insert(0, "指标_解释说明",
                              "解释说明" if current_language == 'zh' else "Explanation")

        # 添加分析结果解读
        interpretation_df = pd.DataFrame([interpretations])
        interpretation_df = interpretation_df.reindex(
            columns=["推荐者", "被动者", "贬损者", "NPS净推荐值"])
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

            # 生成各类型占比情况柱状图
            categories = ["推荐者", "被动者", "贬损者"]
            percentages_list = []
            valid_results = [res for res in all_results if res]
            for result in valid_results:
                base_key = list(result.keys())[0].split('_')[0]
                keys = [f"{base_key}_推荐者比例", f"{base_key}_被动者比例", f"{base_key}_贬损者比例"]
                if all(key in result for key in keys):
                    percentages = [result[key] for key in keys]
                    percentages_list.extend(percentages)
                else:
                    print(f"结果字典中缺少必要的键: {keys}")

            num_files = len(file_paths)
            num_columns = len(question_columns)
            fig, ax = plt.subplots()
            x = np.arange(len(categories))
            width = 0.8 / (num_files * num_columns)
            for i, percentages in enumerate(percentages_list):
                if i < len(file_names) and i < len(valid_results):
                    ax.bar(x + i * width, percentages, width, label=f"{file_names[i]}_{list(valid_results[i].keys())[0].split('_')[0]}")
                else:
                    print(f"索引越界: i = {i}, file_names 长度 = {len(file_names)}, valid_results 长度 = {len(valid_results)}")

            ax.set_title('各类型占比情况' if current_language == 'zh' else 'Percentage of Each Type')
            ax.set_ylabel('比例 (%)' if current_language == 'zh' else 'Percentage (%)')
            ax.set_xticks(x + width * (num_files * num_columns - 1) / 2)
            ax.set_xticklabels(categories)
            ax.legend()

            # 保存图片
            img_path = os.path.splitext(save_path)[0] + '_types.png'
            plt.savefig(img_path)
            plt.close()

            # 生成各分数段占比情况柱状图
            fig, ax = plt.subplots()
            score_bins = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
            x = np.arange(len(score_bins) - 1)
            width = 0.8 / (num_files * num_columns)
            for i, score_percentages in enumerate(all_score_percentages):
                if i < len(file_names) and i < len(valid_results):
                    ax.bar(x + i * width, score_percentages.values, width, label=f"{file_names[i]}_{list(valid_results[i].keys())[0].split('_')[0]}")
                else:
                    print(f"索引越界: i = {i}, file_names 长度 = {len(file_names)}, valid_results 长度 = {len(valid_results)}")

            ax.set_title('各分数段占比情况' if current_language == 'zh' else 'Percentage of Each Score Range')
            ax.set_ylabel('比例 (%)' if current_language == 'zh' else 'Percentage (%)')
            ax.set_xticks(x + width * (num_files * num_columns - 1) / 2)
            ax.set_xticklabels([f"{score_bins[j]}-{score_bins[j + 1]}" for j in range(len(score_bins) - 1)])
            ax.legend()

            # 保存图片
            img_path = os.path.splitext(save_path)[0] + '_scores.png'
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