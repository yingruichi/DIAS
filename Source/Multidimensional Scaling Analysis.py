import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox
import openpyxl
import os
import numpy as np
import pandas as pd
from tkinter import filedialog
import tkinter as tk
import tkinter.simpledialog
import matplotlib.pyplot as plt
import pathlib
from sklearn.manifold import MDS

# 设置 matplotlib 支持中文
plt.rcParams['font.family'] = 'SimHei'
plt.rcParams['axes.unicode_minus'] = False

# 定义语言字典
LANGUAGES = {
    'zh': {
        'title': "多维尺度分析",
        'select_button': "选择文件",
        'analyze_button': "分析文件",
        'file_not_found': "文件不存在，请重新选择。",
        'analysis_success': "分析完成，结果已保存到 {}\n",
        'no_save_path': "未选择保存路径，结果未保存。",
        'analysis_error': "分析文件时出错: {}",
        'switch_language': "切换语言",
        'explanation': {
            "多维尺度分析": "多维尺度分析（MDS）是一种将多维空间中的对象之间的相似性或距离信息可视化的技术。",
        },
        'interpretation': {
            "多维尺度分析": "在多维尺度分析图中，距离较近的点表示对象之间的相似性较高，距离较远的点表示对象之间的相似性较低。",
        }
    },
    'en': {
        'title': "Multidimensional Scaling Analysis",
        'select_button': "Select Files",
        'analyze_button': "Analyze Files",
        'file_not_found': "The file does not exist. Please select again.",
        'analysis_success': "Analysis completed. The results have been saved to {}\n",
        'no_save_path': "No save path selected. The results were not saved.",
        'analysis_error': "An error occurred while analyzing the file: {}",
        'switch_language': "Switch Language",
        'explanation': {
            "多维尺度分析": "Multidimensional Scaling (MDS) is a technique for visualizing the similarity or distance information between objects in a multi-dimensional space.",
        },
        'interpretation': {
            "多维尺度分析": "In the MDS plot, points that are closer together indicate higher similarity between objects, while points that are farther apart indicate lower similarity between objects.",
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
        file_entry.configure(style="TEntry")


def mds_analysis(data):
    try:
        # 进行多维尺度分析
        mds = MDS(n_components=2, random_state=42)
        mds_result = mds.fit_transform(data)
        return mds_result
    except Exception as e:
        print(f"多维尺度分析出错: {e}")
        return None


def analyze_file():
    global current_language
    file_paths = file_entry.get().split(", ")
    if not file_paths or file_paths[0] == "请输入待分析 Excel 文件的完整路径" or file_paths[
        0] == "Please enter the full path of the Excel file to be analyzed":
        file_paths = []
    for file_path in file_paths:
        if not os.path.exists(file_path):
            result_label.config(text=LANGUAGES[current_language]['file_not_found'])
            return
    try:
        all_results = []
        file_names = []
        for file_path in file_paths:
            # 打开 Excel 文件
            df = pd.read_excel(file_path)
            # 将特征名称转换为字符串类型
            df.columns = df.columns.astype(str)

            # 进行多维尺度分析
            mds_result = mds_analysis(df)
            if mds_result is not None:
                all_results.append(mds_result)
                file_names.append(os.path.basename(file_path))

        # 整理数据
        all_data = []
        for i, result in enumerate(all_results):
            for j, point in enumerate(result):
                all_data.append([f"{file_names[i]}_对象{j + 1}", point[0], point[1]])
        headers = ["对象", "维度1", "维度2"]
        df_result = pd.DataFrame(all_data, columns=headers)

        # 添加解释说明
        explanations = LANGUAGES[current_language]['explanation']
        interpretations = LANGUAGES[current_language]['interpretation']
        explanation_df = pd.DataFrame([explanations])
        explanation_df = explanation_df.reindex(
            columns=["多维尺度分析"])
        explanation_df.insert(0, "指标_解释说明",
                              "解释说明" if current_language == 'zh' else "Explanation")

        # 添加分析结果解读
        interpretation_df = pd.DataFrame([interpretations])
        interpretation_df = interpretation_df.reindex(
            columns=["多维尺度分析"])
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

            # 生成多维尺度分析图
            fig, ax = plt.subplots()
            has_labels = False
            for i, result in enumerate(all_results):
                if len(result) > 0:
                    ax.scatter(result[:, 0], result[:, 1], label=file_names[i])
                    has_labels = True
                    for j, point in enumerate(result):
                        ax.annotate(f"对象{j + 1}", (point[0], point[1]))
            if has_labels:
                ax.set_title('多维尺度分析图' if current_language == 'zh' else 'Multidimensional Scaling Plot')
                ax.set_xlabel('维度1' if current_language == 'zh' else 'Dimension 1')
                ax.set_ylabel('维度2' if current_language == 'zh' else 'Dimension 2')
                ax.legend()

            # 保存图片
            img_path = os.path.splitext(save_path)[0] + '_mds.png'
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
