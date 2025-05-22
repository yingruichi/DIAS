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
from sklearn.discriminant_analysis import LinearDiscriminantAnalysis
from sklearn.metrics import classification_report, confusion_matrix

# 定义语言字典
LANGUAGES = {
    'zh': {
        'title': "判别分析",
        'select_button': "选择文件",
        'analyze_button': "分析文件",
        'file_not_found': "文件不存在，请重新选择。",
        'analysis_success': "分析完成，结果已保存到 {}\n",
        'no_save_path': "未选择保存路径，结果未保存。",
        'analysis_error': "分析文件时出错: {}",
        'switch_language': "切换语言",
        'explanation': {
            "判别分析": "根据已知类别的样本数据建立判别函数，对新样本进行分类。",
        },
        'interpretation': {
            "准确率": "分类正确的样本数占总样本数的比例。",
            "精确率": "预测为某一类别的样本中，实际为该类别的样本比例。",
            "召回率": "实际为某一类别的样本中，被正确预测为该类别的样本比例。",
            "F1值": "精确率和召回率的调和平均数，综合衡量分类性能。"
        }
    },
    'en': {
        'title': "Discriminant Analysis",
        'select_button': "Select File",
        'analyze_button': "Analyze File",
        'file_not_found': "The file does not exist. Please select again.",
        'analysis_success': "Analysis completed. The results have been saved to {}\n",
        'no_save_path': "No save path selected. The results were not saved.",
        'analysis_error': "An error occurred while analyzing the file: {}",
        'switch_language': "Switch Language",
        'explanation': {
            "Discriminant Analysis": "Establish a discriminant function based on sample data of known categories to classify new samples.",
        },
        'interpretation': {
            "Accuracy": "The proportion of correctly classified samples to the total number of samples.",
            "Precision": "The proportion of samples actually belonging to a certain category among the samples predicted to belong to that category.",
            "Recall": "The proportion of samples correctly predicted to belong to a certain category among the samples actually belonging to that category.",
            "F1-score": "The harmonic mean of precision and recall, comprehensively measuring the classification performance."
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

        # 假设最后一列是类别变量，其余列是特征变量
        X = df.iloc[:, :-1]
        y = df.iloc[:, -1]

        # 进行判别分析
        lda = LinearDiscriminantAnalysis()
        lda.fit(X, y)
        y_pred = lda.predict(X)

        # 生成分类报告
        report = classification_report(y, y_pred, output_dict=True)
        report_df = pd.DataFrame(report).transpose()

        # 生成混淆矩阵
        cm = confusion_matrix(y, y_pred)
        cm_df = pd.DataFrame(cm, index=lda.classes_, columns=lda.classes_)

        # 添加解释说明
        explanations = LANGUAGES[current_language]['explanation']
        interpretations = LANGUAGES[current_language]['interpretation']
        explanation_df = pd.DataFrame([explanations])
        explanation_df = explanation_df.reindex(
            columns=["判别分析" if current_language == 'zh' else "Discriminant Analysis"])
        explanation_df.insert(0, "统计量", "解释说明" if current_language == 'zh' else "Explanation")

        # 添加分析结果解读
        interpretation_df = pd.DataFrame([interpretations])
        interpretation_df = interpretation_df.reindex(columns=[
            "准确率" if current_language == 'zh' else "Accuracy",
            "精确率" if current_language == 'zh' else "Precision",
            "召回率" if current_language == 'zh' else "Recall",
            "F1值" if current_language == 'zh' else "F1-score"
        ])
        interpretation_df.insert(0, "统计量", "结果解读" if current_language == 'zh' else "Interpretation")

        # 合并数据、解释说明和结果解读
        combined_df = pd.concat([report_df, cm_df, explanation_df, interpretation_df], ignore_index=False)

        # 让用户选择保存路径
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            # 保存到 Excel 文件
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                combined_df.to_excel(writer, index=True)
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

            # 生成结果图片（混淆矩阵可视化）
            desktop_path = pathlib.Path.home() / 'Desktop'
            plot_path = desktop_path / 'discriminant_analysis_plot.png'
            plt.figure()
            plt.imshow(cm, interpolation='nearest', cmap=plt.cm.Blues)
            plt.title('判别分析混淆矩阵' if current_language == 'zh' else 'Confusion Matrix of Discriminant Analysis')
            plt.colorbar()
            tick_marks = np.arange(len(lda.classes_))
            plt.xticks(tick_marks, lda.classes_, rotation=45)
            plt.yticks(tick_marks, lda.classes_)

            thresh = cm.max() / 2.
            for i in range(cm.shape[0]):
                for j in range(cm.shape[1]):
                    plt.text(j, i, format(cm[i, j], 'd'),
                             horizontalalignment="center",
                             color="white" if cm[i, j] > thresh else "black")

            plt.tight_layout()
            plt.ylabel('真实类别' if current_language == 'zh' else 'True label')
            plt.xlabel('预测类别' if current_language == 'zh' else 'Predicted label')
            plt.savefig(plot_path)
            plt.close()

            result_msg = LANGUAGES[current_language]['analysis_success'].format(
                save_path) + f"结果图片已保存到 {plot_path}" if current_language == 'zh' else f"The result image has been saved to {plot_path}"
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
try:
    root.mainloop()
except KeyboardInterrupt:
    print("程序被手动中断。")
    root.destroy()  # 销毁主窗口
