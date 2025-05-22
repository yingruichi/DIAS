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
from sklearn.cluster import DBSCAN
from sklearn.metrics import silhouette_score

# 设置 matplotlib 支持中文
plt.rcParams['font.family'] = 'SimHei'
plt.rcParams['axes.unicode_minus'] = False

# 定义语言字典
LANGUAGES = {
    'zh': {
        'title': "密度聚类分析",
        'select_button': "选择文件",
        'analyze_button': "分析文件",
        'file_not_found': "文件不存在，请重新选择。",
        'analysis_success': "分析完成，结果已保存到 {}\n",
        'no_save_path': "未选择保存路径，结果未保存。",
        'analysis_error': "分析文件时出错: {}",
        'switch_language': "切换语言",
        'explanation': {
            "聚类结果": "每个样本所属的聚类类别",
            "聚类散点图": "展示样本在二维空间中的分布及聚类结果",
            "轮廓系数": "衡量聚类效果的指标，值越接近1表示聚类效果越好"
        },
        'interpretation': {
            "聚类结果": "可用于区分不同样本所属的类别，-1 表示噪声点",
            "聚类散点图": "直观展示样本之间的聚类关系和分布情况",
            "轮廓系数": "若值接近1，说明聚类紧凑且分离度高；若值接近 -1，说明聚类效果差"
        }
    },
    'en': {
        'title': "Density-Based Clustering Analysis",
        'select_button': "Select File",
        'analyze_button': "Analyze File",
        'file_not_found': "The file does not exist. Please select again.",
        'analysis_success': "Analysis completed. The results have been saved to {}\n",
        'no_save_path': "No save path selected. The results were not saved.",
        'analysis_error': "An error occurred while analyzing the file: {}",
        'switch_language': "Switch Language",
        'explanation': {
            "聚类结果": "The cluster label to which each sample belongs",
            "聚类散点图": "Show the distribution of samples in a two-dimensional space and the clustering results",
            "轮廓系数": "An index to measure the clustering effect. A value closer to 1 indicates better clustering."
        },
        'interpretation': {
            "聚类结果": "Can be used to distinguish the categories to which different samples belong. -1 represents noise points.",
            "聚类散点图": "Visually show the clustering relationship and distribution of samples.",
            "轮廓系数": "If the value is close to 1, the clusters are compact and well-separated. If close to -1, the clustering effect is poor."
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


def density_clustering(data, eps=0.5, min_samples=5):
    """
    进行密度聚类分析
    :param data: 输入数据
    :param eps: 邻域的最大距离
    :param min_samples: 形成核心点所需的最小样本数
    :return: 聚类标签
    """
    model = DBSCAN(eps=eps, min_samples=min_samples)
    labels = model.fit_predict(data)
    return labels, model


def plot_clustering(data, labels, save_path):
    """
    绘制聚类散点图
    :param data: 输入数据
    :param labels: 聚类标签
    :param save_path: 图片保存路径
    """
    plt.figure(figsize=(10, 5))
    unique_labels = set(labels)
    colors = [plt.cm.Spectral(each) for each in np.linspace(0, 1, len(unique_labels))]
    for k, col in zip(unique_labels, colors):
        if k == -1:
            # 噪声点用黑色表示
            col = [0, 0, 0, 1]
        class_member_mask = (labels == k)
        xy = data[class_member_mask]
        plt.plot(xy[:, 0], xy[:, 1], 'o', markerfacecolor=tuple(col),
                 markeredgecolor='k', markersize=6)
    plt.title('密度聚类散点图' if current_language == 'zh' else 'Density-Based Clustering Scatter Plot')
    plt.xlabel('特征1' if current_language == 'zh' else 'Feature 1')
    plt.ylabel('特征2' if current_language == 'zh' else 'Feature 2')
    img_path = os.path.splitext(save_path)[0] + '_density_clustering_scatter.png'
    plt.savefig(img_path)
    plt.close()


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
        original_data = df.values  # 使用新变量存储原始数据

        # 进行密度聚类分析
        labels, model = density_clustering(original_data)

        # 计算轮廓系数
        try:
            silhouette_avg = silhouette_score(original_data, labels)
        except ValueError:
            silhouette_avg = "无法计算（可能只有一个聚类或所有样本为噪声点）"

        # 整理数据
        result_data = [
            ["聚类结果", labels.tolist(), ""],
            ["轮廓系数", silhouette_avg, ""]
        ]
        headers = ["统计量", "统计量值", "p值"]
        df = pd.DataFrame(result_data, columns=headers)

        # 添加解释说明
        explanations = LANGUAGES[current_language]['explanation']
        interpretations = LANGUAGES[current_language]['interpretation']
        explanation_df = pd.DataFrame([explanations])
        explanation_df = explanation_df.reindex(columns=["聚类结果", "聚类散点图", "轮廓系数"])
        explanation_df.insert(0, "统计量_解释说明", "解释说明" if current_language == 'zh' else "Explanation")

        # 添加分析结果解读
        interpretation_df = pd.DataFrame([interpretations])
        interpretation_df = interpretation_df.reindex(columns=["聚类结果", "聚类散点图", "轮廓系数"])
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

            # 生成聚类散点图
            if original_data.shape[1] >= 2:
                plot_clustering(original_data, labels, save_path)
            else:
                result_label.config(text=result_msg + "数据维度小于2，无法绘制散点图。", wraplength=400)

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
