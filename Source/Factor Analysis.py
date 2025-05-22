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
from factor_analyzer import FactorAnalyzer
from factor_analyzer.factor_analyzer import calculate_bartlett_sphericity
from factor_analyzer.factor_analyzer import calculate_kmo

# 设置 matplotlib 支持中文
plt.rcParams['font.family'] = 'SimHei'
plt.rcParams['axes.unicode_minus'] = False

# 定义语言字典
LANGUAGES = {
    'zh': {
        'title': "因子分析",
        'select_button': "选择文件",
        'analyze_button': "分析文件",
        'file_not_found': "文件不存在，请重新选择。",
        'analysis_success': "分析完成，结果已保存到 {}\n",
        'no_save_path': "未选择保存路径，结果未保存。",
        'analysis_error': "分析文件时出错: {}",
        'switch_language': "切换语言",
        'explanation': {
            "因子载荷矩阵": "显示每个变量在各个因子上的载荷，反映变量与因子的相关性",
            "共同度": "表示每个变量被因子所解释的方差比例",
            "特征值和方差贡献率": "特征值表示每个因子解释的总方差，方差贡献率表示每个因子解释的方差占总方差的比例",
            "Bartlett球形检验": "检验变量之间是否存在相关性",
            "KMO检验": "衡量变量之间的偏相关性，判断数据是否适合进行因子分析",
            "碎石图": "展示特征值随因子数量的变化情况，帮助确定因子的数量"
        },
        'interpretation': {
            "因子载荷矩阵": "绝对值越大，说明变量与因子的相关性越强",
            "共同度": "值越接近1，说明变量被因子解释的程度越高",
            "特征值和方差贡献率": "特征值大于1的因子通常被保留，方差贡献率越高，说明该因子越重要",
            "Bartlett球形检验": "p值小于0.05时，拒绝原假设，表明变量之间存在相关性，适合进行因子分析",
            "KMO检验": "KMO值大于0.6时，适合进行因子分析",
            "碎石图": "曲线的拐点处通常表示合适的因子数量"
        }
    },
    'en': {
        'title': "Factor Analysis",
        'select_button': "Select File",
        'analyze_button': "Analyze File",
        'file_not_found': "The file does not exist. Please select again.",
        'analysis_success': "Analysis completed. The results have been saved to {}\n",
        'no_save_path': "No save path selected. The results were not saved.",
        'analysis_error': "An error occurred while analyzing the file: {}",
        'switch_language': "Switch Language",
        'explanation': {
            "因子载荷矩阵": "Shows the loadings of each variable on each factor, reflecting the correlation between variables and factors",
            "共同度": "Represents the proportion of variance of each variable explained by the factors",
            "特征值和方差贡献率": "The eigenvalue represents the total variance explained by each factor, and the variance contribution rate represents the proportion of variance explained by each factor to the total variance",
            "Bartlett球形检验": "Tests whether there is a correlation between variables",
            "KMO检验": "Measures the partial correlation between variables to determine whether the data is suitable for factor analysis",
            "碎石图": "Shows the change of eigenvalues with the number of factors, helping to determine the number of factors"
        },
        'interpretation': {
            "因子载荷矩阵": "The larger the absolute value, the stronger the correlation between the variable and the factor",
            "共同度": "The closer the value is to 1, the higher the degree to which the variable is explained by the factors",
            "特征值和方差贡献率": "Factors with eigenvalues greater than 1 are usually retained. The higher the variance contribution rate, the more important the factor",
            "Bartlett球形检验": "When the p-value is less than 0.05, the null hypothesis is rejected, indicating that there is a correlation between variables and factor analysis is suitable",
            "KMO检验": "When the KMO value is greater than 0.6, factor analysis is suitable",
            "碎石图": "The inflection point of the curve usually indicates the appropriate number of factors"
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


def factor_analysis(data):
    """
    进行因子分析
    :param data: 输入数据
    :return: 因子载荷矩阵、共同度、特征值、方差贡献率、Bartlett球形检验结果、KMO检验结果
    """
    # Bartlett球形检验
    chi_square_value, p_value = calculate_bartlett_sphericity(data)

    # KMO检验
    kmo_all, kmo_model = calculate_kmo(data)

    # 创建因子分析对象
    fa = FactorAnalyzer()
    fa.fit(data)

    # 计算特征值和方差贡献率
    ev, v = fa.get_eigenvalues()

    # 确定因子数量
    num_factors = sum(ev > 1)

    # 重新进行因子分析
    fa = FactorAnalyzer(n_factors=num_factors, rotation='varimax')
    fa.fit(data)

    # 获取因子载荷矩阵
    loadings = fa.loadings_

    # 获取共同度
    communalities = fa.get_communalities()

    return loadings, communalities, ev, v, (chi_square_value, p_value), kmo_model


def plot_scree_plot(ev, save_path):
    """
    绘制碎石图
    :param ev: 特征值
    :param save_path: 图片保存路径
    """
    plt.figure(figsize=(10, 5))
    plt.plot(range(1, len(ev) + 1), ev, marker='o')
    plt.title('碎石图' if current_language == 'zh' else 'Scree Plot')
    plt.xlabel('因子数量' if current_language == 'zh' else 'Number of Factors')
    plt.ylabel('特征值' if current_language == 'zh' else 'Eigenvalues')
    img_path = os.path.splitext(save_path)[0] + '_scree_plot.png'
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
        df = pd.read_excel(file_path)
        original_data = df.values

        # 进行因子分析
        loadings, communalities, ev, v, bartlett_result, kmo_result = factor_analysis(df)

        # 整理数据
        factor_names = [f'因子{i + 1}' for i in range(len(loadings[0]))]
        loadings_df = pd.DataFrame(loadings, index=df.columns, columns=factor_names)
        communalities_df = pd.DataFrame(communalities, index=df.columns, columns=['共同度'])
        ev_df = pd.DataFrame(ev, columns=['特征值'])
        v_df = pd.DataFrame(v, columns=['方差贡献率'])
        bartlett_df = pd.DataFrame([bartlett_result], columns=['卡方值', 'p值'], index=['Bartlett球形检验'])
        kmo_df = pd.DataFrame([kmo_result], columns=['KMO值'], index=['KMO检验'])

        # 添加解释说明
        explanations = LANGUAGES[current_language]['explanation']
        interpretations = LANGUAGES[current_language]['interpretation']
        explanation_df = pd.DataFrame([explanations])
        explanation_df = explanation_df.reindex(columns=["因子载荷矩阵", "共同度", "特征值和方差贡献率", "Bartlett球形检验", "KMO检验", "碎石图"])
        explanation_df.insert(0, "统计量_解释说明", "解释说明" if current_language == 'zh' else "Explanation")

        # 添加分析结果解读
        interpretation_df = pd.DataFrame([interpretations])
        interpretation_df = interpretation_df.reindex(columns=["因子载荷矩阵", "共同度", "特征值和方差贡献率", "Bartlett球形检验", "KMO检验", "碎石图"])
        interpretation_df.insert(0, "统计量_结果解读", "结果解读" if current_language == 'zh' else "Interpretation")

        # 合并数据、解释说明和结果解读
        with pd.ExcelWriter('factor_analysis_results.xlsx', engine='openpyxl') as writer:
            loadings_df.to_excel(writer, sheet_name='因子载荷矩阵')
            communalities_df.to_excel(writer, sheet_name='共同度')
            pd.concat([ev_df, v_df], axis=1).to_excel(writer, sheet_name='特征值和方差贡献率')
            bartlett_df.to_excel(writer, sheet_name='Bartlett球形检验')
            kmo_df.to_excel(writer, sheet_name='KMO检验')
            explanation_df.to_excel(writer, sheet_name='解释说明', index=False)
            interpretation_df.to_excel(writer, sheet_name='结果解读', index=False)

            # 自动调整列宽
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
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

        # 让用户选择保存路径
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            # 复制文件到用户指定路径
            import shutil
            shutil.copyfile('factor_analysis_results.xlsx', save_path)

            result_msg = LANGUAGES[current_language]['analysis_success'].format(save_path)
            result_label.config(text=result_msg, wraplength=400)

            # 生成碎石图
            plot_scree_plot(ev, save_path)

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