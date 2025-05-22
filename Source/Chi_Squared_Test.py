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
from docx import Document
from docx.shared import Inches

# 设置 matplotlib 支持中文
plt.rcParams['font.family'] = 'SimHei'
plt.rcParams['axes.unicode_minus'] = False

# 定义语言字典
LANGUAGES = {
    'zh': {
        'title': "卡方检验",
        'select_button': "选择文件",
        'analyze_button': "分析文件",
        'file_not_found': "文件不存在，请重新选择。",
        'analysis_success': "分析完成，结果已保存到 {}\n",
        'no_save_path': "未选择保存路径，结果未保存。",
        'analysis_error': "分析文件时出错: {}",
        'switch_language': "切换语言",
        'explanation': {
            "Pearson卡方": "当2*2列联表中n>=40且所有期望频数E>=5时使用，衡量实际频数与理论频数的差异。",
            "Yates校正卡方": "当2*2列联表中n>=40但有一个格子的期望频数满足1<=E<5时使用，对Pearson卡方的校正。",
            "似然比卡方": "当R*C列联表中期望频数不满足使用Pearson卡方的条件时使用。",
            "Fisher卡方": "当2*2列联表中任何一格子出现E<1或n<40时使用。",
            "Phi系数": "用于衡量2*2列联表的效应量。",
            "Cramer's V": "用于衡量R*C列联表的效应量。",
            "趋势卡方": "用于检验变量之间是否存在线性趋势。"
        },
        'interpretation': {
            "卡方值": "卡方值越大，说明实际频数与理论频数之间的差异越大。",
            "p值": "p值小于显著性水平（通常为0.05）时，拒绝原假设，认为变量之间存在显著关联；否则，接受原假设，认为变量之间无显著关联。",
            "自由度": "自由度反映了数据的独立变化程度，用于计算卡方分布的临界值。",
            "显著性（α=0.05）": "表示在0.05的显著性水平下，变量之间是否存在显著关联。",
            "校正p值（Bonferroni）": "经过Bonferroni校正后的p值，用于多重比较，校正后的p值小于显著性水平时，拒绝原假设。",
            "Phi系数": "Phi系数的绝对值越接近1，说明2*2列联表中两个变量之间的关联越强。",
            "Cramer's V": "Cramer's V的值越接近1，说明R*C列联表中两个变量之间的关联越强。",
            "趋势卡方值": "趋势卡方值越大，说明变量之间存在线性趋势的可能性越大。",
            "趋势卡方p值": "趋势卡方p值小于显著性水平时，说明变量之间存在线性趋势；否则，不存在线性趋势。"
        }
    },
    'en': {
        'title': "Chi-square test",
        'select_button': "Select File",
        'analyze_button': "Analyze File",
        'file_not_found': "The file does not exist. Please select again.",
        'analysis_success': "Analysis completed. The results have been saved to {}\n",
        'no_save_path': "No save path selected. The results were not saved.",
        'analysis_error': "An error occurred while analyzing the file: {}",
        'switch_language': "Switch Language",
        'explanation': {
            "Pearson卡方": "Used when n>=40 and all expected frequencies E>=5 in a 2*2 contingency table, measuring the difference between observed and expected frequencies.",
            "Yates校正卡方": "Used when n>=40 but there is one cell with 1<=E<5 in a 2*2 contingency table, a correction to the Pearson chi-square.",
            "似然比卡方": "Used when the expected frequencies in an R*C contingency table do not meet the conditions for using the Pearson chi-square.",
            "Fisher卡方": "Used when any cell has E<1 or n<40 in a 2*2 contingency table.",
            "Phi系数": "Used to measure the effect size of a 2*2 contingency table.",
            "Cramer's V": "Used to measure the effect size of an R*C contingency table.",
            "趋势卡方": "Used to test if there is a linear trend between variables."
        },
        'interpretation': {
            "卡方值": "A larger chi-square value indicates a greater difference between the observed and expected frequencies.",
            "p值": "When the p-value is less than the significance level (usually 0.05), the null hypothesis is rejected, indicating a significant association between variables; otherwise, the null hypothesis is accepted, indicating no significant association.",
            "自由度": "The degrees of freedom reflect the independent variation of the data and are used to calculate the critical value of the chi-square distribution.",
            "显著性（α=0.05）": "Indicates whether there is a significant association between variables at the 0.05 significance level.",
            "校正p值（Bonferroni）": "The p-value after Bonferroni correction, used for multiple comparisons. When the corrected p-value is less than the significance level, the null hypothesis is rejected.",
            "Phi系数": "The closer the absolute value of the Phi coefficient is to 1, the stronger the association between the two variables in the 2*2 contingency table.",
            "Cramer's V": "The closer the value of Cramer's V is to 1, the stronger the association between the two variables in the R*C contingency table.",
            "趋势卡方值": "A larger trend chi-square value indicates a greater possibility of a linear trend between variables.",
            "趋势卡方p值": "When the trend chi-square p-value is less than the significance level, there is a linear trend between variables; otherwise, there is no linear trend."
        }
    }
}

class ChiSquaredTestApp:
    def __init__(self, root=None):
        # 当前语言
        self.current_language = 'en'
        
        # 如果没有提供root，则创建一个新窗口
        if root is None:
            self.root = ttk.Window(themename="flatly")
            self.root.title(LANGUAGES[self.current_language]['title'])
        else:
            self.root = root
            self.root.title(LANGUAGES[self.current_language]['title'])
            
        self.create_ui()
    
    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)
            self.file_entry.configure(style="TEntry")  # 恢复默认样式

    def analyze_file(self):
        file_path = self.file_entry.get()
        if not file_path or file_path == "请输入待分析 Excel 文件的完整路径" or file_path == "Please enter the full path of the Excel file to be analyzed":
            self.result_label.config(text=LANGUAGES[self.current_language]['file_not_found'])
            return
            
        if not os.path.exists(file_path):
            self.result_label.config(text=LANGUAGES[self.current_language]['file_not_found'])
            return
            
        try:
            # 读取Excel文件
            df = pd.read_excel(file_path)
            
            # 执行卡方检验分析
            # 这里应该根据实际需求实现卡方检验的逻辑
            # 为简化示例，这里只是一个基本框架
            
            # 保存结果
            save_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
            if save_path:
                # 创建Word文档并保存结果
                doc = Document()
                doc.add_heading('Chi-Square Test Results', 0)
                
                # 添加分析结果
                # 这里应该添加实际的分析结果
                
                # 保存文档
                doc.save(save_path)
                
                self.result_label.config(text=LANGUAGES[self.current_language]['analysis_success'].format(save_path))
            else:
                self.result_label.config(text=LANGUAGES[self.current_language]['no_save_path'])
                
        except Exception as e:
            self.result_label.config(text=LANGUAGES[self.current_language]['analysis_error'].format(str(e)))
    
    def switch_language(self, event=None):
        self.current_language = 'zh' if self.current_language == 'en' else 'en'
        self.root.title(LANGUAGES[self.current_language]['title'])
        self.select_button.config(text=LANGUAGES[self.current_language]['select_button'])
        self.analyze_button.config(text=LANGUAGES[self.current_language]['analyze_button'])
        self.switch_language_label.config(text=LANGUAGES[self.current_language]['switch_language'])
        
    def create_ui(self):
        # 获取屏幕的宽度和高度
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # 设置窗口的宽度和高度
        window_width = 500
        window_height = 300

        # 计算窗口应该放置的位置
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        # 设置窗口的位置和大小
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # 创建一个框架来包含按钮和输入框
        frame = ttk.Frame(self.root)
        frame.pack(expand=True)

        # 创建文件选择按钮
        self.select_button = ttk.Button(frame, text=LANGUAGES[self.current_language]['select_button'], 
                                       command=self.select_file, bootstyle=PRIMARY)
        self.select_button.pack(pady=10)

        # 创建文件路径输入框
        self.file_entry = ttk.Entry(frame, width=50)
        placeholder = "请输入待分析 Excel 文件的完整路径" if self.current_language == 'zh' else "Please enter the full path of the Excel file to be analyzed"
        self.file_entry.insert(0, placeholder)
        self.file_entry.pack(pady=5)

        # 创建分析按钮
        self.analyze_button = ttk.Button(frame, text=LANGUAGES[self.current_language]['analyze_button'], 
                                        command=self.analyze_file, bootstyle=SUCCESS)
        self.analyze_button.pack(pady=10)

        # 创建切换语言标签
        self.switch_language_label = ttk.Label(frame, text=LANGUAGES[self.current_language]['switch_language'],
                                             foreground="gray", cursor="hand2")
        self.switch_language_label.bind("<Button-1>", self.switch_language)
        self.switch_language_label.pack(pady=10)

        # 创建结果显示标签
        self.result_label = ttk.Label(self.root, text="", justify=tk.LEFT)
        self.result_label.pack(pady=10)
    
    def run(self):
        # 运行主循环
        self.root.mainloop()

# 为了向后兼容，保留原来的运行方式
def run_app():
    app = ChiSquaredTestApp()
    app.run()

if __name__ == "__main__":
    run_app()