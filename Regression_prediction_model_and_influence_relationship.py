import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import pandas as pd
from tkinter import messagebox
import subprocess
import os

from Regression_Analysis_Selector import RegressionAnalysisSelector

# 移除了Canvas和Scrollbar相关的全局变量
button_frame = None


class RegressionPredictionApp:
    def __init__(self, root=None):
        # 定义语言字典
        self.LANGUAGES = {
            'zh': {
                'title': "(迪亚士) 回归预测模型与影响关系",
                'no_details': "无详细信息",
                'switch_language': "切换语言",
                'copy_success': "内容已复制到剪贴板",
                'open_script': "打开回归分析选择器"
            },
            'en': {
                'title': "(DIAS) Regression prediction model and influence relationship",
                'no_details': "No detailed information",
                'switch_language': "Switch Language",
                'copy_success': "Content has been copied to the clipboard",
                'open_script': "Open Regression Analysis Selector"
            }
        }

        # 当前语言
        self.current_language = 'en'

        # 如果没有提供root，则创建一个新窗口
        if root is None:
            self.root = ttk.Window(themename="flatly")
        else:
            self.root = root
        self.root.title(self.LANGUAGES[self.current_language]['title'])

        self.load_data()
        self.create_ui()

    def load_data(self):
        # 获取当前脚本所在的目录
        current_dir = os.path.dirname(os.path.abspath(__file__))

        # 构建 Excel 文件的相对路径
        excel_file_path = os.path.join(current_dir, 'Excel',
                                       'Regression prediction model and influence relationship.xlsx')

        try:
            # 读取文件，默认第一行作为表头
            df = pd.read_excel(excel_file_path)
            # 从第二行开始获取数据
            data = df[1:]
            self.first_column = data.iloc[:, 0].tolist()
            self.second_column = data.iloc[:, 1].tolist()
            self.third_column = data.iloc[:, 2].tolist()
            # 读取第四列数据
            self.fourth_column = data.iloc[:, 3].tolist()
        except Exception as e:
            print(f"读取 Excel 文件时出错: {e}")
            self.first_column = []
            self.second_column = []
            self.third_column = []
            self.fourth_column = []

        # 按首字母排序
        sorted_data = sorted(zip(self.first_column, self.second_column, self.third_column, self.fourth_column),
                             key=lambda x: str(x[0]).lower())
        if sorted_data:
            self.first_column, self.second_column, self.third_column, self.fourth_column = zip(*sorted_data)
            self.first_column = list(self.first_column)
            self.second_column = list(self.second_column)
            self.third_column = list(self.third_column)
            self.fourth_column = list(self.fourth_column)
        else:
            self.first_column = []
            self.second_column = []
            self.third_column = []
            self.fourth_column = []

    def show_details(self, index):
        # 根据当前语言选择显示第三列或第四列的内容
        if index < len(self.third_column):
            if self.current_language == 'zh':
                details = f"{self.third_column[index]}"
            else:
                details = f"{self.fourth_column[index]}"
        else:
            details = self.LANGUAGES[self.current_language]['no_details']
        self.details_text.delete(1.0, ttk.END)
        self.details_text.insert(ttk.END, details)

    def switch_language(self):
        self.current_language = 'zh' if self.current_language == 'en' else 'en'
        self.root.title(self.LANGUAGES[self.current_language]['title'])
        self.switch_language_label.config(text=self.LANGUAGES[self.current_language]['switch_language'])
        self.open_script_button.config(text=self.LANGUAGES[self.current_language]['open_script'])

        # 更新按钮文本
        for index, button in enumerate(self.button_list):
            if self.current_language == 'zh':
                display_text = self.first_column[index]
            else:
                display_text = self.second_column[index]
            button.config(text=display_text)

    def open_script(self):
        try:
            from Regression_Analysis_Selector import RegressionAnalysisSelector
            RegressionAnalysisSelector(ttk.Toplevel(self.root))
        except Exception as e:
            print(f"打开脚本时出错: {e}")
            messagebox.showerror("错误", f"无法打开回归分析选择器: {str(e)}")

    def create_ui(self):
        global button_frame

        # 获取屏幕的宽度和高度
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # 设置窗口的宽度和高度
        window_width = 1250
        window_height = 700

        # 计算窗口应该放置的位置
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        # 设置窗口的位置和大小
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # 创建一个主框架，用于居中内容
        main_frame = ttk.Frame(self.root)
        main_frame.pack(expand=True, fill=BOTH, padx=10, pady=10)

        # 创建一个框架来放置按钮，使用pack布局
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=X, padx=10, pady=10, expand=True, anchor='center')

        # 存储所有按钮的列表
        self.button_list = []

        # 创建按钮
        fixed_button_width = 30  # 设置固定的按钮宽度
        for index, button_text in enumerate(self.first_column):
            if self.current_language == 'zh':
                display_text = button_text
            else:
                display_text = self.second_column[index]
            button = ttk.Button(
                button_frame,
                text=display_text,
                bootstyle=PRIMARY,
                width=fixed_button_width
            )
            row = index // 4
            col = index % 4
            button.grid(row=row, column=col, padx=5, pady=5, sticky="ew")
            button.bind("<Button-1>", lambda event, i=index: self.show_details(i))
            self.button_list.append(button)

        # 创建详情显示文本框
        self.details_text = ttk.Text(self.root, height=20, wrap=ttk.WORD, font=12)
        self.details_text.pack(pady=5, fill=X, padx=10)

        # 创建一个新的框架来放置打开脚本按钮和切换语言标签
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(pady=5)

        # 创建打开脚本按钮
        self.open_script_button = ttk.Button(
            bottom_frame,
            text=self.LANGUAGES[self.current_language]['open_script'],
            command=self.open_script
        )
        self.open_script_button.pack(side=ttk.LEFT, padx=5)

        # 创建语言切换标签
        self.switch_language_label = ttk.Label(
            bottom_frame,
            text=self.LANGUAGES[self.current_language]['switch_language'],
            foreground='gray',
            cursor='hand2'
        )
        self.switch_language_label.pack(side=ttk.LEFT, padx=5)
        self.switch_language_label.bind("<Button-1>", lambda event: self.switch_language())

    def run(self):
        # 运行主循环
        self.root.mainloop()


# 为了向后兼容，保留原来的运行方式
def run_app():
    app = RegressionPredictionApp()
    app.run()


if __name__ == "__main__":
    run_app()