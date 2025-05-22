import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import pandas as pd
import os
from tkinter import messagebox

# 全局变量
canvas = None
button_frame = None

def on_mousewheel(event):
    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

def center_button_frame():
    # 更新 Canvas 的滚动区域
    button_frame.update_idletasks()
    canvas.config(scrollregion=canvas.bbox(ALL))

    # 计算按钮框架的宽度和高度
    button_frame_width = button_frame.winfo_width()
    button_frame_height = button_frame.winfo_height()
    canvas_width = canvas.winfo_width()
    canvas_height = canvas.winfo_height()

    # 计算水平和垂直偏移量以实现居中
    x_offset = (canvas_width - button_frame_width) // 2 if canvas_width > button_frame_width else 0
    y_offset = (canvas_height - button_frame_height) // 2 if canvas_height > button_frame_height else 0

    # 更新 Canvas 中窗口的位置
    canvas.coords(canvas.find_all()[0], (x_offset, y_offset))

class DatasetApp:
    def __init__(self, root=None):
        # 定义语言字典
        self.LANGUAGES = {
            'zh': {
                'title': "(迪亚士) 数据库",
                'no_details': "无详细信息",
                'switch_language': "切换语言",
                'copy_success': "内容已复制到剪贴板",
            },
            'en': {
                'title': "(DIAS) Dataset",
                'no_details': "No detailed information",
                'switch_language': "Switch Language",
                'copy_success': "Content has been copied to the clipboard",
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
        excel_file_path = os.path.join(current_dir, 'Excel', 'Dataset.xlsx')

        try:
            df = pd.read_excel(excel_file_path)
            # 从第二行开始获取数据
            data = df[1:]
            self.first_column = data.iloc[:, 0].tolist()
            self.second_column = data.iloc[:, 1].tolist()
        except Exception as e:
            print(f"读取 Excel 文件时出错: {e}")
            self.first_column = []
            self.second_column = []

        # 按首字母排序
        sorted_data = sorted(zip(self.first_column, self.second_column), key=lambda x: str(x[0]).lower())
        self.first_column, self.second_column = zip(*sorted_data) if sorted_data else ([], [])
        self.first_column = list(self.first_column)
        self.second_column = list(self.second_column)
        
    def switch_language(self):
        self.current_language = 'zh' if self.current_language == 'en' else 'en'
        self.root.title(self.LANGUAGES[self.current_language]['title'])
        self.switch_language_label.config(text=self.LANGUAGES[self.current_language]['switch_language'])

    def show_details(self, index):
        # 从第二列获取详细信息
        if index < len(self.second_column):
            details = f"{self.second_column[index]}"
        else:
            details = self.LANGUAGES[self.current_language]['no_details']
        self.details_text.delete(1.0, ttk.END)
        self.details_text.insert(ttk.END, details)
        
    def create_ui(self):
        global canvas, button_frame
        
        # 获取屏幕的宽度和高度
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # 设置窗口的宽度和高度
        window_width = 600
        window_height = 600

        # 计算窗口应该放置的位置
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        # 设置窗口的位置和大小
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # 创建一个主框架，用于居中内容
        main_frame = ttk.Frame(self.root)
        main_frame.pack(expand=True, fill=BOTH)

        # 创建一个 Canvas 组件
        canvas = ttk.Canvas(main_frame)
        canvas.pack(side=LEFT, fill=BOTH, expand=True)

        # 创建垂直滚动条
        scrollbar = ttk.Scrollbar(main_frame, command=canvas.yview)
        scrollbar.pack(side=RIGHT, fill=Y)

        # 配置 Canvas 的滚动条
        canvas.configure(yscrollcommand=scrollbar.set)

        # 创建一个框架来放置按钮
        button_frame = ttk.Frame(canvas)

        # 将按钮框架添加到 Canvas 中
        canvas.create_window((0, 0), window=button_frame, anchor=NW)

        # 存储所有按钮的列表
        self.button_list = []

        # 创建按钮
        for index, button_text in enumerate(self.first_column):
            button = ttk.Button(button_frame, text=button_text, bootstyle=PRIMARY)
            button.pack(fill=X, padx=5, pady=5)
            button.bind("<Button-1>", lambda event, i=index: self.show_details(i))
            self.button_list.append(button)

        # 初始居中按钮框架
        center_button_frame()

        # 绑定窗口大小改变事件，重新居中按钮框架
        self.root.bind("<Configure>", lambda event: center_button_frame())

        # 绑定鼠标滚轮事件
        canvas.bind_all("<MouseWheel>", on_mousewheel)

        # 创建详情显示文本框，将高度从 5 改为 3 以使其更窄
        self.details_text = ttk.Text(self.root, height=3, wrap=ttk.WORD, font=12)
        self.details_text.pack(pady=5, fill=X)

        # 创建语言切换标签，点击可切换语言，颜色设为灰色
        self.switch_language_label = ttk.Label(
            self.root, 
            text=self.LANGUAGES[self.current_language]['switch_language'], 
            foreground='gray', 
            cursor='hand2'
        )
        self.switch_language_label.pack(pady=5)
        self.switch_language_label.bind("<Button-1>", lambda event: self.switch_language())
    
    def run(self):
        # 运行主循环
        self.root.mainloop()

# 为了向后兼容，保留原来的run_app函数
def run_app():
    app = DatasetApp()
    app.run()

if __name__ == "__main__":
    run_app()