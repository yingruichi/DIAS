import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import pandas as pd
import os

# 全局变量
button_frame = None


def center_button_frame():
    # 更新按钮框架的大小
    button_frame.update_idletasks()

    # 计算按钮框架的宽度和高度
    button_frame_width = button_frame.winfo_width()
    button_frame_height = button_frame.winfo_height()
    main_frame_width = main_frame.winfo_width()
    main_frame_height = main_frame.winfo_height()

    # 计算水平和垂直偏移量以实现居中
    x_offset = (main_frame_width - button_frame_width) // 2 if main_frame_width > button_frame_width else 0
    y_offset = (main_frame_height - button_frame_height) // 2 if main_frame_height > button_frame_height else 0

    # 更新按钮框架的位置
    button_frame.place(x=x_offset, y=y_offset)


class CorrelationAnalysisApp:
    def __init__(self, root=None):
        # 定义语言字典
        self.LANGUAGES = {
            'zh': {
                'title': "(迪亚士) 相关性分析",
                'no_details': "无详细信息",
                'switch_language': "切换语言",
                'copy_success': "内容已复制到剪贴板",
            },
            'en': {
                'title': "(DIAS) Correlation analysis",
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
        excel_folder = os.path.join(current_dir, 'Excel')
        excel_file_path = os.path.join(excel_folder, 'Correlation analysis.xlsx')

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

        # 更新按钮文本和宽度
        for index, button in enumerate(self.button_list):
            if self.current_language == 'zh':
                display_text = self.first_column[index]
            else:
                display_text = self.second_column[index]
            button.config(text=display_text, width=self.max_button_width)

    def create_ui(self):
        global button_frame

        # 获取屏幕的宽度和高度
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # 设置窗口的宽度和高度
        window_width = 1400
        window_height = 600

        # 计算窗口应该放置的位置
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        # 设置窗口的位置和大小
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # 创建一个主框架，用于居中内容
        global main_frame
        main_frame = ttk.Frame(self.root)
        main_frame.pack(expand=True, fill=BOTH, anchor='n')  # 使用 anchor='n' 让框架在顶部居中

        # 创建一个框架来放置按钮
        button_frame = ttk.Frame(main_frame)

        # 存储所有按钮的列表
        self.button_list = []

        # 记录英语模式下按钮所需的最大宽度
        self.max_button_width = 0

        # 计算中文和英文文本的最大宽度
        max_chinese_width = max(len(str(text)) for text in self.first_column) if self.first_column else 0
        max_english_width = max(len(str(text)) for text in self.second_column) if self.second_column else 0

        # 调整系数 - 根据实际情况调整
        chinese_adjustment = 1.2  # 中文字符通常更宽
        padding = 3  # 额外的边距

        # 计算最终的最大宽度
        adjusted_chinese_width = int(max_chinese_width * chinese_adjustment)
        self.max_button_width = max(adjusted_chinese_width, max_english_width) + padding

        # 创建按钮
        for index, button_text in enumerate(self.first_column):
            if self.current_language == 'zh':
                display_text = button_text
            else:
                display_text = self.second_column[index]
            button = ttk.Button(button_frame, text=display_text, bootstyle=PRIMARY, width=self.max_button_width)
            row = index // 4
            col = index % 4
            button.grid(row=row, column=col, padx=5, pady=5, sticky="ew")
            button.bind("<Button-1>", lambda event, i=index: self.show_details(i))
            self.button_list.append(button)

        # 初始居中按钮框架
        center_button_frame()

        # 绑定窗口大小改变事件，重新居中按钮框架
        self.root.bind("<Configure>", lambda event: center_button_frame())

        # 创建详情显示文本框，增加高度
        self.details_text = ttk.Text(self.root, height=20, wrap=ttk.WORD, font=12)
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
    app = CorrelationAnalysisApp()
    app.run()


if __name__ == "__main__":
    run_app()