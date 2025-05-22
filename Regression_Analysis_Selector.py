import tkinter as tk
from tkinter import messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import subprocess
import os

from Source.Analyzer import AnalyzerApp

# 定义不同语言的文本信息
LANGUAGES = {
    "中文": {
        "title": "回归分析选择器",
        "num_features_label": "特征数量:",
        "sample_size_label": "样本量:",
        "num_dependent_values_label": "因变量取值数量:\n（仅在因变量为分类变量时使用）",
        "relationship_label": "因变量和自变量的关系类型:",
        "relationship_options": ["线性", "非线性"],
        "dependent_variable_label": "因变量类型:",
        "dependent_variable_options": ["连续", "分类"],
        "multicollinearity_label": "数据是否存在多重共线性:",
        "multicollinearity_options": ["是", "否"],
        "independent_variable_label": "自变量类型:",
        "independent_variable_options": ["连续", "分类"],
        "regularization_label": "是否需要正则化:",
        "regularization_options": ["是", "否"],
        "truncation_label": "数据是否有截断情况:",
        "truncation_options": ["是", "否"],
        "outliers_label": "数据是否有异常值:",
        "outliers_options": ["是", "否"],
        "analyze_button": "分析并建议",
        "suggestion_title": "回归分析建议",
        "no_suggestion": "根据输入信息，暂时无法给出合适的建议。",
        "input_error": "请输入有效的数字！",
        "change_language_text": "切换语言",
        "open_other_script_button": "打开分析器",
        "regression_methods": {
            "Ridge Regression": "岭回归",
            "Lasso Regression": "套索回归",
            "PLS Regression": "偏最小二乘回归PLS",
            "Linear Regression OLS": "普通最小二乘线性回归OLS",
            "Linear Regression Robust": "稳健线性回归",
            "Regularized Linear Regression (e.g., Ridge, Lasso)": "正则化线性回归（如岭回归、套索回归）",
            "Binary Logistic Regression": "二元逻辑回归",
            "Regularized Binary Logistic Regression": "正则化二元逻辑回归",
            "Multinomial Logistic Regression": "多项逻辑回归",
            "Regularized Multinomial Logistic Regression": "正则化多项逻辑回归",
            "Polynomial Regression": "多项式回归",
            "Regularized Polynomial Regression": "正则化多项式回归",
            "Linear Tobit Regression": "线性托宾回归",
            "Hierarchical Regression": "分层回归",
            "Stepwise Regression": "逐步回归",
            "Binary Logit Regression": "二元Logit回归",
            "Multinomial Logit Regression": "多分类Logit回归",
            "Ordered Logit Regression": "有序Logit回归",
            "Logistic Regression": "逻辑回归",
            "Linear Regression Lasso": "线性回归Lasso",
            "Linear Regression PLS": "线性回归PLS",
        }
    },
    "English": {
        "title": "Regression Analysis Selector",
        "num_features_label": "Number of features:",
        "sample_size_label": "Sample size:",
        "num_dependent_values_label": "Number of values of the dependent variable:\n(only used when the dependent variable is a categorical variable)",
        "relationship_label": "Relationship type between the dependent and independent variables:",
        "relationship_options": ["Linear", "Nonlinear"],
        "dependent_variable_label": "Type of dependent variable:",
        "dependent_variable_options": ["Continuous", "Categorical"],
        "multicollinearity_label": "Does the data have multicollinearity?",
        "multicollinearity_options": ["Yes", "No"],
        "independent_variable_label": "Type of independent variable:",
        "independent_variable_options": ["Continuous", "Categorical"],
        "regularization_label": "Is regularization required?",
        "regularization_options": ["Yes", "No"],
        "truncation_label": "Does the data have truncation?",
        "truncation_options": ["Yes", "No"],
        "outliers_label": "Does the data have outliers?",
        "outliers_options": ["Yes", "No"],
        "analyze_button": "Analyze and Suggest",
        "suggestion_title": "Regression Analysis Suggestion",
        "no_suggestion": "Based on the input information, no suitable suggestion can be given at present.",
        "input_error": "Please enter valid numbers!",
        "change_language_text": "Change Language",
        "open_other_script_button": "Open Analyzer"
    }
}


class RegressionAnalysisSelector:
    def __init__(self, root=None):
        # 修改初始语言为英文
        self.current_language = 'English'

        # 如果没有提供root，则创建一个新窗口
        if root is None:
            self.root = ttk.Window(themename="flatly")
        else:
            self.root = root
        self.root.title(LANGUAGES[self.current_language]["title"])

        self.create_ui()

    def is_sample_large(self, sample_size, num_features):
        return sample_size > 10 * num_features

    def suggest_linear_continuous(self, multicollinearity, independent_variable_type, regularization, sample_size,
                                  num_features, truncation, outliers):
        suggestions = []
        if multicollinearity == LANGUAGES[self.current_language]["multicollinearity_options"][0]:
            if regularization == LANGUAGES[self.current_language]["regularization_options"][0]:
                suggestions.extend(
                    ["Ridge Regression", "Lasso Regression", "Linear Regression Lasso", "Linear Regression PLS"])
            else:
                suggestions.extend(["PLS Regression"])
        else:
            if self.is_sample_large(sample_size, num_features):
                if regularization == LANGUAGES[self.current_language]["regularization_options"][0]:
                    suggestions.extend(["Regularized Linear Regression (e.g., Ridge, Lasso)"])
                else:
                    suggestions.extend(["Linear Regression OLS", "Linear Regression Robust", "Hierarchical Regression",
                                        "Stepwise Regression"])
            else:
                if regularization == LANGUAGES[self.current_language]["regularization_options"][0]:
                    suggestions.extend(["Regularized Linear Regression (e.g., Ridge, Lasso)", "Linear Regression Lasso",
                                        "Linear Regression PLS"])
                else:
                    suggestions.extend(["PLS Regression"])

            if independent_variable_type == LANGUAGES[self.current_language]["independent_variable_options"][1]:
                suggestions = [s + " (need encoding)" for s in suggestions]

        if truncation == LANGUAGES[self.current_language]["truncation_options"][0]:
            suggestions.append("Linear Tobit Regression")
        if outliers == LANGUAGES[self.current_language]["outliers_options"][0]:
            suggestions.append("Linear Regression Robust")

        return suggestions

    def suggest_linear_categorical(self, independent_variable_type, regularization, sample_size, num_features,
                                   num_dependent_values):
        suggestions = []
        if num_dependent_values == 2:
            if self.is_sample_large(sample_size, num_features):
                if regularization == LANGUAGES[self.current_language]["regularization_options"][0]:
                    suggestions.extend(["Regularized Binary Logistic Regression"])
                else:
                    suggestions.extend(["Binary Logistic Regression", "Binary Logit Regression"])
            else:
                if regularization == LANGUAGES[self.current_language]["regularization_options"][0]:
                    suggestions.extend(["Regularized Binary Logistic Regression"])
                else:
                    suggestions.extend(["Binary Logit Regression"])
        elif num_dependent_values > 2:
            if self.is_sample_large(sample_size, num_features):
                if regularization == LANGUAGES[self.current_language]["regularization_options"][0]:
                    suggestions.extend(["Regularized Multinomial Logistic Regression"])
                else:
                    suggestions.extend(["Multinomial Logistic Regression", "Multinomial Logit Regression"])
            else:
                if regularization == LANGUAGES[self.current_language]["regularization_options"][0]:
                    suggestions.extend(["Regularized Multinomial Logistic Regression"])
                else:
                    suggestions.extend(["Multinomial Logit Regression"])

        if independent_variable_type == LANGUAGES[self.current_language]["independent_variable_options"][1]:
            suggestions = [s + " (need encoding)" for s in suggestions]

        return suggestions

    def suggest_nonlinear(self, independent_variable_type, regularization, sample_size, num_features):
        suggestions = []
        if self.is_sample_large(sample_size, num_features):
            if regularization == LANGUAGES[self.current_language]["regularization_options"][0]:
                suggestions.extend(["Regularized Polynomial Regression"])
            else:
                suggestions.extend(["Polynomial Regression"])
        else:
            if regularization == LANGUAGES[self.current_language]["regularization_options"][0]:
                suggestions.extend(["Regularized Polynomial Regression"])
            else:
                suggestions.extend(["Nonlinear Regression"])

        if independent_variable_type == LANGUAGES[self.current_language]["independent_variable_options"][1]:
            suggestions = [s + " (need encoding)" for s in suggestions]

        return suggestions

    def suggest_regression_method(self):
        try:
            # 获取用户输入
            num_features = int(self.num_features_entry.get())
            sample_size = int(self.sample_size_entry.get())
            relationship_type = self.relationship_var.get()
            dependent_variable_type = self.dependent_variable_var.get()
            multicollinearity = self.multicollinearity_var.get()
            independent_variable_type = self.independent_variable_var.get()
            regularization = self.regularization_var.get()
            truncation = self.truncation_var.get()
            outliers = self.outliers_var.get()

            suggestions = []

            # 根据不同条件建议回归方法
            if relationship_type == LANGUAGES[self.current_language]["relationship_options"][0]:
                if dependent_variable_type == LANGUAGES[self.current_language]["dependent_variable_options"][0]:
                    suggestions = self.suggest_linear_continuous(multicollinearity, independent_variable_type,
                                                                 regularization, sample_size, num_features, truncation,
                                                                 outliers)
                elif dependent_variable_type == LANGUAGES[self.current_language]["dependent_variable_options"][1]:
                    num_dependent_values = int(self.num_dependent_values_entry.get())
                    suggestions = self.suggest_linear_categorical(independent_variable_type, regularization,
                                                                  sample_size, num_features, num_dependent_values)
            elif relationship_type == LANGUAGES[self.current_language]["relationship_options"][1]:
                suggestions = self.suggest_nonlinear(independent_variable_type, regularization, sample_size,
                                                     num_features)

            if suggestions:
                if self.current_language == "中文":
                    chinese_suggestions = []
                    for suggestion in suggestions:
                        base_method = suggestion.split(" (")[0]
                        if base_method in LANGUAGES[self.current_language]["regression_methods"]:
                            chinese_method = LANGUAGES[self.current_language]["regression_methods"][base_method]
                            if " (need encoding)" in suggestion:
                                chinese_method += "（需要编码）"
                            chinese_suggestions.append(chinese_method)
                        else:
                            chinese_suggestions.append(suggestion)
                    suggestion_str = "建议的回归分析方法有：" + "，".join(chinese_suggestions)
                else:
                    suggestion_str = "Suggested regression analysis methods are: " + ", ".join(suggestions)
                # 使用 Toplevel 窗口显示结果
                result_window = ttk.Toplevel(self.root)
                result_window.title(LANGUAGES[self.current_language]["suggestion_title"])

                # 设置弹出框大小
                window_width = 400
                window_height = 200
                result_window.geometry(f"{window_width}x{window_height}")

                # 计算并设置窗口位置使其在桌面正中央
                screen_width = self.root.winfo_screenwidth()
                screen_height = self.root.winfo_screenheight()
                x = (screen_width - window_width) // 2
                y = (screen_height - window_height) // 2
                result_window.geometry(f"{window_width}x{window_height}+{x}+{y}")

                text_widget = tk.Text(result_window, wrap='word')
                text_widget.insert(tk.END, suggestion_str)
                text_widget.pack(padx=10, pady=10, fill='both', expand=True)
                text_widget.config(state='disabled')  # 禁止编辑
            else:
                # 使用 Toplevel 窗口显示无建议信息
                result_window = ttk.Toplevel(self.root)
                result_window.title(LANGUAGES[self.current_language]["suggestion_title"])

                # 设置弹出框大小
                window_width = 400
                window_height = 200
                result_window.geometry(f"{window_width}x{window_height}")

                # 计算并设置窗口位置使其在桌面正中央
                screen_width = self.root.winfo_screenwidth()
                screen_height = self.root.winfo_screenheight()
                x = (screen_width - window_width) // 2
                y = (screen_height - window_height) // 2
                result_window.geometry(f"{window_width}x{window_height}+{x}+{y}")

                text_widget = tk.Text(result_window, wrap='word')
                text_widget.insert(tk.END, LANGUAGES[self.current_language]["no_suggestion"])
                text_widget.pack(padx=10, pady=10, fill='both', expand=True)
                text_widget.config(state='disabled')  # 禁止编辑
        except ValueError:
            messagebox.showerror(LANGUAGES[self.current_language]["suggestion_title"],
                                 LANGUAGES[self.current_language]["input_error"])

    def change_language(self):
        self.current_language = "English" if self.current_language == "中文" else "中文"
        # 更新界面文本
        self.root.title(LANGUAGES[self.current_language]["title"])
        self.num_features_label.config(text=LANGUAGES[self.current_language]["num_features_label"])
        self.sample_size_label.config(text=LANGUAGES[self.current_language]["sample_size_label"])
        self.num_dependent_values_label.config(text=LANGUAGES[self.current_language]["num_dependent_values_label"])
        self.relationship_label.config(text=LANGUAGES[self.current_language]["relationship_label"])
        self.relationship_var.set(LANGUAGES[self.current_language]["relationship_options"][0])
        self.relationship_combobox['values'] = LANGUAGES[self.current_language]["relationship_options"]
        self.relationship_combobox.set(LANGUAGES[self.current_language]["relationship_options"][0])

        self.dependent_variable_label.config(text=LANGUAGES[self.current_language]["dependent_variable_label"])
        self.dependent_variable_var.set(LANGUAGES[self.current_language]["dependent_variable_options"][0])
        self.dependent_variable_combobox['values'] = LANGUAGES[self.current_language]["dependent_variable_options"]
        self.dependent_variable_combobox.set(LANGUAGES[self.current_language]["dependent_variable_options"][0])

        self.multicollinearity_label.config(text=LANGUAGES[self.current_language]["multicollinearity_label"])
        self.multicollinearity_var.set(LANGUAGES[self.current_language]["multicollinearity_options"][1])
        self.multicollinearity_combobox['values'] = LANGUAGES[self.current_language]["multicollinearity_options"]
        self.multicollinearity_combobox.set(LANGUAGES[self.current_language]["multicollinearity_options"][1])

        self.independent_variable_label.config(text=LANGUAGES[self.current_language]["independent_variable_label"])
        self.independent_variable_var.set(LANGUAGES[self.current_language]["independent_variable_options"][0])
        self.independent_variable_combobox['values'] = LANGUAGES[self.current_language]["independent_variable_options"]
        self.independent_variable_combobox.set(LANGUAGES[self.current_language]["independent_variable_options"][0])

        self.regularization_label.config(text=LANGUAGES[self.current_language]["regularization_label"])
        self.regularization_var.set(LANGUAGES[self.current_language]["regularization_options"][1])
        self.regularization_combobox['values'] = LANGUAGES[self.current_language]["regularization_options"]
        self.regularization_combobox.set(LANGUAGES[self.current_language]["regularization_options"][1])

        self.truncation_label.config(text=LANGUAGES[self.current_language]["truncation_label"])
        self.truncation_var.set(LANGUAGES[self.current_language]["truncation_options"][1])
        self.truncation_combobox['values'] = LANGUAGES[self.current_language]["truncation_options"]
        self.truncation_combobox.set(LANGUAGES[self.current_language]["truncation_options"][1])

        self.outliers_label.config(text=LANGUAGES[self.current_language]["outliers_label"])
        self.outliers_var.set(LANGUAGES[self.current_language]["outliers_options"][1])
        self.outliers_combobox['values'] = LANGUAGES[self.current_language]["outliers_options"]
        self.outliers_combobox.set(LANGUAGES[self.current_language]["outliers_options"][1])

        self.analyze_button.config(text=LANGUAGES[self.current_language]["analyze_button"])
        self.change_language_label.config(text=LANGUAGES[self.current_language]["change_language_text"])
        self.open_other_script_button.config(text=LANGUAGES[self.current_language]["open_other_script_button"])

    def open_other_script(self):
        try:
            AnalyzerApp(ttk.Toplevel(self.root))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open the script: {str(e)}")

    def get_max_length(self, options):
        max_length = 0
        for option in options:
            length = len(option)
            if length > max_length:
                max_length = length
        return max_length

    def create_ui(self):
        # 设置界面长宽
        window_width = 1000  # 增加宽度以适应两列布局
        window_height = 700
        self.root.geometry(f"{window_width}x{window_height}")

        # 创建一个主框架，用于布局
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(expand=True, fill='both')

        # 创建左右两个子框架
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side='left', padx=10, fill='y')

        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side='right', padx=10, fill='y')

        # 创建样式对象
        style = ttk.Style()
        # 设置 Combobox 文字居中
        style.configure('TCombobox', anchor='center')

        # 特征数量输入
        self.num_features_label = ttk.Label(left_frame, text=LANGUAGES[self.current_language]["num_features_label"])
        self.num_features_label.pack(pady=5)
        self.num_features_entry = ttk.Entry(left_frame)
        self.num_features_entry.pack(pady=5)

        # 样本量输入
        self.sample_size_label = ttk.Label(left_frame, text=LANGUAGES[self.current_language]["sample_size_label"])
        self.sample_size_label.pack(pady=5)
        self.sample_size_entry = ttk.Entry(left_frame)
        self.sample_size_entry.pack(pady=5)

        # 因变量取值数量输入
        self.num_dependent_values_label = ttk.Label(left_frame,
                                                    text=LANGUAGES[self.current_language]["num_dependent_values_label"])
        self.num_dependent_values_label.pack(pady=5)
        self.num_dependent_values_entry = ttk.Entry(left_frame)
        self.num_dependent_values_entry.pack(pady=5)

        # 因变量和自变量的关系类型选择
        self.relationship_label = ttk.Label(right_frame, text=LANGUAGES[self.current_language]["relationship_label"])
        self.relationship_label.pack(pady=5)
        self.relationship_var = tk.StringVar()
        self.relationship_combobox = ttk.Combobox(right_frame, textvariable=self.relationship_var,
                                                  values=LANGUAGES[self.current_language]["relationship_options"])
        self.relationship_combobox.set(LANGUAGES[self.current_language]["relationship_options"][0])
        self.relationship_combobox.pack(pady=5)

        # 因变量类型选择
        self.dependent_variable_label = ttk.Label(right_frame,
                                                  text=LANGUAGES[self.current_language]["dependent_variable_label"])
        self.dependent_variable_label.pack(pady=5)
        self.dependent_variable_var = tk.StringVar()
        self.dependent_variable_combobox = ttk.Combobox(right_frame, textvariable=self.dependent_variable_var,
                                                        values=LANGUAGES[self.current_language][
                                                            "dependent_variable_options"])
        self.dependent_variable_combobox.set(LANGUAGES[self.current_language]["dependent_variable_options"][0])
        self.dependent_variable_combobox.pack(pady=5)

        # 数据是否存在多重共线性选择
        self.multicollinearity_label = ttk.Label(right_frame,
                                                 text=LANGUAGES[self.current_language]["multicollinearity_label"])
        self.multicollinearity_label.pack(pady=5)
        self.multicollinearity_var = tk.StringVar()
        self.multicollinearity_combobox = ttk.Combobox(right_frame, textvariable=self.multicollinearity_var,
                                                       values=LANGUAGES[self.current_language][
                                                           "multicollinearity_options"])
        self.multicollinearity_combobox.set(LANGUAGES[self.current_language]["multicollinearity_options"][1])
        self.multicollinearity_combobox.pack(pady=5)

        # 自变量类型选择
        self.independent_variable_label = ttk.Label(right_frame,
                                                    text=LANGUAGES[self.current_language]["independent_variable_label"])
        self.independent_variable_label.pack(pady=5)
        self.independent_variable_var = tk.StringVar()
        self.independent_variable_combobox = ttk.Combobox(right_frame, textvariable=self.independent_variable_var,
                                                          values=LANGUAGES[self.current_language][
                                                              "independent_variable_options"])
        self.independent_variable_combobox.set(LANGUAGES[self.current_language]["independent_variable_options"][0])
        self.independent_variable_combobox.pack(pady=5)

        # 是否需要正则化选择
        self.regularization_label = ttk.Label(right_frame,
                                              text=LANGUAGES[self.current_language]["regularization_label"])
        self.regularization_label.pack(pady=5)
        self.regularization_var = tk.StringVar()
        self.regularization_combobox = ttk.Combobox(right_frame, textvariable=self.regularization_var,
                                                    values=LANGUAGES[self.current_language]["regularization_options"])
        self.regularization_combobox.set(LANGUAGES[self.current_language]["regularization_options"][1])
        self.regularization_combobox.pack(pady=5)

        # 数据是否有截断情况选择
        self.truncation_label = ttk.Label(right_frame, text=LANGUAGES[self.current_language]["truncation_label"])
        self.truncation_label.pack(pady=5)
        self.truncation_var = tk.StringVar()
        self.truncation_combobox = ttk.Combobox(right_frame, textvariable=self.truncation_var,
                                                values=LANGUAGES[self.current_language]["truncation_options"])
        self.truncation_combobox.set(LANGUAGES[self.current_language]["truncation_options"][1])
        self.truncation_combobox.pack(pady=5)

        # 数据是否有异常值选择
        self.outliers_label = ttk.Label(right_frame, text=LANGUAGES[self.current_language]["outliers_label"])
        self.outliers_label.pack(pady=5)
        self.outliers_var = tk.StringVar()
        self.outliers_combobox = ttk.Combobox(right_frame, textvariable=self.outliers_var,
                                              values=LANGUAGES[self.current_language]["outliers_options"])
        self.outliers_combobox.set(LANGUAGES[self.current_language]["outliers_options"][1])
        self.outliers_combobox.pack(pady=5)

        # 创建底部框架
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(side=tk.BOTTOM, pady=20)

        # 分析并建议按钮
        self.analyze_button = ttk.Button(bottom_frame, text=LANGUAGES[self.current_language]["analyze_button"],
                                         command=self.suggest_regression_method)
        self.analyze_button.pack(side=tk.LEFT, padx=10)

        # 打开分析器按钮
        self.open_other_script_button = ttk.Button(bottom_frame,
                                                   text=LANGUAGES[self.current_language]["open_other_script_button"],
                                                   command=self.open_other_script)
        self.open_other_script_button.pack(side=tk.LEFT, padx=10)

        # 切换语言按钮
        self.change_language_label = ttk.Button(bottom_frame,
                                                text=LANGUAGES[self.current_language]["change_language_text"],
                                                command=self.change_language)
        self.change_language_label.pack(side=tk.LEFT, padx=10)


if __name__ == "__main__":
    app = RegressionAnalysisSelector()
    app.root.mainloop()
