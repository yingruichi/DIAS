import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import messagebox, PhotoImage
import subprocess
import os

# 导入模块
from Dataset import DatasetApp
from Clustering import ClusteringApp
from Questionnaire_analysis import QuestionnaireAnalysisApp
from Regression_prediction_model_and_influence_relationship import RegressionPredictionApp
from Statistical_Modeling import StatisticalModelingApp
from Econometric_Model import EconometricModelApp
from Difference_analysis import DifferenceAnalysisApp
from Design_scheme_selection_and_comprehensive_evaluation import DesignSchemeSelectionApp
from Data_Description_and_Validation import DataDescriptionApp
from Correlation_analysis import CorrelationAnalysisApp
from Source.Analyzer import AnalyzerApp

# 获取当前脚本所在的目录
current_dir = os.path.dirname(os.path.abspath(__file__))

# 定义语言字典
LANGUAGES = {
    'zh': {
        'title': "迪亚士   设计信息学分析系统",
        'no_details': " 欢迎使用 '迪亚士‘ 设计信息学分析系统 \n 这是一款专为 ‘设计信息学’ 研发的统计分析软件 \n 适合设计学研究者和跨专业研究者。",
        'switch_language': "切换语言",
        'copy_success': "内容已复制到剪贴板",
        'group1': "查找数据",
        'group2': "数据的可靠性",
        'group3': "分析器的分类",
        'group4': "查找分析器并开始分析",
        'details': "详情",
        'copyright': "博士生:池英瑞/ 卡梅里诺大学/ CSC"
    },
    'en': {
        'title': "DIAS   Design Informatics Analysis System",
        'no_details': " Welcome to ‘DIAS’ Design Informatics Analysis System \n This is a statistical analysis software developed specifically for 'Design informatics' \n It is suitable for design researchers and cross-disciplinary researchers.",
        'switch_language': "Switch Language",
        'copy_success': "Content has been copied to the clipboard",
        'group1': "Find data",
        'group2': "Reliability of data",
        'group3': "Classification of analyzers",
        'group4': "Find an analyzer and start analyzing",
        'details': "Details",
        'copyright': "PhD student: Chi Yingrui/ Università degli studi di Camerino/ CSC"
    }
}

# 定义按钮文本的语言字典
BUTTON_TEXTS = {
    'zh': {
        "数据库": "数据库",
        "数据描述与检验": "数据描述与检验",
        "问卷分析": "问卷分析",
        "相关性分析": "相关性分析",
        "差异性分析": "差异性分析",
        "设计方案选择与综合评价": "设计方案选择与综合评价",
        "回归预测模型与影响关系": "回归预测模型与影响关系",
        "聚类": "聚类",
        "统计建模": "统计建模",
        "计量经济模型": "计量经济模型",
        "分析器": "分析器"
    },
    'en': {
        "数据库": "Dataset",
        "数据描述与检验": "Data description and test",
        "问卷分析": "Questionnaire analysis",
        "相关性分析": "Correlation analysis",
        "差异性分析": "Difference analysis",
        "设计方案选择与综合评价": "Design scheme selection and comprehensive evaluation",
        "回归预测模型与影响关系": "Regression prediction model and influence relationship",
        "聚类": "Clustering",
        "统计建模": "Statistical modeling",
        "计量经济模型": "Econometric model",
        "分析器": "Analyzer"
    }
}

# 定义按钮对应的详情信息
DETAILS_INFO = {
    "数据库": {
        'zh': "该功能用于访问数据库，查找数据，是完成数据分析类研究的第一步。",
        'en': "This function is used to access a database and retrieve data, serving as the first step in conducting data-driven analytical research."
    },
    "数据描述与检验": {
        'zh': " 数据描述与检验：\n 对设计数据进行特征概括与假设验证。例如：在包装设计测试中，描述用户反馈数据特征，检验设计方案是否达标。\n 课题：\n 某品牌食品新包装设计测试数据描述与效果检验。 \n 某运动品牌运动鞋新配色设计测试的数据描述与市场接受度检验 \n 沉浸式文旅空间照明设计用户反馈的数据描述与效果验证 \n 智能穿戴设备交互手势设计测试的数据描述与易用性检验 \n 博物馆展陈空间动线设计的数据描述与参观流畅度检验 \n 新能源汽车外观设计用户评价的数据描述与审美趋势检验",
        'en': " Data Description and Testing: \n Summarizes the characteristics of design-related data and validates hypotheses. For example, In packaging design testing, this involves describing user feedback data and testing whether the design meets standards.\n Research Topic: \n Descriptive Analysis and Effectiveness Testing of a New Packaging Design for a Food Brand. \n Descriptive Analysis and Market Acceptance Testing of New Colorway Design for a Sports Shoe Brand \n Descriptive Analysis and Effectiveness Testing of User Feedback on Immersive Cultural Tourism Lighting Design \n Descriptive Analysis and Usability Testing of Interaction Gestures in Smart Wearable Devices \n Descriptive Data Analysis and Flow Testing of Museum Exhibition Space Path Design \n Descriptive Analysis and Aesthetic Trend Testing of User Evaluation on New Energy Vehicle Exterior Design"
    },
    "问卷分析": {
        'zh': "对问卷和数据进行分析，例如：调查问卷的设计是否合理，问卷中的问题是否有关联，收集的数据是否有效、是否有可信度。",
        'en': "Questionnaire analysis involves evaluating both the questionnaire and the data it generates. For example, This includes assessing the rationality of the questionnaire design, examining the logical relationships between questions, and determining the validity and reliability of the collected data."
    },
    "相关性分析": {
        'zh': " 相关性分析：\n 探究设计变量间关联程度。例如：在室内设计中，分析空间布局、色彩搭配与用户舒适度的相关性，优化设计方案。\n 课题：\n 办公空间色彩搭配与员工工作效率的相关性研究及设计优化。\n 产品材质触感与用户购买意愿的相关性研究及设计策略 \n 网页字体排版与用户阅读时长的相关性分析及界面优化 \n 园林景观植物配置与游客停留时间的相关性及设计应用 \n 汽车内饰色彩搭配与驾驶员情绪状态的相关性研究 \n 灯具光照强度分布与办公空间氛围营造的相关性分析",
        'en': " Correlation Analysis: \n Explores the degree of association between design variables. For example, In interior design, this can involve analyzing the correlation between spatial layout, color schemes, and user comfort to optimize design solutions.\n Research Topic: \n A Study on the Correlation Between Office Space Color Schemes and Employee Work Efficiency and Design Optimization \n Study on the Correlation Between Product Material Texture and User Purchase Intention and Design Strategy \n Correlation Analysis Between Web Font Layout and User Reading Duration and Interface Optimization \n Correlation and Design Application of Plant Configuration in Landscape Gardens and Visitor Stay Duration \n Study on the Correlation Between Automobile Interior Color Matching and Driver Emotional States \n Correlation Analysis Between Lighting Intensity Distribution and Office Atmosphere Design"
    },
    "差异性分析": {
        'zh': " 差异性分析：\n 对比不同设计方案、用户群体的差异。例如：通过分析不同年龄段用户对产品外观设计的接受度差异，为设计定位提供依据。\n 课题：\n Z 世代与千禧一代对潮玩产品外观设计的差异性分析。\n 国内外高校校园文化空间设计风格的差异性分析与启示 \n 不同收入阶层对家居智能设备外观设计的差异性研究 \n 线上线下展览空间用户体验的差异性分析与设计改进 \n 东西方传统建筑装饰元素应用的差异性及现代转化研究 \n 老年群体与青年群体对适老化手机界面设计的差异性探讨",
        'en': " Difference Analysis: \n Compares differences among various design solutions or user groups. For instance, analyzing how users of different age groups perceive product appearance can guide design positioning.\n Research Topic: \n A Comparative Study on the Differences in Appearance Design Preferences for Trendy Toys Between Generation Z and Millennials \n Comparative Analysis and Insights on the Design Styles of Campus Cultural Spaces in Domestic and International Universities \n Study on the Design Preference Differences of Smart Home Device Appearances Among Different Income Groups \n User Experience Differences Between Online and Offline Exhibition Spaces and Design Improvement \n Study on the Differences and Modern Transformation of Traditional Architectural Decorative Elements in the East and West \n Discussion on the Interface Design Differences for Elderly and Young User Groups in Aging-Friendly Mobile Applications"
    },
    "设计方案选择与综合评价": {
        'zh': " 综合评价：\n 从多维度评估设计方案优劣。例如：对多个建筑设计方案，综合考虑美观性、实用性、经济性等指标评价排序。\n 课题：\n 城市公共建筑设计方案的综合评价体系构建与应用。 \n 乡村振兴背景下传统村落改造设计方案的综合评价体系构建与应用 \n 城市更新项目中历史建筑活化利用设计方案的综合评价研究 \n 虚拟现实（VR）教育场景设计方案的多维度综合评价体系构建 \n 绿色建筑设计方案的生态、经济、社会指标综合评价与实践 \n 商业综合体公共艺术装置设计方案的综合评价与优选策略",
        'en': " Comprehensive Evaluation: \n Assesses the quality of design proposals from multiple dimensions. For example, For architectural design solutions, it evaluates and ranks them based on aesthetics, functionality, and cost-effectiveness.\n Research Topic: \n Construction and Application of a Comprehensive Evaluation System for Urban Public Building Design Schemes \n Construction and Application of a Comprehensive Evaluation System for Traditional Village Renovation Design under Rural Revitalization \n Comprehensive Evaluation of Design Schemes for the Adaptive Reuse of Historical Buildings in Urban Renewal Projects \n Construction of a Multidimensional Evaluation System for Virtual Reality (VR) Educational Scenario Design \n Comprehensive Evaluation and Practice of Ecological, Economic, and Social Indicators in Green Building Design \n Comprehensive Evaluation and Optimization Strategy for Public Art Installation Design in Commercial Complexes"
    },
    "回归预测模型与影响关系": {
        'zh': " 回归分析：\n 建立设计因素与结果间的数学关系模型。例如：分析产品功能复杂度、价格与销量的关系，指导产品功能设计与定价。\n 课题：\n 智能家居产品功能配置与市场销量的回归分析及设计策略。\n 健身器材功能多样性、价格与消费者购买倾向的回归分析及设计策略 \n 电子书阅读界面字体大小、行距与用户阅读疲劳度的回归分析 \n 主题餐厅空间尺度、装饰风格与客流量的回归分析及设计优化 \n 儿童学习桌椅调节功能、安全性与家长购买决策的回归研究 \n 智能音箱外观尺寸、造型设计与市场占有率的回归分析及设计导向",
        'en': " Regression Analysis: \n Establishes mathematical models between design factors and outcomes. For example, It can be used to analyze the relationship between product feature complexity, pricing, and sales volume to inform product design and pricing strategies.\n Research Topic: \n Regression Analysis of Smart Home Product Feature Configuration and Market Sales, and Design Strategy \n Regression Analysis of Fitness Equipment Function Diversity, Price, and Consumer Purchase Inclination and Design Strategy \n Regression Analysis of Font Size, Line Spacing, and User Reading Fatigue on E-book Interfaces \n Regression Analysis and Design Optimization of Theme Restaurant Space Scale, Decoration Style, and Customer Flow \n Regression Study on the Adjustment Function and Safety of Children's Study Furniture and Parental Purchase Decisions \n Regression Analysis of Smart Speaker Appearance Size, Design Form, and Market Share and Design Guidance"
    },
    "聚类": {
        'zh': " 聚类：\n 可将设计学中具有相似特征的元素、用户或案例分类。例如：在用户调研中，依据用户审美偏好、使用习惯聚类，明确细分群体需求，助力产品差异化设计。\n 课题：\n 基于聚类分析的智能手表用户需求分类与界面设计研究。 \n 基于聚类分析的户外家具用户使用场景分类与形态设计研究 \n 运用聚类对短视频平台用户界面交互偏好分类及设计优化研究 \n 聚类视角下的儿童玩具用户情感需求分类与造型设计实践 \n 基于聚类的城市公共艺术风格分类及在地化设计策略研究 \n 聚类分析在传统服饰纹样风格分类与现代创新设计中的应用",
        'en': " Clustering: \n Used to categorize elements, users, or cases with similar characteristics in design studies. For example, in user research, clustering based on aesthetic preferences and usage habits helps identify the needs of specific user segments and supports differentiated product design.\n Research Topic: \n A Study on User Demand Classification and Interface Design of Smartwatches Based on Cluster Analysis \n A Study on the Classification of Outdoor Furniture User Scenarios and Form Design Based on Cluster Analysis \n Research on UI Interaction Preference Classification and Design Optimization for Short Video Platform Users Using Clustering \n Emotional Needs Classification and Modeling Design Practice for Children’s Toy Users from a Clustering Perspective \n Study on the Style Classification of Urban Public Art and Localization Design Strategy Based on Clustering \n Application of Cluster Analysis in the Classification of Traditional Clothing Pattern Styles and Modern Innovative Design"
    },
    "统计建模": {
        'zh': " 统计建模：\n 将设计问题抽象为数学模型求解。例如：针对用户体验设计，建立模型预测用户对不同交互流程的满意度，优化交互设计。\n 课题：\n 移动应用交互流程设计的统计建模与优化研究。\n 基于统计建模的游戏角色动作流畅度与玩家沉浸感优化研究 \n 统计建模在无障碍公共设施设计人机工程参数优化中的应用 \n 运用统计建模预测智能家电产品功能组合的用户满意度 \n 统计建模在城市公园景观节点布局与游客活动强度关系中的应用 \n 基于统计建模的美妆产品包装设计与消费者购买行为预测研究",
        'en': " Statistical Modeling: \n Abstracts design problems into mathematical models for problem-solving. For example, In user experience design, models can predict user satisfaction with different interaction flows, aiding in interaction design optimization.\n Research Topic: \n Statistical Modeling and Optimization Research on Interaction Flow Design in Mobile Applications \n Optimization of Game Character Motion Smoothness and Player Immersion Based on Statistical Modeling \n Application of Statistical Modeling in the Ergonomic Parameter Optimization of Accessible Public Facility Design \n Using Statistical Modeling to Predict User Satisfaction with Smart Home Product Feature Combinations \n Application of Statistical Modeling in the Relationship Between Urban Park Landscape Node Layout and Visitor Activity Intensity \n Statistical Modeling and Predictive Study of Cosmetic Packaging Design and Consumer Purchase Behavior"
    },
    "计量经济模型": {
        'zh': " 计量经济模型：\n 分析设计决策与经济指标的关系。例如：在文创产品设计中，构建模型分析设计成本、定价与利润的关系，辅助商业决策。\n 课题：\n 文创产品设计要素与经济效益的计量经济模型构建与应用。\n 影视 IP 衍生品设计投入、市场推广与销售收益的计量经济模型构建 \n 快时尚服装款式更新频率、生产成本与利润的计量经济分析 \n 数字艺术展览设计投入与门票收入、衍生品销售的计量经济模型 \n 农产品包装设计成本、品牌溢价与市场份额的计量经济研究 \n 线上教育平台界面设计优化投入与用户留存率的计量经济分析",
        'en': " Econometric Modeling: \n Analyzes the relationship between design decisions and economic indicators. For example, In cultural and creative product design, econometric models can help analyze the relationship between design cost, pricing, and profit to support business decisions.\n Research Topic: \n Construction and Application of an Econometric Model on the Design Elements and Economic Benefits of Cultural and Creative Products \n Construction of an Econometric Model for Design Investment, Marketing, and Sales Revenue of Film IP Derivatives \n Econometric Analysis of Style Update Frequency, Production Cost, and Profit in Fast Fashion Clothing \n Econometric Modeling of Design Investment, Ticket Revenue, and Derivative Sales in Digital Art Exhibitions \n Econometric Study on Agricultural Product Packaging Design Cost, Brand Premium, and Market Share \n Econometric Analysis of Interface Design Optimization Investment and User Retention Rate in Online Education Platforms"
    },
    "分析器": {
        'zh': "本系统共包含103个分析器，都可以在这里找到。它们已经自动采用最优的参数设置，因此，我们可以以最简单的方式使用它们。",
        'en': "This system contains a total of 103 analyzers, which can be found here. They have automatically adopted the optimal parameter settings, so we can use them in the simplest way."
    }
}

# 当前语言
current_language = 'en'

# 用于存储每个按钮的最大宽度
button_max_widths = []

# 记录是否已经显示过详情
has_shown_details = False


def run_script(file_path):
    try:
        # print(file_path)
        if file_path == "Dataset":
            DatasetApp(ttk.Toplevel(root))
        elif file_path == "Clustering":
            ClusteringApp(ttk.Toplevel(root))
        elif file_path == "Questionnaire analysis":
            QuestionnaireAnalysisApp(ttk.Toplevel(root))
        elif file_path == "Regression prediction model and influence relationship":
            RegressionPredictionApp(ttk.Toplevel(root))
        elif file_path == "Statistical Modeling":
            StatisticalModelingApp(ttk.Toplevel(root))
        elif file_path == "Econometric Model":
            EconometricModelApp(ttk.Toplevel(root))
        elif file_path == "Difference analysis":
            DifferenceAnalysisApp(ttk.Toplevel(root))
        elif file_path == "Design scheme selection and comprehensive evaluation":
            DesignSchemeSelectionApp(ttk.Toplevel(root))
        elif file_path == "Data Description and Validation":
            DataDescriptionApp(ttk.Toplevel(root))
        elif file_path == "Correlation analysis":
            CorrelationAnalysisApp(ttk.Toplevel(root))
        elif file_path == "Analyzer":
            AnalyzerApp(ttk.Toplevel(root))
        else:
            import sys
            subprocess.Popen([sys.executable, file_path])
    except Exception as e:
        messagebox.showerror("错误", f"运行脚本时出错: {e}")


def switch_language():
    global current_language, has_shown_details
    current_language = 'zh' if current_language == 'en' else 'en'
    root.title(LANGUAGES[current_language]['title'])
    switch_language_label.config(text=LANGUAGES[current_language]['switch_language'])
    group1_label.config(text=LANGUAGES[current_language]['group1'])
    group2_label.config(text=LANGUAGES[current_language]['group2'])
    group3_label.config(text=LANGUAGES[current_language]['group3'])
    group4_label.config(text=LANGUAGES[current_language]['group4'])
    details_label.config(text=LANGUAGES[current_language]['details'])
    copyright_label.config(text=LANGUAGES[current_language]['copyright'])

    # 更新按钮文本
    for index, button in enumerate(button_list):
        original_text = button_texts[index]
        display_text = BUTTON_TEXTS[current_language][original_text]
        button.config(text=display_text, width=button_max_widths[index])

    # 重置详情显示状态
    has_shown_details = False
    # 如果还没有显示过详情，显示欢迎信息
    if not has_shown_details:
        details_text.delete(1.0, ttk.END)
        details_text.insert(ttk.END, LANGUAGES[current_language]['no_details'])


def show_details(event, text):
    global has_shown_details
    details_text.delete(1.0, ttk.END)
    details_text.insert(ttk.END, DETAILS_INFO[text][current_language])
    has_shown_details = True


def hide_details(event):
    # 鼠标离开时不清除详情内容，保持当前显示
    pass

# 创建主窗口
root = ttk.Window(themename="flatly")
root.title(LANGUAGES[current_language]['title'])

# 加载图标
icon_path = os.path.join(current_dir, 'icon', 'icon.gif')
print(f"尝试加载图标: {icon_path}")  # 添加调试输出

# 检查文件是否存在
if not os.path.exists(icon_path):
    print(f"错误: 图标文件不存在 - {icon_path}")
else:
    # 检查文件是否为有效文件
    if not os.path.isfile(icon_path):
        print(f"错误: 图标路径不是一个文件 - {icon_path}")
    else:
        try:
            icon = PhotoImage(file=icon_path)
            root.iconphoto(True, icon)
            print("图标加载成功")
        except Exception as e:
            print(f"图标加载失败: {str(e)}")
            messagebox.showerror("图标加载错误", f"加载图标时出错: {e}\n\n请确保图标文件存在于指定路径且格式正确。")

# 获取屏幕的宽度和高度
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# 设置窗口的宽度和高度
window_width = 940
window_height = 780

# 计算窗口应该放置的位置
x = (screen_width - window_width) // 2
y = (screen_height - window_height) // 2

# 设置窗口的位置和大小
root.geometry(f"{window_width}x{window_height}+{x}+{y}")

# 创建一个主框架，用于居中内容
main_frame = ttk.Frame(root)
main_frame.pack(expand=True, fill=BOTH, anchor='n')  # 使用 anchor='n' 让框架在顶部居中

# 创建四个子框架来放置每组按钮
group1_frame = ttk.Frame(main_frame)
group1_frame.pack(expand=True, anchor='center')
group2_frame = ttk.Frame(main_frame)
group2_frame.pack(expand=True, anchor='center')
group3_frame = ttk.Frame(main_frame)
group3_frame.pack(expand=True, anchor='center')
group4_frame = ttk.Frame(main_frame)
group4_frame.pack(expand=True, anchor='center')

# 添加每组的标题标签
group1_label = ttk.Label(group1_frame, text=LANGUAGES[current_language]['group1'])
group1_label.pack()
group2_label = ttk.Label(group2_frame, text=LANGUAGES[current_language]['group2'])
group2_label.pack()
group3_label = ttk.Label(group3_frame, text=LANGUAGES[current_language]['group3'])
group3_label.pack()
group4_label = ttk.Label(group4_frame, text=LANGUAGES[current_language]['group4'])
group4_label.pack()

# 存储所有按钮的列表
button_list = []
button_texts = []
file_paths = []


def create_buttons(frame, texts, paths, bootstyle=PRIMARY):
    current_row_frame = ttk.Frame(frame)
    current_row_frame.pack(anchor='center')
    total_width = 0
    # 留出一定的余量
    margin = 20
    for text, path in zip(texts, paths):
        # 计算该按钮在两种语言下的最大宽度
        zh_text = BUTTON_TEXTS['zh'][text]
        en_text = BUTTON_TEXTS['en'][text]
        max_width = max(len(zh_text), len(en_text))

        display_text = BUTTON_TEXTS[current_language][text]
        button = ttk.Button(current_row_frame, text=display_text, bootstyle=bootstyle, width=max_width)
        button.pack(side=ttk.LEFT, padx=5, pady=5)
        button.bind("<Button-1>", lambda event, p=path: run_script(p))
        button.bind("<Enter>", lambda event, t=text: show_details(event, t))
        button.bind("<Leave>", hide_details)
        button_list.append(button)
        button_texts.append(text)
        file_paths.append(path)
        button.update_idletasks()
        # 记录按钮的最大宽度
        button_max_widths.append(max_width)
        # 计算按钮宽度加上左右内边距
        button_width = button.winfo_width() + 10
        if total_width + button_width > window_width - margin:
            current_row_frame = ttk.Frame(frame)
            current_row_frame.pack(anchor='center')
            total_width = button_width
        else:
            total_width += button_width


# 第一行按钮
create_buttons(group1_frame, ["数据库"], ['Dataset'])

# 第二行按钮
create_buttons(group2_frame, ["数据描述与检验", "问卷分析"],
               ['Data Description and Validation', 'Questionnaire analysis'])

# 第三行按钮
third_row_texts = ["相关性分析", "差异性分析", "设计方案选择与综合评价",
                   "回归预测模型与影响关系", "聚类", "统计建模", "计量经济模型"]
third_row_paths = ['Correlation analysis', 'Difference analysis',
                   'Design scheme selection and comprehensive evaluation', 'Regression prediction model and influence relationship',
                   'Clustering', 'Statistical Modeling', 'Econometric Model']
create_buttons(group3_frame, third_row_texts, third_row_paths)

# 第四行按钮，将 bootstyle 设置为 SUCCESS 以显示绿色按钮
create_buttons(group4_frame, ["分析器"], ['Analyzer'], bootstyle=SUCCESS)

# 创建详情框
details_frame = ttk.Frame(main_frame)
details_frame.pack(expand=True, fill=BOTH, padx=10, pady=10)

details_label = ttk.Label(details_frame, text=LANGUAGES[current_language]['details'])
details_label.pack()

# 修改 font 参数，使用元组指定字体和大小
details_text = ttk.Text(details_frame, height=15, font=('TkDefaultFont', 12))
details_text.pack(fill=BOTH, expand=True)

# 初始化详情框内容
details_text.insert(ttk.END, LANGUAGES[current_language]['no_details'])

# 创建语言切换标签，点击可切换语言，颜色设为灰色
switch_language_label = ttk.Label(root, text=LANGUAGES[current_language]['switch_language'], foreground='gray',
                                  cursor='hand2')
switch_language_label.pack(pady=5)
switch_language_label.bind("<Button-1>", lambda event: switch_language())

# 创建版权标签，并设置字体大小为 10
copyright_label = ttk.Label(root, text=LANGUAGES[current_language]['copyright'], foreground='gray', font=('TkDefaultFont', 8))
copyright_label.pack(pady=5)

# 运行主循环
root.mainloop()