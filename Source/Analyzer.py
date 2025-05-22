from operator import imod
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import os
import subprocess
import sys

# 添加父目录到系统路径，以便能够导入子模块
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 导入子模块
from Source.ADF_Test_Analysis import ADFTestAnalysisApp
from Source.Analysis_of_Covariance_ANCOVA import ANCOVAAnalysisApp
from Source.Analytic_Hierarchy_Process_AHP_Analysis import AnalyticHierarchyProcessAHPApp
from Source.Anderson_Darling_Test import AndersonDarlingTestApp
from Source.ARIMA_Model_Analysis import ARIMAModelAnalysisApp
from Source.Bartlett_Test import BartlettTestApp
from Source.Binary_Logistic_Regression_Analysis import BinaryLogisticRegressionAnalysisApp
from Source.Binary_Logit_Regression_Analysis import BinaryLogitRegressionAnalysisApp
from Source.Canonical_Correlation_Analysis import CanonicalCorrelationAnalysisApp
from Source.Chen_Shapiro_Test import ChenShapiroTestApp
from Source.Chi_Square_Goodness_of_Fit_Test import ChiSquareGoodnessOfFitTestApp
from Source.Chi_Squared_Test import ChiSquaredTestApp
from Source.Clustering_Analysis_K_Means import ClusteringAnalysisKMeansApp
# from Source.Cochrans_Q_Test import CochransQTestApp
from Source.Collinearity_Analysis_VIF import CollinearityAnalysisVIFApp
from Source.Composite_Index_Analysis import CompositeIndexAnalysisApp
from Source.Conjoint_Analysis import ConjointAnalysisApp
from Source.Content_Validity_Analysis import ContentValidityAnalysisApp
from Source.Coupling_Coordination_Degree_Model_Analysis import CouplingCoordinationDegreeModelAnalysisApp
from Source.Cramer_von_Mises_Test import CramerVonMisesTestApp
from Source.CRITIC_Weighting_Method_Analysis import CRITICWeightingMethodAnalysisApp
from Source.DAgostino_K_Squared_Test import DAgostinoKSquaredTestApp
from Source.Delphi_Method_Analysis import DelphiMethodAnalysisApp
from Source.DEMATEL_Analysis import DEMATELAnalysisApp
from Source.Density_Based_Clustering_Analysis import DensityBasedClusteringAnalysisApp
from Source.Descriptive_Statistics import DescriptiveStatisticsApp
from Source.Discriminant_Analysis import DiscriminantAnalysisApp
from Source.Efficacy_Coefficient_Analysis import EfficacyCoefficientAnalysisApp
from Source.Entropy_Method_Analysis import EntropyMethodAnalysisApp
from Source.Exponential_Smoothing_Method_Analysis import ExponentialSmoothingMethodAnalysisApp
from Source.Factor_Analysis import FactorAnalysisApp
from Source.Friedman_Test_Analysis import FriedmanTestApp
from Source.Fuzzy_Analytic_Hierarchy_Process_FAHP_Analysis import FuzzyAnalyticHierarchyProcessFAHPApp
from Source.Fuzzy_Comprehensive_Evaluation_Analysis import FuzzyComprehensiveEvaluationAnalysisApp
from Source.Generalized_Estimating_Equations_Analysis import GeneralizedEstimatingEquationsAnalysisApp
from Source.GMM_Estimation_Analysis import GMMEstimationAnalysisApp
from Source.Gray_Prediction_Model_Analysis import GrayPredictionModelAnalysisApp
from Source.Grey_Relational_Analysis import GreyRelationalAnalysisApp
from Source.Hierarchical_Clustering_Analysis import HierarchicalClusteringAnalysisApp
from Source.Hierarchical_Regression_Analysis import HierarchicalRegressionAnalysisApp
from Source.Independence_Weighting_Method_Analysis import IndependenceWeightingMethodAnalysisApp
from Source.Independent_Samples_T_Test_Analysis import IndependentSamplesTTestAnalysisApp
from Source.Information_Entropy_Weight_Method_Analysis import InformationEntropyWeightMethodAnalysisApp
from Source.Jarque_Bera_Test import JarqueBeraTestApp
from Source.KANO_Model_Analysis import KANOModelAnalysisApp
from Source.Kappa_Consistency_Test import KappaConsistencyTestApp
from Source.Kendall_Correlation_Analysis import KendallCorrelationAnalysisApp
from Source.Kendalls_Coordination_Coefficient import KendallsCoordinationCoefficientApp
from Source.Kolmogorov_Smirnov_Test import KolmogorovSmirnovTestApp
from Source.Lasso_Regression_Analysis import LassoRegressionAnalysisApp
from Source.Levene_Test import LeveneTestApp
from Source.Lilliefors_Test import LillieforsTestApp
from Source.Linear_Tobit_Regression_Analysis import LinearTobitRegressionAnalysisApp
from Source.Markov_Prediction_Analysis import MarkovPredictionAnalysisApp
from Source.Mediation_Analysis import MediationAnalysisApp
from Source.Moderated_Mediation_Analysis import ModeratedMediationAnalysisApp
from Source.Moderation_Analysis import ModerationAnalysisApp
from Source.Multi_sample_ANOVA import MultiSampleANOVAApp
from Source.Multidimensional_Scaling_Analysis import MultidimensionalScalingAnalysisApp
from Source.Multidimensional_Scaling_MDS_Analysis import MultidimensionalScalingMDSApp
from Source.Multinomial_Logistic_Regression_Analysis import MultinomialLogisticRegressionApp
from Source.Multinomial_Logit_Regression_Analysis import MultinomialLogitRegressionApp
from Source.Multiple_choice_Question_Analysis import MultipleChoiceQuestionAnalysisApp
from Source.Multivariate_Analysis_of_Variance_MANOVA import MultivariateManovaApp
from Source.NPS_Net_Promoter_Score_Analysis import NPSNetPromoterScoreAnalysisApp
from Source.Obstacle_Degree_Model_Analysis import ObstacleDegreeModelAnalysisApp
from Source.One_Sample_ANOVA import OneSampleANOVAApp
from Source.One_Sample_t_Test_Analysis import OneSampleTTestAnalysisApp
from Source.One_Sample_Wilcoxon_Test_Analysis import OneSampleWilcoxonTestAnalysisApp
from Source.Ordered_Logit_Regression_Analysis import OrderedLogitRegressionAnalysisApp
from Source.Ordinary_Least_Squares_Linear_Regression_Analysis import OrdinaryLeastSquaresLinearRegressionAnalysisApp
from Source.Paired_t_test_Analysis import PairedTTestAnalysisApp
from Source.Paired_Sample_Wilcoxon_Test_Analysis import PairedSampleWilcoxonTestAnalysisApp
from Source.Partial_Correlation_Analysis import PartialCorrelationAnalysisApp
from Source.Partial_Least_Squares_Regression_Analysis import PartialLeastSquaresRegressionAnalysisApp
from Source.Pearson_Correlation_Analysis import PearsonCorrelationAnalysisApp
from Source.Polynomial_Regression_Analysis import PolynomialRegressionAnalysisApp
from Source.Post_hoc_Multiple_Comparisons import PostHocMultipleComparisonsApp
from Source.Price_Sensitivity_Meter_Analysis import PriceSensitivityMeterAnalysisApp
from Source.Principal_Component_Analysis import PrincipalComponentAnalysisApp
from Source.Range_Analysis import RangeAnalysisApp
from Source.Rank_Sum_Ratio_RSR_Analysis import RankSumRatioRSRAnalysisApp
from Source.Regularized_Binary_Logistic_Regression_Analysis import RegularizedBinaryLogisticRegressionAnalysisApp
from Source.Regularized_Multinomial_Logistic_Regression_Analysis import RegularizedMultinomialLogisticRegressionAnalysisApp
from Source.Reliability_Analysis import ReliabilityAnalysisApp
from Source.Reliability_Test_Analysis import ReliabilityTestAnalysisApp
from Source.Repeated_Measures_ANOVA import RepeatedMeasuresANOVAApp
from Source.Ridge_Regression_Analysis import RidgeRegressionAnalysisApp
from Source.Robust_Linear_Regression_Analysis import RobustLinearRegressionAnalysisApp
from Source.Runs_Test import RunsTestApp
from Source.Second_Order_Clustering_Analysis import SecondOrderClusteringAnalysisApp
from Source.Shapiro_Wilk_Test import ShapiroWilkTestApp
from Source.Spearman_Correlation_Analysis import SpearmanCorrelationAnalysisApp
from Source.Split_Half_Reliability_Analysis import SplitHalfReliabilityAnalysisApp
from Source.Stepwise_Regression_Analysis import StepwiseRegressionAnalysisApp
from Source.Test_Retest_Reliability_Analysis import TestRetestReliabilityAnalysisApp
from Source.TOPSIS_Method_Analysis import TOPSISMethodAnalysisApp
from Source.Turf_Combination_Model_Analysis import TurfCombinationModelAnalysisApp
from Source.Two_Sample_ANOVA import TwoSampleANOVAApp
# from Source.Undesirable_SBM_Model_Analysis import UndesirableSBMModelAnalysisApp
from Source.Validity_Analysis import ValidityAnalysisApp
from Source.Within_Group_Inter_Rater_Reliability_rwg_Analysis import WithinGroupInterRaterReliabilityRwgAnalysisApp



# 全局变量
canvas = None
button_frame = None

# 定义语言字典
LANGUAGES = {
    'zh': {
        'title': "(迪亚士) 分析器",
        'switch_language': "切换语言",
        'error_message': "打开脚本 {} 时出错: {}",
        'search_placeholder': "搜索分析方法"
    },
    'en': {
        'title': "(DIAS) Analyzer",
        'switch_language': "Switch Language",
        'error_message': "Error opening script {}: {}",
        'search_placeholder': "Search analysis methods"
    }
}

# 定义模块映射表
MODULE_MAP = {
    "ADF Test Analysis": {
        "class": ADFTestAnalysisApp,
        "description": {
            "zh": "ADF检验分析",
            "en": "ADF Test Analysis"
        }
    },
    "Analysis of Covariance (ANCOVA)": {
        "class": ANCOVAAnalysisApp,
        "description": {
            "zh": "协方差分析",
            "en": "Analysis of Covariance (ANCOVA)"
        }
    },
    "Analytic Hierarchy Process (AHP)": {
        "class": AnalyticHierarchyProcessAHPApp,
        "description": {
            "zh": "层次分析法 AHP 分析",
            "en": "Analytic Hierarchy Process (AHP) Analysis"
        }
    },
    "Anderson-Darling Test": {
        "class": AndersonDarlingTestApp,
        "description": {
            "zh": "Anderson-Darling 检验",
            "en": "Anderson-Darling Test"
        }
    },
    "ARIMA Model Analysis": {
        "class": ARIMAModelAnalysisApp,
        "description": {
            "zh": "ARIMA模型分析",
            "en": "ARIMA Model Analysis"
        }
    },
    "Bartlett Test": {
        "class": BartlettTestApp,
        "description": {
            "zh": "巴特利特检验",
            "en": "Bartlett Test"
        }
    },
    "Binary Logistic Regression Analysis": {
        "class": BinaryLogisticRegressionAnalysisApp,
        "description": {
            "zh": "二元逻辑回归分析",
            "en": "Binary Logistic Regression Analysis"
        }
    },
    "Binary Logit Regression Analysis": {
        "class": BinaryLogitRegressionAnalysisApp,
        "description": {
            "zh": "二元Logit回归分析",
            "en": "Binary Logit Regression Analysis"
        }
    },
    "Canonical Correlation Analysis": {
        "class": CanonicalCorrelationAnalysisApp,
        "description": {
            "zh": "典型相关分析",
            "en": "Canonical Correlation Analysis"
        }
    },
    "Chen-Shapiro Test": {
        "class": ChenShapiroTestApp,
        "description": {
            "zh": "陈-夏普检验",
            "en": "Chen-Shapiro Test"
        }
    },
    "Chi-Square Goodness-of-Fit Test": {
        "class": ChiSquareGoodnessOfFitTestApp,
        "description": {
            "zh": "卡方拟合优度检验",
            "en": "Chi-Square Goodness-of-Fit Test"
        }
    },
    "Chi-Squared Test": {
        "class": ChiSquaredTestApp,
        "description": {
            "zh": "卡方检验",
            "en": "Chi-square test"
        }
    },
    "Clustering Analysis K-Means": {
        "class": ClusteringAnalysisKMeansApp,
        "description": {
            "zh": "聚类分析 K-Means",
            "en": "Clustering Analysis K-Means"
        }
    },
    # "Cochran's Q Test": {
    #     "class": CochransQTestApp,
    #     "description": {
    #         "zh": "Cochran's Q 检验分析",
    #         "en": "Cochran's Q Test Analysis"
    #     }
    # },
    "Collinearity Analysis (VIF)": {
        "class": CollinearityAnalysisVIFApp,
        "description": {
            "zh": "共线性分析 (VIF)",
            "en": "Collinearity Analysis (VIF)"
        }
    },
    "Composite Index Analysis": {
        "class": CompositeIndexAnalysisApp,
        "description": {
            "zh": "综合指数分析",
            "en": "Composite Index Analysis"
        }
    },
    "Conjoint Analysis": {
        "class": ConjointAnalysisApp,
        "description": {
            "zh": "联合分析",
            "en": "Conjoint Analysis"
        }
    },
    "Content Validity Analysis": {
        "class": ContentValidityAnalysisApp,
        "description": {
            "zh": "内容有效性分析",
            "en": "Content Validity Analysis"
        }
    },
    "Coupling Coordination Degree Model Analysis": {
        "class": CouplingCoordinationDegreeModelAnalysisApp,
        "description": {
            "zh": "耦合协调度模型分析",
            "en": "Coupling Coordination Degree Model Analysis"
        }
    },
    "Cramer-von Mises Test": {
        "class": CramerVonMisesTestApp,
        "description": {
            "zh": "Cramer-von Mises 检验",
            "en": "Cramer-von Mises Test"
        }
    },
    "CRITIC Weighting Method Analysis": {
        "class": CRITICWeightingMethodAnalysisApp,
        "description": {
            "zh": "CRITIC 权重法分析",
            "en": "CRITIC Weighting Method Analysis"
        }
    },
    "DAgostino-K-Squared Test": {
        "class": DAgostinoKSquaredTestApp,
        "description": {
            "zh": "DAgostino-K-Squared 检验",
            "en": "DAgostino-K-Squared Test"
        }
    },
    "Delphi Method Analysis": {
        "class": DelphiMethodAnalysisApp,
        "description": {
            "zh": "德尔菲专家法分析",
            "en": "Delphi Method Analysis"
        }
    },
    "DEMATEL Analysis": {
        "class": DEMATELAnalysisApp,
        "description": {
            "zh": "DEMATEL 分析",
            "en": "DEMATEL Analysis"
        }
    },
    "Density-Based Clustering Analysis": {
        "class": DensityBasedClusteringAnalysisApp,
        "description": {
            "zh": "密度聚类分析",
            "en": "Density-Based Clustering Analysis"
        }
    },
    "Descriptive Statistics": {
        "class": DescriptiveStatisticsApp,
        "description": {
            "zh": "描述性统计",
            "en": "Descriptive Statistics"
        }
    },
    "Discriminant Analysis": {
        "class": DiscriminantAnalysisApp,
        "description": {
            "zh": "判别分析",
            "en": "Discriminant Analysis"
        }
    },
    "Efficacy Coefficient Analysis": {
        "class": EfficacyCoefficientAnalysisApp,
        "description": {
            "zh": "功效系数分析",
            "en": "Efficacy Coefficient Analysis"
        }
    },
    "Entropy Method Analysis": {
        "class": EntropyMethodAnalysisApp,
        "description": {
            "zh": "熵权法分析",
            "en": "Entropy Method Analysis"
        }
    },
    "Exponential Smoothing Method Analysis": {
        "class": ExponentialSmoothingMethodAnalysisApp,
        "description": {
            "zh": "指数平滑法分析",
            "en": "Exponential Smoothing Method Analysis"
        }
    },
    "Factor Analysis": {
        "class": FactorAnalysisApp,
        "description": {
            "zh": "因子分析",
            "en": "Factor Analysis"
        }
    },
    "Friedman Test Analysis": {
        "class": FriedmanTestApp,
        "description": {
            "zh": "Friedman 检验分析",
            "en": "Friedman Test Analysis"
        }
    },
    "Fuzzy Analytic Hierarchy Process (FAHP) Analysis": {
        "class": FuzzyAnalyticHierarchyProcessFAHPApp,
        "description": {
            "zh": "模糊层次分析法 FAHP 分析",
            "en": "Fuzzy Analytic Hierarchy Process (FAHP) Analysis"
        }
    },
    "Fuzzy Comprehensive Evaluation Analysis": {
        "class": FuzzyComprehensiveEvaluationAnalysisApp,
        "description": {
            "zh": "模糊综合评价分析",
            "en": "Fuzzy Comprehensive Evaluation Analysis"
        }
    },
    "Generalized Estimating Equations Analysis": {
        "class": GeneralizedEstimatingEquationsAnalysisApp,
        "description": {
            "zh": "广义估计方程分析",
            "en": "Generalized Estimating Equations Analysis"
        }
    },
    "GMM Estimation Analysis": {
        "class": GMMEstimationAnalysisApp,
        "description": {
            "zh": "GMM 估计分析",
            "en": "GMM Estimation Analysis"
        }
    },
    "Gray Prediction Model Analysis": {
        "class": GrayPredictionModelAnalysisApp,
        "description": {
            "zh": "灰色预测模型分析",
            "en": "Gray Prediction Model Analysis"
        }
    },
    "Grey Relational Analysis": {
        "class": GreyRelationalAnalysisApp,
        "description": {
            "zh": "灰色关联分析",
            "en": "Grey Relational Analysis"
        }
    },
    "Hierarchical Clustering Analysis": {
        "class": HierarchicalClusteringAnalysisApp,
        "description": {
            "zh": "分层聚类分析",
            "en": "Hierarchical Clustering Analysis"
        }
    },
    "Hierarchical Regression Analysis": {
        "class": HierarchicalRegressionAnalysisApp,
        "description": {
            "zh": "层次回归分析",
            "en": "Hierarchical Regression Analysis"
        }
    },
    "Independence Weighting Method Analysis": {
        "class": IndependenceWeightingMethodAnalysisApp,
        "description": {
            "zh": "独立性权重法分析",
            "en": "Independence Weighting Method Analysis"
        }
    },
    "Independent Samples T-Test Analysis": {
        "class": IndependentSamplesTTestAnalysisApp,
        "description": {
            "zh": "独立样本 t 检验分析",
            "en": "Independent Samples T-Test Analysis"
        }
    },
    "Information Entropy Weight Method Analysis": {
        "class": InformationEntropyWeightMethodAnalysisApp,
        "description": {
            "zh": "信息量权重法分析",
            "en": "Information Entropy Weight Method Analysis"
        }
    },
    "Jarque-Bera Test": {
        "class": JarqueBeraTestApp,
        "description": {
            "zh": "Jarque-Bera 检验",
            "en": "Jarque-Bera Test"
        }
    },
    "KANO Model Analysis": {
        "class": KANOModelAnalysisApp,
        "description": {
            "zh": "KANO 模型分析",
            "en": "KANO Model Analysis"
        }
    },
    "Kappa Consistency Test": {
        "class": KappaConsistencyTestApp,
        "description": {
            "zh": "Kappa 一致性检验",
            "en": "Kappa Consistency Test"
        }
    },
    "Kendall Correlation Analysis": {
        "class": KendallCorrelationAnalysisApp,
        "description": {
            "zh": "Kendall 相关分析",
            "en": "Kendall Correlation Analysis"
        }
    },
    "Kendall's Coordination Coefficient": {
        "class": KendallsCoordinationCoefficientApp,
        "description": {
            "zh": "Kendall 协和系数分析",
            "en": "Kendall's Coordination Coefficient"
        }
    },
    "Kolmogorov-Smirnov Test": {
        "class": KolmogorovSmirnovTestApp,
        "description": {
            "zh": "Kolmogorov-Smirnov 检验",
            "en": "Kolmogorov-Smirnov Test"
        }
    },
    "Lasso Regression Analysis": {
        "class": LassoRegressionAnalysisApp,
        "description": {
            "zh": "Lasso 回归分析",
            "en": "Lasso Regression Analysis"
        }
    },
    "Levene Test": {
        "class": LeveneTestApp,
        "description": {
            "zh": "Levene 检验",
            "en": "Levene Test"
        }
    },
    "Lilliefors Test": {
        "class": LillieforsTestApp,
        "description": {
            "zh": "Lilliefors 检验",
            "en": "Lilliefors Test"
        }
    },
    "Linear Tobit Regression Analysis": {
        "class": LinearTobitRegressionAnalysisApp,
        "description": {
            "zh": "线性 Tobit 回归分析",
            "en": "Linear Tobit Regression Analysis"
        }
    },
    "Markov Prediction Analysis": {
        "class": MarkovPredictionAnalysisApp,
        "description": {
            "zh": "马尔可夫预测分析",
            "en": "Markov Prediction Analysis"
        }
    },
    "Mediation Analysis": {
        "class": MediationAnalysisApp,
        "description": {
            "zh": "中介作用分析",
            "en": "Mediation Analysis"
        }
    },
    "Moderated Mediation Analysis": {
        "class": ModeratedMediationAnalysisApp,
        "description": {
            "zh": "调节中介作用分析",
            "en": "Moderated Mediation Analysis"
        }
    },
    "Moderation Analysis": {
        "class": ModerationAnalysisApp,
        "description": {
            "zh": "调节作用分析",
            "en": "Moderation Analysis"
        }
    },
    "Multi-sample ANOVA": {
        "class": MultiSampleANOVAApp,
        "description": {
            "zh": "多样本方差分析",
            "en": "Multi-sample ANOVA"
        }
    },
    "Multidimensional Scaling Analysis": {
        "class": MultidimensionalScalingAnalysisApp,
        "description": {
            "zh": "多维缩放分析",
            "en": "Multidimensional Scaling Analysis"
        }
    },
    "Multidimensional Scaling (MDS) Analysis": {
        "class": MultidimensionalScalingMDSApp,
        "description": {
            "zh": "多维缩放 (MDS) 分析",
            "en": "Multidimensional Scaling (MDS) Analysis"
        }
    },
    "Multinomial Logistic Regression Analysis": {
        "class": MultinomialLogisticRegressionApp,
        "description": {
            "zh": "多项逻辑回归分析",
            "en": "Multinomial Logistic Regression Analysis"
        }
    },
    "Multinomial Logit Regression Analysis": {
        "class": MultinomialLogitRegressionApp,
        "description": {
            "zh": "多项 Logit 回归分析",
            "en": "Multinomial Logit Regression Analysis"
        }
    },
    "Multiple Choice Question Analysis": {
        "class": MultipleChoiceQuestionAnalysisApp,
        "description": {
            "zh": "多项选择题分析",
            "en": "Multiple Choice Question Analysis"
        }
    },
    "Multivariate Analysis of Variance (MANOVA)": {
        "class": MultivariateManovaApp,
        "description": {
            "zh": "多变量方差分析",
            "en": "Multivariate Analysis of Variance (MANOVA)"
        }
    },
    "NPS Net Promoter Score Analysis": {
        "class": NPSNetPromoterScoreAnalysisApp,
        "description": {
            "zh": "NPS 净推广者得分分析",
            "en": "NPS Net Promoter Score Analysis"
        }
    },
    "Obstacle Degree Model Analysis": {
        "class": ObstacleDegreeModelAnalysisApp,
        "description": {
            "zh": "障碍度模型分析",
            "en": "Obstacle Degree Model Analysis"
        }
    },
    "One-sample ANOVA": {
        "class": OneSampleANOVAApp,
        "description": {
            "zh": "单样本方差分析",
            "en": "One-sample ANOVA"
        }
    },
    "One-Sample t-Test Analysis": {
        "class": OneSampleTTestAnalysisApp,
        "description": {
            "zh": "单样本 t 检验分析",
            "en": "One-Sample t-Test Analysis"
        }
    },
    "One-Sample Wilcoxon Test Analysis": {
        "class": OneSampleWilcoxonTestAnalysisApp,
        "description": {
            "zh": "单样本Wilcoxon检验分析",
            "en": "One-Sample Wilcoxon Test Analysis"
        }
    },
    "Ordered Logit Regression Analysis": {
        "class": OrderedLogitRegressionAnalysisApp,
        "description": {
            "zh": "有序Logit回归分析",
            "en": "Ordered Logit Regression Analysis"
        }
    },
    "Ordinary Least Squares Linear Regression Analysis": {
        "class": OrdinaryLeastSquaresLinearRegressionAnalysisApp,
        "description": {
            "zh": "普通最小二乘线性回归分析",
            "en": "Ordinary Least Squares Linear Regression Analysis"
        }
    },
    "Paired t-test Analysis": {
        "class": PairedTTestAnalysisApp,
        "description": {
            "zh": "配对 t 检验分析",
            "en": "Paired t-test Analysis"
        }
    },
    "Paired-Sample Wilcoxon Test Analysis": {
        "class": PairedSampleWilcoxonTestAnalysisApp,
        "description": {
            "zh": "配对样本Wilcoxon检验分析",
            "en": "Paired-Sample Wilcoxon Test Analysis"
        }
    },
    "Partial Correlation Analysis": {
        "class": PartialCorrelationAnalysisApp,
        "description": {
            "zh": "偏相关分析",
            "en": "Partial Correlation Analysis"
        }
    },
    "Partial Least Squares (PLS) Analysis": {
        "class": PartialLeastSquaresRegressionAnalysisApp,
        "description": {
            "zh": "偏最小二乘 (PLS) 分析",
            "en": "Partial Least Squares (PLS) Analysis"
        }
    },
    "Pearson Correlation Analysis": {
        "class": PearsonCorrelationAnalysisApp,
        "description": {
            "zh": "皮尔逊相关分析",
            "en": "Pearson Correlation Analysis"
        }
    },
    "Polynomial Regression Analysis": {
        "class": PolynomialRegressionAnalysisApp,
        "description": {
            "zh": "多项式回归分析",
            "en": "Polynomial Regression Analysis"
        }
    },
    "Post-hoc Multiple Comparison Analysis": {
        "class": PostHocMultipleComparisonsApp,
        "description": {
            "zh": "事后多重比较分析",
            "en": "Post-hoc Multiple Comparison Analysis"
        }
    },
    "Price Sensitivity Meter Analysis": {
        "class": PriceSensitivityMeterAnalysisApp,
        "description": {
            "zh": "价格敏感度仪分析",
            "en": "Price Sensitivity Meter Analysis"
        }
    },
    "Principal Component Analysis": {
        "class": PrincipalComponentAnalysisApp,
        "description": {
            "zh": "主成分分析",
            "en": "Principal Component Analysis"
        }
    },
    "Range Analysis": {
        "class": RangeAnalysisApp,
        "description": {
            "zh": "极差分析",
            "en": "Range Analysis"
        }
    },
    "Rank-Sum Ratio (RSR) Analysis": {
        "class": RankSumRatioRSRAnalysisApp,
        "description": {
            "zh": "秩和比(RSR)分析",
            "en": "Rank-Sum Ratio (RSR) Analysis"
        }
    },
    "Regularized Binary Logistic Regression Analysis": {
        "class": RegularizedBinaryLogisticRegressionAnalysisApp,
        "description": {
            "zh": "正则化二元逻辑回归分析",
            "en": "Regularized Binary Logistic Regression Analysis"
        }
    },
    "Regularized Multinomial Logistic Regression Analysis": {
        "class": RegularizedMultinomialLogisticRegressionAnalysisApp,
        "description": {
            "zh": "正则化多项逻辑回归分析",
            "en": "Regularized Multinomial Logistic Regression Analysis"
        }
    },
    "Reliability Analysis": {
        "class": ReliabilityAnalysisApp,
        "description": {
            "zh": "信度分析",
            "en": "Reliability Analysis"
        }
    },
    "Reliability Test Analysis": {
        "class": ReliabilityTestAnalysisApp,
        "description": {
            "zh": "信度检验分析",
            "en": "Reliability Test Analysis"
        }
    },
    "Repeated Measures ANOVA": {
        "class": RepeatedMeasuresANOVAApp,
        "description": {
            "zh": "重复测量方差分析",
            "en": "Repeated Measures ANOVA"
        }
    },
    "Ridge Regression Analysis": {
        "class": RidgeRegressionAnalysisApp,
        "description": {
            "zh": "岭回归分析",
            "en": "Ridge Regression Analysis"
        }
    },
    "Robust Linear Regression Analysis": {
        "class": RobustLinearRegressionAnalysisApp,
        "description": {
            "zh": "稳健线性回归分析",
            "en": "Robust Linear Regression Analysis"
        }
    },
    "Runs Test": {
        "class": RunsTestApp,
        "description": {
            "zh": "游程检验分析",
            "en": "Runs Test Analysis"
        }
    },
    "Second Order Cluster Analysis": {
        "class": SecondOrderClusteringAnalysisApp,
        "description": {
            "zh": "二阶聚类分析",
            "en": "Second Order Cluster Analysis"
        }
    },
    "Shapiro-Wilk Test": {
        "class": ShapiroWilkTestApp,
        "description": {
            "zh": "Shapiro-Wilk 检验",
            "en": "Shapiro-Wilk Test"
        }
    },
    "Spearman Correlation Analysis": {
        "class": SpearmanCorrelationAnalysisApp,
        "description": {
            "zh": "Spearman相关分析",
            "en": "Spearman Correlation Analysis"
        }
    },
    "Split-Half Reliability Analysis": {
        "class": SplitHalfReliabilityAnalysisApp,
        "description": {
            "zh": "半样本信度分析",
            "en": "Split-Half Reliability Analysis"
        }
    },
    "Stepwise Regression Analysis": {
        "class": StepwiseRegressionAnalysisApp,
        "description": {
            "zh": "逐步回归分析",
            "en": "Stepwise Regression Analysis"
        }
    },
    "Test-Retest Reliability Analysis": {
        "class": TestRetestReliabilityAnalysisApp,
        "description": {
            "zh": "重测信度分析",
            "en": "Test-Retest Reliability Analysis"
        }
    },
    "TOPSIS Method Analysis": {
        "class": TOPSISMethodAnalysisApp,
        "description": {
            "zh": "TOPSIS 法分析",
            "en": "TOPSIS Method Analysis"
        }
    },
    "Turf Combination Model Analysis": {
        "class": TurfCombinationModelAnalysisApp,
        "description": {
            "zh": "Turf组合模型分析",
            "en": "Turf Combination Model Analysis"
        }
    },
    "Two-sample ANOVA": {
        "class": TwoSampleANOVAApp,
        "description": {
            "zh": "双样本方差分析",
            "en": "Two-sample ANOVA"
        }
    },
    # "Undesirable SBM Model Analysis": {
    #     "class": UndesirableSBMModelAnalysisApp,
    #     "description": {
    #         "zh": "非期望SBM模型分析",
    #         "en": "Undesirable SBM Model Analysis"
    #     }
    # },
    "Validity Analysis": {
        "class": ValidityAnalysisApp,
        "description": {
            "zh": "效度分析",
            "en": "Validity Analysis"
        }
    },
    "Within-Group Inter-Rater Reliability rwg Analysis": {
        "class": WithinGroupInterRaterReliabilityRwgAnalysisApp,
        "description": {
            "zh": "组内评分者信度rwg分析",
            "en": "Within-Group Inter-Rater Reliability rwg Analysis"
        }
    },
    # 可以继续添加其他模块
    # "Module Name": {
    #     "class": ModuleClass,
    #     "description": {
    #         "zh": "中文描述",
    #         "en": "English Description"
    #     }
    # },
}

def on_mousewheel(event):
    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

def center_button_frame():
    # 更新Canvas的滚动区域
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

    # 更新Canvas中窗口的位置
    canvas.coords(canvas.find_all()[0], (x_offset, y_offset))

class AnalyzerApp:
    def __init__(self, root=None):
        # 当前语言
        self.current_language = 'en'

        # 获取当前脚本所在的目录
        self.project_dir = os.path.dirname(os.path.abspath(__file__))

        # 使用模块映射表替代文件扫描
        self.modules = MODULE_MAP

        # 如果没有提供root，则创建一个新窗口
        if root is None:
            self.root = ttk.Window(themename="flatly")
        else:
            self.root = root
        self.root.title(LANGUAGES[self.current_language]["title"])

        self.create_ui()

    def open_module(self, module_name):
        try:
            # 使用映射表中的类创建应用实例
            module_class = self.modules[module_name]["class"]
            module_class(ttk.Toplevel(self.root))
        except Exception as e:
            self.result_label.config(text=LANGUAGES[self.current_language]['error_message'].format(module_name, e))

    def switch_language(self):
        self.current_language = 'en' if self.current_language == 'zh' else 'zh'
        self.root.title(LANGUAGES[self.current_language]["title"])
        self.language_label.config(text=LANGUAGES[self.current_language]['switch_language'])
        self.search_entry.delete(0, ttk.END)
        self.search_entry.insert(0, LANGUAGES[self.current_language]['search_placeholder'])
        self.search_entry.config(foreground='gray')

        # 更新按钮文本为当前语言
        for button, module_name in zip(self.button_list, self.modules.keys()):
            button_text = self.modules[module_name]["description"][self.current_language]
            button.config(text=button_text)
            button.configure(bootstyle=PRIMARY)

    def search_scripts(self, event=None):
        keyword = self.search_entry.get().strip()
        for button, module_name in zip(self.button_list, self.modules.keys()):
            button_text = button.cget("text")
            if keyword and keyword.lower() in button_text.lower():
                button.configure(bootstyle="danger")
            else:
                button.configure(bootstyle=PRIMARY)

    def on_entry_click(self, event):
        if self.search_entry.get() == LANGUAGES[self.current_language]['search_placeholder']:
            self.search_entry.delete(0, ttk.END)
            self.search_entry.config(foreground='black')

    def on_focusout(self, event):
        if not self.search_entry.get():
            self.search_entry.insert(0, LANGUAGES[self.current_language]['search_placeholder'])
            self.search_entry.config(foreground='gray')

    def create_ui(self):
        global canvas, button_frame

        # 获取屏幕的宽度和高度
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # 设置窗口的宽度和高度
        window_width = 1250
        window_height = 600

        # 计算窗口应该放置的位置
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        # 设置窗口的位置和大小
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # 创建一个主框架，用于居中内容
        main_frame = ttk.Frame(self.root)
        main_frame.pack(expand=True, fill=BOTH)

        # 创建搜索框
        self.search_entry = ttk.Entry(main_frame)
        self.search_entry.insert(0, LANGUAGES[self.current_language]['search_placeholder'])
        self.search_entry.config(foreground='gray')
        self.search_entry.pack(pady=10, padx=10, fill=X)
        self.search_entry.bind("<KeyRelease>", self.search_scripts)
        self.search_entry.bind("<FocusIn>", self.on_entry_click)
        self.search_entry.bind("<FocusOut>", self.on_focusout)

        # 创建一个Canvas组件
        canvas = ttk.Canvas(main_frame)
        canvas.pack(side=LEFT, fill=BOTH, expand=True)

        # 创建垂直滚动条
        scrollbar = ttk.Scrollbar(main_frame, command=canvas.yview)
        scrollbar.pack(side=RIGHT, fill=Y)

        # 配置Canvas的滚动条
        canvas.configure(yscrollcommand=scrollbar.set)

        # 创建一个框架来放置按钮
        button_frame = ttk.Frame(canvas)

        # 将按钮框架添加到Canvas中
        canvas.create_window((0, 0), window=button_frame, anchor=NW)

        # 存储所有按钮的列表
        self.button_list = []

        # 创建按钮
        col = 0
        row = 0
        for module_name in self.modules.keys():
            # 根据当前语言获取按钮文本
            button_text = self.modules[module_name]["description"][self.current_language]
            button = ttk.Button(button_frame, text=button_text,
                               command=lambda m=module_name: self.open_module(m),
                               bootstyle=PRIMARY)
            button.grid(row=row, column=col, padx=5, pady=5)
            self.button_list.append(button)
            col += 1
            if col == 3:
                col = 0
                row += 1

        # 初始居中按钮框架
        center_button_frame()

        # 绑定窗口大小改变事件，重新居中按钮框架
        self.root.bind("<Configure>", lambda event: center_button_frame())

        # 绑定鼠标滚轮事件
        canvas.bind_all("<MouseWheel>", on_mousewheel)

        # 创建语言切换标签
        self.language_label = ttk.Label(self.root, text=LANGUAGES[self.current_language]['switch_language'], cursor="hand2", foreground="gray")
        self.language_label.pack(pady=10)
        self.language_label.bind("<Button-1>", lambda event: self.switch_language())

        # 创建结果显示标签
        self.result_label = ttk.Label(self.root, text="", justify=LEFT)
        self.result_label.pack(pady=10)

    def run(self):
        # 运行主循环
        self.root.mainloop()

# 为了向后兼容，保留原来的运行方式
def run_app():
    app = AnalyzerApp()
    app.run()

if __name__ == "__main__":
    run_app()