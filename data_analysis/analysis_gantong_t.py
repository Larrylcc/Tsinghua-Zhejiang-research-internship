import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from scipy import stats
import os
from datetime import datetime

# 配置中文字体
plt.rcParams['font.sans-serif'] = ['Songti SC']
plt.rcParams['axes.unicode_minus'] = False

# 读取数据
file_path = '/Users/larry/Desktop/长三院儿童康复数据（去重）/集成数据/感统测评.xlsx'
df = pd.read_excel(file_path, sheet_name='Sheet1')

# 定义分析参数
score_columns = [col for col in df.columns if '得分' in col]
group_vars = {
    'gender': {'col': '性别', 'groups': ['男', '女']},
    'age': {'col': '年龄', 'bins': [0, 3, 6, 100], 'labels': ['0-3岁', '4-6岁', '7岁以上']},
    'agency': {'col': '测评机构', 'groups': None}  # 自动获取所有机构
}

# 准备结果存储
results = []

# 创建输出目录
output_dir = '/Users/larry/Desktop/长三院儿童康复数据（去重）/集成数据'
os.makedirs(output_dir, exist_ok=True)

def perform_analysis(df, score_col, group_col, group_name):
    """执行统计检验并返回结果"""
    groups = df.groupby(group_col)[score_col].apply(list)

    # 过滤有效数据
    groups = {k: v for k, v in groups.items() if len(v) > 1}
    if len(groups) < 2:
        return None

    # 选择检验方法
    if len(groups) == 2:
        stat, p = stats.ttest_ind(*groups.values())
        test_type = 't-test'
    else:
        stat, p = stats.f_oneway(*groups.values())
        test_type = 'ANOVA'

    return {
        'score_col': score_col,
        'group_type': group_name,
        'groups': ','.join(map(str, groups.keys())),
        'test_type': test_type,
        'statistic': round(stat, 3),
        'p_value': round(p, 5)
    }

# 对每个得分项进行分析
for score_col in score_columns:
    # 性别分析
    if '性别' in df.columns:
        res = perform_analysis(df, score_col, '性别', 'gender')
        if res: results.append(res)

    # 年龄段分析
    if '年龄' in df.columns:
        df['age_group'] = pd.cut(df['年龄'],
                               bins=group_vars['age']['bins'],
                               labels=group_vars['age']['labels'])
        res = perform_analysis(df, score_col, 'age_group', 'age')
        if res: results.append(res)

    # 测评机构分析
    if '测评机构' in df.columns:
        res = perform_analysis(df, score_col, '测评机构', 'agency')
        if res: results.append(res)

# 转换结果为DataFrame
results_df = pd.DataFrame(results)

# 保存分析结果
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
output_path = f"{output_dir}/statistical_analysis_{timestamp}.xlsx"
results_df.to_excel(output_path, index=False)
print(f"分析结果已保存至：{output_path}")

# 可视化显著结果
significant_df = results_df[results_df['p_value'] < 0.05]
if not significant_df.empty:
    plt.figure(figsize=(12, 8))
    heatmap_data = significant_df.pivot_table(index='score_col',
                                            columns='group_type',
                                            values='p_value')
    sns.heatmap(heatmap_data, annot=True, cmap='YlGnBu', cbar_kws={'label': 'p值'})
    plt.title('显著差异结果分布 (p < 0.05)', fontsize=14)
    plt.xlabel('分组类型', fontsize=12)
    plt.ylabel('测评项目', fontsize=12)

    img_path = f"{output_dir}/significant_results_{timestamp}.png"
    plt.savefig(img_path, bbox_inches='tight', dpi=300)
    plt.close()
    print(f"可视化结果已保存至：{img_path}")
else:
    print("未发现显著差异结果 (p < 0.05)")

print("\n分析完成！")
