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
df = pd.read_excel(file_path, sheet_name='Sheet2')

# 数据预处理
# 确保必要列存在
necessary_cols = ['儿童姓名', '测评日期', '综合得分']
if not all(col in df.columns for col in necessary_cols):
    missing = [col for col in necessary_cols if col not in df.columns]
    print(f"缺少必要列：{missing}")
    exit()

# 转换日期格式并清理数据
df['测评日期'] = pd.to_datetime(df['测评日期'])
df.dropna(subset=['综合得分', '测评日期'], inplace=True)

# 创建输出目录
output_dir = '/Users/larry/Desktop/长三院儿童康复数据（去重）/集成数据/分析结果'
os.makedirs(output_dir, exist_ok=True)

# 初始化结果列表
results = []

# 按儿童分组分析
for name, group in df.groupby('儿童姓名'):
    if len(group) < 2:  # 至少需要两次测评
        continue

    # 按时间排序
    group = group.sort_values('测评日期')

    # 计算综合得分变化
    first = group['综合得分'].iloc[0]
    last = group['综合得分'].iloc[-1]
    slope = stats.linregress(np.arange(len(group)), group['综合得分'])[0]

    results.append({
        '儿童姓名': name,
        '测评次数': len(group),
        '首次得分': first,
        '末次得分': last,
        '变化值': last - first,
        '变化率': (last - first)/first if first !=0 else np.nan,
        '时间斜率': slope,
        '首次日期': group['测评日期'].iloc[0].strftime('%Y-%m-%d'),
        '末次日期': group['测评日期'].iloc[-1].strftime('%Y-%m-%d'),
        '测评数据': group[['测评日期', '综合得分']].values.tolist()
    })

# 转换分析结果
results_df = pd.DataFrame(results)
results_df.sort_values('变化值', ascending=False, inplace=True)

# 保存结果
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
output_path = f"{output_dir}/康复效果分析_{timestamp}.xlsx"
results_df[['儿童姓名', '测评次数', '首次得分', '末次得分', '变化值',
           '变化率', '时间斜率', '首次日期', '末次日期']].to_excel(output_path, index=False)

# ===== 可视化分析 =====

# 1. 综合得分趋势图（前10名改善最明显的儿童）
plt.figure(figsize=(14, 8))
top_children = results_df.head(10)['儿童姓名'].values

for name in top_children:
    child_data = df[df['儿童姓名'] == name].sort_values('测评日期')
    plt.plot(child_data['测评日期'], child_data['综合得分'],
            marker='o', linestyle='-', linewidth=2,
            label=f"{name}")

plt.title('综合得分康复趋势（改善最明显的10名儿童）', fontsize=14)
plt.xlabel('测评日期', fontsize=12)
plt.ylabel('综合得分', fontsize=12)
plt.xticks(rotation=45)
plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
plt.grid(True, linestyle='--', alpha=0.7)
plt.tight_layout()

# 保存趋势图
trend_path = f"{output_dir}/康复趋势图_{timestamp}.png"
plt.savefig(trend_path, bbox_inches='tight', dpi=300)
plt.close()

# 2. 综合得分变化分布
plt.figure(figsize=(12, 10))

# 变化值分布
plt.subplot(2, 1, 1)
sns.histplot(results_df['变化值'], kde=True)
plt.axvline(x=0, color='r', linestyle='--')
plt.title('综合得分变化值分布', fontsize=14)
plt.xlabel('得分变化值', fontsize=12)
plt.ylabel('儿童数量', fontsize=12)

# 变化率分布
plt.subplot(2, 1, 2)
sns.boxplot(y=results_df['变化值'])
plt.title('综合得分变化箱线图', fontsize=14)
plt.ylabel('得分变化值', fontsize=12)

plt.tight_layout()
dist_path = f"{output_dir}/变化分布图_{timestamp}.png"
plt.savefig(dist_path, bbox_inches='tight', dpi=300)
plt.close()

# 3. 前后对比散点图
plt.figure(figsize=(10, 8))
plt.scatter(results_df['首次得分'], results_df['末次得分'], alpha=0.7)
plt.plot([0, 100], [0, 100], 'r--')  # 对角线
plt.title('首次vs末次测评得分对比', fontsize=14)
plt.xlabel('首次测评得分', fontsize=12)
plt.ylabel('末次测评得分', fontsize=12)
plt.grid(True, linestyle='--', alpha=0.7)
plt.tight_layout()

scatter_path = f"{output_dir}/首末对比图_{timestamp}.png"
plt.savefig(scatter_path, bbox_inches='tight', dpi=300)
plt.close()

# 统计检验（配对t检验）
t_stat, p_value = stats.ttest_rel(results_df['末次得分'], results_df['首次得分'])
stat_results = [{
    '分析指标': '综合得分',
    '平均变化值': results_df['变化值'].mean(),
    '中位数变化': results_df['变化值'].median(),
    '标准差': results_df['变化值'].std(),
    't值': t_stat,
    'p值': p_value,
    '有效样本量': len(results_df),
    '改善人数': sum(results_df['变化值'] > 0),
    '改善比例': sum(results_df['变化值'] > 0) / len(results_df)
}]

# 保存统计结果
stat_df = pd.DataFrame(stat_results)
stat_path = f"{output_dir}/统计分析结果_{timestamp}.xlsx"
stat_df.to_excel(stat_path, index=False)

print(f"""
分析完成！结果已保存至：
- 个体分析结果：{output_path}
- 康复趋势图：{trend_path}
- 变化分布图：{dist_path}
- 首末对比图：{scatter_path}
- 统计分析结果：{stat_path}

统计结果摘要：
- 有效样本数：{len(results_df)}名儿童
- 平均变化值：{results_df['变化值'].mean():.2f}
- 改善人数比例：{sum(results_df['变化值'] > 0)/len(results_df):.1%}
- 统计显著性：p值={p_value:.4f} {'(显著)' if p_value < 0.05 else '(不显著)'}
""")
