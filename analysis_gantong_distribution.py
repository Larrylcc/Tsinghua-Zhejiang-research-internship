import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

# 读取Excel文件
file_path = '/Users/larry/Desktop/长三院儿童康复数据（去重）/集成数据/感统测评.xlsx'
sheet_name = 'Sheet1'  # 根据实际工作表名称修改

plt.rcParams['font.sans-serif']=['Songti SC']
plt.rcParams['axes.unicode_minus']=False
try:
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    print("文件读取成功！")
except Exception as e:
    print(f"文件读取失败，错误信息：{str(e)}")
    exit()

# 筛选所有得分列（根据实际列名特征调整）
score_columns = [col for col in df.columns if '得分' in col]

# 检查是否找到得分列
if not score_columns:
    print("未找到包含'得分'的列，请检查列名")
    exit()

# 统计计算
stats = pd.DataFrame({
    '样本量': df[score_columns].count(),
    '最大值': df[score_columns].max(),
    '最小值': df[score_columns].min(),
    '缺失值': df[score_columns].isnull().sum(),
    '中位数': df[score_columns].median(),
    '唯一值数量': df[score_columns].nunique(),
    '平均值': df[score_columns].mean().round(2),
    '方差': df[score_columns].var().round(2),
    '标准差': df[score_columns].std().round(2),
})
stats['变异系数'] = (stats['标准差'] / stats['平均值']).round(2)

# 绘制分布图
n_cols = 2
n_rows = int(np.ceil(len(score_columns) / n_cols))
plt.figure(figsize=(18, 5*n_rows))

# 修改后的绘图代码部分
# 在字体设置后添加全局字号配置（约第8行）
plt.rcParams['font.size'] = 20  # 全局字体大小
plt.rcParams['axes.titlesize'] = 20  # 子图标题字号
plt.rcParams['axes.labelsize'] = 20  # 坐标轴标签字号

# 修改绘图部分（约第45行开始）
for idx, col in enumerate(score_columns, 1):
    plt.subplot(n_rows, n_cols, idx)
    sns.histplot(
        data=df,
        x=col,
        kde=True,
        bins=20,
        color='skyblue',
        edgecolor='black'
    )
    plt.title(f'{col} 分布', fontsize=20)  # 调大子图标题
    plt.xlabel('得分值', fontsize=20)      # 调大X轴标签
    plt.ylabel('频数', fontsize=20)       # 调大Y轴标签
    plt.xticks(fontsize=20)              # X轴刻度字号
    plt.yticks(fontsize=20)              # Y轴刻度字号
    plt.grid(alpha=0.3)

plt.tight_layout()
plt.suptitle('各测评项得分分布', y=1.02, fontsize=24)  # 调大主标题

# 新增保存图片功能
save_path = '/Users/larry/Desktop/长三院儿童康复数据（去重）/集成数据/gantong_distribution.png'
plt.savefig(save_path, bbox_inches='tight', dpi=300)
print(f"\n图表已保存至：{save_path}")

plt.show()

# 输出统计结果
print("\n统计指标汇总：")
print(stats.to_string())  # 使用to_string()保持表格格式

# 可选：保存统计结果到Excel
# stats.to_excel('统计结果.xlsx')