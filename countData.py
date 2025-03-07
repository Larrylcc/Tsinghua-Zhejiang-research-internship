import os
import pandas as pd
import re

# 设定主文件夹路径（请修改为你的实际路径）
main_folder = "/Users/larry/Desktop/长三院儿童康复数据（去重）"

# 设定测评文件类型关键词
categories = [
    "0~6岁儿童发育行为评估报告测评报告", "粗大运动能力(8大项)测评报告", "大运动评估",
    "感统测评报告", "感知觉能力(8大项)测评报告", "感知觉评估",
    "孤独症儿童心理教育评估测评详情", "精细运动能力(8大项)测评报告", "精细运动评估",
    "口部运动测评报告", "情绪行为评估", "情绪与行为能力评估(8大项)测评报告",
    "认知发展能力(8大项)测评报告", "认知能力评估", "社会交往能力(8大项)测评报告",
    "社会交往能力评估", "生活自理能力评估", "生活自理能力评估(8大项)测评报告",
    "心智障碍个别化教育课程评量表", "言语构音测评报告", "语言沟通能力评估表",
    "语言与沟通(8大项)测评报告", "韵母构音测评报告", "注意力操作测评测评报告",
    "注意力问卷测评测评报告", "注意力综合测评测评报告"
]

# 统计结果字典
results = {"测评类型": [], "总数": []}
org_list = []

data = {category: {} for category in categories}

# 遍历机构文件夹
for org in os.listdir(main_folder):
    org_path = os.path.join(main_folder, org)
    if not os.path.isdir(org_path):
        continue

    org_list.append(org)
    for category in categories:
        data[category][org] = 0

    # 遍历文件
    for file in os.listdir(org_path):
        for category in categories:
            if re.search(re.escape(category), file):
                data[category][org] += 1
                break

# 组织数据
for category, counts in data.items():
    results["测评类型"].append(category)
    results["总数"].append(sum(counts.values()))
    for org in org_list:
        if org not in results:
            results[org] = []
        results[org].append(counts.get(org, 0))

# 生成 DataFrame 并导出 Excel
output_file = "测评文件统计.xlsx"
df = pd.DataFrame(results)
df.to_excel(output_file, index=False)
print(f"统计完成，结果已保存至 {output_file}")
