import os
import hashlib
from collections import defaultdict

# 计算文件的哈希值
def get_file_hash(file_path):
    hasher = hashlib.md5()  # 也可以使用 hashlib.sha256()
    with open(file_path, 'rb') as f:
        while chunk := f.read(8192):  # 逐块读取文件，适用于大文件
            hasher.update(chunk)
    return hasher.hexdigest()

# 扫描文件夹，查找重复的 docx 文件
def find_duplicate_docx(folder_path):
    hash_dict = defaultdict(list)

    # 遍历所有 docx 文件
    for root, _, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.docx'):
                file_path = os.path.join(root, file)
                file_hash = get_file_hash(file_path)
                hash_dict[file_hash].append(file_path)

    # 找到重复文件
    duplicates = {hash_val: paths for hash_val, paths in hash_dict.items() if len(paths) > 1}
    return duplicates

# 删除重复文件，只保留一份
def remove_duplicates(duplicates):
    for file_list in duplicates.values():
        for duplicate_file in file_list[1:]:  # 保留第一个文件，删除其他
            os.remove(duplicate_file)
            print(f"Deleted: {duplicate_file}")

# 指定要查找的文件夹
folder_path = "/Users/larry/Desktop/长三院儿童康复数据（去重）"  # 这里替换为你的文件夹路径

# 运行查找和删除
duplicates = find_duplicate_docx(folder_path)
remove_duplicates(duplicates)

print("Duplicate docx files removed successfully!")
