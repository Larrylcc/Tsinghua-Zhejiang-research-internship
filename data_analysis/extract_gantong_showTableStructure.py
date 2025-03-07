import docx
import os

# 指向包含多个docx文件的目录
base_dir = "/Users/larry/Desktop/长三院儿童康复数据（去重）/以诺儿童培智服务中心/感统测评报告 (6).docx"

# 检查路径是否存在
if not os.path.exists(base_dir):
    print("错误：指定的路径不存在: {}".format(base_dir))
    exit(1)

# 如果是目录，则处理目录中的所有docx文件
if os.path.isdir(base_dir):
    docx_files = [f for f in os.listdir(base_dir) if f.endswith('.docx')]
    if not docx_files:
        print("目录中没有找到.docx文件")
        exit(1)

    print("在目录中找到 {} 个docx文件".format(len(docx_files)))

    for docx_file in docx_files:
        docx_path = os.path.join(base_dir, docx_file)
        print("\n处理文件: {}".format(docx_file))

        try:
            document = docx.Document(docx_path)
            for table in document.tables:
                for row_index, row in enumerate(table.rows):
                    for col_index, cell in enumerate(row.cells):
                        print('行列为：({},{})'.format(row_index, col_index))
                        print('单元格内容为：{}'.format(cell.text))
        except Exception as e:
            print("处理文件 {} 时出错: {}".format(docx_file, str(e)))
else:
    # 如果是一个文件，确保它是.docx文件
    if not base_dir.endswith('.docx'):
        print("错误：文件 {} 不是.docx文件".format(base_dir))
        exit(1)

    try:
        document = docx.Document(base_dir)
        for table in document.tables:
            for row_index, row in enumerate(table.rows):
                for col_index, cell in enumerate(row.cells):
                    print('行列为：({},{})'.format(row_index, col_index))
                    print('单元格内容为：{}'.format(cell.text))
    except Exception as e:
        print("处理文件时出错: {}".format(str(e)))
