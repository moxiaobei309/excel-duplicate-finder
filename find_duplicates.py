import pandas as pd

# 读取Excel文件
df = pd.read_excel('sample_data.xlsx')

# 查找Name列的重复记录
duplicates = df[df.duplicated('Name', keep=False)]

# 输出结果
print("重复记录：")
print(duplicates)

# 保存结果到新文件
duplicates.to_excel('duplicates_result.xlsx', index=False)
