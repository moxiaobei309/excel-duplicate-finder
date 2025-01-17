import pandas as pd

# 创建示例数据
data = {
    'ID': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 1, 2, 11, 12, 13, 14, 15, 16, 17, 18],
    'Name': ['Alice', 'Bob', 'Charlie', 'David', 'Eve', 'Frank', 'Grace', 'Heidi', 'Ivan', 'Judy',
             'Alice', 'Bob', 'Kevin', 'Linda', 'Mike', 'Nancy', 'Oscar', 'Peter', 'Quincy', 'Rachel'],
    'Age': [25, 30, 35, 40, 45, 50, 55, 60, 65, 70, 25, 30, 28, 32, 34, 36, 38, 40, 42, 44],
    'Department': ['HR', 'IT', 'Finance', 'IT', 'HR', 'Finance', 'IT', 'HR', 'Finance', 'IT',
                   'HR', 'IT', 'Marketing', 'Sales', 'Marketing', 'Sales', 'Marketing', 'Sales', 'Marketing', 'Sales']
}

# 创建DataFrame
df = pd.DataFrame(data)

# 保存为Excel文件
df.to_excel('sample_data.xlsx', index=False)
