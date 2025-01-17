# Excel 重复数据筛选器

## 项目简介
这是一个用于查找Excel文件中重复数据的图形界面应用程序。它可以帮助用户快速识别和提取指定列中的重复记录。

## 主要功能
- 选择Excel文件
- 自动加载文件列名
- 选择要筛选的列
- 查找重复数据
- 预览重复记录
- 保存筛选结果

## 使用方法
1. 安装依赖：
   ```bash
   pip install pandas openpyxl
   ```

2. 运行程序：
   ```bash
   python excel_duplicate_finder_app.py
   ```

3. 使用步骤：
   - 点击"浏览"按钮选择Excel文件
   - 从下拉列表中选择要筛选的列
   - 点击"查找重复数据"按钮
   - 查看结果预览
   - 点击"保存结果"按钮保存筛选结果

## 依赖说明
- Python 3.x
- pandas
- openpyxl
- tkinter（Python标准库）

## 注意事项
1. 仅支持.xlsx格式的Excel文件
2. 文件大小限制取决于可用内存，建议不超过1GB
3. 程序采用分块读取技术，可有效处理大文件
4. 如果遇到内存不足问题，请确保系统有足够可用内存
5. 保存结果时请确保有足够的磁盘空间

## 文件说明
- `excel_duplicate_finder_app.py`：主程序
- `generate_sample.py`：示例数据生成脚本
- `find_duplicates.py`：命令行版本筛选脚本
- `sample_data.xlsx`：示例数据文件
