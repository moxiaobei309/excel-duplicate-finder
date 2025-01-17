import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

class ExcelDuplicateFinderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 重复数据筛选器")
        self.root.geometry("500x300")
        
        # 文件选择
        self.file_label = tk.Label(root, text="选择Excel文件:")
        self.file_label.pack(pady=5)
        
        self.file_entry = tk.Entry(root, width=50)
        self.file_entry.pack(pady=5)
        
        self.browse_button = tk.Button(root, text="浏览", command=self.browse_file)
        self.browse_button.pack(pady=5)
        
        # 列选择
        self.column_label = tk.Label(root, text="选择要筛选的列:")
        self.column_label.pack(pady=5)
        
        self.column_combobox = ttk.Combobox(root, state="readonly")
        self.column_combobox.pack(pady=5)
        
        # 操作按钮
        self.find_button = tk.Button(root, text="查找重复数据", command=self.find_duplicates)
        self.find_button.pack(pady=20)
        
    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)
            self.load_columns(file_path)
            
    def load_columns(self, file_path):
        try:
            df = pd.read_excel(file_path)
            self.column_combobox['values'] = list(df.columns)
            self.column_combobox.current(0)
        except Exception as e:
            messagebox.showerror("错误", f"无法读取文件: {str(e)}")
            
    def find_duplicates(self):
        file_path = self.file_entry.get()
        column = self.column_combobox.get()
        
        if not file_path:
            messagebox.showwarning("警告", "请先选择文件")
            return
            
        if not column:
            messagebox.showwarning("警告", "请选择要筛选的列")
            return
            
        try:
            # 分块读取Excel文件
            chunksize = 10**6  # 每次读取1百万行
            chunks = []
            for chunk in pd.read_excel(file_path, chunksize=chunksize):
                chunks.append(chunk)
            df = pd.concat(chunks)
            
            # 检查所选列是否存在
            if column not in df.columns:
                messagebox.showerror("错误", f"列 '{column}' 不存在")
                return
                
            # 查找重复数据
            duplicates = df[df.duplicated(column, keep=False)]
            
            # 释放内存
            del chunks
            del df
            
            if duplicates.empty:
                messagebox.showinfo("结果", f"在列 '{column}' 中未找到重复数据")
            else:
                # 显示结果预览
                preview_window = tk.Toplevel(self.root)
                preview_window.title("重复数据预览")
                
                text = tk.Text(preview_window, wrap=tk.NONE)
                text.insert(tk.END, duplicates.to_string())
                text.config(state=tk.DISABLED)
                text.pack(fill=tk.BOTH, expand=True)
                
                # 添加保存按钮
                save_button = tk.Button(
                    preview_window,
                    text="保存结果",
                    command=lambda: self.save_results(duplicates)
                )
                save_button.pack(pady=10)
                
        except pd.errors.EmptyDataError:
            messagebox.showerror("错误", "文件为空或格式不正确")
        except pd.errors.ParserError:
            messagebox.showerror("错误", "无法解析文件内容")
        except Exception as e:
            messagebox.showerror("错误", f"处理文件时出错: {str(e)}")
            
    def save_results(self, duplicates):
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="保存重复数据结果"
        )
        if save_path:
            try:
                duplicates.to_excel(save_path, index=False)
                messagebox.showinfo("成功", f"已保存重复数据到: {save_path}")
            except Exception as e:
                messagebox.showerror("错误", f"保存文件时出错: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelDuplicateFinderApp(root)
    root.mainloop()
