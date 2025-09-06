import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import os

class ExcelCompareTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel数据对比工具")
        self.root.geometry("1000x750")
        
        # 数据存储
        self.df1 = None
        self.df2 = None
        self.file1_path = ""
        self.file2_path = ""
        
        self.setup_ui()
    
    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 文件1选择
        ttk.Label(file_frame, text="文件1:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.file1_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file1_var, width=50).grid(row=0, column=1, padx=(0, 5))
        ttk.Button(file_frame, text="浏览", command=lambda: self.select_file(1)).grid(row=0, column=2)
        
        # 文件2选择
        ttk.Label(file_frame, text="文件2:").grid(row=1, column=0, sticky=tk.W, padx=(0, 5), pady=(5, 0))
        self.file2_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file2_var, width=50).grid(row=1, column=1, padx=(0, 5), pady=(5, 0))
        ttk.Button(file_frame, text="浏览", command=lambda: self.select_file(2)).grid(row=1, column=2, pady=(5, 0))
        
        # 对比设置区域
        compare_frame = ttk.LabelFrame(main_frame, text="对比设置", padding="10")
        compare_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 列选择
        ttk.Label(compare_frame, text="对比列:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.column_var = tk.StringVar()
        self.column_combo = ttk.Combobox(compare_frame, textvariable=self.column_var, width=20)
        self.column_combo.grid(row=0, column=1, padx=(0, 10))
        
        # 对比类型
        ttk.Label(compare_frame, text="对比类型:").grid(row=0, column=2, sticky=tk.W, padx=(10, 5))
        self.compare_type = tk.StringVar(value="相同元素")
        compare_type_combo = ttk.Combobox(compare_frame, textvariable=self.compare_type, 
                                        values=["相同元素", "文件1独有", "文件2独有", "所有差异", "1与2独有"], 
                                        width=15, state="readonly")
        compare_type_combo.grid(row=0, column=3)
        
        # 数据验证选项
        ttk.Label(compare_frame, text="数据验证:").grid(row=1, column=0, sticky=tk.W, padx=(0, 5), pady=(5, 0))
        self.validate_data = tk.BooleanVar(value=True)
        ttk.Checkbutton(compare_frame, text="启用全行数据验证（防止同名不同值）", 
                       variable=self.validate_data).grid(row=1, column=1, columnspan=3, sticky=tk.W, pady=(5, 0))
        
        # 联系电话选项
        ttk.Label(compare_frame, text="显示选项:").grid(row=2, column=0, sticky=tk.W, padx=(0, 5), pady=(5, 0))
        self.show_phone = tk.BooleanVar(value=False)
        ttk.Checkbutton(compare_frame, text="显示联系电话", 
                       variable=self.show_phone).grid(row=2, column=1, sticky=tk.W, pady=(5, 0))
        
        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=(0, 10))
        
        ttk.Button(button_frame, text="加载文件", command=self.load_files).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="开始对比", command=self.compare_data).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="清空结果", command=self.clear_results).pack(side=tk.LEFT)
        
        # 结果显示区域
        result_frame = ttk.LabelFrame(main_frame, text="对比结果", padding="10")
        result_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 结果文本框 - 使用系统兼容字体，设置为可选择状态
        self.result_text = ScrolledText(result_frame, height=20, width=80, font=('Consolas', 12), state='normal')
        self.result_text.pack(fill=tk.BOTH, expand=True)
        # 绑定右键菜单以支持复制
        self.result_text.bind('<Button-3>', self.show_context_menu)
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)
        file_frame.columnconfigure(1, weight=1)
        
    def show_context_menu(self, event):
        """显示右键菜单"""
        try:
            context_menu = tk.Menu(self.root, tearoff=0)
            context_menu.add_command(label="复制", command=self.copy_text)
            context_menu.add_command(label="全选", command=self.select_all_text)
            context_menu.tk_popup(event.x_root, event.y_root)
        except Exception:
            pass
    
    def copy_text(self):
        """复制选中的文本"""
        try:
            self.root.clipboard_clear()
            selected_text = self.result_text.selection_get()
            self.root.clipboard_append(selected_text)
        except tk.TclError:
            # 如果没有选中文本，复制全部内容
            self.root.clipboard_clear()
            all_text = self.result_text.get(1.0, tk.END)
            self.root.clipboard_append(all_text)
    
    def select_all_text(self):
        """全选文本"""
        self.result_text.tag_add(tk.SEL, "1.0", tk.END)
        self.result_text.mark_set(tk.INSERT, "1.0")
        self.result_text.see(tk.INSERT)
    
    def format_element_with_phone(self, element, df):
        """格式化元素显示，根据设置添加联系电话"""
        if not self.show_phone.get():
            return str(element)
        
        # 查找该元素对应的联系电话
        try:
            # 根据当前对比列查找对应行
            column = self.column_var.get()
            if column and column in df.columns:
                matching_rows = df[df[column] == element]
                if not matching_rows.empty and '联系电话' in df.columns:
                    phone = matching_rows['联系电话'].iloc[0]
                    if pd.notna(phone) and str(phone).strip():
                        return f"{element} (联系电话: {phone})"
        except Exception:
            pass
        
        return str(element)
    
    def format_items_two_columns(self, items, prefix="•"):
        """将项目列表格式化为双列显示"""
        if not items:
            return ""
        
        result = ""
        items_list = list(items)
        
        # 计算每个项目的最大长度，用于对齐
        max_len = max(len(str(item)) for item in items_list) if items_list else 0
        max_len = min(max_len, 40)  # 限制最大长度，避免过长
        
        # 双列显示
        for i in range(0, len(items_list), 2):
            left_item = str(items_list[i])
            if i + 1 < len(items_list):
                right_item = str(items_list[i + 1])
                result += f"  {prefix} {left_item:<{max_len}}    {prefix} {right_item}\n"
            else:
                result += f"  {prefix} {left_item}\n"
        
        return result
        
    def select_file(self, file_num):
        """选择Excel文件"""
        file_path = filedialog.askopenfilename(
            title=f"选择文件{file_num}",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if file_path:
            if file_num == 1:
                self.file1_var.set(file_path)
                self.file1_path = file_path
            else:
                self.file2_var.set(file_path)
                self.file2_path = file_path
    
    def load_files(self):
        """加载Excel文件"""
        if not self.file1_path or not self.file2_path:
            messagebox.showerror("错误", "请先选择两个文件")
            return
        
        try:
            # 读取文件
            if self.file1_path.endswith('.csv'):
                self.df1 = pd.read_csv(self.file1_path, encoding='utf-8')
            else:
                self.df1 = pd.read_excel(self.file1_path)
            
            if self.file2_path.endswith('.csv'):
                self.df2 = pd.read_csv(self.file2_path, encoding='utf-8')
            else:
                self.df2 = pd.read_excel(self.file2_path)
            
            # 更新列选择下拉框
            common_columns = list(set(self.df1.columns) & set(self.df2.columns))
            self.column_combo['values'] = common_columns
            
            if common_columns:
                self.column_combo.set(common_columns[0])
            
            # 显示文件信息
            info = f"文件加载成功！\n"
            info += f"文件1: {os.path.basename(self.file1_path)} ({len(self.df1)}行, {len(self.df1.columns)}列)\n"
            info += f"文件2: {os.path.basename(self.file2_path)} ({len(self.df2)}行, {len(self.df2.columns)}列)\n"
            info += f"共同列: {', '.join(common_columns)}\n\n"
            
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, info)
            
            messagebox.showinfo("成功", "文件加载成功！")
            
        except Exception as e:
            messagebox.showerror("错误", f"文件加载失败: {str(e)}")
    
    def compare_data(self):
        """对比数据"""
        if self.df1 is None or self.df2 is None:
            messagebox.showerror("错误", "请先加载文件")
            return
        
        if not self.column_var.get():
            messagebox.showerror("错误", "请选择对比列")
            return
        
        # 清空之前的结果
        self.result_text.delete(1.0, tk.END)
        
        try:
            column = self.column_var.get()
            compare_type = self.compare_type.get()
            validate_data = self.validate_data.get()
            
            if validate_data:
                # 启用数据验证模式：基于全行数据对比
                # 创建以对比列为键的字典，值为完整行数据
                dict1 = {}
                dict2 = {}
                
                for idx, row in self.df1.iterrows():
                    key = str(row[column]) if pd.notna(row[column]) else None
                    if key:
                        dict1[key] = row.to_dict()
                
                for idx, row in self.df2.iterrows():
                    key = str(row[column]) if pd.notna(row[column]) else None
                    if key:
                        dict2[key] = row.to_dict()
                
                set1 = set(dict1.keys())
                set2 = set(dict2.keys())
                
                # 检查数据一致性 - 只验证关键字段：身份证号和名称
                inconsistent_data = []
                key_fields = ['身份证号', '名称']  # 定义关键验证字段
                
                for key in set1 & set2:
                    # 只检查关键字段是否一致
                    is_inconsistent = False
                    for field in key_fields:
                        if field in dict1[key] and field in dict2[key]:
                            val1 = str(dict1[key][field]) if pd.notna(dict1[key][field]) else ''
                            val2 = str(dict2[key][field]) if pd.notna(dict2[key][field]) else ''
                            if val1 != val2:
                                is_inconsistent = True
                                break
                    
                    if is_inconsistent:
                        inconsistent_data.append(key)
                
                # 保存完整数据用于导出
                self.full_data1 = dict1
                self.full_data2 = dict2
                self.inconsistent_keys = inconsistent_data
            else:
                # 传统模式：仅对比指定列
                set1 = set(self.df1[column].dropna().astype(str))
                set2 = set(self.df2[column].dropna().astype(str))
                self.full_data1 = None
                self.full_data2 = None
                self.inconsistent_keys = []
            
            result = ""
            
            # 如果启用数据验证且发现不一致数据，在最上方显示详细警告
            if validate_data and self.inconsistent_keys:
                result += "\n" + "="*80 + "\n"
                result += f"⚠️ 数据一致性警告: 发现 {len(self.inconsistent_keys)} 个重复元素但关键信息不同\n"
                result += "="*80 + "\n\n"
                
                for key in sorted(self.inconsistent_keys):
                    result += f"🔍 重复元素: {key}\n"
                    result += "-" * 60 + "\n"
                    
                    # 显示文件1的信息
                    if key in dict1:
                        data1 = dict1[key]
                        id_card1 = data1.get('身份证号', '未知')
                        name1 = data1.get('名称', '未知')
                        result += f"{name1} 身份证号：{id_card1} 位于文件1\n"
                    
                    # 显示文件2的信息
                    if key in dict2:
                        data2 = dict2[key]
                        id_card2 = data2.get('身份证号', '未知')
                        name2 = data2.get('名称', '未知')
                        result += f"{name2} 身份证号：{id_card2} 位于文件2\n"
                    
                    result += "\n"
                
                result += "="*80 + "\n\n"
            
            result += f"\n=== 对比结果 ===\n"
            result += f"对比列: {column}\n"
            result += f"对比类型: {compare_type}\n"
            result += f"数据验证: {'启用' if validate_data else '关闭'}\n\n"
            
            if compare_type == "相同元素":
                common_elements = set1 & set2
                result += f"相同元素数量: {len(common_elements)}\n"
                result += "相同元素列表:\n"
                # 格式化显示，支持联系电话
                items = sorted(common_elements)
                if items:
                    # 处理联系电话格式化
                    formatted_items = []
                    for item in items:
                        # 优先从文件1获取联系电话，如果没有则从文件2获取
                        formatted_item = self.format_element_with_phone(item, self.df1)
                        if formatted_item == str(item) and self.df2 is not None:
                            formatted_item = self.format_element_with_phone(item, self.df2)
                        formatted_items.append(formatted_item)
                    
                    # 双列显示
                    result += self.format_items_two_columns(formatted_items, "•")
                    
            elif compare_type == "文件1独有":
                unique_in_file1 = set1 - set2
                result += f"文件1独有元素数量: {len(unique_in_file1)}\n"
                result += "文件1独有元素列表:\n"
                # 格式化显示，支持联系电话
                items = sorted(unique_in_file1)
                if items:
                    # 处理联系电话格式化
                    formatted_items = []
                    for item in items:
                        formatted_item = self.format_element_with_phone(item, self.df1)
                        formatted_items.append(formatted_item)
                    
                    # 双列显示
                    result += self.format_items_two_columns(formatted_items, "•")
                    
            elif compare_type == "文件2独有":
                unique_in_file2 = set2 - set1
                result += f"文件2独有元素数量: {len(unique_in_file2)}\n"
                result += "文件2独有元素列表:\n"
                # 格式化显示，支持联系电话
                items = sorted(unique_in_file2)
                if items:
                    # 处理联系电话格式化
                    formatted_items = []
                    for item in items:
                        formatted_item = self.format_element_with_phone(item, self.df2)
                        formatted_items.append(formatted_item)
                    
                    # 双列显示
                    result += self.format_items_two_columns(formatted_items, "•")
                    
            elif compare_type == "1与2独有":
                # 文件1独有和文件2独有的元素
                unique_in_file1 = set1 - set2
                unique_in_file2 = set2 - set1
                
                result += f"文件1独有元素数量: {len(unique_in_file1)}\n"
                result += f"文件2独有元素数量: {len(unique_in_file2)}\n\n"
                
                if unique_in_file1:
                    result += "=== 文件1独有元素 ===\n"
                    # 格式化显示，支持联系电话
                    items = sorted(unique_in_file1)
                    # 处理联系电话格式化
                    formatted_items = []
                    for item in items:
                        formatted_item = self.format_element_with_phone(item, self.df1)
                        formatted_items.append(formatted_item)
                    
                    # 双列显示
                    result += self.format_items_two_columns(formatted_items, "•")
                    result += "\n"
                
                if unique_in_file2:
                    result += "=== 文件2独有元素 ===\n"
                    # 格式化显示，支持联系电话
                    items = sorted(unique_in_file2)
                    # 处理联系电话格式化
                    formatted_items = []
                    for item in items:
                        formatted_item = self.format_element_with_phone(item, self.df2)
                        formatted_items.append(formatted_item)
                    
                    # 双列显示
                    result += self.format_items_two_columns(formatted_items, "•")
                    
            elif compare_type == "所有差异":
                common_elements = set1 & set2
                unique_in_file1 = set1 - set2
                unique_in_file2 = set2 - set1
                
                result += f"相同元素数量: {len(common_elements)}\n"
                result += f"文件1独有数量: {len(unique_in_file1)}\n"
                result += f"文件2独有数量: {len(unique_in_file2)}\n\n"
                
                if common_elements:
                    result += "相同元素:\n"
                    # 格式化显示，支持联系电话
                    items = sorted(common_elements)
                    # 处理联系电话格式化
                    formatted_items = []
                    for item in items:
                        # 优先从文件1获取联系电话，如果没有则从文件2获取
                        formatted_item = self.format_element_with_phone(item, self.df1)
                        if formatted_item == str(item) and self.df2 is not None:
                            formatted_item = self.format_element_with_phone(item, self.df2)
                        formatted_items.append(formatted_item)
                    
                    # 双列显示
                    result += self.format_items_two_columns(formatted_items, "=")
                    result += "\n"
                
                if unique_in_file1:
                    result += "文件1独有元素:\n"
                    # 格式化显示，支持联系电话
                    items = sorted(unique_in_file1)
                    # 处理联系电话格式化
                    formatted_items = []
                    for item in items:
                        formatted_item = self.format_element_with_phone(item, self.df1)
                        formatted_items.append(formatted_item)
                    
                    # 双列显示
                    result += self.format_items_two_columns(formatted_items, "+")
                    result += "\n"
                
                if unique_in_file2:
                    result += "文件2独有元素:\n"
                    # 格式化显示，支持联系电话
                    items = sorted(unique_in_file2)
                    # 处理联系电话格式化
                    formatted_items = []
                    for item in items:
                        formatted_item = self.format_element_with_phone(item, self.df2)
                        formatted_items.append(formatted_item)
                    
                    # 双列显示
                    result += self.format_items_two_columns(formatted_items, "-")
            
            # 保存结果用于导出
            self.last_result = {
                'column': column,
                'type': compare_type,
                'validate_data': validate_data,
                'common': set1 & set2 if compare_type in ["相同元素", "所有差异"] else set(),
                'file1_unique': set1 - set2 if compare_type in ["文件1独有", "所有差异"] else set(),
                'file2_unique': set2 - set1 if compare_type in ["文件2独有", "所有差异"] else set(),
                'complement': (set1 | set2) - (set1 & set2) if compare_type == "1与2独有" else set(),
                'inconsistent': set(self.inconsistent_keys) if validate_data else set()
            }
            
            self.result_text.insert(tk.END, result)
            
        except Exception as e:
            messagebox.showerror("错误", f"对比失败: {str(e)}")
    

    def clear_results(self):
        """清空结果"""
        self.result_text.delete(1.0, tk.END)
        if hasattr(self, 'last_result'):
            delattr(self, 'last_result')

def main():
    root = tk.Tk()
    app = ExcelCompareTool(root)
    root.mainloop()

if __name__ == "__main__":
    main()