from typing import List, Any
import xlwt
from xmindparser import xmind_to_dict
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import win32com.client
import pythoncom
import time

# 增量索引,表的行数
index = 2
big_title=""
def is_file_locked(file_path):
    """
    检查Excel文件是否被占用
    """
    if not os.path.exists(file_path):
        return False
        
    try:
        # 尝试打开文件
        with open(file_path, 'a') as _:
            pass
        return False
    except IOError:
        return True

def resolve_path(dict_,f,file_path,sheet,style,level=0):
    """
    遍历字典结构，当某个层级的所有子节点都不含有 topics 时，打印这些子节点的 title；
    否则继续递归到下一个子节点。
    :param dict_: 当前分支的字典
    :param lists: 存储拼接标题的列表
    :param title: 上一级的标题
    :param level: 当前处理的层级
    :return:
    """
    global index

    try:
        if "topics" not in dict_:
            # 每次写入前检查文件是否被占用
            if is_file_locked(file_path):
                raise IOError("Excel文件正在被其他程序使用，请关闭后重试")
                
            sheet.write(index, 0, level, style)
            sheet.write(index, 1, dict_['title'], style)
            sheet.write(index, 2, dict_['title'], style)
            index += 1
            
            try:
                f.save(file_path)
            except Exception:
                raise IOError("无法保存Excel文件，请确保文件未被打开")
            return
            
        else:
            # 检查子节点是否都不含有 topics
            all_leaf_nodes = all("topics" not in sub_dict for sub_dict in dict_["topics"])

            if all_leaf_nodes:
                # 打印当前层级下的所有子节点的标题
                for  topic in dict_["topics"]:
                    numbered_topics=topic["title"]
                    # numbered_topics=[f"{i + 1}. {sub_dict['title']}" for i, sub_dict in enumerate(dict_['topics'])]
                    # topics_str='\n'.join(numbered_topics)
                    sheet.write(index, 0, level,style)
                    sheet.write(index, 1, dict_['title'],style)
                    sheet.write(index, 2, numbered_topics,style)
                    index += 1
                    f.save(file_path)
                    print(f"Level {level}: {dict_['title']} -> 子节点: {numbered_topics}")
                return
            else:
                # 打印当前层级并继续递归处理
                print(f"Level {level}: {dict_['title']}(继续递归下一级)")
                sheet.write(index, 0, level,style)
                sheet.write(index, 1, dict_['title'],style)
                index+=1
                # print(index)
                f.save(file_path)
                for sub_dict in dict_["topics"]:
                    resolve_path(sub_dict,f,file_path,sheet,style,level + 1)
    except Exception as e:
        print(f"处理节点 {dict_['title']} 时发生错误: {str(e)}")




def xmind_to_excel(list_, excel_path):
    f = xlwt.Workbook()
    # 生成单sheet的Excel文件，sheet名自取
    sheet = f.add_sheet("模块", cell_overwrite_ok=True)
    style = xlwt.XFStyle()
    # 设置单元格的对齐方式
    alignment = xlwt.Alignment()
    alignment.wrap = 1  # 开启换行
    # alignment.vert = xlwt.Alignment.VERT_TOP  # 垂直顶部对齐
    alignment.horz=xlwt.Alignment.HORZ_LEFT
    # alignment.horz = xlwt.Alignment.HORZ_CENTER  # 水平居中
    alignment.vert = xlwt.Alignment.VERT_CENTER  # 垂直居中
    style.alignment = alignment

    # 第一行固定的表头标题
    # 创建加粗样式
    bold_style = xlwt.XFStyle()
    font = xlwt.Font()
    font.bold = True  # 设置字体为加粗
    bold_style.font = font

    row_header = ["级别", "功能点", "测试点", "操作步骤", "预期结果","测试结果","负责人","备注"]
    for i in range(0, len(row_header)):
        sheet.write(0, i, row_header[i],bold_style)

    # 设置列宽度（单位是1/256字符宽度）
    sheet.col(0).width = 256 * 10
    sheet.col(1).width = 256 * 50
    sheet.col(2).width = 256 * 50
    sheet.col(3).width = 256 * 60
    sheet.col(4).width = 256 * 80
    sheet.col(5).width = 256 * 15
    sheet.col(6).width = 256 * 15
    sheet.col(7).width = 256 * 15
    sheet.col(8).width = 256 * 15

    sheet.write(1, 0, 1, style)
    sheet.write(1, 1, list_['title'], style)
    f.save(excel_path)

    for h in range(0, len(list_['topics'])):
        lists: List[Any] = []
        resolve_path(list_['topics'][h],f,excel_path,sheet,style,2)
        # print(lists)

def run(xmind_path):
    try:
        # 检查文件是否存在
        if not os.path.exists(xmind_path):
            raise FileNotFoundError("找不到指定的XMind文件")
            
        # 检查文件扩展名
        if not xmind_path.lower().endswith('.xmind'):
            raise ValueError("请选择正确的XMind文件")
            
        # 将XMind转化成字典
        xmind_dict = xmind_to_dict(xmind_path)
        excel_name = os.path.splitext(os.path.basename(xmind_path))[0] + '.xls'
        excel_path = os.path.join(os.path.dirname(xmind_path), excel_name)
        
        xmind_to_excel(xmind_dict[0]['topic'], excel_path)
        return True, f"转换完成！文件保存路径: {excel_path}"
    except Exception as e:
        return False, f"转换失败：{str(e)}"

class XMindConverterUI:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("XMind转Excel工具")
        self.window.geometry("600x400")  # 增加窗口高度以容纳提示信息
        
        # 创建主框架
        self.main_frame = tk.Frame(self.window, padx=20, pady=20)
        self.main_frame.pack(expand=True)
        
        # 添加标题标签
        self.title_label = tk.Label(
            self.main_frame, 
            text="XMind转Excel工具",
            font=("Arial", 16, "bold")
        )
        self.title_label.pack(pady=(0, 20))
        
        # 添加说明文本
        self.instruction_text = tk.Label(
            self.main_frame,
            text="使用说明：\n" 
                 "1. 点击\"选择XMind文件\"按钮选择要转换的文件\n"
                 "2. 确保目标Excel文件未被打开\n"
                 "3. 点击\"开始转换\"按钮进行转换\n"
                 "4. 转换完成后会在同目录生成Excel文件",
            justify=tk.LEFT,
            font=("Arial", 10)
        )
        self.instruction_text.pack(pady=(0, 20))
        
        # 文件路径框架
        self.path_frame = tk.Frame(self.main_frame)
        self.path_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 添加路径标签
        self.path_label = tk.Label(
            self.path_frame, 
            text="文件路径：",
            font=("Arial", 10)
        )
        self.path_label.pack(side=tk.LEFT)
        
        # 文件路径显示
        self.path_var = tk.StringVar()
        self.path_entry = tk.Entry(
            self.path_frame, 
            textvariable=self.path_var, 
            width=50,
            state='readonly'  # 设置为只读
        )
        self.path_entry.pack(side=tk.LEFT, padx=5)
        
        # 按钮框架
        self.button_frame = tk.Frame(self.main_frame)
        self.button_frame.pack(pady=10)
        
        # 选择文件按钮
        self.select_button = tk.Button(
            self.button_frame, 
            text="选择XMind文件", 
            command=self.select_file,
            width=15,
            height=2
        )
        self.select_button.pack(side=tk.LEFT, padx=5)
        
        # 转换按钮
        self.convert_button = tk.Button(
            self.button_frame, 
            text="开始转换", 
            command=self.convert,
            width=15,
            height=2
        )
        self.convert_button.pack(side=tk.LEFT, padx=5)
        
        # 添加状态标签
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        self.status_label = tk.Label(
            self.main_frame,
            textvariable=self.status_var,
            font=("Arial", 10),
            fg="gray"
        )
        self.status_label.pack(pady=(10, 0))
        
    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="选择XMind文件",
            filetypes=[("XMind文件", "*.xmind"), ("所有文件", "*.*")]
        )
        if file_path:
            self.path_var.set(file_path)
            self.status_var.set("已选择文件，请点击\"开始转换\"")
            
    def convert(self):
        xmind_path = self.path_var.get().strip()
        if not xmind_path:
            messagebox.showerror("错误", "请先选择XMind文件！")
            return
            
        # 获取目标Excel文件路径
        excel_name = os.path.splitext(os.path.basename(xmind_path))[0] + '.xls'
        excel_path = os.path.join(os.path.dirname(xmind_path), excel_name)
        
        # 检查Excel文件是否被占用
        if os.path.exists(excel_path) and is_file_locked(excel_path):
            messagebox.showerror(
                "错误", 
                f"Excel文件 '{excel_name}' 正在被其他程序使用！\n"
                "请关闭该文件后重试。"
            )
            return
            
        self.status_var.set("正在转换...")
        self.window.update()  # 更新界面显示
        
        success, message = run(xmind_path)
        if success:
            self.status_var.set("转换完成")
            messagebox.showinfo("成功", message)
        else:
            self.status_var.set("转换失败")
            messagebox.showerror("错误", message)
            
    def run(self):
        # 设置窗口图标（如果有的话）
        try:
            self.window.iconbitmap("icon.ico")  # 如果有图标文件的话
        except:
            pass
            
        self.window.mainloop()

if __name__ == '__main__':
    # app = XMindConverterUI()
    # app.run()
    test_file = r'C:\Users\v_ahongchen\Desktop\炽墨幻境—蹴鞠王.xmind'  # 替换为实际的文件路径
    print(f"开始处理文件: {test_file}")
    
    success, message = run(test_file)
    print(f"处理结果: success={success}, message={message}")