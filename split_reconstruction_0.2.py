import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from datetime import datetime

# 拆分功能：按列拆分
def split_data_by_column(input_file, output_folder):
    try:
        # 读取 Excel 文件
        df = pd.read_excel(input_file)
        
        # 检查是否选择了输出文件夹
        if not output_folder:
            messagebox.showerror("错误", "未选择输出文件夹！")
            return

        # 遍历每一列拆分
        for col in df.columns:
            # 选择每一列并保存为单独的 Excel 文件
            split_df = df[[col]]
            output_file = os.path.join(output_folder, f'split_{col}.xlsx')  # 输出文件的路径
            
            # 检查文件是否已存在，若存在则重命名文件
            if os.path.exists(output_file):
                # 给文件名添加时间戳来避免覆盖
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_file = os.path.join(output_folder, f'split_{col}_{timestamp}.xlsx')
            
            split_df.to_excel(output_file, index=False)  # 将拆分的列保存为新的 Excel 文件
        
        messagebox.showinfo("完成", f"数据按列拆分完成！文件已保存至：{output_folder}")
    except Exception as e:
        # 如果发生错误，显示错误信息
        messagebox.showerror("错误", f"拆分失败: {str(e)}")

# 重构功能：按列重构
def reconstruct_data(input_files, output_folder):
    try:
        df_list = []
        # 读取每个拆分后的文件并合并
        for file in input_files:
            df = pd.read_excel(file)
            df_list.append(df)
        
        # 将拆分文件按列合并
        merged_df = pd.concat(df_list, axis=1)
        
        # 确保选择了输出文件夹
        if not output_folder:
            messagebox.showerror("错误", "未选择输出文件夹！")
            return

        # 设置输出文件路径
        output_file = os.path.join(output_folder, 'reconstructed_file.xlsx')  # 合并后的文件保存路径
        
        # 检查文件是否已存在，若存在则重命名文件
        if os.path.exists(output_file):
            # 给文件名添加时间戳来避免覆盖
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(output_folder, f'reconstructed_file_{timestamp}.xlsx')

        # 保存重构后的文件
        merged_df.to_excel(output_file, index=False)  # 将合并后的数据保存为新的 Excel 文件
        messagebox.showinfo("完成", f"数据重构完成！文件已保存至：{output_file}")
    except Exception as e:
        # 如果发生错误，显示错误信息
        messagebox.showerror("错误", f"重构失败: {str(e)}")

# 选择拆分文件
def choose_split_file():
    # 弹出文件选择对话框，选择要拆分的 Excel 文件
    input_file = filedialog.askopenfilename(title="选择要拆分的 Excel 文件", filetypes=[("Excel Files", "*.xlsx")])
    if input_file:
        # 弹出文件夹选择对话框，选择保存拆分文件的文件夹
        output_folder = filedialog.askdirectory(title="选择保存拆分文件的文件夹")
        if output_folder:
            # 调用拆分函数进行拆分
            split_data_by_column(input_file, output_folder)

# 选择重构文件
def choose_reconstruct_files():
    # 弹出文件选择对话框，选择要重构的多个 Excel 文件
    input_files = filedialog.askopenfilenames(title="选择要重构的 Excel 文件", filetypes=[("Excel Files", "*.xlsx")])
    if input_files:
        # 将 input_files 从元组转换为列表，确保后续操作可修改列表
        input_files = list(input_files)
        # 创建一个新的窗口来显示选择的文件
        show_selected_files(input_files)

# 显示已选文件的窗口
def show_selected_files(input_files):
    # 创建新的窗口来展示选择的文件
    window = tk.Toplevel()
    window.title("已选择的文件")

    # 设置窗口大小
    window.geometry("400x400")

    # 显示说明文本
    label = tk.Label(window, text="以下是你选择的文件，确认后点击'确认'按钮进行重构：")
    label.pack(pady=10)

    # 在窗口中显示已选择的文件列表
    file_listbox = tk.Listbox(window, selectmode=tk.MULTIPLE, height=10, width=50)
    for file in input_files:
        file_listbox.insert(tk.END, file)  # 将选择的文件逐个插入到列表框中
    file_listbox.pack(pady=10)

    # 添加文件按钮，允许用户继续选择文件
    def add_files():
        new_files = filedialog.askopenfilenames(title="选择要添加的 Excel 文件", filetypes=[("Excel Files", "*.xlsx")])
        if new_files:
            new_files = list(new_files)  # 确保新选择的文件是列表类型
            for file in new_files:
                if file not in input_files:  # 确保没有重复添加文件
                    input_files.append(file)  # 添加新文件到列表
                    file_listbox.insert(tk.END, file)  # 将新文件添加到列表框中

    # 删除文件按钮，允许用户从列表中删除文件
    def remove_file():
        selected_files = list(file_listbox.curselection())  # 获取用户选择的文件
        if selected_files:
            for file_index in selected_files[::-1]:
                file_listbox.delete(file_index)  # 从列表框中删除选中的文件
                input_files.pop(file_index)  # 从文件列表中删除文件
        else:
            messagebox.showwarning("警告", "请先选择要删除的文件")  # 如果没有选择文件，提示用户

    # 创建按钮来添加文件
    add_button = tk.Button(window, text="添加文件", command=add_files)
    add_button.pack(pady=5)

    # 创建按钮来删除文件
    remove_button = tk.Button(window, text="删除已选文件", command=remove_file)
    remove_button.pack(pady=5)

    # 创建确认按钮，点击后执行重构操作
    def confirm_reconstruct():
        selected_files = input_files  # 使用最终的文件列表
        if not selected_files:
            messagebox.showerror("错误", "请至少选择一个文件进行重构")  # 如果没有选择文件，提示用户
            return
        output_folder = filedialog.askdirectory(title="选择保存重构文件的文件夹")  # 选择保存重构文件的文件夹
        if output_folder:
            # 调用重构函数执行文件合并操作
            reconstruct_data(selected_files, output_folder)
            window.destroy()  # 关闭文件选择窗口

    # 确认按钮，执行重构操作
    confirm_button = tk.Button(window, text="确认", command=confirm_reconstruct)
    confirm_button.pack(pady=20)

    # 运行窗口
    window.mainloop()

# 创建 GUI
root = tk.Tk()
root.title("Excel 数据拆分与重构")

# 设置窗口大小
root.geometry("400x200")

# 创建拆分按钮
split_button = tk.Button(root, text="拆分数据按列", command=choose_split_file)
split_button.pack(pady=20)

# 创建重构按钮
reconstruct_button = tk.Button(root, text="重构数据", command=choose_reconstruct_files)
reconstruct_button.pack(pady=20)

# 运行 GUI 主循环
root.mainloop()
