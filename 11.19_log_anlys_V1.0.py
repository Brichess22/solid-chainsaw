import re
import os
import pandas as pd
import tkinter as tk
from tkinter import simpledialog,messagebox


# 合并所有 txt 文件到一个文件中
def merge_txt_files(folder_path, output_file, input_file_encoding):
    try:
        # 获取指定文件夹中的所有文件
        files = os.listdir(folder_path)
        
        # 过滤出所有的 log 文件
        txt_files = [f for f in files if f.endswith('.log')]

        with open(output_file, 'w', encoding='utf-8') as outfile:
            for txt_file in txt_files:
                file_path = os.path.join(folder_path, txt_file)

                # 写入当前处理的文件名
                outfile.write(f"filename:{txt_file}\n")

                # 逐个打开 log 文件并写入到输出文件
                with open(file_path, 'r', encoding=input_file_encoding) as infile:
                    content = infile.read()
                    outfile.write(content)
                    outfile.write('\n')  # 添加换行符以区分文件内容

        print(f"合并完成，输出文件为: {output_file}")
    except Exception as e:
        print(f"合并文件时出错: {e}")


# 从文件中提取 logon 的日期和 IP 信息
def extract_logon_details(file_path, user_pattern):
    results = []
    try:
        # 修改正则表达式，添加容错机制
        pattern = re.compile(user_pattern)
        
        with open(file_path, 'r', encoding='utf-8') as file:
            for line in file:
                match = pattern.search(line)
                if match:
                    if match.group(1) and match.group(2):
                        date_time = match.group(1)
                        ip_port = match.group(2)
                        results.append({"日期时间": date_time, "IP:端口": ip_port})
                    elif match.group(3):
                        # 匹配到文件名
                        file_name = match.group(3)
                        results.append({"文件名": file_name})
    except Exception as e:
        print(f"提取 logon 信息时出错: {e}")

    return results


# 自定义输入对话框
def custom_input_dialog(title, prompt, default_value, width=50):
    # 创建一个新的窗口
    top = tk.Toplevel()
    top.title(title)
    
    # 设置窗口大小
    top.geometry("500x100")  # 自定义窗口的大小
    
    # 添加提示标签
    label = tk.Label(top, text=prompt)
    label.pack(pady=10)
    
    # 添加文本框
    input_text = tk.Entry(top, width=width)
    input_text.insert(0, default_value)  # 默认值
    input_text.pack(pady=10)
    
    # 添加按钮来关闭窗口并返回输入的值
    result = {'value': default_value}
    def on_submit():
        result['value'] = input_text.get()  # 获取输入的值
        top.destroy()  # 关闭窗口
    
    submit_button = tk.Button(top, text="提交", command=on_submit)
    submit_button.pack(pady=10)
    
    # 启动窗口事件循环
    top.wait_window()  # 阻止后续代码执行，直到关闭此窗口
    return result['value']


# 主逻辑
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    
    folder_path = r"gold"  # 存放 log 文件的文件夹路径
    output_file = r"gold/merged.txt"  # 合并后的文件路径
    excel_file = r"gold/logon_details.xlsx"  # 输出的 Excel 文件路径
    # 确认是否要覆盖输出文件
    if os.path.exists(output_file):
        if not messagebox.askyesno("提示", f"输出文件 '{output_file}' 已存在，是否覆盖？"):
            print("操作已取消")
            exit()
    # 确认是否要覆盖输出 Excel 文件
    if os.path.exists(excel_file):
        if not messagebox.askyesno("提示", f"输出 Excel 文件 '{excel_file}' 已存在，是否覆盖？"):
            print("操作已取消")
            exit()
  
    user_pattern = custom_input_dialog(
        title="输入正则表达式", 
        prompt="请输入提取信息的正则表达式：", 
        default_value=r'(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}).*?(\d{1,3}(?:\.\d{1,3}){3}:\d{1,5}).*?logon|filename:(.*?\.log)',
        width=190  # 你可以调整输入框的宽度
    )

    # 输入文件编码
    input_file_encoding = custom_input_dialog(
        title="文件编码",
        prompt="请输入输入文件的编码（如 'GBK', 'utf-8' 等）：",
        default_value="GBK",
        width=20  # 可以调整输入框的宽度
    )

    # 合并文件
    merge_txt_files(folder_path, output_file, input_file_encoding)

    # 提取 logon 信息
    logon_details = extract_logon_details(output_file, user_pattern)
    for detail in logon_details:
        print(detail)

    # 使用 pandas 将结果输出到 Excel 文件
    try:
        df = pd.DataFrame(logon_details)
        df.to_excel(excel_file, index=False)
        print(f"Logon 详情已导出到: {excel_file}")
    except Exception as e:
        print(f"导出 Excel 文件时出错: {e}")
