import re
import os
import sys
import pandas as pd
import tkinter as tk
from tkinter import simpledialog, messagebox

# 合并所有 txt 文件到一个文件中
def merge_txt_files(folder_path, output_file, input_file_encoding):
    try:
        # 检查文件夹路径是否存在
        if not os.path.exists(folder_path):
            print(f"文件夹 '{folder_path}' 不存在")
            return
        # 获取指定文件夹中的所有文件
        files = os.listdir(folder_path)

        if not files:
            print(f"文件夹 '{folder_path}' 中没有文件")
            return

        # 过滤出所有的 log 文件
        txt_files = [f for f in files if f.endswith('.log')]

        if not txt_files:
            print(f"文件夹 '{folder_path}' 中没有 log 文件")
            return
        
        with open(output_file, 'w', encoding='utf-8') as outfile:
            for txt_file in txt_files:
                file_path = os.path.join(folder_path, txt_file)

                # 写入当前处理的文件名
                outfile.write(f"filename:{txt_file}\n")
                try:
                    # 逐个打开 log 文件并写入到输出文件
                   with open(file_path, 'r', encoding=input_file_encoding) as infile:
                        content = infile.read()
                        outfile.write(content)
                        outfile.write('\n')  # 添加换行符以区分文件内容
                except Exception as e:
                    print(f"读取文件 '{file_path}' 时出错: {e}")

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
    except re.error as e:
        print(f"正则表达式错误: {e}")
    except Exception as e:
        print(f"提取 logon 信息时出错: {e}")

    return results


# 自定义输入对话框
def custom_input_dialog(title, prompt, default_value, width=50, win_width="200", win_height="100", font_size=10):
    while True:
        # 创建一个新的窗口
        top = tk.Toplevel()
        top.title(title)

        # 获取屏幕宽度和高度
        ws = top.winfo_screenwidth()
        hs = top.winfo_screenheight()

        # 计算窗口坐标
        x = (ws / 2) - (int(win_width) / 2)
        y = (hs / 2) - (int(win_height) / 2)

        # 设置窗口大小和位置
        top.geometry(win_width + 'x' + win_height + '+' + str(int(x)) + '+' + str(int(y)))

        font = ('Arial', font_size)  # 设置字体和字号

        # 添加提示标签
        label = tk.Label(top, text=prompt, font=font)
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
        submit_button.pack(side=tk.LEFT, padx=50, pady=5, expand=True)

        # 添加按钮关闭窗口
        def on_cancel():
            result['value'] = None  # 取消操作
            top.destroy()  # 关闭窗口
            return result  # 直接返回空值

        cancel_button = tk.Button(top, text="取消", command=on_cancel)
        cancel_button.pack(side=tk.RIGHT, padx=50, pady=5, expand=True)

        # 当用户点击右上角关闭按钮时，关闭窗口
        top.protocol("WM_DELETE_WINDOW", on_cancel)

        # 启动窗口事件循环
        top.wait_window()  # 阻止后续代码执行，直到关闭此窗口

        # 如果用户点击了取消按钮或关闭窗口
        if result['value'] is None:
            return None  # 返回None表示取消操作

        if result['value']:
            return result['value']
        else:
            messagebox.showwarning("输入无效", "请输入有效的内容。")


# 检查路径是否有效
def is_valid_path(path):
    return os.path.exists(path)


# 检查正则表达式是否有效
def is_valid_regex(pattern):
    try:
        re.compile(pattern)
        return True
    except re.error:
        return False


# 检查编码格式是否有效
def is_valid_encoding(encoding):
    try:
        ''.encode(encoding)
        return True
    except (UnicodeEncodeError, TypeError):
        return False


# 主逻辑
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    
    # 获取用户输入的路径，直到有效
    folder_path = custom_input_dialog(
        title="输入文件夹路径",
        prompt="请输入存放 log 文件的文件夹路径：",
        default_value=r"gold",
        width=50,
        win_width="350",
        win_height="150"
    )

    if folder_path is None:
        print("操作已取消")
        sys.exit()  # 退出程序

    while not is_valid_path(folder_path):
        messagebox.showwarning("路径无效", "请输入有效的文件夹路径。")
        folder_path = custom_input_dialog(
            title="输入文件夹路径",
            prompt="请输入存放 log 文件的文件夹路径：",
            default_value=r"gold",
            width=50,
            win_width="350",
            win_height="150"
        )

        if folder_path is None:
            print("操作已取消")
            sys.exit()  # 退出程序

    output_file = custom_input_dialog(
        title="输入输出文件路径",
        prompt="请输入合并后的文件路径：",
        default_value=r"gold/merged.txt",
        width=50,
        win_width="350",
        win_height="150"
    )

    if output_file is None:
        print("操作已取消")
        sys.exit()  # 退出程序

    while not is_valid_path(os.path.dirname(output_file)):
        messagebox.showwarning("路径无效", "请输入有效的输出文件路径。")
        output_file = custom_input_dialog(
            title="输入输出文件路径",
            prompt="请输入合并后的文件路径：",
            default_value=r"gold/merged.txt",
            width=50,
            win_width="350",
            win_height="150"
        )

        if output_file is None:
            print("操作已取消")
            sys.exit()  # 退出程序

    # 确认是否要覆盖输出文件
    if os.path.exists(output_file):
        if not messagebox.askyesno("提示", f"输出文件 '{output_file}' 已存在，是否覆盖？"):
            print("操作已取消")
            exit()

    excel_file = custom_input_dialog(
        title="输入 Excel 文件路径",
        prompt="请输入输出 Excel 文件路径：",
        default_value=r"gold/logon_details.xlsx",
        width=50,
        win_width="350",
        win_height="150"
    )

    if excel_file is None:
        print("操作已取消")
        sys.exit()  # 退出程序

    while not is_valid_path(os.path.dirname(excel_file)):
        messagebox.showwarning("路径无效", "请输入有效的 Excel 文件路径。")
        excel_file = custom_input_dialog(
            title="输入 Excel 文件路径",
            prompt="请输入输出 Excel 文件路径：",
            default_value=r"gold/logon_details.xlsx",
            width=50,
            win_width="350",
            win_height="150"
        )

        if excel_file is None:
            print("操作已取消")
            sys.exit()  # 退出程序

    # 确认是否要覆盖输出 Excel 文件
    if os.path.exists(excel_file):
        if not messagebox.askyesno("提示", f"输出 Excel 文件 '{excel_file}' 已存在，是否覆盖？"):
            print("操作已取消")
            exit()

    user_pattern = custom_input_dialog(
        title="输入正则表达式",
        prompt="请输入提取信息的正则表达式：",
        default_value=r'(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}).*?(\d{1,3}(?:\.\d{1,3}){3}:\d{1,5}).*?logon|filename:(.*?\.log)',
        width=190,
        win_width="550",
        win_height="150",
        font_size=18
    )

    if user_pattern is None:
        print("操作已取消")
        sys.exit()  # 退出程序

    while not is_valid_regex(user_pattern):
        messagebox.showwarning("正则表达式无效", "请输入有效的正则表达式。")
        user_pattern = custom_input_dialog(
            title="输入正则表达式",
            prompt="请输入提取信息的正则表达式：",
            default_value=r'(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}).*?(\d{1,3}(?:\.\d{1,3}){3}:\d{1,5}).*?logon|filename:(.*?\.log)',
            width=190,
            win_width="550",
            win_height="150",
            font_size=18
        )

        if user_pattern is None:
            print("操作已取消")
            sys.exit()  # 退出程序

    # 输入文件编码
    input_file_encoding = custom_input_dialog(
        title="文件编码",
        prompt="请输入输入文件的编码（如 'GBK', 'utf-8' 等）：",
        default_value="GBK",
        width=20,
        win_width="350",
        win_height="150"
    )

    if input_file_encoding is None:
        print("操作已取消")
        sys.exit()  # 退出程序

    while not is_valid_encoding(input_file_encoding):
        messagebox.showwarning("编码无效", "请输入有效的编码格式。")
        input_file_encoding = custom_input_dialog(
            title="文件编码",
            prompt="请输入输入文件的编码（如 'GBK', 'utf-8' 等）：",
            default_value="GBK",
            width=20,
            win_width="350",
            win_height="150"
        )

        if input_file_encoding is None:
            print("操作已取消")
            sys.exit()  # 退出程序

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
