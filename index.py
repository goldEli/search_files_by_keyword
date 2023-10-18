import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import os


def search_files(folder_path):
    file_extensions = ['.doc', '.docx', '.pdf', '.ppt', '.xls', '.xlsx']
    results = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            filename, extension = os.path.splitext(file)
            if extension.lower() in file_extensions:
                file_path = os.path.join(root, file)
                results.append(file_path)
    return results


def on_select_folder():
    folder_path = filedialog.askdirectory()
    folder_label.config(text=folder_path)


def on_submit():
    keyword = keyword_entry.get()
    folder = folder_label.cget("text")
    # result_text.config(text=f"关键字：{keyword}\n选择的文件夹：{folder}")
    result_text.config(text=f"搜索中。。。")

    # 调用函数搜索文件并打印路径
    files = search_files(folder)
    print(files)

    output = ""
    for file in files:
        output += file + "\n"

    result_text.config(text=output)


# 创建主窗口
window = tk.Tk()
window.title("应用")

# 创建样式
style = ttk.Style()
style.configure("TLabel", background="white")
style.configure("TButton", padding=6, relief="flat")

# 创建关键字标签和输入框
keyword_label = ttk.Label(window, text="关键字：")
keyword_label.grid(row=0, column=0, padx=10, pady=10, sticky="e")
keyword_entry = ttk.Entry(window, width=30)
keyword_entry.grid(row=0, column=1, padx=10, pady=10, sticky="w")

# 创建选择文件夹按钮和标签
folder_button = ttk.Button(window, text="选择文件夹", command=on_select_folder)
folder_button.grid(row=1, column=0, padx=10, pady=10, sticky="e")
folder_label = ttk.Label(window, text="未选择文件夹")
folder_label.grid(row=1, column=1, padx=10, pady=10, sticky="w")

# 创建确定按钮
submit_button = ttk.Button(window, text="搜索", command=on_submit)
submit_button.grid(row=2, column=0, columnspan=2, padx=10, pady=10)

# 创建结果展示标签
result_text = ttk.Label(window, text="")
result_text.grid(row=3, column=0, columnspan=2, padx=10, pady=10)
# result_text = tk.Text(window, width=40, height=10)
# result_text.grid(row=3, column=0, columnspan=2, padx=10, pady=10)
# result_text.configure(state="disabled") # 禁止编辑

# 运行主循环
window.mainloop()
