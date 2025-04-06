import tkinter as tk
from tkinter import messagebox

# 假设这四个文件和当前脚本在同一目录下
# 导入四个模块
import dataPreprocessing as test1
import statisticDay as test2
import wagesCalculation as test3
import Statistics as test4

def run_test1():
    try:
        test1.main()  # 假设test1.py中的主函数名为main，根据实际情况修改
        messagebox.showinfo("提示", "test1 执行完成")
    except Exception as e:
        messagebox.showerror("错误", f"test1 执行失败: {str(e)}")


def run_test2():
    try:
        test2.main()  # 假设test2.py中的主函数名为main，根据实际情况修改
        messagebox.showinfo("提示", "test2 执行完成")
    except Exception as e:
        messagebox.showerror("错误", f"test2 执行失败: {str(e)}")


def run_test3():
    try:
        test3.main()  # 假设test3.py中的主函数名为main，根据实际情况修改
        messagebox.showinfo("提示", "test3 执行完成")
    except Exception as e:
        messagebox.showerror("错误", f"test3 执行失败: {str(e)}")


def run_test4():
    try:
        test4.main()  # 假设test4.py中的主函数名为main，根据实际情况修改
        messagebox.showinfo("提示", "test4 执行完成")
    except Exception as e:
        messagebox.showerror("错误", f"test4 执行失败: {str(e)}")


root = tk.Tk()
root.title("依次执行测试脚本")

# 添加提示标签
prompt_label = tk.Label(root, text="请按照 test1、test2、test3、test4 的顺序依次执行。", justify=tk.CENTER)
prompt_label.pack(pady=10)

button1 = tk.Button(root, text="运行 test1", command=run_test1)
button1.pack(pady=10)

button2 = tk.Button(root, text="运行 test2", command=run_test2)
button2.pack(pady=10)

button3 = tk.Button(root, text="运行 test3", command=run_test3)
button3.pack(pady=10)

button4 = tk.Button(root, text="运行 test4", command=run_test4)
button4.pack(pady=10)

root.mainloop()