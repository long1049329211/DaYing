# -*- coding: utf-8 -*-
"""
批量打印工具 (兼容 MS Office / WPS Office)
支持 Word(.doc/.docx) / Excel(.xls/.xlsx) / PDF
Python 3.8+，选择文件夹后可视化勾选文件，首个文件弹出打印设置，其余静默打印
"""

import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import win32com.client as win32
import win32api
import pythoncom

pythoncom.CoInitialize()

# ------------------------- 办公应用获取 -------------------------
def get_word_app():
    try:
        return win32.Dispatch("Word.Application")
    except Exception:
        try:
            return win32.Dispatch("WPS.Application")
        except Exception:
            return None

def get_excel_app():
    try:
        return win32.Dispatch("Excel.Application")
    except Exception:
        try:
            return win32.Dispatch("ET.Application")
        except Exception:
            return None

def setup_printer_dialog(word_app):
    """仅弹出打印设置对话框，不实际打印纸张"""
    doc = word_app.Documents.Add()
    word_app.Dialogs(88).Show()
    doc.Close(0)

# ------------------------- 打印分组处理 -------------------------
def print_word_files(word_app, files):
    """打印 Word 文件组，files 为绝对路径列表，首个文件弹出设置"""
    word_app.Visible = True
    word_app.DisplayAlerts = False

    # 打开第一个文件
    try:
        doc = word_app.Documents.Open(files[0])
    except Exception as e:
        messagebox.showerror("错误", f"无法打开 Word 文件：\n{files[0]}\n\n错误：{e}")
        return False

    result = word_app.Dialogs(88).Show()
    if result == -1:      # 用户点击了“打印”
        doc.Close(0)
    else:
        doc.Close(0)
        ans = messagebox.askyesno("确认", "你取消了第一个文件的打印，\n是否仍使用当前打印机设置打印剩余 Word 文件？")
        if not ans:
            return False

    # 静默打印剩余 Word 文件
    if len(files) > 1:
        word_app.Visible = False
        word_app.ScreenUpdating = False
        for path in files[1:]:
            try:
                doc = word_app.Documents.Open(path)
                doc.PrintOut(Background=False)
                doc.Close(0)
            except Exception as e:
                messagebox.showwarning("跳过文件", f"无法打印 Word：\n{path}\n\n错误：{e}")
        word_app.Visible = True
    else:
        word_app.Visible = False
    return True

def print_excel_files(excel_app, files):
    """打印 Excel 文件组，files 为绝对路径列表，首个文件弹出设置"""
    excel_app.Visible = True
    excel_app.DisplayAlerts = False

    try:
        wb = excel_app.Workbooks.Open(files[0])
    except Exception as e:
        messagebox.showerror("错误", f"无法打开 Excel 文件：\n{files[0]}\n\n错误：{e}")
        return False

    result = excel_app.Dialogs(8).Show()
    if result == True:
        wb.Close(False)
    else:
        wb.Close(False)
        ans = messagebox.askyesno("确认", "你取消了第一个 Excel 文件的打印，\n是否仍使用当前打印机设置打印剩余 Excel 文件？")
        if not ans:
            return False

    if len(files) > 1:
        excel_app.Visible = False
        excel_app.ScreenUpdating = False
        for path in files[1:]:
            try:
                wb = excel_app.Workbooks.Open(path)
                wb.PrintOut()
                wb.Close(False)
            except Exception as e:
                messagebox.showwarning("跳过文件", f"无法打印 Excel：\n{path}\n\n错误：{e}")
        excel_app.Visible = True
    else:
        excel_app.Visible = False
    return True

def print_pdf_files(pdf_files):
    """静默打印所有 PDF 文件"""
    for path in pdf_files:
        try:
            win32api.ShellExecute(0, "print", path, None, None, 0)
        except Exception as e:
            messagebox.showwarning("跳过文件", f"无法打印 PDF：\n{path}\n\n错误：{e}")

# ------------------------- 文件选择窗口 -------------------------
class FileSelector(tk.Toplevel):
    """可视化勾选文件窗口"""
    def __init__(self, parent, all_files):
        super().__init__(parent)
        self.title("选择要打印的文件")
        self.geometry("700x500")
        self.resizable(True, True)

        # 变量存储每个文件的勾选状态
        self.vars = {}
        for f in all_files:
            self.vars[f] = tk.BooleanVar(value=True)   # 默认全部选中

        # 顶部提示
        ttk.Label(self, text="请勾选需要打印的文件（默认全选）：", font=("微软雅黑", 10)).pack(pady=5)

        # 带滚动条的复选框列表
        frame = ttk.Frame(self)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        canvas = tk.Canvas(frame, borderwidth=0)
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=canvas.yview)
        self.check_frame = ttk.Frame(canvas)

        self.check_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.check_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 生成复选框
        for path in all_files:
            cb = ttk.Checkbutton(self.check_frame, text=os.path.basename(path), variable=self.vars[path])
            cb.pack(anchor=tk.W, padx=5, pady=2)

        # 底部按钮
        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="全选", command=self.select_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="全不选", command=self.deselect_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="开始打印", command=self.on_confirm).pack(side=tk.LEFT, padx=20)
        ttk.Button(btn_frame, text="取消", command=self.on_cancel).pack(side=tk.LEFT, padx=5)

        self.selected_files = []   # 最终选择的文件路径
        self.confirmed = False     # 是否点击了“开始打印”

    def select_all(self):
        for var in self.vars.values():
            var.set(True)

    def deselect_all(self):
        for var in self.vars.values():
            var.set(False)

    def on_confirm(self):
        self.selected_files = [path for path, var in self.vars.items() if var.get()]
        if not self.selected_files:
            messagebox.showwarning("提示", "请至少选择一个文件！")
            return
        self.confirmed = True
        self.destroy()

    def on_cancel(self):
        self.selected_files = []
        self.confirmed = False
        self.destroy()

# ------------------------- 主流程 -------------------------
def main():
    root = tk.Tk()
    root.withdraw()

    folder = filedialog.askdirectory(title="选择包含 Word/Excel/PDF 文档的文件夹", initialdir="D:\\")
    if not folder:
        messagebox.showinfo("取消", "未选择文件夹，程序退出。")
        return
    if not os.path.isdir(folder):
        messagebox.showerror("错误", f"文件夹不存在：\n{folder}")
        return

    # 收集所有支持的文件（忽略子文件夹）
    extensions = ('.doc', '.docx', '.xls', '.xlsx', '.pdf')
    all_files = []
    for f in os.listdir(folder):
        full = os.path.join(folder, f)
        if os.path.isfile(full) and f.lower().endswith(extensions):
            all_files.append(full)

    if not all_files:
        messagebox.showinfo("提示", "所选文件夹下没有 Word/Excel/PDF 文件。")
        return

    all_files.sort()

    # 弹出可视化勾选窗口
    selector = FileSelector(root, all_files)
    root.wait_window(selector)   # 等待窗口关闭

    if not selector.confirmed or not selector.selected_files:
        messagebox.showinfo("取消", "已取消打印。")
        return

    selected = selector.selected_files
    word_files  = [f for f in selected if f.lower().endswith(('.doc', '.docx'))]
    excel_files = [f for f in selected if f.lower().endswith(('.xls', '.xlsx'))]
    pdf_files   = [f for f in selected if f.lower().endswith('.pdf')]

    # 准备 Word / Excel 应用（必要时再获取）
    word_app = None
    excel_app = None

    # 1. 处理 Word 文件
    if word_files:
        word_app = get_word_app()
        if word_app is None:
            messagebox.showerror("错误", "无法启动 Word 应用程序，请确认已安装 Office 或 WPS。")
            return
        if not print_word_files(word_app, word_files):
            word_app.Quit()
            return

    # 2. 处理 Excel 文件
    if excel_files:
        excel_app = get_excel_app()
        if excel_app is None:
            messagebox.showerror("错误", "无法启动 Excel 应用程序，请确认已安装 Office 或 WPS。")
            if word_app:
                word_app.Quit()
            return
        if not print_excel_files(excel_app, excel_files):
            excel_app.Quit()
            if word_app:
                word_app.Quit()
            return

    # 3. 处理 PDF 文件
    if pdf_files:
        # 如果之前没有弹出过打印设置，则用临时空白 Word 弹出一次
        if not word_files and not excel_files:
            temp_word = get_word_app()
            if temp_word is None:
                messagebox.showerror("错误", "无法启动 Word 应用程序以设置打印机，请检查 Office/WPS 安装。")
                if word_app:
                    word_app.Quit()
                if excel_app:
                    excel_app.Quit()
                return
            temp_word.Visible = True
            setup_printer_dialog(temp_word)
            temp_word.Quit()
        print_pdf_files(pdf_files)

    # 清理
    if word_app:
        word_app.Quit()
    if excel_app:
        excel_app.Quit()

    messagebox.showinfo("完成", "批量打印已结束。")

if __name__ == "__main__":
    main()