import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter
import os
import warnings

warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", message="Multiple definitions in dictionary")

def parse_pages(page_str):
    pages = []
    page_str = page_str.replace("，", ",").replace(" ", "").strip()
    if not page_str:
        return pages
    for part in page_str.split(","):
        if "-" in part:
            try:
                start, end = map(int, part.split("-"))
                pages.extend(range(start, end + 1))
            except:
                pass
        else:
            try:
                pages.append(int(part))
            except:
                pass
    return sorted(list(set(pages)))

def cut_pdf_gui(input_path, pages_str, output_path):
    pages = parse_pages(pages_str)
    if not os.path.exists(input_path):
        messagebox.showerror("错误", "输入文件不存在！")
        return
    if not pages:
        messagebox.showerror("错误", "请输入有效页码！")
        return
    try:
        reader = PdfReader(input_path)
        writer = PdfWriter()
        total = len(reader.pages)
        valid = []
        for p in pages:
            if 1 <= p <= total:
                writer.add_page(reader.pages[p-1])
                valid.append(p)
        if not valid:
            messagebox.showerror("错误", "无有效页码可截取！")
            return
        with open(output_path, "wb") as f:
            writer.write(f)
        messagebox.showinfo("成功", f"截取完成！\n文件：{output_path}")
    except Exception as e:
        messagebox.showerror("失败", f"截取失败：{str(e)}")

def insert_pdf_gui(main_path, insert_path, pos_str, pages_str, output_path):
    if not os.path.exists(main_path) or not os.path.exists(insert_path):
        messagebox.showerror("错误", "主PDF/插入PDF不存在！")
        return
    try:
        insert_pos = int(pos_str)
    except:
        messagebox.showerror("错误", "插入位置必须是数字！")
        return
    pages = parse_pages(pages_str)
    if not pages:
        messagebox.showerror("错误", "请输入有效插入页码！")
        return
    try:
        main_reader = PdfReader(main_path)
        insert_reader = PdfReader(insert_path)
        writer = PdfWriter()
        main_total = len(main_reader.pages)

        for i in range(insert_pos):
            writer.add_page(main_reader.pages[i])
        insert_total = len(insert_reader.pages)
        for p in pages:
            if 1 <= p <= insert_total:
                writer.add_page(insert_reader.pages[p-1])
        for i in range(insert_pos, main_total):
            writer.add_page(main_reader.pages[i])

        with open(output_path, "wb") as f:
            writer.write(f)
        messagebox.showinfo("成功", f"插入完成！\n文件：{output_path}")
    except Exception as e:
        messagebox.showerror("失败", f"插入失败：{str(e)}")

class PDFToolGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF截取&插入工具")
        self.root.geometry("650x380")
        self.root.resizable(False, False)

        tab_control = ttk.Notebook(root)
        self.tab_cut = ttk.Frame(tab_control)
        self.tab_insert = ttk.Frame(tab_control)
        tab_control.add(self.tab_cut, text="📄 截取PDF页面")
        tab_control.add(self.tab_insert, text="📎 插入PDF页面")
        tab_control.pack(expand=1, fill="both", padx=10, pady=10)

        ttk.Label(self.tab_cut, text="源PDF文件：").grid(row=0, column=0, sticky="w", padx=5, pady=12)
        self.cut_input = tk.StringVar()
        ttk.Entry(self.tab_cut, textvariable=self.cut_input, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(self.tab_cut, text="选择", command=self.select_cut_input).grid(row=0, column=2, padx=5)

        ttk.Label(self.tab_cut, text="截取页码：").grid(row=1, column=0, sticky="w", padx=5, pady=12)
        self.cut_pages = tk.StringVar()
        ttk.Entry(self.tab_cut, textvariable=self.cut_pages, width=50).grid(row=1, column=1, padx=5)
        ttk.Label(self.tab_cut, text="例：1-3,5-7").grid(row=1, column=2, sticky="w")

        ttk.Label(self.tab_cut, text="输出文件：").grid(row=2, column=0, sticky="w", padx=5, pady=12)
        self.cut_output = tk.StringVar()
        ttk.Entry(self.tab_cut, textvariable=self.cut_output, width=50).grid(row=2, column=1, padx=5)
        ttk.Button(self.tab_cut, text="保存", command=self.select_cut_output).grid(row=2, column=2, padx=5)
        ttk.Button(self.tab_cut, text="✅ 开始截取", command=self.do_cut, width=20).grid(row=3, column=1, pady=20)

        ttk.Label(self.tab_insert, text="主PDF（被插入）：").grid(row=0, column=0, sticky="w", padx=5, pady=8)
        self.insert_main = tk.StringVar()
        ttk.Entry(self.tab_insert, textvariable=self.insert_main, width=45).grid(row=0, column=1, padx=5)
        ttk.Button(self.tab_insert, text="选择", command=self.select_insert_main).grid(row=0, column=2, padx=5)

        ttk.Label(self.tab_insert, text="插入PDF（来源）：").grid(row=1, column=0, sticky="w", padx=5, pady=8)
        self.insert_file = tk.StringVar()
        ttk.Entry(self.tab_insert, textvariable=self.insert_file, width=45).grid(row=1, column=1, padx=5)
        ttk.Button(self.tab_insert, text="选择", command=self.select_insert_file).grid(row=1, column=2, padx=5)

        ttk.Label(self.tab_insert, text="插入位置（页后）：").grid(row=2, column=0, sticky="w", padx=5, pady=8)
        self.insert_pos = tk.StringVar()
        ttk.Entry(self.tab_insert, textvariable=self.insert_pos, width=20).grid(row=2, column=1, sticky="w", padx=5)

        ttk.Label(self.tab_insert, text="插入页码：").grid(row=3, column=0, sticky="w", padx=5, pady=8)
        self.insert_pages = tk.StringVar()
        ttk.Entry(self.tab_insert, textvariable=self.insert_pages, width=45).grid(row=3, column=1, padx=5)
        ttk.Label(self.tab_insert, text="例：1-3").grid(row=3, column=2, sticky="w")

        ttk.Label(self.tab_insert, text="输出文件：").grid(row=4, column=0, sticky="w", padx=5, pady=8)
        self.insert_output = tk.StringVar()
        ttk.Entry(self.tab_insert, textvariable=self.insert_output, width=45).grid(row=4, column=1, padx=5)
        ttk.Button(self.tab_insert, text="保存", command=self.select_insert_output).grid(row=4, column=2, padx=5)
        ttk.Button(self.tab_insert, text="✅ 开始插入", command=self.do_insert, width=20).grid(row=5, column=1, pady=15)

    def select_cut_input(self):
        path = filedialog.askopenfilename(filetypes=[("PDF文件", "*.pdf")])
        if path: self.cut_input.set(path)
    def select_cut_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF文件", "*.pdf")])
        if path: self.cut_output.set(path)
    def do_cut(self):
        cut_pdf_gui(self.cut_input.get(), self.cut_pages.get(), self.cut_output.get())

    def select_insert_main(self):
        path = filedialog.askopenfilename(filetypes=[("PDF文件", "*.pdf")])
        if path: self.insert_main.set(path)
    def select_insert_file(self):
        path = filedialog.askopenfilename(filetypes=[("PDF文件", "*.pdf")])
        if path: self.insert_file.set(path)
    def select_insert_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF文件", "*.pdf")])
        if path: self.insert_output.set(path)
    def do_insert(self):
        insert_pdf_gui(self.insert_main.get(), self.insert_file.get(), self.insert_pos.get(), self.insert_pages.get(), self.insert_output.get())

if __name__ == "__main__":
    root = tk.Tk()
    PDFToolGUI(root)
    root.mainloop()
