import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from docx import Document
import pandas as pd
import os


class StudentManager:
    def __init__(self):
        self.students = []
        self.current_file = ""

    def load_word_file(self, filename):
        try:
            doc = Document(filename)
            self.students = []

            # 假设数据存储在Word表格中
            for table in doc.tables:
                for row in table.rows[1:]:  # 跳过表头
                    cells = row.cells
                    if len(cells) >= 4:
                        student = {
                            "姓名": cells[0].text.strip(),
                            "学号": cells[1].text.strip(),
                            "年龄": cells[2].text.strip(),
                            "成绩": cells[3].text.strip()
                        }
                        self.students.append(student)
            self.current_file = filename
            return True
        except Exception as e:
            messagebox.showerror("错误", f"文件读取失败: {str(e)}")
            return False


class StudentApp:
    def __init__(self, root):
        self.manager = StudentManager()
        self.root = root
        self.root.title("学生信息管理系统")
        self.root.geometry("800x600")

        # 创建界面组件
        self.create_widgets()
        self.update_listbox()

    def create_widgets(self):
        # 工具栏
        toolbar = ttk.Frame(self.root)
        toolbar.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(toolbar, text="打开Word文件", command=self.open_file).pack(side=tk.LEFT)
        ttk.Button(toolbar, text="导出Excel", command=self.export_excel).pack(side=tk.LEFT)
        ttk.Button(toolbar, text="添加记录", command=self.add_record).pack(side=tk.LEFT)
        ttk.Button(toolbar, text="删除记录", command=self.delete_record).pack(side=tk.LEFT)

        # 列表显示
        self.listbox = tk.Listbox(self.root, width=100)
        self.listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 详细信息
        detail_frame = ttk.Frame(self.root)
        detail_frame.pack(fill=tk.X, padx=5, pady=5)

        self.entries = {}
        fields = [("姓名", 0), ("学号", 1), ("年龄", 2), ("成绩", 3)]
        for field, column in fields:
            frame = ttk.Frame(detail_frame)
            frame.grid(row=0, column=column, padx=5)
            ttk.Label(frame, text=field + ":").pack()
            entry = ttk.Entry(frame, width=15)
            entry.pack()
            self.entries[field] = entry

        ttk.Button(detail_frame, text="保存修改", command=self.save_edit).grid(row=0, column=4, padx=5)

    def open_file(self):
        filetypes = [("Word文件", "*.docx"), ("所有文件", "*.*")]
        filename = filedialog.askopenfilename(filetypes=filetypes)
        if filename and self.manager.load_word_file(filename):
            self.update_listbox()
            messagebox.showinfo("成功", f"已加载 {len(self.manager.students)} 条记录")

    def update_listbox(self):
        self.listbox.delete(0, tk.END)
        for student in self.manager.students:
            self.listbox.insert(tk.END,
                                f"{student['姓名']} | {student['学号']} | {student['年龄']} | {student['成绩']}")

    def add_record(self):
        new_student = {
            "姓名": self.entries["姓名"].get(),
            "学号": self.entries["学号"].get(),
            "年龄": self.entries["年龄"].get(),
            "成绩": self.entries["成绩"].get()
        }
        if all(new_student.values()):
            self.manager.students.append(new_student)
            self.update_listbox()
            self.clear_entries()
        else:
            messagebox.showwarning("警告", "所有字段必须填写")

    def delete_record(self):
        selection = self.listbox.curselection()
        if selection:
            self.manager.students.pop(selection[0])
            self.update_listbox()

    def save_edit(self):
        selection = self.listbox.curselection()
        if selection:
            index = selection[0]
            self.manager.students[index] = {
                "姓名": self.entries["姓名"].get(),
                "学号": self.entries["学号"].get(),
                "年龄": self.entries["年龄"].get(),
                "成绩": self.entries["成绩"].get()
            }
            self.update_listbox()

    def clear_entries(self):
        for entry in self.entries.values():
            entry.delete(0, tk.END)

    def export_excel(self):
        if not self.manager.students:
            messagebox.showwarning("警告", "没有数据可以导出")
            return

        filetypes = [("Excel文件", "*.xlsx")]
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=filetypes,
            initialdir=os.path.expanduser("~/Desktop")
        )
        if filename:
            try:
                df = pd.DataFrame(self.manager.students)
                df.to_excel(filename, index=False)
                messagebox.showinfo("成功", f"文件已保存到 {filename}")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = StudentApp(root)
    root.mainloop()