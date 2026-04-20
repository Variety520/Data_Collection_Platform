#  file_import页面
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

from main_page import MainPage
from settings_dialog import SettingsDialog


class FileImportPage:
    def __init__(self, root):
        self.root = root
        self.root.title("佰态 · 病例数据采集工作台")
        self.root.geometry("600x400")
        self.root.attributes("-topmost", True)

        self.main_frame = tk.Frame(root)
        self.main_frame.pack(pady=20)

        self.label = tk.Label(
            self.main_frame,
            text="请导入你需要收集的Excel表格以作为索引和工作薄的建立"
        )
        self.label.pack(pady=10)

        self.file_path_var = tk.StringVar()
        self.file_path_entry = tk.Entry(
            self.main_frame,
            textvariable=self.file_path_var,
            width=50
        )
        self.file_path_entry.pack(pady=10)

        self.import_button = tk.Button(
            self.main_frame,
            text="导入Excel文件",
            command=self.import_excel
        )
        self.import_button.pack(pady=10)

        self.author_label = tk.Label(self.main_frame, text="关于作者：\n"
                                                           "开发者：佰态；"
                                                           "邮箱：qiangdehu@gmail")
        self.author_label.pack(pady=10)

        self.notice_label = tk.Label(
            self.main_frame,
            text="使用须知：作者懒，软件简陋，而相对有用！不喜勿喷。\n"
                 "小Tip：直接点击病人ID附近可直接复制住院号哦！"
        )
        self.notice_label.pack(pady=10)

        self.excel_data = None
        self.sheet_name = None
        self.main_page = None

        self.default_settings = {
            "buttons_per_row": 5,
            "visible_rows": 6,
            "button_text_max": 8
        }

    def import_excel(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")]
        )

        if not file_path:
            return

        self.file_path_var.set(file_path)

        try:
            excel_file = pd.ExcelFile(file_path, engine="openpyxl")
            self.sheet_name = excel_file.sheet_names[0]
            self.excel_data = pd.read_excel(
                file_path,
                sheet_name=self.sheet_name,
                engine="openpyxl"
            )

            if self.validate_excel(self.excel_data):
                response = messagebox.askyesno(
                    "确认",
                    "是否要开始建立单个病人工作薄以开始数据收集？"
                )
                if response:
                    self.start_data_collection()
            else:
                messagebox.showerror("错误", "导入的表格不合格，请检查内容")

        except Exception as e:
            messagebox.showerror("错误", "无法读取文件: {0}".format(e))

    def validate_excel(self, excel_data):
        from data_processing import validate_excel
        return validate_excel(excel_data)

    def open_settings_dialog(self):
        dialog = SettingsDialog(self.root, self.default_settings)
        result = dialog.show()

        if result is not None:
            self.default_settings = result.copy()

        return result

    def start_data_collection(self):
        settings = self.open_settings_dialog()
        if settings is None:
            return

        self.main_frame.destroy()
        self.main_page = MainPage(
            self.root,
            self.excel_data,
            self.file_path_var.get(),
            self.sheet_name,
            settings
        )


if __name__ == "__main__":
    root = tk.Tk()
    FileImportPage(root)
    root.mainloop()