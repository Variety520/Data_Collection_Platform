# settings_dialog页面
import tkinter as tk
from tkinter import messagebox


class SettingsDialog:
    def __init__(self, parent, initial_settings):
        self.parent = parent
        self.result = None

        self.top = tk.Toplevel(parent)
        self.top.title("界面设置")
        self.top.geometry("420x300")
        self.top.resizable(False, False)
        self.top.transient(parent)
        self.top.grab_set()

        self.buttons_per_row_var = tk.StringVar(
            value=str(initial_settings.get("buttons_per_row", 5))
        )
        self.visible_rows_var = tk.StringVar(
            value=str(initial_settings.get("visible_rows", 6))
        )
        self.button_text_max_var = tk.StringVar(
            value=str(initial_settings.get("button_text_max", 8))
        )

        self.build_ui()

        self.top.protocol("WM_DELETE_WINDOW", self.on_cancel)
        self.top.bind("<Return>", lambda event: self.on_ok())
        self.top.bind("<Escape>", lambda event: self.on_cancel())

    def build_ui(self):
        main_frame = tk.Frame(self.top, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        title_label = tk.Label(
            main_frame,
            text="设置按钮显示参数",
            font=("微软雅黑", 12, "bold")
        )
        title_label.pack(pady=(0, 15))

        row_1 = tk.Frame(main_frame)
        row_1.pack(fill=tk.X, pady=5)
        tk.Label(row_1, text="每行按钮数:", width=20, anchor="w").pack(side=tk.LEFT)
        tk.Spinbox(
            row_1,
            from_=1,
            to=20,
            textvariable=self.buttons_per_row_var,
            width=10
        ).pack(side=tk.LEFT)

        row_2 = tk.Frame(main_frame)
        row_2.pack(fill=tk.X, pady=5)
        tk.Label(
            row_2,
            text="按钮区可视行数:",
            width=20,
            anchor="w"
        ).pack(side=tk.LEFT)
        tk.Spinbox(
            row_2,
            from_=1,
            to=30,
            textvariable=self.visible_rows_var,
            width=10
        ).pack(side=tk.LEFT)

        row_3 = tk.Frame(main_frame)
        row_3.pack(fill=tk.X, pady=5)
        tk.Label(
            row_3,
            text="按钮文字显示字数:",
            width=20,
            anchor="w"
        ).pack(side=tk.LEFT)
        tk.Spinbox(
            row_3,
            from_=1,
            to=30,
            textvariable=self.button_text_max_var,
            width=10
        ).pack(side=tk.LEFT)

        note_label = tk.Label(
            main_frame,
            text="说明：超过可视行数后，字段按钮区会自动出现滚动效果。",
            fg="gray"
        )
        note_label.pack(pady=(15, 10), anchor="w")

        button_frame = tk.Frame(main_frame)
        button_frame.pack(side=tk.BOTTOM, pady=(10, 0))

        tk.Button(button_frame, text="确定", width=10, command=self.on_ok).pack(
            side=tk.LEFT, padx=5
        )
        tk.Button(button_frame, text="取消", width=10, command=self.on_cancel).pack(
            side=tk.LEFT, padx=5
        )

    def on_ok(self):
        try:
            buttons_per_row = int(self.buttons_per_row_var.get())
            visible_rows = int(self.visible_rows_var.get())
            button_text_max = int(self.button_text_max_var.get())
        except ValueError:
            messagebox.showerror("错误", "设置项必须是整数", parent=self.top)
            return

        if buttons_per_row <= 0 or visible_rows <= 0 or button_text_max <= 0:
            messagebox.showerror("错误", "设置项必须大于 0", parent=self.top)
            return

        self.result = {
            "buttons_per_row": buttons_per_row,
            "visible_rows": visible_rows,
            "button_text_max": button_text_max
        }
        self.top.destroy()

    def on_cancel(self):
        self.result = None
        self.top.destroy()

    def show(self):
        self.parent.wait_window(self.top)
        return self.result