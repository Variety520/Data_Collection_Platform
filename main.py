# main页面
import tkinter as tk
import tkinter.font as tkfont
from file_import import FileImportPage


def setup_global_fonts(root):
    available_fonts = set(tkfont.families(root))

    if "微软雅黑" in available_fonts:
        font_family = "微软雅黑"
    elif "Microsoft YaHei UI" in available_fonts:
        font_family = "Microsoft YaHei UI"
    else:
        # 兜底：用 Tk 默认逻辑字体，不强行指定不存在的中文字体
        font_family = None

    font_names = [
        "TkDefaultFont",
        "TkTextFont",
        "TkMenuFont",
        "TkHeadingFont",
        "TkCaptionFont",
        "TkSmallCaptionFont",
        "TkIconFont",
        "TkTooltipFont"
    ]

    for font_name in font_names:
        try:
            f = tkfont.nametofont(font_name)
            if font_family is not None:
                f.configure(family=font_family, size=10)
            else:
                f.configure(size=10)
        except tk.TclError:
            pass


def main():
    root = tk.Tk()
    setup_global_fonts(root)
    FileImportPage(root)
    root.mainloop()


if __name__ == "__main__":
    main()